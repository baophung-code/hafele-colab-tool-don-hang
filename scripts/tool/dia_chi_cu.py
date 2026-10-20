# scripts/tool/dia_chi_cu.py
# -*- coding: utf-8 -*-
"""Suy Tỉnh/Quận/Phường CŨ thuần từ chữ trong địa chỉ SAP. Thuần logic.

KHÔNG import openpyxl. KHÔNG gọi mạng. KHÔNG dùng toạ độ/geocode.
Nhờ vậy test được toàn bộ bằng dữ liệu bịa, không chạm internet.
Ranh giới này có test canh: xem test_dia_chi_cu_khong_import_openpyxl_va_khong_goi_mang.
"""
import csv
import difflib
import re
import unicodedata

# Nguong tuong dong toi thieu de chap nhan khop mo (fuzzy) ten Quan/Huyen khach tu go
# tay bi du/thieu/lap vai ky tu (vd 'Dong Xoaii' go du 1 chu 'i' so voi 'Dong Xoai') -
# xem DuLieu.chuan_cap_stm(). 0.85 du chat de khong khop nham 2 ten quan khac nhau
# nhung gan giong (da do thu tren danh_muc_stm.csv that).
_NGUONG_MO_TEN_QUAN = 0.85

# Tiền tố cấp hành chính, khớp ở ĐẦU chuỗi và phải có khoảng trắng theo sau.
# Không dùng \b giữa chuỗi: 'Thành' trong 'Kho Thành Đạt' sẽ bị cắt oan.
# Thứ tự trong nhóm là first-match-wins: cụm 2 từ phải đứng TRƯỚC từ đơn cùng gốc
# ('thi xa' trước 'xa', 'thanh pho' trước 'p'), nếu không 'Thị xã Dĩ An' sẽ hỏng.
_TIEN_TO = re.compile(
    r"^(thanh pho|thi xa|thi tran|quan|huyen|tinh|phuong|xa|tp|tx|tt|p|q)\s+",
    re.IGNORECASE)

# 'Tịnh' (dấu nặng, tên riêng vd huyện Tịnh Biên - An Giang) mất dấu thanh
# sau NFD-strip giống hệt tiền tố 'Tỉnh ' (dấu hỏi) -> phải chặn TRƯỚC khi bỏ
# dấu, lúc còn phân biệt được 2 dấu thanh.
_NGOAI_LE_TINH = re.compile(r"^tịnh\s+", re.IGNORECASE)

# Cùng họ bug với 'Tịnh Biên': huyện Quản Bạ (Hà Giang), Quan Hóa/Quan Sơn (Thanh
# Hóa) là TÊN RIÊNG thật, nhưng 'Quản'/'Quan' mất dấu thanh sau NFD-strip giống
# hệt tiền tố 'Quận ' (dấu nặng) -> _TIEN_TO cắt oan thành 'Bạ'/'Hóa'/'Sơn'. Phát
# hiện 20/10/2026 qua scripts/kiem_tra_ten_quan_stm.py. CHỈ liệt kê đích danh 2 tên
# riêng chắc chắn này (không chặn mọi 'Quan X' nói chung) để không phá khả năng
# nhận 'Quan 1' = 'Quận 1' viết tắt không dấu — tiền tố 'quận' (dấu nặng) thật vẫn
# bị cắt bình thường ở nhánh khác.
_TU_DAU_QUAN = re.compile(r"^(qu[aậảâ]n)\s+", re.IGNORECASE)


def _la_ten_rieng_quan(s):
    m = _TU_DAU_QUAN.match(s.strip())
    if not m or m.group(1).lower() in ("quận",):
        return False
    phan_sau = s.strip()[m.end():]
    phan_sau = unicodedata.normalize("NFD", phan_sau)
    phan_sau = "".join(c for c in phan_sau if unicodedata.category(c) != "Mn").lower()
    tu_dau = m.group(1).lower()
    if tu_dau == "quản":
        return phan_sau.startswith("ba")
    return phan_sau.startswith("hoa") or phan_sau.startswith("son")


def chuan_ten(s):
    """'Quận Ba Đình' -> 'ba dinh'. Quy y->i để bắt 'Quy Nhơn' == 'Qui Nhơn'."""
    s = str(s)
    la_ten_rieng = bool(_NGOAI_LE_TINH.match(s.strip())) or _la_ten_rieng_quan(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D").lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    if not la_ten_rieng:
        s = _TIEN_TO.sub("", s).strip()
    return s.replace("y", "i")


def _giu_dau(s):
    """Chuẩn hoá NHẸ: hạ chữ thường, gộp khoảng trắng — GIỮ NGUYÊN mọi dấu (kể cả dấu
    thanh) và tiền tố. Dùng làm TRỌNG TÀI phân xử cuối khi mơ hồ do chuan_ten() cố ý
    bỏ dấu thanh (để chịu lỗi gõ sai dấu) làm 2 tên khác nhau trông giống hệt nhau (vd
    'Thạnh' và 'Thành' đều về 'thanh') — KHÔNG thay chuan_ten() chính, chỉ dùng để so
    khớp phụ khi đã mơ hồ, không làm giảm khả năng chịu lỗi gõ sai dấu ở nơi khác."""
    s = unicodedata.normalize("NFC", str(s)).lower()
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()


class DuLieu:
    """Ba nguồn tham chiếu đã nạp sẵn vào bộ nhớ.

    Nhận file-like object (không phải đường dẫn) để test truyền StringIO vào được.
    """

    def __init__(self, f_crosswalk, f_danh_muc, f_alias):
        self.ten_quan = {}            # chuan_ten(quan_cu) -> [(tinh_cu, quan_cu), ...]
        self.phuong_moi = {}          # chuan_ten(phuong_moi) -> [(tinh_cu, quan_cu), ...]
        # (loai_rieng, chuan_ten(phuong_moi)) -> [...] — phân biệt Phường/Xã/Thị trấn,
        # xem _loai_rieng_phuong(). Ưu tiên tra ở đây trước khi rơi về phuong_moi gộp chung.
        self.phuong_moi_co_cap = {}
        self.phuong_cu_theo_quan = {}  # (tinh_cu, quan_cu) -> {chuan_ten(phuong_cu): phuong_cu}
        # chuan_ten(phuong_cu) -> [(tinh_cu, quan_cu), ...] — tra theo tên phường CŨ ngay
        # cả khi CHƯA biết quận (bổ sung cho phuong_moi vốn chỉ tra theo tên phường MỚI sau
        # sáp nhập — tên cũ/mới thường khác hẳn, vd 'Bình Hưng Hòa A' (cũ) -> 'Bình Hưng
        # Hòa' (mới, bỏ hậu tố A), nếu SAP chỉ ghi tên cũ thì tra theo tên mới không ra).
        self.phuong_cu_flat = {}
        # (tinh_cu, quan_cu) -> {_giu_dau(tên phường cũ/mới của quận đó), ...} — trọng
        # tài phân xử mơ hồ do dấu thanh bị gộp mất trong chuan_ten() (xem _giu_dau()).
        self._phuong_dau = {}
        self._tinh_moi = {}            # chuan_ten(tinh_cu) -> tên tỉnh MỚI 2025

        for d in csv.DictReader(f_crosswalk):
            k = (d["tinh_cu"].strip(), d["quan_cu"].strip())
            self._tinh_moi.setdefault(chuan_ten(k[0]), d.get("tinh_moi", "").strip())
            ds = self.ten_quan.setdefault(chuan_ten(k[1]), [])
            if k not in ds:
                ds.append(k)
            pm_goc = d.get("phuong_moi", "").strip()
            pm = chuan_ten(pm_goc)
            if pm:
                self.phuong_moi.setdefault(pm, []).append(k)
                loai_pm = _loai_rieng_phuong(pm_goc)
                if loai_pm:
                    self.phuong_moi_co_cap.setdefault((loai_pm, pm), []).append(k)
                self._phuong_dau.setdefault(k, set()).add(_giu_dau(pm_goc))
            phuong_cu = d.get("phuong_cu", "").strip()
            if phuong_cu:
                self.phuong_cu_theo_quan.setdefault(k, {})[chuan_ten(phuong_cu)] = phuong_cu
                self._phuong_dau.setdefault(k, set()).add(_giu_dau(phuong_cu))
                ds_flat = self.phuong_cu_flat.setdefault(chuan_ten(phuong_cu), [])
                if k not in ds_flat:
                    ds_flat.append(k)

        self._alias = {}
        for d in csv.DictReader(f_alias):
            self._alias[(d["loai"], chuan_ten(d["ten_crosswalk"]))] = chuan_ten(d["ten_stm"])

        self._stm_tinh = {}       # chuan_ten -> tên hiển thị
        self._stm_quan = {}       # (chuan_ten tỉnh, chuan_ten quận) -> (tên tỉnh, tên quận)
        for d in csv.DictReader(f_danh_muc):
            ten = d["ten"].strip()
            if not ten or ten == "INVALID":
                continue
            if d["loai"] == "tinh":
                self._stm_tinh[chuan_ten(ten)] = ten
            else:
                tinh = d["tinh"].strip()
                if tinh:
                    self._stm_quan[(chuan_ten(tinh), chuan_ten(ten))] = (tinh, ten)

    def tinh_moi_cua(self, tinh_cu):
        """Tỉnh CŨ -> tên tỉnh MỚI 2025 đã nuốt nó. Không biết -> ''."""
        return self._tinh_moi.get(chuan_ten(tinh_cu), "")

    def chuan_cap_stm(self, tinh, quan):
        """(tỉnh, quận) bất kỳ -> đúng chính tả danh mục STM. Không có -> None.

        Không chỉ kiểm hợp lệ mà còn ĐỔI CHÍNH TẢ: nhánh dự phòng cho ra
        'Huyện Long Điền' trong khi danh mục STM ghi 'Long Điền' — giữ dạng thô
        là STM từ chối. Đo trên data thật: 2/295 đơn.

        Khớp đúng thất bại -> thử khớp MỜ (fuzzy) theo tên quận, CHỈ trong phạm vi
        ĐÚNG tỉnh đã biết (ít ứng viên, ít rủi ro khớp nhầm) và CHỈ nhận khi duy
        nhất 1 ứng viên vượt ngưỡng — bắt được lỗi gõ dư/thiếu vài ký tự của khách
        tự điền tay (vd 'Đồng Xoàii' -> 'Đồng Xoài'), không đoán bừa khi còn mơ hồ.
        """
        tc, qc = chuan_ten(tinh), chuan_ten(quan)
        ten = self._stm_quan.get((tc, qc))
        if ten is not None:
            return ten
        ung_vien = [qc_stm for (t, qc_stm) in self._stm_quan if t == tc]
        if not ung_vien:
            return None
        khop = difflib.get_close_matches(qc, ung_vien, n=2, cutoff=_NGUONG_MO_TEN_QUAN)
        if len(khop) == 1:
            return self._stm_quan.get((tc, khop[0]))
        return None

    def ten_stm(self, tinh_cu, quan_cu):
        """Tên crosswalk -> tên đúng như danh mục STM. Không có -> None.

        Không bịa: ngoài danh mục thì trả None để tầng trên để trống + cảnh báo,
        chứ không ghi 'INVALID' (đó là giá trị lỗi của chính STM).
        """
        tc = self._alias.get(("tinh", chuan_ten(tinh_cu)), chuan_ten(tinh_cu))
        qc = self._alias.get(("quan", chuan_ten(quan_cu)), chuan_ten(quan_cu))
        return self._stm_quan.get((tc, qc))

    def ten_stm_tinh(self, tinh_cu):
        """Chỉ riêng cấp TỈNH -> tên đúng như danh mục STM. Không có -> None.

        Dùng khi quận không hợp lệ/không xác định nhưng tỉnh vẫn nên hiển thị
        đúng chính tả STM (vd 'Thành phố Hồ Chí Minh' -> 'Hồ Chí Minh').
        """
        tc = self._alias.get(("tinh", chuan_ten(tinh_cu)), chuan_ten(tinh_cu))
        return self._stm_tinh.get(tc)


# Mảnh là cấp QUẬN/HUYỆN khi bắt đầu bằng 1 trong các từ này (không gồm phường/xã).
# LƯU Ý: "thành phố"/"tp" KHÔNG ở đây — nó mơ hồ 2 cấp (TP thuộc tỉnh vs TP trực thuộc TW,
# vd "Thành phố Dĩ An" là quận-cấp nhưng "Thành phố Hồ Chí Minh" là tỉnh-cấp), rơi vào
# "khong_ro" để suy_dvhc_tu_text() tự thử khớp cả 2 vai trò.
_QUAN_PRE = re.compile(
    r"^(thi\s*x[aã]|qu[aậ]n|huy[eệ]n|tx\.?|q\.?)\s+", re.IGNORECASE)
# Mảnh là cấp PHƯỜNG/XÃ khi bắt đầu bằng 1 trong các từ này.
_PHUONG_PRE = re.compile(
    r"^(ph[uư][ơo]ng|x[aã]|th[iị]\s*tr[aấ]n|p\.?|tt\.?)\s+", re.IGNORECASE)
# Mảnh là cấp TỈNH khi bắt đầu bằng 1 trong các từ này.
# "thành phố"/"tp" KHÔNG ở đây cùng lý do trên (xem _QUAN_PRE).
_TINH_PRE = re.compile(r"^(t[iỉ]nh)\s+", re.IGNORECASE)

# Phân biệt RIÊNG Phường/Xã/Thị trấn trong nhóm "phuong" ở _PHUONG_PRE — dùng để
# KHÔNG lẫn 'Phường Thanh Thủy' (mới, đô thị) với 'Xã Thanh Thủy' (mới, nông thôn,
# ở TỈNH KHÁC) dù chuan_ten() cắt tiền tố làm chúng trùng chữ. Bug thật đã vấp
# 20/10/2026: SAP 'Phường Thanh Thủy' bị báo mơ hồ với cả 'Xã Thanh Thủy' (Phú
# Thọ) và 'Xã Thanh Thủy' (Hà Giang) — trong khi 2 nơi đó không hề có Phường nào
# tên này, chỉ trùng chữ sau khi cắt tiền tố.
def _chuan_ten_truoc_gach_ngang(manh_goc):
    """Phần TRƯỚC dấu '-' đệm khoảng trắng hai bên (vd 'Dương Đông - Phú Quốc' ->
    'Dương Đông'). CHỈ dùng làm phương án dự phòng khi tra nguyên mảnh thất bại — SAP hay
    nối thêm tên huyện/quận vào cùng mảnh với tên phường qua dấu gạch ngang thay vì tách
    bằng dấu phẩy. Không đụng tới cấp Tỉnh ở nơi khác (tên tỉnh có thể tự chứa gạch ngang
    như 'Bà Rịa - Vũng Tàu', nơi đó vẫn phải khớp nguyên văn)."""
    truoc = re.split(r"\s+-\s+", str(manh_goc), maxsplit=1)[0]
    return chuan_ten(truoc)


_PHUONG_RIENG = re.compile(r"^(ph[uư][ơo]ng|p\.?)\s+", re.IGNORECASE)
_XA_RIENG = re.compile(r"^(x[aã])\s+", re.IGNORECASE)
_THITRAN_RIENG = re.compile(r"^(th[iị]\s*tr[aấ]n|tt\.?)\s+", re.IGNORECASE)


def _loai_rieng_phuong(manh_goc):
    """-> "phuong" | "xa" | "thi_tran" | None. Chạy TRƯỚC chuan_ten() như _phan_loai_manh."""
    t = unicodedata.normalize("NFD", str(manh_goc).strip())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = t.replace("đ", "d").replace("Đ", "D").lower()
    t = re.sub(r"\s+", " ", t).strip()
    if _PHUONG_RIENG.match(t):
        return "phuong"
    if _XA_RIENG.match(t):
        return "xa"
    if _THITRAN_RIENG.match(t):
        return "thi_tran"
    return None


def _phan_loai_manh(manh_goc):
    """-> (cap, chuan_ten(manh_goc)). cap in {"quan","phuong","tinh","khong_ro"}.

    PHẢI phân loại theo tiền tố TRƯỚC khi gọi chuan_ten() (vốn cắt bỏ tiền tố để so khớp) —
    nếu không 'Phường 11' và 'Quận 11' sẽ trở thành cùng 1 chuỗi '11' sau khi cắt, gây nhận
    nhầm cấp hành chính. Bug thật đã vấp lúc đo thử trên data SAP thật 14/10/2026."""
    t = unicodedata.normalize("NFD", str(manh_goc).strip())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = t.replace("đ", "d").replace("Đ", "D").lower()
    t = re.sub(r"\s+", " ", t).strip()
    if _PHUONG_PRE.match(t):
        cap = "phuong"
    elif _QUAN_PRE.match(t):
        cap = "quan"
    elif _TINH_PRE.match(t):
        cap = "tinh"
    else:
        cap = "khong_ro"
    return cap, chuan_ten(manh_goc)


def _loc_dau_thanh(cands, manh_goc_list, du_lieu):
    """Mơ hồ nhiều quận (không lọc được bằng tỉnh/STM) nhưng chỉ vì chuan_ten() bỏ dấu
    thanh làm 2 tên phường khác nhau trông giống hệt nhau (vd 'Thạnh'/'Thành', xem
    _giu_dau()) -> so lại bằng bản GIỮ dấu thanh, ứng viên nào có tên phường (cũ hoặc
    mới) khớp đúng 1 trong các mảnh gốc thì chọn. Đúng 1 ứng viên khớp -> hết mơ hồ;
    0 hoặc ≥2 ứng viên khớp -> vẫn mơ hồ như cũ (không ép ra kết quả khi còn lưỡng lự)."""
    dau_goc = {_giu_dau(m) for m in manh_goc_list}
    khop = [c for c in cands if du_lieu._phuong_dau.get(c, set()) & dau_goc]
    return khop[0] if len(khop) == 1 else None


def suy_dvhc_tu_text(dia_chi_goc, ten_tinh_rg, du_lieu):
    """Suy Tỉnh/Quận/Phường CŨ thuần từ chữ trong địa chỉ SAP — KHÔNG toạ độ, KHÔNG mạng.

    -> dict(tinh_cu, quan_cu, phuong_cu, do_tin, ly_do). `ly_do` rỗng nếu suy được quận;
    "mo_ho" nếu khớp nhiều ứng viên không lọc nổi còn 1; "khong_co_quan" nếu không mảnh nào
    khớp. `tinh_cu` vẫn có thể có giá trị (từ chữ hoặc từ `ten_tinh_rg`) ngay cả khi
    `quan_cu` rỗng — 2 việc tách biệt, xem spec §2.5.
    """
    manh = [m.strip() for m in str(dia_chi_goc).split(",") if m.strip()]
    phan = [_phan_loai_manh(m) for m in manh]
    manh_quan = [ct for cap, ct in phan if cap == "quan"]
    manh_phuong = [ct for cap, ct in phan if cap == "phuong"]
    # (ct, loai_rieng, manh_goc) song song với manh_phuong — loai_rieng để tra
    # phuong_moi_co_cap (xem _loai_rieng_phuong()), manh_goc để thử fallback cắt trước
    # dấu gạch ngang khi tra bằng ct thất bại (xem _chuan_ten_truoc_gach_ngang()).
    manh_phuong_loai = [(ct, _loai_rieng_phuong(m), m)
                         for m, (cap, ct) in zip(manh, phan) if cap == "phuong"]
    manh_tinh = [ct for cap, ct in phan if cap == "tinh"]
    manh_khong_ro = [ct for cap, ct in phan if cap == "khong_ro"]
    manh_khong_ro_goc = [(ct, m) for m, (cap, ct) in zip(manh, phan) if cap == "khong_ro"]

    tinh_rg_chuan = chuan_ten(ten_tinh_rg) if ten_tinh_rg else None

    def khop_tinh(tinh_cu):
        tc = chuan_ten(tinh_cu)
        if tc in manh_tinh or tc in manh_khong_ro:
            return True
        tm = du_lieu.tinh_moi_cua(tinh_cu)
        tm_chuan = chuan_ten(tm) if tm else None
        if tm_chuan and (tm_chuan in manh_tinh or tm_chuan in manh_khong_ro):
            return True
        if tinh_rg_chuan and (tinh_rg_chuan == tc or tinh_rg_chuan == tm_chuan):
            return True
        return False

    def stm_nhat_quan(cands):
        """Nhiều ứng viên khác nhau ở cấp CŨ nhưng CÙNG map về 1 tên quận
        trên danh mục STM (vd 'Thành phố Cao Lãnh' / 'Huyện Cao Lãnh' đều
        chỉ còn 1 dòng 'Cao Lãnh' trong danh_muc_stm.csv) -> không còn mơ hồ
        với OUTPUT cuối cùng (cột Quận ra giống hệt nhau dù chọn ứng viên
        nào), trả ứng viên đầu tiên. Khác tên STM, hoặc có ứng viên không
        nằm trong danh mục STM -> None, vẫn coi là mơ hồ thật."""
        cands = list(cands)
        if len(cands) < 2:
            return None
        ten = [du_lieu.ten_stm(tc, qc) for tc, qc in cands]
        if all(t is not None for t in ten) and len(set(ten)) == 1:
            return cands[0]
        return None

    chon = None
    uv = []
    for ct in manh_quan + manh_khong_ro:
        for k in du_lieu.ten_quan.get(ct, []):
            if k not in uv:
                uv.append(k)
    def tinh_mo_ho(uv_tho, loc):
        """Quận mơ hồ (không lọc còn 1) nhưng vẫn có thể BIẾT TỈNH, theo thứ tự ưu tiên:
        1) ứng viên đã lọc theo Rg/text (`loc`) cùng 1 tỉnh -> chắc nhất, có xác nhận;
        2) KHÔNG lọc được ứng viên nào (`loc` rỗng) nhưng toàn bộ ứng viên thô (`uv_tho`,
           trước khi lọc) vốn dĩ đã cùng 1 tỉnh -> tỉnh không hề mơ hồ, chỉ quận mới mơ hồ,
           không cần Rg/text xác nhận gì thêm;
        3) còn lại thì tin vào Rg nếu có.
        Chỉ trả rỗng khi thật sự không còn gì để bám vào — xoá sạch tỉnh đã biết chỉ vì
        quận mơ hồ là bug thật đã vấp 16/10/2026 (45/295 đơn data thật)."""
        tinhs = {k[0] for k in loc}
        if len(tinhs) == 1:
            return next(iter(tinhs))
        tinhs_tho = {k[0] for k in uv_tho}
        if len(tinhs_tho) == 1:
            return next(iter(tinhs_tho))
        return ten_tinh_rg or ""

    if len(uv) == 1:
        chon = uv[0]
    elif len(uv) > 1:
        loc = [k for k in uv if khop_tinh(k[0])]
        if len(loc) == 1:
            chon = loc[0]
        else:
            chon = stm_nhat_quan(loc or uv)
            if chon is None:
                return {"tinh_cu": tinh_mo_ho(uv, loc), "quan_cu": "", "phuong_cu": "",
                        "do_tin": "THAP", "ly_do": "mo_ho", "ung_vien": loc or uv}

    def _tim_ung_vien_phuong(ct, manh_goc):
        """ct -> ứng viên, thử theo thứ tự: tên MỚI cùng cấp -> tên MỚI gộp chung ->
        tên CŨ (phuong_cu_flat) -> nếu vẫn rỗng và manh_goc có dấu gạch ngang, thử lại
        toàn bộ chuỗi trên với phần TRƯỚC dấu gạch ngang (xem _chuan_ten_truoc_gach_ngang)."""
        tim = du_lieu.phuong_moi.get(ct, []) or du_lieu.phuong_cu_flat.get(ct, [])
        if not tim:
            alt = _chuan_ten_truoc_gach_ngang(manh_goc)
            if alt and alt != ct:
                tim = du_lieu.phuong_moi.get(alt, []) or du_lieu.phuong_cu_flat.get(alt, [])
        return tim

    if chon is None:
        cand = []
        for ct, loai, m in manh_phuong_loai:
            # Ưu tiên tra ĐÚNG cấp (Phường/Xã/Thị trấn) trước — chỉ rơi về tra gộp
            # chung (level-agnostic, gồm cả tên cũ + fallback gạch ngang) khi không có
            # ứng viên nào cùng cấp, để không mất recall với trường hợp dữ liệu
            # crosswalk thiếu/không rõ cấp.
            cung_cap = du_lieu.phuong_moi_co_cap.get((loai, ct), []) if loai else []
            cand += cung_cap if cung_cap else _tim_ung_vien_phuong(ct, m)
        for ct, m in manh_khong_ro_goc:
            cand += _tim_ung_vien_phuong(ct, m)
        bo = {(c[0], c[1]) for c in cand}
        if len(bo) == 1:
            chon = next(iter(bo))
        elif len(bo) > 1:
            loc = {b for b in bo if khop_tinh(b[0])}
            if len(loc) == 1:
                chon = next(iter(loc))
            else:
                chon = stm_nhat_quan(loc or bo)
                if chon is None:
                    manh_phuong_raw = ([m for _, _, m in manh_phuong_loai]
                                        + [m for _, m in manh_khong_ro_goc])
                    chon = _loc_dau_thanh(loc or bo, manh_phuong_raw, du_lieu)
                if chon is None:
                    return {"tinh_cu": tinh_mo_ho(bo, loc), "quan_cu": "", "phuong_cu": "",
                            "do_tin": "THAP", "ly_do": "mo_ho", "ung_vien": list(loc or bo)}

    if chon is None:
        return {"tinh_cu": ten_tinh_rg or "", "quan_cu": "", "phuong_cu": "",
                "do_tin": "THAP", "ly_do": "khong_co_quan"}

    tinh_cu, quan_cu = chon
    phuong_cu = ""
    ds_phuong = du_lieu.phuong_cu_theo_quan.get(chon, {})
    for ct in manh_phuong:
        if ct in ds_phuong:
            phuong_cu = ds_phuong[ct]
            break

    return {"tinh_cu": tinh_cu, "quan_cu": quan_cu, "phuong_cu": phuong_cu,
            "do_tin": "CAO", "ly_do": ""}


# Mảnh là cấp hành chính khi BẮT ĐẦU bằng một trong các từ này.
# 'Khu Phố' không dính vì 'Phố' != 'Phường'. 'Tỉnh Lộ' bị loại riêng bên dưới.
_CAP = re.compile(
    r"^(phuong|xa|thi tran|quan|huyen|thi xa|thanh pho|tinh|tp|tt|tx|p|q)\s",
    re.IGNORECASE)
# Tên đường bắt đầu bằng từ trông giống cấp hành chính -> KHÔNG cắt.
_KHONG_CAT = re.compile(r"^(tinh lo|quan lo|huyen lo)\s", re.IGNORECASE)


def _la_manh_hanh_chinh(manh):
    t = unicodedata.normalize("NFD", str(manh).strip())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = t.replace("đ", "d").replace("Đ", "D").lower()
    t = re.sub(r"\s+", " ", t).strip()
    if _KHONG_CAT.match(t):
        return False
    return bool(_CAP.match(t))


def cat_duoi_hanh_chinh(dia_chi):
    """Bỏ các mảnh cấp hành chính ở CUỐI chuỗi, giữ nguyên phần số nhà + đường."""
    manh = [m.strip() for m in str(dia_chi).split(",")]
    while len(manh) > 1 and _la_manh_hanh_chinh(manh[-1]):
        manh.pop()
    return ", ".join(m for m in manh if m)


def ghep_dia_chi_cu(dia_chi_goc, phuong_cu, quan_cu, tinh_cu):
    """'<số nhà + đường>, Phường <cũ>, <Quận cũ>, <Tỉnh cũ>'. Cấp rỗng thì bỏ qua."""
    phan = [cat_duoi_hanh_chinh(dia_chi_goc)]
    phan += [x.strip() for x in (phuong_cu, quan_cu, tinh_cu) if x and x.strip()]
    return ", ".join(p for p in phan if p)


def _chon_ung_vien_bang_geocode(dia_chi_goc, ung_vien, lay_dvhc_cu_tu_geocode):
    """Mơ hồ (nhiều quận trùng tên) không lọc được bằng chữ -> hỏi geocode CHỈ để PHÂN XỬ
    giữa các ứng viên đã suy ra sẵn, không tự ý tin thẳng cấp hành chính geocode trả về
    (đã cấm từ nấc A2 cũ dùng Goong, 7,3% sai hết dù gắn CAO — xem CLAUDE.md). Geocode chỉ
    được chấp nhận khi khớp ĐÚNG 1 trong các ứng viên; khớp ứng viên lạ hoặc không khớp cái
    nào -> bỏ qua, giữ nguyên mơ hồ. `lay_dvhc_cu_tu_geocode` không hàm nào được ném
    exception (tiêm được cho test / None để tắt tính năng)."""
    if not ung_vien or lay_dvhc_cu_tu_geocode is None:
        return None
    kq = lay_dvhc_cu_tu_geocode(dia_chi_goc)
    if not kq:
        return None
    tc, qc = chuan_ten(kq[0]), chuan_ten(kq[1])
    khop = [u for u in ung_vien if chuan_ten(u[0]) == tc and chuan_ten(u[1]) == qc]
    return khop[0] if len(khop) == 1 else None


def tra_dia_chi(dia_chi_goc, ten_tinh_rg, du_lieu, lay_dvhc_cu_tu_geocode=None):
    """Cửa duy nhất cho dung_bang.py. -> dict(N, O, P, do_tin, ly_do).

    `ly_do`: "" (suy được quận) | "mo_ho" | "khong_co_quan" | "khong_hop_le_stm".
    `lay_dvhc_cu_tu_geocode`: hàm injectable `dia_chi -> (tinh_cu, quan_cu) | None` (xem
    trackasia_api.dvhc_cu) — CHỈ gọi khi suy chữ ra "mo_ho" và có ứng viên, dùng để phân xử
    (không dùng để tự suy từ đầu). None (mặc định) = tắt hẳn, hành vi y như trước.
    """
    if not str(dia_chi_goc).strip():
        return {"N": "", "O": "", "P": "", "do_tin": "THAP", "ly_do": "khong_co_quan"}

    s = suy_dvhc_tu_text(dia_chi_goc, ten_tinh_rg, du_lieu)
    if s["ly_do"] == "mo_ho":
        chon = _chon_ung_vien_bang_geocode(
            dia_chi_goc, s.get("ung_vien"), lay_dvhc_cu_tu_geocode)
        if chon is not None:
            s = {"tinh_cu": chon[0], "quan_cu": chon[1], "phuong_cu": "",
                 "do_tin": "CAO", "ly_do": ""}
    tinh_hien_thi = du_lieu.ten_stm_tinh(s["tinh_cu"]) or s["tinh_cu"]
    if s["ly_do"]:
        kq = {"N": "", "O": tinh_hien_thi, "P": "", "do_tin": s["do_tin"], "ly_do": s["ly_do"]}
        if s["ly_do"] == "mo_ho" and s.get("ung_vien"):
            kq["ung_vien"] = s["ung_vien"]
        return kq

    ten = du_lieu.ten_stm(s["tinh_cu"], s["quan_cu"])
    if ten is None:
        return {"N": "", "O": tinh_hien_thi, "P": "", "do_tin": "THAP",
                "ly_do": "khong_hop_le_stm",
                "tinh_tho": s["tinh_cu"], "quan_tho": s["quan_cu"]}
    tinh_stm, quan_stm = ten
    N = ghep_dia_chi_cu(dia_chi_goc, s["phuong_cu"], quan_stm, tinh_stm)
    return {"N": N, "O": tinh_stm, "P": quan_stm, "do_tin": "CAO", "ly_do": ""}

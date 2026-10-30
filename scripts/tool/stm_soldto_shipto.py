# -*- coding: utf-8 -*-
"""Logic quyết định tái dùng Sold-to/Ship-to có sẵn — tra cache local (xem stm_cache.py),
KHÔNG gọi mạng. Khi khớp mờ ra 1 Ship-to có sẵn, lấy cả Tỉnh/Quận/Địa chỉ đã đăng ký trên STM
(không chỉ địa chỉ như bản cũ) - xem spec §2.5: cache Ship-to đã kiêm vai trò "bộ nhớ địa chỉ".

Không khớp cache + SAP CÓ sẵn Mã Ship to gốc -> GIỮ NGUYÊN mã gốc đó (đổi hướng 29/10/2026):
STM tự khớp địa chỉ/đề xuất tạo mới lúc lên đơn, mã Hafele cũ gắn trên SAP vẫn là đầu mối hữu
ích để STM tra. Không khớp cache + SAP KHÔNG có mã (nhóm DECO) -> để trống + cảnh báo
"Thiếu Mã Ship to" (tô vàng cột L, nhóm NHOM_SAP - khách phải báo bộ phận SAP hoặc tự gõ mã).
KHÔNG tự sinh mã mới, KHÔNG gọi geocode, KHÔNG tự ghi cache ngay. Việc ghi điểm mới vào cache
dời sang main.xac_nhan_da_upload_don_hang(), chạy SAU KHI khách xác nhận đã upload thật lên STM
- và chỉ ghi dòng có L TRỐNG (DECO), dòng giữ mã gốc SAP không ghi (không biết STM nhận mã đó
hay tự tạo mã khác, giữ gap đã chấp nhận 20/10/2026).

Luồng tách 2 bước (xem docs/superpowers/specs/2026-10-15-tach-2-buoc-soldto-shipto-design.md):
- Bước 1 (main.chay()): gọi xu_ly(..., cho_phep_tao_moi=False) - chỉ tái dùng, không đụng
  dòng không khớp - để khách kịp sửa địa chỉ trong file upload trước.
- Bước 2 (main.tao_file_soldto_shipto()): gọi xu_ly(..., cho_phep_tao_moi=True) trên dữ liệu
  đọc lại từ file upload đã sửa - lúc này mới để trống Mã Ship to cho các dòng không khớp.
"""
from .dia_chi_so_khop import parse_address, same_place
from .dia_chi_cu import chuan_ten, suy_dvhc_tu_text

# same_place()/parse_address() được tune cho BBGH (gom địa chỉ SAP sạch, xem
# docs/QUYLUAT_GOM_BBGH.md §6) - KHÔNG sửa 2 hàm đó. Ở đây so khớp địa chỉ cache (vốn đến từ
# STM, nhập tay tự do, lẫn tên người/SĐT), rủi ro khớp nhầm khác hẳn: chặn thêm 1 lớp riêng
# cho việc reuse Ship-to, không đụng module dùng chung với BBGH.
_SO_TOKEN_TOI_THIEU = 4


def _du_tin_cay_de_tai_dung(toks_moi, toks_sp):
    """Chỉ tái dùng khi CẢ 2 địa chỉ đủ dài để tin cậy (>= 4 token có nghĩa) - chặn case
    địa chỉ cache quá ngắn/chung chung (vd chỉ có tên người + quận + tỉnh) trùng tình cờ với
    1 phần nhỏ của địa chỉ SAP dài hơn nhưng khác hẳn số nhà (bug thật gặp 07/10/2026:
    'Kho Anh Tuan, Quan 1, HCM' (3 token) bị coi là same_place với địa chỉ SAP có số nhà 123
    hoàn toàn khác, chỉ vì trùng tình cờ tên người 'Tuan' + 'Quan 1' + 'HCM')."""
    if len(toks_moi) < _SO_TOKEN_TOI_THIEU or len(toks_sp) < _SO_TOKEN_TOI_THIEU:
        return False
    return same_place(toks_moi, toks_sp)


def _cache_dong_bo_dia_chi(sp, du_lieu, canh_bao, bo_nho):
    """True = tin bản ghi cache, False = BỎ QUA (tỉnh/quận cache mâu thuẫn với chính
    địa chỉ text của nó - vd bản ghi tạo tay nhập Hà Nội/Ba Đình nhưng địa chỉ ghi
    Dĩ An, Bình Dương, phát hiện 29/10/2026). Chỉ bác khi suy chữ từ địa chỉ ra tỉnh
    với độ tin cậy CAO và khác tỉnh cache; mơ hồ hoặc không ra -> cho qua, không bác
    oan. Kết quả cache theo chuỗi địa chỉ để không gọi lặp."""
    dia_chi = sp.get("dia_chi", "") or ""
    if dia_chi in bo_nho:
        return bo_nho[dia_chi]
    s = suy_dvhc_tu_text(dia_chi, "", du_lieu)
    if s.get("ly_do") or s.get("do_tin") != "CAO" or not s.get("tinh_cu"):
        bo_nho[dia_chi] = True
        return True
    tinh_text = du_lieu.ten_stm_tinh(s["tinh_cu"]) or s["tinh_cu"]
    tinh_cache = sp.get("tinh", "") or ""
    if chuan_ten(tinh_text) != chuan_ten(tinh_cache):
        canh_bao.add(
            f"Ship-to {sp.get('ma','?')} tren STM dang sai Tinh "
            f"(cache: {tinh_cache}, dia chi text suy ra: {tinh_text}) - bo qua, khong dung")
        bo_nho[dia_chi] = False
        return False
    bo_nho[dia_chi] = True
    return True


_LY_DO_DVHC = (
    "Không xác định được Quận/Huyện", "Có nhiều Quận/Huyện cũ ứng với địa chỉ này",
    "Quận không có trong danh sách STM chấp nhận", "Không xác định được Tỉnh",
    "Tỉnh độ tin cậy thấp",
)
LY_DO_THIEU_MA_SHIPTO = "Thiếu Mã Ship to"


def _dung_format_dia_chi_cache(sp):
    """Chốt 30/10/2026: địa chỉ cache phải kết thúc bằng '..., <Quận cache>, <Tỉnh cache>'
    thì mới coi là format chuẩn để tái dùng. Có thể có 'Phường X' đứng trước Quận. Không
    đúng format -> coi như KHÔNG khớp, rơi vào nhánh tạo mới (giữ mã gốc SAP hoặc để trống
    DECO) - tránh ghi đè địa chỉ file upload bằng chuỗi cache lộn xộn từ STM cũ (nhập tay
    tự do, có khi chỉ 'HCM' hoặc dính tên người/SĐT)."""
    manh = [m.strip() for m in (sp.get("dia_chi") or "").split(",") if m.strip()]
    if len(manh) < 2:
        return False
    quan_cache = sp.get("quan", "") or ""
    tinh_cache = sp.get("tinh", "") or ""
    if not quan_cache or not tinh_cache:
        return False
    return (chuan_ten(manh[-1]) == chuan_ten(tinh_cache)
            and chuan_ten(manh[-2]) == chuan_ten(quan_cache))


def _nhom_theo_dia_chi(rows):
    nhom = {}
    for r in rows:
        nhom.setdefault(r.get("N", ""), []).append(r)
    return nhom


def _xu_ly_thieu_ma_shipto(rows_khong_ma, ship_to_hien_co, thong_ke, du_lieu, bo_nho_check):
    """Dòng SAP không có sẵn Mã Ship to (vd nhóm DECO) - theo xác nhận nghiệp vụ 16/10/2026:
    STM tự check địa chỉ khi xử lý đơn, khớp ĐÚNG 1 điểm có sẵn trong cùng Sold-to thì dùng
    luôn, không thì tự đề xuất chọn/tạo - KHÔNG phải lỗi phải báo SAP. Ở đây không có mã nào
    để neo nên PHẢI khớp CHẮC CHẮN (100% token địa chỉ) mới dám tự điền giúp, khác hẳn ngưỡng
    mờ (same_place, >= 4 token) dùng cho case đã có mã. Không khớp -> để trống."""
    for rs in _nhom_theo_dia_chi(rows_khong_ma).values():
        r0 = rs[0]
        if "O" in r0["_thieu"] or "P" in r0["_thieu"]:
            continue
        _, toks_moi = parse_address(r0["N"])
        if not toks_moi:
            continue
        ma_khop = ban_ghi_khop = None
        for sp in ship_to_hien_co:
            if not sp.get("ma"):
                continue  # ban ghi khong ma (diem moi da xac nhan truoc day) - khong tai dung
            phones_sp, toks_sp = parse_address(sp["dia_chi"])
            if phones_sp:
                continue
            if toks_moi == toks_sp:
                if not _cache_dong_bo_dia_chi(sp, du_lieu, thong_ke["canh_bao_cache_sai"], bo_nho_check):
                    continue
                if not _dung_format_dia_chi_cache(sp):
                    thong_ke["canh_bao_cache_sai"].add(
                        f"Ship-to {sp.get('ma','?')} tren STM dia chi sai format "
                        "(khong ket thuc bang 'Quan, Tinh') - bo qua, khong tai dung")
                    continue
                ma_khop, ban_ghi_khop = sp["ma"], sp
                break
        if ma_khop is None:
            # DECO - khong co ma goc SAP de giu, de trong + canh bao "Thieu Ma Ship to" de
            # khach biet (NHOM_SAP: khong tu sua Excel duoc, phai bao SAP hoac go tay neu biet).
            for r in rs:
                if "L" not in r["_thieu"]:
                    r["_thieu"].append("L")
                if LY_DO_THIEU_MA_SHIPTO not in r["_ly_do"]:
                    r["_ly_do"].append(LY_DO_THIEU_MA_SHIPTO)
            continue
        for r in rs:
            r["L"] = ma_khop
            r["N"] = ban_ghi_khop["dia_chi"]
            r["O"] = ban_ghi_khop["tinh"]
            r["P"] = ban_ghi_khop["quan"]
            for c in ("L", "N", "O", "P"):
                if c in r["_thieu"]:
                    r["_thieu"].remove(c)
            r["_ly_do"] = [ly for ly in r["_ly_do"]
                           if ly not in _LY_DO_DVHC and ly != LY_DO_THIEU_MA_SHIPTO]
        thong_ke["shipto_tai_dung"].add(ma_khop)


def xu_ly(dong, cache, du_lieu, cho_phep_tao_moi=True):
    """Sửa trực tiếp dong[i]['L'/'N'/'O'/'P'] khi tái dùng. Không khớp cache + cho_phep_tao_moi
    -> xoá trống dong[i]['L'] (để STM tự tạo điểm lúc xử lý đơn) - KHÔNG đụng gì tới `cache`.

    `du_lieu`: instance dia_chi_cu.DuLieu (hoặc fake cùng interface) - dùng
    `du_lieu.chuan_cap_stm(tinh, quan)` để validate ứng viên cache: ứng viên khớp địa chỉ
    nhưng tỉnh/quận KHÔNG đúng danh mục STM (bản ghi cũ sai/lệch ĐVHC) bị bỏ qua, không dùng
    để tái dùng - chỉ ghi cảnh báo vào `canh_bao_cache_sai`.

    `cho_phep_tao_moi=False` (Bước 1): chỉ chạy nhánh tái dùng, dòng không khớp giữ nguyên Mã
    Ship to gốc từ SAP - để khách kịp sửa file upload trước khi bấm nút 3.

    -> dict thống kê: `shipto_tai_dung` (mã Ship-to tái dùng được), `shipto_tao_moi` (mã
    Sold-to có điểm giao MỚI, Mã Ship to đã để trống), `bo_qua_thieu_dia_chi` (mã Ship-to gốc
    chưa đủ Tỉnh/Quận để xét), `canh_bao_cache_sai` (bản ghi cache sai ĐVHC, bị bỏ qua)."""
    thong_ke = {"shipto_tai_dung": set(), "shipto_tao_moi": set(),
                "canh_bao_cache_sai": set(), "bo_qua_thieu_dia_chi": set()}
    # Nhớ kết quả check "cache có đồng bộ tỉnh với địa chỉ text của nó" theo chuỗi địa
    # chỉ - không gọi suy_dvhc_tu_text lặp cho cùng 1 địa chỉ.
    bo_nho_check = {}

    theo_soldto = {}
    for r in dong:
        ma_soldto = r.get("K", "")
        if not ma_soldto:
            continue
        theo_soldto.setdefault(ma_soldto, []).append(r)

    for ma_soldto, rows in theo_soldto.items():
        ship_to_hien_co = cache["shipto"].get(ma_soldto, [])

        # Nhóm theo CẶP (Mã Ship-to gốc, Địa chỉ) chứ không chỉ theo Mã Ship-to gốc - SAP
        # có khi dùng chung 1 mã Ship-to (vd placeholder/mặc định) cho nhiều delivery có
        # địa chỉ giao hàng khác hẳn nhau (bug thật đã vấp 21/10/2026).
        theo_shipto = {}
        rows_khong_ma = []
        for r in rows:
            ma_shipto = r.get("L", "")
            if ma_shipto:
                theo_shipto.setdefault((ma_shipto, r.get("N", "")), []).append(r)
            else:
                rows_khong_ma.append(r)

        _xu_ly_thieu_ma_shipto(rows_khong_ma, ship_to_hien_co, thong_ke, du_lieu, bo_nho_check)

        for (ma_shipto, _dia_chi_nhom), rs in theo_shipto.items():
            ma_khop = None
            ban_ghi_khop = None
            _, toks_moi = parse_address(rs[0]["N"])
            for sp in ship_to_hien_co:
                if not sp.get("ma"):
                    continue  # diem moi da xac nhan truoc day, khong co ma - khong tai dung
                phones_sp, toks_sp = parse_address(sp["dia_chi"])
                if phones_sp:
                    # Địa chỉ cột N trong file upload đã tách riêng tên người/SĐT ra cột
                    # AA (xem chuan_hoa.thong_tin_nguoi_nhan) nên luôn sạch. Bản ghi cache
                    # còn lẫn SĐT trong địa chỉ là dữ liệu không sạch (tạo tay dính tên
                    # người + SĐT trên STM trước đây) - không đáng tin để so khớp dù overlap
                    # token cao đến đâu, bỏ qua hẳn không xét ứng viên này.
                    continue
                if not _du_tin_cay_de_tai_dung(toks_moi, toks_sp):
                    continue
                if du_lieu.chuan_cap_stm(sp["tinh"], sp["quan"]) is None:
                    thong_ke["canh_bao_cache_sai"].add(
                        f"Ship-to {sp['ma']} tren STM dang sai DVHC "
                        f"(tinh={sp['tinh']}, quan={sp['quan']}) - bo qua, khong dung")
                    continue
                if not _cache_dong_bo_dia_chi(sp, du_lieu, thong_ke["canh_bao_cache_sai"], bo_nho_check):
                    continue
                if not _dung_format_dia_chi_cache(sp):
                    thong_ke["canh_bao_cache_sai"].add(
                        f"Ship-to {sp['ma']} tren STM dia chi sai format "
                        "(khong ket thuc bang 'Quan, Tinh') - bo qua, khong tai dung")
                    continue
                ma_khop = sp["ma"]
                ban_ghi_khop = sp
                break

            if ma_khop:
                for r in rs:
                    r["L"] = ma_khop
                    r["N"] = ban_ghi_khop["dia_chi"]
                    r["O"] = ban_ghi_khop["tinh"]
                    r["P"] = ban_ghi_khop["quan"]
                    for c in ("N", "O", "P"):
                        if c in r["_thieu"]:
                            r["_thieu"].remove(c)
                    r["_ly_do"] = [ly for ly in r["_ly_do"] if ly not in _LY_DO_DVHC]
                thong_ke["shipto_tai_dung"].add(ma_khop)
                continue

            if not cho_phep_tao_moi:
                continue

            r0 = rs[0]
            if "O" in r0["_thieu"] or "P" in r0["_thieu"]:
                thong_ke["bo_qua_thieu_dia_chi"].add(ma_shipto)
                continue

            # Khong khop cache nhung SAP co san ma goc -> GIU NGUYEN ma do (doi huong 29/10/2026):
            # STM van tu khop dia chi/de xuat tao moi luc len don, ma Hafele cu la dau moi huu
            # ich de STM tra. Khong xoa L nua, khong canh bao - khac nhanh DECO (khong co ma de
            # neo, phai bao thieu).
            thong_ke["shipto_tao_moi"].add(ma_soldto)

    return thong_ke

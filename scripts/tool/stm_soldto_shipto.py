# -*- coding: utf-8 -*-
"""Logic quyết định tạo mới/tái dùng Sold-to/Ship-to — tra cache local (xem stm_cache.py),
KHÔNG gọi mạng. Khi khớp mờ ra 1 Ship-to có sẵn, lấy cả Tỉnh/Quận/Địa chỉ đã đăng ký trên STM
(không chỉ địa chỉ như bản cũ) - xem spec §2.5: cache Ship-to đã kiêm vai trò "bộ nhớ địa chỉ".

Luồng tách 2 bước (xem docs/superpowers/specs/2026-10-15-tach-2-buoc-soldto-shipto-design.md):
- Bước 1 (main.chay()): gọi xu_ly(..., cho_phep_tao_moi=False) - chỉ tái dùng, không đề xuất
  tạo mới, không ghi cache - để khách kịp sửa địa chỉ trong file upload trước.
- Bước 2 (main.tao_file_soldto_shipto()): gọi xu_ly(..., cho_phep_tao_moi=True) trên dữ liệu
  đọc lại từ file upload đã sửa - lúc này mới đề xuất tạo mới + sinh mã + ghi cache.
"""
import random

from .dia_chi_so_khop import parse_address, same_place
from .stm_cache import them_shipto, them_soldto

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
    hoàn toàn khác, chỉ vì trùng tên người 'Tuan' + 'Quan 1' + 'HCM')."""
    if len(toks_moi) < _SO_TOKEN_TOI_THIEU or len(toks_sp) < _SO_TOKEN_TOI_THIEU:
        return False
    return same_place(toks_moi, toks_sp)


def _so_ngau_nhien_10_chu_so():
    return f"{random.randint(0, 9_999_999_999):010d}"


def sinh_ma_shipto_moi(ma_soldto, da_dung, sinh_so=_so_ngau_nhien_10_chu_so):
    """Mã Ship-to mới = '<ma_soldto>-<10 chữ số>' (có dấu '-' phân cách), không trùng mã nào
    trong `da_dung` (tập hợp toàn bộ mã Ship-to đã biết - TOÀN HỆ THỐNG, không chỉ của 1
    Sold-to, vì mã Ship-to trên STM là duy nhất toàn hệ thống - bug thật từng gặp khi dùng mã
    ngắn dễ trùng, xem HANDOFF §0f). `sinh_so` tiêm được cho test; mặc định sinh số ngẫu nhiên
    thật."""
    for _ in range(1000):
        ma = f"{ma_soldto}-{sinh_so()}"
        if ma not in da_dung:
            return ma
    raise RuntimeError(
        f"Khong sinh duoc ma Ship-to moi khong trung cho Sold-to {ma_soldto} sau 1000 lan thu")


def _tat_ca_ma_shipto(cache):
    return {sp["ma"] for ds in cache["shipto"].values() for sp in ds}


_LY_DO_DVHC = (
    "Không xác định được Quận/Huyện", "Quận/Huyện hoặc Phường/Xã trùng tên nhiều nơi",
    "Quận không có trong danh sách STM chấp nhận", "Không xác định được Tỉnh",
    "Tỉnh độ tin cậy thấp",
)


def _nhom_theo_dia_chi(rows):
    nhom = {}
    for r in rows:
        nhom.setdefault(r.get("N", ""), []).append(r)
    return nhom


def _xu_ly_thieu_ma_shipto(rows_khong_ma, ship_to_hien_co, thong_ke):
    """Dòng SAP không có sẵn Mã Ship to (vd nhóm DECO) - theo xác nhận nghiệp vụ 16/10/2026:
    STM tự check địa chỉ khi xử lý đơn, khớp ĐÚNG 1 điểm có sẵn trong cùng Sold-to thì dùng
    luôn, không thì tự đề xuất chọn/tạo - KHÔNG phải lỗi phải báo SAP. Ở đây không có mã nào
    để neo nên PHẢI khớp CHẮC CHẮN (100% token địa chỉ) mới dám tự điền giúp, khác hẳn ngưỡng
    mờ (same_place, >= 4 token) dùng cho case đã có mã. Không khớp -> để trống, KHÔNG tự đề
    xuất tạo Ship-to mới (để STM tự lo lúc lên đơn, không cần geocode trước)."""
    for rs in _nhom_theo_dia_chi(rows_khong_ma).values():
        r0 = rs[0]
        if "O" in r0["_thieu"] or "P" in r0["_thieu"]:
            continue
        _, toks_moi = parse_address(r0["N"])
        if not toks_moi:
            continue
        ma_khop = ban_ghi_khop = None
        for sp in ship_to_hien_co:
            phones_sp, toks_sp = parse_address(sp["dia_chi"])
            if phones_sp:
                continue
            if toks_moi == toks_sp:
                ma_khop, ban_ghi_khop = sp["ma"], sp
                break
        if ma_khop is None:
            continue
        for r in rs:
            r["L"] = ma_khop
            r["N"] = ban_ghi_khop["dia_chi"]
            r["O"] = ban_ghi_khop["tinh"]
            r["P"] = ban_ghi_khop["quan"]
            for c in ("L", "N", "O", "P"):
                if c in r["_thieu"]:
                    r["_thieu"].remove(c)
            r["_ly_do"] = [ly for ly in r["_ly_do"] if ly not in _LY_DO_DVHC]
        thong_ke["shipto_tai_dung"].add(ma_khop)


def xu_ly(dong, cache, du_lieu, lay_toa_do=None, cho_phep_tao_moi=True):
    """Sửa trực tiếp dong[i]['L'/'N'/'O'/'P'] khi tái dùng; cập nhật `cache` ngay khi đề
    xuất tạo mới (để lần gọi sau trong cùng batch, hoặc lần chạy sau nếu cache được ghi lại
    ra file, không đề xuất trùng).

    `du_lieu`: instance dia_chi_cu.DuLieu (hoặc fake cùng interface) - dùng
    `du_lieu.chuan_cap_stm(tinh, quan)` để validate ứng viên cache: ứng viên khớp địa chỉ
    nhưng tỉnh/quận KHÔNG đúng danh mục STM (bản ghi cũ sai/lệch ĐVHC) bị bỏ qua, không dùng
    để tái dùng - chỉ ghi cảnh báo vào `canh_bao_cache_sai`.

    `lay_toa_do`: hàm `dia_chi -> (lat, lng) | None` (xem trackasia_api.toa_do) - BẮT BUỘC
    phải tra được toạ độ mới tạo Ship-to mới (STM import Partner không tự geocode - điểm tạo
    ra sẽ không có toạ độ nếu để trống, trải nghiệm thật đã gặp). Không tra được (None, hoặc
    không truyền `lay_toa_do`) -> KHÔNG tạo qua file soldto_shipto, để nguyên địa chỉ trong
    file upload đơn hàng cho STM tự tạo điểm (không toạ độ) lúc xử lý đơn - xem
    `bo_qua_khong_co_toa_do`.

    `cho_phep_tao_moi=False` (Bước 1): chỉ chạy nhánh tái dùng, không đề xuất tạo mới, không
    đụng cache['soldto'/'shipto'] - để khách kịp sửa file upload trước khi tạo thật ở Bước 2.

    -> dict thống kê, trong đó `can_tao_moi` là phần sẽ ghi ra file Excel upload Sold-to/Ship-to
    (xem stm_api.ghi_file_soldto_shipto)."""
    thong_ke = {"soldto_moi": set(), "shipto_moi": set(), "shipto_tai_dung": set(),
                "loi": set(), "canh_bao_cache_sai": set(), "bo_qua_thieu_dia_chi": set(),
                "bo_qua_khong_co_toa_do": set(),
                "can_tao_moi": {"soldto": [], "shipto": []}}

    theo_soldto = {}
    for r in dong:
        ma_soldto = r.get("K", "")
        if not ma_soldto:
            continue
        theo_soldto.setdefault(ma_soldto, []).append(r)

    for ma_soldto, rows in theo_soldto.items():
        # Luôn thử khớp địa chỉ với Ship-to có sẵn trước, kể cả khi dòng đang bị tô vàng vì
        # geocode không chắc (thiếu O/P) - 1 khớp đúng giải quyết luôn phần không chắc đó
        # (lấy O/P đã đăng ký trên STM). Chỉ khi KHÔNG khớp được mới xét có nên đề xuất tạo
        # mới hay không (và khi đó mới cần dữ liệu đủ tin cậy - không tạo mới từ địa chỉ
        # chưa xác định được quận/tỉnh).
        ship_to_hien_co = cache["shipto"].get(ma_soldto, [])

        theo_shipto = {}
        rows_khong_ma = []
        for r in rows:
            ma_shipto = r.get("L", "")
            if ma_shipto:
                theo_shipto.setdefault(ma_shipto, []).append(r)
            else:
                rows_khong_ma.append(r)

        _xu_ly_thieu_ma_shipto(rows_khong_ma, ship_to_hien_co, thong_ke)

        for ma_shipto, rs in theo_shipto.items():
            ma_khop = None
            ban_ghi_khop = None
            _, toks_moi = parse_address(rs[0]["N"])
            for sp in ship_to_hien_co:
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
                    r["_ly_do"] = [ly for ly in r["_ly_do"] if ly not in (
                        "Không xác định được Quận/Huyện",
                        "Quận/Huyện hoặc Phường/Xã trùng tên nhiều nơi",
                        "Quận không có trong danh sách STM chấp nhận",
                        "Không xác định được Tỉnh", "Tỉnh độ tin cậy thấp")]
                thong_ke["shipto_tai_dung"].add(ma_khop)
                continue

            if not cho_phep_tao_moi:
                continue

            r0 = rs[0]
            if "O" in r0["_thieu"] or "P" in r0["_thieu"]:
                # Không khớp được Ship-to có sẵn, và địa chỉ chưa đủ tin cậy (thiếu
                # Tỉnh/Quận) -> không đề xuất tạo mới từ địa chỉ chưa chắc này, để khách tự
                # sửa file rồi bấm lại nút 3.
                thong_ke["bo_qua_thieu_dia_chi"].add(ma_shipto)
                continue

            toa_do = lay_toa_do(r0["N"]) if lay_toa_do else None
            if toa_do is None:
                # Không tra được toạ độ (chưa cấu hình, mạng lỗi, key bị chặn, không ra kết
                # quả) -> KHÔNG tạo Ship-to qua file soldto_shipto (điểm tạo ra sẽ không có
                # toạ độ trên STM). Để nguyên địa chỉ trong file upload đơn hàng, STM sẽ tự
                # tạo điểm (không toạ độ) lúc xử lý đơn - chấp nhận được, không ép tạo bằng
                # mọi giá.
                thong_ke["bo_qua_khong_co_toa_do"].add(ma_shipto)
                continue
            lat, lng = toa_do

            ma_moi = sinh_ma_shipto_moi(ma_soldto, _tat_ca_ma_shipto(cache))
            for r in rs:
                r["L"] = ma_moi

            if ma_soldto not in cache["soldto"]:
                them_soldto(cache, ma_soldto)
                thong_ke["can_tao_moi"]["soldto"].append(ma_soldto)
                thong_ke["soldto_moi"].add(ma_soldto)

            them_shipto(cache, ma_soldto, ma_moi, r0["N"], r0["O"], r0["P"])
            thong_ke["can_tao_moi"]["shipto"].append(
                {"ma_soldto": ma_soldto, "ma_shipto": ma_moi, "dia_chi": r0["N"],
                 "tinh": r0["O"], "quan": r0["P"], "lat": lat, "lng": lng})
            thong_ke["shipto_moi"].add(ma_moi)

    return thong_ke

# scripts/tool/chuan_hoa.py
# -*- coding: utf-8 -*-
"""Dọn text địa chỉ, lọc ghi chú giao hàng, tách quận, dựng tên Ship-to.

Vì sao lọc ghi chú theo CỤM TỪ chứ không theo từ đơn:
  từ đơn dính oan tên địa danh — 'Xã Giao Thủy' có chữ "giao", 'Kho Thành Đạt'
  có chữ "kho", 'Đại lý Nhận Phát' có chữ "nhận". Đo trên data thật:
  cụm từ bắt đúng 5/5 ghi chú, cắt nhầm 0/767 giá trị địa chỉ thuần
  và 0/22 tên địa danh bẫy.
"""
import re
import unicodedata


def bo_dau(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def _chuan(s):
    return re.sub(r"\s+", " ", bo_dau(s).lower()).strip()


def la_ghi_chu(s, cum_tu):
    """True nếu ô này là ghi chú giao hàng chứ không phải địa chỉ."""
    if not isinstance(s, str) or not s.strip():
        return False
    t = _chuan(s)
    return any(_chuan(c) in t for c in cum_tu)


_TIEN_TO_DINH_LIEN = re.compile(
    r"\b(Phường|Xã|Thị trấn|Quận|Huyện|Thành phố|Thị xã|Tỉnh)(\S)")


def _tach_tien_to_dinh_lien(s):
    """SAP đôi khi export 2 field dính liền không khoảng trắng khi field trước kết
    thúc đúng bằng tiền tố cấp hành chính và field sau là tên riêng viết hoa (vd
    'PhườngBình Dương' -> 'Phường Bình Dương'). CHỈ chèn khoảng trắng khi ký tự ngay
    sau là CHỮ HOA — chữ thường theo sau có thể chỉ là trùng chữ ngẫu nhiên, không
    phải lỗi dính field, không được tách bừa. Bug thật 20/10/2026: DO.XLSX Street 4
    = 'PhườngBình Dương,' (xem CLAUDE.md)."""
    def _chen(m):
        tien_to, ky_tu = m.group(1), m.group(2)
        return f"{tien_to} {ky_tu}" if ky_tu.isupper() else m.group(0)
    return _TIEN_TO_DINH_LIEN.sub(_chen, str(s))


def don_dia_chi(s):
    """Gộp khoảng trắng, bỏ dấu phẩy lặp và dấu phẩy/gạch ở cuối."""
    t = _tach_tien_to_dinh_lien(s)
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"\s*,\s*(?=,)", "", t)      # ',,' -> ','
    t = re.sub(r"(\s*,)+\s*$", "", t)       # bỏ dấu phẩy cuối chuỗi
    t = re.sub(r"^[\s,;-]+", "", t)
    return re.sub(r"[\s,;-]+$", "", t).strip()


def ghep_dia_chi(dong, cot_dia_chi, cum_tu):
    """Ghép các cột địa chỉ của DO thành ô 'Địa chỉ'.

    KHÔNG bao giờ lấy 'c/o name' — cột đó chứa tên người nhận + SĐT
    (271/296 đơn có SĐT), và không cột địa chỉ nào khác chứa SĐT (0/296).
    """
    phan = []
    for c in cot_dia_chi:
        v = dong.get(c, "")
        if not isinstance(v, str):
            v = str(v) if v is not None else ""
        v = v.strip()
        if not v or la_ghi_chu(v, cum_tu):
            continue
        phan.append((c, v))

    # "House No." + "Street" là 1 cụm số nhà + tên đường, viết liền không phẩy. Các cột
    # sau đó (Street 2.. City) thường là phường/quận/tỉnh -> PHẢI ngăn bằng dấu phẩy, nếu
    # không dia_chi_cu.py sẽ dính nó vào mảnh trước, không tách được mảnh phường/quận để
    # suy luận (bug thật 16/10/2026: SAP không tự có dấu phẩy cuối field như field "Street"
    # ở ví dụ cũ tình cờ có — không được dựa vào may rủi đó).
    ghep = ""
    for i, (c, v) in enumerate(phan):
        if i == 0:
            ghep = v
        elif phan[i - 1][0] == "House No.":
            ghep += " " + v
        else:
            ghep += ", " + v
    return don_dia_chi(ghep)


def thong_tin_nguoi_nhan(dong, cot_dia_chi, cum_tu):
    """Phần KHÔNG phải địa chỉ: 'c/o name' (tên + SĐT) + các ghi chú bị loại khỏi địa chỉ.

    Là phần bù của ghep_dia_chi — cùng tiêu chí la_ghi_chu nên không sót, không trùng.
    """
    phan = []
    co = dong.get("c/o name", "")
    if isinstance(co, str) and co.strip():
        phan.append(co.strip())
    for c in cot_dia_chi:
        v = dong.get(c, "")
        if isinstance(v, str) and v.strip() and la_ghi_chu(v, cum_tu):
            phan.append(v.strip())
    return "; ".join(phan)


def hoa_dau_tu(s):
    """'Thành phố MỸ THO' -> 'Thành Phố Mỹ Tho'. Giữ nguyên cụm toàn số ('12')."""
    ra = []
    for w in str(s).split():
        ra.append(w if w.isdigit() else w[:1].upper() + w[1:].lower())
    return " ".join(ra)


_PREFIX_SHIP = re.compile(r"^SHIP\s*[ _]?\s*TO_|^SHIP_", re.IGNORECASE)


def ten_ship_to(name, name2):
    """Bỏ prefix SHIP_/SHIP TO_ ở đầu Name rồi nối với Name 2."""
    a = _PREFIX_SHIP.sub("", str(name or "")).strip()
    b = str(name2 or "").strip()
    return re.sub(r"\s+", " ", f"{a} {b}").strip()

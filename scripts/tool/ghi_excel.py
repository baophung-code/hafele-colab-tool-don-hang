# scripts/tool/ghi_excel.py
# -*- coding: utf-8 -*-
"""Ghi file kết quả: sheet 'Upload' (27 cột).

Từ 05/11/2026: bỏ hẳn sheet ẩn `_TenSoldTo` (gắn với luồng Sold-to/Ship-to gỡ cùng đợt —
STM tự làm việc đó qua TMS-15494). Nhưng các hàm đọc lại/ghi đè (doc_lai_upload,
ghi_de_upload, xoa_canh_bao_*) vẫn GIỮ cho Bước 3 "Quét lại sau khi sửa tay" —
khách sửa Tỉnh/Quận/Địa chỉ/Ngày/Số lượng trực tiếp trong Excel, bấm Bước 3 để tool
re-validate + xoá cảnh báo tô vàng đúng dòng đã hợp lệ."""
import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .xlsx_shared_strings import chuyen_sang_shared_strings

# Đúng thứ tự template. Cột trống giữa "Số lượng" và "Storage location" (khoá
# nội bộ "H") đã bỏ khỏi file output theo yêu cầu khách 25/09 -> không xuất hiện
# trong TEN_COT/KHOA, dù dung_bang.py vẫn còn giữ khoá "H" nội bộ (luôn rỗng).
# "Ghi chú đơn hàng" (khoá nội bộ "Z") nằm ngoài template gốc, thêm 25/09 theo
# yêu cầu khách — hiện CHƯA có nguồn data, luôn để trống (dung_bang.py không
# gán giá trị cho khoá "Z").
TEN_COT = [
    "Mã ODO", "Mã đơn hàng", "Mã đặt hàng", "Ngày gửi y/c", "Mã hàng hóa",
    "Tên hàng hóa/đơn vị tính", "Số lượng", "Storage location",
    "Mã điểm lấy hàng", "Mã Sold to", "Mã Ship to", "Tên Ship to", "Địa chỉ",
    "Tỉnh điểm giao hàng", "Quận điểm giao hàng", "Mã hình thức thanh toán",
    "Giá chuyến COD", "Khách hàng trả phí vận chuyển (Có/ Không)", "ETD", "ETA",
    "Mã HU", "Ngày bán hàng", "Ghi chú đơn hàng", "Loại vận chuyển", "Loại dịch vụ",
    "Thông tin người nhận", "Lý do cảnh báo",
]

# Z ("Ghi chú đơn hàng") chèn giữa W và X -> không dùng range chữ cái liên tục
# được nữa, phải liệt kê tay đúng thứ tự cột thật. "AB" (cuối, "Lý do cảnh báo")
# không phải khoá thường có trong dong[i] -> ghi riêng từ r["_ly_do"], xem _sheet_upload.
KHOA = [k for k in "ABCDEFGHIJKLMNOPQRSTUVW" if k != "H"] + ["Z", "X", "Y", "AA", "AB"]

VANG = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
DINH_DANG_NGAY = {"D": "dd/mm/yyyy", "T": "dd/mm/yyyy hh:mm",
                  "U": "dd/mm/yyyy hh:mm", "W": "dd/mm/yyyy hh:mm"}
RONG_COT = {"M": 38, "N": 60, "O": 20, "P": 20, "D": 13, "T": 17, "U": 17, "W": 17, "AA": 45}

# Khách không rành flow tool, chỉ cần biết: cái nào TỰ SỬA được trên Excel,
# cái nào phải BÁO LẠI vì thiếu ngay từ dữ liệu SAP gốc (tool không có gì để suy).
NHOM_SAP = "Thiếu dữ liệu SAP gốc"
NHOM_DIA_CHI = "Cần soát địa chỉ"
NHOM_KHAC = "Khác"
GIAI_THICH_NHOM = {
    NHOM_SAP: "không tự sửa trên Excel được — báo lại bộ phận SAP",
    NHOM_DIA_CHI: "có thể sửa trực tiếp trên Excel",
    NHOM_KHAC: "xem cột Chi tiết",
}
_LOAI_TO_NHOM = {
    "Thiếu ngày gửi / ETD (cột D, T)": NHOM_SAP,
    "Số lượng ≤ 0": NHOM_SAP,
    "Tỉnh SAP và tỉnh theo toạ độ khác nhau": NHOM_DIA_CHI,
    "Địa chỉ cần soát": NHOM_DIA_CHI,
    "Có nhiều Quận/Huyện cũ ứng với địa chỉ này": NHOM_DIA_CHI,
    "Không xác định được Quận/Huyện": NHOM_DIA_CHI,
    "Quận không có trong danh sách STM chấp nhận": NHOM_DIA_CHI,
    "Không xác định được Tỉnh": NHOM_DIA_CHI,
    "Tỉnh độ tin cậy thấp": NHOM_DIA_CHI,
}


def nhom_cua(loai):
    return _LOAI_TO_NHOM.get(loai, NHOM_KHAC)


def don_theo_nhom(canh_bao):
    """-> {nhóm: set(mã ODO)}, sắp xếp nhóm đông đơn nhất trước."""
    theo_nhom = {}
    for c in canh_bao:
        theo_nhom.setdefault(nhom_cua(c["loai"]), set()).add(c["ma_odo"])
    return dict(sorted(theo_nhom.items(), key=lambda kv: -len(kv[1])))


def ghi_file(duong_dan, dong):
    wb = openpyxl.Workbook()
    _sheet_upload(wb.active, dong)
    try:
        wb.save(duong_dan)
    except PermissionError:
        raise PermissionError(
            f"Không ghi được '{duong_dan}'. File đang mở trong Excel — hãy đóng lại rồi chạy lại."
        )
    finally:
        wb.close()
    chuyen_sang_shared_strings(duong_dan)


def _sheet_upload(ws, dong):
    ws.title = "Upload"
    ws.append(TEN_COT)
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.freeze_panes = "A2"

    for r in dong:
        gia_tri = [r.get(k, "") for k in KHOA[:-1]]
        chi_tiet_theo_loai = r.get("_chi_tiet", {})
        ly_do_day_du = [
            f"{loai} — {chi_tiet_theo_loai[loai]}" if loai in chi_tiet_theo_loai else loai
            for loai in r.get("_ly_do", [])
        ]
        gia_tri.append("; ".join(ly_do_day_du))
        ws.append(gia_tri)
        i = ws.max_row
        thieu = set(r.get("_thieu", []))
        for j, k in enumerate(KHOA, start=1):
            o = ws.cell(i, j)
            if k in thieu:
                o.fill = VANG
            if k in DINH_DANG_NGAY and isinstance(o.value, datetime.datetime):
                o.number_format = DINH_DANG_NGAY[k]

    for k, w in RONG_COT.items():
        ws.column_dimensions[get_column_letter(KHOA.index(k) + 1)].width = w


LOAI_THIEU_NGAY = "Thiếu ngày gửi / ETD (cột D, T)"
LOAI_SO_LUONG_AM = "Số lượng ≤ 0"


def doc_lai_upload(duong_dan):
    """Đọc lại sheet 'Upload' (có thể đã được khách sửa tay) -> list dict theo khoá
    K/L/M/N/O/P + '_hang' (số dòng Excel thật, bắt đầu từ 2) để `ghi_de_upload` biết ghi
    đè đúng ô. Dùng cho Bước 3."""
    wb = openpyxl.load_workbook(duong_dan, data_only=True)
    try:
        ws = wb["Upload"]
        ket_qua = []
        so_cot = len(KHOA)
        for i in range(2, ws.max_row + 1):
            if all(ws.cell(i, j).value in (None, "") for j in range(1, so_cot + 1)):
                continue
            d = {"_hang": i}
            for k in ("K", "L", "M", "N", "O", "P"):
                v = ws.cell(i, KHOA.index(k) + 1).value
                d[k] = "" if v is None else v
            ket_qua.append(d)
        return ket_qua
    finally:
        wb.close()


def ghi_de_upload(duong_dan, thay_doi):
    """Ghi đè đúng ô N/O/P (và L nếu truyền) của file theo list dict có '_hang' + khoá cần
    sửa. Không đụng ô khác, không đụng cột 'Lý do cảnh báo' (xoá riêng qua
    xoa_canh_bao_dia_chi)."""
    if not thay_doi:
        return
    wb = openpyxl.load_workbook(duong_dan)
    try:
        ws = wb["Upload"]
        for d in thay_doi:
            for k in ("L", "N", "O", "P"):
                if k in d:
                    ws.cell(d["_hang"], KHOA.index(k) + 1).value = d[k]
        try:
            wb.save(duong_dan)
        except PermissionError:
            raise PermissionError(
                f"Không ghi được '{duong_dan}'. File đang mở trong Excel — hãy đóng lại rồi "
                "bấm lại Bước 3."
            )
    finally:
        wb.close()
    chuyen_sang_shared_strings(duong_dan)


def doc_lai_canh_bao_sap(duong_dan):
    """Đọc lại cột D, T, G, AB sheet 'Upload' — Bước 3 dùng để biết dòng nào khách đã tự
    điền tay ngày gửi/ETD hoặc sửa số lượng."""
    wb = openpyxl.load_workbook(duong_dan, data_only=True)
    try:
        ws = wb["Upload"]
        ket_qua = []
        for i in range(2, ws.max_row + 1):
            d = {"_hang": i, "AB": ws.cell(i, KHOA.index("AB") + 1).value or ""}
            for k in ("D", "T", "G"):
                d[k] = ws.cell(i, KHOA.index(k) + 1).value
            ket_qua.append(d)
        return ket_qua
    finally:
        wb.close()


def _xoa_khoi_cot_ab(ws, hang, giu_lai):
    """giu_lai(loai) -> bool. Giữ lại các mục cảnh báo thoả điều kiện, xoá phần còn lại."""
    col_ab = KHOA.index("AB") + 1
    o = ws.cell(hang, col_ab)
    if not o.value:
        return
    con_lai = [p for p in str(o.value).split("; ")
               if p and giu_lai(p.split(" — ")[0])]
    o.value = "; ".join(con_lai)


def xoa_canh_bao_sap_goc(duong_dan, hang_da_sua_ngay, hang_da_sua_so_luong):
    """Xoá cảnh báo Thiếu ngày gửi/ETD và Số lượng ≤ 0 khỏi các dòng khách đã tự điền tay
    hợp lệ trực tiếp trong Excel + bỏ tô vàng đúng cột đã sửa, không đụng cảnh báo khác."""
    if not hang_da_sua_ngay and not hang_da_sua_so_luong:
        return
    wb = openpyxl.load_workbook(duong_dan)
    try:
        ws = wb["Upload"]
        for hang in hang_da_sua_ngay:
            _xoa_khoi_cot_ab(ws, hang, lambda loai: loai != LOAI_THIEU_NGAY)
            for k in ("D", "T"):
                ws.cell(hang, KHOA.index(k) + 1).fill = PatternFill(fill_type=None)
        for hang in hang_da_sua_so_luong:
            _xoa_khoi_cot_ab(ws, hang, lambda loai: loai != LOAI_SO_LUONG_AM)
            ws.cell(hang, KHOA.index("G") + 1).fill = PatternFill(fill_type=None)
        try:
            wb.save(duong_dan)
        except PermissionError:
            raise PermissionError(
                f"Không ghi được '{duong_dan}'. File đang mở trong Excel — hãy đóng lại rồi "
                "bấm lại Bước 3."
            )
    finally:
        wb.close()
    chuyen_sang_shared_strings(duong_dan)


def xoa_canh_bao_dia_chi(duong_dan, cac_hang):
    """Xoá mọi cảnh báo thuộc NHOM_DIA_CHI khỏi cột AB + bỏ tô vàng N/O/P của các dòng đã
    quét lại hợp lệ ở Bước 3. Không đụng nhóm khác (vd Thiếu Mã đơn hàng)."""
    if not cac_hang:
        return
    wb = openpyxl.load_workbook(duong_dan)
    try:
        ws = wb["Upload"]
        for hang in cac_hang:
            _xoa_khoi_cot_ab(ws, hang, lambda loai: nhom_cua(loai) != NHOM_DIA_CHI)
            for k in ("N", "O", "P"):
                ws.cell(hang, KHOA.index(k) + 1).fill = PatternFill(fill_type=None)
        try:
            wb.save(duong_dan)
        except PermissionError:
            raise PermissionError(
                f"Không ghi được '{duong_dan}'. File đang mở trong Excel — hãy đóng lại rồi "
                "bấm lại Bước 3."
            )
    finally:
        wb.close()
    chuyen_sang_shared_strings(duong_dan)

# -*- coding: utf-8 -*-
"""Layout + hàm ghi file Excel 'upload Sold-to/Ship-to' — khách/admin tự import qua UI Import
Excel có sẵn của STM. Tool không còn gọi API tạo mới (chốt 14/10/2026, xem spec)."""
import openpyxl

COT_HEADER_SOLDTO_SHIPTO = [
    "STT", "Loại", "Mã đối tác hệ thống", "Mã đối tác sử dụng", "Tên đối tác hệ thống",
    "Nhóm NPP / địa chỉ", "Mã địa chỉ hệ thống", "Tên địa chỉ hệ thống", "Mã địa chỉ sử dụng",
    "Tên địa chỉ sử dụng", "Địa chỉ", "Tỉnh thành", "Quận/Huyện", "Phường/Xã", "SĐT", "Fax",
    "Email", "Ghi chú 1", "Ghi chú 2", "Ghi chú 3", "Loại Điểm", "Mã khu vực", "Kinh độ",
    "Vĩ độ", "XK hạn lưu container", "XK hạn giữ rỗng", "NK hạn lưu container",
    "NK hạn giữ rỗng", "TG đóng hàng container", "TG dỡ hàng container",
    "TG đóng hàng phân phối", "TG dỡ hàng phân phối", "TG bốc hàng mỗi tấn phân phối",
    "TG dỡ hàng mỗi tấn phân phối", "TG bốc hàng mỗi khối phân phối",
    "TG dỡ hàng mỗi khối phân phối", "Người đại diện", "Chức vụ", "Mã số thuế", "Website",
    "Người nhận", "Di động", "CCCD", "Giờ mở cửa", "Giờ đóng cửa", "Trọng tải xe tối thiểu",
    "Trọng tải xe tối đa", "Mã xác nhận", "Mẫu thông báo gửi SMS", "Loại P/T ưu tiên",
    "Ghi chú 1 (khách)", "Số pallet tối đa", "KM SMS", "Bán kính nhận hàng (mét)",
    "Kho trả về", "Khu vực đội xe", "Mã đơn hàng đầu tiên", "ETD đơn hàng đầu tiên",
]


def _dong_soldto(stt, ma, ten=""):
    ten = ten or ma  # khong tra duoc ten that (vd SAP thieu "Sold-to party") -> dung tam ma
    return [stt, "Distributor", ma, ma, ten, "", ma, ma, ma, ma, "", "", "", ""] + \
        [""] * (len(COT_HEADER_SOLDTO_SHIPTO) - 14)


def _dong_shipto(stt, ma_soldto, ma_shipto, dia_chi, tinh, quan, ten="", lat="", lng=""):
    ten = ten or ma_shipto  # tuong tu _dong_soldto - thieu ten that thi tam dung ma
    dong = [""] * len(COT_HEADER_SOLDTO_SHIPTO)
    dong[0], dong[1] = stt, "Distributor"
    dong[2] = dong[3] = dong[4] = ma_soldto
    dong[6] = ma_shipto
    dong[7] = dong[9] = ten
    dong[8] = ma_shipto
    dong[10], dong[11], dong[12] = dia_chi, tinh, quan
    dong[22] = lng  # "Kinh độ" = longitude
    dong[23] = lat  # "Vĩ độ" = latitude
    return dong


def ghi_file_soldto_shipto(duong_dan, can_tao_moi):
    """can_tao_moi: {"soldto": [ma, ...], "ten_soldto": {ma: ten}, "shipto": [{"ma_soldto",
    "ma_shipto","dia_chi","tinh","quan","ten"}, ...]}. "ten"/"ten_soldto" lấy từ SAP (xem
    dung_bang.py "_ten_soldto" và "M") - rỗng hoặc thiếu hẳn khoá (file can_tao_moi cũ chưa
    có tên) thì tạm dùng mã, không crash. Không ghi file nếu không có gì cần tạo."""
    if not can_tao_moi["soldto"] and not can_tao_moi["shipto"]:
        return

    ten_soldto = can_tao_moi.get("ten_soldto", {})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(COT_HEADER_SOLDTO_SHIPTO)

    stt = 1
    for ma in can_tao_moi["soldto"]:
        ws.append(_dong_soldto(stt, ma, ten_soldto.get(ma, "")))
        stt += 1
    for sp in can_tao_moi["shipto"]:
        ws.append(_dong_shipto(stt, sp["ma_soldto"], sp["ma_shipto"], sp["dia_chi"],
                                sp["tinh"], sp["quan"], sp.get("ten", ""),
                                sp.get("lat", ""), sp.get("lng", "")))
        stt += 1

    wb.save(duong_dan)
    wb.close()

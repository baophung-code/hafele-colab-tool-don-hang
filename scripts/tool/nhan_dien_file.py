# scripts/tool/nhan_dien_file.py
# -*- coding: utf-8 -*-
"""Nhận diện file SAP bằng CHỮ KÝ CỘT, không dựa vào tên file.

Khách hay đổi tên file ('DO thang 8.xlsx'), nên không thể khớp theo tên.
Mỗi chữ ký là tập cột phải có ĐỦ. 'Delivery' có ở cả DO lẫn Item nên chữ ký
phải kèm cột đặc trưng riêng để hai file không lẫn nhau.
"""
from pathlib import Path

from .doc_excel import doc_bang

CHU_KY = {
    "DO": {"Delivery", "Handling Unit", "Sold-to pt", "Ac.GI date"},
    "Item": {"Delivery", "Material", "Delivery quantity", "Base Unit of Measure"},
    "SO": {"ERP Document", "Planned Pick Date", "Actual Time of Goods Issue"},
    "UOM": {"Shipment number", "Ship-To"},
}


def loai_file(ten_cot):
    """-> 'DO' | 'Item' | 'SO' | 'UOM' | None"""
    co = set(ten_cot)
    for loai, can in CHU_KY.items():
        if can <= co:
            return loai
    return None


def nhan_dien_thu_muc(thu_muc):
    """Quét thư mục -> (tim_thay: dict[loai, Path], bo_qua: list[Path]).

    Hai file cùng loại -> lấy file sửa gần nhất.
    """
    tim_thay, bo_qua = {}, []
    for p in sorted(Path(thu_muc).iterdir()):
        if not p.is_file() or p.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        if p.name.startswith("~$"):        # file tạm Excel mở ra khi đang edit
            continue
        try:
            ten_cot, _ = doc_bang(p)
        except Exception:
            bo_qua.append(p)
            continue
        loai = loai_file(ten_cot)
        if loai is None:
            bo_qua.append(p)
        elif loai not in tim_thay or p.stat().st_mtime > tim_thay[loai].stat().st_mtime:
            tim_thay[loai] = p
    return tim_thay, bo_qua

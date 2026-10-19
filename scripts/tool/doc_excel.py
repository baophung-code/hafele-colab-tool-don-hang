# scripts/tool/doc_excel.py
# -*- coding: utf-8 -*-
"""Đọc file .XLSX do SAP xuất ra.

Ba cái bẫy của file SAP (đã đo trên data thật):
  - header dính khoảng trắng thừa
  - tên cột trùng nhau: DO có 2 cột 'WUn' + 2 cột 'VUn', UOM có 2 cột 'Route'
  - ô rỗng trả về None chứ không phải ''
"""
import os

import openpyxl


def doc_bang(duong_dan):
    """Đọc sheet đầu tiên -> (ten_cot, dong).

    ten_cot: list[str], cột trùng tên được đánh số '__2', '__3'...
    dong:    list[dict], bỏ các dòng rỗng hoàn toàn.
    """
    if not os.path.exists(duong_dan):
        raise FileNotFoundError(f"Không thấy file: {duong_dan}")

    wb = openpyxl.load_workbook(duong_dan, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        cac_dong = ws.iter_rows(values_only=True)
        try:
            header = next(cac_dong)
        except StopIteration:
            return [], []

        ten_cot = _dat_ten_cot(header)
        ket_qua = []
        for gia_tri in cac_dong:
            if all(v is None or v == "" for v in gia_tri):
                continue
            d = {}
            for i, ten in enumerate(ten_cot):
                v = gia_tri[i] if i < len(gia_tri) else None
                d[ten] = "" if v is None else v
            ket_qua.append(d)
        return ten_cot, ket_qua
    finally:
        wb.close()


def _dat_ten_cot(header):
    ten_cot, da_dung = [], {}
    for i, raw in enumerate(header):
        ten = str(raw).strip() if raw is not None else ""
        if not ten:
            ten = f"cot_{i}"
        if ten in da_dung:
            da_dung[ten] += 1
            ten = f"{ten}__{da_dung[ten]}"
        else:
            da_dung[ten] = 1
        ten_cot.append(ten)
    return ten_cot

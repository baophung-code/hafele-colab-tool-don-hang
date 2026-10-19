# scripts/tool/cau_hinh.py
# -*- coding: utf-8 -*-
"""Đọc config.ini và bang_ma_tinh.csv.

Nguyên tắc: file thiếu / gõ sai -> dùng mặc định + ghi vào .canh_bao, KHÔNG crash.
Khách sửa nhầm config mà tool chết là hỏng cả buổi làm việc của họ.
"""
import configparser
import csv
import os

CUM_TU_MAC_DINH = [
    "gọi khách", "xác nhận", "nhận hàng", "lấy hàng", "xuống kho",
    "trước khi giao", "liên hệ", "vui lòng", "khách book", "báo trước",
    "lh:", "lưu ý", "ghi chú",
]

COT_DIA_CHI_MAC_DINH = [
    "House No.", "Street", "Street 2", "Street 3", "Street 4", "Street 5", "City",
]

# Key dùng chung công khai trên trang playground của Track Asia (xác nhận gọi thật được
# 15/10/2026) - chỉ dùng ở khối lượng nhỏ (tra toạ độ khi tạo Ship-to MỚI, không phải toàn bộ
# địa chỉ), rủi ro bị chặn/giới hạn không báo trước. Khách có key riêng thì sửa trong
# config.ini, mục [trackasia].
KEY_TRACKASIA_MAC_DINH = "public_key"


class CauHinh:
    def __init__(self):
        self.etd_cong_gio = 8
        self.nguon_ngay = "Planned Pick Date"
        self.cum_tu = list(CUM_TU_MAC_DINH)
        self.cot_dia_chi = list(COT_DIA_CHI_MAC_DINH)
        self.canh_bao = []
        self.bat_stm = False
        self.stm_base_url = "https://api-stm-prod.smartlogvn.com"
        self.stm_domain = "hafele.smartlogvn.com"
        self.stm_functionid = "37"
        self.stm_customerid = 7
        self.trackasia_key = KEY_TRACKASIA_MAC_DINH
        self.go_mo_ho_quan = True


def doc_config(duong_dan):
    c = CauHinh()
    if not os.path.exists(duong_dan):
        return c

    p = configparser.ConfigParser()
    try:
        p.read(duong_dan, encoding="utf-8")
    except Exception as e:
        c.canh_bao.append(f"config.ini đọc không được ({e}). Dùng toàn bộ giá trị mặc định.")
        return c

    def so_nguyen(section, khoa, mac_dinh):
        if not p.has_option(section, khoa):
            return mac_dinh
        raw = p.get(section, khoa).strip()
        try:
            return int(raw)
        except ValueError:
            c.canh_bao.append(f"config.ini: {khoa} = '{raw}' không phải số. Dùng mặc định {mac_dinh}.")
            return mac_dinh

    c.etd_cong_gio = so_nguyen("gio", "etd_cong_gio", 8)

    if p.has_option("ngay", "nguon"):
        v = p.get("ngay", "nguon").strip()
        if v:
            c.nguon_ngay = v

    if p.has_option("ghi_chu", "cum_tu"):
        ds = [x.strip() for x in p.get("ghi_chu", "cum_tu").splitlines() if x.strip()]
        if ds:
            c.cum_tu = ds

    if p.has_option("dia_chi", "cot"):
        ds = [x.strip() for x in p.get("dia_chi", "cot").split(",") if x.strip()]
        if ds:
            c.cot_dia_chi = ds

    if p.has_section("stm"):
        v = p.get("stm", "bat", fallback="khong").strip().lower()
        c.bat_stm = v in ("co", "1", "true", "yes")
        c.stm_base_url = p.get("stm", "base_url", fallback=c.stm_base_url).strip()
        c.stm_domain = p.get("stm", "d", fallback=c.stm_domain).strip()
        c.stm_functionid = p.get("stm", "functionid", fallback=c.stm_functionid).strip()
        c.stm_customerid = so_nguyen("stm", "customerid", c.stm_customerid)

    if p.has_section("trackasia"):
        v = p.get("trackasia", "key", fallback=c.trackasia_key).strip()
        if v:
            c.trackasia_key = v
        v2 = p.get("trackasia", "go_mo_ho_quan", fallback="co").strip().lower()
        c.go_mo_ho_quan = v2 in ("co", "1", "true", "yes")

    return c


def doc_bang_ma_tinh(duong_dan):
    """-> dict[ma_rg] = (ten_tinh, do_tin_cay). Thiếu file / đọc lỗi (VD sai encoding) -> {}."""
    if not os.path.exists(duong_dan):
        return {}
    bang = {}
    try:
        with open(duong_dan, encoding="utf-8-sig", newline="") as f:
            for d in csv.DictReader(f):
                rg = (d.get("rg") or "").strip()
                ten = (d.get("ten_tinh") or "").strip()
                if rg and ten:
                    bang[rg] = (ten, (d.get("do_tin_cay") or "cao").strip())
    except Exception:
        return {}
    return bang

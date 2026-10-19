# -*- coding: utf-8 -*-
"""Gọi Track Asia textsearch để lấy toạ độ (lat/lng) cho 1 địa chỉ.

CHỈ dùng khi tạo Ship-to MỚI trên STM (xem main.tao_file_soldto_shipto) - không dùng cho toàn
bộ địa chỉ như Goong trước đây (14/10/2026 đã bỏ Goong vì khối lượng gọi quá lớn). Ở đây khối
lượng nhỏ (chỉ điểm giao mới, không phải mọi dòng đơn) nên chấp nhận gọi mạng lại.

Không hàm nào được ném exception - lỗi mạng/parse -> trả None, để tầng trên tự quyết định bỏ
qua (xem stm_soldto_shipto.xu_ly, tham số lay_toa_do).
"""
import json
import urllib.parse
import urllib.request

_URL = "https://maps.track-asia.com/api/v2/place/textsearch/json"


def toa_do(dia_chi, key, timeout=10, mo_url=urllib.request.urlopen):
    """-> (lat, lng) hoặc None nếu không tra được (mạng lỗi, key bị chặn, không có kết quả,
    response không parse được). `mo_url` tiêm được cho test, mặc định gọi mạng thật."""
    query = urllib.parse.urlencode({
        "key": key, "query": dia_chi, "radius": 5000,
        "new_admin": "true", "include_old_admin": "true",
    })
    try:
        with mo_url(f"{_URL}?{query}", timeout=timeout) as r:
            kq = json.loads(r.read().decode("utf-8"))
        loc = kq["results"][0]["geometry"]["location"]
        return loc["lat"], loc["lng"]
    except Exception:
        return None


def dvhc_cu(dia_chi, key, timeout=10, mo_url=urllib.request.urlopen):
    """-> (tinh_cu, quan_cu) từ `old_address_components` của cùng API trên, hoặc None.

    CHỈ dùng làm TRỌNG TÀI giữa các ứng viên đã suy ra sẵn bằng chữ khi mơ hồ (xem
    dia_chi_cu.py:chon_ung_vien_bang_geocode) — KHÔNG tự ý tin thẳng kết quả geocode
    (đã cấm từ nấc A2, 7,3% sai hết dù gắn độ tin cậy CAO, xem CLAUDE.md). Không hàm nào
    được ném exception - lỗi mạng/parse/thiếu field -> trả None."""
    query = urllib.parse.urlencode({
        "key": key, "query": dia_chi, "radius": 5000,
        "new_admin": "true", "include_old_admin": "true",
    })
    try:
        with mo_url(f"{_URL}?{query}", timeout=timeout) as r:
            kq = json.loads(r.read().decode("utf-8"))
        cac_mang = kq["results"][0]["old_address_components"]
        tinh = quan = None
        for c in cac_mang:
            if "administrative_area_level_1" in c["types"]:
                tinh = c["long_name"]
            elif "administrative_area_level_2" in c["types"]:
                quan = c["long_name"]
        if tinh and quan:
            return tinh, quan
        return None
    except Exception:
        return None

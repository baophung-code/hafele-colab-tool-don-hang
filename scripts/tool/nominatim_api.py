# -*- coding: utf-8 -*-
"""Gọi Nominatim (OpenStreetMap) để lấy toạ độ (lat/lng) cho 1 địa chỉ.

DỰ PHÒNG khi Track Asia lỗi/bị chặn (xem main.tao_file_soldto_shipto) - không cần đăng ký/API
key, nhưng chính sách dùng công bằng của OSM Foundation giới hạn RẤT khắt khe: tối đa 1
request/giây, bắt buộc có User-Agent mô tả rõ ứng dụng (thiếu là bị từ chối/chặn IP). Chỉ nên
gọi khi thật sự cần (số lượng nhỏ, đã thử Track Asia trước).

Không hàm nào được ném exception - lỗi mạng/parse -> trả None.
"""
import json
import time
import urllib.parse
import urllib.request

_URL = "https://nominatim.openstreetmap.org/search"
_USER_AGENT = "HafeleSTMTool/1.0 (noi bo Hafele - tao file upload don hang STM)"


def toa_do(dia_chi, timeout=10, mo_url=urllib.request.urlopen, cho_doi=time.sleep):
    """-> (lat, lng) hoặc None nếu không tra được (mạng lỗi, không có kết quả, response không
    parse được). `mo_url`/`cho_doi` tiêm được cho test; mặc định gọi mạng thật + chờ 1 giây
    trước mỗi lần gọi để tôn trọng giới hạn 1 request/giây."""
    cho_doi(1)
    query = urllib.parse.urlencode({
        "q": dia_chi, "format": "json", "limit": 1, "countrycodes": "vn",
    })
    req = urllib.request.Request(f"{_URL}?{query}", headers={"User-Agent": _USER_AGENT})
    try:
        with mo_url(req, timeout=timeout) as r:
            kq = json.loads(r.read().decode("utf-8"))
        return float(kq[0]["lat"]), float(kq[0]["lon"])
    except Exception:
        return None

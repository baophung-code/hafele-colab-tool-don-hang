# -*- coding: utf-8 -*-
"""Cache local Sold-to/Ship-to đã có trên STM — thay cho việc tool tự gọi API mỗi lần chạy.
Seed 1 lần bằng dữ liệu thật (xem scripts/seed_stm_cache.py), từ đó tool chỉ đọc/ghi file này.
"""
import json


def doc_cache(duong_dan):
    """-> dict {"soldto": [...], "shipto": {ma_soldto: [{"ma","dia_chi","tinh","quan"}, ...]}}.
    Thiếu file / file hỏng -> cache rỗng, KHÔNG crash."""
    try:
        with open(duong_dan, encoding="utf-8") as f:
            kq = json.load(f)
    except (OSError, ValueError):
        return {"soldto": [], "shipto": {}}
    kq.setdefault("soldto", [])
    kq.setdefault("shipto", {})
    return kq


def ghi_cache(duong_dan, cache):
    with open(duong_dan, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def them_diem_moi_khong_ma(cache, ma_soldto, dia_chi, tinh, quan):
    """Ghi 1 địa chỉ MỚI (Mã Ship to đã để trống trong file upload đơn hàng, xem
    stm_soldto_shipto.xu_ly) vào cache - CHẠY SAU KHI khách xác nhận đã upload thật lên STM
    (xem main.xac_nhan_da_upload_don_hang), không ghi lạc quan ngay lúc xử lý. Bản ghi
    KHÔNG có 'ma' thật (STM tự sinh mã lúc xử lý đơn, tool không biết) - 'ma' để rỗng, và
    stm_soldto_shipto.xu_ly() sẽ KHÔNG BAO GIỜ coi bản ghi rỗng mã là ứng viên để tái dùng
    (xem các chỗ `if not sp.get("ma"): continue`). Không thêm trùng nếu đã có bản ghi cùng
    địa chỉ (so khớp nguyên văn, địa chỉ đưa vào đây đã được main.py chuẩn hoá/so khớp)."""
    ds = cache["shipto"].setdefault(ma_soldto, [])
    if any(ban_ghi["dia_chi"] == dia_chi for ban_ghi in ds):
        return
    ds.append({"ma": "", "dia_chi": dia_chi, "tinh": tinh, "quan": quan})

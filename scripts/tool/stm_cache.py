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


def them_soldto(cache, ma_soldto):
    if ma_soldto not in cache["soldto"]:
        cache["soldto"].append(ma_soldto)


def them_shipto(cache, ma_soldto, ma_shipto, dia_chi, tinh, quan):
    ds = cache["shipto"].setdefault(ma_soldto, [])
    if any(ban_ghi["ma"] == ma_shipto for ban_ghi in ds):
        return
    ds.append({"ma": ma_shipto, "dia_chi": dia_chi, "tinh": tinh, "quan": quan})

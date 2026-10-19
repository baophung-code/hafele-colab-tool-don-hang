# scripts/tool/gui_logic.py
# -*- coding: utf-8 -*-
"""Logic thuần cho GUI: chỉ đụng filesystem, KHÔNG import tkinter.

Tách riêng để test được mà không cần màn hình (Tk cần display, CI/sandbox
không có sẽ lỗi). gui.py mới là nơi dựng cửa sổ, gọi các hàm ở đây.
"""
import shutil
from pathlib import Path


def chuan_bi_input(thu_muc_vao, duong_dan_file):
    """Xoá sạch file cũ trong thu_muc_vao rồi copy các file khách vừa chọn vào.

    Tool nhận diện DO/Item/SO/UOM bằng chữ ký cột (xem nhan_dien_file.py),
    không theo tên file -> copy nguyên tên là đủ, không cần đổi tên.
    -> số file đã copy.
    """
    thu_muc_vao = Path(thu_muc_vao)
    thu_muc_vao.mkdir(parents=True, exist_ok=True)
    for p in thu_muc_vao.iterdir():
        if p.is_file():
            p.unlink()
    n = 0
    for f in duong_dan_file:
        shutil.copy(f, thu_muc_vao / Path(f).name)
        n += 1
    return n

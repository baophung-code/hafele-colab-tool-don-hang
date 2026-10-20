# scripts/tool/main.py
# -*- coding: utf-8 -*-
"""Điểm vào của tool. Khách double-click CHAY_TOOL.exe là chạy hàm main()."""
import datetime
import os
import subprocess
import sys
import traceback
from pathlib import Path

from .cau_hinh import doc_bang_ma_tinh, doc_config
from .dia_chi_cu import DuLieu, tra_dia_chi
from .doc_excel import doc_bang
from .dung_bang import dung_bang
from .ghi_excel import GIAI_THICH_NHOM, don_theo_nhom, doc_lai_upload, ghi_de_upload, ghi_file, nhom_cua
from .nhan_dien_file import nhan_dien_thu_muc
from . import nominatim_api, stm_api, stm_cache, stm_soldto_shipto, trackasia_api


_GIO_VN = datetime.timezone(datetime.timedelta(hours=7))


def _gio_vn_hien_tai():
    """Giờ Việt Nam (UTC+7) dùng để đặt tên file — máy Windows khách để sẵn giờ VN
    nên không khác gì trước đây, nhưng máy chủ Colab mặc định chạy UTC, lệch 7
    tiếng nếu dùng datetime.now() thường (bug thật 19/10/2026, tên file tạo trên
    Colab hiện giờ sáng dù tạo lúc chiều)."""
    return datetime.datetime.now(_GIO_VN)


class LoiNguoiDung(Exception):
    """Lỗi khách tự sửa được — in gọn, không in traceback."""


def _ghi_log_canh_bao(thu_muc_ra, ten_file, dong_log):
    """Ghi các dòng cảnh báo CHI TIẾT (vốn in hàng loạt ra console, dễ gây hoang mang khi
    nhiều) ra 1 file .txt riêng cạnh file upload — console chỉ in số lượng tóm tắt, khách
    cần xem chi tiết thì mở file này. -> Path file log, hoặc None nếu không có gì để ghi."""
    if not dong_log:
        return None
    duong_dan = Path(thu_muc_ra) / ten_file
    duong_dan.write_text("\n".join(dong_log), encoding="utf-8")
    return duong_dan


def thu_muc_goc():
    """Chỗ đặt exe (lúc đóng gói) hoặc gốc project (lúc chạy bằng python)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]


def thu_muc_du_lieu():
    """Chỗ đặt crosswalk/danh mục/bảng tra cứu + cache Sold-to/Ship-to.

    Đóng gói exe -> nằm ở '<cạnh exe>/du_lieu' (build_exe.py chép vào đó).
    Chạy bằng python -> nằm ở 'scripts/tool/du_lieu'.
    """
    goc = thu_muc_goc()
    canh_exe = goc / "du_lieu"
    if (canh_exe / "crosswalk_2025.csv").exists():
        return str(canh_exe)
    return str(Path(__file__).resolve().parent / "du_lieu")


def _nap_du_lieu_dia_chi(duong_dan_du_lieu):
    """-> DuLieu hoặc None nếu thiếu/lỗi 1 trong 3 file tham chiếu."""
    d = Path(duong_dan_du_lieu)
    try:
        with open(d / "crosswalk_2025.csv", encoding="utf-8-sig") as f1, \
             open(d / "danh_muc_stm.csv", encoding="utf-8-sig") as f2, \
             open(d / "alias_ten.csv", encoding="utf-8-sig") as f3:
            return DuLieu(f1, f2, f3)
    except Exception:
        return None


def _tao_lay_dvhc_cu_geocode(cfg):
    """-> hàm dia_chi -> (tinh_cu, quan_cu) | None, CHỈ dùng để phân xử khi suy chữ mơ hồ
    (xem dia_chi_cu.tra_dia_chi, tham số lay_dvhc_cu_tu_geocode). Tắt bằng
    config.ini [trackasia] go_mo_ho_quan = khong."""
    if not cfg.go_mo_ho_quan:
        return None
    return lambda dia_chi: trackasia_api.dvhc_cu(dia_chi, cfg.trackasia_key)


def tao_tra_dc(duong_dan_du_lieu, cfg=None):
    """-> hàm tra_dc cho dung_bang, hoặc None nếu thiếu dữ liệu tham chiếu."""
    du_lieu = _nap_du_lieu_dia_chi(duong_dan_du_lieu)
    if du_lieu is None:
        print("  [!] Không đọc được dữ liệu địa chỉ. Cột Tỉnh/Quận sẽ để trống.")
        return None
    lay_geocode = _tao_lay_dvhc_cu_geocode(cfg) if cfg is not None else None
    return lambda dia_chi, tinh: tra_dia_chi(dia_chi, tinh, du_lieu, lay_geocode)


def chay(thu_muc_vao, thu_muc_ra, duong_dan_config, duong_dan_bang_tinh):
    thu_muc_vao, thu_muc_ra = Path(thu_muc_vao), Path(thu_muc_ra)
    if not thu_muc_vao.is_dir():
        raise LoiNguoiDung(
            f"Không thấy folder input tại:\n  {thu_muc_vao}\n"
            "Hãy tạo folder 'input' cạnh CHAY_TOOL.exe rồi copy 4 file SAP vào."
        )
    thu_muc_ra.mkdir(parents=True, exist_ok=True)

    cfg = doc_config(duong_dan_config)
    for m in cfg.canh_bao:
        print(f"  [!] {m}")
    bang_tinh = doc_bang_ma_tinh(duong_dan_bang_tinh)
    if not bang_tinh:
        print("  [!] Không đọc được bang_ma_tinh.csv — cột Tỉnh sẽ để trống.")

    print("\n[1/4] Doc file input")
    tim_thay, bo_qua = nhan_dien_thu_muc(thu_muc_vao)
    for p in bo_qua:
        print(f"  [!] Bỏ qua '{p.name}' — không nhận diện được.")

    if not tim_thay and not bo_qua:
        raise LoiNguoiDung(
            f"Folder input chưa có file nào:\n  {thu_muc_vao}\n"
            "Hãy copy 4 file SAP (DO, Item, SO, UOM) vào rồi chạy lại. Xem HUONG_DAN.html."
        )
    thieu = [x for x in ("DO", "Item", "SO", "UOM") if x not in tim_thay]
    if thieu:
        raise LoiNguoiDung(
            f"Thiếu file {', '.join(thieu)}.\n"
            "Hãy copy đủ 4 file SAP (DO, Item, SO, UOM) vào folder input rồi chạy lại."
        )

    for loai, p in tim_thay.items():
        print(f"  - {loai}: {p.name}")

    do_ten_cot, do_dong = doc_bang(tim_thay["DO"])
    _, item_dong = doc_bang(tim_thay["Item"])
    _, so_dong = doc_bang(tim_thay["SO"])
    _, uom_dong = doc_bang(tim_thay["UOM"])

    print("\n[2/4] Dia chi (DVHC cu)")
    tra_dc = tao_tra_dc(thu_muc_du_lieu(), cfg)
    if tra_dc is None:
        tra_dc = lambda dia_chi, tinh: {"N": dia_chi, "O": tinh, "P": "",
                                         "do_tin": "THAP", "ly_do": "khong_co_quan"}
    dong, canh_bao = dung_bang(do_ten_cot, do_dong, item_dong, so_dong, uom_dong,
                               cfg, bang_tinh, tra_dc)

    if cfg.bat_stm:
        print("\n[3/4] Sold-to / Ship-to tren STM (tai dung, chua tao moi)")
        du_lieu_dia_chi = _nap_du_lieu_dia_chi(thu_muc_du_lieu())
        if du_lieu_dia_chi is None:
            print("  [!] Khong doc duoc du lieu dia chi - bo qua buoc tai dung Ship-to.")
        else:
            duong_dan_cache = Path(thu_muc_du_lieu()) / "stm_cache.json"
            cache = stm_cache.doc_cache(duong_dan_cache)
            thong_ke_stm = stm_soldto_shipto.xu_ly(dong, cache, du_lieu_dia_chi,
                                                     cho_phep_tao_moi=False)
            if thong_ke_stm["shipto_tai_dung"]:
                print(f"  - Tai dung (da co san): {len(thong_ke_stm['shipto_tai_dung'])} Ship-to")
            canh_bao_stm = sorted(thong_ke_stm["canh_bao_cache_sai"])
            if canh_bao_stm:
                duong_dan_log = _ghi_log_canh_bao(
                    thu_muc_ra, f"canhbao_shipto_{_gio_vn_hien_tai():%Y%m%d_%H%M%S}.txt",
                    canh_bao_stm)
                print(f"  [!] {len(canh_bao_stm)} Ship-to trong cache bi bo qua do sai DVHC "
                      f"- xem chi tiet: {duong_dan_log.name}")
            print("  - Da doi chieu voi cache co san; dia chi nao con can tao moi se duoc "
                  "xu ly o buoc tao Sold-to/Ship-to tiep theo.")
            # xu_ly() co the da go bot ly_do (khop cache theo dia chi, ke ca nhom thieu han
            # Ma Ship to) truc tiep tren dong[i]["_ly_do"] - canh_bao (tra ve tu dung_bang)
            # khong tu dong cap nhat theo, phai loc lai o day de bang tong ket cuoi khong
            # con nhac ten canh bao da duoc go.
            con_lai = {(r["A"], ly) for r in dong for ly in r["_ly_do"]}
            canh_bao = [c for c in canh_bao if (c["ma_odo"], c["loai"]) in con_lai]

    if not dong:
        raise LoiNguoiDung(
            "Không dựng được dòng nào. Thường là do file Item và file DO không cùng một lô "
            "xuất (không có mã ODO nào trùng nhau). Kiểm tra lại 4 file có cùng đợt không."
        )

    duong_dan_ra = _duong_dan_file_ra(thu_muc_ra)
    ghi_file(duong_dan_ra, dong)

    so_odo = len({r["A"] for r in dong})
    print("\n[4/4] Ket qua")
    print(f"  Xong: {len(dong)} dòng / {so_odo} đơn -> {duong_dan_ra.name}")
    if canh_bao:
        print("\n  Có cảnh báo (ô tô vàng trong file, xem chi tiết ở đây):")
        for nhom, dons in don_theo_nhom(canh_bao).items():
            print(f"  [{nhom} - {GIAI_THICH_NHOM[nhom]}] {len(dons)} đơn")
            theo_loai = {}
            for c in canh_bao:
                if nhom_cua(c["loai"]) == nhom:
                    theo_loai[c["loai"]] = theo_loai.get(c["loai"], 0) + 1
            for k, v in sorted(theo_loai.items(), key=lambda x: -x[1]):
                print(f"      - {k}: {v} đơn")
    return duong_dan_ra


def _tao_lay_toa_do(cfg):
    """-> hàm dia_chi -> (lat, lng) | None. Track Asia làm chính (gọi được ngay, không cần
    đăng ký). Track Asia lỗi/bị chặn (key dùng chung không có SLA) thì mới rơi xuống Nominatim
    (OpenStreetMap, không cần key nhưng giới hạn rất khắt khe - 1 request/giây - nên chỉ dùng
    làm phương án cuối, không gọi trước)."""
    def lay_toa_do(dia_chi):
        kq = trackasia_api.toa_do(dia_chi, cfg.trackasia_key)
        if kq is not None:
            return kq
        return nominatim_api.toa_do(dia_chi)
    return lay_toa_do


def tao_file_soldto_shipto(thu_muc_ra, duong_dan_du_lieu, duong_dan_config):
    """Bước 2: đọc lại file upload_STM_*.xlsx MỚI NHẤT trong thu_muc_ra (đã được khách sửa
    tay), xử lý Sold-to/Ship-to với cho_phep_tao_moi=True, ghi đè lại đúng file đó + xuất
    upload_soldto_shipto_*.xlsx nếu có đề xuất tạo mới (chỉ khi tra được toạ độ qua Track
    Asia - không có toạ độ thì không tạo, xem stm_soldto_shipto.xu_ly).
    -> (duong_dan_file_upload, duong_dan_file_soldto_shipto_hoac_None)."""
    thu_muc_ra = Path(thu_muc_ra)
    cac_file = sorted(thu_muc_ra.glob("upload_STM_*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not cac_file:
        raise LoiNguoiDung(
            f"Không thấy file upload_STM_*.xlsx nào trong:\n  {thu_muc_ra}\n"
            "Hãy bấm nút '2. Chạy' trước để tạo file upload đơn hàng."
        )
    duong_dan_file = cac_file[-1]

    du_lieu_dia_chi = _nap_du_lieu_dia_chi(duong_dan_du_lieu)
    if du_lieu_dia_chi is None:
        raise LoiNguoiDung(
            "Không đọc được dữ liệu địa chỉ tham chiếu (crosswalk/danh mục STM) trong:\n"
            f"  {duong_dan_du_lieu}\n"
            "Không thể xác định Tỉnh/Quận có đúng danh mục STM không."
        )

    cfg = doc_config(duong_dan_config)
    lay_toa_do = _tao_lay_toa_do(cfg)

    dong = doc_lai_upload(duong_dan_file)
    dong_log_dvhc_sai = []
    for d in dong:
        thieu = [c for c in ("O", "P") if not str(d.get(c, "")).strip()]
        if not thieu:
            # Khách tự gõ tay Tinh/Quan trong file upload (khong trong o thieu) nhung ten do
            # khong co trong danh_muc_stm.csv (go sai chinh ta, hoac ten khong duoc STM chap
            # nhan) - neu khong chan o day, xu_ly() se coi la "du tin cay" va de xuat tao
            # Ship-to moi voi Tinh/Quan sai, ghi thang vao file import STM. Chan lai giong het
            # nhu truong hop con thieu O/P (khong tao moi, cho khach sua roi bam lai nut 3).
            chuan = du_lieu_dia_chi.chuan_cap_stm(d["O"], d["P"])
            if chuan is None:
                thieu = ["O", "P"]
                dong_cb = (
                    f"Dong {d['_hang']}: Tinh/Quan '{d['O']}'/'{d['P']}' ban tu dien khong co "
                    "trong danh muc STM - sua lai dung ten roi bam lai nut 3")
                dong_log_dvhc_sai.append(dong_cb)
                # In thang ra console (khac cac canh bao khac chi ghi file) - khach can sua
                # LIEN va chay lai nut 3 ngay, khong phai mo file .txt moi biet.
                print(f"  [!] {dong_cb}")
            else:
                d["O"], d["P"] = chuan
        d["_thieu"] = thieu
        d["_ly_do"] = []

    duong_dan_cache = Path(duong_dan_du_lieu) / "stm_cache.json"
    cache = stm_cache.doc_cache(duong_dan_cache)
    thong_ke = stm_soldto_shipto.xu_ly(dong, cache, du_lieu_dia_chi, lay_toa_do,
                                        cho_phep_tao_moi=True)
    stm_cache.ghi_cache(duong_dan_cache, cache)

    thay_doi = [{"_hang": d["_hang"], "L": d["L"], "N": d["N"], "O": d["O"], "P": d["P"]}
                for d in dong]
    ghi_de_upload(duong_dan_file, thay_doi)

    dong_log = list(dong_log_dvhc_sai)
    dong_log += list(sorted(thong_ke["canh_bao_cache_sai"]))
    dong_log += [f"Bo qua Ship-to '{ma}' - con thieu Tinh/Quan, sua file roi bam lai nut 3"
                 for ma in sorted(thong_ke["bo_qua_thieu_dia_chi"])]
    dong_log += [f"Bo qua Ship-to '{ma}' - khong tra duoc toa do (Track Asia loi/khoa). "
                 "Dia chi van giu nguyen trong file upload don hang - STM se tu tao diem "
                 "(khong toa do) luc xu ly don."
                 for ma in sorted(thong_ke["bo_qua_khong_co_toa_do"])]
    duong_dan_log = _ghi_log_canh_bao(
        thu_muc_ra, f"canhbao_soldto_shipto_{_gio_vn_hien_tai():%Y%m%d_%H%M%S}.txt", dong_log)
    if duong_dan_log:
        print(f"  Co {len(dong_log)} canh bao chi tiet - xem file: {duong_dan_log.name}")

    n_soldto = len(thong_ke["can_tao_moi"]["soldto"])
    n_shipto = len(thong_ke["can_tao_moi"]["shipto"])
    print(f"  De xuat tao moi: {n_soldto} Sold-to, {n_shipto} Ship-to")
    if thong_ke["shipto_tai_dung"]:
        print(f"  Tai dung (da co san): {len(thong_ke['shipto_tai_dung'])} Ship-to")

    duong_dan_stm = None
    if n_soldto or n_shipto:
        duong_dan_stm = thu_muc_ra / (
            f"upload_soldto_shipto_{_gio_vn_hien_tai():%Y%m%d_%H%M%S}.xlsx")
        stm_api.ghi_file_soldto_shipto(duong_dan_stm, thong_ke["can_tao_moi"])
        print(f"  Da ghi file can tao moi: {duong_dan_stm.name}")
        print("    (import file nay qua UI Import Excel cua STM truoc khi dung don)")

    return duong_dan_file, duong_dan_stm


def _duong_dan_file_ra(thu_muc_ra):
    """Tên file kèm giây để giảm khả năng trùng; nếu vẫn trùng (2 lần chạy
    trong cùng 1 giây) thì tự thêm hậu tố _2, _3... để không đè file cũ."""
    goc = f"upload_STM_{_gio_vn_hien_tai():%Y%m%d_%H%M%S}"
    duong_dan = thu_muc_ra / f"{goc}.xlsx"
    i = 2
    while duong_dan.exists():
        duong_dan = thu_muc_ra / f"{goc}_{i}.xlsx"
        i += 1
    return duong_dan


def main():
    goc = thu_muc_goc()
    print("=" * 62)
    print("  TAO FILE UPLOAD DON HANG LEN STM  -  HAFELE")
    print("=" * 62)
    try:
        ra = chay(goc / "input", goc / "output", goc / "config.ini", goc / "bang_ma_tinh.csv")
        _mo_thu_muc(ra.parent)
    except LoiNguoiDung as e:
        print(f"\n[LOI] {e}")
    except Exception:
        log = goc / "output" / "log_loi.txt"
        log.parent.mkdir(parents=True, exist_ok=True)
        log.write_text(traceback.format_exc(), encoding="utf-8")
        print(f"\n[LOI KHONG MONG DOI] Đã ghi chi tiết vào:\n  {log}\n"
              "Hãy gửi file này cho bộ phận kỹ thuật.")
    print()
    input("Nhan Enter de dong cua so...")


def _mo_thu_muc(p):
    try:
        if os.name == "nt":
            os.startfile(p)               # noqa: S606
        else:
            subprocess.run(["xdg-open", str(p)], check=False)
    except Exception:
        pass


if __name__ == "__main__":
    main()

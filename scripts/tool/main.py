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
from .ghi_excel import (GIAI_THICH_NHOM, don_theo_nhom, doc_lai_upload, ghi_de_upload,
                         ghi_file, nhom_cua, xoa_canh_bao_dia_chi)
from .nhan_dien_file import nhan_dien_thu_muc
from . import stm_cache, stm_soldto_shipto, trackasia_api


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


def tao_file_soldto_shipto(thu_muc_ra, duong_dan_du_lieu, duong_dan_config):
    """Bước 2: đọc lại file upload_STM_*.xlsx MỚI NHẤT trong thu_muc_ra (đã được khách sửa
    tay), xử lý Sold-to/Ship-to với cho_phep_tao_moi=True - tái dùng điểm khớp cache (ghi đè
    Mã Ship to/Địa chỉ/Tỉnh/Quận), điểm không khớp để TRỐNG Mã Ship to cho STM tự tạo lúc xử
    lý đơn, ghi đè lại đúng file đó. KHÔNG còn xuất file Excel riêng, KHÔNG gọi geocode lấy
    toạ độ, KHÔNG tự ghi cache (xem main.xac_nhan_da_upload_don_hang).
    -> duong_dan_file_upload."""
    thu_muc_ra = Path(thu_muc_ra)
    cac_file = sorted(thu_muc_ra.glob("preview_upload_STM_*.xlsx"),
                       key=lambda p: p.stat().st_mtime)
    if not cac_file:
        raise LoiNguoiDung(
            f"Không thấy file preview_upload_STM_*.xlsx nào trong:\n  {thu_muc_ra}\n"
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

    # Vong 1 (khong goi mang): check Tinh/Quan khach tu dien co dung danh_muc_stm.csv
    # khong TRUOC KHI dong cham vao Sold-to/Ship-to (khop cache...). Loi go nho (vd thieu/du
    # vai ky tu) tu sua duoc qua khop mo
    # (xem chuan_cap_stm) thi IM LANG sua luon, khong can bao - chi bao that su khong
    # khop duoc gi. Con BAT KY 1 dong nao van sai sau khi da thu sua -> DUNG o day, KHONG
    # chay tiep buoc Sold-to/Ship-to cho ca nhung dong dung trong cung lan nay (tranh phi
    # cong goi mang/tao Ship-to roi phai lam lai khi khach sua xong con sai khac).
    dong = doc_lai_upload(duong_dan_file)
    dong_log_dvhc_sai = []
    hang_da_sua_dung = []
    for d in dong:
        thieu = [c for c in ("O", "P") if not str(d.get(c, "")).strip()]
        chuan = None
        if not thieu:
            chuan = du_lieu_dia_chi.chuan_cap_stm(d["O"], d["P"])
        if chuan is None:
            # Chua co Tinh/Quan hop le (thieu hoac sai chinh ta qua xa de fuzzy sua duoc) -
            # thu quet lai tu Dia chi + Tinh da sua tay, CHAP NHAN CA ten DVHC CU lan MOI
            # 2025 (vd khach go ten phuong/tinh moi) - dung lai dung engine suy_dvhc_tu_text
            # qua crosswalk_2025.csv (von da co san ca ten cu/moi), KHONG goi mang/geocode.
            kq = tra_dia_chi(d["N"], d["O"], du_lieu_dia_chi)
            if not kq["ly_do"]:
                chuan = (kq["O"], kq["P"])
                if kq["N"]:
                    d["N"] = kq["N"]
        if chuan is None:
            thieu = ["O", "P"]
            dong_cb = (
                f"Dong {d['_hang']}: Tinh/Quan '{d['O']}'/'{d['P']}' ban tu dien khong co "
                "trong danh muc STM, du da thu nhan dien ca ten DVHC cu lan moi tu Dia chi - "
                "sua lai dung ten roi bam lai nut 3")
            dong_log_dvhc_sai.append(dong_cb)
            print(f"  [!] {dong_cb}")
        else:
            d["O"], d["P"] = chuan
            thieu = []
            hang_da_sua_dung.append(d["_hang"])
        d["_thieu"] = thieu
        d["_ly_do"] = []

    # Ghi lai ngay nhung o da tu sua chinh ta/nhan dien lai tu Dia chi (ca N neu dia chi
    # duoc rap lai qua tra_dia_chi) - du con sai cho khac, lan chay sau khach khong phai
    # sua lai nhung cho da tu dung.
    ghi_de_upload(duong_dan_file, [{"_hang": d["_hang"], "N": d["N"], "O": d["O"], "P": d["P"]}
                                    for d in dong])
    xoa_canh_bao_dia_chi(duong_dan_file, hang_da_sua_dung)

    if dong_log_dvhc_sai:
        duong_dan_log = _ghi_log_canh_bao(
            thu_muc_ra, f"canhbao_soldto_shipto_{_gio_vn_hien_tai():%Y%m%d_%H%M%S}.txt",
            dong_log_dvhc_sai)
        print()
        print(f"  [!!!] CON {len(dong_log_dvhc_sai)} DONG TINH/QUAN SAI - CHUA CHAY DUOC "
              "BUOC SOLD-TO/SHIP-TO!")
        print(f"        Mo file '{duong_dan_file.name}' vua tai ve, sua dung ten "
              "Tinh/Quan o cac dong bi bao sai (xem chi tiet trong file canh bao "
              f"'{duong_dan_log.name if duong_dan_log else ''}'),")
        print("        roi CHAY LAI Buoc 2 (chon dung file vua sua) de xu ly tiep.")
        return duong_dan_file

    duong_dan_cache = Path(duong_dan_du_lieu) / "stm_cache.json"
    cache = stm_cache.doc_cache(duong_dan_cache)
    thong_ke = stm_soldto_shipto.xu_ly(dong, cache, du_lieu_dia_chi, cho_phep_tao_moi=True)
    # KHONG ghi cache o day - doi den main.xac_nhan_da_upload_don_hang(), chay SAU KHI khach
    # xac nhan da upload that len STM (sua gap kien truc da ghi nhan 20/10/2026: cache khong
    # nen lac quan ghi ngay luc de xuat).

    thay_doi = [{"_hang": d["_hang"], "L": d["L"], "N": d["N"], "O": d["O"], "P": d["P"]}
                for d in dong]
    ghi_de_upload(duong_dan_file, thay_doi)

    dong_log = list(sorted(thong_ke["canh_bao_cache_sai"]))
    dong_log += [f"Bo qua Ship-to '{ma}' - con thieu Tinh/Quan, sua file roi bam lai nut 3"
                 for ma in sorted(thong_ke["bo_qua_thieu_dia_chi"])]
    duong_dan_log = _ghi_log_canh_bao(
        thu_muc_ra, f"canhbao_soldto_shipto_{_gio_vn_hien_tai():%Y%m%d_%H%M%S}.txt", dong_log)
    if duong_dan_log:
        print(f"  Co {len(dong_log)} canh bao chi tiet - xem file: {duong_dan_log.name}")

    print(f"  Tai dung (da co san): {len(thong_ke['shipto_tai_dung'])} Ship-to")
    if thong_ke["shipto_tao_moi"]:
        print(f"  {len(thong_ke['shipto_tao_moi'])} Sold-to co diem giao MOI - da de trong "
              "Ma Ship to trong file upload, STM se tu tao diem luc xu ly don. Sau khi "
              "upload file nay thanh cong len STM, chay 'Xac nhan da upload don hang' de "
              "tool nho lai cho lan sau.")

    if duong_dan_file.name.startswith("preview_"):
        duong_dan_cuoi = duong_dan_file.with_name(duong_dan_file.name[len("preview_"):])
        duong_dan_file.rename(duong_dan_cuoi)
        duong_dan_file = duong_dan_cuoi
        print(f"  Da doi ten file thanh '{duong_dan_file.name}' - san sang upload len STM.")

    return duong_dan_file


def xac_nhan_da_upload_don_hang(thu_muc_ra, duong_dan_du_lieu):
    """Chạy SAU KHI khách đã tự import thành công file upload_STM_*.xlsx (đơn hàng, KHÔNG
    còn tiền tố 'preview_') lên STM. Đọc lại đúng file đó, tìm các dòng tool đã để TRỐNG Mã
    Ship to vì không khớp cache (nhóm 'điểm mới', xem stm_soldto_shipto.xu_ly) - các điểm
    này giờ khách đã upload thật, STM đã tự tạo điểm khi xử lý đơn - ghi địa chỉ/tỉnh/quận
    (KHÔNG kèm mã - tool không biết STM sinh mã gì) vào cache để các lần chạy sau không còn
    coi là hoàn toàn mới.
    -> (duong_dan_file, so_diem_moi_ghi_vao_cache)."""
    thu_muc_ra = Path(thu_muc_ra)
    cac_file = sorted(thu_muc_ra.glob("upload_STM_*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not cac_file:
        raise LoiNguoiDung(
            f"Không thấy file upload_STM_*.xlsx (đã hoàn tất, không còn tiền tố "
            f"'preview_') nào trong:\n  {thu_muc_ra}\n"
            "Hãy bấm '3. Tạo file Sold-to/Ship-to trên STM' và sửa hết lỗi trước, upload "
            "file đó lên STM, rồi mới xác nhận."
        )
    duong_dan_file = cac_file[-1]

    dong = doc_lai_upload(duong_dan_file)
    duong_dan_cache = Path(duong_dan_du_lieu) / "stm_cache.json"
    cache = stm_cache.doc_cache(duong_dan_cache)

    diem_moi = {}
    for d in dong:
        ma_soldto = d.get("K", "")
        if not ma_soldto or str(d.get("L", "")).strip():
            continue
        dia_chi, tinh, quan = d.get("N", ""), d.get("O", ""), d.get("P", "")
        if not (dia_chi and tinh and quan):
            continue
        diem_moi[(ma_soldto, dia_chi)] = (tinh, quan)

    for (ma_soldto, dia_chi), (tinh, quan) in diem_moi.items():
        stm_cache.them_diem_moi_khong_ma(cache, ma_soldto, dia_chi, tinh, quan)
    stm_cache.ghi_cache(duong_dan_cache, cache)

    print(f"  Da ghi {len(diem_moi)} diem moi vao cache (khong kem ma Ship-to).")
    return duong_dan_file, len(diem_moi)


def _duong_dan_file_ra(thu_muc_ra):
    """Tên file kèm giây để giảm khả năng trùng; nếu vẫn trùng (2 lần chạy
    trong cùng 1 giây) thì tự thêm hậu tố _2, _3... để không đè file cũ.

    Tiền tố 'preview_' đánh dấu đây là file CHƯA chắc chắn đúng — khách cần mở Excel soát
    lại trước khi bấm nút 3. Nút 3 chạy xong không còn lỗi Tỉnh/Quận/Sold-to-Ship-to nào sẽ
    tự đổi tên bỏ tiền tố này (xem tao_file_soldto_shipto)."""
    goc = f"preview_upload_STM_{_gio_vn_hien_tai():%Y%m%d_%H%M%S}"
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

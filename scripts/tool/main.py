# scripts/tool/main.py
# -*- coding: utf-8 -*-
"""Điểm vào của tool. Khách double-click CHAY_TOOL.exe là chạy hàm main().

Từ 05/11/2026: tool chỉ còn 1 bước — tạo file upload từ 4 file SAP. Việc tái dùng/tạo
Sold-to/Ship-to đã giao hẳn cho STM xử lý qua popup "Danh sách điểm giao" (TMS-15494,
QA DONE 04/11/2026). Tool chỉ cần đảm bảo cột Địa chỉ/Tỉnh/Quận ở ĐVHC cũ chuẩn để STM
khớp 70% có nghĩa — các ô thiếu/chưa chắc vẫn tô vàng để khách sửa tay trong Excel trước
khi upload lên STM."""
import datetime
import os
import subprocess
import sys
import traceback
from pathlib import Path

from .cau_hinh import doc_bang_ma_tinh, doc_config
from .dia_chi_cu import DuLieu, tra_dia_chi
from .doc_excel import doc_bang
from .dung_bang import _mo_ta_mo_ho, dung_bang
from .ghi_excel import (GIAI_THICH_NHOM, LOAI_SO_LUONG_AM, LOAI_THIEU_NGAY,
                         LOAI_THIEU_SHIP_TO, LOAI_THIEU_SOLD_TO,
                         cap_nhat_canh_bao_dia_chi, doc_lai_canh_bao_sap, doc_lai_upload,
                         don_theo_nhom, ghi_de_upload, ghi_file, nhom_cua,
                         xoa_canh_bao_dia_chi, xoa_canh_bao_sap_goc)
from .nhan_dien_file import nhan_dien_thu_muc
from . import trackasia_api


_GIO_VN = datetime.timezone(datetime.timedelta(hours=7))


def _gio_vn_hien_tai():
    """Giờ Việt Nam (UTC+7) dùng để đặt tên file — máy Windows khách để sẵn giờ VN
    nên không khác gì trước đây, nhưng máy chủ Colab mặc định chạy UTC, lệch 7
    tiếng nếu dùng datetime.now() thường (bug thật 19/10/2026, tên file tạo trên
    Colab hiện giờ sáng dù tạo lúc chiều)."""
    return datetime.datetime.now(_GIO_VN)


class LoiNguoiDung(Exception):
    """Lỗi khách tự sửa được — in gọn, không in traceback."""


def thu_muc_goc():
    """Chỗ đặt exe (lúc đóng gói) hoặc gốc project (lúc chạy bằng python)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]


def thu_muc_du_lieu():
    """Chỗ đặt crosswalk/danh mục/bảng tra cứu.

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
            f4 = None
            p4 = d / "phuong_pre2020_hcm.csv"
            if p4.exists():
                with open(p4, encoding="utf-8-sig") as f4_open:
                    return DuLieu(f1, f2, f3, f4_open)
            return DuLieu(f1, f2, f3, f4)
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

    print("\n[1/3] Doc file input")
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

    print("\n[2/3] Dia chi (DVHC cu)")
    tra_dc = tao_tra_dc(thu_muc_du_lieu(), cfg)
    if tra_dc is None:
        tra_dc = lambda dia_chi, tinh: {"N": dia_chi, "O": tinh, "P": "",
                                         "do_tin": "THAP", "ly_do": "khong_co_quan"}
    dong, canh_bao = dung_bang(do_ten_cot, do_dong, item_dong, so_dong, uom_dong,
                               cfg, bang_tinh, tra_dc)

    if not dong:
        raise LoiNguoiDung(
            "Không dựng được dòng nào. Thường là do file Item và file DO không cùng một lô "
            "xuất (không có mã ODO nào trùng nhau). Kiểm tra lại 4 file có cùng đợt không."
        )

    duong_dan_ra = _duong_dan_file_ra(thu_muc_ra)
    ghi_file(duong_dan_ra, dong)

    so_odo = len({r["A"] for r in dong})
    print("\n[3/3] Ket qua")
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


def quet_lai_sau_sua_tay(thu_muc_ra, duong_dan_du_lieu, duong_dan_file=None):
    """Bước 2: khách đã mở file upload_STM_*.xlsx, sửa tay Địa chỉ/Tỉnh/Quận/Ngày gửi/
    Số lượng/Mã Sold to/Mã Ship to trực tiếp trong Excel. Bước này đọc lại, re-validate,
    và XOÁ cảnh báo + bỏ tô vàng ở các dòng giờ đã hợp lệ. Dòng nào vẫn mơ hồ địa chỉ
    thì RE-WRITE lý do theo nội dung MỚI (không giữ lý do stale từ Bước 1).

    Logic quét Tỉnh/Quận — ưu tiên cột N (Địa chỉ) vì khách thường sửa chỗ này:
      1) Luôn chạy tra_dia_chi(N) lại từ đầu, engine thuần chữ, chấp nhận cả tên ĐVHC
         CŨ lẫn MỚI 2025 (crosswalk có sẵn cả 2). Ra 1 cặp duy nhất → dùng, xoá cảnh báo.
      2) Nếu vẫn mơ hồ/không khớp: thử (O, P) user điền. Đủ 2 ô + khớp danh_muc_stm.csv
         (fuzzy nhẹ qua chuan_cap_stm) → tin user.
      3) Cả 2 không được → rewrite cảnh báo với lý do MỚI suy từ địa chỉ user hiện có.

    duong_dan_file: đường dẫn file cụ thể (Colab upload). Nếu None → scan thu_muc_ra lấy
    file upload_STM_*.xlsx mới nhất (GUI .exe, hành vi cũ).

    -> (duong_dan_file, so_dong_quet_duoc, so_dong_van_sai)."""
    if duong_dan_file is not None:
        duong_dan_file = Path(duong_dan_file)
        if not duong_dan_file.exists():
            raise LoiNguoiDung(f"Không thấy file: {duong_dan_file}")
    else:
        thu_muc_ra = Path(thu_muc_ra)
        cac_file = sorted(thu_muc_ra.glob("upload_STM_*.xlsx"),
                          key=lambda p: p.stat().st_mtime)
        if not cac_file:
            raise LoiNguoiDung(
                f"Không thấy file upload_STM_*.xlsx nào trong:\n  {thu_muc_ra}\n"
                "Hãy bấm Bước 1 trước để tạo file upload, hoặc truyền thẳng đường dẫn "
                "file đã sửa tay (Colab)."
            )
        duong_dan_file = cac_file[-1]

    du_lieu_dia_chi = _nap_du_lieu_dia_chi(duong_dan_du_lieu)
    if du_lieu_dia_chi is None:
        raise LoiNguoiDung(
            "Không đọc được dữ liệu địa chỉ tham chiếu (crosswalk/danh mục STM) trong:\n"
            f"  {duong_dan_du_lieu}"
        )

    dong = doc_lai_upload(duong_dan_file)
    hang_da_sua_dung = []
    hang_van_sai = []
    thay_doi = []
    canh_bao_moi = []
    for d in dong:
        kq = tra_dia_chi(d["N"], d["O"], du_lieu_dia_chi)
        chuan = None
        if not kq["ly_do"]:
            chuan = (kq["O"], kq["P"])
            if kq["N"]:
                d["N"] = kq["N"]
        elif str(d["O"]).strip() and str(d["P"]).strip():
            chuan = du_lieu_dia_chi.chuan_cap_stm(d["O"], d["P"])
        if chuan is not None:
            d["O"], d["P"] = chuan
            hang_da_sua_dung.append(d["_hang"])
            thay_doi.append({"_hang": d["_hang"], "N": d["N"],
                             "O": d["O"], "P": d["P"]})
        else:
            hang_van_sai.append(d["_hang"])
            canh_bao_moi.append(_tinh_canh_bao_dia_chi_moi(d, kq))

    ghi_de_upload(duong_dan_file, thay_doi)
    xoa_canh_bao_dia_chi(duong_dan_file, hang_da_sua_dung)
    cap_nhat_canh_bao_dia_chi(duong_dan_file, canh_bao_moi)

    dong_sap = doc_lai_canh_bao_sap(duong_dan_file)
    hang_da_sua_ngay = [
        d["_hang"] for d in dong_sap
        if LOAI_THIEU_NGAY in str(d["AB"])
        and d["D"] not in (None, "") and d["T"] not in (None, "")
    ]
    hang_da_sua_so_luong = [
        d["_hang"] for d in dong_sap
        if LOAI_SO_LUONG_AM in str(d["AB"])
        and isinstance(d["G"], (int, float)) and d["G"] > 0
    ]
    hang_da_sua_sold_to = [
        d["_hang"] for d in dong_sap
        if LOAI_THIEU_SOLD_TO in str(d["AB"])
        and str(d.get("K", "") or "").strip()
    ]
    hang_da_sua_ship_to = [
        d["_hang"] for d in dong_sap
        if LOAI_THIEU_SHIP_TO in str(d["AB"])
        and str(d.get("L", "") or "").strip()
    ]
    xoa_canh_bao_sap_goc(duong_dan_file, hang_da_sua_ngay, hang_da_sua_so_luong,
                         hang_da_sua_sold_to, hang_da_sua_ship_to)

    print(f"\n  Đã quét lại: {len(hang_da_sua_dung)} dòng Tỉnh/Quận đã hợp lệ, "
          f"{len(hang_van_sai)} dòng vẫn còn sai")
    if (hang_da_sua_ngay or hang_da_sua_so_luong
            or hang_da_sua_sold_to or hang_da_sua_ship_to):
        print(f"  Đã xoá cảnh báo: {len(hang_da_sua_ngay)} dòng ngày gửi/ETD, "
              f"{len(hang_da_sua_so_luong)} dòng số lượng, "
              f"{len(hang_da_sua_sold_to)} dòng Mã Sold to, "
              f"{len(hang_da_sua_ship_to)} dòng Mã Ship to")
    if hang_van_sai:
        print(f"  [!] {len(hang_van_sai)} dòng vẫn còn Tỉnh/Quận không khớp danh mục STM — "
              "sửa cột Địa chỉ cho rõ hơn hoặc gõ thẳng Tỉnh+Quận rồi bấm lại Bước 2.")

    return duong_dan_file, len(hang_da_sua_dung), len(hang_van_sai)


def _tinh_canh_bao_dia_chi_moi(d, kq):
    """Tính lý do cảnh báo địa chỉ MỚI cho 1 dòng vẫn sai sau khi re-quét từ N. Trả
    dict {_hang, canh_bao, o_vang} để ghi_excel.cap_nhat_canh_bao_dia_chi() xài."""
    canh_bao = []
    o_vang = []
    ly_do = kq.get("ly_do", "")
    if ly_do == "mo_ho":
        canh_bao.append((
            "Có nhiều Quận/Huyện cũ ứng với địa chỉ này",
            f"{_mo_ta_mo_ho(kq)} — chọn 1 tên, gõ vào cột Quận rồi bấm Bước 2",
        ))
        o_vang += ["N", "P"]
    elif ly_do == "khong_hop_le_stm":
        canh_bao.append((
            "Quận không có trong danh sách STM chấp nhận",
            f"Suy ra '{kq.get('quan_tho', '')}' ({kq.get('tinh_tho', '')}) "
            "nhưng STM không có tên này, cần tự tra đúng tên",
        ))
        o_vang += ["N", "P"]
    else:
        if not str(d["P"]).strip():
            canh_bao.append((
                "Không xác định được Quận/Huyện",
                "Địa chỉ SAP không có chữ khớp được quận/phường nào, cần điền tay",
            ))
            o_vang += ["N", "P"]
    if not str(kq.get("O", "") or d["O"]).strip():
        canh_bao.append((
            "Không xác định được Tỉnh",
            "Địa chỉ SAP không ghi tên tỉnh, cần điền tay",
        ))
        if "O" not in o_vang:
            o_vang.append("O")
    return {"_hang": d["_hang"], "canh_bao": canh_bao, "o_vang": o_vang}


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

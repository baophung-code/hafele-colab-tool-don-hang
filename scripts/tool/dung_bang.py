# scripts/tool/dung_bang.py
# -*- coding: utf-8 -*-
"""Logic nghiệp vụ: 4 bảng SAP -> danh sách dòng của file upload.

Module này KHÔNG đọc file và KHÔNG ghi file — chỉ nhận list[dict] và trả
list[dict]. Nhờ vậy test được toàn bộ nghiệp vụ bằng dữ liệu bịa nhỏ gọn.

Sơ đồ dựng:
    DO (dedupe theo Delivery)     <- quyết định ĐƠN NÀO được upload
      └─ join Item theo Delivery  <- quyết định MỖI ĐƠN RA MẤY DÒNG
           └─ join UOM theo Mã đơn hàng  <- lấy Mã Ship to
"""
import datetime

from .chuan_hoa import ghep_dia_chi, ten_ship_to, thong_tin_nguoi_nhan

COT_TRONG = ["F", "H", "R", "S"]
MA_HINH_THUC_THANH_TOAN_MAC_DINH = "Congno"

# Cặp (Mã điểm lấy hàng, Mã Ship to) coi là chuyển kho nội bộ, không phải giao khách.
KHO_NOI_BO = {"VN10", "VN20"}

# Mã hàng là phí (vận chuyển/lắp đặt/phạt hợp đồng...), không phải hàng hóa thật
# -> loại bỏ khỏi file upload, không cảnh báo. Khách chốt danh sách 20/10/2026.
MA_HANG_PHI = {
    "999.11.001", "999.11.009", "999.11.064", "999.11.310", "999.11.311", "999.11.313",
    "999.11.314", "999.11.315", "999.11.316", "999.11.317", "999.11.319", "999.11.321",
    "999.11.322", "999.11.323", "999.11.324", "999.11.327", "999.11.328", "999.11.329",
    "999.11.330", "999.11.332", "999.11.333", "999.11.335", "999.11.338", "999.11.345",
    "999.11.346", "999.11.348", "999.11.349", "999.11.369", "999.11.370", "999.11.371",
    "999.11.372", "999.11.373", "999.11.374", "999.11.376", "999.11.377", "999.11.378",
    "999.11.379", "999.11.380", "999.11.382", "999.11.386", "999.11.400", "999.11.410",
    "999.11.411", "999.11.417", "999.11.418", "999.11.419", "999.11.420", "999.11.421",
    "999.11.422", "999.11.423", "999.11.424", "999.11.425", "999.11.426", "999.11.428",
    "999.11.429", "999.11.430", "999.11.431", "999.11.432",
}


def dung_bang(do_ten_cot, do_dong, item_dong, so_dong, uom_dong, cfg, bang_tinh, tra_dc):
    """-> (dong, canh_bao)

    tra_dc: hàm (dia_chi_goc, ten_tinh_rg) -> dict(N, O, P, do_tin, ly_do) — xem
            scripts/tool/dia_chi_cu.py:tra_dia_chi(), cửa duy nhất để suy Tỉnh/Quận/Phường
            CŨ từ chữ. Bắt buộc truyền (không còn nhánh "cách cũ" dùng tach_quan()).
    """
    canh_bao = []

    # Cột A của DO.XLSX có tên dài bất thường ('ExterLieferung nAngel.IstWA Dat  VB')
    # và SAP có thể đổi -> đọc theo INDEX 0 cho chắc.
    ten_cot_ma_don = do_ten_cot[0]

    do_theo_odo = {}
    for d in do_dong:                       # DO.XLSX có 353 dòng trùng lặp hoàn toàn
        odo = str(d.get("Delivery", "")).strip()
        if odo and odo not in do_theo_odo:
            do_theo_odo[odo] = d

    ship_to = {}
    for u in uom_dong:
        ma = str(u.get("Shipment number", "")).strip()
        if ma:
            ship_to.setdefault(ma, str(u.get("Ship-To", "")).strip())

    ngay_kh, gio_xuat = {}, {}
    for s in so_dong:
        odo = str(s.get("ERP Document", "")).strip()
        if not odo:
            continue
        v = s.get(cfg.nguon_ngay)
        if odo not in ngay_kh and isinstance(v, datetime.datetime):
            ngay_kh[odo] = v
        g = s.get("Actual Time of Goods Issue")
        if odo not in gio_xuat and isinstance(g, datetime.time):
            gio_xuat[odo] = g

    dong, da_bao = [], set()

    def bao(odo, loai, chi_tiet=""):
        if (odo, loai) in da_bao:
            return
        da_bao.add((odo, loai))
        canh_bao.append({"ma_odo": odo, "loai": loai, "chi_tiet": chi_tiet})

    for it in item_dong:
        odo = str(it.get("Delivery", "")).strip()
        d = do_theo_odo.get(odo)
        if d is None:
            continue                        # ngoài phạm vi driver — bỏ, không cảnh báo

        if str(it.get("Material", "")).strip() in MA_HANG_PHI:
            continue                        # mã phí, không phải hàng hóa — bỏ, không cảnh báo

        thieu, ly_do_dong, chi_tiet_dong = [], [], {}

        def danh_dau(loai, cot_moi, chi_tiet=""):
            bao(odo, loai, chi_tiet)
            for c in cot_moi:
                if c not in thieu:
                    thieu.append(c)
            if loai not in ly_do_dong:
                ly_do_dong.append(loai)
                if chi_tiet:
                    chi_tiet_dong[loai] = chi_tiet

        ma_don = str(d.get(ten_cot_ma_don, "")).strip()
        if not ma_don:
            danh_dau("Thiếu Mã đơn hàng", ["B"],
                      "SAP không có dữ liệu cho delivery này")

        # Thiếu Mã đặt hàng (Reference Document) không cảnh báo - khách xác nhận
        # 21/10/2026 không cần báo, khác với Mã đơn hàng vẫn giữ cảnh báo.
        ma_dat = str(it.get("Reference Document", "")).strip()

        # Thiếu Mã Ship to (vd nhóm DECO) không còn cảnh báo - đã xác nhận nghiệp vụ
        # 16/10/2026: STM tự khớp địa chỉ/đề xuất tạo điểm trong cùng Sold-to lúc lên
        # đơn, không cần tool báo lại hay tô vàng.
        l = ship_to.get(ma_don, "")

        ngay = ngay_kh.get(odo)
        if ngay is None:
            danh_dau("Thiếu ngày gửi / ETD (cột D, T)", ["D", "T"],
                      f"SAP không có {cfg.nguon_ngay} cho đơn này")
            cot_d = cot_t = ""
        else:
            cot_d = ngay
            cot_t = ngay + datetime.timedelta(hours=cfg.etd_cong_gio)
        cot_u = ""   # ETA khong con dung - chot 14/10/2026

        rg = str(d.get("Rg", "")).strip()
        ten_tinh_rg, tin_cay_rg = bang_tinh.get(rg, ("", ""))

        dia_chi = ghep_dia_chi(d, cfg.cot_dia_chi, cfg.cum_tu)
        dia_chi_goc = dia_chi

        dc = tra_dc(dia_chi, ten_tinh_rg)
        ten_tinh, quan = dc["O"], dc["P"]
        if dc["N"]:
            dia_chi = dc["N"]

        if not quan:
            if dc["ly_do"] == "mo_ho":
                ds_ung_vien = " hoặc ".join(f"{q} ({t})" for t, q in dc.get("ung_vien", []))
                danh_dau("Quận/Huyện trùng tên nhiều nơi", ["N", "P"],
                          f"Chưa rõ là: {ds_ung_vien or 'không rõ'} — tự chọn đúng rồi điền tay")
            elif dc["ly_do"] == "khong_hop_le_stm":
                danh_dau("Quận không có trong danh sách STM chấp nhận", ["N", "P"],
                          f"Suy ra '{dc.get('quan_tho', '')}' ({dc.get('tinh_tho', '')}) "
                          f"nhưng STM không có tên này, cần tự tra đúng tên")
            else:
                danh_dau("Không xác định được Quận/Huyện", ["N", "P"],
                          f"Địa chỉ SAP không có chữ khớp được quận/phường nào, cần điền tay")

        if not ten_tinh:
            danh_dau("Không xác định được Tỉnh", ["O"],
                      "Địa chỉ SAP không ghi tên tỉnh, cần điền tay")
        elif tin_cay_rg == "thap" and dc["do_tin"] != "CAO":
            danh_dau("Tỉnh độ tin cậy thấp", ["O"], f"Tool đoán {ten_tinh}, cần soát lại")

        sl = it.get("Delivery quantity", "")
        if isinstance(sl, (int, float)) and sl <= 0:
            danh_dau("Số lượng ≤ 0", ["G"], f"mã hàng {it.get('Material', '')} có số lượng {sl}")

        ngay_ban = d.get("Ac.GI date")
        if isinstance(ngay_ban, datetime.datetime):
            g = gio_xuat.get(odo)
            cot_w = ngay_ban.replace(hour=g.hour, minute=g.minute, second=g.second) if g else ngay_ban
        else:
            cot_w = ""

        j = str(it.get("Plant", "")).strip()
        loai_dich_vu = "Chuyển kho" if {j, l} == KHO_NOI_BO else "Giao hàng"

        r = {
            "A": odo, "B": ma_don, "C": ma_dat,
            "D": cot_d, "E": str(it.get("Material", "")).strip(), "G": sl,
            "I": str(it.get("Storage Location", "")).strip(), "J": j,
            "K": str(d.get("Sold-to pt", "")).strip(), "L": l,
            "M": ten_ship_to(d.get("Name", ""), d.get("Name 2", "")),
            "_ten_soldto": str(d.get("Sold-to party", "")).strip(),
            "N": dia_chi, "O": ten_tinh, "P": quan,
            "Q": MA_HINH_THUC_THANH_TOAN_MAC_DINH,
            "T": cot_t, "U": cot_u, "V": str(d.get("Handling Unit", "")).strip(), "W": cot_w,
            "X": "LTL", "Y": loai_dich_vu,
            "AA": thong_tin_nguoi_nhan(d, cfg.cot_dia_chi, cfg.cum_tu),
            "_thieu": thieu, "_ly_do": ly_do_dong, "_chi_tiet": chi_tiet_dong,
            "_dia_chi_goc": dia_chi_goc,
        }
        for c in COT_TRONG:
            r[c] = ""
        dong.append(r)

    return dong, canh_bao

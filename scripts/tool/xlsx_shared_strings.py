# -*- coding: utf-8 -*-
"""openpyxl (từ bản 2.6 trở đi) luôn ghi cell chuỗi dạng inline string
(``<c t="inlineStr"><is><t>...</t></is></c>``), KHÔNG còn dùng bảng shared strings
(xl/sharedStrings.xml) như Excel thật vẫn làm — đây là hành vi cố định của thư viện, không
phải do code của tool. Phát hiện 21/10/2026: màn hình "Tạo mới đơn hàng excel" của STM (ảnh
khách gửi) đọc file .xlsx bằng 1 bộ phân tích xlsx riêng CHỈ hiểu shared strings, bỏ qua hoàn
toàn cell inline string — file tool xuất ra, nếu CHƯA từng mở bằng Excel thật (Excel luôn tự
ghi lại bằng shared strings mỗi khi lưu) mà upload thẳng, mọi cột kiểu chữ (mã ODO, mã hàng,
địa chỉ...) hiện trống trơn, chỉ còn cột ngày/số (kiểu "n", không qua shared strings) là có dữ
liệu — đúng triệu chứng trong ảnh khách báo.

Hàm `chuyen_sang_shared_strings()` hậu xử lý file .xlsx NGAY SAU KHI `wb.save()` ghi xong, viết
lại toàn bộ cell inline string (mọi sheet trong file) sang shared strings, để file tải về dùng
được ngay cả khi khách KHÔNG mở qua Excel. Chỉ dùng `zipfile` + `xml.etree.ElementTree` có sẵn
trong Python - không thêm phụ thuộc mới (lxml), giữ đúng triết lý "chỉ dùng openpyxl" của tool.
"""
import re
import zipfile
import xml.etree.ElementTree as ET

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_XML = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("", _NS_MAIN)


def _noi_dung_is(the_is):
    """-> text của 1 thẻ <is> (inline string) - gộp mọi <t> con, kể cả nằm trong <r> (rich text).
    Không có <t> nào (cell rỗng) -> ""."""
    return "".join(t.text or "" for t in the_is.iter(f"{{{_NS_MAIN}}}t"))


def _chuyen_1_sheet(xml_bytes):
    """-> (xml_bytes_moi, [text đã dùng theo đúng thứ tự xuất hiện]) nếu sheet có inline
    string cần chuyển, hoặc (None, []) nếu sheet không có gì để đổi."""
    root = ET.fromstring(xml_bytes)
    da_dung = []
    co_doi = False
    for c in root.iter(f"{{{_NS_MAIN}}}c"):
        if c.get("t") != "inlineStr":
            continue
        the_is = c.find(f"{{{_NS_MAIN}}}is")
        text = _noi_dung_is(the_is) if the_is is not None else ""
        if the_is is not None:
            c.remove(the_is)
        if text == "":
            del c.attrib["t"]
            co_doi = True
            continue
        da_dung.append(text)
        c.set("t", "s")
        ET.SubElement(c, f"{{{_NS_MAIN}}}v")
        co_doi = True
    if not co_doi:
        return None, []
    return ET.tostring(root, encoding="unicode").encode("utf-8"), da_dung


def _gan_chi_so(xml_bytes, chi_so_theo_chuoi):
    """Sau khi đã biết chỉ số shared string cuối cùng của từng chuỗi, điền vào các thẻ <v>
    rỗng vừa tạo ở `_chuyen_1_sheet` theo đúng thứ tự xuất hiện."""
    root = ET.fromstring(xml_bytes)
    it = iter(chi_so_theo_chuoi)
    for c in root.iter(f"{{{_NS_MAIN}}}c"):
        if c.get("t") != "s":
            continue
        v = c.find(f"{{{_NS_MAIN}}}v")
        if v is not None and not v.text:
            v.text = str(next(it))
    return ET.tostring(root, encoding="unicode").encode("utf-8")


def _tao_shared_strings_xml(thu_tu, tong_so_lan):
    sst = ET.Element(f"{{{_NS_MAIN}}}sst",
                      {"count": str(tong_so_lan), "uniqueCount": str(len(thu_tu))})
    for s in thu_tu:
        si = ET.SubElement(sst, f"{{{_NS_MAIN}}}si")
        t = ET.SubElement(si, f"{{{_NS_MAIN}}}t")
        t.text = s
        if s != s.strip():
            t.set(f"{{{_NS_XML}}}space", "preserve")
    return ET.tostring(sst, encoding="unicode").encode("utf-8")


def _them_content_type(data):
    phan_them = (
        b'<Override PartName="/xl/sharedStrings.xml" ContentType='
        b'"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    )
    return data.replace(b"</Types>", phan_them + b"</Types>")


def _them_quan_he_shared_strings(data):
    da_dung = {int(m) for m in re.findall(rb'Id="rId(\d+)"', data)}
    rid_moi = max(da_dung, default=0) + 1
    phan_them = (
        f'<Relationship Id="rId{rid_moi}" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    ).encode("utf-8")
    return data.replace(b"</Relationships>", phan_them + b"</Relationships>")


def chuyen_sang_shared_strings(duong_dan):
    """Đọc file .xlsx tại `duong_dan`, chuyển toàn bộ cell inline string (mọi sheet) sang
    shared strings, ghi đè lại đúng file đó. Không có cell chuỗi nào (hiếm, vd file test
    rỗng) -> không đổi gì. Lỗi đọc/parse bất kỳ -> bỏ qua, giữ nguyên file gốc (đây là bước
    tăng tương thích, không phải bước bắt buộc — lỗi ở đây không được làm gián đoạn luồng
    chính ghi file upload)."""
    try:
        with zipfile.ZipFile(duong_dan) as zf:
            noi_dung = {n: zf.read(n) for n in zf.namelist()}
    except Exception:
        return

    try:
        ten_sheet = [n for n in noi_dung
                     if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        chuoi_moi_sheet = {}  # ten -> (xml_moi, [text theo dung thu tu xuat hien])
        chi_so_theo_chuoi = {}
        thu_tu = []
        tong_so_lan = 0
        for ten in ten_sheet:
            xml_moi, da_dung = _chuyen_1_sheet(noi_dung[ten])
            if xml_moi is None:
                continue
            chuoi_moi_sheet[ten] = (xml_moi, da_dung)
            for text in da_dung:
                if text not in chi_so_theo_chuoi:
                    chi_so_theo_chuoi[text] = len(thu_tu)
                    thu_tu.append(text)
                tong_so_lan += 1

        if not thu_tu:
            return

        for ten, (xml_moi, da_dung) in chuoi_moi_sheet.items():
            chi_so_cua_sheet = [chi_so_theo_chuoi[t] for t in da_dung]
            noi_dung[ten] = _gan_chi_so(xml_moi, chi_so_cua_sheet)

        noi_dung["xl/sharedStrings.xml"] = _tao_shared_strings_xml(thu_tu, tong_so_lan)
        noi_dung["[Content_Types].xml"] = _them_content_type(noi_dung["[Content_Types].xml"])
        noi_dung["xl/_rels/workbook.xml.rels"] = _them_quan_he_shared_strings(
            noi_dung["xl/_rels/workbook.xml.rels"])
    except Exception:
        return

    with zipfile.ZipFile(duong_dan, "w", zipfile.ZIP_DEFLATED) as zf:
        for ten, data in noi_dung.items():
            zf.writestr(ten, data)

"""Tạo 2 file Word mẫu để thử chương trình: mau_dung.docx (đúng chuẩn) và mau_loi.docx (cố tình mắc lỗi).

Chạy:  python tests/tao_mau.py [thư_mục_đầu_ra]
"""
from __future__ import annotations

import os
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

NB = "\u00a0"  # dấu cách không ngắt
# Hai chuỗi dưới đây đã được thử với LibreOffice (Liberation Serif = số đo chữ y hệt Times New Roman), lề 2 cm, cỡ 14:
# đoạn văn dòng 2 chỉ còn đúng 1 chữ ("thời"); LONG_TITLE in đậm, căn giữa thì dòng 2 chỉ còn "Tống".
ORPHAN_TEXT = "Phòng giao dịch đã phối hợp với Ủy ban nhân dân các xã và các tổ chức kịp thời"
LONG_TITLE = "V/v kiểm tra việc sử dụng vốn vay của các hộ nghèo tại các xã trên địa bàn xã Tống"


def setup(doc: Document, font="Times New Roman", size=14, margins=(2.0, 2.0, 3.0, 1.5)):
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = (Cm(m) for m in margins)
    st = doc.styles["Normal"]
    st.font.name = font
    st.element.rPr.rFonts.set(qn("w:eastAsia"), font)
    st.font.size = Pt(size)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.line_spacing = 1.15


def para(container, text="", bold=False, italic=False, size=None, align=None, indent=None, after=None, keep_next=False):
    p = container.add_paragraph()
    if text:
        r = p.add_run(text)
        r.bold, r.italic = bold, italic
        if size:
            r.font.size = Pt(size)
    if align is not None:
        p.alignment = align
    if indent is not None:
        p.paragraph_format.first_line_indent = Cm(indent)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if keep_next:
        p.paragraph_format.keep_with_next = True
    return p


def cell_par(cell, text, first=False, **kw):
    if first and cell.paragraphs and not cell.paragraphs[0].text:
        p = cell.paragraphs[0]
        r = p.add_run(text)
        r.bold, r.italic = kw.get("bold", False), kw.get("italic", False)
        if kw.get("size"):
            r.font.size = Pt(kw["size"])
        if kw.get("align") is not None:
            p.alignment = kw["align"]
        p.paragraph_format.space_after = Pt(0)
        return p
    p = para(cell, text, after=0, **kw)
    return p


def no_borders(table):
    tbl = table._tbl
    borders = OxmlElement("w:tblBorders")
    for b in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{b}")
        e.set(qn("w:val"), "nil")
        borders.append(e)
    tbl.tblPr.append(borders)


def cant_split(row):
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))


def build(path, *, good: bool):
    doc = Document()
    setup(doc, font="Times New Roman" if good else "Times New Roman", size=14,
          margins=(2.0, 2.0, 3.0, 1.5) if good else (2.0, 2.0, 2.0, 2.0))
    c = AL.CENTER

    # ---- đầu văn bản
    t = doc.add_table(rows=1, cols=2)
    no_borders(t)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    L, R = t.rows[0].cells
    cell_par(L, "NGÂN HÀNG CHÍNH SÁCH XÃ HỘI", True, bold=True, size=13, align=c)
    cell_par(L, "PHÒNG GIAO DỊCH HUYỆN PHÙ CỪ", bold=True, size=13, align=c)
    if good:
        cell_par(R, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", True, bold=True, size=13, align=c)
        cell_par(R, "Độc lập - Tự do - Hạnh phúc", bold=True, size=14, align=c)
        cell_par(L, "Số: 05/NHCS-TD", size=13, align=c)
        cell_par(R, f"Phù Cừ, ngày{NB}05{NB}tháng{NB}10{NB}năm{NB}2026", italic=True, size=14, align=c)
    else:
        cell_par(R, "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM", True, bold=True, size=13, align=c)
        cell_par(R, "Độc lập – tự do – hạnh phúc", size=14, align=c)
        cell_par(L, "Số 5/nhcs-td", size=13, align=c)
        cell_par(R, "Phù Cừ, ngày 5 tháng 03 năm 2026", size=14, align=c)

    para(doc, "")
    if good:
        para(doc, "BÁO CÁO", bold=True, size=14, align=c, after=0, keep_next=True)
        para(doc, "Kết quả kiểm tra sử dụng vốn vay tại các xã trên địa bàn huyện", bold=True, size=14, align=c, keep_next=True)
        para(doc, "Kính gửi: Ủy ban nhân dân huyện Phù Cừ", size=14, align=c)
        body = [
            "Thực hiện kế hoạch kiểm tra hằng năm, Phòng giao dịch huyện Phù Cừ đã kiểm tra việc sử dụng vốn vay "
            f"của hộ nghèo tại các xã theo Nghị định số{NB}30/2020/NĐ-CP và Điều{NB}5 của Quy chế cho vay. Kết quả cụ thể như sau:",
            "Tổng dư nợ đến ngày 30 tháng 9 năm 2026 là 98.835.000 đồng, lãi suất 0,66%/tháng; không có hộ vay nào sử dụng vốn sai mục đích.",
            "Phòng giao dịch sẽ tiếp tục phối hợp với Ủy ban nhân dân các xã, các tổ chức hội, đoàn thể để thực hiện tốt công tác này.",
        ]
        for b in body:
            para(doc, b, size=14, align=AL.JUSTIFY, indent=1.27)
    else:
        para(doc, "báo cáo", size=14, align=AL.LEFT, after=0)
        para(doc, "V/v kiểm tra sử dụng vốn vay  tại các xã .", size=14, align=c)
        para(doc, "Kính gửi Ủy Ban Nhân Dân huyện Phù Cừ", size=14, align=AL.LEFT)
        body = [
            "    Thực hiện kế hoạch kiểm tra hàng năm, Ngân hàng Chính sách Xã hội huyện Phù Cừ đã kiểm tra việc sử lý vốn vay của "
            "Hộ Nghèo,tại các xã theo nghị định số 30/2020/NĐ - CP và điều 5 của Quy chế cho vay(gọi tắt là NHCSXH ) nhằm bổ xung kiểm soát.",
            "Tổng dư nợ là 98,835,000 đồng, lãi suất 0.66%/tháng.Không có hộ vay nào ko trả nợ nghành tín dụng. theo báo cáo và và đánh giá chung.",
            "Hoà bình và hòa thuận, kí tên rồi ký nhận Nguyễn VănAn gởi cho giám đốc.",
        ]
        for b in body:
            p = para(doc, b, size=12, align=AL.LEFT)
        para(doc, ORPHAN_TEXT, size=14, align=AL.JUSTIFY, indent=1.27)
    para(doc, "")

    # ---- chữ ký + nơi nhận
    sig = doc.add_table(rows=1, cols=2)
    no_borders(sig)
    if good:
        cant_split(sig.rows[0])
    A, B = sig.rows[0].cells
    cell_par(A, "Nơi nhận:" if good else "Nơi nhận", True, bold=True, italic=True, size=12)
    if good:
        for ln in ("- Như trên;", "- Ban giám đốc;", "- Lưu: VT, TD."):
            cell_par(A, ln, size=11)
        cell_par(B, "KT. GIÁM ĐỐC", True, bold=True, size=14, align=c)
        cell_par(B, "PHÓ GIÁM ĐỐC", bold=True, size=14, align=c)
        cell_par(B, "")
        cell_par(B, "Nguyễn Văn An", bold=True, size=14, align=c)
    else:
        for ln in ("• Như trên", "- Ban giám đốc", "- Lưu VT"):
            cell_par(A, ln, size=11)
        cell_par(B, "KT GIÁM ĐỐC", True, size=14, align=c)
        cell_par(B, "PHÓ GIÁM ĐỐC", size=14, align=c)
        cell_par(B, "Nguyễn Văn An", size=14, align=c)
    doc.save(path)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "mau")
    os.makedirs(out, exist_ok=True)
    build(os.path.join(out, "mau_dung.docx"), good=True)
    build(os.path.join(out, "mau_loi.docx"), good=False)
    print("Đã tạo", out)

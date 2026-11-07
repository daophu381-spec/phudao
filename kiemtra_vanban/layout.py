"""Phát hiện lỗi bố cục sau khi dàn trang: chữ một mình một dòng (dòng mồ côi), cụm từ bị tách dòng,
tiêu đề mục rơi cuối trang, chữ ký rơi sang trang sau...

File Word không lưu chỗ ngắt dòng nên phải chuyển sang PDF (LibreOffice, hoặc Microsoft Word qua docx2pdf)
rồi đo vị trí từng dòng chữ bằng pdfplumber.
"""
from __future__ import annotations

import os
import re
import shutil
import statistics
import subprocess
import tempfile
from dataclasses import dataclass, field

from .models import CAT_LAYOUT, Issue, SEV_ERROR, SEV_HINT, SEV_WARN, short


@dataclass
class Seg:
    page: int
    top: float
    bottom: float
    x0: float
    x1: float
    words: list = field(default_factory=list)   # [(text, x0, x1, bold)]
    size: float = 12.0
    line_no: int = 0

    @property
    def text(self):
        return " ".join(w[0] for w in self.words)

    @property
    def cx(self):
        return (self.x0 + self.x1) / 2

    @property
    def bold(self):
        return bool(self.words) and sum(w[3] for w in self.words) >= len(self.words) * 0.7


class LayoutUnavailable(Exception):
    pass


# ---------------------------------------------------------------- chuyển DOCX -> PDF
def docx_to_pdf(path: str, outdir: str) -> str:
    """Trả đường dẫn PDF. Thử LibreOffice, sau đó Microsoft Word (docx2pdf)."""
    out_pdf = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    soffice = shutil.which("soffice") or shutil.which("libreoffice") or shutil.which("soffice.exe")
    if soffice is None:
        for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if os.path.exists(cand):
                soffice = cand
    if soffice:
        prof = tempfile.mkdtemp(prefix="lo_profile_")
        try:
            subprocess.run([soffice, f"-env:UserInstallation=file:///{prof.replace(os.sep, '/').lstrip('/')}", "--headless",
                            "--convert-to", "pdf", "--outdir", outdir, path],
                           check=True, capture_output=True, timeout=180)
            if os.path.exists(out_pdf):
                return out_pdf
        except Exception:
            pass
        finally:
            shutil.rmtree(prof, ignore_errors=True)
    try:
        from docx2pdf import convert      # cần cài Microsoft Word
        convert(path, out_pdf)
        if os.path.exists(out_pdf):
            return out_pdf
    except Exception:
        pass
    raise LayoutUnavailable("Không chuyển được sang PDF: cần cài LibreOffice (miễn phí) hoặc Microsoft Word + 'pip install docx2pdf'")


# ---------------------------------------------------------------- đọc PDF thành các dòng
def read_segments(pdf_path: str) -> tuple[list[Seg], list[tuple[float, float]]]:
    try:
        import pdfplumber
    except ImportError as e:
        raise LayoutUnavailable("Thiếu thư viện pdfplumber: chạy 'pip install pdfplumber'") from e
    segs: list[Seg] = []
    sizes = []
    with pdfplumber.open(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            sizes.append((float(page.width), float(page.height)))
            words = page.extract_words(extra_attrs=["size", "fontname"], keep_blank_chars=False, x_tolerance=1.5, y_tolerance=2)
            words.sort(key=lambda w: (round(w["top"]), w["x0"]))
            lines: list[list[dict]] = []
            for w in words:
                if lines and abs(lines[-1][0]["top"] - w["top"]) <= max(2.5, w["size"] * 0.3):
                    lines[-1].append(w)
                else:
                    lines.append([w])
            ln = 0
            for line in lines:
                line.sort(key=lambda w: w["x0"])
                gap_limit = max(26.0, statistics.median(w["size"] for w in line) * 2.2)
                cur = [line[0]]
                groups = []
                for w in line[1:]:
                    if w["x0"] - cur[-1]["x1"] > gap_limit:
                        groups.append(cur)
                        cur = [w]
                    else:
                        cur.append(w)
                groups.append(cur)
                for g in groups:
                    ln += 1
                    segs.append(Seg(pno, min(w["top"] for w in g), max(w["bottom"] for w in g),
                                    min(w["x0"] for w in g), max(w["x1"] for w in g),
                                    [(w["text"], w["x0"], w["x1"], "bold" in w["fontname"].lower()) for w in g],
                                    statistics.median(w["size"] for w in g), ln))
    return segs, sizes


# ---------------------------------------------------------------- phân tích
INDENT_TOL = 40.0   # cho phép dòng đầu đoạn lùi vào (1–1,27 cm ≈ 28–36 pt)
TIGHT_PREV = {"điều", "khoản", "điểm", "chương", "mục", "ngày", "tháng", "năm", "số", "số:", "lúc", "giờ", "quý"}
UNITS = {"đồng", "triệu", "tỷ", "nghìn", "%", "vnđ", "vnd", "hộ", "xã", "km", "kg", "người", "tổ"}


def analyze(segs: list[Seg], sizes: list[tuple[float, float]], max_orphan_words: int = 1) -> list[Issue]:
    issues: list[Issue] = []
    if not segs:
        return issues
    body = [s for s in segs if len(s.words) >= 6]
    if len(body) < 3:
        body = [s for s in segs if len(s.words) >= 3] or segs
    xs0 = sorted(s.x0 for s in body)
    xs1 = sorted(s.x1 for s in body)
    x_left = xs0[int(len(xs0) * 0.1)]
    x_right = xs1[int(len(xs1) * 0.9)]
    width = max(x_right - x_left, 1)

    # độ cao dòng thường gặp
    by_page: dict[int, list[Seg]] = {}
    for s in segs:
        by_page.setdefault(s.page, []).append(s)
    pitches = []
    for pg, ss in by_page.items():
        ss = sorted(ss, key=lambda s: (s.x0 // 40, s.top))
        for a, b in zip(ss, ss[1:]):
            d = b.top - a.top
            if 8 < d < 40 and abs(a.x0 - b.x0) < 6 and len(a.words) >= 5:
                pitches.append(round(d, 1))
    pitch = statistics.median(pitches) if pitches else 18.0

    def prev_line(s: Seg, ss: list[Seg]):
        best = None
        for q in ss:
            if q is s or q.top >= s.top - 1:
                continue
            dy = s.top - q.top
            if dy > pitch * 1.12 or dy < pitch * 0.5:
                continue
            overlap = min(q.x1, s.x1 + INDENT_TOL) - max(q.x0, s.x0)
            if overlap <= 0:
                continue
            if best is None or q.top > best.top:
                best = q
        return best

    def next_line(s: Seg, ss: list[Seg]):
        best = None
        for q in ss:
            if q is s or q.top <= s.top + 1:
                continue
            dy = q.top - s.top
            if dy > pitch * 1.12 or dy < pitch * 0.5:
                continue
            overlap = min(q.x1, s.x1 + INDENT_TOL) - max(q.x0, s.x0)
            if overlap <= 0 and abs(q.x0 - s.x0) > 30:
                continue
            if best is None or q.top < best.top:
                best = q
        return best

    for pg, ss in sorted(by_page.items()):
        ph = sizes[pg - 1][1]
        ss = sorted(ss, key=lambda s: (s.top, s.x0))
        for s in ss:
            if s.top < 45 or s.bottom > ph - 45:
                continue                      # đầu trang/chân trang (số trang)
            txt = s.text.strip(" .,;:")
            if not txt or re.fullmatch(r"[\d.\-–]+", txt):
                continue
            pl = prev_line(s, ss)
            nw = len(s.words)
            few = nw <= max_orphan_words and (nw == 1 or len(s.text) <= 12)
            if few and pl is not None and len(pl.words) >= 2:
                reaches_right = pl.x1 >= x_right - 0.05 * width - 3
                at_left = s.x0 <= x_left + 6 or abs(s.x0 - pl.x0) <= 6
                centered = abs(s.cx - pl.cx) < 6 and (pl.x1 - pl.x0) >= 0.45 * width and len(pl.words) >= 3
                if reaches_right and at_left:
                    issues.append(Issue(CAT_LAYOUT, SEV_ERROR, f"Trang {pg}, dòng {s.line_no}", f"…{pl.text[-35:]} ⏎ [{s.text}]",
                                        "Chữ một mình một dòng (dòng cuối đoạn chỉ còn 1 chữ)",
                                        "Rút gọn/đổi câu cho gọn lại, hoặc nối chữ cuối với chữ này bằng dấu cách không ngắt (Ctrl+Shift+Space)",
                                        pg * 1000 + s.line_no))
                elif centered:
                    issues.append(Issue(CAT_LAYOUT, SEV_WARN, f"Trang {pg}, dòng {s.line_no}", f"{pl.text} ⏎ [{s.text}]",
                                        "Tiêu đề/trích yếu bị rơi chữ: dòng dưới chỉ còn 1 chữ",
                                        "Ngắt dòng chủ động bằng Shift+Enter ở ranh giới cụm từ, hai dòng cân đối",
                                        pg * 1000 + s.line_no))
            # cụm từ bị tách dòng: "ngày | 05", "Điều | 5", "98.835.000 | đồng"
            nl = next_line(s, ss)
            if nl is not None and s.words and nl.words and s.x1 >= x_right - 0.08 * width:
                last = s.words[-1][0].lower().strip(",.;:")
                first = nl.words[0][0].lower()
                broken = None
                if last in TIGHT_PREV and re.match(r"^(\d|[ivx]+\b|[a-zđ]\)|\.\.\.|…)", first):
                    broken = f"'{s.words[-1][0]}' | '{nl.words[0][0]}'"
                elif re.fullmatch(r"[\d.,]+", last) and first.strip(".,;") in UNITS:
                    broken = f"'{s.words[-1][0]}' | '{nl.words[0][0]}'"
                if broken:
                    issues.append(Issue(CAT_LAYOUT, SEV_WARN, f"Trang {pg}, dòng {s.line_no}", f"…{s.text[-28:]} ⏎ {nl.text[:20]}…",
                                        f"Cụm từ không nên tách dòng: {broken}",
                                        "Nối bằng dấu cách không ngắt (Ctrl+Shift+Space)", pg * 1000 + s.line_no))

    issues += _page_level(segs, by_page, sizes)
    issues += _two_line_titles(segs, by_page, x_left, width, pitch)
    return issues


def _two_line_titles(segs, by_page, x_left, width, pitch):
    out = []
    for pg, ss in sorted(by_page.items()):
        ss = sorted(ss, key=lambda s: (s.top, s.x0))
        for a in ss:
            if not a.bold or len(a.words) < 2 or (a.x1 - a.x0) < 0.35 * width:
                continue
            for b in ss:
                if b is a or not (pitch * 0.5 < b.top - a.top < pitch * 1.12) or not b.bold:
                    continue
                if abs(a.cx - b.cx) > 6:
                    continue
                if (b.x1 - b.x0) > (a.x1 - a.x0) * 1.12 and len(b.words) >= 3:
                    # chỉ báo khi đúng 2 dòng (không có dòng thứ 3 cùng khối)
                    third = [c for c in ss if c is not a and c is not b and pitch * 0.5 < c.top - b.top < pitch * 1.12
                             and abs(c.cx - b.cx) <= 6 and c.bold]
                    if not third:
                        out.append(Issue(CAT_LAYOUT, SEV_HINT, f"Trang {pg}, dòng {a.line_no}", f"{a.text} ⏎ {b.text}",
                                         "Tiêu đề/trích yếu 2 dòng: dòng dưới dài hơn dòng trên (nên để dòng trên dài hơn hoặc bằng)",
                                         "Ngắt lại bằng Shift+Enter ở ranh giới cụm từ", pg * 1000 + a.line_no))
    return out


def _page_level(segs, by_page, sizes):
    out = []
    n = len(by_page)
    pages = sorted(by_page)
    real = {pg: [s for s in by_page[pg] if sizes[pg - 1][1] - 45 > s.bottom and s.top > 45 and not re.fullmatch(r"[\d\s.\-–]+", s.text.strip())]
            for pg in pages}
    if n > 1:
        last = pages[-1]
        cnt = len({round(s.top) for s in real[last]})
        if 0 < cnt <= 3:
            txt = " ⏎ ".join(s.text for s in sorted(real[last], key=lambda s: (s.top, s.x0)))
            out.append(Issue(CAT_LAYOUT, SEV_ERROR, f"Trang {last}", short(txt, 0, 60),
                             f"Trang cuối chỉ có {cnt} dòng (rơi trang): nội dung hoặc khối chữ ký bị đẩy sang trang mới",
                             "Rút gọn nội dung/giãn dòng để dồn về trang trước, hoặc đưa thêm nội dung sang trang cuối cho cân", last * 1000))
        if cnt == 0:
            out.append(Issue(CAT_LAYOUT, SEV_ERROR, f"Trang {last}", "", "Trang cuối trống (dòng trống hoặc ngắt trang thừa)", "Xóa dòng trống/ngắt trang thừa", last * 1000))
    # tiêu đề mục cuối trang
    for pg in pages[:-1]:
        ss = sorted(real[pg], key=lambda s: (s.top, s.x0))
        if not ss:
            continue
        lastline = ss[-1]
        if re.match(r"^(?:[IVX]+|\d+(?:\.\d+)*)[.)]\s+\S", lastline.text) and len(lastline.text) < 90 and lastline.bold:
            out.append(Issue(CAT_LAYOUT, SEV_ERROR, f"Trang {pg}", lastline.text,
                             "Tiêu đề mục nằm trơ trọi ở cuối trang, nội dung của nó sang trang sau",
                             "Đặt 'Keep with next' cho tiêu đề, hoặc ngắt trang trước tiêu đề", pg * 1000 + lastline.line_no))
    # ngày tháng và chữ ký khác trang
    date_pg = [s.page for s in segs if re.search(r"ngày\s+\S*\s*tháng\s+\S*\s*năm", s.text)]
    sign_pg = [s.page for s in segs if re.match(r"^\(\s*Ký", s.text)]
    noinhan_pg = [s.page for s in segs if re.match(r"^Nơi nhận", s.text)]
    if date_pg and sign_pg and date_pg[-1] != sign_pg[-1]:
        out.append(Issue(CAT_LAYOUT, SEV_ERROR, f"Trang {date_pg[-1]} và {sign_pg[-1]}", "",
                         "Dòng ngày tháng và chữ ký nằm khác trang: khối chữ ký bị tách",
                         "Khóa khối chữ ký (cantSplit, keep together) và dồn nội dung về một trang", sign_pg[-1] * 1000))
    if noinhan_pg and sign_pg and noinhan_pg[-1] != sign_pg[-1]:
        out.append(Issue(CAT_LAYOUT, SEV_WARN, f"Trang {noinhan_pg[-1]} và {sign_pg[-1]}", "",
                         "'Nơi nhận' và chữ ký nằm khác trang", "Đặt cùng trang (cùng một hàng bảng, cantSplit)", sign_pg[-1] * 1000))
    return out


def analyze_docx(path: str, max_orphan_words: int = 1, keep_pdf: str | None = None):
    """Chuyển DOCX sang PDF và phân tích. Trả (issues, segs, pdf_path)."""
    tmp = tempfile.mkdtemp(prefix="nd30_")
    try:
        pdf = docx_to_pdf(path, tmp)
        segs, sizes = read_segments(pdf)
        issues = analyze(segs, sizes, max_orphan_words)
        if keep_pdf:
            shutil.copy(pdf, keep_pdf)
        return issues, segs, len(sizes)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def analyze_pdf(path: str, max_orphan_words: int = 1):
    segs, sizes = read_segments(path)
    return analyze(segs, sizes, max_orphan_words), segs, len(sizes)

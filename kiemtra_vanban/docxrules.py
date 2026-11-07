"""Đọc file .docx và kiểm tra thể thức, phông chữ, cỡ chữ, lề theo Nghị định 30/2020/NĐ-CP."""
from __future__ import annotations

import re
import unicodedata
from collections import Counter

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH as ALIGN
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .models import (CAT_FONT, CAT_FORM, CAT_PUNCT, Issue, Para, SEV_ERROR, SEV_HINT, SEV_WARN, short)

NBSP = " "
FONT_OK = "times new roman"
DOC_TYPE_NAMES = ["TỜ TRÌNH", "THÔNG BÁO", "BÁO CÁO", "KẾ HOẠCH", "QUYẾT ĐỊNH", "BIÊN BẢN", "CHƯƠNG TRÌNH",
                  "HƯỚNG DẪN", "NGHỊ QUYẾT", "CHỈ THỊ", "THÔNG TƯ", "ĐỀ ÁN", "PHƯƠNG ÁN", "GIẤY MỜI",
                  "GIẤY GIỚI THIỆU", "GIẤY ỦY QUYỀN", "ĐỀ CƯƠNG", "QUY CHẾ", "QUY ĐỊNH", "NỘI QUY"]
BULLETS = ("-", "–", "—", "•", "+", "*", "·", "●")


def fold(s: str) -> str:
    """Bỏ dấu, bỏ ký tự không phải chữ/số, viết thường: để so khớp chịu được sai dấu."""
    s = unicodedata.normalize("NFD", s.replace("đ", "d").replace("Đ", "D"))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace(NBSP, " ")).strip()


# ---------------------------------------------------------------- đọc và làm phẳng
class DocxFile:
    def __init__(self, path: str):
        self.doc = Document(path)
        self.paras: list[Para] = []
        self._flatten()

    def _flatten(self):
        n = [0]
        t = [0]

        def add_par(el, parent, in_table, where):
            n[0] += 1
            par = Paragraph(el, parent)
            self.paras.append(Para(len(self.paras), par.text, where.format(n=n[0]), in_table, par))

        def walk_table(tbl_el, parent, label):
            t[0] += 1
            tno = t[0]
            tbl = Table(tbl_el, parent)
            seen = set()
            for ri, row in enumerate(tbl.rows, 1):
                for ci, cell in enumerate(row.cells, 1):
                    if id(cell._tc) in seen:
                        continue
                    seen.add(id(cell._tc))
                    for ch in cell._tc.iterchildren():
                        if ch.tag == qn("w:p"):
                            add_par(ch, cell, True, f"Bảng {tno}, hàng {ri}, cột {ci}")
                        elif ch.tag == qn("w:tbl"):
                            walk_table(ch, cell, label)

        body = self.doc.element.body
        for ch in body.iterchildren():
            if ch.tag == qn("w:p"):
                add_par(ch, self.doc, False, "Đoạn {n}")
            elif ch.tag == qn("w:tbl"):
                walk_table(ch, self.doc, "")

    # ----- phân giải thuộc tính qua chuỗi style
    def _style_chain(self, style):
        while style is not None:
            yield style
            style = style.base_style

    def _defaults(self, kind: str):
        dd = self.doc.styles.element.find(qn("w:docDefaults"))
        if dd is None:
            return None
        return dd.find(qn(f"w:{kind}Default"))

    def run_font(self, run: Run, par: Paragraph):
        if run.font.name:
            return run.font.name
        if run.style is not None:
            for s in self._style_chain(run.style):
                if s.font.name:
                    return s.font.name
        for s in self._style_chain(par.style):
            if s.font.name:
                return s.font.name
        d = self._defaults("rPr")
        if d is not None:
            rf = d.find(".//" + qn("w:rFonts"))
            if rf is not None:
                return rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
        return None

    def run_size(self, run: Run, par: Paragraph):
        if run.font.size:
            return run.font.size.pt
        if run.style is not None:
            for s in self._style_chain(run.style):
                if s.font.size:
                    return s.font.size.pt
        for s in self._style_chain(par.style):
            if s.font.size:
                return s.font.size.pt
        d = self._defaults("rPr")
        if d is not None:
            sz = d.find(".//" + qn("w:sz"))
            if sz is not None:
                return int(sz.get(qn("w:val"))) / 2
        return None

    def run_flag(self, run: Run, par: Paragraph, attr: str) -> bool:
        v = getattr(run, attr)
        if v is not None:
            return bool(v)
        if run.style is not None:
            for s in self._style_chain(run.style):
                v = getattr(s.font, attr)
                if v is not None:
                    return bool(v)
        for s in self._style_chain(par.style):
            v = getattr(s.font, attr)
            if v is not None:
                return bool(v)
        return False

    def runs(self, par: Paragraph):
        out = []
        for r in par._p.iter(qn("w:r")):
            run = Run(r, par)
            if run.text and run.text.strip():
                out.append(run)
        return out

    def par_prop(self, par: Paragraph, prop: str):
        v = getattr(par.paragraph_format, prop)
        if v is not None:
            return v
        for s in self._style_chain(par.style):
            v = getattr(s.paragraph_format, prop)
            if v is not None:
                return v
        return None

    def all_bold(self, par):
        rs = self.runs(par)
        return bool(rs) and all(self.run_flag(r, par, "bold") for r in rs)

    def all_italic(self, par):
        rs = self.runs(par)
        return bool(rs) and all(self.run_flag(r, par, "italic") for r in rs)

    def sizes(self, par):
        return {self.run_size(r, par) for r in self.runs(par)} - {None}

    def alignment(self, par):
        return self.par_prop(par, "alignment")


def _agg(out: list[Issue], items: list[Issue], limit: int = 4, label: str = ""):
    """Gom nhiều lỗi cùng loại: ít thì liệt kê, nhiều thì gộp thành 1 dòng."""
    if len(items) <= limit:
        out.extend(items)
        return
    first = items[0]
    where = ", ".join(i.where for i in items[:limit]) + f" … (tổng {len(items)} chỗ)"
    out.append(Issue(first.category, first.severity, where, first.found, first.message, first.fix, first.order))


# ---------------------------------------------------------------- kiểm tra
def check_docx(f: DocxFile) -> list[Issue]:
    out: list[Issue] = []
    out += check_page(f)
    out += check_fonts(f)
    out += check_body_format(f)
    out += check_header_block(f)
    out += check_title(f)
    out += check_closing(f)
    out += check_misc(f)
    return out


def check_page(f: DocxFile) -> list[Issue]:
    out = []
    note = " (nếu là biểu mẫu cấp trên ban hành thì giữ nguyên, chỉ ghi nhận)"
    for si, sec in enumerate(f.doc.sections, 1):
        w, h = sec.page_width, sec.page_height
        if w is None or h is None:
            continue
        wc, hc = w.cm, h.cm
        if abs(wc - 21.0) > 0.3 or abs(hc - 29.7) > 0.3:
            out.append(Issue(CAT_FONT, SEV_WARN, f"Trang (phần {si})", f"{wc:.1f} × {hc:.1f} cm",
                             "Khổ giấy phải là A4 đặt dọc (21 × 29,7 cm)", "A4 dọc", 0))
        spec = [("trên", sec.top_margin, 20, 25), ("dưới", sec.bottom_margin, 20, 25),
                ("trái", sec.left_margin, 30, 35), ("phải", sec.right_margin, 15, 20)]
        for name, val, lo, hi in spec:
            if val is None:
                continue
            mm = val.mm
            if mm < lo - 1.0 or mm > hi + 1.0:
                out.append(Issue(CAT_FONT, SEV_HINT, f"Trang (phần {si})", f"lề {name} {mm:.0f} mm",
                                 f"Lề {name} nên trong khoảng {lo}–{hi} mm theo NĐ30" + note, f"{lo}–{hi} mm", 0))
    return out


def check_fonts(f: DocxFile) -> list[Issue]:
    out: list[Issue] = []
    bad_font, bad_color, unresolved = [], [], 0
    size_counter: Counter = Counter()
    per_par_sizes = {}
    for p in f.paras:
        par = p.obj
        fonts, colors = set(), set()
        for r in f.runs(par):
            name = f.run_font(r, par)
            if name is None:
                unresolved += 1
            elif name.lower() != FONT_OK:
                fonts.add(name)
            try:
                if r.font.color is not None and r.font.color.rgb is not None and str(r.font.color.rgb) not in ("000000",):
                    colors.add(str(r.font.color.rgb))
            except Exception:
                pass
            sz = f.run_size(r, par)
            if sz:
                per_par_sizes.setdefault(p.index, set()).add(sz)
                if not p.in_table:
                    size_counter[sz] += len(r.text.strip())
        if fonts:
            bad_font.append(Issue(CAT_FONT, SEV_ERROR, p.where, short(p.text, 0, 20),
                                  f"Phông chữ không phải Times New Roman: {', '.join(sorted(fonts))}", "Times New Roman", p.index))
        if colors:
            bad_color.append(Issue(CAT_FONT, SEV_WARN, p.where, short(p.text, 0, 20),
                                   f"Chữ không màu đen (mã màu {', '.join(sorted(colors))})", "Màu đen (Automatic)", p.index))
    _agg(out, bad_font, 5)
    _agg(out, bad_color, 5)

    if size_counter:
        dom = size_counter.most_common(1)[0][0]
        if dom not in (13, 13.0, 14, 14.0):
            out.append(Issue(CAT_FONT, SEV_WARN, "Toàn văn bản", f"Cỡ chữ nội dung chủ yếu là {dom:g}",
                             "Cỡ chữ nội dung phải là 13 hoặc 14", "13 hoặc 14", 0))
        mixed = []
        for p in f.paras:
            if p.in_table or len(p.text.strip()) < 20:
                continue
            sz = per_par_sizes.get(p.index, set())
            if sz and (len(sz) > 1 or (dom not in sz)):
                # trích yếu, nơi nhận, ghi chú cỡ nhỏ hơn không tính nếu là dòng ngắn đứng riêng
                mixed.append(Issue(CAT_FONT, SEV_WARN, p.where, short(p.text, 0, 24),
                                   "Cỡ chữ lệch so với phần còn lại của nội dung: " + ", ".join(f"{s:g}" for s in sorted(sz))
                                   + f" (nội dung chính {dom:g})", f"{dom:g}", p.index))
        _agg(out, mixed, 5)
    if unresolved and not any(i.message.startswith("Phông") for i in out):
        fn = Counter(f.run_font(r, p.obj) for p in f.paras for r in f.runs(p.obj))
        if not any((k or "").lower() == FONT_OK for k in fn):
            out.append(Issue(CAT_FONT, SEV_WARN, "Toàn văn bản", "Phông chữ theo chủ đề (theme)",
                             "Không xác định được phông chữ vì file dùng phông theo theme (có thể là Calibri/Aptos)",
                             "Chọn toàn bộ, đặt Times New Roman", 0))
    return out


def check_body_format(f: DocxFile) -> list[Issue]:
    out: list[Issue] = []
    not_just, no_indent, no_space, bad_ls = [], [], [], []
    for p in f.paras:
        txt = p.text.strip()
        if p.in_table or len(txt) < 80 or txt.startswith(BULLETS):
            continue
        par = p.obj
        al = f.alignment(par)
        if al not in (ALIGN.JUSTIFY, ALIGN.DISTRIBUTE):
            not_just.append(Issue(CAT_FONT, SEV_HINT, p.where, short(p.text, 0, 24),
                                  "Đoạn văn nên căn đều hai lề (Justify)", "Căn đều hai lề", p.index))
        fi = f.par_prop(par, "first_line_indent")
        left = f.par_prop(par, "left_indent")
        if not re.match(r"^[ \t" + NBSP + "]", p.text):
            cm = fi.cm if fi is not None else 0
            if (fi is None or cm < 0.9 or cm > 1.4) and not (left and left.cm > 0.9 and cm == 0 and False):
                no_indent.append(Issue(CAT_FONT, SEV_HINT, p.where, short(p.text, 0, 24),
                                       f"Chưa lùi đầu dòng đúng 1–1,27 cm (đang {cm:.2f} cm)", "Lùi đầu dòng 1 cm hoặc 1,27 cm", p.index))
        sb, sa = f.par_prop(par, "space_before"), f.par_prop(par, "space_after")
        tot = (sb.pt if sb is not None else 0) + (sa.pt if sa is not None else 0)
        if tot < 6:
            no_space.append(Issue(CAT_FONT, SEV_HINT, p.where, short(p.text, 0, 24),
                                  f"Khoảng cách giữa các đoạn nhỏ hơn 6 pt (đang {tot:g} pt)", "Spacing ≥ 6 pt", p.index))
        ls = f.par_prop(par, "line_spacing")
        if isinstance(ls, float) and not (0.99 <= ls <= 1.51):
            bad_ls.append(Issue(CAT_FONT, SEV_WARN, p.where, short(p.text, 0, 24),
                                f"Giãn dòng {ls:g} lines nằm ngoài khoảng 1,0 đến 1,5", "Từ đơn đến 1,5 lines", p.index))
    for lst in (not_just, no_indent, no_space, bad_ls):
        _agg(out, lst, 3)
    return out


def check_header_block(f: DocxFile) -> list[Issue]:
    out: list[Issue] = []
    found_qh = found_tn = False
    for p in f.paras:
        txt = norm_ws(p.text)
        fz = fold(txt)
        par = p.obj
        if "conghoaxahoichunghiavietnam" in fz and len(fz) < 40:
            found_qh = True
            if txt != "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM":
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Quốc hiệu phải viết IN HOA đúng chuẩn (chú ý dấu 'HÒA')",
                                 "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", p.index))
            if not f.all_bold(par):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Quốc hiệu phải in đậm", "In đậm", p.index))
            sz = f.sizes(par)
            if sz and not all(12 <= s <= 13.5 for s in sz):
                out.append(Issue(CAT_FORM, SEV_HINT, p.where, txt, "Quốc hiệu cỡ chữ 12–13", "12 hoặc 13", p.index))
        elif "doclaptudohanhphuc" in fz and len(fz) < 30:
            found_tn = True
            if txt != "Độc lập - Tự do - Hạnh phúc":
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt,
                                 "Tiêu ngữ phải viết 'Độc lập - Tự do - Hạnh phúc' (viết hoa chữ đầu mỗi cụm, gạch nối có cách hai bên)",
                                 "Độc lập - Tự do - Hạnh phúc", p.index))
            if not f.all_bold(par):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Tiêu ngữ phải in đậm", "In đậm", p.index))
            sz = f.sizes(par)
            if sz and not all(13 <= s <= 14.5 for s in sz):
                out.append(Issue(CAT_FORM, SEV_HINT, p.where, txt, "Tiêu ngữ cỡ chữ 13–14", "13 hoặc 14", p.index))
            if f.all_italic(par):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Tiêu ngữ in đứng, không nghiêng", "In đứng", p.index))
        # dòng gạch gõ bằng ký tự
        if re.fullmatch(r"[_\-–—─=\s]{5,}", p.text.strip() or "x"):
            out.append(Issue(CAT_FORM, SEV_WARN, p.where, short(p.text, 0, 20),
                             "Đường kẻ ngang gõ bằng ký tự (___ hoặc ---); phải dùng Shape line hoặc viền dưới của ô/đoạn", "Dùng viền hoặc Shape line", p.index))
        # dòng địa danh, ngày tháng
        if re.search(r"ngày\s*\d{0,2}\s*tháng", txt) and re.match(r"^\S[^\n]{0,40}?[, ]\s*ngày", txt) and len(txt) < 70:
            if not re.match(r"^[^,]+,\s*ngày", txt):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Thiếu dấu phẩy sau địa danh", "Địa danh, ngày … tháng … năm …", p.index))
            if not f.all_italic(par):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Dòng địa danh, ngày tháng phải in nghiêng", "In nghiêng", p.index))
            if re.search(r"ngày [\d…. ]+ tháng|tháng [\d…. ]+ năm", p.text):
                out.append(Issue(CAT_LAYOUT_HINT, SEV_HINT, p.where, txt,
                                 "Nên dùng dấu cách không ngắt (Ctrl+Shift+Space) trong 'ngày … tháng … năm …' để dòng không bị gãy",
                                 "", p.index))
    if not found_qh:
        out.append(Issue(CAT_FORM, SEV_HINT, "Toàn văn bản", "", "Không tìm thấy Quốc hiệu (bỏ qua nếu không phải văn bản hành chính hoàn chỉnh)", "", 0))
    if not found_tn:
        out.append(Issue(CAT_FORM, SEV_HINT, "Toàn văn bản", "", "Không tìm thấy Tiêu ngữ 'Độc lập - Tự do - Hạnh phúc'", "", 0))
    return out


CAT_LAYOUT_HINT = "Rơi chữ, rơi dòng, rơi trang"


def _next_nonempty(f: DocxFile, idx: int):
    for q in f.paras[idx + 1:]:
        if q.text.strip():
            return q
    return None


def check_title(f: DocxFile) -> list[Issue]:
    out: list[Issue] = []
    for p in f.paras:
        txt = norm_ws(p.text).rstrip(":")
        if not txt:
            continue
        par = p.obj
        hit = next((t for t in DOC_TYPE_NAMES if fold(txt) == fold(t)), None)
        if hit:
            if txt != hit:
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Tên loại văn bản phải viết IN HOA", hit, p.index))
            if not f.all_bold(par):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Tên loại văn bản phải in đậm", "In đậm", p.index))
            if f.alignment(par) != ALIGN.CENTER:
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Tên loại văn bản phải căn giữa", "Căn giữa", p.index))
            sz = f.sizes(par)
            if sz and not all(13 <= s <= 14.5 for s in sz):
                out.append(Issue(CAT_FORM, SEV_HINT, p.where, txt, "Tên loại văn bản cỡ chữ 13–14", "13 hoặc 14", p.index))
            ty = _next_nonempty(f, p.index)
            if ty is not None:
                tt = norm_ws(ty.text)
                tp = ty.obj
                if not f.all_bold(tp):
                    out.append(Issue(CAT_FORM, SEV_ERROR, ty.where, short(tt, 0, 30), "Trích yếu phải in đậm", "In đậm", ty.index))
                if f.alignment(tp) != ALIGN.CENTER:
                    out.append(Issue(CAT_FORM, SEV_ERROR, ty.where, short(tt, 0, 30), "Trích yếu phải căn giữa", "Căn giữa", ty.index))
                if f.all_italic(tp):
                    out.append(Issue(CAT_FORM, SEV_ERROR, ty.where, short(tt, 0, 30), "Trích yếu in đứng, không nghiêng", "In đứng", ty.index))
                if tt.endswith("."):
                    out.append(Issue(CAT_FORM, SEV_ERROR, ty.where, short(tt, len(tt) - 20, len(tt)), "Trích yếu không kết thúc bằng dấu chấm", "Bỏ dấu chấm", ty.index))
                if len(tt) > 20 and tt.isupper():
                    out.append(Issue(CAT_FORM, SEV_ERROR, ty.where, short(tt, 0, 30), "Trích yếu viết in thường (chỉ hoa chữ đầu), không IN HOA cả dòng", "", ty.index))
        m = re.match(r"^kính\s*gửi\b\s*(.?)", txt, re.I)
        if m:
            if not txt.startswith("Kính gửi"):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, short(txt, 0, 20), "'Kính gửi' phải viết hoa chữ K, thường các chữ còn lại", "Kính gửi:", p.index))
            if not re.match(r"^kính\s*gửi\s*:", norm_ws(p.text), re.I):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, short(txt, 0, 20), "Thiếu dấu hai chấm sau 'Kính gửi'", "Kính gửi:", p.index))
    return out


def check_closing(f: DocxFile) -> list[Issue]:
    out: list[Issue] = []
    P = f.paras
    for i, p in enumerate(P):
        txt = norm_ws(p.text)
        if re.fullmatch(r"nơi\s*nhận\s*:?", txt, re.I):
            if txt != "Nơi nhận:":
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Phải viết 'Nơi nhận:' (có dấu hai chấm, chỉ hoa chữ N)", "Nơi nhận:", p.index))
            par = p.obj
            if not (f.all_bold(par) and f.all_italic(par)):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "'Nơi nhận:' phải in đậm và nghiêng", "In đậm, nghiêng", p.index))
            sz = f.sizes(par)
            if sz and not all(11.5 <= s <= 12.5 for s in sz):
                out.append(Issue(CAT_FORM, SEV_HINT, p.where, txt, "'Nơi nhận:' cỡ chữ 12", "12", p.index))
            lines = []
            for q in P[i + 1:]:
                t = q.text.strip()
                if not t or q.in_table != p.in_table:
                    break
                if not t.startswith(BULLETS):
                    break
                lines.append(q)
            for k, q in enumerate(lines):
                t = q.text.strip()
                last = k == len(lines) - 1
                if not t.startswith("- "):
                    out.append(Issue(CAT_FORM, SEV_ERROR, q.where, short(t, 0, 24),
                                     "Danh sách nơi nhận phải bắt đầu bằng gạch đầu dòng '- ' (gạch nối + 1 dấu cách)", "- " + t.lstrip("".join(BULLETS) + " "), q.index))
                if not last and not t.endswith(";"):
                    out.append(Issue(CAT_FORM, SEV_ERROR, q.where, short(t, max(0, len(t) - 20), len(t)),
                                     "Mỗi nơi nhận kết thúc bằng dấu chấm phẩy ';'", t.rstrip(".,") + ";", q.index))
                if last and not t.endswith("."):
                    out.append(Issue(CAT_FORM, SEV_ERROR, q.where, short(t, max(0, len(t) - 20), len(t)),
                                     "Dòng cuối của danh sách nơi nhận kết thúc bằng dấu chấm", t.rstrip(";,") + ".", q.index))
                if re.match(r"^[-–•]?\s*Lưu\b", t) and not re.match(r"^[-–•]?\s*Lưu:", t):
                    out.append(Issue(CAT_FORM, SEV_ERROR, q.where, short(t, 0, 24), "Thiếu dấu hai chấm sau 'Lưu'", "- Lưu: VT, …", q.index))
                szs = f.sizes(q.obj)
                if szs and not all(10.5 <= s <= 12 for s in szs):
                    out.append(Issue(CAT_FORM, SEV_HINT, q.where, short(t, 0, 24), "Danh sách nơi nhận cỡ chữ 11", "11", q.index))
                if f.all_italic(q.obj):
                    out.append(Issue(CAT_FORM, SEV_ERROR, q.where, short(t, 0, 24), "Danh sách nơi nhận in đứng, không nghiêng", "In đứng", q.index))
            if lines and not any(re.match(r"^[-–•]?\s*Lưu\s*:", l.text.strip()) for l in lines):
                out.append(Issue(CAT_FORM, SEV_HINT, p.where, "Nơi nhận", "Không thấy dòng '- Lưu: VT, …' ở cuối danh sách nơi nhận", "- Lưu: VT, …", p.index))
            # hàng bảng chứa nơi nhận phải khóa không tách trang
            out += _check_row_lock(f, p)
        m = re.match(r"^(KT|TM|TL|TUQ|Q)(\.?)(?=\s|$)", txt)
        if m and txt[len(m.group(0)):].strip().isupper():
            if m.group(2) != ".":
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, f"Thiếu dấu chấm sau '{m.group(1)}'", f"{m.group(1)}. " + txt[len(m.group(0)):].strip(), p.index))
            if not f.all_bold(p.obj):
                out.append(Issue(CAT_FORM, SEV_ERROR, p.where, txt, "Quyền hạn, chức vụ người ký phải in đậm", "In đậm", p.index))
            out += _check_row_lock(f, p)
        if re.fullmatch(r"\(?\s*[Kk]ý[^)]*\)?", txt) and len(txt) < 40:
            if f.all_italic(p.obj) is False and False:
                pass
    return out


def _check_row_lock(f: DocxFile, p: Para) -> list[Issue]:
    if not p.in_table:
        return []
    tr = p.obj._p.getparent()
    while tr is not None and tr.tag != qn("w:tr"):
        tr = tr.getparent()
    if tr is None:
        return []
    trpr = tr.find(qn("w:trPr"))
    if trpr is None or trpr.find(qn("w:cantSplit")) is None:
        return [Issue(CAT_LAYOUT_HINT, SEV_WARN, p.where, short(p.text, 0, 24),
                      "Hàng bảng chứa chữ ký / nơi nhận chưa khóa 'Không cho ngắt hàng qua trang' (cantSplit): có thể bị rớt sang trang sau",
                      "Table Properties > Row > bỏ chọn 'Allow row to break across pages'", p.index)]
    return []


def check_misc(f: DocxFile) -> list[Issue]:
    out: list[Issue] = []
    # dòng trống liên tiếp
    run = 0
    blanks = []
    for p in f.paras:
        if p.in_table:
            run = 0
            continue
        if not p.text.strip():
            run += 1
            if run == 2:
                blanks.append(Issue(CAT_FORM, SEV_HINT, p.where, "(dòng trống)",
                                    "Dùng nhiều dòng trống liên tiếp để cách đoạn/ép trang; nên dùng Spacing before/after", "Xóa dòng trống thừa", p.index))
        else:
            run = 0
    _agg(out, blanks, 3)
    # tiêu đề mục chưa keep-with-next
    heads = []
    for p in f.paras:
        t = p.text.strip()
        if p.in_table or len(t) > 130 or not re.match(r"^(?:[IVX]+|\d+(?:\.\d+)*|[A-ZĐ])[.)]\s+\S", t):
            continue
        if f.all_bold(p.obj) and not f.par_prop(p.obj, "keep_with_next"):
            heads.append(Issue(CAT_LAYOUT_HINT, SEV_HINT, p.where, short(t, 0, 30),
                               "Tiêu đề mục chưa đặt 'Keep with next': có thể nằm trơ trọi cuối trang", "Paragraph > Line and Page Breaks > Keep with next", p.index))
    _agg(out, heads, 3)
    return out

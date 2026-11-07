"""Luật kiểm tra trên văn bản thuần: chính tả, viết hoa, dấu câu, dấu cách, số tiền, ngày tháng."""
from __future__ import annotations

import calendar
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field

from .models import (CAT_CAP, CAT_FORM, CAT_PUNCT, CAT_SPELL, SEV_ERROR, SEV_HINT, SEV_WARN,
                     Issue, Para, short)
from .vietnamese import check_syllable, to_new_style

NBSP = " "
WORD_RE = re.compile(r"[^\W\d_]+", re.UNICODE)
SKIP_SPAN_RE = re.compile(r"https?://\S+|www\.\S+|\S+@\S+\.\S+|\b[\w-]+\.(?:docx?|xlsx?|pdf|pptx?|vn|com|net|org|gov)\b",
                          re.IGNORECASE)
LOWER = "a-zàáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ"
UPPER = "A-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ"

# ---------------------------------------------------------------- dữ liệu
# Lỗi từ/cụm từ hay gặp (âm tiết riêng lẻ vẫn đúng nên bộ kiểm tra âm tiết không bắt được)
MISSPELL = [
    ("hàng tháng", "hằng tháng"), ("hàng năm", "hằng năm"), ("hàng ngày", "hằng ngày"),
    ("hàng quý", "hằng quý"), ("hàng tuần", "hằng tuần"),
    ("xử dụng", "sử dụng"), ("sử lý", "xử lý"), ("xữ lý", "xử lý"), ("xử lí", "xử lý"),
    ("sắp sếp", "sắp xếp"), ("xắp xếp", "sắp xếp"), ("sát nhập", "sáp nhập"), ("xát nhập", "sáp nhập"),
    ("bổ xung", "bổ sung"), ("lãi xuất", "lãi suất"), ("suất sắc", "xuất sắc"), ("sơ xuất", "sơ suất"),
    ("tập chung", "tập trung"), ("chặt chẻ", "chặt chẽ"), ("xem sét", "xem xét"),
    ("rà xoát", "rà soát"), ("dà soát", "rà soát"), ("dám đốc", "giám đốc"), ("dám sát", "giám sát"),
    ("tuyên chuyền", "tuyên truyền"), ("chử ký", "chữ ký"), ("chỉnh chu", "chỉn chu"),
    ("nhất chí", "nhất trí"), ("triễn khai", "triển khai"), ("bảo hiễm", "bảo hiểm"),
    ("kiễm tra", "kiểm tra"), ("kiễm soát", "kiểm soát"), ("sai xót", "sai sót"),
    ("chính xách", "chính sách"), ("chánh sách", "chính sách"), ("tính dụng", "tín dụng"),
    ("tín dung", "tín dụng"), ("tiết kiểm", "tiết kiệm"), ("tiết kiêm", "tiết kiệm"),
    ("kế hoặc", "kế hoạch"), ("dải ngân", "giải ngân"), ("giãi ngân", "giải ngân"),
    ("dải quyết", "giải quyết"), ("giãi quyết", "giải quyết"), ("chân trọng", "trân trọng"),
    ("í kiến", "ý kiến"), ("khẩn chương", "khẩn trương"), ("xơ kết", "sơ kết"),
    ("dành được", "giành được"), ("dành lấy", "giành lấy"), ("dành thắng lợi", "giành thắng lợi"),
    ("dành quyền", "giành quyền"), ("gởi", "gửi"), ("nghành", "ngành"), ("chổ ở", "chỗ ở"),
    ("xát thực", "sát thực"), ("xát sao", "sát sao"), ("trụ xở", "trụ sở"), ("chụ sở", "trụ sở"),
]
MISSPELL_RE = [(re.compile(r"(?<![\w])" + r"\s+".join(map(re.escape, a.split())) + r"(?![\w])", re.I), a, b)
               for a, b in MISSPELL]

def compile_pairs(pairs):
    return [(re.compile(r"(?<![\w])" + r"\s+".join(map(re.escape, a.split())) + r"(?![\w])", re.I), a, b) for a, b in pairs]


# Gõ tắt kiểu nhắn tin
SLANG = {"ko": "không", "hok": "không", "đc": "được", "dc": "được", "vs": "với", "bik": "biết",
         "wa": "quá", "ntn": "như thế nào", "k": "không"}

# Tên cơ quan: (viết đúng NĐ30, cho phép viết thường hoàn toàn?)
ORG_NAMES = [
    ("Ngân hàng Chính sách xã hội", False),
    ("Ủy ban nhân dân", True),
    ("Hội đồng nhân dân", True),
    ("Hội Nông dân", False),
    ("Hội Liên hiệp Phụ nữ", False),
    ("Hội Cựu chiến binh", False),
    ("Đoàn Thanh niên Cộng sản Hồ Chí Minh", False),
    ("Mặt trận Tổ quốc Việt Nam", False),
    ("Chính phủ", True),
    ("Quốc hội", True),
]


def _org_regex(name: str) -> re.Pattern:
    parts = []
    for w in name.split():
        w2 = w.replace("Ủy", "(?:ủy|uỷ)").replace("ủy", "(?:ủy|uỷ)")
        parts.append(w2 if "(?:" in w2 else re.escape(w))
    return re.compile(r"(?<![\w])" + r"\s+".join(parts) + r"(?![\w])", re.I)


ORG_RE = [(_org_regex(n), n, lo) for n, lo in ORG_NAMES]

COMMON_NOUNS = {"hộ", "nghèo", "vay", "vốn", "nợ"}   # không được viết hoa tùy tiện giữa câu
ADMIN_UNITS = ["Thành phố", "Thị trấn", "Tỉnh", "Huyện", "Xã", "Thôn", "Xóm", "Phường", "Quận", "Làng"]
DOC_TYPES = ["Nghị định", "Quyết định", "Thông tư", "Luật", "Nghị quyết", "Chỉ thị", "Công văn",
             "Tờ trình", "Kế hoạch", "Báo cáo", "Thông báo", "Hướng dẫn", "Biên bản", "Chương trình"]
FUNC_DUP = {"của", "và", "các", "những", "được", "trong", "tại", "với", "để", "về", "theo", "này"}

WHITELIST = {"kg", "km", "cm", "mm", "ml", "ha", "tr", "kw", "kwh", "mb", "gb", "tb", "vnd", "vnđ", "usd",
             "ok", "pdf", "docx", "xlsx", "tt", "tp", "ts", "ths", "gs", "pgs", "bs", "stt", "vt", "qđ",
             "ttr", "cv", "tb", "ubnd", "hđnd", "nhcs", "nhcsxh", "ndcp", "nđ", "cp", "tw", "cn", "tl",
             "tm", "kt", "tuq", "kh", "đk", "đt", "sđt", "cmnd", "cccd", "mst", "stk", "atm", "sms",
             "internet", "email", "online", "app", "bidv", "agribank", "vietinbank", "vpbank"}
ROMAN_RE = re.compile(r"^(?=[ivxlcdm]+$)m*(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$", re.I)

IY_PAIRS = [("kí", "ký"), ("kĩ", "kỹ"), ("lí", "lý"), ("mĩ", "mỹ"), ("quí", "quý"), ("qui", "quy"),
            ("sĩ", "sỹ")]


@dataclass
class Options:
    kieu_dau: str = "moi"          # moi | cu | bo-qua
    whitelist: set = field(default_factory=set)
    check_spelling: bool = True
    extra_misspell: list = field(default_factory=list)   # [(sai, đúng)] do người dùng tự thêm


@dataclass
class DocObservations:
    tone: list = field(default_factory=list)    # (Issue-tham-chiếu-tạm, style, fixed)
    iy: Counter = field(default_factory=Counter)
    iy_where: dict = field(default_factory=dict)


# ---------------------------------------------------------------- tiện ích
def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def _skip_spans(text: str):
    return [(m.start(), m.end()) for m in SKIP_SPAN_RE.finditer(text)]


def _in(spans, pos: int) -> bool:
    return any(a <= pos < b for a, b in spans)


def sentence_start(text: str, pos: int) -> bool:
    prefix = text[:pos].rstrip(" \t" + NBSP)
    if not prefix:
        return True
    if re.fullmatch(r"(?:[-–•+*]|\(?\d{1,2}[.)]|[a-zđ][.)])", prefix):
        return True
    return prefix[-1] in ".!?:"


def _match_case(src: str, repl: str) -> str:
    if src.isupper() and len(src) > 3:
        return repl.upper()
    return repl[:1].upper() + repl[1:] if src[:1].isupper() else repl


# ---------------------------------------------------------------- kiểm tra từng đoạn
def check_paragraph(p: Para, prev: Para | None, opts: Options, obs: DocObservations) -> list[Issue]:
    text = nfc(p.text)
    out: list[Issue] = []
    if not text.strip():
        return out

    def add(cat, sev, found, msg, fix=""):
        out.append(Issue(cat, sev, p.where, found, msg, fix, p.index))

    spans = _skip_spans(text)
    stripped = text.strip(" \t" + NBSP)

    # ---------- dấu cách
    if re.match(r"^[ \t" + NBSP + r"]+\S", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, 0, 6), "Có dấu cách/tab ở đầu dòng (không dùng dấu cách để thụt đầu dòng hay căn chỉnh)",
            "Xóa, dùng thụt đầu dòng (First line indent) hoặc căn lề")
    if re.search(r"\S[ \t" + NBSP + r"]+$", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, len(text) - 6, len(text)), "Thừa dấu cách ở cuối dòng", "Xóa dấu cách cuối dòng")
    for m in re.finditer(r"(?<=\S)[ " + NBSP + r"]{2,}(?=\S)", text):
        if _in(spans, m.start()):
            continue
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Hai dấu cách liền nhau trở lên", "Chỉ để 1 dấu cách")
    for m in re.finditer(r"(?<=[^\s.…])[ " + NBSP + r"]+([,;:!?])", text):
        if _in(spans, m.start()):
            continue
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), f"Có dấu cách trước dấu '{m.group(1)}'",
            "Dấu câu dính sát chữ đứng trước")
    for m in re.finditer(r"(?<=[^\s.…\d])[ " + NBSP + r"]+(\.)(?![.\d])", text):
        if _in(spans, m.start()):
            continue
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Có dấu cách trước dấu chấm", "Dấu chấm dính sát chữ đứng trước")
    for m in re.finditer(r"(?<=[^\W\d_)\]\"”’'%]),(?=[^\s\d\"”’')\]}»,.;:!?…])|(?<=\d),(?=[^\W\d_])|(?<=[^\W\d_)\]\"”’'%]);(?=[^\s\"”’')\]}»,.;:!?…])", text):
        if _in(spans, m.start()):
            continue
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end() + 1), f"Thiếu dấu cách sau dấu '{m.group(0)}'", "Thêm 1 dấu cách sau dấu câu")
    for m in re.finditer(r"(?<=[^\W\d_]):(?=[^\W\d_])", text):
        if _in(spans, m.start()):
            continue
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end() + 1), "Thiếu dấu cách sau dấu hai chấm", "Thêm 1 dấu cách sau ':'")
    for m in re.finditer(r"(?<=[" + LOWER + r"]{2})\.(?=[" + UPPER + r"][" + LOWER + r"])", text):
        if _in(spans, m.start()):
            continue
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end() + 1), "Thiếu dấu cách sau dấu chấm", "Thêm 1 dấu cách sau dấu chấm")
    for m in re.finditer(r"\(\s+\S|\S\s+\)", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Có dấu cách bên trong ngoặc đơn", "Ngoặc dính sát chữ: (nội dung)")
    for m in re.finditer(r"(?<=[^\W\d_])\((?=[^\W\d_])(?!s\))", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Thiếu dấu cách trước ngoặc đơn", "Thêm dấu cách trước '('")
    for m in re.finditer(r"\)(?=[^\W_])", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end() + 1), "Thiếu dấu cách sau ngoặc đơn", "Thêm dấu cách sau ')'")
    for m in re.finditer(r"“[ " + NBSP + r"]+\S|\S[ " + NBSP + r"]+”", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Có dấu cách bên trong dấu ngoặc kép", "Ngoặc kép dính sát chữ")
    for m in re.finditer(r'"([^"]*)"', text):
        g = m.group(1)
        if g != g.strip():
            add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Có dấu cách bên trong dấu ngoặc kép", "Ngoặc kép dính sát chữ")
    for m in re.finditer(r"(?<=[0-9A-ZĐ])[ " + NBSP + r"]+/|/[ " + NBSP + r"]+(?=[0-9A-ZĐ])", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Dấu gạch chéo trong số/ký hiệu không được cách (30/2020/NĐ-CP, V/v)",
            "Bỏ dấu cách quanh '/'")
    for m in re.finditer(r"(\d+/(?:\w+/)*[A-ZĐ]+)[ " + NBSP + r"]+-[ " + NBSP + r"]*([A-ZĐ])|(\d+/(?:\w+/)*[A-ZĐ]+)[ " + NBSP + r"]*-[ " + NBSP + r"]+([A-ZĐ])", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Gạch nối trong ký hiệu văn bản không được cách (NĐ-CP)", "Bỏ dấu cách quanh '-'")
    for m in re.finditer(r"(?<![.\d])\.\.(?!\.)|,,|;;|\.,|,\.", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Dấu câu bị gõ lặp hoặc liền kề nhau", "Giữ lại 1 dấu thích hợp")
    for m in re.finditer(r"(?<=\d)[ " + NBSP + r"]+%", text):
        add(CAT_PUNCT, SEV_HINT, short(text, m.start(), m.end()), "Có dấu cách giữa số và dấu %", "Viết liền: 0,66%")

    # ---------- số tiền, số thập phân
    for m in re.finditer(r"(?<![\d.,])\d{1,3}(?:,\d{3})+(?![\d,])(?=\s*(?:đồng|VNĐ|VND|nghìn|triệu|tỷ|đ\b))", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Số tiền dùng dấu phẩy phân cách hàng nghìn",
            m.group(0).replace(",", "."))
    for m in re.finditer(r"(?<![\d.,])\d{4,}(?![\d.,])(?=[ " + NBSP + r"]*(?:đồng|VNĐ|VND)\b)", text):
        n = m.group(0)
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Số tiền chưa có dấu chấm phân cách hàng nghìn",
            f"{int(n):,}".replace(",", "."))
    for m in re.finditer(r"(?<![\d.,])\d{1,2}\.\d{1,2}[ " + NBSP + r"]*%", text):
        add(CAT_PUNCT, SEV_ERROR, short(text, m.start(), m.end()), "Phần thập phân dùng dấu phẩy, không dùng dấu chấm",
            m.group(0).replace(".", ","))

    # ---------- chính tả: cụm từ sai thường gặp
    covered = []
    for rx, a, b in MISSPELL_RE + compile_pairs(opts.extra_misspell):
        for m in rx.finditer(text):
            covered.append((m.start(), m.end()))
            add(CAT_SPELL, SEV_ERROR, short(text, m.start(), m.end()), f"Sai chính tả: '{m.group(0)}'", _match_case(m.group(0), b))
    for m in re.finditer(r"(?<![\w])(\w+)\s+\1(?![\w])", text, re.I):
        if m.group(1).lower() in FUNC_DUP:
            add(CAT_SPELL, SEV_ERROR, short(text, m.start(), m.end()), f"Lặp từ '{m.group(1)}'", m.group(1))

    # ---------- chính tả: từng âm tiết
    if opts.check_spelling:
        for m in WORD_RE.finditer(text):
            w = m.group(0)
            lw = w.lower()
            if len(w) < 2 or _in(spans, m.start()) or _in(covered, m.start()):
                continue
            if lw in SLANG:
                if lw == "k" and len(w) == 1:
                    continue
                add(CAT_SPELL, SEV_ERROR, short(text, m.start(), m.end()), f"Viết tắt kiểu nhắn tin '{w}' không dùng trong văn bản hành chính",
                    SLANG[lw])
                continue
            if lw in WHITELIST or lw in opts.whitelist or ROMAN_RE.match(lw):
                continue
            if w.isupper():
                if not check_syllable(w).ok:
                    continue            # chữ viết tắt IN HOA
            elif any(c.isupper() for c in w[1:]):
                continue                # lẫn hoa thường: để luật "chữ hoa giữa từ" xử lý
            r = check_syllable(w)
            if not r.ok:
                titled = w[:1].isupper() and w[1:].islower()
                if titled and not r.suggestion:
                    add(CAT_SPELL, SEV_WARN, short(text, m.start(), m.end()),
                        f"'{w}' không phải âm tiết tiếng Việt hợp lệ ({r.problem}); nếu là tên riêng nước ngoài thì bỏ qua",
                        "Thêm vào tu_dien_rieng.txt nếu đúng")
                else:
                    add(CAT_SPELL, SEV_ERROR, short(text, m.start(), m.end()), f"Sai chính tả '{w}': {r.problem}", r.suggestion)
            elif r.style in ("moi", "cu") and opts.kieu_dau != "bo-qua":
                obs.tone.append((p, text, m.start(), m.end(), w, r.style))
            if lw in dict(IY_PAIRS) or lw in {b for _, b in IY_PAIRS}:
                if w[:1].islower() or sentence_start(text, m.start()):   # bỏ họ "Lý" giữa câu
                    obs.iy[lw] += 1
                    obs.iy_where.setdefault(lw, []).append((p, text, m.start(), m.end()))

    # ---------- chữ hoa giữa từ (thiếu dấu cách / gõ nhầm Caps)
    for m in re.finditer(r"[a-zàáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ][A-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ]+[a-zàáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]*", text):
        if _in(spans, m.start()):
            continue
        add(CAT_CAP, SEV_WARN, short(text, m.start(), m.end()), "Chữ hoa xuất hiện giữa từ (thiếu dấu cách hoặc bật nhầm Caps Lock)",
            "Kiểm tra lại: tách từ hoặc viết thường")

    # ---------- viết hoa
    if not (stripped.isupper() or p.in_table and len(stripped) < 40):
        words = [w for w in stripped.split() if w[:1].isalpha()]
        heading_like = len(stripped) < 90 and words and sum(w[0].isupper() for w in words) / len(words) > 0.6
        if not heading_like:
            for m in WORD_RE.finditer(text):
                w = m.group(0)
                if w.lower() in COMMON_NOUNS and w[0].isupper() and w[1:].islower() and not sentence_start(text, m.start()):
                    add(CAT_CAP, SEV_ERROR, short(text, m.start(), m.end()),
                        f"Viết hoa tùy tiện danh từ chung '{w}' (NĐ30 Phụ lục VI)", w.lower())
    for rx, canon, lower_ok in ORG_RE:
        for m in rx.finditer(text):
            g = nfc(m.group(0))
            g = re.sub(r"\s+", " ", g)
            if g == canon or g.isupper() or (lower_ok and g == g.lower()):
                continue
            if g.replace("Uỷ", "Ủy") == canon:
                continue           # sai kiểu dấu thanh: đã có luật riêng
            add(CAT_CAP, SEV_WARN, short(text, m.start(), m.end()), f"Tên cơ quan viết hoa chưa đúng NĐ30: '{g}'", canon)
    unit_alt = "|".join(ADMIN_UNITS)
    for m in re.finditer(r"(?<![\w])(" + unit_alt + r")[ " + NBSP + r"]+(?=[" + UPPER + r"])", text):
        if sentence_start(text, m.start()):
            continue
        add(CAT_CAP, SEV_ERROR, short(text, m.start(), m.end() + 8),
            f"Danh từ chung chỉ đơn vị hành chính đứng trước tên riêng viết thường: '{m.group(1).lower()} …'", m.group(1).lower())
    for m in re.finditer(r"(?<![\w])(điều|chương)[ " + NBSP + r"]+(\d+|[IVXLC]+)(?![\w])", text):
        add(CAT_CAP, SEV_ERROR, short(text, m.start(), m.end()), f"'{m.group(1).capitalize()}' kèm số phải viết hoa chữ đầu",
            f"{m.group(1).capitalize()} {m.group(2)}")
    for m in re.finditer(r"(?<![\w])(Khoản|Điểm)[ " + NBSP + r"]+(\d+|[a-zđ])(?![\w])", text):
        if sentence_start(text, m.start()):
            continue
        add(CAT_CAP, SEV_ERROR, short(text, m.start(), m.end()), f"'{m.group(1)}' giữa câu viết thường",
            f"{m.group(1).lower()} {m.group(2)}")
    dt = "|".join(t.lower() for t in DOC_TYPES)
    for m in re.finditer(r"(?<![\w])(" + dt + r")[ " + NBSP + r"]+số[ " + NBSP + r"]+\d", text):
        add(CAT_CAP, SEV_ERROR, short(text, m.start(), m.end()),
            f"Tên loại văn bản khi viện dẫn cụ thể phải viết hoa chữ đầu: '{m.group(1).capitalize()} số …'",
            m.group(1).capitalize() + " số")

    # ---------- viết hoa đầu câu
    for m in re.finditer(r"([" + LOWER + r"]{2,})([.!?])[ " + NBSP + r"]+([" + LOWER + r"])", text):
        pre = text[max(0, m.start() - 3): m.start(1)]
        if _in(spans, m.start()) or m.group(1) in ("tp", "tt", "ts", "ths", "gs", "pgs", "bs", "vv", "stt", "đc") or pre.endswith("."):
            continue
        add(CAT_CAP, SEV_ERROR, short(text, m.start(), m.end()), "Đầu câu (sau dấu chấm/chấm hỏi/chấm than) phải viết hoa",
            (re.match(r"\S+", text[m.start(3):]).group(0)).capitalize())
    if not p.in_table and prev is not None:
        body = re.sub(r"^(?:[-–•+*]\s*|\(?\d{1,2}[.)]\s*|[a-zđ][.)]\s+)", "", stripped)
        pt = prev.text.strip()
        if len(body) >= 15 and body[:1].islower() and pt and pt[-1] in ".!?:" and not text.startswith(("v/v", "(")):
            add(CAT_CAP, SEV_WARN, short(text, 0, 12), "Đoạn mới (sau dấu chấm/hai chấm xuống dòng) bắt đầu bằng chữ thường",
                body[:1].upper() + body[1:12])
    if re.match(r"^v/v\b", stripped, re.I) and not stripped.startswith("V/v"):
        add(CAT_FORM, SEV_ERROR, short(text, 0, 8), "Trích yếu công văn phải viết 'V/v' (chữ V hoa, v thường)", "V/v" + stripped[3:])
    if not p.in_table and len(re.sub(r"[^A-ZÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐa-zàáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]", "", stripped)) >= 60 and stripped.isupper():
        add(CAT_CAP, SEV_WARN, short(text, 0, 40), "Cả đoạn dài viết IN HOA (không viết hoa cả câu trong thân văn bản)", "Viết thường, chỉ hoa chữ đầu câu")

    # ---------- ngày tháng, số ký hiệu
    out.extend(check_dates(p, text))
    out.extend(check_so_hieu(p, text))
    return out


def check_dates(p: Para, text: str) -> list[Issue]:
    out = []
    for m in re.finditer(r"(?<![\w])ngày[ " + NBSP + r"]+(\d{1,2})[ " + NBSP + r"]+tháng[ " + NBSP + r"]+(\d{1,2})[ " + NBSP + r"]+năm[ " + NBSP + r"]+(\d{4})(?!\d)", text):
        d, mo, y = m.group(1), m.group(2), m.group(3)
        di, moi, yi = int(d), int(mo), int(y)
        found = short(text, m.start(), m.end())
        if not 1 <= moi <= 12 or not 1 <= di <= calendar.monthrange(yi, moi)[1] if 1 <= moi <= 12 else True:
            out.append(Issue(CAT_FORM, SEV_ERROR, p.where, found, "Ngày tháng không tồn tại", "", p.index))
            continue
        want_d = f"{di:02d}" if di < 10 else str(di)
        want_m = f"{moi:02d}" if moi < 3 else str(moi)
        if d != want_d or mo != want_m:
            fixed = f"ngày {want_d} tháng {want_m} năm {y}"
            out.append(Issue(CAT_FORM, SEV_ERROR, p.where, found,
                             "Ngày nhỏ hơn 10 và tháng 1, 2 phải thêm số 0 (tháng 3 đến 12 không thêm 0)", fixed, p.index))
    return out


def check_so_hieu(p: Para, text: str) -> list[Issue]:
    out = []
    s = text.strip()
    m = re.match(r"^[Ss][ốo]([ " + NBSP + r"]*:?)[ " + NBSP + r"]*(\d+)/", s)
    if not m:
        return out
    found = short(s, 0, 14)
    if not s.startswith("Số"):
        out.append(Issue(CAT_FORM, SEV_ERROR, p.where, found, "Chữ 'Số' phải viết hoa chữ S", "Số", p.index))
    if ":" not in m.group(1):
        out.append(Issue(CAT_FORM, SEV_ERROR, p.where, found, "Thiếu dấu hai chấm sau 'Số'", f"Số: {m.group(2)}/…", p.index))
    elif re.match(r"^[Ss][ốo][ " + NBSP + r"]+:", s):
        out.append(Issue(CAT_PUNCT, SEV_ERROR, p.where, found, "Có dấu cách trước dấu hai chấm sau 'Số'", f"Số: {m.group(2)}/…", p.index))
    num = m.group(2)
    if len(num) == 1 and num != "0":
        out.append(Issue(CAT_FORM, SEV_ERROR, p.where, found, "Số văn bản nhỏ hơn 10 phải thêm số 0", f"Số: 0{num}/…", p.index))
    sym = s[m.end():].split()[0] if s[m.end():].split() else ""
    if sym and re.search(r"[a-zàáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ]", sym) and not re.search(r"[A-ZĐ]", sym):
        out.append(Issue(CAT_FORM, SEV_WARN, p.where, found, "Ký hiệu văn bản (sau số) phải viết IN HOA (có thể trừ chữ viết tắt loại văn bản như TTr, CV)", "", p.index))
    return out


# ---------------------------------------------------------------- tổng hợp toàn văn bản
def finish_document(obs: DocObservations, opts: Options) -> list[Issue]:
    out: list[Issue] = []
    if obs.tone and opts.kieu_dau != "bo-qua":
        want = "moi" if opts.kieu_dau == "moi" else "cu"
        n_new = sum(1 for t in obs.tone if t[5] == "moi")
        n_old = sum(1 for t in obs.tone if t[5] == "cu")
        mixed = n_new > 0 and n_old > 0
        for p, text, a, b, w, style in obs.tone:
            if style == want:
                continue
            fixed = to_new_style(w) if want == "moi" else w
            if want == "cu":
                from .vietnamese import to_old_style
                fixed = to_old_style(w)
            out.append(Issue(CAT_SPELL, SEV_WARN if mixed else SEV_HINT, p.where, short(text, a, b),
                             "Văn bản dùng lẫn hai kiểu bỏ dấu (hòa/hoà, thủy/thuỷ)" if mixed else
                             "Kiểu bỏ dấu cũ; Quốc hiệu và NĐ30 dùng kiểu mới (hòa, thủy, khỏe, ủy)", fixed, p.index))
    for a, b in IY_PAIRS:
        ca, cb = obs.iy.get(a, 0), obs.iy.get(b, 0)
        if ca and cb:
            keep, drop = (a, b) if ca > cb else (b, a)
            for p, text, s, e in obs.iy_where.get(drop, []):
                out.append(Issue(CAT_SPELL, SEV_WARN, p.where, short(text, s, e),
                                 f"Văn bản dùng lẫn '{a}' và '{b}': cần thống nhất một cách viết", keep, p.index))
    return out


def check_text_document(paras: list[Para], opts: Options | None = None) -> list[Issue]:
    opts = opts or Options()
    obs = DocObservations()
    issues: list[Issue] = []
    prev = None
    for p in paras:
        issues.extend(check_paragraph(p, prev, opts, obs))
        if p.text.strip() and not p.in_table:
            prev = p
    issues.extend(finish_document(obs, opts))
    return issues

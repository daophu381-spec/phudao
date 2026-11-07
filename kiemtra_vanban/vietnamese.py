"""Kiểm tra một âm tiết có đúng quy tắc ghép vần tiếng Việt hay không.

Không cần từ điển: dựa vào cấu trúc  [phụ âm đầu] + [vần: nguyên âm] + [âm cuối]
và các quy tắc chính tả (c/k/q, g/gh, ng/ngh, âm cuối c-ch-p-t chỉ đi với dấu sắc/nặng,
vị trí đặt dấu thanh...). Bắt được phần lớn lỗi gõ nhầm, sai dấu, thừa/thiếu chữ.
Không bắt được lỗi "âm tiết đúng nhưng từ sai" (vd. "xử dụng") - phần đó do từ điển lỗi thường gặp xử lý.
"""
from __future__ import annotations

import unicodedata
from dataclasses import dataclass

# Dấu thanh (combining): huyền, sắc, hỏi, ngã, nặng
TONE_MARKS = {"̀": "huyền", "́": "sắc", "̉": "hỏi", "̃": "ngã", "̣": "nặng"}
VOWELS = set("aăâeêioôơuưy")
MODIFIED = set("ăâêôơư")
ONSETS = ["ngh", "ng", "nh", "ch", "gh", "gi", "kh", "ph", "qu", "th", "tr",
          "b", "c", "d", "đ", "g", "h", "k", "l", "m", "n", "p", "r", "s", "t", "v", "x"]
CODAS = ["ng", "nh", "ch", "c", "m", "n", "p", "t"]

NUCLEI_BOTH = {"a", "e", "ê", "i", "y", "o", "ô", "ơ", "u", "ư", "oa", "oe", "uê", "uy"}
NUCLEI_CLOSED = {"ă", "â", "iê", "yê", "uô", "ươ", "oă", "uâ", "uyê", "oo"}
NUCLEI_OPEN = {"ia", "ya", "ua", "ưa", "ai", "ay", "ao", "au", "âu", "ây", "eo", "êu", "iu",
               "oi", "ôi", "ơi", "ui", "ưi", "ưu", "oai", "oay", "oeo", "uây", "uôi", "ươi",
               "ươu", "iêu", "yêu", "uyu", "uya", "uơ"}
ALL_NUCLEI = NUCLEI_BOTH | NUCLEI_CLOSED | NUCLEI_OPEN


@dataclass
class SyllableResult:
    ok: bool
    problem: str = ""        # mô tả lỗi
    suggestion: str = ""     # gợi ý sửa (có thể rỗng)
    style: str = ""          # "moi" (hòa, thủy, khỏe) | "cu" (hoà, thuỷ, khoẻ) | "" (không áp dụng)

    @property
    def style_old(self) -> bool:
        return self.style == "cu"


def _split_tone(word: str):
    """Tách dấu thanh. Trả (chuỗi_không_dấu_thanh, chỉ_số_chữ_mang_dấu, tên_dấu)."""
    base, idx, tone = [], -1, ""
    for i, ch in enumerate(unicodedata.normalize("NFC", word)):
        d = unicodedata.normalize("NFD", ch)
        keep = []
        for c in d:
            if c in TONE_MARKS:
                idx, tone = i, TONE_MARKS[c]
            else:
                keep.append(c)
        base.append(unicodedata.normalize("NFC", "".join(keep)))
    return "".join(base), idx, tone


def _put_tone(base: str, idx: int, tone_name: str) -> str:
    mark = next(k for k, v in TONE_MARKS.items() if v == tone_name)
    ch = unicodedata.normalize("NFC", unicodedata.normalize("NFD", base[idx]) + mark)
    return base[:idx] + ch + base[idx + 1:]


def _has_vowel(s: str) -> bool:
    return any(c in VOWELS for c in s)


def parse_syllable(low: str):
    """Tách (onset, nucleus, coda) từ âm tiết đã bỏ dấu thanh, viết thường. None nếu không tách được."""
    if not low:
        return None
    for on in sorted(ONSETS, key=len, reverse=True) + [""]:
        if on and not low.startswith(on):
            continue
        rest = low[len(on):]
        if on == "gi" and not _has_vowel(rest):
            continue          # "gì", "gin": gi- không phải phụ âm đầu
        if on == "qu" and not rest:
            continue
        if not rest:
            continue
        for coda in CODAS + [""]:
            if coda and not rest.endswith(coda):
                continue
            nuc = rest[: len(rest) - len(coda)]
            if nuc and nuc in ALL_NUCLEI:
                return on, nuc, coda
    return None


def check_syllable(word: str, old_style_ok: bool = True) -> SyllableResult:
    """Kiểm tra 1 âm tiết (một dãy chữ cái, không có dấu cách)."""
    base, tone_idx, tone = _split_tone(word)
    low = base.lower()
    if any(c in "fjwz" for c in low):
        return SyllableResult(True)  # từ nước ngoài / ký hiệu, bỏ qua
    parsed = parse_syllable(low)
    if parsed is None:
        return SyllableResult(False, "không đúng cấu tạo âm tiết tiếng Việt")
    onset, nuc, coda = parsed

    first = nuc[0]
    if onset == "c" and first in "iyeê":
        return SyllableResult(False, "'c' không đứng trước i, e, ê, y", _swap_onset(word, "c", "k"))
    if onset == "k" and first not in "iyeê":
        return SyllableResult(False, "'k' chỉ đứng trước i, e, ê, y", _swap_onset(word, "k", "c"))
    if onset == "g" and first in "ieê" and not low.startswith("gi"):
        return SyllableResult(False, "'g' không đứng trước i, e, ê (dùng 'gh')", _swap_onset(word, "g", "gh"))
    if onset == "gh" and first not in "ieê":
        return SyllableResult(False, "'gh' chỉ đứng trước i, e, ê", _swap_onset(word, "gh", "g"))
    if onset == "ng" and first in "ieê":
        return SyllableResult(False, "'ng' không đứng trước i, e, ê (dùng 'ngh')", _swap_onset(word, "ng", "ngh"))
    if onset == "ngh" and first not in "ieê":
        return SyllableResult(False, "'ngh' chỉ đứng trước i, e, ê", _swap_onset(word, "ngh", "ng"))
    if onset == "qu" and first in "uo":
        return SyllableResult(False, "sau 'qu' không thể là u/o")

    if coda in ("ch", "nh") and nuc[-1] not in "aêiy":
        return SyllableResult(False, f"vần '{nuc}' không đi với '{coda}'")
    if coda in ("c", "ng") and nuc in ("i", "ê", "y"):
        return SyllableResult(False, f"vần '{nuc}' không đi với '{coda}'")
    if coda and nuc in NUCLEI_OPEN:
        return SyllableResult(False, f"vần '{nuc}' không có âm cuối '{coda}'")
    if not coda and nuc in NUCLEI_CLOSED:
        return SyllableResult(False, f"vần '{nuc}' phải có âm cuối")
    if coda in ("c", "ch", "p", "t") and tone not in ("sắc", "nặng"):
        return SyllableResult(False, f"âm cuối '{coda}' chỉ đi với dấu sắc hoặc dấu nặng"
                              + (f" (đang là dấu {tone})" if tone else " (đang không dấu)"))

    # ----- vị trí dấu thanh -----
    if tone_idx >= 0:
        s = len(onset)               # chỉ số bắt đầu của vần trong chuỗi
        # nếu onset là 'gi' mà vần bắt đầu bằng 'i' (gì, gìn) đã được xử lý ở parse
        vowel_positions = list(range(s, s + len(nuc)))
        exp, style_dep = _expected_mark(nuc, coda, vowel_positions)
        if style_dep:
            if tone_idx not in vowel_positions[:2]:
                return SyllableResult(False, "dấu thanh đặt sai chữ", _put_tone(base, vowel_positions[0], tone))
            return SyllableResult(True, style="cu" if tone_idx == vowel_positions[1] else "moi")
        if tone_idx != exp:
            return SyllableResult(False, "dấu thanh đặt sai chữ", _put_tone(base, exp, tone))
    return SyllableResult(True)


def _expected_mark(nuc: str, coda: str, pos: list[int]):
    """Trả (chỉ số chữ cần mang dấu, có phụ thuộc kiểu cũ/mới hay không)."""
    if len(nuc) == 1:
        return pos[0], False
    mods = [i for i, c in enumerate(nuc) if c in MODIFIED]
    if mods:
        return pos[mods[-1]], False
    if len(nuc) == 2:
        if nuc in ("oa", "oe", "uy") and not coda:
            return pos[0], True
        if nuc in ("oa", "oe", "uy", "oo", "oă"):
            return pos[1], False
        return pos[0], False
    return pos[1], False   # 3 nguyên âm: dấu ở giữa (oai, oay, oeo, uyu)


def _swap_onset(word: str, old: str, new: str) -> str:
    base, idx, tone = _split_tone(word)
    if not base.lower().startswith(old):
        return ""
    rep = new.upper() if word[: len(old)].isupper() and len(word) > 1 and word[len(old):len(old) + 1].isupper() \
        else (new.capitalize() if word[0].isupper() else new)
    res = rep + base[len(old):]
    if idx >= 0:
        res = _put_tone(res, idx - len(old) + len(new), tone)
    return res


def to_old_style(word: str) -> str:
    """hòa -> hoà (chỉ dùng cho thử nghiệm)."""
    base, idx, tone = _split_tone(word)
    return _put_tone(base, idx + 1, tone) if idx >= 0 else word


def to_new_style(word: str) -> str:
    """hoà -> hòa."""
    base, idx, tone = _split_tone(word)
    res = check_syllable(word)
    if idx < 0 or not res.style_old:
        return word
    return _put_tone(base, idx - 1, tone)

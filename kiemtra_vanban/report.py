from __future__ import annotations

import html
import json
from collections import Counter

from .models import CATEGORIES, SEV_ERROR, SEV_ORDER, SEV_WARN, SEV_HINT, Issue


def dedupe(issues: list[Issue]) -> list[Issue]:
    seen, out = set(), []
    for i in issues:
        k = (i.category, i.where, i.found, i.message, i.fix)
        if k not in seen:
            seen.add(k)
            out.append(i)
    return out


def organize(issues: list[Issue], only_errors: bool = False) -> dict[str, list[Issue]]:
    issues = dedupe(issues)
    if only_errors:
        issues = [i for i in issues if i.severity == SEV_ERROR]
    groups: dict[str, list[Issue]] = {c: [] for c in CATEGORIES}
    for i in issues:
        groups.setdefault(i.category, []).append(i)
    for g in groups.values():
        g.sort(key=lambda i: (SEV_ORDER[i.severity], i.order))
    return {k: v for k, v in groups.items() if v}


def summary(issues: list[Issue]) -> str:
    c = Counter(i.severity for i in dedupe(issues))
    return f"{c[SEV_ERROR]} lỗi, {c[SEV_WARN]} cảnh báo, {c[SEV_HINT]} gợi ý"


def to_text(name: str, issues: list[Issue], notes: list[str], only_errors=False, color=False) -> str:
    def paint(s, code):
        return f"\033[{code}m{s}\033[0m" if color else s
    sev_color = {SEV_ERROR: "31", SEV_WARN: "33", SEV_HINT: "36"}
    lines = [paint(f"KẾT QUẢ KIỂM TRA: {name}", "1"), paint("Tổng: " + summary(issues), "1")]
    lines += [f"  ! {n}" for n in notes]
    n = 0
    for cat, items in organize(issues, only_errors).items():
        lines += ["", paint(f"== {cat} ({len(items)}) ==", "1")]
        for i in items:
            n += 1
            lines.append(f"{n:>3}. [{paint(i.severity, sev_color[i.severity])}] {i.where}")
            if i.found:
                lines.append(f"       Chỗ lỗi : {i.found}")
            lines.append(f"       Lỗi     : {i.message}")
            if i.fix:
                lines.append(f"       Sửa thành: {i.fix}")
    if n == 0:
        lines += ["", "Không phát hiện lỗi nào."]
    return "\n".join(lines)


def to_markdown(name: str, issues: list[Issue], notes: list[str], only_errors=False) -> str:
    esc = lambda s: s.replace("|", "\\|").replace("\n", " ")
    lines = [f"# Kết quả kiểm tra: {name}", "", f"**Tổng:** {summary(issues)}", ""]
    lines += [f"> {n}" for n in notes]
    n = 0
    for cat, items in organize(issues, only_errors).items():
        lines += ["", f"## {cat} ({len(items)})", "", "| # | Mức | Vị trí | Chỗ lỗi | Lỗi | Sửa thành |", "|---|---|---|---|---|---|"]
        for i in items:
            n += 1
            lines.append(f"| {n} | {i.severity} | {esc(i.where)} | {esc(i.found)} | {esc(i.message)} | {esc(i.fix)} |")
    return "\n".join(lines) + "\n"


def to_html(name: str, issues: list[Issue], notes: list[str], only_errors=False) -> str:
    e = html.escape
    cls = {SEV_ERROR: "err", SEV_WARN: "warn", SEV_HINT: "hint"}
    parts = [f"<!doctype html><html lang='vi'><meta charset='utf-8'><title>Kiểm tra {e(name)}</title>",
             "<style>body{font:15px/1.5 system-ui,sans-serif;max-width:1100px;margin:24px auto;padding:0 16px;color:#222}"
             "table{border-collapse:collapse;width:100%;margin:8px 0 24px}td,th{border:1px solid #ddd;padding:6px 8px;vertical-align:top}"
             "th{background:#f4f4f4;text-align:left}.err{color:#b00020;font-weight:600}.warn{color:#a05a00;font-weight:600}"
             ".hint{color:#00698f;font-weight:600}.note{background:#fff7e0;padding:8px 12px;border-left:4px solid #e0a800;margin:6px 0}"
             "code{background:#f2f2f2;padding:1px 4px;border-radius:3px}</style>",
             f"<h1>Kết quả kiểm tra: {e(name)}</h1><p><b>Tổng:</b> {e(summary(issues))}</p>"]
    parts += [f"<div class='note'>{e(n)}</div>" for n in notes]
    n = 0
    for cat, items in organize(issues, only_errors).items():
        parts.append(f"<h2>{e(cat)} ({len(items)})</h2><table><tr><th>#</th><th>Mức</th><th>Vị trí</th><th>Chỗ lỗi</th><th>Lỗi</th><th>Sửa thành</th></tr>")
        for i in items:
            n += 1
            parts.append(f"<tr><td>{n}</td><td class='{cls[i.severity]}'>{e(i.severity)}</td><td>{e(i.where)}</td>"
                         f"<td><code>{e(i.found)}</code></td><td>{e(i.message)}</td><td>{e(i.fix)}</td></tr>")
        parts.append("</table>")
    if n == 0:
        parts.append("<p>Không phát hiện lỗi nào.</p>")
    return "".join(parts)


def to_json(name: str, issues: list[Issue], notes: list[str]) -> str:
    return json.dumps({"file": name, "tong": summary(issues), "ghi_chu": notes,
                       "loi": [i.to_dict() for i in dedupe(issues)]}, ensure_ascii=False, indent=2)

"""Chạy:  python -m kiemtra_vanban  <file.docx | file.pdf | file.txt | thư mục>  [tùy chọn]"""
from __future__ import annotations

import argparse
import os
import sys

from . import report
from .models import Para, SEV_ERROR
from .textrules import Options, check_text_document


def load_whitelist(extra: str | None) -> set:
    words = set()
    cands = [os.path.join(os.getcwd(), "tu_dien_rieng.txt"), os.path.join(os.path.dirname(__file__), "..", "tu_dien_rieng.txt")]
    if extra:
        cands.append(extra)
    for c in cands:
        if os.path.isfile(c):
            with open(c, encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip().lower()
                    if line and not line.startswith("#"):
                        words.update(line.split())
    return words


def load_extra_misspell(extra: str | None) -> list:
    """Đọc file 'sai => đúng' (mỗi dòng một cặp, dòng bắt đầu bằng # là ghi chú)."""
    pairs = []
    cands = [os.path.join(os.getcwd(), "loi_chinh_ta_rieng.txt"), os.path.join(os.path.dirname(__file__), "..", "loi_chinh_ta_rieng.txt")]
    if extra:
        cands.append(extra)
    for c in cands:
        if os.path.isfile(c):
            with open(c, encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip()
                    if line and not line.startswith("#") and "=>" in line:
                        a, b = (x.strip() for x in line.split("=>", 1))
                        if a and b:
                            pairs.append((a.lower(), b))
    return pairs


def check_file(path: str, args) -> tuple[list, list[str]]:
    ext = os.path.splitext(path)[1].lower()
    opts = Options(kieu_dau=args.kieu_dau, whitelist=load_whitelist(args.tu_dien), check_spelling=not args.bo_chinh_ta,
                   extra_misspell=load_extra_misspell(args.loi_rieng))
    notes: list[str] = []
    issues = []
    if ext in (".docx", ".docm", ".dotx"):
        from .docxrules import DocxFile, check_docx
        f = DocxFile(path)
        issues += check_text_document(f.paras, opts)
        issues += check_docx(f)
        if not args.bo_bo_cuc:
            from . import layout
            try:
                li, _, pages = layout.analyze_docx(path, args.chu_mo_coi, args.luu_pdf)
                issues += li
                notes.append(f"Đã dàn trang thử: {pages} trang (kết quả rơi chữ/rơi dòng phụ thuộc phông và máy dàn trang; "
                             f"máy không có Times New Roman có thể lệch nhẹ so với Word)")
            except layout.LayoutUnavailable as e:
                notes.append(f"Bỏ qua kiểm tra rơi chữ/rơi dòng: {e}")
    elif ext == ".pdf":
        from . import layout
        try:
            li, segs, pages = layout.analyze_pdf(path, args.chu_mo_coi)
        except layout.LayoutUnavailable as e:
            sys.exit(str(e))
        paras = [Para(i, s.text, f"Trang {s.page}, dòng {s.line_no}") for i, s in enumerate(segs)]
        issues += check_text_document(paras, opts)
        issues += li
        notes.append("File PDF: chỉ kiểm tra chính tả, dấu câu và bố cục dòng; thể thức Word (phông, lề, in đậm...) cần file .docx")
    else:
        with open(path, encoding="utf-8", errors="replace") as fh:
            paras = [Para(i, ln.rstrip("\n"), f"Dòng {i + 1}") for i, ln in enumerate(fh)]
        issues += check_text_document(paras, opts)
        notes.append("File văn bản thuần: chỉ kiểm tra chính tả và dấu câu")
    return issues, notes


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="kiemtra_vanban",
                                 description="Kiểm tra lỗi chính tả, viết hoa, dấu câu và thể thức theo Nghị định 30/2020/NĐ-CP")
    ap.add_argument("duong_dan", help="File .docx/.pdf/.txt hoặc thư mục chứa các file .docx")
    ap.add_argument("--md", metavar="FILE", help="ghi báo cáo Markdown")
    ap.add_argument("--html", metavar="FILE", help="ghi báo cáo HTML (mở bằng trình duyệt)")
    ap.add_argument("--json", metavar="FILE", help="ghi báo cáo JSON")
    ap.add_argument("--chi-loi", action="store_true", help="chỉ hiện mức LỖI, ẩn cảnh báo và gợi ý")
    ap.add_argument("--bo-bo-cuc", action="store_true", help="không dàn trang thử (bỏ kiểm tra rơi chữ/dòng/trang)")
    ap.add_argument("--bo-chinh-ta", action="store_true", help="không kiểm tra từng âm tiết tiếng Việt")
    ap.add_argument("--chu-mo-coi", type=int, default=1, metavar="N", help="dòng cuối đoạn có ≤ N chữ (rất ngắn) bị coi là mồ côi (mặc định 1)")
    ap.add_argument("--kieu-dau", choices=["moi", "cu", "bo-qua"], default="moi",
                    help="kiểu bỏ dấu thanh: moi = hòa/thủy (như Quốc hiệu, mặc định), cu = hoà/thuỷ")
    ap.add_argument("--tu-dien", metavar="FILE", help="file chứa từ/tên riêng được bỏ qua khi kiểm tra chính tả (mỗi dòng một từ)")
    ap.add_argument("--loi-rieng", metavar="FILE", help="file 'sai => đúng' (mỗi dòng một cặp) để thêm các lỗi riêng của cơ quan")
    ap.add_argument("--luu-pdf", metavar="FILE", help="lưu bản PDF dàn trang thử để xem lại")
    args = ap.parse_args(argv)

    if os.path.isdir(args.duong_dan):
        files = sorted(os.path.join(args.duong_dan, n) for n in os.listdir(args.duong_dan)
                       if n.lower().endswith(".docx") and not n.startswith("~$"))
    else:
        files = [args.duong_dan]
    if not files:
        print("Không có file nào để kiểm tra.")
        return 2

    color = sys.stdout.isatty()
    worst = 0
    for path in files:
        if not os.path.exists(path):
            print(f"Không tìm thấy file: {path}")
            worst = 2
            continue
        name = os.path.basename(path)
        issues, notes = check_file(path, args)
        print(report.to_text(name, issues, notes, args.chi_loi, color))
        print()
        base = os.path.splitext(path)[0] if len(files) > 1 else None
        for opt, ext, fn in ((args.md, ".md", report.to_markdown), (args.html, ".html", report.to_html)):
            target = (base + "_kiemtra" + ext) if (base and opt) else opt
            if target:
                with open(target, "w", encoding="utf-8") as fh:
                    fh.write(fn(name, issues, notes, args.chi_loi))
        if args.json:
            target = (base + "_kiemtra.json") if base else args.json
            with open(target, "w", encoding="utf-8") as fh:
                fh.write(report.to_json(name, issues, notes))
        if any(i.severity == SEV_ERROR for i in issues):
            worst = max(worst, 1)
    return worst


if __name__ == "__main__":
    sys.exit(main())

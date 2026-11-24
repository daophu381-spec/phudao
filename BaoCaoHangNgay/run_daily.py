#!/usr/bin/env python3
"""Chay hang ngay: tai file -> Gemini (nhan vien) -> Claude (giam doc) -> bao cao .docx
Dung:  python run_daily.py            (chay that)
       python run_daily.py --dry-run  (khong goi AI, chi thu tai file + tao docx mau)"""
import json, re, subprocess, sys, urllib.request, datetime, shutil
from pathlib import Path
from docx import Document

BASE = Path(__file__).parent
cfg = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
dry = "--dry-run" in sys.argv
today = datetime.date.today().isoformat()
dl_dir, out_dir, log_dir = BASE / "tai_ve" / today, BASE / "bao_cao", BASE / "log"
for d in (dl_dir, out_dir, log_dir): d.mkdir(parents=True, exist_ok=True)

def run(cmd, stdin=None):
    r = subprocess.run(cmd, input=stdin, capture_output=True, text=True, encoding="utf-8", timeout=1800)
    if r.returncode != 0: raise RuntimeError(f"{cmd[0]} loi: {r.stderr[:500]}")
    return r.stdout.strip()

# 1) Tai file
f = dl_dir / cfg["ten_file"]
urllib.request.urlretrieve(cfg["url"], f)
print("Da tai:", f)

# 2) Nhan vien Gemini
if dry:
    staff = "- (dry-run) so lieu mau A\n- (dry-run) so lieu mau B"
    md = "# Bao cao " + today + "\n\n## Tom tat\n- Chay thu thanh cong\n\n## Nhan vien bao cao\n" + staff
else:
    staff = run([shutil.which(cfg["gemini_cmd"]) or cfg["gemini_cmd"], "-p", f'{cfg["yeu_cau_nhan_vien"]} @{f}'])
    (dl_dir / "gemini_nhap.md").write_text(staff, encoding="utf-8")
    # 3) Giam doc Claude
    prompt = f'{cfg["yeu_cau_giam_doc"]}\n\nFile goc: {f}\n\n--- BAO CAO CUA NHAN VIEN ---\n{staff}'
    md = run([shutil.which(cfg["claude_cmd"]) or cfg["claude_cmd"], "-p", prompt, "--allowedTools", "Read"])

# 4) Xuat Word
doc = Document()
for line in md.splitlines():
    m = re.match(r"^(#{1,3})\s+(.*)", line)
    if m: doc.add_heading(m.group(2), level=len(m.group(1)))
    elif re.match(r"^\s*[-*]\s+", line): doc.add_paragraph(re.sub(r"^\s*[-*]\s+", "", line), style="List Bullet")
    elif line.strip(): doc.add_paragraph(line.replace("**", ""))
out = out_dir / f"BaoCao_{today}.docx"
doc.save(out)
print("Da tao:", out)

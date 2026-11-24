# Bao cao hang ngay: Claude (giam doc) + Gemini (nhan vien)

## Cai dat 1 lan
1. Cai Python 3 + `pip install python-docx`
2. Cai Gemini CLI (`npm i -g @google/gemini-cli`) va Claude Code (`npm i -g @anthropic-ai/claude-code`); chay `gemini` va `claude` 1 lan de dang nhap bang tai khoan Pro.
3. Sua `config.json`: `url` (link file), `ten_file`, va 2 cau yeu cau.
4. Thu: `python run_daily.py --dry-run` roi `python run_daily.py`.

## Dat gio chay co dinh (vi du 07:30 moi ngay)
**Windows** (CMD):
`schtasks /create /tn "BaoCaoHangNgay" /sc daily /st 07:30 /tr "python C:\duong\dan\BaoCaoHangNgay\run_daily.py"`

**macOS/Linux**: `crontab -e` them dong:
`30 7 * * * cd /duong/dan/BaoCaoHangNgay && python3 run_daily.py >> log/cron.log 2>&1`

May phai bat luc do. Ket qua: `bao_cao/BaoCao_YYYY-MM-DD.docx`.

#!/usr/bin/env python3
"""Tự động hoá công việc lặp lại hằng ngày trên Google Chrome.

Cách hoạt động: điều khiển CHÍNH Chrome của bạn (qua cổng debug 9222) nên giữ
nguyên đăng nhập. Công việc mô tả trong file JSON (xem tasks/vi_du.json).

Lệnh:
  python chrome_auto.py mo-chrome            # mở Chrome ở chế độ cho phép điều khiển
  python chrome_auto.py chay tasks/x.json    # chạy một công việc
  python chrome_auto.py chay tasks/x.json --thu   # chạy thử: chỉ in các bước, không bấm
"""
import argparse, json, os, unicodedata, platform, re, shutil, subprocess, sys, time
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
PORT = 9222
PROFILE = HERE / "chrome_profile"      # profile riêng (đăng nhập 1 lần, dùng mãi)
OUT = HERE / "ket_qua"                 # ảnh chụp, file tải về, log


def tim_chrome():
    cands = {
        "Windows": [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")],
        "Darwin": ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"],
    }.get(platform.system(), [])
    for c in cands:
        if os.path.exists(c):
            return c
    for n in ("google-chrome", "google-chrome-stable", "chromium", "chrome"):
        p = shutil.which(n)
        if p:
            return p
    sys.exit("Không tìm thấy Chrome. Hãy sửa hàm tim_chrome() cho đúng đường dẫn.")


def mo_chrome():
    PROFILE.mkdir(exist_ok=True)
    subprocess.Popen([tim_chrome(), f"--remote-debugging-port={PORT}",
                      f"--user-data-dir={PROFILE}", "--no-first-run"])
    print(f"Đã mở Chrome (cổng {PORT}). Lần đầu hãy đăng nhập các trang cần dùng, "
          "rồi để cửa sổ này mở và chạy lệnh 'chay'.")


def log(msg):
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    print(line, flush=True)
    OUT.mkdir(exist_ok=True)
    with open(OUT / f"log_{datetime.now():%Y-%m-%d}.txt", "a", encoding="utf-8") as f:
        f.write(line + "\n")


def bo_dau(x):
    """'Nguyễn Văn Đức' -> 'Nguyen Van Duc' (cả đ/Đ)."""
    x = x.replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in unicodedata.normalize("NFD", x) if unicodedata.category(c) != "Mn")


def thay_bien(v, bien):
    if isinstance(v, str):
        return re.sub(r"\{(\w+)\}", lambda m: str(bien.get(m.group(1), m.group(0))), v)
    return v


def thuc_hien(page, ctx, buoc, bien):
    b = {k: thay_bien(v, bien) for k, v in buoc.items()}
    a = b["lam"]
    t = b.get("cho_toi_da_giay", 20) * 1000
    sel = b.get("chon")  # CSS selector, hoặc "text=Phê duyệt", "role=button[name='In']"
    if a == "mo_trang":
        page.goto(b["url"], timeout=t)
    elif a == "bam":
        page.locator(sel).first.click(timeout=t)
    elif a == "nhap":
        nd = b["noi_dung"]
        if b.get("bo_dau"):           # ô chỉ nhận không dấu
            nd = bo_dau(nd)
        if b.get("viet_hoa"):
            nd = nd.upper()
        o = page.locator(sel).first
        if b.get("go_tung_phim"):     # gõ như người thật (kích hoạt bộ lọc phím của trang)
            o.click(timeout=t); o.fill("", timeout=t)
            o.press_sequentially(nd, delay=b.get("tre_ms", 60))
        else:
            o.fill(nd, timeout=t)
        # Đọc lại giá trị: nếu trang từ chối/đổi ký tự thì báo lỗi, không đi tiếp
        thuc = o.input_value(timeout=t)
        if thuc != nd and not b.get("khong_kiem_tra"):
            raise ValueError(f"Ô nhập khác dự kiến: mong '{nd}', thực tế '{thuc}'")
        if b.get("khong_co_canh_bao") and page.locator(b["khong_co_canh_bao"]).first.is_visible():
            raise ValueError("Trang đang hiện cảnh báo ở ô vừa nhập")
    elif a == "phim":
        page.keyboard.press(b["phim"])
    elif a == "chon_muc":
        page.locator(sel).first.select_option(b["gia_tri"], timeout=t)
    elif a == "cho":
        time.sleep(float(b.get("giay", 1)))
    elif a == "cho_xuat_hien":
        page.locator(sel).first.wait_for(timeout=t)
    elif a == "chup_anh":
        ten = OUT / thay_bien(b.get("file", "anh_{ngay}_{gio}.png"), bien)
        ten.parent.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(ten), full_page=True)
        log(f"   đã lưu {ten}")
    elif a == "tai_ve":  # bấm vào nút tải và lưu file
        with page.expect_download(timeout=t) as d:
            page.locator(sel).first.click(timeout=t)
        ten = OUT / thay_bien(b.get("file", d.value.suggested_filename), bien)
        ten.parent.mkdir(parents=True, exist_ok=True)
        d.value.save_as(str(ten))
        log(f"   đã tải {ten}")
    elif a == "doc_chu":  # in nội dung phần tử ra log
        log("   > " + page.locator(sel).first.inner_text(timeout=t).strip())
    elif a == "xac_nhan":  # dừng hỏi người dùng trước việc quan trọng (duyệt, gửi...)
        if input(f"   ⚠ {b.get('hoi', 'Tiếp tục?')} [y/N]: ").strip().lower() != "y":
            raise KeyboardInterrupt("Người dùng dừng.")
    else:
        raise ValueError(f"Hành động lạ: {a}")


def chay(file_cv, thu=False):
    cv = json.loads(Path(file_cv).read_text(encoding="utf-8"))
    now = datetime.now()
    bien = {"ngay": now.strftime("%Y-%m-%d"), "gio": now.strftime("%H%M%S"),
            "ngay_vn": now.strftime("%d/%m/%Y"), **cv.get("bien", {})}
    log(f"== {cv.get('ten', file_cv)} ({len(cv['cac_buoc'])} bước)")
    if thu:
        for i, s in enumerate(cv["cac_buoc"], 1):
            log(f"{i}. {s.get('mo_ta', s['lam'])}  {thay_bien(json.dumps(s, ensure_ascii=False), bien)}")
        return
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        try:
            br = p.chromium.connect_over_cdp(f"http://127.0.0.1:{PORT}")
        except Exception:
            sys.exit("Chưa kết nối được Chrome. Chạy trước: python chrome_auto.py mo-chrome")
        ctx = br.contexts[0]
        page = ctx.new_page()
        nghi = cv.get("nghi_giua_cac_buoc_giay", 0.5)
        try:
            for i, s in enumerate(cv["cac_buoc"], 1):
                log(f"{i}. {s.get('mo_ta', s['lam'])}")
                try:
                    thuc_hien(page, ctx, s, bien)
                except KeyboardInterrupt:
                    raise
                except Exception as e:
                    if s.get("bo_qua_loi"):
                        log(f"   bỏ qua lỗi: {str(e).splitlines()[0]}")
                    else:
                        anh = OUT / f"loi_{bien['ngay']}_{bien['gio']}.png"
                        OUT.mkdir(exist_ok=True)
                        page.screenshot(path=str(anh))
                        log(f"   ❌ LỖI bước {i}: {str(e).splitlines()[0]} (ảnh: {anh})")
                        return 1
                time.sleep(nghi)
            log("✅ Xong.")
            return 0
        except KeyboardInterrupt as e:
            log(f"Đã dừng: {e}")
            return 2


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawTextHelpFormatter)
    sp = ap.add_subparsers(dest="lenh", required=True)
    sp.add_parser("mo-chrome")
    c = sp.add_parser("chay"); c.add_argument("file"); c.add_argument("--thu", action="store_true")
    a = ap.parse_args()
    if a.lenh == "mo-chrome":
        mo_chrome()
    else:
        sys.exit(chay(a.file, a.thu))

# -*- coding: utf-8 -*-
"""
NHẬP TỪ EXCEL VÀO PHẦN MỀM NGHIỆP VỤ TRÊN CHROME BẰNG PHÍM (Enter, F5, F2, Lên/Xuống...)

- Điều khiển chính cửa sổ Chrome bạn đang mở và đã đăng nhập (không mở Chrome mới).
- Đọc từng dòng Excel, làm đúng các bước ghi trong file kich_ban.txt.
- Không gửi dữ liệu đi đâu; chỉ gõ phím vào trang đang mở.

Chạy:  python NhapTuExcel.py                -> chạy tất cả các dòng
       python NhapTuExcel.py --thu          -> chỉ chạy dòng đầu tiên để thử
       python NhapTuExcel.py --tu 15        -> chạy từ dòng 15 của Excel (chạy tiếp khi bị dừng)
       python NhapTuExcel.py --tu 15 --den 20
Dừng gấp: bấm Ctrl+C trong cửa sổ đen.
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys
import time

try:
    from openpyxl import load_workbook
    from playwright.sync_api import sync_playwright
except ImportError:
    print("Chưa cài thư viện. Chạy file CaiDat.bat (hoặc: pip install playwright openpyxl) rồi thử lại.")
    input("Nhấn Enter để thoát...")
    sys.exit(1)

THU_MUC = os.path.dirname(os.path.abspath(__file__))
CONG_CHROME = "http://localhost:9222"


# ---------------------------------------------------------------- Excel
def chu(v, ten_cot=""):
    """Đổi giá trị ô Excel thành chữ đúng như nhìn thấy (giữ số 0 đầu CCCD, ngày dd/mm/yyyy)."""
    if v is None:
        return ""
    if isinstance(v, dt.datetime) or isinstance(v, dt.date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    # Excel hay làm mất số 0 đầu của CCCD (12 số) / CMND (9 số)
    if re.search(r"cccd|cmnd|can_?cuoc|giay_?to|dinh_?danh", ten_cot, re.I) and s.isdigit():
        if len(s) == 11:
            s = "0" + s
        elif len(s) == 8:
            s = "0" + s
    return s


def doc_excel(duong_dan, ten_sheet=None):
    wb = load_workbook(duong_dan, data_only=True, read_only=True)
    ws = wb[ten_sheet] if ten_sheet else wb.worksheets[0]
    dong = list(ws.iter_rows(values_only=True))
    if not dong:
        return [], []
    tieu_de = [chu(x) for x in dong[0]]
    ds = []
    for so, r in enumerate(dong[1:], start=2):
        if r is None or all(x is None or str(x).strip() == "" for x in r):
            continue
        ds.append((so, {tieu_de[i]: chu(r[i], tieu_de[i]) for i in range(min(len(tieu_de), len(r))) if tieu_de[i]}))
    return tieu_de, ds


# ---------------------------------------------------------------- Kịch bản
def doc_kich_ban(duong_dan):
    cai_dat = {"TRANG": "", "NGHI": 0.3, "EXCEL": "du_lieu.xlsx", "SHEET": ""}
    buoc = []
    with open(duong_dan, encoding="utf-8-sig") as f:
        for so, dong in enumerate(f, start=1):
            dong = dong.strip()
            if not dong or dong.startswith("#"):
                continue
            lenh, _, tham_so = dong.partition(" ")
            lenh = lenh.upper()
            tham_so = tham_so.strip()
            if lenh in ("TRANG", "EXCEL", "SHEET"):
                cai_dat[lenh] = tham_so
            elif lenh == "NGHI":
                cai_dat["NGHI"] = float(tham_so.replace(",", "."))
            elif lenh in ("BAM", "GO", "XOA", "PHIM", "CHO", "CHO_THAY", "CHON", "DUNG"):
                buoc.append((so, lenh, tham_so))
            else:
                raise SystemExit(f"kich_ban.txt dòng {so}: không hiểu lệnh '{lenh}'")
    return cai_dat, buoc


def thay_cot(mau, du_lieu, so_dong_kb):
    def _thay(m):
        ten = m.group(1).strip()
        if ten not in du_lieu:
            raise SystemExit(f"kich_ban.txt dòng {so_dong_kb}: Excel không có cột '{ten}'. "
                             f"Các cột đang có: {', '.join(du_lieu.keys())}")
        return du_lieu[ten]
    return re.sub(r"\{([^}]+)\}", _thay, mau)


# ---------------------------------------------------------------- Trình duyệt
def tim_trong_moi_khung(page, selector):
    """Tìm phần tử trong trang và cả các khung (iframe) bên trong."""
    for khung in page.frames:
        try:
            loc = khung.locator(selector)
            n = loc.count()
            for i in range(n):
                if loc.nth(i).is_visible():
                    return loc.nth(i)
        except Exception:
            pass
    return None


def tim_chu_trong_moi_khung(page, chu_can_tim):
    for khung in page.frames:
        try:
            loc = khung.get_by_text(chu_can_tim, exact=True)
            n = loc.count()
            for i in range(n - 1, -1, -1):  # ưu tiên cái hiện sau cùng (bảng vừa bật lên)
                if loc.nth(i).is_visible():
                    return loc.nth(i)
        except Exception:
            pass
    return None


def cho_den_khi(ham, giay=15):
    het = time.time() + giay
    while time.time() < het:
        kq = ham()
        if kq is not None:
            return kq
        time.sleep(0.2)
    return None


def chon_trang(browser, tu_khoa):
    cac_trang = [pg for ctx in browser.contexts for pg in ctx.pages
                 if not pg.url.startswith(("devtools://", "chrome-extension://"))]
    if tu_khoa:
        khop = [pg for pg in cac_trang if tu_khoa.lower() in (pg.title() + " " + pg.url).lower()]
        if len(khop) >= 1:
            return khop[0]
        print(f"Không thấy tab nào có chữ '{tu_khoa}' trong tiêu đề/địa chỉ.")
    if len(cac_trang) == 1:
        return cac_trang[0]
    print("Các tab đang mở:")
    for i, pg in enumerate(cac_trang, start=1):
        print(f"  {i}. {pg.title()}  ({pg.url[:80]})")
    so = int(input("Chọn số thứ tự tab cần nhập liệu: "))
    return cac_trang[so - 1]


# ---------------------------------------------------------------- Chạy 1 khách hàng
class BoQua(Exception):
    pass


def chay_mot_dong(page, buoc, du_lieu, nghi):
    kb = page.keyboard
    for so_kb, lenh, ts in buoc:
        if lenh == "BAM":
            el = cho_den_khi(lambda: tim_trong_moi_khung(page, ts))
            if el is None:
                raise RuntimeError(f"[kịch bản dòng {so_kb}] không thấy ô '{ts}' trên màn hình")
            el.click()
        elif lenh == "GO":
            gia_tri = thay_cot(ts, du_lieu, so_kb)
            if gia_tri:
                kb.type(gia_tri, delay=20)
        elif lenh == "XOA":
            kb.press("Control+A")
            kb.press("Backspace")
        elif lenh == "PHIM":
            m = re.match(r"^(\S+)(?:\s+x\s*(\d+))?$", ts, re.I)
            if not m:
                raise SystemExit(f"kich_ban.txt dòng {so_kb}: viết sai lệnh PHIM '{ts}' (vd: PHIM ArrowDown x2)")
            for _ in range(int(m.group(2) or 1)):
                kb.press(m.group(1))
                time.sleep(nghi)
            continue
        elif lenh == "CHO":
            time.sleep(float(ts.replace(",", ".")))
            continue
        elif lenh == "CHO_THAY":
            muc = thay_cot(ts, du_lieu, so_kb)
            ham = (lambda: tim_trong_moi_khung(page, muc[4:].strip())) if muc.lower().startswith("css:") \
                else (lambda: tim_chu_trong_moi_khung(page, muc))
            if cho_den_khi(ham) is None:
                raise RuntimeError(f"[kịch bản dòng {so_kb}] chờ 15 giây không thấy '{muc}' hiện ra")
        elif lenh == "CHON":
            muc = thay_cot(ts, du_lieu, so_kb)
            el = cho_den_khi(lambda: tim_chu_trong_moi_khung(page, muc), giay=10)
            if el is None:
                raise RuntimeError(f"[kịch bản dòng {so_kb}] trong danh sách không có dòng '{muc}'")
            el.click()
        elif lenh == "DUNG":
            loi_nhac = ts or "Kiểm tra màn hình"
            tl = input(f"   >> {loi_nhac}. Enter = làm tiếp | b = bỏ qua khách này | q = dừng hẳn: ").strip().lower()
            if tl == "q":
                raise KeyboardInterrupt
            if tl == "b":
                raise BoQua()
        time.sleep(nghi)


# ---------------------------------------------------------------- Chính
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--kich-ban", default=os.path.join(THU_MUC, "kich_ban.txt"))
    ap.add_argument("--excel", default="")
    ap.add_argument("--thu", action="store_true", help="chỉ chạy dòng đầu tiên")
    ap.add_argument("--tu", type=int, default=0, help="chạy từ dòng Excel số này")
    ap.add_argument("--den", type=int, default=0, help="chạy đến dòng Excel số này")
    ap.add_argument("--cong", default=CONG_CHROME)
    a = ap.parse_args()

    cai_dat, buoc = doc_kich_ban(a.kich_ban)
    file_excel = a.excel or cai_dat["EXCEL"]
    if not os.path.isabs(file_excel):
        file_excel = os.path.join(THU_MUC, file_excel)
    tieu_de, ds = doc_excel(file_excel, cai_dat["SHEET"] or None)
    ds = [(so, d) for so, d in ds if (not a.tu or so >= a.tu) and (not a.den or so <= a.den)]
    if a.thu:
        ds = ds[:1]
    if not ds:
        print("Excel không có dòng nào cần nhập.")
        return
    print(f"Excel: {os.path.basename(file_excel)} | Cột: {', '.join(t for t in tieu_de if t)}")
    print(f"Sẽ nhập {len(ds)} khách hàng (dòng {ds[0][0]} đến {ds[-1][0]}).")

    nhat_ky = os.path.join(THU_MUC, "nhat_ky.csv")
    moi = not os.path.exists(nhat_ky)
    with sync_playwright() as p, open(nhat_ky, "a", newline="", encoding="utf-8-sig") as f_log:
        log = csv.writer(f_log)
        if moi:
            log.writerow(["Thoi_gian", "Dong_Excel", "Ket_qua", "Ghi_chu", "Du_lieu"])
        try:
            browser = p.chromium.connect_over_cdp(a.cong)
        except Exception:
            print("Không kết nối được Chrome. Hãy đóng hết Chrome rồi mở bằng file MoChrome.bat, đăng nhập, "
                  "mở đúng màn hình nhập liệu rồi chạy lại.")
            input("Nhấn Enter để thoát...")
            return
        page = chon_trang(browser, cai_dat["TRANG"])
        print(f"Nhập vào tab: {page.title()}")
        input("Mở sẵn màn hình nhập khách hàng mới, rồi nhấn Enter ở đây để bắt đầu (Ctrl+C để dừng)...")
        page.bring_to_front()

        xong = 0
        for so, du_lieu in ds:
            ten = next(iter(du_lieu.values()), "")
            print(f"- Dòng {so}: {ten} ...", end=" ", flush=True)
            bay_gio = dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            try:
                chay_mot_dong(page, buoc, du_lieu, cai_dat["NGHI"])
                print("xong")
                xong += 1
                log.writerow([bay_gio, so, "XONG", "", " | ".join(du_lieu.values())])
            except BoQua:
                print("bỏ qua")
                log.writerow([bay_gio, so, "BO_QUA", "", " | ".join(du_lieu.values())])
            except KeyboardInterrupt:
                print("\nĐã dừng theo yêu cầu.")
                log.writerow([bay_gio, so, "DUNG", "người dùng dừng", " | ".join(du_lieu.values())])
                print(f"Kiểm tra lại khách dòng {so} trên phần mềm. Chạy tiếp: python NhapTuExcel.py --tu {so}")
                break
            except Exception as e:
                print("LỖI")
                print(f"   {e}")
                log.writerow([bay_gio, so, "LOI", str(e), " | ".join(du_lieu.values())])
                print(f"Đã dừng để không nhập bừa. Xem lại màn hình, rồi chạy tiếp: python NhapTuExcel.py --tu {so}")
                break
            f_log.flush()
        print(f"Đã nhập xong {xong}/{len(ds)} khách. Nhật ký: nhat_ky.csv")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nĐã dừng.")

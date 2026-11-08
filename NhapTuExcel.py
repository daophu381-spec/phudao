# -*- coding: utf-8 -*-
"""
NHẬP TỪ EXCEL VÀO PHẦN MỀM NGHIỆP VỤ TRÊN CHROME BẰNG PHÍM (Enter, F5, F2, Lên/Xuống...)

- Điều khiển chính cửa sổ Chrome bạn đang mở và đã đăng nhập (không mở Chrome mới).
- Đọc từng dòng Excel, làm đúng các bước ghi trong file kich_ban.txt.
- Không gửi dữ liệu đi đâu; chỉ gõ phím vào trang đang mở.

Cách chạy: nháy đúp file BatDau.bat rồi chọn số trong menu.
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


# ---------------------------------------------------------------- Ghi kịch bản tự động
JS_GHI = r"""
(() => {
  if (window.__dangGhiPhim) return; window.__dangGhiPhim = 1;
  const gui = (x) => { try { window.__ghiPhim(x); } catch (e) {} };
  const sel = (el) => {
    if (el.id && document.querySelectorAll('#' + CSS.escape(el.id)).length === 1) return '#' + CSS.escape(el.id);
    if (el.name && document.getElementsByName(el.name).length === 1) return el.tagName.toLowerCase() + '[name="' + el.name + '"]';
    const p = []; let n = el;
    while (n && n.nodeType === 1 && n !== document.body) {
      let i = 1, s = n; while ((s = s.previousElementSibling)) if (s.tagName === n.tagName) i++;
      p.unshift(n.tagName.toLowerCase() + ':nth-of-type(' + i + ')'); n = n.parentElement;
    }
    return 'body > ' + p.join(' > ');
  };
  const laO = (el) => el && el.matches && el.matches('input, textarea, select, [contenteditable="true"]');
  const gt = (el) => (el && 'value' in el) ? el.value : null;
  document.addEventListener('focusin', (e) => { if (laO(e.target)) gui({loai: 'vao', sel: sel(e.target), giaTri: gt(e.target)}); }, true);
  document.addEventListener('mousedown', (e) => {
    const t = e.target;
    if (laO(t)) gui({loai: 'chuot_o', sel: sel(t)});
    else gui({loai: 'chuot_chu', chu: (t.innerText || '').trim().slice(0, 80)});
  }, true);
  const DB = /^(Enter|Tab|Escape|F\d{1,2}|ArrowUp|ArrowDown|ArrowLeft|ArrowRight|PageUp|PageDown|Home|End|Delete|Insert)$/;
  document.addEventListener('keydown', (e) => {
    const k = e.key;
    if (['Control', 'Alt', 'Shift', 'Meta'].includes(k)) return;
    if (!DB.test(k) && !e.ctrlKey && !e.altKey) return;
    let ten = (e.ctrlKey ? 'Control+' : '') + (e.altKey ? 'Alt+' : '') + (e.shiftKey && (DB.test(k) || e.ctrlKey || e.altKey) ? 'Shift+' : '') + (k.length === 1 ? k.toUpperCase() : k);
    const a = document.activeElement;
    gui({loai: 'phim', phim: ten, sel: laO(a) ? sel(a) : null, giaTri: laO(a) ? gt(a) : null});
  }, true);
})();
"""


def _chuan(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _tim_cot(gia_tri, mau):
    """Giá trị vừa gõ trùng với cột nào của dòng Excel mẫu."""
    g = _chuan(gia_tri)
    if not g:
        return None
    for cot, v in mau.items():
        if _chuan(v) == g:
            return cot
    so = re.sub(r"\D", "", g)
    for cot, v in mau.items():  # ngày gõ 05082021 hay 5/8/2021 vẫn nhận ra
        vs = re.sub(r"\D", "", str(v))
        if so and len(so) >= 6 and vs == so:
            return cot
    return None


def tao_kich_ban(su_kien, mau):
    dong = []
    da_bam_dau = False
    gia_tri_luc_vao = {}
    t_truoc = None

    def go(sel, gia_tri):
        cu = gia_tri_luc_vao.get(sel, "")
        if gia_tri is None or gia_tri == cu:
            return
        if cu:
            dong.append("XOA")
        cot = _tim_cot(gia_tri, mau)
        dong.append(f"GO {{{cot}}}" if cot else f"GO {gia_tri}")
        gia_tri_luc_vao[sel] = gia_tri

    for t, sk in su_kien:
        if t_truoc is not None and t - t_truoc > 1.5 and dong:
            dong.append("CHO 1")
        t_truoc = t
        loai = sk.get("loai")
        if loai == "vao":
            gia_tri_luc_vao[sk["sel"]] = sk.get("giaTri") or ""
        elif loai == "chuot_o":
            if not da_bam_dau or dong:
                dong.append(f"BAM {sk['sel']}")
            da_bam_dau = True
        elif loai == "chuot_chu":
            chu_bam = sk.get("chu") or ""
            if chu_bam and "\n" not in chu_bam:
                cot = _tim_cot(chu_bam, mau)
                dong.append(f"CHON {{{cot}}}" if cot else f"CHON {chu_bam}")
        elif loai == "phim":
            if sk.get("sel"):
                go(sk["sel"], sk.get("giaTri"))
            phim = sk["phim"]
            if dong and dong[-1].startswith(f"PHIM {phim}"):
                m = re.match(r"^PHIM (\S+)(?: x(\d+))?$", dong[-1])
                dong[-1] = f"PHIM {phim} x{int(m.group(2) or 1) + 1}"
            else:
                dong.append(f"PHIM {phim}")
            if phim.startswith("F") and phim[1:].isdigit():
                dong.append("CHO 0.8")  # chờ bảng danh mục bật lên
    # bỏ các CHO thừa đứng liền nhau
    gon = []
    for d in dong:
        if d.startswith("CHO") and gon and gon[-1].startswith("CHO"):
            continue
        gon.append(d)
    return gon


def ghi_kich_ban(page, file_excel, sheet, duong_dan_kb):
    _, ds = doc_excel(file_excel, sheet or None)
    if not ds:
        print("File Excel chưa có dữ liệu (dòng 1 là tên cột, từ dòng 2 là khách hàng).")
        return None
    so_mau, mau = ds[0]
    print("\n=========== GHI KỊCH BẢN ===========")
    print(f"Bạn hãy nhập TAY khách hàng ở dòng {so_mau} của Excel, đúng như mọi ngày:")
    for cot, v in mau.items():
        print(f"     {cot}: {v}")
    print("Máy sẽ ghi lại các phím bạn bấm (Enter, F5, F2, mũi tên...) và các chữ bạn gõ.")
    print("  - Bắt đầu bằng việc BẤM CHUỘT vào ô đầu tiên, sau đó cố gắng chỉ dùng bàn phím.")
    print("  - Gõ đúng như trong Excel (để máy nhận ra ô nào lấy cột nào).")
    print("  - Lưu xong khách đó (và bấm F2 nếu cần về màn hình nhập mới) thì quay lại cửa sổ đen này nhấn Enter.\n")

    su_kien = []
    page.expose_binding("__ghiPhim", lambda nguon, x: su_kien.append((time.time(), x)))
    page.add_init_script(JS_GHI)
    for khung in page.frames:
        try:
            khung.evaluate(JS_GHI)
        except Exception:
            pass

    import threading
    xong = threading.Event()
    threading.Thread(target=lambda: (input(">>> Nhập xong khách mẫu thì nhấn Enter ở đây... "), xong.set()),
                     daemon=True).start()
    page.bring_to_front()
    da_tiem = set()
    while not xong.is_set():
        page.wait_for_timeout(300)
        for khung in page.frames:  # khung mới hiện ra cũng phải ghi
            if khung not in da_tiem:
                try:
                    khung.evaluate(JS_GHI)
                    da_tiem.add(khung)
                except Exception:
                    pass

    buoc = tao_kich_ban(su_kien, mau)
    if not buoc:
        print("Không ghi được phím nào. Hãy chắc là bạn nhập trên đúng tab Chrome rồi ghi lại.")
        return None
    if os.path.exists(duong_dan_kb):
        os.replace(duong_dan_kb, duong_dan_kb.replace(".txt", "_cu.txt"))
    from urllib.parse import urlparse
    may_chu = urlparse(page.url).netloc or page.title()
    with open(duong_dan_kb, "w", encoding="utf-8") as f:
        f.write("# Kịch bản do máy tự ghi lúc " + dt.datetime.now().strftime("%d/%m/%Y %H:%M") + "\n")
        f.write("# Sửa được bằng Notepad. Hướng dẫn các lệnh: xem README.\n")
        f.write(f"TRANG {may_chu}\nEXCEL {os.path.basename(file_excel)}\n")
        if sheet:
            f.write(f"SHEET {sheet}\n")
        f.write("NGHI 0.3\n\n")
        f.write("\n".join(buoc) + "\n")
        f.write("\n# Dừng để bạn nhìn lại sau mỗi khách. Chạy ổn rồi thì xoá dòng dưới đây.\n")
        f.write("DUNG Xong 1 khách - kiểm tra trên màn hình\n")
    print("\nĐã ghi kịch bản:")
    for b in buoc:
        print("   " + b)
    khong_khop = [b for b in buoc if b.startswith("GO ") and "{" not in b]
    if khong_khop:
        print("\nChú ý: các dòng sau gõ chữ CỐ ĐỊNH (không thấy trùng cột Excel nào), mọi khách sẽ gõ y như vậy:")
        for b in khong_khop:
            print("   " + b)
        print("Nếu chỗ đó phải lấy theo Excel: lần ghi sau gõ đúng y như ô Excel, hoặc mở kịch bản sửa thành GO {Tên_cột}.")
    return so_mau


# ---------------------------------------------------------------- Chạy
def ket_noi_chrome(p, cong):
    try:
        return p.chromium.connect_over_cdp(cong)
    except Exception:
        print("\nKHÔNG KẾT NỐI ĐƯỢC CHROME.")
        print("Hãy: 1) Đóng hết Chrome  2) Mở Chrome bằng file MoChrome.bat  3) Đăng nhập, mở màn hình nhập liệu")
        print("     4) Chạy lại công cụ này.")
        return None


def chay(page, cai_dat, buoc, ds, file_tien_do):
    nhat_ky = os.path.join(THU_MUC, "nhat_ky.csv")
    moi = not os.path.exists(nhat_ky)
    xong = 0
    with open(nhat_ky, "a", newline="", encoding="utf-8-sig") as f_log:
        log = csv.writer(f_log)
        if moi:
            log.writerow(["Thoi_gian", "Dong_Excel", "Ket_qua", "Ghi_chu", "Du_lieu"])
        page.bring_to_front()
        for so, du_lieu in ds:
            ten = next(iter(du_lieu.values()), "")
            print(f"- Dòng {so}: {ten} ...", end=" ", flush=True)
            bay_gio = dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            try:
                chay_mot_dong(page, buoc, du_lieu, cai_dat["NGHI"])
                print("xong")
                xong += 1
                log.writerow([bay_gio, so, "XONG", "", " | ".join(du_lieu.values())])
                ghi_tien_do(file_tien_do, so)
            except BoQua:
                print("bỏ qua")
                log.writerow([bay_gio, so, "BO_QUA", "", " | ".join(du_lieu.values())])
                ghi_tien_do(file_tien_do, so)
            except KeyboardInterrupt:
                print("\nĐã dừng theo yêu cầu.")
                log.writerow([bay_gio, so, "DUNG", "người dùng dừng", " | ".join(du_lieu.values())])
                print(f"Hãy xem khách dòng {so} trên phần mềm đã lưu chưa. Lần chạy sau công cụ sẽ hỏi bắt đầu từ dòng {so}.")
                break
            except Exception as e:
                print("LỖI")
                print(f"   {e}")
                log.writerow([bay_gio, so, "LOI", str(e), " | ".join(du_lieu.values())])
                print(f"Đã dừng để không nhập bừa. Xem lại màn hình. Lần chạy sau công cụ sẽ hỏi bắt đầu từ dòng {so}.")
                break
            f_log.flush()
    print(f"\nĐã nhập {xong}/{len(ds)} khách. Chi tiết từng dòng: nhat_ky.csv")


def doc_tien_do(f):
    try:
        return int(open(f, encoding="utf-8").read().strip())
    except Exception:
        return 0


def ghi_tien_do(f, so):
    with open(f, "w", encoding="utf-8") as g:
        g.write(str(so))


def hoi_file_excel(mac_dinh):
    if os.path.exists(mac_dinh):
        return mac_dinh
    print(f"\nChưa thấy file Excel '{os.path.basename(mac_dinh)}' trong thư mục công cụ.")
    print("Hãy KÉO file Excel của bạn thả vào cửa sổ đen này rồi nhấn Enter")
    print("(hoặc copy file đó vào thư mục công cụ và đổi tên thành du_lieu.xlsx):")
    while True:
        duong = input("> ").strip().strip('"').strip("'").strip()
        if duong.lower().endswith((".xlsx", ".xlsm")) and os.path.exists(duong):
            return duong
        print("Không đúng file Excel (.xlsx). Thử lại:")


def menu(a):
    kb_path = a.kich_ban
    file_tien_do = os.path.join(THU_MUC, "tien_do.txt")
    co_kb = os.path.exists(kb_path)
    cai_dat, buoc = doc_kich_ban(kb_path) if co_kb else ({"TRANG": "", "NGHI": 0.3, "EXCEL": "du_lieu.xlsx", "SHEET": ""}, [])
    file_excel = a.excel or cai_dat["EXCEL"]
    if not os.path.isabs(file_excel):
        file_excel = os.path.join(THU_MUC, file_excel)

    print("==================================================")
    print("   NHẬP TỰ ĐỘNG TỪ EXCEL VÀO PHẦN MỀM (bằng phím)")
    print("==================================================")
    print("  1. Ghi kịch bản: tôi nhập tay 1 khách, máy học theo")
    print("  2. Thử nhập 1 khách tiếp theo")
    print("  3. Nhập tất cả các khách còn lại")
    print("  4. Mở kịch bản bằng Notepad để xem/sửa")
    print("  0. Thoát")
    chon = input("Chọn số rồi nhấn Enter: ").strip()
    if chon not in ("1", "2", "3", "4"):
        return
    if chon == "4":
        if co_kb:
            os.startfile(kb_path) if hasattr(os, "startfile") else print(open(kb_path, encoding="utf-8").read())
        else:
            print("Chưa có kịch bản. Chọn 1 để ghi.")
        return
    if chon in ("2", "3") and not buoc:
        print("Chưa có kịch bản. Hãy chọn 1 để ghi kịch bản trước.")
        return

    file_excel = hoi_file_excel(file_excel)
    with sync_playwright() as p:
        browser = ket_noi_chrome(p, a.cong)
        if browser is None:
            return
        page = chon_trang(browser, "" if chon == "1" else cai_dat["TRANG"])
        print(f"Đang làm việc với tab: {page.title()}")
        if chon == "1":
            input("Mở sẵn màn hình nhập khách hàng MỚI trên Chrome, rồi nhấn Enter ở đây...")
            so_mau = ghi_kich_ban(page, file_excel, cai_dat.get("SHEET"), kb_path)
            if so_mau:
                ghi_tien_do(file_tien_do, so_mau)
                print(f"\nKhách dòng {so_mau} bạn đã nhập tay. Lần sau chọn 2 để máy thử nhập khách tiếp theo.")
            return

        _, ds = doc_excel(file_excel, cai_dat["SHEET"] or None)
        da_xong = doc_tien_do(file_tien_do)
        con_lai = [x for x in ds if x[0] > da_xong]
        if not con_lai:
            print("Đã nhập hết các dòng trong Excel.")
            tl = input("Gõ số dòng muốn nhập lại từ đó (hoặc Enter để thoát): ").strip()
            if not tl.isdigit():
                return
            con_lai = [x for x in ds if x[0] >= int(tl)]
        else:
            tl = input(f"Bắt đầu từ dòng {con_lai[0][0]} của Excel ({con_lai[0][1] and next(iter(con_lai[0][1].values()))})? "
                       f"Enter = đồng ý, hoặc gõ số dòng khác: ").strip()
            if tl.isdigit():
                con_lai = [x for x in ds if x[0] >= int(tl)]
        if chon == "2":
            con_lai = con_lai[:1]
        if not con_lai:
            print("Không có dòng nào để nhập.")
            return
        print(f"Sẽ nhập {len(con_lai)} khách (dòng {con_lai[0][0]} đến {con_lai[-1][0]}). Dừng gấp: Ctrl+C.")
        input("Mở sẵn màn hình nhập khách hàng mới trên Chrome, rồi nhấn Enter ở đây để bắt đầu...")
        chay(page, cai_dat, buoc, con_lai, file_tien_do)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--kich-ban", default=os.path.join(THU_MUC, "kich_ban.txt"))
    ap.add_argument("--excel", default="")
    ap.add_argument("--cong", default=CONG_CHROME)
    a = ap.parse_args()
    menu(a)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nĐã dừng.")

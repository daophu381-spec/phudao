/* ============================================================
   ĐIỀN NHANH: HỌ VÀ TÊN – SỐ CCCD – NGÀY CẤP  (dán vào F12 > Console)
   - Chỉ điền vào các ô trên trang đang mở. KHÔNG gửi dữ liệu đi đâu.
   - Cách dùng: copy 1 dòng 3 ô trong Excel (Họ tên | Số CCCD | Ngày cấp),
     chạy đoạn này, dán vào hộp hỏi, Enter. Thứ tự cột không quan trọng.
   - Lần sau trong cùng trang: gõ  dien()  rồi Enter (hoặc nhấn mũi tên lên).
   ============================================================ */
(() => {
  /* (1) Tự nhận ô bị sai hoặc không thấy? Ghi "selector" của ô vào đây.
         Cách lấy: chuột phải vào ô > Kiểm tra (Inspect) > xem id="..."
         rồi ghi '#id' . Ví dụ: hoTen: '#txtHoTen'  */
  const CHON_O = { hoTen: '', soCCCD: '', ngayCap: '' };

  /* (2) Có muốn tự VIẾT HOA họ tên không: true / false */
  const VIET_HOA_HO_TEN = false;

  /* (3) Từ khoá nhận diện ô (viết không dấu, chữ thường).
         co  = nhãn/tên ô có chứa 1 trong các chữ này thì nhận
         tru = có chứa các chữ này thì bỏ qua */
  const TU_KHOA = {
    ngayCap: {
      co: ['ngay cap', 'ngaycap', 'issue date', 'issuedate', 'issued date', 'dateofissue', 'date of issue'],
      tru: ['het han', 'sinh']
    },
    soCCCD: {
      co: ['cccd', 'can cuoc', 'cmnd', 'cmt', 'so giay to', 'sogiayto', 'giay to tuy than', 'dinh danh',
           'idcard', 'id card', 'idno', 'id number', 'identity', 'so ho chieu'],
      tru: ['ngay', 'noi cap', 'noicap', 'date', 'loai giay', 'loaigiay']
    },
    hoTen: {
      co: ['ho va ten', 'ho ten', 'hoten', 'hovaten', 'fullname', 'full name', 'ten khach hang', 'tenkhachhang',
           'ten kh', 'tenkh', 'customer name', 'customername'],
      tru: ['ho ten cha', 'ho ten me', 'ho ten vo', 'ho ten chong', 'uy quyen', 'thua ke', 'nguoi giam ho']
    }
  };

  const TEN_O = { hoTen: 'Họ và tên', soCCCD: 'Số CCCD', ngayCap: 'Ngày cấp' };

  const boDau = (s) => String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[đĐ]/g, 'd').toLowerCase()
    .replace(/[_\-:*()\[\]]+/g, ' ').replace(/\s+/g, ' ').trim();

  /* ---- Tách dòng dán từ Excel thành họ tên / CCCD / ngày cấp ---- */
  const laNgay = (s) => /^\s*(\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{4}|\d{4}[\/.\-]\d{1,2}[\/.\-]\d{1,2})\s*$/.test(s);
  const chuanNgay = (s) => {
    const p = s.trim().split(/[\/.\-]/);
    let d, m, y;
    if (p[0].length === 4) { y = p[0]; m = p[1]; d = p[2]; } else { d = p[0]; m = p[1]; y = p[2]; }
    return { d: d.padStart(2, '0'), m: m.padStart(2, '0'), y: y };
  };
  /* Mã QR trên thẻ CCCD gắn chip (quét bằng máy quét/điện thoại):
     SốCCCD|SốCMND cũ|Họ tên|Ngày sinh ddmmyyyy|Giới tính|Địa chỉ|Ngày cấp ddmmyyyy */
  const docQR = (dong) => {
    const p = dong.trim().split('|');
    if (p.length < 7 || !/^\d{12}$/.test(p[0].trim())) return null;
    const nc = p[6].trim();
    const kq = {
      hoTen: p[2].trim().replace(/\s+/g, ' '),
      soCCCD: p[0].trim(),
      ngayCap: /^\d{8}$/.test(nc) ? nc.slice(0, 2) + '/' + nc.slice(2, 4) + '/' + nc.slice(4) : nc
    };
    if (VIET_HOA_HO_TEN) kq.hoTen = kq.hoTen.toLocaleUpperCase('vi');
    return kq;
  };
  const tachDong = (dong) => {
    const qr = docQR(dong);
    if (qr) return qr;
    let phan = dong.split('\t');
    if (phan.length < 2) phan = dong.split(/[|;\n]/);
    if (phan.length < 2) phan = dong.split(',');
    phan = phan.map((x) => x.trim()).filter(Boolean);
    const kq = { hoTen: '', soCCCD: '', ngayCap: '' };
    for (const x of phan) {
      const so = x.replace(/[\s.]/g, '');
      if (!kq.ngayCap && laNgay(x)) { kq.ngayCap = x; continue; }
      if (!kq.soCCCD && /^\d{8,12}$/.test(so)) {
        /* Excel hay làm mất số 0 đầu: CCCD 11 số -> thêm 0 cho đủ 12, CMND 8 số -> đủ 9 */
        kq.soCCCD = so.length === 11 ? '0' + so : (so.length === 8 ? '0' + so : so);
        continue;
      }
      if (!kq.hoTen) kq.hoTen = x.replace(/\s+/g, ' ');
    }
    if (VIET_HOA_HO_TEN) kq.hoTen = kq.hoTen.toLocaleUpperCase('vi');
    return kq;
  };

  /* ---- Lấy tất cả trang con (iframe cùng hệ thống) ---- */
  const layTrang = (doc, ds) => {
    ds.push(doc);
    for (const f of doc.querySelectorAll('iframe, frame')) {
      try { if (f.contentDocument) layTrang(f.contentDocument, ds); } catch (e) { /* khác tên miền, bỏ qua */ }
    }
    return ds;
  };

  const LOAI_BO_QUA = ['hidden', 'button', 'submit', 'reset', 'checkbox', 'radio', 'file', 'image', 'password'];
  const dangHien = (el) => {
    if (el.disabled) return false;
    if (!(el.offsetWidth || el.offsetHeight || el.getClientRects().length)) return false;
    const st = el.ownerDocument.defaultView.getComputedStyle(el);
    return st.visibility !== 'hidden' && st.display !== 'none';
  };

  /* Chữ mô tả 1 ô: nhãn, placeholder, name, id + chữ đứng ngay trước ô */
  const moTa = (el) => {
    const doc = el.ownerDocument;
    const ds = [];
    if (el.labels) for (const l of el.labels) ds.push(l.innerText);
    const by = el.getAttribute('aria-labelledby');
    if (by) for (const id of by.split(/\s+/)) { const n = doc.getElementById(id); if (n) ds.push(n.innerText); }
    for (const a of ['aria-label', 'placeholder', 'title', 'name', 'id', 'formcontrolname', 'ng-model', 'data-field']) {
      const v = el.getAttribute(a); if (v) ds.push(v);
    }
    let gan = '';
    let n = el;
    for (let cap = 0; cap < 4 && n && !gan; cap++) {
      let p = n.previousElementSibling;
      while (p && !gan) {
        const t = (p.innerText || p.textContent || '').trim();
        if (t && t.length <= 80) gan = t;
        else if (t) break;
        p = p.previousElementSibling;
      }
      n = n.parentElement;
    }
    return { rieng: boDau(ds.join(' | ')), gan: boDau(gan) };
  };

  const khop = (chu, tk) => tk.co.some((k) => chu.includes(k)) && !tk.tru.some((k) => chu.includes(k));

  const timO = () => {
    const tatCa = [];
    for (const doc of layTrang(document, [])) {
      for (const el of doc.querySelectorAll('input, textarea')) {
        if (LOAI_BO_QUA.includes((el.type || '').toLowerCase())) continue;
        if (dangHien(el)) tatCa.push(el);
      }
    }
    const daDung = new Set();
    const o = {};
    for (const k of ['ngayCap', 'soCCCD', 'hoTen']) {
      if (CHON_O[k]) {
        for (const doc of layTrang(document, [])) {
          const el = doc.querySelector(CHON_O[k]);
          if (el) { o[k] = el; break; }
        }
        if (o[k]) { daDung.add(o[k]); continue; }
      }
      const conLai = tatCa.filter((el) => !daDung.has(el)).map((el) => ({ el, m: moTa(el) }));
      /* Ưu tiên khớp nhãn/tên riêng của ô, sau đó mới tới chữ đứng trước ô */
      const t = conLai.find((x) => khop(x.m.rieng, TU_KHOA[k])) || conLai.find((x) => khop(x.m.gan, TU_KHOA[k]));
      if (t) { o[k] = t.el; daDung.add(t.el); }
    }
    return o;
  };

  /* Gán giá trị kiểu "người gõ" để cả React/Angular/jQuery đều nhận */
  const ganGiaTri = (el, v) => {
    const w = el.ownerDocument.defaultView;
    const proto = el.tagName === 'TEXTAREA' ? w.HTMLTextAreaElement.prototype : w.HTMLInputElement.prototype;
    const set = Object.getOwnPropertyDescriptor(proto, 'value').set;
    set.call(el, v);
    for (const ev of ['input', 'change', 'keyup']) el.dispatchEvent(new w.Event(ev, { bubbles: true }));
    el.dispatchEvent(new w.Event('blur'));
    el.style.outline = '3px solid #2e7d32';
    setTimeout(() => { el.style.outline = ''; }, 2500);
  };

  const dien = (dong) => {
    if (dong === undefined) {
      dong = prompt('Dán 1 dòng từ Excel: Họ tên | Số CCCD | Ngày cấp\n(copy 3 ô cạnh nhau trong Excel rồi Ctrl+V)');
    }
    if (!dong) return;
    const dl = typeof dong === 'object' ? dong : tachDong(dong);
    const o = timO();
    const bang = [];
    const thieu = [];
    for (const k of ['hoTen', 'soCCCD', 'ngayCap']) {
      let v = dl[k] || '';
      const el = o[k];
      if (!v) { bang.push({ O: TEN_O[k], GiaTri: '(không có trong dòng dán)', KetQua: 'bỏ qua' }); continue; }
      if (!el) { thieu.push(TEN_O[k]); bang.push({ O: TEN_O[k], GiaTri: v, KetQua: 'KHÔNG TÌM THẤY Ô' }); continue; }
      if (k === 'ngayCap' && laNgay(v)) {
        const n = chuanNgay(v);
        v = (el.type === 'date') ? n.y + '-' + n.m + '-' + n.d : n.d + '/' + n.m + '/' + n.y;
      }
      ganGiaTri(el, v);
      bang.push({ O: TEN_O[k], GiaTri: v, KetQua: 'đã điền vào ' + (el.id ? '#' + el.id : (el.name || el.tagName)) });
    }
    console.table(bang);
    if (thieu.length) {
      alert('Không tìm thấy ô: ' + thieu.join(', ') + '.\nMở F12 > Console xem chi tiết, hoặc ghi selector vào CHON_O ở đầu đoạn mã.');
    }
  };

  window.dien = dien;
  dien();
})();

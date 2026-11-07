# Tự động Phê duyệt + In (HP M402) – bấm 1 lần

Công cụ nhỏ chạy trên Windows: **bấm 1 nút** là tự bấm lần lượt *Phê duyệt → (Đồng ý) → In → In ở hộp thoại máy in*, để bạn rảnh tay ký chứng từ.
Công cụ chỉ bấm chuột/phím thay bạn, không sửa phần mềm giao dịch và không đọc hay gửi dữ liệu đi đâu.

## 1. Cài đặt (1 lần)

1. Tải và cài **AutoHotkey v2** tại <https://www.autohotkey.com> (chọn *Download v2.0*).
   *Máy cơ quan khoá cài phần mềm thì nhờ bộ phận tin học cài giúp hoặc hỏi xem có được dùng không.*
2. Tải file `TuDongPheDuyetIn.ahk` về, để vào một thư mục riêng (ví dụ `D:\TuDongIn\`).
3. Nháy đúp vào file. Một nút nổi **▶ PHÊ DUYỆT + IN** hiện ra (luôn nằm trên cùng, kéo đi đâu cũng được).
4. Để công cụ tự chạy khi bật máy: nhấn `Win + R`, gõ `shell:startup`, rồi tạo shortcut của file `.ahk` vào thư mục vừa mở.

### Tự chọn máy in đang cắm (không cần vào Windows chọn tay)

Công cụ **tự làm** các việc sau, bạn không cần cài đặt gì thêm:

- Tắt chế độ *“Let Windows manage my default printer”* (để Windows không tự đổi lại máy in mặc định lung tung).
- Tìm máy in đang **bật/đang cắm**, ưu tiên máy có tên chứa `M402`, rồi đặt làm **máy in mặc định**. Bỏ qua các máy in ảo (PDF, XPS, OneNote, Fax...).
- Làm việc này khi mở công cụ, **mỗi 20 giây**, và **ngay trước mỗi lần bấm In**. Đi giao dịch đổi sang cắm máy in khác thì công cụ tự đổi theo.
- Nhấn **F10** để chọn ngay và xem tên máy in vừa được đặt.

Nếu máy in của bạn không có chữ `M402` trong tên, mở `cau_hinh.ini`, sửa dòng `TenChua=` trong mục `[MayIn]` (để trống = lấy máy in thật đầu tiên đang online).

**Cách thủ công (không dùng công cụ):** mở PowerShell và chạy một lần, hoặc lưu thành file `.ps1`:

```powershell
# Đặt máy in có tên chứa "M402" làm mặc định
(New-Object -ComObject WScript.Network).SetDefaultPrinter((Get-Printer | Where-Object Name -like '*M402*' | Select-Object -First 1).Name)
# Tắt "Let Windows manage my default printer"
Set-ItemProperty 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name LegacyDefaultPrinterMode -Value 1
```

## 2. Ghi các bước (1 lần)

Mở phần mềm giao dịch, **phóng to toàn màn hình** (lần nào dùng cũng phóng to để các nút luôn ở cùng chỗ), chọn sẵn 1 chứng từ cần duyệt, rồi:

| Việc | Cách làm |
|---|---|
| Nút **Phê duyệt** | Đặt chuột lên nút (không bấm) → nhấn **F8** → OK. Sau đó bấm tay nút đó. |
| Hộp hỏi **Đồng ý / Có** (nếu có) | Đặt chuột lên nút → **F8** → OK, rồi bấm tay. |
| Nút **In** | Đặt chuột lên nút → **F8** → OK, rồi bấm tay. |
| Hộp thoại máy in | Nhấn **F6** (ghi phím Enter) → OK. Hoặc đặt chuột lên nút *In/Print* → **F8**. |
| Phần mềm xử lý chậm | Nhấn **F5** để chèn thêm vài giây chờ ở đúng chỗ đó. |

Khi hộp “Xác nhận cửa sổ” hiện ra: nếu tên cửa sổ có phần **thay đổi theo từng chứng từ** (số CT, tên khách hàng, ngày…), hãy xoá phần đó, chỉ giữ phần cố định (ví dụ chỉ giữ chữ `Giao dịch lưu động`). Không chắc thì cứ để nguyên.

Nhấn **F1** để xem lại danh sách bước. Ghi sai thì nhấn **F7** để xoá và ghi lại.

## 3. Dùng hằng ngày

1. Kiểm tra và chọn chứng từ trên phần mềm.
2. Bấm nút **▶ PHÊ DUYỆT + IN** (hoặc nhấn **F9**).
3. Trong lúc máy tự bấm và in, bạn ký chứng từ trước đó.

Đang chạy muốn dừng gấp: nhấn **Esc**. Gặp lỗi (không thấy cửa sổ, mạng chậm…) công cụ sẽ **dừng lại và báo**, không bấm bừa. Hãy xem lại chứng từ trên phần mềm rồi mới chạy tiếp.

## 4. Tinh chỉnh trong `cau_hinh.ini`

File này tự tạo cạnh file `.ahk`, mở bằng Notepad:

- `NghiSauMoiBuoc_ms=700`: số mili-giây nghỉ sau mỗi lần bấm. Máy chậm thì tăng lên `1500`.
- `ChoCuaSoToiDa_giay=20`: thời gian tối đa chờ một cửa sổ (ví dụ hộp thoại in) hiện ra.
- `[BuocN]`: từng bước đã ghi (toạ độ X/Y, tên cửa sổ, phím). Sửa được trực tiếp.

Sửa xong thì nhấp chuột phải biểu tượng **H** dưới khay hệ thống và chọn *Reload Script*.

## 5. Nếu phần mềm chạy trên trình duyệt (Chrome/Edge)

Có thể bỏ hẳn hộp thoại in: tạo shortcut Chrome riêng với tham số
`"C:\Program Files\Google\Chrome\Application\chrome.exe" --kiosk-printing`
(đóng hết Chrome trước khi mở bằng shortcut này). Khi đó bấm In là in ngay ra máy in mặc định (HP M402), nên không cần ghi bước F6 cho hộp thoại in.

## Lưu ý

- Công cụ bấm theo **vị trí trên màn hình**. Nếu phần mềm đổi giao diện, đổi độ phân giải/zoom màn hình, hoặc cửa sổ không phóng to như lúc ghi, hãy nhấn **F7** và ghi lại.
- Việc **duyệt** vẫn do bạn chịu trách nhiệm: luôn xem chứng từ trước khi bấm nút. Công cụ chỉ thay bạn các thao tác bấm.

---

# Điền nhanh Họ tên – Số CCCD – Ngày cấp vào Chrome (F12)

File `DienCCCD.js`: copy 1 dòng trong Excel (3 ô **Họ tên | Số CCCD | Ngày cấp**, thứ tự nào cũng được), chạy mã, dán vào hộp hỏi, Enter là cả 3 ô được điền, ô vừa điền viền xanh. Mã chỉ điền vào trang đang mở, **không gửi dữ liệu đi đâu**.

- Tự tìm ô theo nhãn/tên ô (*Họ và tên, Tên khách hàng, Số CCCD/CMND, Số giấy tờ, Ngày cấp…*), kể cả ô nằm trong khung (iframe).
- Tự sửa số CCCD bị Excel làm mất số 0 đầu (11 số → thêm 0 cho đủ 12).
- Ngày `5/8/2021`, `05-08-2021`, `2021-08-05` đều được đổi về `05/08/2021` (ô kiểu lịch thì đổi đúng định dạng của ô).

## Cách 1 – Dán vào Console (F12)

1. Mở trang nhập liệu, nhấn **F12** → tab **Console**.
2. Lần đầu Chrome sẽ chặn dán: gõ chữ `allow pasting` rồi Enter.
3. Mở `DienCCCD.js`, copy hết, dán vào Console, Enter.
4. Hộp hỏi hiện ra → **Ctrl+V** dòng vừa copy trong Excel → Enter.
5. Người tiếp theo: copy dòng mới trong Excel, bấm vào Console, gõ `dien()` Enter (hoặc nhấn **↑** rồi Enter). Tải lại trang thì dán lại mã như bước 3.

**Lưu sẵn để khỏi dán lại:** F12 → tab **Sources** → **Snippets** → *New snippet*, dán mã, Ctrl+S. Từ đó chỉ cần F12 → Sources → Snippets → chọn snippet → **Ctrl+Enter**.

## Cách 2 – Nút bấm trên thanh dấu trang (nhanh nhất, không cần F12)

1. Nhấn **Ctrl+Shift+B** để hiện thanh dấu trang. Chuột phải lên thanh → *Thêm trang…*
2. Tên: `Điền CCCD`. Ô URL: dán **toàn bộ** nội dung file `DienCCCD_bookmarklet.txt` (bắt đầu bằng `javascript:`).
3. Dùng: copy dòng trong Excel → bấm nút **Điền CCCD** → Ctrl+V → Enter.

## Nếu điền sai ô / không tìm thấy ô

Chuột phải vào ô đó → **Kiểm tra (Inspect)** → xem `id="..."`. Mở `DienCCCD.js`, sửa dòng `CHON_O` ở đầu, ví dụ:

```js
const CHON_O = { hoTen: '#txtHoTen', soCCCD: '#txtSoGiayTo', ngayCap: '#dtNgayCap' };
```

Sửa xong nhớ tạo lại bookmarklet (hoặc gửi lại để được tạo giúp).

## Lưu ý

- Chỉ dán vào F12 mã mình đọc hiểu hoặc lấy từ nguồn tin cậy – dán mã lạ vào Console là cách kẻ gian hay dùng để lấy tài khoản.
- Máy cơ quan có quy định về công cụ tự động thì hỏi bộ phận tin học trước khi dùng.
- Luôn **nhìn lại 3 ô trước khi bấm Lưu**; mã chỉ điền thay, không bấm Lưu.

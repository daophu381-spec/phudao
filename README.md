# Tự động Phê duyệt + In (HP M402) – bấm 1 lần

Công cụ nhỏ chạy trên Windows: **bấm 1 nút** là tự bấm lần lượt *Phê duyệt → (Đồng ý) → In → In ở hộp thoại máy in*, để bạn rảnh tay ký chứng từ.
Công cụ chỉ bấm chuột/phím thay bạn, không sửa phần mềm giao dịch và không đọc hay gửi dữ liệu đi đâu.

## 1. Cài đặt (1 lần)

1. Tải và cài **AutoHotkey v2** tại <https://www.autohotkey.com> (chọn *Download v2.0*).
   *Máy cơ quan khoá cài phần mềm thì nhờ bộ phận tin học cài giúp hoặc hỏi xem có được dùng không.*
2. Tải file `TuDongPheDuyetIn.ahk` về, để vào một thư mục riêng (ví dụ `D:\TuDongIn\`).
3. Nháy đúp vào file. Một nút nổi **▶ PHÊ DUYỆT + IN** hiện ra (luôn nằm trên cùng, kéo đi đâu cũng được).
4. Để công cụ tự chạy khi bật máy: nhấn `Win + R`, gõ `shell:startup`, rồi tạo shortcut của file `.ahk` vào thư mục vừa mở.

**Nên làm thêm:** vào *Settings → Printers*, đặt **HP LaserJet Pro M402** làm *máy in mặc định* và tắt mục *“Let Windows manage my default printer”*. Khi đó hộp thoại in luôn chọn sẵn M402, chỉ cần bấm Enter.

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

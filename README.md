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

**Quét mã QR trên thẻ CCCD gắn chip:** dán (hoặc dùng máy quét mã vạch bắn) chuỗi QR của thẻ vào hộp hỏi, dạng `001090012345|123456789|Nguyễn Văn An|01011990|Nam|Địa chỉ|05082021`. Máy tự lấy đúng Họ tên, Số CCCD, Ngày cấp (bỏ qua ngày sinh, địa chỉ). Máy quét phải gõ được tiếng Việt có dấu (chế độ Unicode); máy quét rẻ hay làm hỏng dấu tên – thử trước bằng Notepad.

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

---

# Nhập hàng loạt từ Excel bằng phím (Enter, F5, F2, Lên/Xuống)

**Không cần biết lập trình.** Bạn nhập tay 1 khách hàng như mọi ngày, máy ghi lại bạn đã bấm những phím gì, rồi tự làm y hệt cho các khách còn lại trong Excel. Không phải viết code, không phải tìm mã ô.

Công cụ chỉ gõ phím vào Chrome của bạn, không gửi dữ liệu đi đâu.

## A. Cài đặt (làm 1 lần, khoảng 10 phút)

1. **Cài Python**: vào <https://www.python.org/downloads> → bấm nút *Download Python* màu vàng → mở file vừa tải.
   ⚠️ Ở màn hình cài đầu tiên, **tick ô "Add python.exe to PATH"** ở dưới cùng, rồi bấm *Install Now*.
   (Đã cài rồi mà quên tick thì gỡ đi cài lại, nhớ tick.)
2. Tải cả thư mục công cụ về máy, để vào ví dụ `D:\NhapTuDong\`.
3. Nháy đúp **`CaiDat.bat`**. Chờ chữ chạy xong, hiện *Xong* thì nhấn phím bất kỳ để đóng.
4. **Đóng hết Chrome**, rồi nháy đúp **`MoChrome.bat`**. Một cửa sổ Chrome mới mở ra → đăng nhập phần mềm nghiệp vụ như bình thường. Lần sau Chrome này nhớ đăng nhập.
   👉 Mỗi lần muốn dùng công cụ, luôn mở Chrome bằng `MoChrome.bat` (có thể kéo file này ra màn hình Desktop cho tiện).

## B. Chuẩn bị file Excel

- Đặt tên file là **`du_lieu.xlsx`** và để cùng thư mục công cụ (hoặc lúc chạy kéo thả file vào cửa sổ đen cũng được).
- **Dòng 1** là tên cột, đặt tên gì cũng được (*Họ và tên, Số CCCD, Ngày cấp, Mã thôn...*).
- **Từ dòng 2**, mỗi dòng là 1 khách hàng.
- Cột có số 0 đứng đầu (mã thôn `03`, mã tổ...) thì bôi đen cột → chuột phải → *Format Cells* → chọn **Text** rồi mới gõ.

## C. Dạy máy (làm 1 lần cho mỗi loại màn hình)

1. Trên Chrome (mở bằng `MoChrome.bat`), vào màn hình **nhập khách hàng mới**.
2. Nháy đúp **`BatDau.bat`** → gõ **`1`** → Enter → Enter.
3. Cửa sổ đen hiện thông tin khách ở **dòng 2** của Excel. Bạn chuyển sang Chrome và **nhập tay khách đó** như mọi ngày:
   - Bấm chuột vào ô đầu tiên, sau đó cố gắng **chỉ dùng bàn phím** (Enter, F5, mũi tên, F2...).
   - Gõ **đúng y như trong Excel** (để máy biết ô nào lấy cột nào).
   - Lưu xong (và bấm F2 về màn hình nhập mới nếu bạn hay làm vậy).
4. Quay lại cửa sổ đen, nhấn **Enter**. Máy in ra các bước nó đã học, ví dụ:
   ```
   BAM #txtHoTen
   GO {Họ và tên}
   PHIM Enter
   GO {Số CCCD}
   PHIM Enter
   PHIM F5
   GO {Mã thôn}
   PHIM Enter
   ...
   ```
   Thấy chữ trong ngoặc `{ }` là tên cột Excel → máy đã hiểu đúng ô đó lấy cột nào.

Nhập tay sai thì cứ làm lại bước C, máy ghi đè kịch bản mới (bản cũ giữ trong `kich_ban_cu.txt`).

## D. Cho máy nhập

1. Mở màn hình nhập khách hàng mới trên Chrome.
2. Nháy đúp **`BatDau.bat`**:
   - Gõ **`2`** → máy nhập **1 khách tiếp theo** để bạn xem thử.
   - Gõ **`3`** → máy nhập **tất cả khách còn lại**.
3. Máy hỏi *"Bắt đầu từ dòng ... ?"* → nhấn Enter (máy tự nhớ đã nhập tới đâu). Nhấn Enter lần nữa để bắt đầu.
4. Sau mỗi khách, máy **dừng lại chờ bạn nhìn màn hình**: đúng thì nhấn **Enter** để làm khách tiếp; gõ **`b`** để bỏ qua; gõ **`q`** để dừng.

- Đang chạy muốn dừng gấp: bấm vào cửa sổ đen, nhấn **Ctrl+C**.
- Gặp trục trặc (không thấy ô, danh mục không bật lên...) máy **tự dừng**, không gõ bừa. Lần sau chạy, máy tự hỏi bắt đầu lại từ đúng dòng đó.
- Trong lúc máy chạy, **đừng bấm chuột sang tab khác**.
- Kết quả từng khách ghi trong `nhat_ky.csv` (mở bằng Excel).

## E. Khi đã quen (không bắt buộc)

- Chạy ổn rồi, muốn máy **không dừng hỏi sau mỗi khách**: chọn **`4`** trong menu (mở kịch bản bằng Notepad), xoá dòng cuối `DUNG ...`, bấm Ctrl+S lưu lại.
- Mạng chậm, máy gõ nhanh quá làm sót: trong kịch bản sửa `NGHI 0.3` thành `NGHI 0.8`.
- Các lệnh có thể viết trong kịch bản: xem file mẫu `kich_ban_mau.txt` (có giải thích từng lệnh).
- F5 do công cụ bấm **không** làm tải lại trang; Unikey cũng không làm sai dấu.

> Máy cơ quan có quy định về cài phần mềm/công cụ tự động thì hỏi bộ phận tin học trước. Việc kiểm tra và lưu hồ sơ vẫn là trách nhiệm của người nhập.

---

# Kiểm tra vay chồng chéo khi nhận Giấy ủy quyền (Mẫu 01/UQ)

Mỗi hộ chỉ 1 người đứng tên vay. Khi hộ nộp giấy ủy quyền: **người được ủy quyền (mục II) là người vay**, các thành viên còn lại (mục I) phải tra xem đã có CIF/đang vay chưa. Phần tra này máy làm thay được, và **chỉ tra, không sửa gì trên hệ thống**.

## 1. Lấy danh sách từ giấy vào `tra_cuu.xlsx`

Dòng 1 là tên cột, nên có: `Họ và tên | Ngày sinh | Số CCCD | Vai trò` (thêm cột `Hộ` để ghi tên chủ hộ nếu tra nhiều hộ một lúc). Mỗi thành viên 1 dòng.

- **Gõ tay**: chỉ cần gõ số CCCD (và năm sinh), tên có thể để trống – kết quả tra sẽ hiện tên trên hệ thống để đối chiếu.
- **Dùng AI đọc ảnh** (chỉ khi cơ quan cho phép gửi ảnh giấy tờ khách hàng ra ngoài): mở file `Doc_giay_uy_quyen_bang_AI.txt`, làm theo hướng dẫn trong đó, copy bảng kết quả dán vào Excel.
- Máy tự soát lỗi số CCCD: **thiếu/thừa chữ số**, hoặc **số thứ 4–6 không khớp năm sinh** (ví dụ sinh 1990 thì CCCD dạng `xxx090...` hoặc `xxx190...`). Dòng sai sẽ tô vàng trong kết quả.

## 2. Dạy máy tra cứu (1 lần)

1. Mở Chrome bằng `MoChrome.bat`, vào **màn hình tra cứu khách hàng/CIF**.
2. Nháy đúp `BatDau.bat` → chọn **`4`** → Enter → Enter.
3. Tra tay **người ở dòng 2** của `tra_cuu.xlsx` như mọi ngày: bấm chuột vào ô tra cứu, gõ số CCCD, bấm phím tìm (Enter/F5...).
4. Khi kết quả hiện ra: **giữ phím Ctrl và bấm chuột** vào từng chỗ cần lấy (số CIF, tên, dư nợ, trạng thái...). Chỗ được chọn có viền đỏ.
5. Bấm phím để màn hình sẵn sàng tra người tiếp (F2/Esc... nếu cần), quay lại cửa sổ đen nhấn **Enter**.

## 3. Tra cả danh sách

`BatDau.bat` → chọn **`5`** → Enter → Enter. Máy tra lần lượt từng người, xong tự mở file **`ket_qua_tra_cuu.xlsx`**: mỗi người 1 dòng, thêm các cột `KQ1, KQ2...` là chữ máy đọc được ở các chỗ bạn đã Ctrl+bấm, và cột `Canh_bao`.

- Thành viên (không phải người vay) mà **có CIF đang có dư nợ** → hộ đang vay chồng chéo, xử lý theo quy định.
- Dòng **tô vàng** phải xem lại tay: số CCCD nghi sai, hoặc kết quả giống hệt người trước (máy đọc lúc màn hình chưa kịp đổi – khi đó trong kịch bản tra cứu tăng `CHO 1` thành `CHO 2`).
- Người không có CIF máy sẽ ghi đúng chữ phần mềm hiện ra (vd *Không tìm thấy*) hoặc `(không thấy)`.

Kết quả tra cứu chỉ để hỗ trợ; quyết định cho vay vẫn theo kiểm tra của cán bộ.

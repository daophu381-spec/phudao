# Kiểm tra văn bản theo Nghị định 30/2020/NĐ-CP

Chương trình Python đọc file Word (`.docx`) và **chỉ ra** (không tự sửa) những lỗi mà sếp hay bắt:
chính tả, viết hoa sai, dấu cách/dấu câu, thể thức Nghị định 30, và **chữ một mình một dòng**.
Kết quả là bảng lỗi nhóm theo loại, mỗi dòng có vị trí, chỗ lỗi, lý do và gợi ý sửa.

## Cài đặt

```bash
pip install -r requirements.txt
```

Để bắt lỗi **rơi chữ / rơi dòng / rơi trang** cần dàn trang thử, nên máy cần **một** trong hai:

- **LibreOffice** (miễn phí, <https://www.libreoffice.org>), hoặc
- **Microsoft Word** + `pip install docx2pdf`.

Không có cả hai thì chương trình vẫn chạy, chỉ bỏ qua nhóm "Rơi chữ, rơi dòng, rơi trang" và báo rõ trong kết quả.

## Cách dùng

```bash
python -m kiemtra_vanban bao_cao.docx                     # in kết quả ra màn hình
python -m kiemtra_vanban bao_cao.docx --html kq.html      # thêm báo cáo HTML (mở bằng trình duyệt)
python -m kiemtra_vanban D:\HoSo                          # kiểm tra mọi file .docx trong thư mục
python -m kiemtra_vanban bao_cao.docx --chi-loi           # chỉ hiện mức LỖI
python -m kiemtra_vanban bao_cao.docx --chu-mo-coi 2      # dòng cuối ≤ 2 chữ ngắn cũng coi là mồ côi
```

Trên Windows: **kéo thả file `.docx` lên `kiem_tra.bat`**, báo cáo tự mở trong trình duyệt.

Các tùy chọn khác: `--md`, `--json`, `--bo-bo-cuc`, `--bo-chinh-ta`, `--kieu-dau moi|cu|bo-qua`, `--luu-pdf` (xem bản dàn trang thử). Chạy `python -m kiemtra_vanban -h` để xem đủ.
Mã thoát: `0` không có LỖI, `1` có LỖI (tiện dùng trong script).

## Chương trình bắt những lỗi nào

| Nhóm | Ví dụ |
|---|---|
| **Chính tả** | `nghành`→`ngành`, `ghái`, `ngừơi`, `nhât` (âm cuối c/ch/p/t sai dấu), `ko`, lặp từ `và và`; từ sai thường gặp: `hàng tháng`→`hằng tháng`, `xử dụng`, `sử lý`, `sát nhập`→`sáp nhập`, `bổ xung`, `lãi xuất`, `gởi`...; lẫn `hoà/hòa`, `kí/ký` |
| **Viết hoa** | `Hộ Nghèo`, `Tỉnh Hưng Yên`, `điều 5`, `nghị định số 30/...`, `Ngân hàng Chính sách Xã hội`, `Ủy Ban Nhân Dân`, đầu câu viết thường, `VănAn` |
| **Dấu câu, dấu cách** | 2 dấu cách, cách đầu/cuối dòng, cách trước `, . ; :`, thiếu cách sau dấu câu, `( abc )`, `NĐ - CP`, `98,835,000 đồng`, `0.66%` |
| **Phông, cỡ chữ, trình bày** | không phải Times New Roman, cỡ chữ lệch, chữ không đen, lề, khổ giấy, căn đều, lùi đầu dòng, giãn dòng, khoảng cách đoạn |
| **Thể thức NĐ30** | Quốc hiệu, Tiêu ngữ, `Số: 05/…`, ngày tháng (`ngày 05 tháng 3`), địa danh in nghiêng, tên loại văn bản, trích yếu, `Kính gửi:`, `Nơi nhận:` (gạch đầu dòng `- `, `;` và `.`), `KT.`/`TM.`, `- Lưu: VT` |
| **Rơi chữ, rơi dòng, rơi trang** | **chữ một mình một dòng** (cuối đoạn và cuối tiêu đề/trích yếu), `ngày ⏎ 05`, `Điều ⏎ 5`, `98.835.000 ⏎ đồng`, tiêu đề mục cuối trang, trang cuối chỉ vài dòng, chữ ký khác trang với ngày tháng, hàng chữ ký chưa khóa `cantSplit` |

Mức độ: **LỖI** (chắc chắn sai), **CẢNH BÁO** (nhiều khả năng sai, cần người xác nhận: tên riêng, tên đơn vị), **GỢI Ý** (lệch khuyến nghị NĐ30).
Lề, cỡ chữ, giãn dòng chỉ ở mức cảnh báo/gợi ý, vì **biểu mẫu cấp trên ban hành thì giữ nguyên bố cục**.

## Tùy chỉnh cho cơ quan của bạn

- `tu_dien_rieng.txt`: từ/tên riêng **không** bị coi là sai (tên người, tên công ty nước ngoài...).
- `loi_chinh_ta_rieng.txt`: thêm các lỗi sếp hay bắt, dạng `sai => đúng`, chương trình sẽ bắt từ đó trở đi.

Hai file này đặt cạnh chương trình hoặc trong thư mục đang chạy là tự được đọc.

## Giới hạn (nên biết)

- **Chính tả dựa trên cấu tạo âm tiết, không có từ điển.** Bắt tốt lỗi gõ thừa/thiếu/đảo chữ và sai dấu ở âm cuối, nhưng
  **không bắt được lỗi mà âm tiết vẫn hợp lệ**, ví dụ thiếu dấu thanh (`tin dụng` thay vì `tín dụng`). Đo thử trên văn bản mẫu:
  gõ thừa chữ 100%, đảo chữ 92%, thiếu chữ giữa 58%, **thiếu dấu thanh chỉ khoảng 30%**. Lỗi loại này hãy bổ sung vào `loi_chinh_ta_rieng.txt`.
  (Đã thử so cụm từ với bộ từ điển mở, bắt được 98% nhưng báo nhầm khoảng 3 chỗ/dòng văn bản đúng nên không dùng.)
- Chỗ ngắt dòng do máy dàn trang quyết định. LibreOffice dùng phông Liberation Serif (cùng số đo với Times New Roman) nên
  rất sát Word, nhưng vẫn có thể lệch nhẹ. Nên **xem lại bằng mắt ở Word** những chỗ được báo rơi chữ/rơi trang.
- Không kiểm tra nội dung, văn phong, đúng sai số liệu; không tự sửa file.
- Tên riêng nước ngoài viết hoa chữ đầu bị báo mức CẢNH BÁO; thêm vào `tu_dien_rieng.txt` để bỏ qua.

## Kiểm thử

```bash
python tests/tao_mau.py            # tạo mau_dung.docx và mau_loi.docx trong tests/mau/ để xem thử
python -m unittest discover -s tests -v
```

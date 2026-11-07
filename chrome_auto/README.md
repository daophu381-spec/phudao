# Tự động hoá Chrome bằng Python

Điều khiển **chính Chrome của bạn** (giữ nguyên đăng nhập), làm lần lượt các bước ghi trong file JSON.

## Cài (1 lần)
```
pip install -r requirements.txt
python chrome_auto.py mo-chrome     # mở Chrome riêng, đăng nhập các trang cần dùng
```

## Dùng hằng ngày
```
python chrome_auto.py chay tasks/vi_du.json --thu   # xem thử các bước, không bấm
python chrome_auto.py chay tasks/vi_du.json         # chạy thật
```
Lỗi ở bước nào thì dừng, chụp ảnh `ket_qua/loi_*.png` và ghi log `ket_qua/log_*.txt`.

## Các hành động (`"lam"`)
| lam | tham số | việc |
|---|---|---|
| `mo_trang` | `url` | mở địa chỉ |
| `bam` | `chon` | bấm phần tử |
| `nhap` | `chon`, `noi_dung` | điền ô nhập |
| `phim` | `phim` | nhấn phím (`Enter`, `Control+S`…) |
| `chon_muc` | `chon`, `gia_tri` | chọn trong danh sách thả xuống |
| `cho` / `cho_xuat_hien` | `giay` / `chon` | chờ |
| `chup_anh` | `file` | chụp cả trang |
| `tai_ve` | `chon`, `file` | bấm nút tải và lưu file |
| `doc_chu` | `chon` | in chữ của phần tử ra log |
| `xac_nhan` | `hoi` | dừng hỏi y/N (nên đặt trước bước **duyệt/gửi**) |

`chon` là CSS (`#btn-ok`) hoặc `text=Phê duyệt`. Bước nào cho phép lỗi thì thêm `"bo_qua_loi": true`.
Biến dùng được: `{ngay}` `{ngay_vn}` `{gio}` và biến tự đặt trong `"bien"`.

Tìm `chon` nhanh: chạy `playwright codegen` rồi thao tác thử, nó sinh ra selector cho bạn.

## Chạy tự động mỗi ngày
- Windows: *Task Scheduler* → chạy `python chrome_auto.py chay tasks/<file>.json`.
- Linux/Mac: `cron`.
(Chrome mở bằng `mo-chrome` phải đang chạy.)

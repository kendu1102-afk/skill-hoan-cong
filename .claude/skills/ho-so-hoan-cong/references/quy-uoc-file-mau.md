# Quy ước file mẫu

File mẫu là một workbook Excel (.xlsx). Xem ví dụ đầy đủ ở `assets/file-mau-minh-hoa.xlsx`.

## Các loại sheet

| Sheet | Vai trò |
|---|---|
| `TRINH_TU` | Trình tự nghiệm thu cho từng loại công tác |
| `THONG_TIN` | Thông tin chung của công trình |
| Tên bắt đầu bằng `_` hoặc `HD` | Ghi chú/hướng dẫn - skill bỏ qua |
| Mọi sheet còn lại | Một biểu mẫu (biên bản, phiếu kiểm tra, biên bản lấy mẫu...) |

## Sheet TRINH_TU

Dòng đầu là tiêu đề (nhận theo chữ, không cần đúng thứ tự cột):

| Cột | Ý nghĩa |
|---|---|
| Mã công tác | Mã ngắn tự đặt: `DAP`, `CPDD`, `BTN`... |
| Loại công tác | Tên loại công tác |
| Từ khóa | Cụm từ xuất hiện trong tên công tác dự toán, cách nhau bằng `;`. Không phân biệt dấu/hoa thường. Cụm dài nhất khớp sẽ thắng |
| Mã hiệu định mức | Tiền tố mã hiệu (`AD.11;AD.12`) - ưu tiên hơn từ khóa |
| Bước | Số thứ tự bước - quyết định thứ tự ngày khi soát hồ sơ |
| Nội dung bước | Mô tả bước |
| Biểu mẫu | **Đúng tên sheet** biểu mẫu dùng cho bước này |
| Tiêu chuẩn áp dụng | Được điền vào trường `{{tieu_chuan}}` |
| Tần suất / ghi chú | Tần suất lấy mẫu, cách chia đợt... |

Một loại công tác chiếm nhiều dòng (mỗi bước một dòng); từ dòng thứ hai để trống 4 cột đầu.

## Sheet THONG_TIN

Cột A tên trường (chữ thường không dấu, gạch dưới), cột B giá trị mặc định, cột C mô tả. Mọi
trường ở đây dùng được trong tất cả biểu mẫu.

## Ô điền trong biểu mẫu

- `{{ten_truong}}` - trường đơn. Có thể đứng lẫn trong câu: `Lý trình: {{ly_trinh}}`.
  Ô chỉ chứa đúng một trường sẽ giữ kiểu số, tiện cho công thức.
- `{{ten_bang.ten_cot}}` - dòng bảng lặp. Chỉ cần làm **một dòng mẫu**; skill tự chèn thêm dòng
  theo số liệu, giữ nguyên kẻ khung, ô gộp ngang và chiều cao dòng. `{{ten_bang.stt}}` tự đánh số.
- Thứ tự tìm giá trị: trường của biên bản → thông tin chung. Không có → để trống + tô vàng.

## Tên trường chuẩn (dùng lại để các biểu mẫu chia sẻ dữ liệu)

Tự điền từ dự toán/TRINH_TU: `ten_cong_viec`, `hang_muc`, `don_vi`, `khoi_luong` (thiết kế/dự
toán), `ma_hieu`, `loai_cong_tac`, `tieu_chuan`, `noi_dung_buoc`, `ghi_chu_buoc`.

Điền từ bản vẽ: `ly_trinh`, `ban_ve_so`, và các cột thiết kế trong bảng (`kt.vi_tri`,
`kt.noi_dung`, `kt.thiet_ke`, `kt.cho_phep`).

Cần số liệu hiện trường: `so_bien_ban`, `ngay`, `ngay_nghiem_thu`, `ngay_lay_mau`,
`ngay_kiem_tra` (mọi trường bắt đầu bằng `ngay` đều được soát trình tự, định dạng dd/mm/yyyy),
`gio_bat_dau`, `gio_ket_thuc`, `khoi_luong_nghiem_thu`, `ket_luan`, `danh_gia_chat_luong`,
các cột `kt.thuc_te`, `kt.sai_so`, `kt.danh_gia`.

Có thể đặt thêm bất kỳ tên trường nào - skill đọc tên từ file mẫu chứ không có danh sách cứng.

## Thêm biểu mẫu mới

1. Thêm một sheet, trình bày như bản in thật, gắn `{{trường}}` vào chỗ cần điền.
2. Thêm dòng vào `TRINH_TU` ở công tác và bước tương ứng, cột Biểu mẫu ghi tên sheet.
3. Chạy `quet_file_mau.py` - không còn cảnh báo là xong.

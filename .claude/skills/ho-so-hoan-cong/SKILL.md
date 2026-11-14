---
name: ho-so-hoan-cong
description: Lập hồ sơ hoàn công / hồ sơ quản lý chất lượng (QLCL) công trình cầu đường và bảo dưỡng thường xuyên đường bộ từ file dự toán + bản vẽ, dựa trên file mẫu của người dùng (trình tự nghiệm thu, biên bản nghiệm thu, phiếu kiểm tra, biên bản lấy mẫu, phiếu yêu cầu nghiệm thu). Dùng skill này bất cứ khi nào người dùng đưa dự toán, bản vẽ hoặc trình tự nghiệm thu và muốn làm hồ sơ hoàn công, hồ sơ nghiệm thu, hồ sơ QLCL, biên bản nghiệm thu công việc/vật liệu, biên bản lấy mẫu, phiếu kiểm tra, danh mục hồ sơ; khi họ muốn điền số liệu từ bản vẽ vào biên bản; và cả khi họ gửi file mẫu mới hoặc thêm biểu mẫu/bảng biểu để "cập nhật skill" - kể cả khi họ không nhắc tới chữ "skill".
---

# Lập hồ sơ hoàn công từ dự toán + bản vẽ

Người dùng là kỹ sư QLDA/QLCL cầu đường. Sản phẩm cuối là một file Excel hồ sơ hoàn công: mỗi
biên bản/phiếu là một sheet đúng biểu mẫu của họ, đủ trình tự cho từng công tác trong dự toán,
số liệu thiết kế điền từ bản vẽ, kèm sheet danh mục. Hồ sơ này sẽ bị thanh tra, kiểm toán soi,
nên **đúng và truy được nguồn** quan trọng hơn là điền cho kín.

## Nguyên tắc số liệu (đọc trước khi làm)

Hồ sơ hoàn công ghi lại việc **đã thực sự xảy ra** ngoài hiện trường. Vì vậy:

- **Được điền từ dự toán / bản vẽ:** tên công việc, hạng mục, lý trình, kích thước và cao độ
  *thiết kế*, khối lượng thiết kế, số hiệu bản vẽ, tiêu chuẩn áp dụng (lấy từ TRINH_TU), sai số
  cho phép (nếu file mẫu hoặc người dùng đã nêu).
- **Chỉ điền khi người dùng cung cấp** (nhật ký thi công, sổ đo, phiếu kết quả thí nghiệm, tin
  báo hiện trường): số đo thực tế, kết quả thí nghiệm, ngày giờ, số biên bản, tên người ký,
  khối lượng nghiệm thu thực tế, kết luận đạt/không đạt.
- Cái gì chưa có thì **để `null`** - script sẽ để trống và tô vàng. Không tự nghĩ ra số đo "cho
  đẹp": một số liệu bịa lọt vào hồ sơ pháp lý là rủi ro lớn cho người dùng khi giải trình, còn
  một ô vàng thì họ nhìn là biết phải bổ sung.
- Thấy dự toán và bản vẽ lệch nhau (khối lượng, kích thước, lý trình) thì không tự chọn bên
  nào - ghi lại và báo người dùng; phát hiện sai khác là một phần công việc của họ.

## Đầu vào

| Thứ | Bắt buộc | Ghi chú |
|---|---|---|
| File mẫu (.xlsx) | Có | `assets/file-mau.xlsx` nếu đã lưu; nếu chưa có thì dùng `assets/file-mau-minh-hoa.xlsx` và **nói rõ** đang dùng mẫu minh họa |
| Dự toán (.xlsx/.xls/.xlsm) | Có | |
| Bản vẽ (PDF; ảnh) | Nên có | DWG không đọc trực tiếp được - đề nghị xuất PDF |
| Trình tự nghiệm thu | Tùy | Nếu người dùng đưa riêng thì ưu tiên hơn sheet TRINH_TU |
| Số liệu hiện trường | Tùy | Có thì điền, không có thì để trống |

## Quy trình

Các script nằm trong `scripts/`, chạy bằng `python3`. Làm việc trong một thư mục tạm; file kết
quả đặt ở nơi người dùng mở được.

### 1. Quét file mẫu

```bash
python3 scripts/quet_file_mau.py [file_mau.xlsx] -o mau.json
```

Cho biết các loại công tác + bước nghiệm thu (sheet `TRINH_TU`), các biểu mẫu và trường
`{{...}}` cần điền của từng biểu mẫu, các trường thông tin chung (sheet `THONG_TIN`). Đọc kỹ
phần `canh_bao`. Nếu người dùng gửi file mẫu mới hoặc biểu mẫu chưa gắn trường → làm theo mục
"Cập nhật file mẫu" bên dưới trước khi đi tiếp.

### 2. Đọc dự toán

```bash
python3 scripts/doc_du_toan.py du_toan.xlsx -o du_toan.json
```

Script tự dò bảng trên mọi sheet. Nếu nhận ra nhiều sheet, chọn sheet dự toán chi tiết (có mã
hiệu và diễn giải khối lượng), không phải bảng tổng hợp vật tư hay phân tích đơn giá. Nếu báo
khối lượng là công thức chưa tính, nhờ người dùng mở file bằng Excel và lưu lại.

### 3. Lập khung kế hoạch

```bash
python3 scripts/lap_khung_ke_hoach.py du_toan.json mau.json -o ke_hoach.json
```

Mỗi công tác dự toán khớp được với TRINH_TU (theo mã hiệu định mức, rồi từ khóa) sinh ra đủ các
bước. Xử lý danh sách **chưa khớp**: công tác phụ không nghiệm thu riêng (vận chuyển, tưới
ẩm...) thì chuyển vào `cong_tac_khong_nghiem_thu_rieng` trong ke_hoach.json kèm lý do; công tác
cần nghiệm thu mà file mẫu chưa có trình tự thì báo người dùng bổ sung - đừng tự đặt ra trình tự
nghiệm thu thay họ.

### 4. Đọc bản vẽ và hoàn thiện ke_hoach.json

Đọc `references/doc-ban-ve.md` để biết cần lấy gì từ bản vẽ, và `references/ke-hoach-json.md`
cho cấu trúc file. Việc ở bước này:

- Điền `chung` (tên công trình, gói thầu, chủ đầu tư, địa điểm...) từ khung tên bản vẽ, thuyết
  minh, trang bìa dự toán.
- Chia công tác thành các đợt/đoạn nghiệm thu theo thực tế thi công (theo lớp đắp, theo đoạn lý
  trình, theo cấu kiện) - nhân bản bản ghi và ghi `dot`. Nếu người dùng chưa cho biết cách chia
  thì giữ một đợt cho cả khối lượng và hỏi lại ở phần báo cáo.
- Điền trường từ bản vẽ: lý trình, số bản vẽ, các dòng bảng kiểm tra (vị trí cọc, nội dung kiểm
  tra, giá trị thiết kế). Cột thực tế/sai số/đánh giá để trống trừ khi có số liệu hiện trường.
- Với dự án lớn (hàng trăm biên bản), viết một đoạn Python ngắn để sửa ke_hoach.json thay vì
  sửa tay từng bản ghi.

### 5. Sinh hồ sơ và soát

```bash
python3 scripts/tao_ho_so.py ke_hoach.json -o "Ho so hoan cong - <ten cong trinh>.xlsx"
python3 scripts/kiem_tra_ho_so.py ke_hoach.json "<file vừa tạo>" --du-toan du_toan.json
```

`tao_ho_so.py` tạo thêm `<tên>_con-thieu.json` liệt kê từng ô còn thiếu. `kiem_tra_ho_so.py`
bắt các lỗi kiểm toán hay hỏi: sai trình tự ngày giữa các bước, ngày ngoài thời gian thi công,
trùng số biên bản, công tác dự toán chưa có biên bản, khối lượng nghiệm thu cộng dồn lệch dự
toán. Sửa ke_hoach.json rồi chạy lại cho tới khi hết LỖI; CẢNH BÁO thì đưa vào báo cáo.

### 6. Báo cáo cho người dùng

Gửi file Excel và nói ngắn gọn, bằng tiếng Việt:

1. Đã lập bao nhiêu biên bản/phiếu cho bao nhiêu công tác.
2. Số liệu nào lấy từ bản vẽ (kèm số bản vẽ) - để họ kiểm lại được.
3. Còn thiếu gì cần số liệu hiện trường (gom theo loại, không liệt kê từng ô).
4. Sai khác dự toán - bản vẽ đã phát hiện.
5. Công tác chưa có trình tự trong file mẫu.

## Cập nhật file mẫu (tính năng mở rộng của skill)

Skill không mã hóa cứng biểu mẫu nào - mọi thứ đọc từ file mẫu, nên "thêm tính năng" nghĩa là
sửa file mẫu. Chi tiết quy ước ở `references/quy-uoc-file-mau.md`. Khi người dùng gửi file mẫu
mới hoặc thêm bảng biểu:

1. Chạy `quet_file_mau.py` trên file họ gửi.
2. Biểu mẫu nào chưa có `{{trường}}` (file xuất thẳng từ phần mềm QLCL thường như vậy): chạy
   `quet_file_mau.py <file> --chi-tiet "<sheet>"`, xác định ô nào là chỗ điền, rồi dùng openpyxl
   ghi `{{ten_truong}}` vào các ô đó trên **bản sao** của file. Dùng lại tên trường chuẩn trong
   `references/quy-uoc-file-mau.md` để các biểu mẫu chia sẻ được dữ liệu. Chỉ thêm chữ vào ô,
   không đổi định dạng.
3. Biểu mẫu mới chưa nằm trong TRINH_TU: hỏi người dùng nó thuộc bước nào của công tác nào, rồi
   thêm dòng.
4. Quét lại cho tới khi hết cảnh báo, cho người dùng xem danh sách trường đã gắn, rồi lưu thành
   `assets/file-mau.xlsx` trong thư mục skill (nếu skill nằm trong kho git thì commit) để các
   lần sau tự dùng.

## Giới hạn cần nói thật với người dùng

- Bản vẽ scan mờ hoặc chữ nhỏ: đọc được nhưng dễ nhầm số - nêu rõ số liệu nào đọc không chắc.
- Hình ảnh/logo và biểu đồ trong biểu mẫu không được sao sang sheet mới (giới hạn của openpyxl);
  công thức nằm bên dưới bảng nhân dòng không tự dịch vùng tham chiếu.
- File `.xls` được chuyển tự động nếu máy có LibreOffice; không thì nhờ người dùng lưu `.xlsx`.

# Đọc bản vẽ để điền hồ sơ

## Cách đọc

- PDF vẽ từ CAD (có lớp chữ): thử `pdfplumber` lấy chữ từng trang trước - nhanh và chính xác với
  bảng khối lượng, bảng tọa độ, bảng thống kê cốt thép. Sau đó mở trang dưới dạng ảnh (Read với
  `pages`) để hiểu hình vẽ và kích thước ghi trên hình.
- PDF scan / ảnh: đọc bằng hình. Bản vẽ A1/A3 thu nhỏ rất dễ nhầm số (3↔8, 1↔7, dấu phẩy) -
  cắt/phóng to vùng cần đọc (render 150-200 dpi theo vùng) và ghi nhận số nào không chắc.
- Bộ bản vẽ dài: đọc mục lục/danh mục bản vẽ trước, rồi chỉ mở các tờ liên quan tới công tác
  đang lập hồ sơ. Không cần đọc cả bộ trong một lượt.
- DWG không đọc được - đề nghị người dùng xuất PDF (hoặc DXF).

## Cần lấy gì

| Nguồn trong bản vẽ | Dùng cho |
|---|---|
| Khung tên, trang bìa, thuyết minh | `chung`: tên công trình, gói thầu, chủ đầu tư, tư vấn thiết kế, địa điểm, tiêu chuẩn thiết kế |
| Bình đồ, trắc dọc | Phạm vi lý trình từng hạng mục, tên cọc, cao độ thiết kế tim/vai |
| Trắc ngang điển hình + trắc ngang chi tiết | Bề rộng nền/mặt/lề, độ dốc ngang, mái taluy, chiều dày từng lớp kết cấu, vị trí rãnh |
| Bản vẽ kết cấu áo đường | Thứ tự và chiều dày lớp → thứ tự nghiệm thu các lớp, yêu cầu độ chặt (K95/K98), loại vật liệu |
| Bản vẽ cống, rãnh, tường chắn, cầu | Kích thước hình học cấu kiện, mác bê tông, loại và đường kính cốt thép, cao độ đáy/đỉnh |
| Bảng khối lượng trên bản vẽ | Đối chiếu với dự toán |
| Ghi chú chung | Yêu cầu vật liệu, tiêu chuẩn thi công nghiệm thu thiết kế chỉ định |

Ghi số hiệu bản vẽ cho từng số liệu lấy ra (`ban_ve_so`) - đây là thứ người dùng cần khi giải
trình nguồn gốc số liệu.

## Lập bảng kiểm tra từ bản vẽ

Với phiếu kiểm tra kích thước hình học: mỗi dòng là một (vị trí, nội dung kiểm tra, giá trị
thiết kế). Vị trí lấy theo các cọc/mặt cắt có trong trắc ngang hoặc theo mật độ kiểm tra mà
tiêu chuẩn ở TRINH_TU / người dùng yêu cầu. Nội dung kiểm tra thường là bề rộng, cao độ, chiều
dày, độ dốc ngang - nhưng lấy theo đúng các cột/mục trong biểu mẫu của người dùng, đừng thêm
hạng mục kiểm tra mà biểu mẫu và tiêu chuẩn của họ không yêu cầu. Sai số cho phép chỉ điền khi
file mẫu, TRINH_TU hoặc người dùng đã nêu; nếu tự tra theo tiêu chuẩn thì nói rõ nguồn và nhờ
họ xác nhận.

## Đối chiếu dự toán - bản vẽ

Trong lúc đọc, so nhanh những thứ dễ lệch và ghi lại để báo cáo (không tự sửa):

- Khối lượng dự toán so với bảng khối lượng trên bản vẽ hoặc tính lại từ kích thước
  (dài × rộng × dày), nhớ quy đổi đơn vị định mức (100m3, 100m2, tấn).
- Chiều dày lớp, mác bê tông, loại vật liệu trong tên công tác so với bản vẽ.
- Phạm vi lý trình trong diễn giải khối lượng so với bình đồ.
- Công tác có trên bản vẽ nhưng không có trong dự toán và ngược lại.

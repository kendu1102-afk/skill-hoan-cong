# Cấu trúc ke_hoach.json

`lap_khung_ke_hoach.py` tạo khung; Claude hoàn thiện; `tao_ho_so.py` đọc để sinh hồ sơ.

```json
{
  "file_mau": "đường dẫn file mẫu",
  "chung": { "ten_cong_trinh": "...", "ngay_khoi_cong": "01/03/2026", "ngay_hoan_thanh": null },
  "bien_ban": [
    {
      "mau": "PHIEU_KT",                 // tên sheet biểu mẫu
      "ten_sheet": "",                   // trống -> tự đặt "012_PHIEU_KT"; tối đa 31 ký tự
      "cong_tac": "3 Làm móng CPĐD ...", // nhóm các bước của cùng một công tác
      "dot": "Đoạn Km10+000-Km10+400",   // tùy chọn: đợt/đoạn/lớp khi một công tác nghiệm thu nhiều lần
      "ma_cong_tac": "CPDD",
      "buoc": "4",                       // thứ tự bước - dùng để soát trình tự ngày
      "dong_du_toan": 10,                // dòng trong sheet dự toán - dùng để soát khối lượng
      "truong": { "ly_trinh": "Km10+000 - Km10+400", "ngay_kiem_tra": null },
      "bang": {
        "kt": [ { "vi_tri": "Km10+000", "noi_dung": "Bề rộng móng (m)", "thiet_ke": 7.5, "thuc_te": null } ]
      },
      "_cot_bang": { "kt": ["stt", "vi_tri", "..."] }   // gợi ý tên cột, script bỏ qua
    }
  ],
  "cong_tac_chua_co_trinh_tu": [ ... ],
  "cong_tac_khong_nghiem_thu_rieng": [ { "ma_hieu": "...", "ten": "...", "ly_do": "..." } ]
}
```

Lưu ý:

- Giá trị `null` hoặc thiếu = chưa có số liệu → ô để trống, tô vàng, ghi vào `_con-thieu.json`.
- Giữ số là số (`7.5`, không phải `"7,5"`) để Excel tính được; ngày ghi chuỗi `dd/mm/yyyy`.
- Chia đợt: nhân bản toàn bộ các bước của công tác cho mỗi đợt, đặt cùng `dot`, và chia
  `khoi_luong_nghiem_thu` sao cho tổng các đợt bằng khối lượng thực hiện.
- Thứ tự phần tử trong `bien_ban` là thứ tự sheet trong hồ sơ và trong danh mục.
- Bước lấy mẫu / nghiệm thu vật liệu dùng chung cho nhiều công tác (cùng nguồn vật liệu) thì chỉ
  giữ một bộ, không lặp theo từng công tác - hỏi người dùng nếu không rõ.

#!/usr/bin/env python3
"""Sinh assets/file-mau-minh-hoa.xlsx - file mẫu MINH HỌA quy ước của skill.

Đây chỉ là ví dụ để người dùng thấy cách tổ chức file mẫu; biểu mẫu và trình tự thật do người
dùng cung cấp (lưu thành assets/file-mau.xlsx). Chạy lại script này khi cần tạo lại file minh họa.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

F = "Times New Roman"
VIEN = Border(*(Side(style="thin"),) * 4)
GIUA = Alignment(horizontal="center", vertical="center", wrap_text=True)
TRAI = Alignment(horizontal="left", vertical="center", wrap_text=True)
XAM = PatternFill("solid", fgColor="D9D9D9")


class Mau:
    """Trình dựng biểu mẫu khổ A4 dọc, 8 cột A..H."""

    def __init__(self, wb, ten, rong=(6, 16, 16, 12, 12, 12, 12, 14)):
        self.ws, self.r, self.n = wb.create_sheet(ten), 1, len(rong)
        for i, w in enumerate(rong):
            self.ws.column_dimensions["ABCDEFGHIJ"[i]].width = w
        ps = self.ws.page_setup
        ps.paperSize, ps.orientation, ps.fitToWidth, ps.fitToHeight = 9, "portrait", 1, 0
        self.ws.sheet_properties.pageSetUpPr.fitToPage = True
        self.ws.page_margins.left, self.ws.page_margins.right = 0.6, 0.4

    def dong(self, chu="", dam=False, giua=False, co=12, nghieng=False, cao=None, c1=1, c2=None):
        c2 = c2 or self.n
        o = self.ws.cell(self.r, c1, chu)
        o.font = Font(name=F, size=co, bold=dam, italic=nghieng)
        o.alignment = GIUA if giua else TRAI
        if c2 > c1:
            self.ws.merge_cells(start_row=self.r, start_column=c1, end_row=self.r, end_column=c2)
        self.ws.row_dimensions[self.r].height = cao or (18 if len(chu) < 95 else 34)
        self.r += 1

    def quoc_hieu(self, tieu_de, so="Số: {{so_bien_ban}}"):
        self.dong("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", dam=True, giua=True)
        self.dong("Độc lập - Tự do - Hạnh phúc", dam=True, giua=True)
        self.dong("{{dia_diem}}, ngày {{ngay}}", giua=True, nghieng=True)
        self.dong()
        self.dong(tieu_de, dam=True, giua=True, co=14, cao=24)
        self.dong(so, giua=True)
        self.dong()
        self.dong("Công trình: {{ten_cong_trinh}}", dam=True)
        self.dong("Gói thầu: {{goi_thau}}")
        self.dong("Hạng mục: {{hang_muc}}")
        self.dong("Địa điểm xây dựng: {{dia_diem}}")

    def thanh_phan(self, thi_nghiem=False):
        self.dong("Thành phần trực tiếp tham gia:", dam=True)
        self.dong("a) Đại diện Tư vấn giám sát: {{tu_van_giam_sat}}")
        self.dong("    Ông/Bà: {{dai_dien_tvgs}}  -  Chức vụ: {{chuc_vu_tvgs}}")
        self.dong("b) Đại diện Nhà thầu thi công: {{nha_thau}}")
        self.dong("    Ông/Bà: {{ky_thuat_nha_thau}}  -  Chức vụ: Kỹ thuật thi công trực tiếp")
        if thi_nghiem:
            self.dong("c) Đại diện đơn vị thí nghiệm: {{don_vi_thi_nghiem}}")

    def bang(self, tieu_de, o_mau, gop=()):
        """tieu_de/o_mau: danh sách theo cột; gop: các cặp (c1, c2) gộp ngang trên cả 2 dòng."""
        for dong, dam in ((tieu_de, True), (o_mau, False)):
            for c, chu in enumerate(dong, 1):
                o = self.ws.cell(self.r, c, chu)
                o.font, o.border, o.alignment = Font(name=F, size=11, bold=dam), VIEN, GIUA
                if dam:
                    o.fill = XAM
            for c1, c2 in gop:
                self.ws.merge_cells(start_row=self.r, start_column=c1, end_row=self.r, end_column=c2)
            self.ws.row_dimensions[self.r].height = 32 if dam else 30
            self.r += 1

    def chu_ky(self, ben=("TƯ VẤN GIÁM SÁT", "NHÀ THẦU THI CÔNG"), ten=("{{dai_dien_tvgs}}", "{{ky_thuat_nha_thau}}")):
        self.dong()
        nua = self.n // 2
        for chu, dam, cao in ((ben, True, 18), (("(Ký, ghi rõ họ tên)",) * 2, False, 18), (("", ""), False, 60), (ten, True, 18)):
            for i, (c1, c2) in enumerate(((1, nua), (nua + 1, self.n))):
                o = self.ws.cell(self.r, c1, chu[i])
                o.font, o.alignment = Font(name=F, size=12, bold=dam, italic=not dam), GIUA
                self.ws.merge_cells(start_row=self.r, start_column=c1, end_row=self.r, end_column=c2)
            self.ws.row_dimensions[self.r].height = cao
            self.r += 1
        self.ws.print_area = f"A1:{'ABCDEFGHIJ'[self.n - 1]}{self.r - 1}"


def main():
    wb = Workbook()
    hd = wb.active
    hd.title = "HD_SU_DUNG"
    hd.column_dimensions["A"].width = 120
    for chu in [
        "FILE MẪU MINH HỌA - chỉ để thể hiện quy ước, KHÔNG phải biểu mẫu chính thức của dự án.",
        "1. Sheet TRINH_TU: mỗi loại công tác + các bước nghiệm thu; cột 'Biểu mẫu' ghi đúng tên sheet biểu mẫu.",
        "2. Sheet THONG_TIN: thông tin chung của công trình (cột A tên trường, cột B giá trị, cột C mô tả).",
        "3. Mỗi sheet còn lại là một biểu mẫu. Ô cần điền ghi {{ten_truong}}; dòng bảng lặp ghi {{ten_bang.ten_cot}}.",
        "4. Thêm biểu mẫu mới = thêm 1 sheet + thêm dòng ở TRINH_TU. Sheet tên bắt đầu bằng '_' hoặc 'HD' được bỏ qua.",
        "5. Tiêu chuẩn ghi ở TRINH_TU do người lập file mẫu chịu trách nhiệm rà soát hiệu lực trước khi dùng.",
    ]:
        hd.append([chu])
        hd.cell(hd.max_row, 1).alignment = TRAI

    # ---------------- TRINH_TU ----------------
    tt = wb.create_sheet("TRINH_TU")
    td = ["Mã công tác", "Loại công tác", "Từ khóa nhận diện trong dự toán (cách nhau bằng ;)", "Mã hiệu định mức (tiền tố)",
          "Bước", "Nội dung bước", "Biểu mẫu (tên sheet)", "Tiêu chuẩn áp dụng", "Tần suất / ghi chú"]
    tt.append(td)
    yc, kt, nt = "PHIEU_YC_NT", "PHIEU_KT", "BB_NT_CONG_VIEC"
    lm, vl = "BB_LAY_MAU", "BB_NT_VAT_LIEU"
    du_lieu = [
        ("DAO", "Đào nền đường, đào khuôn", "đào nền đường; đào khuôn đường; đào đất; đào rãnh; đào móng", "AB.2;AB.3", [
            ("1", "Phiếu yêu cầu nghiệm thu", yc, "", ""),
            ("2", "Kiểm tra kích thước hình học, cao độ nền đào", kt, "TCVN 9436:2012", "Theo trắc ngang thiết kế"),
            ("3", "Nghiệm thu công việc xây dựng", nt, "TCVN 9436:2012; TCVN 4447:2012", "")]),
        ("DAP", "Đắp nền đường", "đắp nền đường; đắp đất; đắp cát; đắp lề", "AB.6", [
            ("1", "Lấy mẫu vật liệu đắp", lm, "TCVN 9436:2012", "Theo nguồn/mỏ vật liệu và khối lượng"),
            ("2", "Nghiệm thu vật liệu trước khi sử dụng", vl, "TCVN 9436:2012", ""),
            ("3", "Phiếu yêu cầu nghiệm thu", yc, "", "Theo từng lớp đắp"),
            ("4", "Kiểm tra kích thước hình học, cao độ, độ chặt", kt, "TCVN 9436:2012", "Theo từng lớp đắp"),
            ("5", "Nghiệm thu công việc xây dựng", nt, "TCVN 9436:2012", "Theo từng lớp đắp")]),
        ("CPDD", "Móng cấp phối đá dăm", "cấp phối đá dăm; cpđd; móng đá dăm", "AD.11", [
            ("1", "Lấy mẫu cấp phối đá dăm", lm, "TCVN 8859:2011", ""),
            ("2", "Nghiệm thu vật liệu trước khi sử dụng", vl, "TCVN 8859:2011", ""),
            ("3", "Phiếu yêu cầu nghiệm thu", yc, "", ""),
            ("4", "Kiểm tra kích thước hình học, cao độ, độ chặt", kt, "TCVN 8859:2011", ""),
            ("5", "Nghiệm thu công việc xây dựng", nt, "TCVN 8859:2011", "")]),
        ("BTN", "Mặt đường bê tông nhựa", "bê tông nhựa; btn; thảm nhựa; bù vênh", "AD.23", [
            ("1", "Lấy mẫu vật liệu, hỗn hợp BTN", lm, "TCVN 13567-1:2022", ""),
            ("2", "Nghiệm thu vật liệu trước khi sử dụng", vl, "TCVN 13567-1:2022", ""),
            ("3", "Phiếu yêu cầu nghiệm thu", yc, "", ""),
            ("4", "Kiểm tra kích thước hình học, cao độ, độ bằng phẳng", kt, "TCVN 13567-1:2022", ""),
            ("5", "Nghiệm thu công việc xây dựng", nt, "TCVN 13567-1:2022", "")]),
        ("BT", "Bê tông, bê tông cốt thép", "bê tông móng; bê tông thân; bê tông rãnh; bê tông tấm đan; đổ bê tông; bê tông m", "AF.1;AF.2;AF.3;AG.1", [
            ("1", "Lấy mẫu vật liệu (xi măng, cát, đá, nước)", lm, "TCVN 4453:1995", ""),
            ("2", "Nghiệm thu vật liệu trước khi sử dụng", vl, "TCVN 4453:1995", ""),
            ("3", "Phiếu yêu cầu nghiệm thu", yc, "", ""),
            ("4", "Kiểm tra kích thước hình học kết cấu", kt, "TCVN 4453:1995", ""),
            ("5", "Nghiệm thu công việc xây dựng", nt, "TCVN 4453:1995", "")]),
    ]
    for ma, loai, tu_khoa, mh, cac_buoc in du_lieu:
        for i, (b, nd, bm, tc, gc) in enumerate(cac_buoc):
            tt.append([ma, loai, tu_khoa, mh, b, nd, bm, tc, gc] if i == 0 else ["", "", "", "", b, nd, bm, tc, gc])
    for i, w in enumerate((12, 28, 46, 18, 7, 44, 22, 30, 32)):
        tt.column_dimensions["ABCDEFGHI"[i]].width = w
    for row in tt.iter_rows():
        for o in row:
            o.border, o.alignment, o.font = VIEN, TRAI, Font(name=F, size=11, bold=o.row == 1)
            if o.row == 1:
                o.fill = XAM
    tt.freeze_panes = "A2"

    # ---------------- THONG_TIN ----------------
    ti = wb.create_sheet("THONG_TIN")
    ti.append(["Tên trường", "Giá trị", "Mô tả"])
    for dong in [
        ("ten_cong_trinh", None, "Tên công trình / dự án"),
        ("goi_thau", None, "Tên gói thầu"),
        ("dia_diem", None, "Địa điểm xây dựng"),
        ("chu_dau_tu", None, "Tên chủ đầu tư"),
        ("tu_van_giam_sat", None, "Tên đơn vị tư vấn giám sát"),
        ("nha_thau", None, "Tên nhà thầu thi công"),
        ("don_vi_thi_nghiem", None, "Phòng thí nghiệm (LAS-XD ...)"),
        ("dai_dien_tvgs", None, "Họ tên giám sát viên ký biên bản"),
        ("chuc_vu_tvgs", "Giám sát viên", "Chức vụ"),
        ("dai_dien_nha_thau", None, "Họ tên chỉ huy trưởng"),
        ("ky_thuat_nha_thau", None, "Họ tên kỹ thuật thi công trực tiếp"),
        ("ho_so_thiet_ke", None, "Hồ sơ thiết kế BVTC được duyệt (số, ngày quyết định)"),
        ("ngay_khoi_cong", None, "dd/mm/yyyy"),
        ("ngay_hoan_thanh", None, "dd/mm/yyyy"),
    ]:
        ti.append(dong)
    for i, w in enumerate((24, 60, 50)):
        ti.column_dimensions["ABC"[i]].width = w
    for row in ti.iter_rows():
        for o in row:
            o.border, o.alignment, o.font = VIEN, TRAI, Font(name=F, size=11, bold=o.row == 1)

    # ---------------- BB_LAY_MAU ----------------
    m = Mau(wb, lm)
    m.quoc_hieu("BIÊN BẢN LẤY MẪU VẬT LIỆU TẠI HIỆN TRƯỜNG")
    m.dong("Phục vụ công tác: {{ten_cong_viec}}")
    m.thanh_phan(thi_nghiem=True)
    m.dong("Thời gian lấy mẫu: {{ngay_lay_mau}}  -  Địa điểm lấy mẫu: {{dia_diem_lay_mau}}")
    m.dong("Các bên cùng tiến hành lấy mẫu như sau:", dam=True)
    m.bang(["TT", "Tên vật liệu", "Nguồn gốc", "Vị trí lấy mẫu", "", "Số lượng mẫu", "Chỉ tiêu thí nghiệm", "Tiêu chuẩn thử"],
           ["{{mau.stt}}", "{{mau.ten_vat_lieu}}", "{{mau.nguon_goc}}", "{{mau.vi_tri_lay}}", "", "{{mau.so_luong_mau}}",
            "{{mau.chi_tieu}}", "{{mau.tieu_chuan_thu}}"], gop=[(4, 5)])
    m.dong("Mẫu được niêm phong, có chữ ký các bên và chuyển về phòng thí nghiệm: {{don_vi_thi_nghiem}}.")
    m.chu_ky()

    # ---------------- BB_NT_VAT_LIEU ----------------
    m = Mau(wb, vl)
    m.quoc_hieu("BIÊN BẢN NGHIỆM THU VẬT LIỆU, SẢN PHẨM\nTRƯỚC KHI SỬ DỤNG VÀO CÔNG TRÌNH")
    m.ws.row_dimensions[5].height = 40
    m.dong("1. Đối tượng nghiệm thu: vật liệu sử dụng cho công tác {{ten_cong_viec}}", dam=True)
    m.dong("2. Thành phần trực tiếp nghiệm thu:", dam=True)
    m.thanh_phan()
    m.dong("3. Thời gian nghiệm thu: {{ngay_nghiem_thu}} tại {{dia_diem}}", dam=True)
    m.dong("4. Căn cứ nghiệm thu: {{ho_so_thiet_ke}}; tiêu chuẩn {{tieu_chuan}}; kết quả thí nghiệm kèm theo.", dam=True)
    m.dong("5. Đánh giá vật liệu:", dam=True)
    m.bang(["TT", "Tên vật liệu", "Nguồn gốc", "Đơn vị", "Khối lượng", "Phiếu kết quả thí nghiệm số", "", "Kết quả"],
           ["{{vat_lieu.stt}}", "{{vat_lieu.ten_vat_lieu}}", "{{vat_lieu.nguon_goc}}", "{{vat_lieu.don_vi}}", "{{vat_lieu.khoi_luong}}",
            "{{vat_lieu.phieu_thi_nghiem}}", "", "{{vat_lieu.ket_qua}}"], gop=[(6, 7)])
    m.dong("6. Kết luận: {{ket_luan}}", dam=True)
    m.chu_ky()

    # ---------------- PHIEU_YC_NT ----------------
    m = Mau(wb, yc)
    m.quoc_hieu("PHIẾU YÊU CẦU NGHIỆM THU")
    m.dong("Kính gửi: {{tu_van_giam_sat}}", dam=True)
    m.dong("Nhà thầu {{nha_thau}} đã thi công xong và tự kiểm tra, đánh giá đạt yêu cầu đối với công việc:")
    m.dong("- Tên công việc: {{ten_cong_viec}}")
    m.dong("- Vị trí / lý trình: {{ly_trinh}}")
    m.dong("- Khối lượng đề nghị nghiệm thu: {{khoi_luong_nghiem_thu}} {{don_vi}}")
    m.dong("Đề nghị Tư vấn giám sát tổ chức nghiệm thu vào: {{thoi_gian_nghiem_thu}} tại hiện trường.")
    m.chu_ky(ben=("TƯ VẤN GIÁM SÁT (nhận phiếu)", "NHÀ THẦU THI CÔNG"), ten=("{{dai_dien_tvgs}}", "{{dai_dien_nha_thau}}"))

    # ---------------- PHIEU_KT ----------------
    m = Mau(wb, kt)
    m.quoc_hieu("PHIẾU KIỂM TRA KÍCH THƯỚC HÌNH HỌC", so="Số: {{so_bien_ban}}  (kèm theo biên bản nghiệm thu công việc xây dựng)")
    m.dong("Công việc kiểm tra: {{ten_cong_viec}}", dam=True)
    m.dong("Vị trí / lý trình: {{ly_trinh}}  -  Ngày kiểm tra: {{ngay_kiem_tra}}")
    m.dong("Tiêu chuẩn áp dụng: {{tieu_chuan}}  -  Bản vẽ số: {{ban_ve_so}}")
    m.bang(["TT", "Vị trí (cọc / lý trình)", "Nội dung kiểm tra", "Thiết kế", "Thực tế", "Sai số", "Sai số cho phép", "Đánh giá"],
           ["{{kt.stt}}", "{{kt.vi_tri}}", "{{kt.noi_dung}}", "{{kt.thiet_ke}}", "{{kt.thuc_te}}", "{{kt.sai_so}}",
            "{{kt.cho_phep}}", "{{kt.danh_gia}}"])
    m.dong("Kết luận: {{ket_luan}}", dam=True)
    m.chu_ky()

    # ---------------- BB_NT_CONG_VIEC ----------------
    m = Mau(wb, nt)
    m.quoc_hieu("BIÊN BẢN NGHIỆM THU CÔNG VIỆC XÂY DỰNG")
    m.dong("1. Đối tượng nghiệm thu: {{ten_cong_viec}}", dam=True)
    m.dong("    Vị trí / lý trình: {{ly_trinh}}")
    m.dong("2. Thành phần trực tiếp nghiệm thu:", dam=True)
    m.thanh_phan()
    m.dong("3. Thời gian nghiệm thu: bắt đầu {{gio_bat_dau}}, kết thúc {{gio_ket_thuc}}, ngày {{ngay_nghiem_thu}} tại hiện trường.", dam=True)
    m.dong("4. Đánh giá công việc xây dựng đã thực hiện:", dam=True)
    m.dong("a) Tài liệu làm căn cứ nghiệm thu:")
    m.dong("    - Phiếu yêu cầu nghiệm thu của nhà thầu số {{so_phieu_yeu_cau}};")
    m.dong("    - Hồ sơ thiết kế bản vẽ thi công: {{ho_so_thiet_ke}}; bản vẽ số {{ban_ve_so}};")
    m.dong("    - Tiêu chuẩn, quy chuẩn áp dụng: {{tieu_chuan}};")
    m.dong("    - Kết quả thí nghiệm, phiếu kiểm tra kèm theo: {{tai_lieu_kem_theo}};")
    m.dong("    - Nhật ký thi công và các văn bản khác có liên quan.")
    m.dong("b) Về chất lượng công việc xây dựng: {{danh_gia_chat_luong}}")
    m.dong("c) Khối lượng nghiệm thu: {{khoi_luong_nghiem_thu}} {{don_vi}}  (khối lượng theo thiết kế/dự toán: {{khoi_luong}} {{don_vi}})")
    m.dong("d) Các ý kiến khác (nếu có): {{y_kien_khac}}")
    m.dong("5. Kết luận: {{ket_luan}}", dam=True)
    m.chu_ky()

    out = Path(__file__).resolve().parent.parent / "assets" / "file-mau-minh-hoa.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print("Đã tạo", out)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Quét file mẫu hồ sơ hoàn công -> mau.json (trình tự nghiệm thu + danh sách biểu mẫu + trường cần điền).

Đây là cơ chế "tự cập nhật" của skill: người dùng thêm sheet biểu mẫu hoặc thêm dòng vào
sheet TRINH_TU thì lần quét sau skill tự biết, không phải sửa code.

  python quet_file_mau.py [file_mau.xlsx] -o mau.json
  python quet_file_mau.py file_mau.xlsx --chi-tiet "BB_NT_CONG_VIEC"   # in mọi ô có chữ của 1 sheet
"""
import argparse
import re

from chung import (SHEET_THONG_TIN, SHEET_TRINH_TU, chuan_hoa, file_mau_mac_dinh, ghi_json,
                   la_sheet_bo_qua, mo_xlsx, tach_placeholder)

COT_TRINH_TU = {
    "ma_cong_tac": r"ma cong tac|ma loai",
    "loai_cong_tac": r"loai cong tac|ten cong tac|cong tac",
    "tu_khoa": r"tu khoa",
    "ma_hieu": r"ma hieu|ma dinh muc",
    "buoc": r"^buoc|thu tu|^tt$|^stt$",
    "noi_dung": r"noi dung|ten buoc",
    "bieu_mau": r"bieu mau|sheet|mau bien ban",
    "tieu_chuan": r"tieu chuan|quy chuan|can cu",
    "ghi_chu": r"tan suat|ghi chu",
}


def doc_trinh_tu(ws):
    cot, dong_td = {}, None
    for r in range(1, min(ws.max_row, 15) + 1):
        tam = {}
        for c in range(1, ws.max_column + 1):
            t = chuan_hoa(ws.cell(r, c).value)
            for vai, mau in COT_TRINH_TU.items():
                if t and vai not in tam and re.search(mau, t):
                    tam[vai] = c
                    break
        if {"loai_cong_tac", "bieu_mau"} <= tam.keys():
            cot, dong_td = tam, r
            break
    if not cot:
        return None, ["Sheet TRINH_TU không có dòng tiêu đề hợp lệ (cần ít nhất cột 'Loại công tác' và 'Biểu mẫu')."]

    ds, hien, loi = [], None, []
    for r in range(dong_td + 1, ws.max_row + 1):
        v = {vai: ws.cell(r, c).value for vai, c in cot.items()}
        v = {k: (str(x).strip() if x is not None else "") for k, x in v.items()}
        if not any(v.values()):
            continue
        if v.get("loai_cong_tac"):
            hien = {
                "ma_cong_tac": v.get("ma_cong_tac") or f"CT{len(ds) + 1:02d}",
                "loai_cong_tac": v["loai_cong_tac"],
                "tu_khoa": [t.strip() for t in re.split(r"[;\n]", v.get("tu_khoa", "")) if t.strip()],
                "ma_hieu": [t.strip() for t in re.split(r"[;,\n]", v.get("ma_hieu", "")) if t.strip()],
                "buoc": [],
            }
            ds.append(hien)
        if hien is None:
            loi.append(f"TRINH_TU dòng {r}: có bước nhưng chưa khai báo loại công tác phía trên.")
            continue
        if v.get("bieu_mau") or v.get("noi_dung"):
            hien["buoc"].append({
                "buoc": v.get("buoc") or str(len(hien["buoc"]) + 1),
                "noi_dung": v.get("noi_dung", ""),
                "bieu_mau": v.get("bieu_mau", ""),
                "tieu_chuan": v.get("tieu_chuan", ""),
                "ghi_chu": v.get("ghi_chu", ""),
                "dong": r,
            })
    return ds, loi


def doc_thong_tin(ws):
    """Sheet THONG_TIN: cột A = tên trường, B = giá trị mặc định, C = mô tả."""
    kq = []
    for r in range(1, ws.max_row + 1):
        ten = ws.cell(r, 1).value
        if isinstance(ten, str) and re.fullmatch(r"[a-z][a-z0-9_]*", ten.strip()):
            kq.append({"truong": ten.strip(), "mac_dinh": ws.cell(r, 2).value, "mo_ta": ws.cell(r, 3).value or ""})
    return kq


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file", nargs="?")
    ap.add_argument("-o", "--out", default="mau.json")
    ap.add_argument("--chi-tiet", metavar="SHEET", help="in toàn bộ ô có nội dung của sheet (để gắn {{trường}} cho biểu mẫu mới)")
    a = ap.parse_args()
    duong_dan = a.file or file_mau_mac_dinh()
    wb = mo_xlsx(duong_dan)

    if a.chi_tiet:
        ws = wb[a.chi_tiet]
        print(f"Sheet {ws.title} - vùng {ws.dimensions} - ô gộp: {[str(m) for m in ws.merged_cells.ranges][:80]}")
        for row in ws.iter_rows():
            for c in row:
                if c.value not in (None, ""):
                    print(f"{c.coordinate}\t{str(c.value)[:120]!r}")
        return

    canh_bao, trinh_tu, thong_tin, bieu_mau = [], [], [], {}
    if SHEET_TRINH_TU in wb.sheetnames:
        trinh_tu, loi = doc_trinh_tu(wb[SHEET_TRINH_TU])
        trinh_tu = trinh_tu or []
        canh_bao += loi
    else:
        canh_bao.append("File mẫu chưa có sheet TRINH_TU - cần người dùng cung cấp trình tự nghiệm thu.")
    if SHEET_THONG_TIN in wb.sheetnames:
        thong_tin = doc_thong_tin(wb[SHEET_THONG_TIN])

    for ws in wb.worksheets:
        if ws.title in (SHEET_TRINH_TU, SHEET_THONG_TIN) or la_sheet_bo_qua(ws.title):
            continue
        truong, bang = tach_placeholder(ws)
        bieu_mau[ws.title] = {"truong": truong, "bang": bang, "vung": ws.dimensions}
        if not truong and not bang:
            canh_bao.append(
                f"Biểu mẫu '{ws.title}' chưa có ô {{{{trường}}}} nào -> chưa điền tự động được. "
                f"Chạy --chi-tiet \"{ws.title}\" rồi gắn {{{{trường}}}} vào các ô cần điền (xem references/quy-uoc-file-mau.md)."
            )

    dung = set()
    for ct in trinh_tu:
        if not ct["tu_khoa"] and not ct["ma_hieu"]:
            canh_bao.append(f"Loại công tác '{ct['loai_cong_tac']}' chưa có từ khóa/mã hiệu -> không tự khớp được với dự toán.")
        for b in ct["buoc"]:
            if b["bieu_mau"] and b["bieu_mau"] not in bieu_mau:
                canh_bao.append(f"TRINH_TU dòng {b['dong']}: biểu mẫu '{b['bieu_mau']}' không có sheet tương ứng.")
            dung.add(b["bieu_mau"])
    for ten in bieu_mau:
        if ten not in dung:
            canh_bao.append(f"Biểu mẫu '{ten}' chưa được trình tự nào sử dụng (mới thêm? hãy bổ sung vào TRINH_TU).")

    ghi_json(a.out, {"file_mau": str(duong_dan), "thong_tin_chung": thong_tin, "trinh_tu": trinh_tu,
                     "bieu_mau": bieu_mau, "canh_bao": canh_bao})
    print(f"File mẫu: {duong_dan}")
    print(f"- {len(trinh_tu)} loại công tác, {sum(len(c['buoc']) for c in trinh_tu)} bước nghiệm thu")
    print(f"- {len(bieu_mau)} biểu mẫu: {', '.join(bieu_mau)}")
    print(f"- {len(thong_tin)} trường thông tin chung")
    for c in canh_bao:
        print("CẢNH BÁO:", c)
    print(f"Đã ghi {a.out}")


if __name__ == "__main__":
    main()

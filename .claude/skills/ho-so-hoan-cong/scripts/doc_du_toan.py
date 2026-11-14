#!/usr/bin/env python3
"""Đọc file dự toán Excel (.xlsx/.xlsm/.xls) -> JSON danh sách công tác.

Tự dò dòng tiêu đề (STT / Mã hiệu / Tên công tác / Đơn vị / Khối lượng) trên từng sheet,
nên dùng được với dự toán xuất từ G8, GXD, ACITT, Delta, Eta... hoặc bảng tự lập.

  python doc_du_toan.py du_toan.xlsx -o du_toan.json
  python doc_du_toan.py du_toan.xlsx --liet-ke          # chỉ xem sheet nào đọc được
  python doc_du_toan.py du_toan.xlsx --sheet "Dự toán"  # chỉ đọc 1 sheet
"""
import argparse
import re
import sys

from chung import chuan_hoa, ghi_json, mo_xlsx

MAU_COT = {
    "stt": r"^(stt|tt|so tt|so thu tu)$",
    "ma_hieu": r"(ma hieu|ma so|so hieu|mhdm|ma dm|ma dinh muc|ma cv|ma cong tac|ma don gia)",
    "ten": r"(ten cong|noi dung cong|danh muc|hang muc|ten vat tu, cong|dien giai cong|noi dung$|ten hang muc)",
    "don_vi": r"^(don vi|don vi tinh|dvt|dv|đv)$",
    "khoi_luong": r"^(khoi luong|kl|k\.luong|khoi luong thiet ke|khoi luong du toan|so luong)",
    "don_gia": r"^don gia",
    "thanh_tien": r"^thanh tien",
}


def so(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        t = v.strip().replace(" ", "")
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", t):  # 1.234,56
            t = t.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"-?\d+,\d+", t):
            t = t.replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return None
    return None


def do_tieu_de(ws):
    """Tìm dòng tiêu đề; trả về (dòng, {vai trò: cột}) hoặc None."""
    for r in range(1, min(ws.max_row, 60) + 1):
        cot = {}
        for rr in (r, r + 1):  # tiêu đề có thể gộp 2 dòng
            for c in range(1, min(ws.max_column, 40) + 1):
                t = chuan_hoa(ws.cell(rr, c).value)
                if not t:
                    continue
                for vai, mau in MAU_COT.items():
                    if vai not in cot and re.search(mau, t):
                        cot[vai] = c
                        break
            if rr == r and not ("ten" in cot and "don_vi" in cot):
                break
        if {"ten", "don_vi", "khoi_luong"} <= cot.keys():
            return r, cot
    return None


def doc_sheet(ws, ws_ct):
    kq = do_tieu_de(ws)
    if not kq:
        return None
    dong_td, cot = kq
    cong_tac, hang_muc, hien, thieu_gt = [], "", None, 0

    def o(r, vai):
        return ws.cell(r, cot[vai]).value if vai in cot else None

    for r in range(dong_td + 1, ws.max_row + 1):
        ten = o(r, "ten")
        if ten is None or str(ten).strip() == "":
            continue
        ten = re.sub(r"\s+", " ", str(ten)).strip()
        dv, kl_raw = o(r, "don_vi"), o(r, "khoi_luong")
        kl, stt, mh = so(kl_raw), o(r, "stt"), o(r, "ma_hieu")
        if chuan_hoa(ten) in ("ten cong tac", "noi dung cong viec") or (
            so(ten) is not None and so(stt) is not None
        ):
            continue  # dòng đánh số cột (1) (2) (3)...
        co_ct = ws_ct is not None and isinstance(ws_ct.cell(r, cot["khoi_luong"]).value, str)
        if dv and str(dv).strip() and (kl is not None or co_ct):
            if kl is None:
                thieu_gt += 1
            hien = {
                "dong": r,
                "stt": str(stt).strip() if stt is not None else "",
                "ma_hieu": str(mh).strip() if mh is not None else "",
                "ten": ten,
                "don_vi": str(dv).strip(),
                "khoi_luong": round(kl, 4) if kl is not None else None,
                "hang_muc": hang_muc,
                "dien_giai": [],
            }
            cong_tac.append(hien)
        elif stt not in (None, "") or (mh in (None, "") and ten.isupper()):
            hang_muc, hien = ten, None  # dòng nhóm / hạng mục
        elif hien is not None and len(hien["dien_giai"]) < 40:
            # dòng diễn giải bóc tách khối lượng nằm dưới công tác
            so_lieu = [
                ws.cell(r, c).value
                for c in range(cot["ten"] + 1, min(ws.max_column, 40) + 1)
                if isinstance(ws.cell(r, c).value, (int, float))
            ]
            hien["dien_giai"].append(ten + (f"  {so_lieu}" if so_lieu else ""))
    return {
        "sheet": ws.title,
        "dong_tieu_de": dong_td,
        "so_cong_tac": len(cong_tac),
        "thieu_gia_tri_cong_thuc": thieu_gt,
        "cong_tac": cong_tac,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("-o", "--out", default="du_toan.json")
    ap.add_argument("--sheet", help="chỉ đọc sheet này")
    ap.add_argument("--liet-ke", action="store_true", help="chỉ liệt kê các sheet nhận dạng được")
    a = ap.parse_args()

    wb = mo_xlsx(a.file, data_only=True)
    try:
        wb_ct = mo_xlsx(a.file, data_only=False)
    except SystemExit:
        wb_ct = None
    ds = []
    for ws in wb.worksheets:
        if a.sheet and ws.title != a.sheet:
            continue
        kq = doc_sheet(ws, wb_ct[ws.title] if wb_ct and ws.title in wb_ct.sheetnames else None)
        if kq and kq["so_cong_tac"]:
            ds.append(kq)
    if not ds:
        sys.exit(
            "Không nhận dạng được bảng dự toán nào (cần các cột Tên công tác / Đơn vị / Khối lượng). "
            f"Các sheet có trong file: {wb.sheetnames}"
        )
    for s in ds:
        print(f"- Sheet '{s['sheet']}': {s['so_cong_tac']} công tác (tiêu đề ở dòng {s['dong_tieu_de']})")
        if s["thieu_gia_tri_cong_thuc"]:
            print(
                f"  CẢNH BÁO: {s['thieu_gia_tri_cong_thuc']} công tác có khối lượng là công thức chưa được tính "
                "(file chưa từng mở/lưu bằng Excel). Mở file bằng Excel, lưu lại rồi đọc lại."
            )
    if a.liet_ke:
        return
    ghi_json(a.out, {"file": a.file, "bang": ds})
    print(f"Đã ghi {a.out}. Nếu có nhiều sheet, chọn sheet dự toán chi tiết (thường nhiều công tác và có diễn giải nhất).")


if __name__ == "__main__":
    main()

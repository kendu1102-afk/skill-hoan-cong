#!/usr/bin/env python3
"""Ghép dự toán với trình tự nghiệm thu trong file mẫu -> khung ke_hoach.json.

Mỗi công tác dự toán khớp được với một loại công tác trong TRINH_TU sẽ sinh ra đủ các bước
(biên bản lấy mẫu, phiếu kiểm tra, biên bản nghiệm thu...). Các trường chưa biết để null
- Claude điền tiếp từ bản vẽ; cái gì cần số liệu hiện trường thì giữ null.

  python lap_khung_ke_hoach.py du_toan.json mau.json -o ke_hoach.json [--sheet "Dự toán"]
"""
import argparse
import sys

from chung import chuan_hoa, doc_json, ghi_json


def khop(cong_tac, trinh_tu):
    """Chọn loại công tác theo mã hiệu định mức (ưu tiên) rồi tới từ khóa dài nhất."""
    ten, mh = chuan_hoa(cong_tac["ten"]), chuan_hoa(cong_tac["ma_hieu"]).replace(" ", "")
    tot, diem_tot = None, 0
    for ct in trinh_tu:
        diem = 0
        for m in ct["ma_hieu"]:
            m = chuan_hoa(m).replace(" ", "")
            if m and mh.startswith(m):
                diem = max(diem, 1000 + len(m))
        for k in ct["tu_khoa"]:
            k = chuan_hoa(k)
            if k and k in ten:
                diem = max(diem, len(k))
        if diem > diem_tot:
            tot, diem_tot = ct, diem
    return tot


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("du_toan")
    ap.add_argument("mau")
    ap.add_argument("-o", "--out", default="ke_hoach.json")
    ap.add_argument("--sheet", help="sheet dự toán cần dùng (mặc định: sheet nhiều công tác nhất)")
    a = ap.parse_args()

    dt, mau = doc_json(a.du_toan), doc_json(a.mau)
    bang = dt["bang"]
    if a.sheet:
        bang = [b for b in bang if b["sheet"] == a.sheet]
        if not bang:
            sys.exit(f"Không có sheet '{a.sheet}' trong {a.du_toan}")
    b = max(bang, key=lambda x: x["so_cong_tac"])
    ten_chung = {t["truong"] for t in mau["thong_tin_chung"]}

    bien_ban, chua_khop = [], []
    for ct in b["cong_tac"]:
        loai = khop(ct, mau["trinh_tu"])
        if not loai:
            chua_khop.append({k: ct[k] for k in ("stt", "ma_hieu", "ten", "don_vi", "khoi_luong", "hang_muc")})
            continue
        for buoc in loai["buoc"]:
            bm = mau["bieu_mau"].get(buoc["bieu_mau"])
            if not bm:
                continue
            biet = {
                "ten_cong_viec": ct["ten"], "don_vi": ct["don_vi"], "khoi_luong": ct["khoi_luong"],
                "ma_hieu": ct["ma_hieu"], "hang_muc": ct["hang_muc"], "loai_cong_tac": loai["loai_cong_tac"],
                "tieu_chuan": buoc["tieu_chuan"], "noi_dung_buoc": buoc["noi_dung"], "ghi_chu_buoc": buoc["ghi_chu"],
            }
            truong = {t: biet.get(t) for t in bm["truong"] if t not in ten_chung or t in biet}
            bien_ban.append({
                "mau": buoc["bieu_mau"],
                "ten_sheet": "",  # để trống -> tự đặt theo số thứ tự
                "cong_tac": f"{ct['stt']} {ct['ten']}".strip(),
                "ma_cong_tac": loai["ma_cong_tac"],
                "buoc": buoc["buoc"],
                "dong_du_toan": ct["dong"],
                "truong": truong,
                "bang": {ten: [] for ten in bm["bang"]},
                "_cot_bang": bm["bang"],
            })

    ghi_json(a.out, {
        "file_mau": mau["file_mau"],
        "chung": {t["truong"]: t["mac_dinh"] for t in mau["thong_tin_chung"]},
        "bien_ban": bien_ban,
        "cong_tac_chua_co_trinh_tu": chua_khop,
    })
    print(f"Sheet dự toán '{b['sheet']}': {b['so_cong_tac']} công tác -> {len(bien_ban)} biên bản/phiếu trong khung.")
    if chua_khop:
        print(f"CHƯA KHỚP {len(chua_khop)} công tác (không có trình tự trong file mẫu):")
        for c in chua_khop:
            print(f"  - [{c['ma_hieu']}] {c['ten']} ({c['khoi_luong']} {c['don_vi']})")
        print("  -> Xem có phải công tác phụ không cần nghiệm thu riêng; nếu cần, đề nghị người dùng bổ sung TRINH_TU.")
    print(f"Đã ghi {a.out}")


if __name__ == "__main__":
    main()

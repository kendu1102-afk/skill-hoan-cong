#!/usr/bin/env python3
"""Soát lại hồ sơ đã sinh trước khi giao: lỗi logic mà thanh tra/kiểm toán hay bắt.

Kiểm tra:
  1. Còn sót {{trường}} chưa thay trong file kết quả.
  2. Trùng số biên bản.
  3. Trình tự ngày: trong cùng một công tác, bước sau không được có ngày sớm hơn bước trước
     (lấy mẫu -> nghiệm thu vật liệu -> kiểm tra -> nghiệm thu công việc).
  4. Ngày nằm ngoài thời gian thi công (chung.ngay_khoi_cong .. chung.ngay_hoan_thanh) nếu có khai báo.
  5. Công tác dự toán chưa có biên bản nào (nếu truyền --du-toan).
  6. Tổng khối lượng nghiệm thu các đợt so với khối lượng dự toán (trường khoi_luong_nghiem_thu).

  python kiem_tra_ho_so.py ke_hoach.json "Ho so.xlsx" [--du-toan du_toan.json]
"""
import argparse
import re
from collections import defaultdict
from datetime import date, datetime

from chung import doc_json, mo_xlsx


def ngay(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", v)
        if m:
            try:
                return date(int(m[3]), int(m[2]), int(m[1]))
            except ValueError:
                return None
    return None


def khoa_buoc(b):
    m = re.match(r"\d+(\.\d+)?", str(b))
    return float(m[0]) if m else 0.0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ke_hoach")
    ap.add_argument("ho_so")
    ap.add_argument("--du-toan")
    a = ap.parse_args()
    kh = doc_json(a.ke_hoach)
    loi, canh_bao = [], []

    wb = mo_xlsx(a.ho_so)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "{{" in c.value:
                    loi.append(f"[{ws.title}!{c.coordinate}] còn sót trường chưa thay: {c.value[:60]}")

    chung = kh.get("chung") or {}
    kc, ht = ngay(chung.get("ngay_khoi_cong")), ngay(chung.get("ngay_hoan_thanh"))
    so_bb, nhom, kl_nt = defaultdict(list), defaultdict(list), defaultdict(float)
    for i, bb in enumerate(kh["bien_ban"], 1):
        t = bb.get("truong") or {}
        nhan = bb.get("ten_sheet") or f"#{i}"
        if t.get("so_bien_ban"):
            so_bb[str(t["so_bien_ban"])].append(nhan)
        ds_ngay = [(k, ngay(v)) for k, v in t.items() if k.startswith("ngay") and ngay(v)]
        for k, d in ds_ngay:
            if kc and d < kc:
                loi.append(f"[{nhan}] {k} = {d:%d/%m/%Y} trước ngày khởi công {kc:%d/%m/%Y}")
            if ht and d > ht:
                loi.append(f"[{nhan}] {k} = {d:%d/%m/%Y} sau ngày hoàn thành {ht:%d/%m/%Y}")
        if ds_ngay:
            nhom[(bb.get("cong_tac", ""), bb.get("dot", ""))].append((khoa_buoc(bb.get("buoc", 0)), max(d for _, d in ds_ngay), nhan))
        if isinstance(t.get("khoi_luong_nghiem_thu"), (int, float)) and bb.get("dong_du_toan"):
            kl_nt[bb["dong_du_toan"]] += t["khoi_luong_nghiem_thu"]

    for so, ds in so_bb.items():
        if len(ds) > 1:
            loi.append(f"Trùng số biên bản '{so}': {', '.join(ds)}")
    for (ct, dot), ds in nhom.items():
        ds.sort()
        for (b1, d1, n1), (b2, d2, n2) in zip(ds, ds[1:]):
            if b2 > b1 and d2 < d1:
                loi.append(f"Sai trình tự ngày ở '{ct}' {dot}: [{n2}] (bước {b2:g}, {d2:%d/%m/%Y}) sớm hơn [{n1}] (bước {b1:g}, {d1:%d/%m/%Y})")

    if a.du_toan:
        dt = doc_json(a.du_toan)
        co = {bb.get("dong_du_toan") for bb in kh["bien_ban"]}
        bo_qua = {(c["ma_hieu"], c["ten"]) for c in kh.get("cong_tac_khong_nghiem_thu_rieng", [])}
        b = max(dt["bang"], key=lambda x: x["so_cong_tac"])
        for c in b["cong_tac"]:
            if c["dong"] not in co and (c["ma_hieu"], c["ten"]) not in bo_qua:
                canh_bao.append(f"Công tác dự toán chưa có biên bản: [{c['ma_hieu']}] {c['ten']}")
            elif c["dong"] in kl_nt and c["khoi_luong"]:
                lech = kl_nt[c["dong"]] - c["khoi_luong"]
                if abs(lech) > 0.005 * abs(c["khoi_luong"]):
                    canh_bao.append(f"Khối lượng nghiệm thu cộng dồn {kl_nt[c['dong']]:g} khác dự toán {c['khoi_luong']:g} {c['don_vi']}: {c['ten']}")

    print(f"KẾT QUẢ SOÁT: {len(loi)} lỗi, {len(canh_bao)} cảnh báo")
    for x in loi:
        print("LỖI:", x)
    for x in canh_bao:
        print("CẢNH BÁO:", x)
    raise SystemExit(1 if loi else 0)


if __name__ == "__main__":
    main()

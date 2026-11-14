#!/usr/bin/env python3
"""Sinh file hồ sơ hoàn công (.xlsx) từ ke_hoach.json và file mẫu.

Mỗi phần tử trong ke_hoach["bien_ban"] -> 1 sheet sao từ biểu mẫu (giữ nguyên định dạng, ô gộp,
khổ in), điền {{trường}} và nhân dòng cho các bảng {{bang.cot}}. Ô chưa có số liệu được để trống
và TÔ VÀNG để người làm hồ sơ biết chỗ cần bổ sung số liệu hiện trường.

  python tao_ho_so.py ke_hoach.json -o "Ho so hoan cong.xlsx" [--mau file_mau.xlsx]
"""
import argparse
import copy
import re
from pathlib import Path

from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.cell_range import CellRange

from chung import (PH, SHEET_THONG_TIN, SHEET_TRINH_TU, doc_json, file_mau_mac_dinh, ghi_json,
                   mo_xlsx)

VANG = PatternFill("solid", fgColor="FFFF00")


def ten_sheet_hop_le(ten, da_dung):
    ten = re.sub(r"[\[\]\*\?/\\:]", "-", str(ten)).strip("' ")[:31] or "BB"
    goc, i = ten, 2
    while ten.lower() in da_dung:
        duoi = f"({i})"
        ten, i = goc[: 31 - len(duoi)] + duoi, i + 1
    da_dung.add(ten.lower())
    return ten


def sao_sheet(wb, nguon, ten):
    ws = wb.copy_worksheet(nguon)
    ws.title = ten
    # copy_worksheet bỏ sót các thiết lập in ấn -> sao thêm cho bản in giống mẫu
    if nguon.print_area:
        ws.print_area = nguon.print_area
    ws.print_title_rows = nguon.print_title_rows
    ws.print_title_cols = nguon.print_title_cols
    for thuoc_tinh in ("oddHeader", "oddFooter", "evenHeader", "evenFooter"):
        setattr(ws, thuoc_tinh, copy.copy(getattr(nguon, thuoc_tinh)))
    ws.sheet_properties.pageSetUpPr = copy.copy(nguon.sheet_properties.pageSetUpPr)
    ws.sheet_view.zoomScale = nguon.sheet_view.zoomScale
    ws.sheet_view.showGridLines = nguon.sheet_view.showGridLines
    ws.freeze_panes = nguon.freeze_panes
    for dv in nguon.data_validations.dataValidation:
        ws.add_data_validation(copy.copy(dv))
    return ws


def chen_dong(ws, dong, k):
    """Chèn k dòng ngay dưới `dong`, sao định dạng + nội dung + ô gộp ngang của dòng đó."""
    cao = {r: d.height for r, d in ws.row_dimensions.items() if r > dong and d.height}
    cao_mau = ws.row_dimensions[dong].height
    gop_ngang = [(m.min_col, m.max_col) for m in ws.merged_cells.ranges if m.min_row == m.max_row == dong]
    ws.insert_rows(dong + 1, k)
    for m in ws.merged_cells.ranges:  # openpyxl không tự dời ô gộp
        if m.min_row > dong:
            m.shift(row_shift=k)
        elif m.max_row > dong:
            m.expand(down=k)
    for r in cao:
        ws.row_dimensions[r].height = None
    for r, h in cao.items():
        ws.row_dimensions[r + k].height = h
    for i in range(1, k + 1):
        ws.row_dimensions[dong + i].height = cao_mau
        for c in range(1, ws.max_column + 1):
            nguon, dich = ws.cell(dong, c), ws.cell(dong + i, c)
            if nguon.has_style:
                dich._style = copy.copy(nguon._style)
            if type(nguon).__name__ != "MergedCell":
                dich.value = nguon.value
        for c1, c2 in gop_ngang:
            ws.merge_cells(start_row=dong + i, start_column=c1, end_row=dong + i, end_column=c2)
    if ws.print_area:
        vung = []
        for v in (ws.print_area if isinstance(ws.print_area, list) else str(ws.print_area).split(",")):
            cr = CellRange(v.split("!")[-1].replace("$", ""))
            if cr.max_row >= dong:
                cr.expand(down=k)
            vung.append(cr.coord)
        ws.print_area = ",".join(vung)


def dien_o(o, tra, thieu, nhan):
    """Thay các {{trường}} trong một ô; trả về None, ghi nhận trường thiếu."""
    v = o.value
    if not (isinstance(v, str) and "{{" in v):
        return
    ten_ph = PH.findall(v)
    co_thieu = False
    if PH.fullmatch(v.strip()):  # ô chỉ có 1 trường -> giữ kiểu số/ngày
        gt = tra(ten_ph[0])
        co_thieu = gt in (None, "")
        o.value = None if co_thieu else gt
    else:
        def thay(m):
            nonlocal co_thieu
            gt = tra(m.group(1).strip())
            if gt in (None, ""):
                co_thieu = True
                return "........"
            if isinstance(gt, float):
                gt = f"{gt:,.3f}".rstrip("0").rstrip(".").replace(",", "X").replace(".", ",").replace("X", ".")
            return str(gt)
        o.value = PH.sub(thay, v)
    if co_thieu:
        o.fill = VANG
        thieu.append({"o": o.coordinate, "truong": [t for t in ten_ph if tra(t) in (None, "")], "ngu_canh": nhan})


def dien_sheet(ws, bb, chung, thieu):
    truong, bang = bb.get("truong") or {}, bb.get("bang") or {}
    # 1) nhân dòng bảng - làm từ dưới lên để chỉ số dòng phía trên không đổi
    dong_bang = {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "{{" in c.value:
                for t in PH.findall(c.value):
                    if "." in t:
                        dong_bang[c.row] = t.split(".", 1)[0]
    for dong in sorted(dong_bang, reverse=True):
        ten_bang = dong_bang[dong]
        muc = bang.get(ten_bang) or []
        if len(muc) > 1:
            chen_dong(ws, dong, len(muc) - 1)
        for i in range(max(len(muc), 1)):
            dl = muc[i] if muc else {}

            def tra(t, dl=dl, i=i):
                if "." in t:
                    cot = t.split(".", 1)[1]
                    return dl.get(cot, i + 1 if cot == "stt" and muc else None)
                return truong.get(t, chung.get(t))
            for c in range(1, ws.max_column + 1):
                dien_o(ws.cell(dong + i, c), tra, thieu, f"bảng {ten_bang} dòng {i + 1}")
        if not muc:
            thieu.append({"o": f"dòng {dong}", "truong": [f"{ten_bang}.*"], "ngu_canh": f"bảng {ten_bang} chưa có dòng nào"})
    # 2) trường đơn
    for row in ws.iter_rows():
        for c in row:
            dien_o(c, lambda t: truong.get(t, chung.get(t)), thieu, "")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ke_hoach")
    ap.add_argument("-o", "--out", default="Ho-so-hoan-cong.xlsx")
    ap.add_argument("--mau", help="file mẫu (mặc định: lấy trong ke_hoach.json hoặc assets/)")
    a = ap.parse_args()

    kh = doc_json(a.ke_hoach)
    duong_mau = a.mau or kh.get("file_mau") or file_mau_mac_dinh()
    wb = mo_xlsx(duong_mau)
    chung = kh.get("chung") or {}
    sheet_goc = list(wb.sheetnames)
    da_dung = {s.lower() for s in sheet_goc} | {"danh_muc"}

    bao_cao, muc_luc = [], []
    for i, bb in enumerate(kh["bien_ban"], 1):
        if bb["mau"] not in wb.sheetnames:
            bao_cao.append({"stt": i, "loi": f"Không có biểu mẫu '{bb['mau']}' trong file mẫu", "cong_tac": bb.get("cong_tac", "")})
            continue
        ten = ten_sheet_hop_le(bb.get("ten_sheet") or f"{i:03d}_{bb['mau']}", da_dung)
        ws = sao_sheet(wb, wb[bb["mau"]], ten)
        thieu = []
        dien_sheet(ws, bb, chung, thieu)
        bb["ten_sheet"] = ten
        t = bb.get("truong") or {}
        muc_luc.append((i, t.get("so_bien_ban") or "", bb["mau"], bb.get("cong_tac", ""),
                        t.get("ngay") or t.get("ngay_nghiem_thu") or t.get("ngay_lay_mau") or "", ten, len(thieu)))
        if thieu:
            bao_cao.append({"stt": i, "sheet": ten, "mau": bb["mau"], "cong_tac": bb.get("cong_tac", ""), "thieu": thieu})

    # sheet THONG_TIN: ghi giá trị thực của công trình; bỏ các sheet biểu mẫu gốc và TRINH_TU
    if SHEET_THONG_TIN in wb.sheetnames:
        ws = wb[SHEET_THONG_TIN]
        for r in range(1, ws.max_row + 1):
            k = ws.cell(r, 1).value
            if isinstance(k, str) and k.strip() in chung:
                ws.cell(r, 2).value = chung[k.strip()]
    for s in sheet_goc:
        if s != SHEET_THONG_TIN:
            del wb[s]

    dm = wb.create_sheet("DANH_MUC", 0)
    dm.append(["DANH MỤC HỒ SƠ HOÀN CÔNG"])
    dm["A1"].font = Font(bold=True, size=14)
    dm.append([chung.get("ten_cong_trinh") or ""])
    dm.append([])
    dm.append(["STT", "Số biên bản", "Loại biểu mẫu", "Công tác / hạng mục", "Ngày", "Sheet", "Số ô còn thiếu"])
    for c in dm[4]:
        c.font = Font(bold=True)
    for d in muc_luc:
        dm.append(list(d))
        o = dm.cell(dm.max_row, 6)
        o.hyperlink, o.style = f"#'{d[5]}'!A1", "Hyperlink"
        if d[6]:
            dm.cell(dm.max_row, 7).fill = VANG
    for cot, rong in zip("ABCDEFG", (6, 16, 22, 60, 13, 30, 14)):
        dm.column_dimensions[cot].width = rong
    wb.active = 0

    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    ghi_json(a.ke_hoach, kh)  # lưu lại tên sheet đã đặt để kiem_tra_ho_so.py đối chiếu
    duong_bc = str(Path(a.out).with_suffix("")) + "_con-thieu.json"
    ghi_json(duong_bc, bao_cao)
    so_o = sum(len(b.get("thieu", [])) for b in bao_cao)
    print(f"Đã tạo {a.out}: {len(muc_luc)} biên bản/phiếu.")
    print(f"Còn {so_o} ô chưa có số liệu (đã tô vàng) trên {sum(1 for b in bao_cao if 'thieu' in b)} sheet -> chi tiết: {duong_bc}")
    for b in bao_cao:
        if "loi" in b:
            print("LỖI:", b["loi"], "-", b["cong_tac"])


if __name__ == "__main__":
    main()

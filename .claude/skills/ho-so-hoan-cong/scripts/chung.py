"""Hàm dùng chung cho các script của skill ho-so-hoan-cong."""
import json
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path

# {{ten_truong}} hoặc {{ten_bang.ten_cot}}
PH = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")

SHEET_TRINH_TU = "TRINH_TU"
SHEET_THONG_TIN = "THONG_TIN"
THU_MUC_SKILL = Path(__file__).resolve().parent.parent


def chuan_hoa(s):
    """Bỏ dấu tiếng Việt, viết thường, gộp khoảng trắng - để so khớp từ khóa."""
    if s is None:
        return ""
    s = str(s).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def la_sheet_bo_qua(ten):
    """Sheet bắt đầu bằng '_' hoặc 'HD' là ghi chú/hướng dẫn, không phải biểu mẫu."""
    return ten.startswith("_") or chuan_hoa(ten).startswith("hd")


def file_mau_mac_dinh():
    """File mẫu của người dùng (assets/file-mau.xlsx) nếu có, không thì mẫu minh họa."""
    for ten in ("file-mau.xlsx", "file-mau-minh-hoa.xlsx"):
        p = THU_MUC_SKILL / "assets" / ten
        if p.exists():
            return p
    sys.exit("Không tìm thấy file mẫu trong assets/. Hãy truyền --mau <đường dẫn>.")


def mo_xlsx(duong_dan, **kw):
    """Mở workbook; tự chuyển .xls sang .xlsx bằng LibreOffice nếu cần."""
    import openpyxl

    p = Path(duong_dan)
    if not p.exists():
        sys.exit(f"Không tìm thấy file: {p}")
    if p.suffix.lower() == ".xls":
        if not shutil.which("soffice"):
            sys.exit("File .xls cần LibreOffice để chuyển đổi. Hãy lưu lại thành .xlsx rồi thử lại.")
        tam = Path(tempfile.mkdtemp(prefix="hshc_"))
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(tam), str(p)],
            check=True, capture_output=True, timeout=300,
        )
        p = tam / (p.stem + ".xlsx")
    if p.suffix.lower() == ".xlsm":
        kw.setdefault("keep_vba", True)
    return openpyxl.load_workbook(p, **kw)


def doc_json(p):
    return json.loads(Path(p).read_text(encoding="utf-8"))


def ghi_json(p, du_lieu):
    Path(p).parent.mkdir(parents=True, exist_ok=True)
    Path(p).write_text(json.dumps(du_lieu, ensure_ascii=False, indent=1), encoding="utf-8")


def tach_placeholder(ws):
    """Trả về (truong, bang): tập tên trường đơn và {tên bảng: [cột]} của một sheet."""
    truong, bang = [], {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "{{" in c.value:
                for ten in PH.findall(c.value):
                    if "." in ten:
                        b, cot = ten.split(".", 1)
                        bang.setdefault(b, [])
                        if cot not in bang[b]:
                            bang[b].append(cot)
                    elif ten not in truong:
                        truong.append(ten)
    return truong, bang

import io
import re

import openpyxl


def num(v):
    """Return v as float, or None for blanks / '-' / text."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if v == v else None
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if not s or s in "-–":
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def text(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


def load_wb(src):
    """Open an xlsx/xlsm from a path or bytes (values only, read-only)."""
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    return openpyxl.load_workbook(src, read_only=True, data_only=True)


def sheet_rows(ws, max_col=40):
    """All rows of a worksheet as lists, capped at max_col (some sheets declare XFD)."""
    return [list(r) for r in ws.iter_rows(min_col=1, max_col=max_col, values_only=True)]


def col_letter(idx):
    """0-based column index -> Excel letters."""
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1

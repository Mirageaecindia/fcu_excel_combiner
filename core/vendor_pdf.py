"""Vendor deviation report supplied as PDF.

The PDF is turned into rows of cells from positioned words (pdfplumber in the
CLI, pdf.js in the browser both feed `parse_words`). Data rows of the vendor
template have a fixed column layout (same as the Excel file), so rows with
enough cells are mapped by position; anything else raises a clear error.
"""
import re

from .util import num, text
from .vendor_excel import VendorRow

# 0-based positions in the 31-column vendor template (A..AE)
POS = dict(ref=0, space=1, model=12, flow=13, total=15, sens=16, power=25, qty=27)
MIN_CELLS = 28
GAP = 6.0          # pt: wider gap between words starts a new cell
ROW_TOL = 3.0      # pt: words closer than this vertically share a row


def words_to_rows(words):
    """words: iterable of dicts {page, x0, x1, top, text} -> list of cell lists."""
    by_page = {}
    for w in words:
        by_page.setdefault(w["page"], []).append(w)
    rows = []
    for page in sorted(by_page):
        ws = sorted(by_page[page], key=lambda w: (w["top"], w["x0"]))
        line, last_top = [], None
        for w in ws:
            if last_top is not None and abs(w["top"] - last_top) > ROW_TOL:
                rows.append(_cells(line))
                line = []
            line.append(w)
            last_top = w["top"] if not line[:-1] else last_top
        if line:
            rows.append(_cells(line))
    return rows


def _cells(line):
    line = sorted(line, key=lambda w: w["x0"])
    cells, cur, prev_x1 = [], [], None
    for w in line:
        if prev_x1 is not None and w["x0"] - prev_x1 > GAP:
            cells.append(" ".join(cur))
            cur = []
        cur.append(w["text"])
        prev_x1 = w["x1"]
    if cur:
        cells.append(" ".join(cur))
    return cells


def parse_words(words):
    rows = words_to_rows(words)
    out = []
    for i, r in enumerate(rows):
        if len(r) < MIN_CELLS:
            continue
        total = num(r[POS["total"]])
        if total is None or not text(r[POS["ref"]]):
            continue
        flow = num(r[POS["flow"]])
        power = num(r[POS["power"]])
        out.append(VendorRow(
            ref=text(r[POS["ref"]]), space=text(r[POS["space"]]), qty=num(r[POS["qty"]]),
            total_kw=total, sens_kw=num(r[POS["sens"]]),
            flow_ls=None if flow is None else flow / 3.6,
            power_kw=None if power is None else power / 1000.0,
            model=text(r[POS["model"]]), row=i + 1,
        ))
    if not out:
        raise ValueError(
            "Could not read any unit rows from the vendor PDF. The PDF must contain the "
            "vendor deviation table as selectable text (scanned images are not supported).")
    return out


def words_from_pdf(path):
    """CLI helper: positioned words via pdfplumber."""
    import pdfplumber
    words = []
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages):
            for w in page.extract_words(keep_blank_chars=False, use_text_flow=False):
                words.append(dict(page=pno, x0=w["x0"], x1=w["x1"], top=w["top"], text=w["text"]))
    return words


def parse(path):
    return parse_words(words_from_pdf(path))

"""Fill the reference workbook (used as template) by patching its XML in place.

openpyxl would drop the ActiveX button, logos and validations of the reference,
so the sheet XML is edited as text instead. Everything outside the touched
cells (header rows 1-12, protection, merges, hidden helper columns, ...) is
copied through byte for byte. The template already locks rows 1-12 and the
formula columns and leaves the data-entry cells unlocked, so styles are kept.
"""
import io
import re
import zipfile
from xml.sax.saxutils import escape

from .util import col_index

SHEET = "xl/worksheets/sheet1.xml"
FIRST, LAST = 13, 1012                      # data rows of the template
INPUT_COLS = ["D", "E", "F", "G", "H", "I", "J", "K", "N", "O", "P",
              "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF"]
FORMULA_COLS = ("L", "M", "Q", "R")
HEADER_CELLS = {"plot": "Y3", "project": "Y4", "consultant": "Y5"}

_ROW = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)
_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
_ATTR = re.compile(r'([\w:]+)="([^"]*)"')
_V = re.compile(r"<v[^>]*>.*?</v>|<v\s*/>", re.S)


def fmt_num(v):
    if float(v).is_integer():
        return str(int(v))
    return repr(round(float(v), 6))


def _cell_xml(col, row, attrs, inner):
    a = "".join(f' {k}="{v}"' for k, v in attrs.items())
    return f'<c r="{col}{row}"{a}>{inner}</c>' if inner else f'<c r="{col}{row}"{a}/>'


def _value_inner(value):
    """-> (type attr or None, inner xml) for a python value."""
    if value is None or value == "":
        return None, ""
    if isinstance(value, (int, float)):
        return None, f"<v>{fmt_num(value)}</v>"
    return "inlineStr", f'<is><t xml:space="preserve">{escape(str(value))}</t></is>'


def _parse_cells(inner):
    cells = []
    for m in _CELL.finditer(inner or ""):
        cells.append([m.group(1), dict(_ATTR.findall(m.group(3))), m.group(4) or ""])
    return cells


def _strip_cached(attrs, inner):
    """Drop the cached result of a formula cell so the app recalculates it."""
    if "<f" not in inner:
        return attrs, inner
    return {k: v for k, v in attrs.items() if k != "t"}, _V.sub("", inner)


def _get(cells, col):
    idx = col_index(col)
    for c in cells:
        if c[0] == col:
            return c
    c = [col, {}, ""]
    pos = next((i for i, x in enumerate(cells) if col_index(x[0]) > idx), len(cells))
    cells.insert(pos, c)
    return c


def _set(cells, col, value):
    """Replace the content of a cell, keeping its style."""
    c = _get(cells, col)
    t, inner = _value_inner(value)
    attrs = {k: v for k, v in c[1].items() if k != "t"}
    if t:
        attrs["t"] = t
    c[1], c[2] = attrs, inner


def _cached(cells, col, value, is_text=False):
    """Replace the cached result of a formula cell (None = no cached value)."""
    for c in cells:
        if c[0] != col:
            continue
        c[1] = {k: v for k, v in c[1].items() if k != "t"}
        c[2] = _V.sub("", c[2])
        if value is not None:
            if is_text:
                c[1]["t"] = "str"
                c[2] += f'<v xml:space="preserve">{escape(value)}</v>'
            else:
                c[2] += f"<v>{fmt_num(value)}</v>"


def _row_xml(r, attrs_s, cells):
    body = "".join(_cell_xml(c, r, a, i) for c, a, i in cells)
    return f'<row r="{r}"{attrs_s}>{body}</row>'


def _refresh_formulas(cells, d):
    """Cached values for L (density), M (DCV), Q, R of a filled row."""
    area, people = d.get("I"), d.get("J")
    dens = people / area * 100 if (area and people is not None) else None
    _cached(cells, "L", dens)
    req = " " if dens is None else ("Required" if (area >= 100 and dens >= 25) else " ")
    _cached(cells, "M", req, is_text=True)
    _cached(cells, "Q", d["N"] / area if (area and d.get("N") is not None) else None)
    _cached(cells, "R", d["P"] / area if (area and d.get("P") is not None) else None)


def fill(template, rows, header=None, hide_empty=True):
    """template: bytes of the reference xlsx.  rows: list of dicts col-letter -> value.

    Formula columns L, M, Q, R keep the template formulas (cached values are
    refreshed for the filled rows, the workbook recalculates on open).
    """
    cap = LAST - FIRST + 1
    if len(rows) > cap:
        raise ValueError(f"Too many units ({len(rows)}); the template holds {cap} rows")
    zin = zipfile.ZipFile(io.BytesIO(template))
    files = {n: zin.read(n) for n in zin.namelist()}
    sheet = files[SHEET].decode("utf-8")

    hdr = {}
    for k, ref in HEADER_CELLS.items():
        if (header or {}).get(k):
            hdr[int(re.sub(r"\D", "", ref))] = (re.sub(r"\d", "", ref), header[k])

    def do_row(m):
        r, attrs_s, inner = int(m.group(1)), m.group(2), m.group(3)
        if r in hdr:
            cells = _parse_cells(inner)
            _set(cells, hdr[r][0], hdr[r][1])
            return _row_xml(r, attrs_s, cells)
        if r < FIRST:
            return m.group(0)
        cells = _parse_cells(inner)
        in_data = FIRST <= r <= LAST
        data = rows[r - FIRST] if in_data and r - FIRST < len(rows) else None
        if in_data:
            for col in INPUT_COLS:
                _set(cells, col, (data or {}).get(col))
            if data:
                _refresh_formulas(cells, data)
            else:
                for col in FORMULA_COLS:
                    _cached(cells, col, None)
            attrs_s = re.sub(r'\s+hidden="1"', "", attrs_s)
            if not data and hide_empty:
                attrs_s += ' hidden="1"'
        for c in cells:                     # all other formulas: drop stale cached values
            if not (in_data and c[0] in FORMULA_COLS):
                c[1], c[2] = _strip_cached(c[1], c[2])
        return _row_xml(r, attrs_s, cells)

    files[SHEET] = _ROW.sub(do_row, sheet).encode("utf-8")

    ss = files["xl/sharedStrings.xml"].decode("utf-8")
    files["xl/sharedStrings.xml"] = ss.replace("<r><t>L/s.m</t></r>", "<r><t>(L/s)/m</t></r>").encode("utf-8")

    wb = files["xl/workbook.xml"].decode("utf-8")
    wb = re.sub(r"<calcPr([^>]*?)/>",
                lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="\d"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wb)
    files["xl/workbook.xml"] = wb.encode("utf-8")
    if "xl/calcChain.xml" in files:
        del files["xl/calcChain.xml"]
        ct = files["[Content_Types].xml"].decode("utf-8")
        files["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct).encode("utf-8")
        rl = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
        files["xl/_rels/workbook.xml.rels"] = re.sub(r"<Relationship [^>]*calcChain[^>]*/>", "", rl).encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return out.getvalue()

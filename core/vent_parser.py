"""Ventilation calculation workbooks -> list of VentSpace.

Only sheets that carry a 'proposed no. of people' column are used (the
'FAHU-xx Common Area' sheets); it is detected by header text, not sheet name.
"""
import re
from dataclasses import dataclass

from .util import load_wb, num, sheet_rows, text


@dataclass
class VentSpace:
    name: str
    area: float | None
    people: float | None
    fresh_air: float | None
    sheet: str = ""
    file: str = ""
    row: int = 0


def _labels(rows, top, bottom, ncols):
    out = []
    for c in range(ncols):
        parts = []
        for r in range(top, min(bottom, len(rows))):
            if c < len(rows[r]) and rows[r][c] is not None:
                parts.append(text(rows[r][c]).upper())
        out.append(" ".join(parts))
    return out


def parse(src, file_label=""):
    wb = load_wb(src)
    spaces = []
    for ws in wb.worksheets:
        rows = sheet_rows(ws, 30)
        if len(rows) < 6:
            continue
        ncols = 30
        lab = _labels(rows, 0, 6, ncols)

        def find(*pats, exclude=()):
            for c, l in enumerate(lab):
                if all(re.search(p, l) for p in pats) and not any(re.search(p, l) for p in exclude):
                    return c
            return None

        c_people = find(r"PROPOSED\s*NO.*PEOPLE")
        if c_people is None:
            continue
        c_name = find(r"DESCRIPTION|SPACE NAME")
        c_area = find(r"^AREA\b")
        c_fa = find(r"PROPOSED FRESH AIR")
        if c_name is None or c_area is None:
            continue
        start = 6
        for i in range(start, len(rows)):
            r = rows[i] + [None] * (ncols - len(rows[i]))
            name = text(r[c_name])
            area = num(r[c_area])
            if not name or area is None:
                continue
            spaces.append(VentSpace(
                name=name, area=area, people=num(r[c_people]),
                fresh_air=num(r[c_fa]) if c_fa is not None else None,
                sheet=ws.title, file=file_label, row=i + 1,
            ))
    wb.close()
    return spaces

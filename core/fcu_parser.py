"""FCU schedule -> list of FcuUnit.

A unit row has a numeric total coil load (col D); the rows below it with no
load are the spaces served by that unit (area in C, air flow L/s in F).
"""
from dataclasses import dataclass, field

from .util import load_wb, num, sheet_rows, text


@dataclass
class FcuUnit:
    tag: str
    zone: str
    area: float | None
    total_kw: float | None      # per unit
    sens_kw: float | None       # per unit
    qty: float
    fcu_type: str
    flow: float | None          # sum of space air flows (L/s)
    spaces: list = field(default_factory=list)   # [(name, area, flow)]
    sheet: str = ""
    row: int = 0

    @property
    def ref(self):
        return self.tag or self.zone.lstrip("& ")

    @property
    def calc_total_kw(self):
        return None if self.total_kw is None else self.total_kw * self.qty

    @property
    def calc_sens_kw(self):
        return None if self.sens_kw is None else self.sens_kw * self.qty

    @property
    def keys(self):
        """Names that may identify this unit in other files."""
        return [self.zone] + [s[0] for s in self.spaces if s[0]]


def _find_header(rows):
    for i, r in enumerate(rows[:15]):
        if text(r[0]).lower().replace(" ", "") == "fcutag":
            return i
    return None


def parse(src, sheets=None):
    wb = load_wb(src)
    units = []
    for ws in wb.worksheets:
        if sheets and ws.title not in sheets:
            continue
        rows = sheet_rows(ws, 16)
        h = _find_header(rows)
        if h is None:
            continue
        cur = None
        for i in range(h + 1, len(rows)):
            r = rows[i] + [None] * (16 - len(rows[i]))
            name = text(r[1])
            if not name:
                continue
            if num(r[3]) is not None:           # unit row
                cur = FcuUnit(
                    tag=text(r[0]), zone=name, area=num(r[2]), total_kw=num(r[3]),
                    sens_kw=num(r[4]), qty=num(r[10]) or 1.0, fcu_type=text(r[13]),
                    flow=num(r[5]), sheet=ws.title, row=i + 1,
                )
                units.append(cur)
            elif cur is not None and (num(r[2]) is not None or num(r[5]) is not None):
                cur.spaces.append((name, num(r[2]), num(r[5])))
        wb_units = [u for u in units if u.sheet == ws.title]
        for u in wb_units:
            if u.spaces:
                fl = [s[2] for s in u.spaces if s[2] is not None]
                u.flow = float(sum(fl)) if fl else u.flow
    wb.close()
    return units

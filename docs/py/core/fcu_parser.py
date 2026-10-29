"""FCU schedule -> list of FcuUnit.

A row with an FCU Tag starts a unit. Rows with an empty tag are sub parts of the
unit above (spaces served by it): their area and air flow belong to that unit.
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
    flow: float | None          # sum of the sub rows' air flow (L/s)
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
        """Names that may identify this unit in other files (the tag is not unique)."""
        return [self.zone] + [s[0] for s in self.spaces if s[0]]


def _find_header(rows):
    for i, r in enumerate(rows[:15]):
        if text(r[0]).lower().replace(" ", "") == "fcutag":
            return i
    return None


def parse(src, sheets=None, warnings=None):
    """warnings: optional list that receives human readable notes."""
    warn = warnings if warnings is not None else []
    wb = load_wb(src)
    units = []
    for ws in wb.worksheets:
        if sheets and ws.title not in sheets:
            continue
        rows = sheet_rows(ws, 16)
        h = _find_header(rows)
        if h is None:
            continue
        first = len(units)
        cur, merged, orphan = None, [], 0
        for i in range(h + 1, len(rows)):
            r = rows[i] + [None] * (16 - len(rows[i]))
            tag, name = text(r[0]), text(r[1])
            if tag:
                cur = FcuUnit(
                    tag=tag, zone=name, area=num(r[2]), total_kw=num(r[3]), sens_kw=num(r[4]),
                    qty=num(r[10]) or 1.0, fcu_type=text(r[13]), flow=num(r[5]),
                    sheet=ws.title, row=i + 1,
                )
                units.append(cur)
                continue
            if not name and num(r[2]) is None and num(r[5]) is None:
                continue
            if cur is None:
                orphan += 1
                continue
            if num(r[3]) is not None:
                merged.append(name or f"row {i + 1}")
            cur.spaces.append((name, num(r[2]), num(r[5])))
        for u in units[first:]:
            if u.spaces:
                fl = [s[2] for s in u.spaces if s[2] is not None]
                if fl:
                    u.flow = float(sum(fl))
                if u.area is None:
                    ar = [s[1] for s in u.spaces if s[1] is not None]
                    u.area = float(sum(ar)) if ar else None
        if merged:
            warn.append(f"{ws.title}: {len(merged)} rows have load values but no FCU Tag and were added to the unit "
                        f"above (e.g. {', '.join(merged[:3])}). Fill the FCU Tag column for every unit.")
        if orphan:
            warn.append(f"{ws.title}: {orphan} rows above the first FCU Tag were ignored.")
    wb.close()
    return units

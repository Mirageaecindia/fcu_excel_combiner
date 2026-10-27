"""Vendor deviation report (xlsx/xlsm) -> list of VendorRow (OFFERED values)."""
import re
from dataclasses import dataclass

from .util import load_wb, num, sheet_rows, text


@dataclass
class VendorRow:
    ref: str
    space: str
    qty: float | None
    total_kw: float | None       # per unit
    sens_kw: float | None        # per unit
    flow_ls: float | None        # per unit, L/s
    power_kw: float | None       # per unit, kW
    model: str = ""
    row: int = 0

    @property
    def keys(self):
        return [k for k in (self.ref, self.space) if k]


def list_sheets(src):
    wb = load_wb(src)
    names = list(wb.sheetnames)
    wb.close()
    return names


def default_sheet(names, hint="plot 04"):
    norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())
    for n in names:
        if norm(n) in (norm(hint), "plot4"):
            return n
    return next((n for n in names if norm(hint) in norm(n)), names[0])


def _hdr_rows(rows):
    for i, r in enumerate(rows[:12]):
        if any(text(c).lower() == "unit reference" for c in r[:3]):
            return i
    return None


def parse(src, sheet):
    wb = load_wb(src)
    rows = sheet_rows(wb[sheet], 34)
    wb.close()
    return parse_rows(rows)


def parse_rows(rows):
    h = _hdr_rows(rows)
    if h is None:
        raise ValueError("Vendor sheet has no 'Unit Reference' header row")
    n = max(len(r) for r in rows)
    lab = []
    for c in range(n):
        lab.append(" ".join(text(rows[r][c]).lower() for r in range(h, h + 3)
                            if c < len(rows[r]) and rows[r][c] is not None))
    model = next((c for c, l in enumerate(lab) if re.search(r"\bmodel\b", l)), None)
    if model is None:
        raise ValueError("Vendor sheet has no 'Model' column (OFFERED block)")

    def find(pat, start=model, exclude=r"%"):
        for c in range(start, n):
            if re.search(pat, lab[c]) and not re.search(exclude, lab[c]):
                return c
        return None

    c_total = find(r"total cooling")
    c_sens = find(r"sensible cooling")
    c_flow = find(r"air\s*flow")
    c_power = find(r"power input")
    c_qty = find(r"total qty") or find(r"\bqty\b")
    if c_total is None or c_sens is None:
        raise ValueError("Vendor sheet: OFFERED total/sensible cooling capacity columns not found")
    flow_div = 1.0 if (c_flow is not None and "l/s" in lab[c_flow]) else 3.6   # m3/h -> L/s
    power_div = 1.0 if (c_power is not None and re.search(r"\bkw\b", lab[c_power])) else 1000.0

    out = []
    for i in range(h + 1, len(rows)):
        r = rows[i] + [None] * (n - len(rows[i]))
        ref = text(r[0])
        total = num(r[c_total])
        if not ref or total is None:
            continue
        f = num(r[c_flow]) if c_flow is not None else None
        p = num(r[c_power]) if c_power is not None else None
        out.append(VendorRow(
            ref=ref, space=text(r[1]), qty=num(r[c_qty]) if c_qty is not None else None,
            total_kw=total, sens_kw=num(r[c_sens]),
            flow_ls=None if f is None else f / flow_div,
            power_kw=None if p is None else p / power_div,
            model=text(r[model]), row=i + 1,
        ))
    return out

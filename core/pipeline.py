"""High level API shared by the CLI and the browser (Pyodide) front end.

    s = Session(fcu_paths, vent_paths, vendor_path=None, vendor_sheet=None)
    s.preview()                 -> JSON-able matching table for manual review
    s.build(options, overrides) -> xlsx bytes
"""
import os
from pathlib import Path

from . import fcu_parser, matcher, vendor_excel, vendor_pdf, vent_parser
from .selection import FIELDS, apply_all
from .xlsx_writer import fill

TEMPLATE = Path(__file__).resolve().parent.parent / "docs" / "assets" / "template.xlsx"


def _read(p):
    with open(p, "rb") as f:
        return f.read()


def vendor_sheets(path):
    return vendor_excel.list_sheets(path)


class Session:
    def __init__(self, fcu_paths, vent_paths=(), vendor_path=None, vendor_sheet=None,
                 pdf_words=None, vendor_min=0.8, vent_min=0.4):
        self.units = []
        for p in fcu_paths:
            self.units += fcu_parser.parse(p)
        if not self.units:
            raise ValueError("No FCU units found: the FCU schedule needs a header row with 'FCU Tag'.")
        if not list(vent_paths):
            raise ValueError("The ventilation calculation is required.")
        self.vent = []
        for p in vent_paths:
            self.vent += vent_parser.parse(p, os.path.basename(str(p)))
        self.vendor = None
        self.vendor_sheet = vendor_sheet
        if pdf_words is not None:
            self.vendor = vendor_pdf.parse_words(pdf_words)
        elif vendor_path:
            if str(vendor_path).lower().endswith(".pdf"):
                self.vendor = vendor_pdf.parse(vendor_path)
            else:
                names = vendor_excel.list_sheets(vendor_path)
                self.vendor_sheet = vendor_sheet or vendor_excel.default_sheet(names)
                self.vendor = vendor_excel.parse(vendor_path, self.vendor_sheet)
        self.vm = matcher.match_vendor(self.units, self.vendor, vendor_min) if self.vendor else None
        self.nm = matcher.match_vent(self.units, self.vent, vent_min) if self.vent else None

    def _vlabel(self, vi):
        v = self.vent[vi]
        f = lambda x: "–" if x is None else f"{x:g}"
        return f"{v.name} · {f(v.people)} people · {f(v.fresh_air)} L/s · {f(v.area)} m²"

    # ---- review table -------------------------------------------------
    def preview(self):
        rows = []
        for i, u in enumerate(self.units):
            r = {"i": i, "ref": u.ref, "zone": u.zone, "sheet": u.sheet, "qty": u.qty,
                 "area": u.area, "total": u.calc_total_kw, "sens": u.calc_sens_kw, "flow": u.flow}
            if self.vm is not None:
                m = self.vm[i]
                r["vendor"] = {
                    "best": m["best"], "score": round(m["score"], 2),
                    "cands": [{"i": vi, "s": round(s, 2), "label": f"{self.vendor[vi].ref} · {self.vendor[vi].model}"}
                              for vi, s in m["cands"]],
                }
            if self.nm is not None:
                n = self.nm[i]
                r["vent"] = {
                    "amb": n["amb"],
                    "hits": [{"space": h[0], "i": h[1], "s": round(h[2], 2), "label": self._vlabel(h[1])} for h in n["hits"]],
                    "cands": [{"i": vi, "s": round(s, 2), "label": self._vlabel(vi)} for vi, s in n["cands"]],
                }
            rows.append(r)
        return {"units": rows, "vendor_count": len(self.vendor or []), "vent_count": len(self.vent),
                "vendor_sheet": self.vendor_sheet}

    # ---- output -----------------------------------------------------------
    def vendor_for(self, i, overrides):
        if self.vendor is None:
            return None
        ov = (overrides or {}).get(str(i), {}).get("vendor")
        if ov is not None:
            return None if ov < 0 else self.vendor[ov]
        b = self.vm[i]["best"]
        return None if b is None else self.vendor[b]

    def vent_for(self, i, overrides):
        """-> (people, fresh_air) or (None, None)."""
        if not self.vent:
            return None, None
        ov = (overrides or {}).get(str(i), {}).get("vent")
        if ov is not None:
            if ov < 0:
                return None, None
            v = self.vent[ov]
            return v.people, v.fresh_air
        people = fa = None
        for _, vi, _ in self.nm[i]["hits"]:
            v = self.vent[vi]
            if v.people is not None:
                people = (people or 0) + v.people
            if v.fresh_air is not None:
                fa = (fa or 0) + v.fresh_air
        return people, fa

    def build_rows(self, options=None, overrides=None):
        o = options or {}
        sec_pct, sec_dir = o.get("pct", 0), o.get("direction", "increase")
        sec_apply = o.get("apply_to") or {}
        fb_pct, fb_dir = o.get("fallback_pct", 0), o.get("fallback_direction", "increase")
        fb_apply = o.get("fallback_apply_to") or {}
        fb_qty = o.get("qty")
        rows, warnings = [], []
        for i, u in enumerate(self.units):
            people, fa = self.vent_for(i, overrides)
            d = {"D": u.ref, "E": "Chilled Water", "F": "FCU", "G": u.zone, "H": 1,
                 "I": u.area, "J": people, "K": fa,
                 "N": u.calc_total_kw, "O": u.calc_sens_kw, "P": u.flow}
            first = None
            if self.vendor is not None:
                v = self.vendor_for(i, overrides)
                if v is None:
                    warnings.append(f"{u.ref}: no vendor match, first/second selection left empty")
                else:
                    first = {"qty": v.qty if v.qty is not None else u.qty, "total": v.total_kw,
                             "sensible": v.sens_kw, "flow": v.flow_ls, "power": v.power_kw}
            else:
                calc = {"total": u.calc_total_kw, "sensible": u.calc_sens_kw, "flow": u.flow, "power": None}
                first = {"qty": fb_qty, **apply_all(calc, fb_pct, fb_dir, fb_apply)}
            if first:
                d.update({"S": first["qty"], "T": first["total"], "U": first["sensible"],
                          "V": first["flow"], "W": first["power"]})
                sec = apply_all(first, sec_pct, sec_dir, sec_apply)
                d.update({"Y": first["qty"], "Z": sec["total"], "AA": sec["sensible"],
                          "AB": sec["flow"], "AC": sec["power"]})
            rows.append(d)
        return rows, warnings

    def build(self, options=None, overrides=None, template=None):
        rows, warnings = self.build_rows(options, overrides)
        o = options or {}
        tpl = template if template is not None else _read(TEMPLATE)
        data = fill(tpl, rows, header=o.get("header"), hide_empty=o.get("hide_empty", True))
        return data, warnings

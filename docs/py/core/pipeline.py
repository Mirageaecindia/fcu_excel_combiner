"""High level API shared by the CLI and the browser (Pyodide) front end.

    s = Session(fcu_paths, vent_paths=(), vendor=[...], people=[...])
    s.preview()                 -> JSON-able matching table for manual review
    s.build(options, overrides) -> (xlsx bytes, warnings)

vendor items : "file.xlsm", "file.xlsm::plot 04", "file.pdf", {"path", "sheet"} or {"words", "name"}
people items : "report.pdf", {"path"} or {"pages": [text per page], "name"}
"""
import os
from pathlib import Path

from . import fcu_parser, matcher, people_pdf, vendor_excel, vendor_pdf, vent_parser
from .selection import apply_all
from .xlsx_writer import fill

TEMPLATE = Path(__file__).resolve().parent.parent / "docs" / "assets" / "template.xlsx"


def _read(p):
    with open(p, "rb") as f:
        return f.read()


def vendor_sheets(path):
    return vendor_excel.list_sheets(path)


def _items(x):
    if x is None:
        return []
    return [x] if isinstance(x, (str, os.PathLike, dict)) else list(x)


def _vendor_item(it):
    if isinstance(it, dict):
        return it
    path, _, sheet = str(it).partition("::")
    return {"path": path, "sheet": sheet or None}


class Session:
    def __init__(self, fcu_paths, vent_paths=(), vendor=None, people=None, vendor_min=0.8, people_min=0.4):
        self.fcu_warnings = []
        self.units = []
        for p in _items(fcu_paths):
            self.units += fcu_parser.parse(p, warnings=self.fcu_warnings)
        if not self.units:
            raise ValueError("No FCU units found: the FCU schedule needs a header row with 'FCU Tag' "
                             "and the FCU Tag filled on the first row of every unit.")
        self.vent = []
        for p in _items(vent_paths):
            self.vent += vent_parser.parse(p, os.path.basename(str(p)))
        self.pdf_people = []
        for it in _items(people):
            if isinstance(it, dict) and "pages" in it:
                self.pdf_people += people_pdf.parse_pages(it["pages"], it.get("name", ""))
            else:
                self.pdf_people += people_pdf.parse(it["path"] if isinstance(it, dict) else it)
        self.vendor, self.vendor_sheets = None, []
        items = [_vendor_item(i) for i in _items(vendor)]
        if items:
            self.vendor, seen = [], set()
            for it in items:
                if "words" in it:
                    rows = vendor_pdf.parse_words(it["words"])
                elif str(it["path"]).lower().endswith(".pdf"):
                    rows = vendor_pdf.parse(it["path"])
                else:
                    names = vendor_excel.list_sheets(it["path"])
                    sheet = it.get("sheet") or vendor_excel.default_sheet(names)
                    self.vendor_sheets.append(sheet)
                    rows = vendor_excel.parse(it["path"], sheet)
                fresh = {matcher.norm(r.ref) for r in rows}
                self.vendor += [r for r in rows if matcher.norm(r.ref) not in seen]   # earlier file wins
                seen |= fresh
        self.people_all = self.pdf_people + self.vent
        self.vm = matcher.match_vendor(self.units, self.vendor, vendor_min) if self.vendor else None
        self.pm = matcher.match_people(self.units, self.pdf_people, self.vent, people_min) if self.people_all else None

    # ---- review table -------------------------------------------------
    def _plabel(self, i):
        v = self.people_all[i]
        f = lambda x: "–" if x is None else f"{x:g}"
        return f"{v.name} · {f(v.people)} people · {f(v.area)} m² · [{v.source}]"

    def preview(self):
        rows = []
        for i, u in enumerate(self.units):
            r = {"i": i, "ref": u.ref, "zone": u.zone, "sheet": u.sheet, "qty": u.qty,
                 "area": u.area, "total": u.calc_total_kw, "sens": u.calc_sens_kw, "flow": u.flow,
                 "subs": len(u.spaces)}
            if self.vm is not None:
                m = self.vm[i]
                r["vendor"] = {
                    "best": m["best"], "score": round(m["score"], 2),
                    "cands": [{"i": vi, "s": round(s, 2), "label": f"{self.vendor[vi].ref} · {self.vendor[vi].model}"}
                              for vi, s in m["cands"]],
                }
            if self.pm is not None:
                n = self.pm[i]
                r["people"] = {
                    "hits": [{"space": h[0], "i": h[1], "s": round(h[2], 2), "label": self._plabel(h[1]),
                              "p": self.people_all[h[1]].people} for h in n["hits"]],
                    "cands": [{"i": vi, "s": round(s, 2), "label": self._plabel(vi), "p": self.people_all[vi].people}
                              for vi, s in n["cands"]],
                    "value": self.people_for(i, None, None)[0],
                }
            rows.append(r)
        return {"units": rows, "vendor_count": len(self.vendor or []), "people_count": len(self.people_all),
                "pdf_people_count": len(self.pdf_people), "warnings": self.fcu_warnings}

    # ---- values ------------------------------------------------------------
    def vendor_for(self, i, overrides):
        if self.vendor is None:
            return None
        ov = (overrides or {}).get(str(i), {}).get("vendor")
        if ov is not None:
            return None if ov < 0 else self.vendor[ov]
        b = self.vm[i]["best"]
        return None if b is None else self.vendor[b]

    def people_for(self, i, overrides, default):
        """-> (people or None, 'override'|'match'|'default'|None)."""
        ov = (overrides or {}).get(str(i), {}).get("people")
        if ov is not None:
            return (None, None) if ov < 0 else (self.people_all[ov].people, "override")
        if self.pm is not None:
            vals = [self.people_all[h[1]].people for h in self.pm[i]["hits"] if self.people_all[h[1]].people is not None]
            if vals:
                return float(sum(vals)), "match"
        if default is not None:
            return default, "default"
        return None, None

    def build_rows(self, options=None, overrides=None):
        o = options or {}
        sec_pct, sec_dir = o.get("pct", 0), o.get("direction", "increase")
        sec_apply = o.get("apply_to") or {}
        fb_pct, fb_dir = o.get("fallback_pct", 0), o.get("fallback_direction", "increase")
        fb_apply = o.get("fallback_apply_to") or {}
        fb_qty = o.get("qty")
        default_people = o.get("people_default")
        rows, warnings = [], []
        for i, u in enumerate(self.units):
            people, _ = self.people_for(i, overrides, default_people)
            d = {"D": u.ref, "E": "Chilled Water", "F": "FCU", "G": u.zone, "H": 1,
                 "I": u.area, "J": people,                       # K (Outdoor Air) is always left empty
                 "N": u.calc_total_kw, "O": u.calc_sens_kw, "P": u.flow}
            first = None
            if self.vendor is not None:
                v = self.vendor_for(i, overrides)
                if v is None:
                    warnings.append(f"{u.ref} ({u.zone}): no vendor match, first/second selection left empty")
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

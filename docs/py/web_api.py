"""Thin JSON API used by app.js inside Pyodide (files are written to /work by the page)."""
import json

from core import pipeline, vendor_excel

_session = None
TEMPLATE = "/assets/template.xlsx"


def vendor_sheets(path):
    names = pipeline.vendor_sheets(path)
    return json.dumps({"sheets": names, "default": vendor_excel.default_sheet(names)})


def analyze(fcu, vent, vendor, sheet, pdf_words_json, vendor_min, vent_min):
    global _session
    words = json.loads(pdf_words_json) if pdf_words_json else None
    _session = pipeline.Session(
        list(fcu), list(vent), None if words else (vendor or None), sheet or None,
        pdf_words=words, vendor_min=float(vendor_min), vent_min=float(vent_min))
    return json.dumps(_session.preview())


def build(options_json, overrides_json, out_path):
    with open(TEMPLATE, "rb") as f:
        tpl = f.read()
    data, warnings = _session.build(json.loads(options_json), json.loads(overrides_json), template=tpl)
    with open(out_path, "wb") as f:
        f.write(data)
    return json.dumps({"warnings": warnings, "rows": len(_session.units)})

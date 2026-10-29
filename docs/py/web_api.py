"""Thin JSON API used by app.js inside Pyodide (files are written to /work by the page)."""
import json

from core import pipeline, vendor_excel

_session = None
TEMPLATE = "/assets/template.xlsx"


def vendor_sheets(path):
    names = pipeline.vendor_sheets(path)
    return json.dumps({"sheets": names, "default": vendor_excel.default_sheet(names)})


def analyze(fcu, vent, vendor_json, people_json, vendor_min, people_min):
    """vendor_json: [{"path","sheet"} | {"words","name"}], people_json: [{"path"} | {"pages","name"}]."""
    global _session
    _session = pipeline.Session(
        list(fcu), list(vent), json.loads(vendor_json), json.loads(people_json),
        vendor_min=float(vendor_min), people_min=float(people_min))
    return json.dumps(_session.preview())


def build(options_json, overrides_json, out_path):
    with open(TEMPLATE, "rb") as f:
        tpl = f.read()
    data, warnings = _session.build(json.loads(options_json), json.loads(overrides_json), template=tpl)
    with open(out_path, "wb") as f:
        f.write(data)
    return json.dumps({"warnings": warnings, "rows": len(_session.units)})

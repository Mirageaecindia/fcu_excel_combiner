"""Command line front end.

    python cli.py --fcu F.xlsx [F2.xlsx ...] [--vendor D.xlsm::"plot 04" D2.pdf ...]
                  [--people-pdf HAP.pdf ...] [--vent V.xlsx ...] [--default-people 2]
"""
import argparse
import sys
from pathlib import Path

from core.pipeline import Session
from core.selection import FIELDS


def _apply(s):
    chosen = {x.strip() for x in s.split(",") if x.strip()}
    bad = chosen - set(FIELDS)
    if bad:
        raise argparse.ArgumentTypeError(f"unknown field(s) {sorted(bad)}; choose from {FIELDS}")
    return {f: f in chosen for f in FIELDS}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Merge FCU schedule, people report and vendor report "
                                             "into the thermal load summary workbook.")
    ap.add_argument("--fcu", nargs="+", required=True, help="FCU schedule .xlsx (one or more; FCU Tag must be filled)")
    ap.add_argument("--people-pdf", nargs="*", default=[], help="HAP 'Space Input Data' PDF(s): No. of People per space")
    ap.add_argument("--default-people", type=float, help="No. of People for rows without a match (omit = leave empty)")
    ap.add_argument("--vent", nargs="*", default=[], help="optional ventilation calculation .xlsx (people fallback)")
    ap.add_argument("--vendor", nargs="*", default=[],
                    help="vendor deviation report(s): .xlsx/.xlsm (optionally file::sheet) or .pdf")
    ap.add_argument("--pct", type=float, default=0, help="Second Selection: percentage change")
    ap.add_argument("--dir", choices=["increase", "decrease"], default="increase")
    ap.add_argument("--apply", type=_apply, default=_apply(",".join(FIELDS)),
                    help="Second Selection columns to change (comma list of total,sensible,flow,power; default all)")
    ap.add_argument("--fallback-pct", type=float, default=0, help="no vendor file: First Selection = Calculated load +/- this %%")
    ap.add_argument("--fallback-dir", choices=["increase", "decrease"], default="increase")
    ap.add_argument("--fallback-apply", type=_apply, default=_apply(",".join(FIELDS)))
    ap.add_argument("--qty", type=float, help="no vendor file: First Selection quantity written to every row")
    ap.add_argument("--plot"), ap.add_argument("--project"), ap.add_argument("--consultant")
    ap.add_argument("--keep-empty-rows", action="store_true", help="do not hide unused template rows")
    ap.add_argument("--vendor-min", type=float, default=0.8, help="minimum fuzzy score for automatic vendor match")
    ap.add_argument("--people-min", type=float, default=0.4, help="minimum score for the ventilation fallback match")
    ap.add_argument("-o", "--out", default="output/AC-Unit-Schedule.xlsx")
    a = ap.parse_args(argv)

    if not a.vendor and a.qty is None:
        ap.error("without --vendor you must give --qty (First Selection quantity)")
    s = Session(a.fcu, a.vent, a.vendor, a.people_pdf, vendor_min=a.vendor_min, people_min=a.people_min)
    options = {
        "pct": a.pct, "direction": a.dir, "apply_to": a.apply,
        "fallback_pct": a.fallback_pct, "fallback_direction": a.fallback_dir,
        "fallback_apply_to": a.fallback_apply, "qty": a.qty, "people_default": a.default_people,
        "header": {"plot": a.plot, "project": a.project, "consultant": a.consultant},
        "hide_empty": not a.keep_empty_rows,
    }
    data, warnings = s.build(options)
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    print(f"{len(s.units)} units, {len(s.pdf_people)} HAP spaces, {len(s.vent)} vent spaces, "
          f"{len(s.vendor) if s.vendor else 0} vendor rows -> {out}")
    for w in s.fcu_warnings + warnings:
        print("warning:", w, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())

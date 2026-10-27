"""Command line front end:  python cli.py --fcu F.xlsx --vent V1.xlsx V2.xlsx [--vendor D.xlsm|D.pdf]"""
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
    ap = argparse.ArgumentParser(description="Merge FCU schedule, ventilation calculation and vendor report "
                                             "into the thermal load summary workbook.")
    ap.add_argument("--fcu", nargs="+", required=True, help="FCU schedule .xlsx (one or more)")
    ap.add_argument("--vent", nargs="*", default=[], help="ventilation calculation .xlsx (one per building)")
    ap.add_argument("--vendor", help="vendor deviation report (.xlsx/.xlsm/.pdf); optional")
    ap.add_argument("--vendor-sheet", help="sheet to read from the vendor workbook (default: plot 04 / first)")
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
    ap.add_argument("--vent-min", type=float, default=0.75, help="minimum fuzzy score for automatic vent match")
    ap.add_argument("-o", "--out", default="output/AC-Unit-Schedule.xlsx")
    a = ap.parse_args(argv)

    if not a.vendor and a.qty is None:
        ap.error("without --vendor you must give --qty (First Selection quantity)")
    s = Session(a.fcu, a.vent, a.vendor, a.vendor_sheet, vendor_min=a.vendor_min, vent_min=a.vent_min)
    options = {
        "pct": a.pct, "direction": a.dir, "apply_to": a.apply,
        "fallback_pct": a.fallback_pct, "fallback_direction": a.fallback_dir,
        "fallback_apply_to": a.fallback_apply, "qty": a.qty,
        "header": {"plot": a.plot, "project": a.project, "consultant": a.consultant},
        "hide_empty": not a.keep_empty_rows,
    }
    data, warnings = s.build(options)
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    print(f"{len(s.units)} units, {len(s.vent)} vent spaces, "
          f"{len(s.vendor) if s.vendor else 0} vendor rows -> {out}")
    for w in warnings:
        print("warning:", w, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())

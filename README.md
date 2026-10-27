# FCU Excel Combiner

Merges an **FCU schedule**, **ventilation calculation(s)** and an optional **vendor deviation report** (xlsx/xlsm/pdf)
into the *Summary of Thermal Load Calculation* workbook (the reference file is used as the template, so logos,
button, dropdowns and sheet protection are kept; rows 1-12 stay locked, data rows are editable).

## Use in the browser (GitHub Pages)
Publish the `docs/` folder (Settings > Pages > Deploy from branch > `/docs`). Python runs in the browser via Pyodide;
nothing is uploaded. Steps: choose files, set the percentages, **Analyze**, review matches, **Generate Excel**.

## Use locally
    pip install -r requirements.txt
    python cli.py --fcu input/fcu/*.xlsx --vent input/vent/*.xlsx --vendor input/vendor/*.xlsm --pct 5 -o output/out.xlsx
    # no vendor file: First Selection = calculated load +/- %, quantity from --qty
    python cli.py --fcu ... --vent ... --qty 1 --fallback-pct 10 --pct 5 --dir decrease

## Rules
- Calculated AC load: FCU schedule (kW x unit qty, air flow = sum of the unit's spaces).
- People / Outdoor air: ventilation sheets with a "Proposed no. of people" column, matched by space name.
  Only unambiguous matches are filled automatically; the rest can be chosen in the review table.
- First Selection: vendor OFFERED values (flow m3/h -> L/s, power W -> kW). Second Selection: First Selection
  +/- percentage on the chosen columns; quantity is always copied.
- Check figures (kW/m2, (L/s)/m2), occupant density and DCV are Excel formulas from the template.
- Voltages are left empty. The sheet password is the reference file's.

## Maintenance
After editing `core/` run `python build_pages.py` (copies it into `docs/py`). Tests: `pytest`.
Vendor PDFs need selectable text laid out like the vendor template (31 columns); this is untested on real PDFs.

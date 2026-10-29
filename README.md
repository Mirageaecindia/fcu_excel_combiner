# DM_extrator

Merges **FCU schedule(s)**, a **people report** (HAP *Space Input Data* PDFs) and an optional **vendor deviation report**
(xlsx/xlsm/pdf) into the *Summary of Thermal Load Calculation* workbook. The reference file is the template, so logos,
button, dropdowns and sheet protection are kept; rows 1-12 stay locked, data rows are editable.

## Prerequisites
- **FCU Tag** must be filled on the first row of every unit. A row with an empty tag is a sub part of the tag above it:
  its area and air flow are added to that unit (FCU-03 with sub rows 124 + 11 L/s = one unit, 135 L/s).
  The Units Reference column is the FCU Tag. Rows above the first tag are ignored; rows with load values but no tag are
  merged into the unit above and reported as a warning.
- **No. of People** comes from the HAP *Space Input Data* PDF(s): the *Occupancy ... Person* value under each space
  heading (`#001-BD1-B-Pump Room`). Spaces are matched by code number + building (BD#), ranked by floor area and name.
  Without a PDF you can give a default value or leave it empty. An (optional) ventilation calculation is used as a
  fallback for spaces that have no PDF match.
- **Outdoor Air** is always left empty.

## Use in the browser (GitHub Pages)
Publish the `docs/` folder (Settings > Pages > Deploy from branch > `/docs`; `docs/.nojekyll` is required).
Python runs in the browser via Pyodide; nothing is uploaded. Every input accepts several files (Add files / x to remove).
Steps: add files, set the percentages, **Analyze**, review the matches, **Generate Excel**.

## Use locally
    pip install -r requirements.txt
    python cli.py --fcu input/fcu/*.xlsx --vendor input/vendor/x.xlsm::"plot 04" --people-pdf hap1.pdf hap2.pdf --default-people 2 --pct 5
    # no vendor file: First Selection = calculated load +/- %, quantity from --qty
    python cli.py --fcu ... --qty 1 --fallback-pct 10 --pct 5 --dir decrease

## Other rules
- Calculated AC load: FCU schedule (kW x unit qty, air flow = sum of the unit's sub rows).
- First Selection: vendor OFFERED values (flow m3/h -> L/s, power W -> kW). Second Selection: First Selection
  +/- percentage on the chosen columns; quantity is always copied.
- Check figures (kW/m2, (L/s)/m2), occupant density and DCV are Excel formulas from the template.
- Data cells are white in the output (the reference colours were only a guide); voltages are left empty.
- The sheet password is the reference file's.

## Maintenance
After editing `core/` run `python build_pages.py` (copies it into `docs/py`). Tests: `pytest`.
Vendor and people PDFs need selectable text (scans are not supported).

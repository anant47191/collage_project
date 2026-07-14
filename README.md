# Highlighted-Column Auto-Dashboard (prototype)

Upload an Excel file with certain column headers highlighted in **yellow**,
and get an instant dashboard built from just those columns — every time,
automatically, no matter what the columns are.

Built and tested against a real manufacturing/quality-inspection export
(defect tracking: Minor/Major/HR flags, Shift, Operator, Defect Code, etc.),
but it's generic — it works with any spreadsheet as long as the headers you
care about are highlighted yellow.

## Run it locally

```bash
cd highlighted-dashboard
pip install -r requirements.txt
python app.py
```

Then open **http://localhost:8000**, drag in your `.xlsx` file, and the
dashboard builds itself.

## How it decides what to show

1. **Finds the highlighted columns.** It scans row 1 of the first sheet for
   header cells filled yellow, and keeps only those columns — by position,
   so it's not confused if you have two columns both named "Shift" and only
   one is highlighted.
2. **Classifies each one automatically:**
   - **Date** — a column that's a real date, or whose name contains "date"
     → used for a trend-over-time chart.
   - **Measure** — a numeric column whose name isn't an ID/code/label (e.g.
     `Minor`, `Major`, `HR`) → summed for KPI cards and charts.
   - **Dimension** — everything else (`Shift`, `Operator`, `Defect Code`,
     `Week Code`, etc.) → used to group and break down the measures.
3. **Builds the dashboard:**
   - KPI cards: total record count + total of each measure.
   - A trend chart of the measures over time (if a date column was found).
   - One breakdown chart per dimension column, showing the measures summed
     by that dimension's top categories (e.g. "Minor / Major / HR by
     Operator") — or, if there are no measures, a simple count breakdown.

This means the exact same file structure with different columns highlighted
will produce a different dashboard next time — you control what's shown just
by changing which headers are yellow.

## If nothing shows up

- Make sure the fill color is applied to the **header cell itself** (row 1),
  not the data cells below it.
- Only `.xlsx` / `.xlsm` files are supported (not `.csv`, since CSV has no
  cell formatting).
- If a highlighted column has only one distinct value in the whole file
  (like a column that's always "1"), it's skipped — there's nothing to chart.

## Extending this

- **Multiple sheets**: currently reads the first sheet only; easy to add a
  sheet picker if your files have more than one.
- **Persistence**: save each upload's dashboard so you can compare weeks/months.
- **Scheduled ingestion**: point this at a folder or email inbox so new
  exports get dashboarded automatically without manual upload.
- **Deploy**: containerize with Docker and host it so your team can use it
  without running it locally.

## Building this further with an AI coding agent

This project is clean enough to hand to **Claude Code** for any of the
above — e.g. "add a sheet picker" or "save dashboard history to SQLite."
It can work directly in this folder, run the app, and iterate.

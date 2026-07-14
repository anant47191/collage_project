"""
Highlighted-Column Auto-Dashboard — backend
Reads an Excel file, finds header columns highlighted in yellow, and builds
a dashboard automatically from just those columns.

Run with:  python app.py
Then open: http://localhost:8000
"""

import io
import re

import openpyxl
import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="Highlighted-Column Dashboard Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------------------------
# Step 1: find which header cells (row 1) are highlighted yellow.
# We match by column position, not name, because Excel exports often reuse
# the same header name (e.g. "Shift", "Operator") in more than one place.
# ---------------------------------------------------------------------------

def is_yellow(cell):
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    fg = fill.fgColor
    if fg is None:
        return False
    rgb = fg.rgb
    if isinstance(rgb, str) and len(rgb) == 8:
        try:
            r, g, b = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
            return r > 200 and g > 200 and b < 120
        except ValueError:
            return False
    return False


def get_highlighted_columns(raw_bytes):
    """Returns a list of (1-based column index, header name) for yellow header cells
    in row 1 of the first worksheet."""
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    ws = wb.worksheets[0]
    cols = []
    for cell in ws[1]:
        if cell.value is None:
            continue
        if is_yellow(cell):
            cols.append((cell.column, str(cell.value).strip()))
    return cols


# ---------------------------------------------------------------------------
# Step 2: classify each highlighted column as a date, a measure (something to
# sum/count, e.g. defect flags), or a dimension (something to group by, e.g.
# Shift, Operator, Defect Code) — using the column name and its data.
# ---------------------------------------------------------------------------

DIMENSION_TOKENS = {
    "code", "no", "number", "id", "qr", "qrcode", "barcode", "mould",
    "operator", "shift", "week", "plant", "area", "spec", "seal",
}


def classify_columns(df):
    date_cols, measure_cols, dim_cols = [], [], []
    for col in df.columns:
        s = df[col]
        lname = col.lower()
        tokens = set(re.split(r"[^a-z0-9]+", lname))

        if pd.api.types.is_datetime64_any_dtype(s):
            date_cols.append(col)
            continue

        if "date" in tokens and s.dtype == object:
            parsed = pd.to_datetime(s, errors="coerce")
            if parsed.notna().mean() > 0.8:
                df[col] = parsed
                date_cols.append(col)
                continue

        is_dimension_name = bool(tokens & DIMENSION_TOKENS)
        if pd.api.types.is_numeric_dtype(s) and not is_dimension_name:
            measure_cols.append(col)
        else:
            dim_cols.append(col)
    return date_cols, measure_cols, dim_cols


# ---------------------------------------------------------------------------
# Step 3: build KPI cards and charts from the classified columns.
# ---------------------------------------------------------------------------

def build_kpis(df, measure_cols):
    kpis = {"Total Records": int(len(df))}
    for col in measure_cols:
        total = df[col].sum(skipna=True)
        kpis[f"Total {col}"] = round(float(total), 2) if pd.notna(total) else 0
    return kpis


def build_trend_chart(df, date_col, measure_cols):
    d = df.dropna(subset=[date_col]).copy()
    if d.empty:
        return None
    d["__date__"] = pd.to_datetime(d[date_col]).dt.date

    if measure_cols:
        grouped = d.groupby("__date__")[measure_cols].sum(numeric_only=True).sort_index()
        series = [{"name": m, "values": [round(float(v), 2) for v in grouped[m]]} for m in measure_cols]
    else:
        grouped = d.groupby("__date__").size().sort_index()
        series = [{"name": "Count", "values": [int(v) for v in grouped.values]}]

    labels = [x.strftime("%b %d") for x in grouped.index]
    return {"type": "line", "title": f"Trend Over Time ({date_col})", "labels": labels, "series": series}


def build_dimension_chart(df, dim_col, measure_cols, top_n=10):
    if df[dim_col].nunique(dropna=True) <= 1:
        return None

    if measure_cols:
        grouped = df.groupby(dim_col)[measure_cols].sum(numeric_only=True)
        grouped["__total__"] = grouped.sum(axis=1)
        grouped = grouped.sort_values("__total__", ascending=False).head(top_n)
        labels = [str(x) for x in grouped.index]
        series = [{"name": m, "values": [round(float(v), 2) for v in grouped[m]]} for m in measure_cols]
        title = f"{' / '.join(measure_cols)} by {dim_col}"
    else:
        counts = df[dim_col].value_counts(dropna=True).head(top_n)
        if counts.empty:
            return None
        labels = [str(x) for x in counts.index]
        series = [{"name": "Count", "values": [int(v) for v in counts.values]}]
        title = f"{dim_col} Breakdown"

    return {"type": "bar", "title": title, "labels": labels, "series": series}


@app.post("/api/analyze")
async def analyze(file: UploadFile = File(...)):
    raw = await file.read()
    name = file.filename or ""

    if not name.lower().endswith((".xlsx", ".xlsm", ".xltx")):
        raise HTTPException(400, "Please upload an Excel file (.xlsx) with highlighted header cells.")

    try:
        highlighted = get_highlighted_columns(raw)
    except Exception as e:
        raise HTTPException(400, f"Could not read the file: {e}")

    if not highlighted:
        raise HTTPException(
            400,
            "No yellow-highlighted header cells were found in row 1. "
            "Highlight the column headers you want in the dashboard and re-upload."
        )

    try:
        df_full = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=0)
    except Exception as e:
        raise HTTPException(400, f"Could not read the file: {e}")

    if df_full.empty:
        raise HTTPException(400, "The uploaded file has no data rows.")

    # Select highlighted columns by position (handles duplicate header names).
    idxs = [idx - 1 for idx, _ in highlighted if 0 <= idx - 1 < df_full.shape[1]]
    df = df_full.iloc[:, idxs].copy()

    # De-duplicate column labels for display (e.g. two "Shift" columns).
    final_names, seen = [], {}
    for _, n in highlighted:
        seen[n] = seen.get(n, 0) + 1
        final_names.append(n if seen[n] == 1 else f"{n} ({seen[n]})")
    df.columns = final_names

    date_cols, measure_cols, dim_cols = classify_columns(df)

    kpis = build_kpis(df, measure_cols)

    charts = []
    for dcol in date_cols[:1]:  # one trend chart is enough
        chart = build_trend_chart(df, dcol, measure_cols)
        if chart:
            charts.append(chart)

    for dim in dim_cols[:8]:
        chart = build_dimension_chart(df, dim, measure_cols)
        if chart:
            charts.append(chart)

    if not charts:
        charts.append({
            "type": "bar",
            "title": "Row Preview",
            "labels": [str(i) for i in df.index[:20]],
            "series": [{"name": df.columns[0], "values": [0] * min(20, len(df))}],
        })

    return {
        "highlighted_columns": final_names,
        "column_roles": {
            "date": date_cols,
            "measures": measure_cols,
            "dimensions": dim_cols,
        },
        "row_count": int(len(df)),
        "kpis": kpis,
        "charts": charts,
    }


# Serve the frontend
app.mount("/", StaticFiles(directory="static", html=True), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False)

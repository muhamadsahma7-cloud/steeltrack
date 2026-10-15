"""Excel in/out for SteelTrack: the master database template
(master_database_final.xlsx, sheet 'Master Database', 25 columns) and the
report workbooks."""

from __future__ import annotations

import io
from datetime import date, datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# template header -> parts column. The four '(kg)' columns are not stored:
# a stage's kg is the part weight once that stage has a date, else 0.
TEMPLATE = [
    ("Priority", "priority"),
    ("Work Order", "work_order"),
    ("Assembly Mark", "assembly_mark"),
    ("Sub Assembly", "sub_assembly"),
    ("Part Mark", "part_mark"),
    ("No.", "qty"),
    ("Name", "name"),
    ("Profile", "profile"),
    ("kg/m", "kg_per_m"),
    ("Length (mm)", "length_mm"),
    ("Weight (kg)", "weight_kg"),
    ("Profile 2", "profile2"),
    ("Grade", "grade"),
    ("Remark", "remark"),
    ("Painting requirement", "painting_req"),
    ("FIT UP (kg)", None),
    ("FIT UP Date", "fitup_date"),
    ("WELDING (kg)", None),
    ("WELDING Date", "welding_date"),
    ("BLASTING & PAINTING (kg)", None),
    ("BLASTING & PAINTING Date", "paint_date"),
    ("BLASTING & PAINTING D.O. No.", "paint_do"),
    ("SEND TO SITE (kg)", None),
    ("SEND TO SITE Date", "site_date"),
    ("SEND TO SITE D.O. No.", "site_do"),
]
_KG_OF = {"FIT UP (kg)": "fitup_date", "WELDING (kg)": "welding_date",
          "BLASTING & PAINTING (kg)": "paint_date", "SEND TO SITE (kg)": "site_date"}
_NUMFMT = {"Priority": "0", "No.": "0", "kg/m": "#,##0.000", "Length (mm)": "#,##0.0",
           "Weight (kg)": "#,##0.00", **{k: "#,##0.00" for k in _KG_OF}}

REQUIRED = ["Work Order", "Assembly Mark", "Part Mark", "Weight (kg)"]
_TEXT = ["work_order", "assembly_mark", "sub_assembly", "part_mark", "name", "profile",
         "profile2", "grade", "remark", "painting_req", "paint_do", "site_do"]
_DATES = ["fitup_date", "welding_date", "paint_date", "site_date"]


def _clean_text(v) -> str | None:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return None if s.lower() in ("", "nan", "none", "nat") else s


def _clean_date(v) -> str | None:
    """Excel gives real dates, 'YYYY-MM-DD' text or blanks -> 'YYYY-MM-DD' / None."""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    s = _clean_text(v)
    if s is None:
        return None
    try:
        return pd.to_datetime(s, dayfirst=False).strftime("%Y-%m-%d")
    except Exception:
        return s


def parse_master(file) -> pd.DataFrame:
    """Template workbook -> DataFrame with parts columns. Raises ValueError
    with a readable message when the sheet doesn't look like the template."""
    xl = pd.ExcelFile(file)
    sheet = "Master Database" if "Master Database" in xl.sheet_names else xl.sheet_names[0]
    # dtype=object keeps '001' / '012' leading zeros on WO and D.O. numbers
    raw = xl.parse(sheet, dtype=object)
    raw.columns = [str(c).strip() for c in raw.columns]
    missing = [c for c in REQUIRED if c not in raw.columns]
    if missing:
        raise ValueError(f"Sheet '{sheet}' is missing column(s): {', '.join(missing)}")
    out = pd.DataFrame(index=raw.index)
    for head, col in TEMPLATE:
        if col is None:
            continue
        out[col] = raw[head] if head in raw.columns else None
    for c in _TEXT:
        out[c] = out[c].map(_clean_text)
    # Excel may hand back WO 1 for '001' when the cell is numeric
    out["work_order"] = out["work_order"].map(
        lambda s: s[:-2] if s and s.endswith(".0") and s[:-2].isdigit() else s)
    for c in ("paint_do", "site_do"):
        out[c] = out[c].map(lambda s: s[:-2] if s and s.endswith(".0") and s[:-2].isdigit() else s)
    for c in _DATES:
        out[c] = out[c].map(_clean_date)
    for c in ("priority", "qty"):
        out[c] = pd.to_numeric(out[c], errors="coerce").round().astype("Int64")
    for c in ("kg_per_m", "length_mm", "weight_kg"):
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out = out[out["assembly_mark"].notna()].reset_index(drop=True)
    out["weight_kg"] = out["weight_kg"].fillna(0.0)
    return out


def _write_sheet(ws, df: pd.DataFrame, numfmt: dict | None = None, widths: dict | None = None):
    hdr_fill = PatternFill("solid", fgColor="1F3A5F")
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in df.itertuples(index=False):
        ws.append([None if (v is None or (not isinstance(v, str) and pd.isna(v))) else
                   (v.item() if hasattr(v, "item") else v) for v in row])
    numfmt = numfmt or {}
    for j, col in enumerate(df.columns, 1):
        letter = get_column_letter(j)
        if col in numfmt:
            for cell in ws[letter][1:]:
                cell.number_format = numfmt[col]
        w = (widths or {}).get(col) or min(max(len(str(col)), 8) + 2, 40)
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32


def master_df(parts: pd.DataFrame) -> pd.DataFrame:
    """parts -> the 25-column template layout."""
    out = pd.DataFrame()
    for head, col in TEMPLATE:
        if col is None:
            dcol = _KG_OF[head]
            done = parts[dcol].fillna("").astype(str).str.strip() != ""
            out[head] = parts["weight_kg"].where(done, 0.0).astype(float)
        else:
            out[head] = parts[col]
    return out


def master_xlsx(parts: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Master Database"
    _write_sheet(ws, master_df(parts), _NUMFMT)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def workbook(sheets: dict[str, pd.DataFrame], numfmt: dict | None = None,
             title_rows: dict[str, list[str]] | None = None) -> bytes:
    """Generic multi-sheet export. title_rows puts a few heading lines above a
    sheet's table (report name, period, generated time)."""
    wb = Workbook()
    wb.remove(wb.active)
    for name, df in sheets.items():
        ws = wb.create_sheet(name[:31])
        _write_sheet(ws, df, numfmt or {})
        heads = (title_rows or {}).get(name)
        if heads:
            ws.insert_rows(1, len(heads) + 1)
            for i, t in enumerate(heads, 1):
                ws.cell(i, 1, t).font = Font(bold=(i == 1), size=14 if i == 1 else 10)
            ws.freeze_panes = f"A{len(heads) + 3}"
            ws.auto_filter.ref = None
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

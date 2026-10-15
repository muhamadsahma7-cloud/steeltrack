"""Excel in/out for SteelTrack.

Master database layout = the 'Progress' sheet of MASTER DATABASE 722-1.xlsx
(26 columns, one row per part of ONE copy of an assembly):

    WEIGHT        weight of one piece      -> row weight = No. x WEIGHT
    TOTAL WEIGHT  heads a block of rows (down to the next TOTAL WEIGHT):
                  block weight x number of identical copies. A block is one
                  sub assembly (RFT6-1) or a whole assembly listed sub by sub
                  (B132-1..10)        -> asm_qty = TOTAL / block weight, stored
                  on every part row = copies of that row's sub assembly
    NAEC Scope    NO = not fabricated by NAEC (kept, left out of progress)
    Welded?       NO = a single piece (purlin, plain beam): no fit-up / welding
    Painting status  Yes = must be blasted & painted before site
"""

from __future__ import annotations

import io
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# (header, parts column). Progress columns are filled from the progress table.
HEADERS = [
    ("Work order no", "work_order"), ("Priority no", "priority"), ("NAEC Scope", "scope"),
    ("Welded?", "welded"), ("Area", "area"), ("Assembly Mark", "assembly_mark"), ("Sub Assembly", "sub_assembly"),
    ("Part Mark", "part_mark"), ("No.", "qty"), ("Name", "name"), ("Profile", "profile"),
    ("MATERIAL", "grade"), ("Length (Mm)", "length_mm"), ("AREA", "surface_area"),
    ("WEIGHT", "unit_weight"), ("TOTAL WEIGHT", "total_weight"),
    ("Painting status", "painting_req"), ("Remark", "remark"),
]
# stage -> (kg header, date header)
STAGE_HEADS = {
    "fitup": ("Fit Up (Kg)", "Fit Up Date"),
    "welding": ("Welding (Kg)", "Welding Date"),
    "paint": ("Blasting & Painting (Kg)", "Blasting & Painting Date"),
    "site": ("Send To Site (Kg)", "Send To Site Date"),
}
DO_HEAD = "DO no."
ALL_HEADERS = ([h for h, _ in HEADERS] + [h for p in STAGE_HEADS.values() for h in p]
               + [DO_HEAD])

REQUIRED = ["Work order no", "Assembly Mark", "Part Mark", "No.", "WEIGHT"]
_TEXT = ["work_order", "scope", "welded", "area", "assembly_mark", "sub_assembly", "part_mark", "name",
         "profile", "grade", "painting_req", "remark"]


def _clean_text(v) -> str | None:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return None if s.lower() in ("", "nan", "none", "nat") else s


def _clean_date(v) -> str | None:
    """Real dates, Excel serial numbers (46303) or text -> 'YYYY-MM-DD'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d")
    s = _clean_text(v)
    if s is None:
        return None
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%Y-%m-%d")
    except Exception:
        return None


def _map_headers(cols) -> dict[str, str]:
    """sheet header -> canonical header. 'Area' (zone) and 'AREA' (surface
    m2) differ only by case, so AREA in capitals is matched exactly first."""
    canon = {h.lower(): h for h in ALL_HEADERS}
    out = {}
    for c in cols:
        s = " ".join(str(c).split())
        if s == "AREA" or s.upper() == "AREA.1":
            out[c] = "AREA"
        elif s.lower() == "area":
            out[c] = "Area"
        elif s.lower() in canon:
            out[c] = canon[s.lower()]
    return out


def parse_master(file) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Master workbook -> (parts, progress, warnings). Raises ValueError with
    a readable message when the sheet doesn't look like the template."""
    xl = pd.ExcelFile(file)
    sheet = "Progress" if "Progress" in xl.sheet_names else xl.sheet_names[0]
    raw = xl.parse(sheet, dtype=object)      # object keeps '001' and serial dates as-is
    raw = raw.rename(columns=_map_headers(raw.columns))
    missing = [c for c in REQUIRED if c not in raw.columns]
    if missing:
        raise ValueError(f"Sheet '{sheet}' is missing column(s): {', '.join(missing)}")
    for h in ALL_HEADERS:
        if h not in raw.columns:
            raw[h] = None
    raw = raw[raw["Assembly Mark"].map(_clean_text).notna()].reset_index(drop=True)

    p = pd.DataFrame({col: raw[h] for h, col in HEADERS})
    for c in _TEXT:
        p[c] = p[c].map(_clean_text)
    for c in ("priority", "qty"):
        p[c] = pd.to_numeric(p[c], errors="coerce").round().astype("Int64")
    p["qty"] = p["qty"].fillna(1)
    for c in ("length_mm", "surface_area", "unit_weight", "total_weight"):
        p[c] = pd.to_numeric(p[c], errors="coerce")
    p["unit_weight"] = p["unit_weight"].fillna(0.0)
    p["weight_kg"] = p["qty"].astype(float) * p["unit_weight"]
    p["scope"] = p["scope"].fillna("YES").str.upper()
    p["welded"] = p["welded"].fillna("YES").str.upper()      # older sheets have no Welded?
    p["unit"] = p["sub_assembly"].fillna(p["assembly_mark"])

    # copies = TOTAL WEIGHT / weight of the block of rows it heads
    warn: list[str] = []
    p["_blk"] = (p["total_weight"].notna()
                 | p["assembly_mark"].ne(p["assembly_mark"].shift())).cumsum()
    g = p.groupby("_blk").agg(copy=("weight_kg", "sum"), tw=("total_weight", "first"),
                              first=("unit", "first"), last=("unit", "last"),
                              wo=("work_order", "first"))
    q = (g["tw"] / g["copy"].where(g["copy"] > 0)).round()
    g["asm_qty"] = q.fillna(1).clip(lower=1).astype(int)
    off = g[(g["tw"].notna()) & (g["copy"] > 0)
            & ((g["tw"] - g["copy"] * g["asm_qty"]).abs() > 0.05 * g["tw"])]
    for _, r in off.iterrows():
        who = r["first"] if r["first"] == r["last"] else f"{r['first']} … {r['last']}"
        warn.append(f"{who} (WO {r['wo']}): TOTAL WEIGHT {r['tw']:,.1f} kg is not a whole "
                    f"multiple of its parts ({r['copy']:,.1f} kg); "
                    f"taken as {int(r['asm_qty'])} cop{'y' if r['asm_qty'] == 1 else 'ies'}.")
    p["asm_qty"] = p["_blk"].map(g["asm_qty"])
    p = p.drop(columns=["_blk"])
    # one sub assembly must have one copy count
    mixed = p.groupby(["work_order", "assembly_mark", "unit"])["asm_qty"].nunique()
    for k in mixed[mixed > 1].index:
        warn.append(f"Sub assembly {k[2]} (WO {k[0]}) spans TOTAL WEIGHT blocks with "
                    "different copy counts; the highest is used.")
    p["asm_qty"] = p.groupby(["work_order", "assembly_mark", "unit"])["asm_qty"].transform("max")

    # progress already recorded in the sheet -> one booking per unit + stage
    paint = p["painting_req"].fillna("").str.upper().isin(["YES", "Y"])
    weld = p["welded"] != "NO"
    rows = []
    for s, (kh, dh) in STAGE_HEADS.items():
        d = raw[dh].map(_clean_date)
        kg = pd.to_numeric(raw[kh], errors="coerce")
        has = d.notna() | (kg > 0)
        if not has.any():
            continue
        # copies done per part row: from its kg when given, else all copies
        per = (kg / p["weight_kg"].where(p["weight_kg"] > 0)).round()
        per = per.where(kg > 0, p["asm_qty"].where(d.notna(), 0)).fillna(0)
        per = per.clip(upper=p["asm_qty"]).astype(int)
        t = p[["work_order", "assembly_mark", "unit"]].assign(
            done=per.where(has, 0), date=d,
            do_no=raw[DO_HEAD].map(_clean_text) if s == "site" else None,
            applies=paint if s == "paint" else weld if s in ("fitup", "welding") else True)
        t = t[t["applies"]]
        for k, gg in t.groupby(["work_order", "assembly_mark", "unit"], dropna=False):
            n = int(gg["done"].min())             # every part of the unit must be there
            if n <= 0:
                continue
            dos = gg["do_no"].dropna()
            rows.append({"work_order": k[0], "assembly_mark": k[1], "unit": k[2],
                         "stage": s, "qty": n, "date": gg["date"].dropna().max(),
                         "do_no": dos.mode().iloc[0] if len(dos) else None,
                         "done_by": "from Excel"})
    prog = pd.DataFrame(rows, columns=["work_order", "assembly_mark", "unit", "stage", "qty",
                                       "date", "do_no", "done_by"])
    prog = _backfill(prog, p, paint, weld)
    p = p.drop(columns=["unit"])
    return p, prog, warn


def _backfill(prog: pd.DataFrame, p: pd.DataFrame, paint: pd.Series,
              weld: pd.Series) -> pd.DataFrame:
    """A unit recorded at a later stage (e.g. sent to site) went through the
    earlier ones too, even if the sheet has no date for them. Add those
    earlier stages with NO date: they count in the totals but not in any
    day's output."""
    if prog.empty:
        return prog
    unit_req = (p.assign(_p=paint, _w=weld)
                .groupby(["work_order", "assembly_mark", "unit"], dropna=False)[["_p", "_w"]]
                .any())
    have = {(r.work_order, r.assembly_mark, r.unit, r.stage): r.qty
            for r in prog.itertuples()}
    extra = []
    for (wo, am, un), r in unit_req.iterrows():
        chain = (["fitup", "welding"] if r["_w"] else []) + (["paint"] if r["_p"] else []) + ["site"]
        top = 0
        for s in reversed(chain):
            top = max(top, have.get((wo, am, un, s), 0))
            missing = top - have.get((wo, am, un, s), 0)
            if missing > 0:
                extra.append({"work_order": wo, "assembly_mark": am, "unit": un, "stage": s,
                              "qty": missing, "date": None, "do_no": None,
                              "done_by": "from Excel (no date)"})
    return pd.concat([prog, pd.DataFrame(extra)], ignore_index=True) if extra else prog


# ------------------------------------------------------------------ export
_HDR_FILL = PatternFill("solid", fgColor="1F3A5F")


def _write_sheet(ws, df: pd.DataFrame, numfmt: dict | None = None):
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = _HDR_FILL
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
        ws.column_dimensions[letter].width = min(max(len(str(col)), 8) + 3, 40)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32


def master_df(parts: pd.DataFrame) -> pd.DataFrame:
    """parts (with per-stage d_<stage> done copies, last_<stage> dates and
    site_dos) -> the 26-column master layout. Stage kg = row weight x copies done."""
    out = pd.DataFrame({h: parts[c] for h, c in HEADERS})
    for s, (kh, dh) in STAGE_HEADS.items():
        done = parts[f"d_{s}"].fillna(0)
        out[kh] = (parts["weight_kg"] * done).where(done > 0)
        out[dh] = pd.to_datetime(parts[f"last_{s}"].where(done > 0), errors="coerce")
    out[DO_HEAD] = parts["site_dos"].replace("", None)
    return out[ALL_HEADERS]


def master_xlsx(parts: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Progress"
    fmt = {"AREA": "0.0", "WEIGHT": "0.0", "TOTAL WEIGHT": "#,##0.0",
           **{kh: "#,##0.0" for kh, _ in STAGE_HEADS.values()},
           **{dh: "dd/mm/yyyy" for _, dh in STAGE_HEADS.values()}}
    _write_sheet(ws, master_df(parts), fmt)
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

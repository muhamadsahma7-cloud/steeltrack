"""Database access layer for SteelTrack.

Connects server-side through SQLAlchemy using the connection string in
.streamlit/secrets.toml ([connections.steel] url = ...). Production is a
Supabase Postgres database (Session pooler URL); a local SQLite file works
too, for testing. The connection uses the database owner's credentials, so
it bypasses RLS - keep the secret on the server, never in the browser.

All SQL here is written to run on both Postgres and SQLite.

Data model
    parts     one row per part of ONE copy of an assembly (master sheet row);
              asm_qty = how many identical copies of that assembly are made
    progress  quantity bookings: <qty> copies of sub assembly <unit> passed
              <stage> on <date>. A unit's done qty for a stage = sum(qty).
"""

from __future__ import annotations

import hashlib
import hmac
import secrets as _secrets

import pandas as pd
import streamlit as st
from sqlalchemy import inspect, text

CONN = "steel"


@st.cache_resource
def _conn():
    # Streamlit's SQLConnection: pools, reconnects, reads [connections.steel].
    return st.connection(CONN, type="sql")


def dialect() -> str:
    return _conn().engine.dialect.name          # 'postgresql' | 'sqlite'


def query(sql: str, params: dict | None = None, ttl: int = 30) -> pd.DataFrame:
    """Read query -> DataFrame (cached for `ttl` seconds)."""
    return _conn().query(sql, params=params or {}, ttl=ttl)


def engine():
    return _conn().engine


def execute(sql: str, params: dict | None = None) -> int:
    """Write query, committed. Returns rows affected."""
    with _conn().session as s:
        r = s.execute(text(sql), params or {})
        s.commit()
        return r.rowcount


def transaction(steps: list[tuple[str, dict]]) -> list[int]:
    """Several writes in ONE transaction; any error rolls all of them back."""
    out: list[int] = []
    with _conn().session as s:
        for sql, p in steps:
            out.append(s.execute(text(sql), p or {}).rowcount)
        s.commit()
    return out


# ------------------------------------------------------------------ schema
# The app creates its own tables on first start, so a brand-new (empty)
# Supabase project only needs its connection string in secrets.
PART_COLS = [
    "work_order", "priority", "scope", "area", "assembly_mark", "sub_assembly",
    "part_mark", "qty", "name", "profile", "grade", "length_mm", "surface_area",
    "unit_weight", "weight_kg", "total_weight", "asm_qty", "painting_req", "remark",
    "updated_at",
]
PROGRESS_COLS = ["work_order", "assembly_mark", "unit", "stage", "qty", "date", "do_no",
                 "done_by", "username", "ts"]


def _types(d: str) -> tuple[str, str]:
    pk = "BIGSERIAL PRIMARY KEY" if d == "postgresql" else "INTEGER PRIMARY KEY AUTOINCREMENT"
    fl = "DOUBLE PRECISION" if d == "postgresql" else "REAL"
    return pk, fl


def _part_coldefs(fl: str) -> dict[str, str]:
    return {
        "work_order": "TEXT", "priority": "INTEGER", "scope": "TEXT", "area": "TEXT",
        "assembly_mark": "TEXT", "sub_assembly": "TEXT", "part_mark": "TEXT",
        "qty": "INTEGER", "name": "TEXT", "profile": "TEXT", "grade": "TEXT",
        "length_mm": fl, "surface_area": fl, "unit_weight": fl, "weight_kg": fl,
        "total_weight": fl, "asm_qty": "INTEGER", "painting_req": "TEXT", "remark": "TEXT",
        "updated_at": "TEXT",
    }


def _ddl(d: str) -> list[str]:
    pk, fl = _types(d)
    cols = ", ".join(f"{c} {t}" for c, t in _part_coldefs(fl).items())
    return [
        f"CREATE TABLE IF NOT EXISTS parts (id {pk}, {cols})",
        "CREATE INDEX IF NOT EXISTS idx_parts_unit ON parts (work_order, assembly_mark, sub_assembly)",
        f"""CREATE TABLE IF NOT EXISTS progress (
            id {pk}, work_order TEXT, assembly_mark TEXT, unit TEXT, stage TEXT,
            qty INTEGER NOT NULL, date TEXT, do_no TEXT, done_by TEXT, username TEXT, ts TEXT)""",
        "CREATE INDEX IF NOT EXISTS idx_progress_unit ON progress (work_order, assembly_mark, unit, stage)",
        f"""CREATE TABLE IF NOT EXISTS progress_log (
            id {pk}, ts TEXT, username TEXT, stage TEXT, action TEXT,
            work_order TEXT, assembly_mark TEXT, unit TEXT, date_value TEXT,
            do_no TEXT, parts INTEGER, weight_kg {fl})""",
        """CREATE TABLE IF NOT EXISTS user_credentials (
            username TEXT PRIMARY KEY, password_hash TEXT NOT NULL,
            permission TEXT NOT NULL DEFAULT '', created_at TEXT)""",
        f"CREATE TABLE IF NOT EXISTS user_log (id {pk}, username TEXT, ts TEXT)",
        """CREATE TABLE IF NOT EXISTS project_settings (
            key TEXT PRIMARY KEY, value TEXT, updated_at TEXT)""",
    ]


_TABLES = ["parts", "progress", "progress_log", "user_credentials", "user_log",
           "project_settings"]


@st.cache_resource
def ensure_schema() -> bool:
    """Create missing tables/columns; seed a first admin when there are no
    users. Returns True when the default admin was just created."""
    d = dialect()
    eng = engine()
    with eng.begin() as cx:
        for stmt in _ddl(d):
            cx.execute(text(stmt))
        # parts made by an earlier version of the app: add the new columns
        have = {c["name"] for c in inspect(cx).get_columns("parts")}
        for c, t in _part_coldefs(_types(d)[1]).items():
            if c not in have:
                cx.execute(text(f"ALTER TABLE parts ADD COLUMN {c} {t}"))
        if d == "postgresql":
            # Supabase publishes the public schema over its REST API. RLS on
            # with no policies shuts that door; this app's owner login
            # bypasses RLS, so nothing here changes for the app itself.
            for t in _TABLES:
                cx.execute(text(f"ALTER TABLE {t} ENABLE ROW LEVEL SECURITY"))
        n = cx.execute(text("SELECT count(*) FROM user_credentials")).scalar()
        if not n:
            cx.execute(text("""INSERT INTO user_credentials
                               (username, password_hash, permission, created_at)
                               VALUES ('admin', :h, 'all', '')"""),
                       {"h": hash_password("admin123")})
    return not n


# -------------------------------------------------------------------- auth
_ITER = 120_000


def hash_password(pw: str) -> str:
    salt = _secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), _ITER).hex()
    return f"pbkdf2${_ITER}${salt}${h}"


def verify_password(pw: str, stored: str) -> bool:
    try:
        _, it, salt, h = str(stored).split("$")
        got = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), int(it)).hex()
        return hmac.compare_digest(got, h)
    except Exception:
        return False


def check_login(username: str, password: str) -> str | None:
    """Permission string on success, else None."""
    df = query("SELECT password_hash, permission FROM user_credentials WHERE username = :u",
               {"u": username}, ttl=0)
    if df.empty or not verify_password(password, df.iloc[0]["password_hash"]):
        return None
    return str(df.iloc[0]["permission"] or "")


def set_password(username: str, pw: str) -> None:
    execute("UPDATE user_credentials SET password_hash = :h WHERE username = :u",
            {"h": hash_password(pw), "u": username})


# ---------------------------------------------------------------- settings
def get_settings() -> dict:
    try:
        df = query("SELECT key, value FROM project_settings", ttl=0)
        return dict(zip(df["key"], df["value"]))
    except Exception:
        return {}


def set_settings(values: dict, ts: str) -> None:
    transaction([
        ("""INSERT INTO project_settings (key, value, updated_at) VALUES (:k, :v, :t)
            ON CONFLICT (key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at""",
         {"k": k, "v": None if v is None else str(v), "t": ts})
        for k, v in values.items()
    ])


# ------------------------------------------------------- import & backups
# These take the SQLAlchemy engine/connection explicitly so they can be
# exercised outside a Streamlit run.
_DETAIL = ["priority", "scope", "area", "qty", "name", "profile", "grade", "length_mm",
           "surface_area", "unit_weight", "weight_kg", "total_weight", "asm_qty",
           "painting_req", "remark"]
_KEY = ["work_order", "assembly_mark", "sub_assembly", "part_mark"]
_UNIT = ["work_order", "assembly_mark", "unit"]


def _is_pg(cx) -> bool:
    return cx.engine.dialect.name == "postgresql"


def snapshot(cx, stamp: str) -> str:
    """Copy parts + progress to *_bak_<stamp>, inside the caller's transaction."""
    for t in ("parts", "progress"):
        cx.execute(text(f"CREATE TABLE {t}_bak_{stamp} AS SELECT * FROM {t}"))
        if _is_pg(cx):
            cx.execute(text(f"ALTER TABLE {t}_bak_{stamp} ENABLE ROW LEVEL SECURITY"))
    return f"parts_bak_{stamp}"


def _fix_seq(cx) -> None:
    if _is_pg(cx):
        for t in ("parts", "progress"):
            cx.execute(text(f"SELECT setval(pg_get_serial_sequence('{t}','id'), "
                            f"coalesce((SELECT max(id) FROM {t}), 0) + 1, false)"))


def _tables(cx_or_none=None) -> set[str]:
    if dialect() == "postgresql":
        q = ("SELECT table_name AS t FROM information_schema.tables "
             "WHERE table_schema='public'")
    else:
        q = "SELECT name AS t FROM sqlite_master WHERE type='table'"
    return set(query(q, ttl=0)["t"])


def backups() -> list[str]:
    return sorted((t for t in _tables() if t.startswith("parts_bak_")), reverse=True)


def drop_backup(name: str) -> None:
    _check_bak(name)
    tabs = _tables()
    transaction([(f"DROP TABLE {t}", {})
                 for t in (name, name.replace("parts_bak_", "progress_bak_")) if t in tabs])


def _check_bak(name: str) -> None:
    if not name.startswith("parts_bak_") or not name.replace("_", "").isalnum():
        raise ValueError("not a backup table")


def restore(eng, name: str, stamp: str) -> str:
    """Put backup `name` back (parts, and progress when that backup has it);
    the current data is backed up first."""
    _check_bak(name)
    prog = name.replace("parts_bak_", "progress_bak_")
    has_prog = prog in _tables()
    with eng.begin() as cx:
        cur = snapshot(cx, stamp)
        for t, b in (("parts", name), ("progress", prog)):
            if t == "progress" and not has_prog:
                continue
            cols = ", ".join(c["name"] for c in inspect(cx).get_columns(b))
            cx.execute(text(f"DELETE FROM {t}"))
            cx.execute(text(f"INSERT INTO {t} ({cols}) SELECT {cols} FROM {b}"))
        _fix_seq(cx)
    return cur


def _insert(cx, table: str, df: pd.DataFrame, cols: list[str]) -> None:
    if not df.empty:
        df[[c for c in cols if c in df.columns]].to_sql(
            table, cx, if_exists="append", index=False, method="multi", chunksize=400)


def import_parts(eng, parts: pd.DataFrame, progress: pd.DataFrame, mode: str,
                 ts: str, stamp: str) -> str:
    """Load a parsed master sheet.

    replace  swap all parts AND all progress for the file's.
    merge    update part details, append parts not seen before, and take the
             file's progress only for sub assemblies + stages that have no
             progress booked in the app yet.
    Parts are matched on WO + assembly + sub assembly + part mark + nth
    occurrence (a sheet may repeat a row)."""
    parts = parts.copy()
    parts["updated_at"] = ts
    progress = progress.copy()
    if not progress.empty:
        progress["ts"] = ts
        progress["username"] = "import"
    with eng.begin() as cx:
        snap = snapshot(cx, stamp)
        if mode == "replace":
            cx.execute(text("DELETE FROM progress"))
            cx.execute(text("DELETE FROM parts"))
            _insert(cx, "parts", parts, PART_COLS)
            _insert(cx, "progress", progress, PROGRESS_COLS)
            _fix_seq(cx)
            return (f"Replaced the database with {len(parts):,} parts and "
                    f"{len(progress):,} progress records. Backup: `{snap}`.")
        old = pd.read_sql(text("SELECT id, work_order, assembly_mark, sub_assembly, part_mark "
                               "FROM parts ORDER BY id"), cx)
        for d in (old, parts):
            for c in _KEY:
                d["_" + c] = d[c].fillna("").astype(str).str.strip()
            d["_n"] = d.groupby(["_" + c for c in _KEY]).cumcount()
        on = ["_" + c for c in _KEY] + ["_n"]
        m = parts.merge(old[on + ["id"]], on=on, how="left")
        upd, ins = m[m["id"].notna()], m[m["id"].isna()]
        if not upd.empty:
            stage = upd[["id"] + _DETAIL].copy()
            stage["id"] = stage["id"].astype(int)
            cx.execute(text("DROP TABLE IF EXISTS import_stage"))
            stage.to_sql("import_stage", cx, index=False, method="multi", chunksize=400)
            sets = ", ".join(f"{c} = s.{c}" for c in _DETAIL)
            cx.execute(text(f"UPDATE parts SET {sets}, updated_at = :t "
                            f"FROM import_stage AS s WHERE parts.id = s.id"), {"t": ts})
            cx.execute(text("DROP TABLE import_stage"))
        _insert(cx, "parts", ins, PART_COLS)
        added_prog = 0
        if not progress.empty:
            have = pd.read_sql(text("SELECT DISTINCT work_order, assembly_mark, unit, stage "
                                    "FROM progress"), cx)
            k = _UNIT + ["stage"]
            new = progress.merge(have.assign(_has=1), on=k, how="left")
            new = new[new["_has"].isna()]
            _insert(cx, "progress", new, PROGRESS_COLS)
            added_prog = len(new)
        _fix_seq(cx)
        left = len(old) - len(upd)
    return (f"Updated {len(upd):,} parts, added {len(ins):,} new parts"
            + (f" and {added_prog:,} progress records from the file" if added_prog else "")
            + (f"; {left:,} parts in the database are not in this file and were kept"
               if left else "")
            + f". Backup: `{snap}`.")

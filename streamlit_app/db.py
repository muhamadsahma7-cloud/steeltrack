"""Database access layer for SteelTrack.

Connects server-side through SQLAlchemy using the connection string in
.streamlit/secrets.toml ([connections.steel] url = ...). Production is a
Supabase Postgres database (Session pooler URL); a local SQLite file works
too, for testing. The connection uses the database owner's credentials, so
it bypasses RLS - keep the secret on the server, never in the browser.

All SQL here is written to run on both Postgres and SQLite.
"""

from __future__ import annotations

import hashlib
import hmac
import secrets as _secrets

import pandas as pd
import streamlit as st
from sqlalchemy import text

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
    "priority", "work_order", "assembly_mark", "sub_assembly", "part_mark",
    "qty", "name", "profile", "kg_per_m", "length_mm", "weight_kg", "profile2",
    "grade", "remark", "painting_req",
    "fitup_date", "fitup_by", "welding_date", "welding_by",
    "paint_date", "paint_do", "paint_by", "site_date", "site_do", "site_by",
    "updated_at",
]


def _ddl(d: str) -> list[str]:
    pk = "BIGSERIAL PRIMARY KEY" if d == "postgresql" else "INTEGER PRIMARY KEY AUTOINCREMENT"
    fl = "DOUBLE PRECISION" if d == "postgresql" else "REAL"
    return [
        f"""CREATE TABLE IF NOT EXISTS parts (
            id {pk},
            priority INTEGER, work_order TEXT, assembly_mark TEXT NOT NULL,
            sub_assembly TEXT, part_mark TEXT, qty INTEGER, name TEXT, profile TEXT,
            kg_per_m {fl}, length_mm {fl}, weight_kg {fl}, profile2 TEXT, grade TEXT,
            remark TEXT, painting_req TEXT,
            fitup_date TEXT, fitup_by TEXT, welding_date TEXT, welding_by TEXT,
            paint_date TEXT, paint_do TEXT, paint_by TEXT,
            site_date TEXT, site_do TEXT, site_by TEXT, updated_at TEXT)""",
        "CREATE INDEX IF NOT EXISTS idx_parts_unit ON parts (work_order, assembly_mark, sub_assembly)",
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


_TABLES = ["parts", "progress_log", "user_credentials", "user_log", "project_settings"]


@st.cache_resource
def ensure_schema() -> bool:
    """Create missing tables; seed a first admin when there are no users.
    Returns True when the default admin was just created."""
    d = dialect()
    with _conn().session as s:
        for stmt in _ddl(d):
            s.execute(text(stmt))
        if d == "postgresql":
            # Supabase publishes the public schema over its REST API. RLS on
            # with no policies shuts that door; this app's owner login
            # bypasses RLS, so nothing here changes for the app itself.
            for t in _TABLES:
                s.execute(text(f"ALTER TABLE {t} ENABLE ROW LEVEL SECURITY"))
        n = s.execute(text("SELECT count(*) FROM user_credentials")).scalar()
        seeded = False
        if not n:
            s.execute(text("""INSERT INTO user_credentials
                              (username, password_hash, permission, created_at)
                              VALUES ('admin', :h, 'all', '')"""),
                      {"h": hash_password("admin123")})
            seeded = True
        s.commit()
    return seeded


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
_DETAIL = ["priority", "qty", "name", "profile", "kg_per_m", "length_mm", "weight_kg",
           "profile2", "grade", "remark", "painting_req"]
_PROGRESS = ["fitup_date", "welding_date", "paint_date", "paint_do", "site_date", "site_do"]
_KEY = ["work_order", "assembly_mark", "sub_assembly", "part_mark"]


def _is_pg(cx) -> bool:
    return cx.engine.dialect.name == "postgresql"


def snapshot(cx, stamp: str) -> str:
    """Copy the parts table to parts_bak_<stamp>, inside the caller's transaction."""
    name = f"parts_bak_{stamp}"
    cx.execute(text(f"CREATE TABLE {name} AS SELECT * FROM parts"))
    if _is_pg(cx):
        cx.execute(text(f"ALTER TABLE {name} ENABLE ROW LEVEL SECURITY"))
    return name


def _fix_seq(cx) -> None:
    if _is_pg(cx):
        cx.execute(text("SELECT setval(pg_get_serial_sequence('parts','id'), "
                        "coalesce((SELECT max(id) FROM parts), 0) + 1, false)"))


def backups() -> list[str]:
    if dialect() == "postgresql":
        q = ("SELECT table_name AS t FROM information_schema.tables "
             "WHERE table_schema='public' AND table_name LIKE 'parts_bak_%'")
    else:
        q = "SELECT name AS t FROM sqlite_master WHERE type='table' AND name LIKE 'parts_bak_%'"
    return sorted(query(q, ttl=0)["t"].tolist(), reverse=True)


def restore(eng, name: str, stamp: str) -> str:
    """Replace parts with backup `name`; the current data is backed up first."""
    if not name.startswith("parts_bak_") or not name.replace("_", "").isalnum():
        raise ValueError("not a backup table")
    with eng.begin() as cx:
        cur = snapshot(cx, stamp)
        cx.execute(text("DELETE FROM parts"))
        cx.execute(text(f"INSERT INTO parts SELECT * FROM {name}"))
        _fix_seq(cx)
    return cur


def import_parts(eng, new: pd.DataFrame, mode: str, ts: str, stamp: str) -> str:
    """Load a parsed master sheet. mode 'replace' swaps the whole table;
    'merge' updates part details, fills progress only where the app has none,
    and appends rows it hasn't seen. Rows are matched on WO + assembly +
    sub assembly + part mark + nth occurrence (the sheet repeats some rows)."""
    new = new.copy()
    new["updated_at"] = ts
    cols = [c for c in PART_COLS if c in new.columns]
    with eng.begin() as cx:
        snap = snapshot(cx, stamp)
        if mode == "replace":
            cx.execute(text("DELETE FROM parts"))
            new[cols].to_sql("parts", cx, if_exists="append", index=False,
                             method="multi", chunksize=400)
            _fix_seq(cx)
            return f"Replaced the database with {len(new):,} parts. Backup: `{snap}`."
        old = pd.read_sql(text("SELECT id, work_order, assembly_mark, sub_assembly, part_mark "
                               "FROM parts ORDER BY id"), cx)
        for d in (old, new):
            for c in _KEY:
                d["_" + c] = d[c].fillna("").astype(str).str.strip()
            d["_n"] = d.groupby(["_" + c for c in _KEY]).cumcount()
        on = ["_" + c for c in _KEY] + ["_n"]
        m = new.merge(old[on + ["id"]], on=on, how="left")
        upd, ins = m[m["id"].notna()], m[m["id"].isna()]
        if not upd.empty:
            stage = upd[["id"] + _DETAIL + _PROGRESS].copy()
            stage["id"] = stage["id"].astype(int)
            cx.execute(text("DROP TABLE IF EXISTS import_stage"))
            stage.to_sql("import_stage", cx, index=False, method="multi", chunksize=400)
            sets = ", ".join(
                [f"{c} = s.{c}" for c in _DETAIL]
                # never erase progress already booked in the app
                + [f"{c} = coalesce(nullif(parts.{c},''), s.{c})" for c in _PROGRESS])
            cx.execute(text(f"UPDATE parts SET {sets}, updated_at = :t "
                            f"FROM import_stage AS s WHERE parts.id = s.id"), {"t": ts})
            cx.execute(text("DROP TABLE import_stage"))
        if not ins.empty:
            ins[cols].to_sql("parts", cx, if_exists="append", index=False,
                             method="multi", chunksize=400)
        left = len(old) - len(upd)
    return (f"Updated {len(upd):,} parts, added {len(ins):,} new parts"
            + (f"; {left:,} parts in the database are not in this file and were kept"
               if left else "")
            + f". Backup: `{snap}`.")

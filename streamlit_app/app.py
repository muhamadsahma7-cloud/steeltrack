"""SteelTrack - steel structure fabrication progress, online.

Sister app of Team Piping: same sign-in, look and layout, built around the
steel master database (master_database_final.xlsx). Progress is booked per
SUB ASSEMBLY through four stages:

    Fit-up -> Welding -> Blasting & painting (D.O.) -> Send to site (D.O.)

A stage's kg is the part weight once that stage has a date.

Run locally:
    cd streamlit_app
    python -m streamlit run app.py
"""

from __future__ import annotations

import base64
import html
import re
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import streamlit as st
import db
import reports

MYT = ZoneInfo("Asia/Kuala_Lumpur")


def _today() -> date:
    """Today in Malaysia (Streamlit Cloud servers run on UTC)."""
    return datetime.now(MYT).date()


def _now_ts() -> str:
    return datetime.now(MYT).strftime("%Y-%m-%d %H:%M:%S")


def _stamp() -> str:
    return datetime.now(MYT).strftime("%Y%m%d_%H%M%S")


_LOGO_PATH = Path(__file__).parent / "assets" / "naec_logo.jpg"
try:
    LOGO_URI = "data:image/jpeg;base64," + base64.b64encode(_LOGO_PATH.read_bytes()).decode()
except Exception:
    LOGO_URI = ""

st.set_page_config(
    page_title="SteelTrack — NAEC Malaysia",
    page_icon=str(_LOGO_PATH) if _LOGO_PATH.exists() else "🏗️",
    layout="wide",
)

# I-beam cross-section mark
_BEAM_SVG = (
    "<svg width='30' height='30' viewBox='0 0 32 32' "
    "style='vertical-align:-6px;margin-right:8px'>"
    "<path d='M5 4h22v5H19v14h8v5H5v-5h8V9H5z' fill='none' "
    "stroke='#4d8dff' stroke-width='2.6' stroke-linejoin='round'/>"
    "<circle cx='16' cy='16' r='2.6' fill='#f59f00'/></svg>"
)


def brand_html(project: str = "") -> str:
    sub = html.escape(project) if project else "NAEC MALAYSIA SDN BHD"
    return (
        "<div style='line-height:1.12;margin:.1rem 0 .35rem'>"
        "<div style='font-size:1.6rem;font-weight:800;color:var(--ink,#f1f5f9)'>"
        f"{_BEAM_SVG}SteelTrack</div>"
        "<div style='font-size:.72rem;font-weight:700;letter-spacing:.16em;"
        f"color:var(--accent,#4d8dff);margin-top:3px;text-transform:uppercase'>{sub}</div></div>"
    )


# --------------------------------------------------------------- theme
_TOK_DARK = {
    "accent": "#4d8dff", "accent2": "#22c55e", "ink": "#f1f5f9", "text": "#e2e8f0",
    "line": "#2b3a52", "panel": "#1a2436", "muted": "#94a3b8", "bg": "#0f172a",
    "inputbg": "#0f1a2e", "inputbd": "#35507a", "inputtx": "#e6edf7", "inputph": "#7a8db0",
    "sb1": "#152036", "sb2": "#111a2e", "sb3": "#161327",
    "glass": "rgba(255,255,255,.06)", "glassbd": "rgba(255,255,255,.12)",
    "glasshov": "rgba(77,141,255,.18)", "pillhov": "rgba(255,255,255,.09)",
    "pillact": "rgba(77,141,255,.20)", "sbshadow": "0 2px 14px rgba(0,0,0,.35)",
    "st_scope": "#94a3b8", "st_fitup": "#fbbf24", "st_welding": "#4d8dff",
    "st_paint": "#c084fc", "st_site": "#34d399", "st_hold": "#fb7185",
}
_TOK_LIGHT = {
    "accent": "#2f6feb", "accent2": "#12b886", "ink": "#1f2933", "text": "#1f2933",
    "line": "#e2e8f0", "panel": "#ffffff", "muted": "#5b6b80", "bg": "#f6f8fb",
    "inputbg": "#ffffff", "inputbd": "#cbd5e1", "inputtx": "#1f2933", "inputph": "#94a3b8",
    "sb1": "#eef3fb", "sb2": "#f3f6fc", "sb3": "#f0f2fb",
    "glass": "rgba(255,255,255,.72)", "glassbd": "rgba(15,23,42,.10)",
    "glasshov": "rgba(47,111,235,.10)", "pillhov": "rgba(255,255,255,.9)",
    "pillact": "rgba(47,111,235,.12)", "sbshadow": "0 2px 10px rgba(15,23,42,.06)",
    "st_scope": "#64748b", "st_fitup": "#f59e0b", "st_welding": "#2f6feb",
    "st_paint": "#9333ea", "st_site": "#10b981", "st_hold": "#f43f5e",
}

_STATIC_CSS = """
.block-container{padding-top:2rem;max-width:1320px;position:relative;z-index:1}
section[data-testid="stSidebar"]{z-index:2}
.stTextInput input,.stNumberInput input,.stDateInput input,.stTextArea textarea,
[data-baseweb="input"],[data-baseweb="select"]>div,[data-baseweb="textarea"]{
  background:var(--inputbg)!important;color:var(--inputtx)!important;
  border:1px solid var(--inputbd)!important;border-radius:8px!important}
.stTextInput input::placeholder,.stTextArea textarea::placeholder{color:var(--inputph)!important}
h1{font-weight:700;letter-spacing:-.01em;color:var(--ink)}
h2{margin-top:.3rem;padding-bottom:.35rem;border-bottom:2px solid var(--line);color:var(--ink)}
h3{color:var(--accent);font-weight:600}
section[data-testid="stSidebar"]{
  background:linear-gradient(180deg,var(--sb1) 0%,var(--sb2) 55%,var(--sb3) 100%);
  border-right:1px solid var(--glassbd)}
section[data-testid="stSidebar"] .stButton>button{
  background:var(--glass);border:1px solid var(--glassbd);box-shadow:var(--sbshadow);
  border-radius:12px;color:var(--ink)}
section[data-testid="stSidebar"] .stButton>button:hover{
  background:var(--glasshov);border-color:var(--accent);color:var(--accent)}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"]{gap:.12rem}
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"]{margin-bottom:0!important}
.tp-navgrp{font-size:.66rem;font-weight:700;letter-spacing:.14em;text-transform:uppercase;
  color:var(--muted);margin:1rem 0 .2rem .6rem}
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]{
  background:transparent!important;border-radius:8px;margin:0;
  padding:.3rem .6rem;border-left:3px solid transparent;transition:background .12s ease}
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover{background:var(--pillhov)!important}
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] p{color:var(--text)!important;font-weight:500}
section[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] [data-testid="stIconMaterial"]{color:var(--muted)!important}
section[data-testid="stSidebar"] .st-key-tp_nav_on a[data-testid="stPageLink-NavLink"]{
  background:var(--pillact)!important;border-left-color:var(--accent)}
section[data-testid="stSidebar"] .st-key-tp_nav_on a[data-testid="stPageLink-NavLink"] p,
section[data-testid="stSidebar"] .st-key-tp_nav_on a[data-testid="stPageLink-NavLink"] [data-testid="stIconMaterial"]{
  color:var(--accent)!important;font-weight:650}
.tp-acct{display:flex;align-items:center;gap:.65rem;margin:1.2rem 0 .45rem;
  padding-top:1rem;border-top:1px solid var(--line)}
.tp-av{flex:none;width:34px;height:34px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-weight:700;background:var(--pillact);color:var(--accent)}
.tp-who{line-height:1.25;color:var(--ink);font-size:.9rem;min-width:0}
.tp-who span{display:block;color:var(--muted);font-size:.76rem}
section[data-testid="stSidebar"] [data-testid="stButtonGroup"]>div{flex-wrap:nowrap}
div[data-testid="stMetric"]{
  background:var(--panel);border:1px solid var(--line);border-left:4px solid var(--accent);
  border-radius:12px;padding:14px 16px}
[data-testid="stMetricLabel"] p{font-size:.9rem;font-weight:600;line-height:1.3;color:var(--muted)}
[data-testid="stMetricValue"],[data-testid="stMetricValue"] *{
  white-space:nowrap!important;color:var(--ink)!important;
  font-size:clamp(1.05rem,1.7vw,1.5rem)!important;font-variant-numeric:tabular-nums}
[data-testid="stMetricDelta"],[data-testid="stMetricDelta"] *{color:var(--muted)!important}
[data-testid="stMetricDelta"] svg{display:none}
.stButton>button,.stDownloadButton>button,.stForm button{border-radius:9px;font-weight:600}
div[data-testid="stDataFrame"],div[data-testid="stTable"]{border:1px solid var(--line);border-radius:10px}
.stTabs [aria-selected="true"]{color:var(--accent)!important}
div[data-testid="stAlert"]{border-radius:10px}
[data-testid="stProgress"] [role="progressbar"] div:has(>div:empty){background:var(--line)}
[data-testid="stProgress"] [role="progressbar"] div:empty{background:var(--accent2)}
hr{margin:1rem 0;border-color:var(--line)}
body:not(:has(.vega-embed)) #vg-tooltip-element{display:none!important}
.st-stage{font-size:.78rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;
  margin-bottom:.25rem}
"""
# stage-coloured metric edges + progress fills: containers keyed sg_<stage>
_STATIC_CSS += "".join(
    f'.st-key-sg_{k} [data-testid="stMetric"]{{border-left-color:var(--st_{k})}}'
    f'.st-key-sg_{k} [data-testid="stProgress"] [role="progressbar"] div:empty{{background:var(--st_{k})}}'
    f'.st-key-sg_{k} .st-stage{{color:var(--st_{k})}}'
    for k in ("fitup", "welding", "paint", "site", "scope")
)

_FORCE_TMPL = """
.stApp,[data-testid="stAppViewContainer"],[data-testid="stMain"]{
  background:@bg@!important;color:@text@!important}
[data-testid="stHeader"]{background:transparent!important}
.stMarkdown p,.stMarkdown li,.stMarkdown td,.stMarkdown th,
[data-testid="stWidgetLabel"] p,label p{color:@text@!important}
[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] p{color:@muted@!important}
h1,h2,h3,h4{color:@ink@!important}
[data-testid="stExpander"]{background:@panel@!important;border:1px solid @line@!important}
[data-testid="stExpander"] summary,[data-testid="stExpander"] summary *{color:@text@!important}
"""


def is_dark() -> bool:
    c = st.session_state.get("theme_choice")
    if c is None:
        qp = st.query_params.get("theme")
        if qp in ("light", "dark"):
            c = st.session_state["theme_choice"] = qp
    if c in ("light", "dark"):
        return c == "dark"
    try:
        return st.context.theme.type == "dark"
    except Exception:
        return False


def tok(k: str) -> str:
    return (_TOK_DARK if is_dark() else _TOK_LIGHT)[k]


def inject_css() -> None:
    t = _TOK_DARK if is_dark() else _TOK_LIGHT
    root = ":root{" + "".join(f"--{k}:{v};" for k, v in t.items()) + "}"
    force = _FORCE_TMPL
    for k in ("bg", "text", "ink", "panel", "line", "muted"):
        force = force.replace(f"@{k}@", t[k])
    st.markdown("<style>" + root + _STATIC_CSS + force + "</style>", unsafe_allow_html=True)
    if LOGO_URI and st.session_state.get("user"):
        st.markdown(
            f"<style>.stApp::before{{content:'';position:fixed;inset:0;z-index:0;"
            f"pointer-events:none;background:url('{LOGO_URI}') no-repeat center 45%;"
            f"background-size:min(42vw,460px);opacity:.04;filter:grayscale(1)}}</style>",
            unsafe_allow_html=True)


def dark_alt(chart):
    d = is_dark()
    axis = "#9fb0c9" if d else "#5b6b80"
    grid = "#26344c" if d else "#e3e8ef"
    return (chart.configure(background="transparent").configure_view(strokeWidth=0)
            .configure_axis(labelColor=axis, titleColor=axis, gridColor=grid, domainColor=grid)
            .configure_legend(labelColor=axis, titleColor=axis, orient="top"))


# --------------------------------------------------------------- splashes
def welcome_splash() -> None:
    """One-shot weld-bead flourish right after sign-in (CSS/SVG only)."""
    if not st.session_state.pop("just_logged_in", False):
        return
    who = html.escape(str(st.session_state.get("user", ""))[:40])
    st.markdown("""
<div class="tp-welcome">
  <svg viewBox="0 0 220 220" width="150" height="150" aria-hidden="true">
    <defs><linearGradient id="tpwHeat" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#ffffff"/><stop offset=".35" stop-color="#ffd27a"/>
      <stop offset=".7" stop-color="#ff8a1f"/><stop offset="1" stop-color="#ff5a12"/>
    </linearGradient></defs>
    <path d="M50 52h120v22h-44v72h44v22H50v-22h44V74H50z" fill="none" stroke="#24344f"
          stroke-width="9" stroke-linejoin="round"/>
    <path class="tpw-bead" d="M50 52h120v22h-44v72h44v22H50v-22h44V74H50z" fill="none"
          stroke="url(#tpwHeat)" stroke-width="9" stroke-linejoin="round"
          pathLength="100" stroke-dasharray="100" stroke-dashoffset="100"/>
  </svg>
  <div class="tpw-hi">Welcome</div>
  <div class="tpw-name">__WHO__</div>
  <div class="tpw-sub">Access granted &nbsp;·&nbsp; Steel Structure Tracker</div>
</div>
<style>
.tp-welcome{position:fixed;inset:0;z-index:2147483000;pointer-events:none;
  display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;
  background:radial-gradient(1200px 720px at 50% 42%,#16233b 0%,#0b1220 72%);
  animation:tpw-out .7s ease-in 2.1s forwards}
@keyframes tpw-out{to{opacity:0;visibility:hidden}}
.tpw-bead{animation:tpw-lay 1.1s ease-out .15s forwards;filter:drop-shadow(0 0 6px #ff8a1f88)}
@keyframes tpw-lay{to{stroke-dashoffset:0}}
.tpw-hi{font-size:2.4rem;font-weight:800;letter-spacing:.14em;text-transform:uppercase;
  color:#f1f5f9;opacity:0;transform:translateY(14px);animation:tpw-rise .5s ease-out 1.15s forwards}
.tpw-name{font-size:1.15rem;font-weight:700;letter-spacing:.06em;color:#4d8dff;opacity:0;
  transform:translateY(12px);animation:tpw-rise .5s ease-out 1.35s forwards}
.tpw-sub{font-family:ui-monospace,monospace;font-size:.72rem;letter-spacing:.22em;
  text-transform:uppercase;color:#8aa0bd;opacity:0;animation:tpw-rise .5s ease-out 1.55s forwards}
@keyframes tpw-rise{to{opacity:1;transform:translateY(0)}}
@media (prefers-reduced-motion:reduce){.tp-welcome *{animation:none!important}
  .tpw-bead{stroke-dashoffset:0}.tpw-hi,.tpw-name,.tpw-sub{opacity:1;transform:none}}
</style>""".replace("__WHO__", who), unsafe_allow_html=True)


def logout_splash() -> None:
    if not st.session_state.pop("just_logged_out", False):
        return
    idle = st.session_state.pop("logout_reason", "") == "idle"
    hi = "Timed out" if idle else "Signed out"
    sub = "Inactive for 30 minutes · session closed" if idle else "Session closed · see you tomorrow"
    st.markdown(f"""
<div class="tp-logout"><div class="lo-hi">{hi}</div><div class="lo-sub">{sub}</div></div>
<style>
.tp-logout{{position:fixed;inset:0;z-index:2147483000;pointer-events:none;display:flex;
  flex-direction:column;align-items:center;justify-content:center;gap:14px;
  background:radial-gradient(1200px 720px at 50% 42%,#131c2c 0%,#0b1220 72%);
  animation:lo-out .6s ease 1.4s forwards}}
@keyframes lo-out{{to{{opacity:0;visibility:hidden}}}}
.lo-hi{{font-size:2.2rem;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#e8edf3}}
.lo-sub{{font-family:ui-monospace,monospace;font-size:.72rem;letter-spacing:.2em;
  text-transform:uppercase;color:#8aa0bd}}
</style>""", unsafe_allow_html=True)


# --------------------------------------------------------------- start-up
inject_css()
logout_splash()

try:
    _seeded = db.ensure_schema()
except Exception as e:                       # no / wrong secrets, DB down
    st.error("Cannot reach the database. Check `[connections.steel]` in "
             "`.streamlit/secrets.toml` (or the app's Secrets on Streamlit Cloud).")
    st.caption(f"{type(e).__name__}: {e}")
    st.stop()

SETTINGS = db.get_settings()
PROJECT = SETTINGS.get("project_name", "")


def login_gate() -> None:
    if st.session_state.get("user"):
        return
    img = (f"<img src='{LOGO_URI}' alt='NAEC Malaysia' style='display:block;width:230px;"
           f"margin:0 auto 14px;background:#fff;padding:16px 20px;border-radius:18px;"
           f"box-shadow:0 8px 30px rgba(0,0,0,.35)'>" if LOGO_URI else "")
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        st.markdown(
            "<div style='text-align:center'>" + img +
            "<div style='font-size:1.7rem;font-weight:800;color:var(--ink);letter-spacing:.04em'>"
            "STEELTRACK</div>"
            "<div style='font-size:.72rem;font-weight:700;letter-spacing:.16em;"
            "color:var(--accent);margin-top:3px'>STEEL STRUCTURE PROGRESS"
            + (f" · {html.escape(PROJECT.upper())}" if PROJECT else "") + "</div></div>",
            unsafe_allow_html=True)
        st.subheader("Sign in")
        with st.form("login"):
            u = st.text_input("Username")
            p = st.text_input("Password", type="password")
            ok = st.form_submit_button("Sign in", use_container_width=True)
        if db.check_login("admin", "admin123") is not None:
            st.info("First start: sign in as **admin** / **admin123**, then change the "
                    "password (sidebar → Change password).")
    if ok:
        perm = db.check_login(u.strip(), p)
        if perm is None:
            st.error("Invalid username or password.")
        else:
            st.cache_data.clear()
            st.session_state.update(user=u.strip(), permission=perm, just_logged_in=True)
            try:
                db.execute("INSERT INTO user_log (username, ts) VALUES (:u, :t)",
                           {"u": u.strip(), "t": _now_ts()})
            except Exception:
                pass
            st.rerun()
    st.stop()


_PAGE_NAMES = ["Overview", "Update progress", "Assemblies", "Reports",
               "Activity", "Data admin", "Users"]


def _slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


if not st.session_state.get("user"):
    # register page URLs so a deep link survives the sign-in rerun
    st.navigation([st.Page(lambda: None, title=n, url_path=_slug(n), default=(i == 0))
                   for i, n in enumerate(_PAGE_NAMES)], position="hidden")
login_gate()
welcome_splash()

# ---- access control -------------------------------------------------
# user_credentials.permission = 'all' or a comma-separated list of tokens
STAGES = {
    "fitup":   dict(label="Fit-up", date="fitup_date", by="fitup_by", do=None,
                    token="Update Fit-Up"),
    "welding": dict(label="Welding", date="welding_date", by="welding_by", do=None,
                    token="Update Welding"),
    "paint":   dict(label="Blasting & painting", date="paint_date", by="paint_by",
                    do="paint_do", token="Update Painting"),
    "site":    dict(label="Send to site", date="site_date", by="site_by", do="site_do",
                    token="Update Site Delivery"),
}
KNOWN_TOKENS = [s["token"] for s in STAGES.values()]
ADMIN = "__admin__"
PAGE_PERMS = {
    "Overview": [],
    "Update progress": KNOWN_TOKENS,
    "Assemblies": [],
    "Reports": [],
    "Activity": [ADMIN],
    "Data admin": [ADMIN],
    "Users": [ADMIN],
}


def _toks(perm: str) -> set[str]:
    return {t.strip() for t in str(perm or "").split(",") if t.strip()}


def is_admin() -> bool:
    return st.session_state.get("permission") == "all"


def has_token(t: str) -> bool:
    return is_admin() or t in _toks(st.session_state.get("permission", ""))


def can_see(page: str) -> bool:
    need = PAGE_PERMS.get(page, [])
    if is_admin() or not need:
        return True
    return ADMIN not in need and any(has_token(t) for t in need)


IDLE_LIMIT_S = 1800


def _sign_out(reason: str = "") -> None:
    st.cache_data.clear()
    st.session_state.clear()
    st.session_state["just_logged_out"] = True
    st.session_state["logout_reason"] = reason


@st.fragment(run_every="1s")
def sidebar_clock() -> None:
    la = st.session_state.get("last_active")
    if la is not None and time.time() - la >= IDLE_LIMIT_S:
        _sign_out("idle")
        st.rerun(scope="app")
        return
    now = datetime.now(MYT)
    cd = ""
    if la is not None and IDLE_LIMIT_S - (time.time() - la) <= 60:
        left = max(0, int(IDLE_LIMIT_S - (time.time() - la)))
        cd = (f"<br><span style='font-size:11px;color:var(--st_fitup)'>"
              f"Auto sign-out in {left // 60}:{left % 60:02d}</span>")
    st.markdown(
        "<div style='font-family:ui-monospace,monospace;font-size:12px;color:var(--muted);"
        f"line-height:1.5'>{now:%a %d %b}  ·  <span style='font-weight:700;"
        f"color:var(--accent)'>{now:%H:%M:%S}</span> MYT{cd}</div>",
        unsafe_allow_html=True)


with st.sidebar:
    st.session_state["last_active"] = time.time()
    st.markdown(brand_html(PROJECT), unsafe_allow_html=True)
    sidebar_clock()


# --------------------------------------------------------------- data
@st.cache_data(ttl=30, show_spinner=False)
def load_parts() -> pd.DataFrame:
    df = db.query("SELECT * FROM parts ORDER BY id", ttl=0)
    for c in ("work_order", "assembly_mark", "sub_assembly", "part_mark", "name", "profile",
              "grade", "painting_req", "remark", "profile2",
              "fitup_date", "welding_date", "paint_date", "site_date", "paint_do", "site_do",
              "fitup_by", "welding_by", "paint_by", "site_by"):
        df[c] = df[c].fillna("").astype(str).str.strip()
    for c in ("weight_kg", "kg_per_m", "length_mm"):
        df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    df["weight_kg"] = df["weight_kg"].fillna(0.0)
    df["priority"] = pd.to_numeric(df["priority"], errors="coerce").astype("Int64")
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").astype("Int64")
    for s in STAGES.values():
        df[s["date"]] = df[s["date"]].str[:10]
    df["unit"] = df["sub_assembly"].where(df["sub_assembly"] != "", df["assembly_mark"])
    df["paint_req"] = df["painting_req"].str.upper().eq("YES")
    return df


# the same unit key the UPDATE statements match on
_UNIT_SQL = ("coalesce(work_order,'') = :wo AND assembly_mark = :am "
             "AND coalesce(nullif(trim(sub_assembly),''), assembly_mark) = :unit")


def done(df: pd.DataFrame, s: str, asof: str | None = None) -> pd.Series:
    d = df[STAGES[s]["date"]]
    m = d != ""
    return m & (d <= asof) if asof else m


def applies(df: pd.DataFrame, s: str) -> pd.Series:
    return df["paint_req"] if s == "paint" else pd.Series(True, index=df.index)


def prev_done(df: pd.DataFrame, s: str) -> pd.Series:
    if s == "fitup":
        return pd.Series(True, index=df.index)
    if s == "welding":
        return done(df, "fitup")
    if s == "paint":
        return done(df, "welding")
    return done(df, "paint").where(df["paint_req"], done(df, "welding"))


def ready(df: pd.DataFrame, s: str) -> pd.Series:
    return applies(df, s) & ~done(df, s) & prev_done(df, s)


# SQL twins of applies()/prev_done(): an UPDATE only touches parts that
# are really due, even if the page data is a few seconds stale
_PREV_SQL = {
    "fitup": "1=1",
    "welding": "coalesce(fitup_date,'') <> ''",
    "paint": ("upper(coalesce(painting_req,'')) = 'YES' "
              "AND coalesce(welding_date,'') <> ''"),
    "site": ("(CASE WHEN upper(coalesce(painting_req,'')) = 'YES' "
             "THEN coalesce(paint_date,'') ELSE coalesce(welding_date,'') END) <> ''"),
}
_NEXT_SQL = {     # clearing a stage is blocked while a later stage is booked
    "fitup": "coalesce(welding_date,'') = ''",
    "welding": "coalesce(paint_date,'') = '' AND coalesce(site_date,'') = ''",
    "paint": "coalesce(site_date,'') = ''",
    "site": "1=1",
}

LEVELS = ["Not started", "In progress", "Fit-up done", "Welded", "Painted / ready", "At site"]


def part_level(df: pd.DataFrame) -> pd.Series:
    lv = pd.Series(0, index=df.index)
    lv = lv.mask(done(df, "fitup"), 1)
    lv = lv.mask(done(df, "fitup") & done(df, "welding"), 2)
    lv = lv.mask((lv == 2) & (done(df, "paint") | ~df["paint_req"]), 3)
    lv = lv.mask(done(df, "site"), 4)
    return lv


def group_status(lv: pd.Series) -> str:
    lo, hi = int(lv.min()), int(lv.max())
    if lo == 0:
        return LEVELS[1] if hi > 0 else LEVELS[0]
    return LEVELS[lo + 1]


def stage_pcts(df: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    """weight + % complete per stage, grouped by `by`."""
    t = df.assign(_w=df["weight_kg"])
    for s in STAGES:
        t[f"_a_{s}"] = t["weight_kg"].where(applies(t, s), 0.0)
        t[f"_d_{s}"] = t["weight_kg"].where(applies(t, s) & done(t, s), 0.0)
    g = t.groupby(by, dropna=False).agg(
        **{"parts": ("weight_kg", "size"), "weight_kg": ("_w", "sum")},
        **{f"_a_{s}": (f"_a_{s}", "sum") for s in STAGES},
        **{f"_d_{s}": (f"_d_{s}", "sum") for s in STAGES})
    for s, cfg in STAGES.items():
        g[cfg["label"] + " %"] = (g[f"_d_{s}"] / g[f"_a_{s}"].where(g[f"_a_{s}"] > 0) * 100).fillna(0)
    g = g.drop(columns=[c for c in g.columns if c.startswith("_")]).reset_index()
    return g.rename(columns={"parts": "Parts", "weight_kg": "Weight (kg)"})


PCT_COLS = [s["label"] + " %" for s in STAGES.values()]


def show_table(df: pd.DataFrame, name: str, *, height: int | None = None,
               download: bool = True) -> None:
    cfg = {}
    for c in df.columns:
        if c in PCT_COLS:
            cfg[c] = st.column_config.ProgressColumn(c, min_value=0, max_value=100, format="%.0f%%")
        elif pd.api.types.is_float_dtype(df[c]):
            cfg[c] = st.column_config.NumberColumn(c, format="%.2f")
    kw = {"height": height} if height else {}
    st.dataframe(df, use_container_width=True, hide_index=True, column_config=cfg, **kw)
    if download:
        st.download_button("⬇ CSV", df.to_csv(index=False).encode(), file_name=f"{name}.csv",
                           mime="text/csv", key=f"dl_{name}")


def t_(kg: float) -> str:
    return f"{kg / 1000:,.2f} t"


def _refresh() -> None:
    st.cache_data.clear()


def _no_data_hint() -> bool:
    df = load_parts()
    if df.empty:
        st.info("No parts loaded yet. An admin can import the master database "
                "Excel in **Data admin**.")
        return True
    return False


# --------------------------------------------------------------- Overview
def page_overview() -> None:
    st.header("Project dashboard")
    if _no_data_hint():
        return
    df = load_parts()
    c1, c2, c3 = st.columns([1.2, 1.2, 2])
    asof = c1.date_input("As of", value=_today(), max_value=_today(), format="DD/MM/YYYY",
                         help="See the project as it stood at the end of this day.")
    wos = sorted(df["work_order"].unique())
    pick = c2.multiselect("Work order", wos, placeholder="All work orders")
    if pick:
        df = df[df["work_order"].isin(pick)]
    a = asof.isoformat()
    total = df["weight_kg"].sum()

    with st.container(key="sg_scope"):
        k = st.columns(4)
        k[0].metric("Total scope", t_(total), f"{len(df):,} parts", delta_color="off")
        k[1].metric("Assemblies", f"{df['assembly_mark'].nunique():,}",
                    f"{df.groupby(['work_order', 'unit']).ngroups:,} sub-assemblies",
                    delta_color="off")
        k[2].metric("Needs painting", t_(df.loc[df['paint_req'], 'weight_kg'].sum()),
                    f"{df['paint_req'].mean() * 100:.0f}% of parts" if len(df) else None,
                    delta_color="off")
        k[3].metric("Work orders", f"{df['work_order'].nunique():,}")

    st.subheader("Progress by stage")
    week0 = (asof - timedelta(days=6)).isoformat()
    cols = st.columns(4)
    for col, (s, cfg) in zip(cols, STAGES.items()):
        ap = applies(df, s)
        scope = df.loc[ap, "weight_kg"].sum()
        dn = df.loc[ap & done(df, s, a), "weight_kg"].sum()
        today_kg = df.loc[ap & (df[cfg["date"]] == a), "weight_kg"].sum()
        wk = df.loc[ap & (df[cfg["date"]] >= week0) & (df[cfg["date"]] <= a), "weight_kg"].sum()
        pct = dn / scope * 100 if scope else 0
        with col.container(key=f"sg_{s}"):
            st.markdown(f"<div class='st-stage'>{cfg['label']}</div>", unsafe_allow_html=True)
            st.metric(f"{pct:.1f}% complete", t_(dn), f"balance {t_(scope - dn)}",
                      delta_color="off")
            st.progress(min(pct / 100, 1.0))
            st.caption(f"Today **{today_kg:,.0f} kg** · last 7 days **{wk / 1000:,.2f} t**")

    st.subheader("Assembly status")
    # status as of the picked date: blank out dates after it
    dfa = df.copy()
    for cfg in STAGES.values():
        dfa[cfg["date"]] = dfa[cfg["date"]].where(dfa[cfg["date"]] <= a, "")
    dfa["_lv"] = part_level(dfa)
    stat = dfa.groupby(["work_order", "assembly_mark"])["_lv"].apply(group_status)
    w = dfa.groupby(["work_order", "assembly_mark"])["weight_kg"].sum()
    sdf = (pd.DataFrame({"status": stat, "w": w}).groupby("status")
           .agg(assemblies=("w", "size"), weight=("w", "sum"))
           .reindex(LEVELS).fillna(0).reset_index())
    m = st.columns(len(LEVELS))
    for col, r in zip(m, sdf.itertuples()):
        col.metric(r.status, f"{int(r.assemblies):,}", t_(r.weight), delta_color="off")

    st.subheader("Daily output")
    days = st.segmented_control("Window", [14, 30, 60, 90], default=30,
                                format_func=lambda n: f"{n} days", key="ov_win") or 30
    start = (asof - timedelta(days=days - 1)).isoformat()
    rows = []
    for s, cfg in STAGES.items():
        x = df[applies(df, s) & (df[cfg["date"]] >= start) & (df[cfg["date"]] <= a)]
        g = x.groupby(cfg["date"])["weight_kg"].sum()
        rows += [{"date": d, "stage": cfg["label"], "kg": v} for d, v in g.items()]
    order = [c["label"] for c in STAGES.values()]
    colors = [tok(f"st_{s}") for s in STAGES]
    if rows:
        cd = pd.DataFrame(rows)
        cd["date"] = pd.to_datetime(cd["date"])
        ch = (alt.Chart(cd).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2)
              .encode(x=alt.X("yearmonthdate(date):T", title=None,
                              axis=alt.Axis(format="%d %b", labelAngle=0)),
                      xOffset=alt.XOffset("stage:N", sort=order),
                      y=alt.Y("kg:Q", title="kg"),
                      color=alt.Color("stage:N", sort=order, title=None,
                                      scale=alt.Scale(domain=order, range=colors)),
                      tooltip=[alt.Tooltip("yearmonthdate(date):T", title="Date", format="%a %d %b"),
                               "stage:N", alt.Tooltip("kg:Q", format=",.1f")])
              .properties(height=300))
        st.altair_chart(dark_alt(ch), use_container_width=True)
    else:
        st.caption("No progress booked in this window.")

    st.subheader("Cumulative progress")
    rows = []
    for s, cfg in STAGES.items():
        x = df[applies(df, s) & done(df, s, a)]
        g = x.groupby(cfg["date"])["weight_kg"].sum().sort_index().cumsum()
        rows += [{"date": d, "stage": cfg["label"], "t": v / 1000} for d, v in g.items()]
    if rows:
        cd = pd.DataFrame(rows)
        cd["date"] = pd.to_datetime(cd["date"], errors="coerce")
        ch = (alt.Chart(cd.dropna()).mark_line(point=True, interpolate="step-after")
              .encode(x=alt.X("date:T", title=None, axis=alt.Axis(format="%d %b")),
                      y=alt.Y("t:Q", title="tonnes"),
                      color=alt.Color("stage:N", sort=order, title=None,
                                      scale=alt.Scale(domain=order, range=colors)),
                      tooltip=[alt.Tooltip("date:T", format="%a %d %b"), "stage:N",
                               alt.Tooltip("t:Q", format=",.2f", title="tonnes")])
              .properties(height=280))
        rule = alt.Chart(pd.DataFrame({"t": [total / 1000]})).mark_rule(
            strokeDash=[5, 4], color=tok("muted")).encode(y="t:Q")
        st.altair_chart(dark_alt(ch + rule), use_container_width=True)
        st.caption(f"Dashed line = total scope {t_(total)}.")

    t1, t2 = st.tabs(["By priority", "By member type"])
    with t1:
        p = stage_pcts(dfa, ["priority"]).sort_values("priority", na_position="last")
        p["priority"] = p["priority"].astype("string").fillna("(none)")
        show_table(p.rename(columns={"priority": "Priority"}), "progress_by_priority")
    with t2:
        p = stage_pcts(dfa, ["name"]).sort_values("Weight (kg)", ascending=False)
        show_table(p.rename(columns={"name": "Member type"}), "progress_by_member")


# --------------------------------------------------------------- Update
def _unit_table(df: pd.DataFrame, s: str) -> pd.DataFrame:
    """One row per sub-assembly that has parts due for stage `s`."""
    cfg = STAGES[s]
    r = df[ready(df, s)]
    if r.empty:
        return pd.DataFrame()
    prev = {"welding": "fitup_date", "paint": "welding_date"}.get(s)
    heavy = (df.sort_values("weight_kg", ascending=False)
             .drop_duplicates(["work_order", "assembly_mark", "unit"])
             .set_index(["work_order", "assembly_mark", "unit"]))
    allp = df.groupby(["work_order", "assembly_mark", "unit"]).size()
    g = r.groupby(["work_order", "assembly_mark", "unit"]).agg(
        priority=("priority", "min"), ready_parts=("weight_kg", "size"),
        weight_kg=("weight_kg", "sum"),
        prev_date=(prev, "max") if prev else ("weight_kg", "size"))
    g["parts"] = allp.reindex(g.index).values
    g["main member"] = (heavy.reindex(g.index)["name"] + " " +
                        heavy.reindex(g.index)["profile"]).values
    if s == "site":
        g["prev_date"] = r.assign(_p=r["paint_date"].where(r["paint_req"], r["welding_date"])) \
            .groupby(["work_order", "assembly_mark", "unit"])["_p"].max()
    g = g.reset_index()
    out = pd.DataFrame({
        "✓": False,
        "WO": g["work_order"], "Assembly": g["assembly_mark"], "Sub assembly": g["unit"],
        "Priority": g["priority"], "Main member": g["main member"],
        "Parts due": g["ready_parts"].astype(int), "Parts": g["parts"].astype(int),
        "Weight (kg)": g["weight_kg"].round(2),
    })
    if s != "fitup":
        out["Previous stage"] = g["prev_date"].astype(str)
    return out.sort_values(["Priority", "Assembly", "Sub assembly"], na_position="last")


def page_update() -> None:
    st.header("Update progress")
    if _no_data_hint():
        return
    mine = [s for s, c in STAGES.items() if has_token(c["token"])]
    if not mine:
        st.warning("Your account can't update any stage.")
        return
    s = st.segmented_control("Stage", mine, default=mine[0], key="up_stage",
                             format_func=lambda k: STAGES[k]["label"]) or mine[0]
    cfg = STAGES[s]
    df = load_parts()

    f1, f2, f3, f4 = st.columns([1, 1.3, 1.3, 1.6])
    wo = f1.selectbox("Work order", ["All"] + sorted(df["work_order"].unique()), key="up_wo")
    pri = f2.multiselect("Priority", sorted(df["priority"].dropna().astype(int).unique()),
                         key="up_pri", placeholder="All")
    names = f3.multiselect("Member type", sorted(df["name"].unique()), key="up_name",
                           placeholder="All")
    q = f4.text_input("Search assembly / sub assembly", key="up_q",
                      placeholder="e.g. R2-B1 or R2-H39")

    units = _unit_table(df, s)
    if units.empty:
        st.success(f"Nothing is waiting for {cfg['label'].lower()} — every eligible part is booked.")
    else:
        if wo != "All":
            units = units[units["WO"] == wo]
        if pri:
            keys = set(df[df["priority"].isin(pri)][["work_order", "unit"]]
                       .itertuples(index=False, name=None))
            units = units[[k in keys for k in zip(units["WO"], units["Sub assembly"])]]
        if names:
            keys = set(df[df["name"].isin(names)][["work_order", "unit"]]
                       .itertuples(index=False, name=None))
            units = units[[k in keys for k in zip(units["WO"], units["Sub assembly"])]]
        if q.strip():
            ql = q.strip().lower()
            units = units[units["Assembly"].str.lower().str.contains(ql, regex=False)
                          | units["Sub assembly"].str.lower().str.contains(ql, regex=False)]

    prereq = {"fitup": "", "welding": " (fit-up done)", "paint": " (welded, painting required)",
              "site": " (painted, or welded when no painting)"}[s]
    st.caption(f"Sub-assemblies with parts due for **{cfg['label'].lower()}**{prereq}: "
               f"**{len(units):,}** · {units['Weight (kg)'].sum() if len(units) else 0:,.1f} kg")

    if not units.empty:
        ver = st.session_state.setdefault("up_ver", 0)
        sel_all = st.checkbox("Select all shown", key=f"up_all_{ver}")
        units["✓"] = sel_all
        ed = st.data_editor(
            units, hide_index=True, use_container_width=True, height=420,
            disabled=[c for c in units.columns if c != "✓"],
            column_config={"✓": st.column_config.CheckboxColumn("✓", width="small"),
                           "Weight (kg)": st.column_config.NumberColumn(format="%.2f")},
            key=f"up_ed_{s}_{ver}_{sel_all}_{wo}_{len(units)}")
        picked = ed[ed["✓"]]
        st.markdown(f"**Selected: {len(picked):,} sub-assemblies · "
                    f"{picked['Weight (kg)'].sum():,.1f} kg**")

        with st.form("up_form"):
            c1, c2, c3 = st.columns(3)
            d = c1.date_input(f"{cfg['label']} date", value=_today(), max_value=_today(),
                              format="DD/MM/YYYY")
            do_no = (c2.text_input("D.O. No.", placeholder="e.g. 012") if cfg["do"] else None)
            who = c3.text_input("Done by / supervisor", value=st.session_state["user"])
            go = st.form_submit_button(f"Confirm {cfg['label'].lower()}", type="primary",
                                       use_container_width=True)
        if go:
            if picked.empty:
                st.error("Tick at least one sub-assembly.")
            elif cfg["do"] and not (do_no or "").strip():
                st.error("Enter the D.O. number.")
            else:
                _book(s, picked, d.isoformat(), (do_no or "").strip() or None, who.strip())

    if st.session_state.get("up_msg"):
        st.success(st.session_state.pop("up_msg"))

    with st.expander(f"Undo — clear a booked {cfg['label'].lower()} date"):
        _undo_ui(s, df)


def _book(s: str, picked: pd.DataFrame, d: str, do_no: str | None, who: str) -> None:
    cfg = STAGES[s]
    ts = _now_ts()
    sets = f"{cfg['date']} = :d, {cfg['by']} = :by, updated_at = :t"
    if cfg["do"]:
        sets += f", {cfg['do']} = :do"
    rows = picked.to_dict("records")
    steps = []
    for r in rows:
        p = {"wo": r["WO"], "am": r["Assembly"], "unit": r["Sub assembly"], "d": d,
             "by": who or None, "t": ts, "do": do_no}
        steps.append((f"UPDATE parts SET {sets} WHERE {_UNIT_SQL} "
                      f"AND coalesce({cfg['date']},'') = '' AND {_PREV_SQL[s]}", p))
    n = db.transaction(steps)
    log = []
    for r, k in zip(rows, n):
        if k:
            log.append(("""INSERT INTO progress_log (ts, username, stage, action, work_order,
                           assembly_mark, unit, date_value, do_no, parts, weight_kg)
                           VALUES (:t, :u, :s, 'set', :wo, :am, :unit, :d, :do, :n, :w)""",
                        {"t": ts, "u": st.session_state["user"], "s": cfg["label"],
                         "wo": r["WO"], "am": r["Assembly"], "unit": r["Sub assembly"],
                         "d": d, "do": do_no, "n": k, "w": float(r["Weight (kg)"])}))
    if log:
        db.transaction(log)
    _refresh()
    st.session_state["up_ver"] = st.session_state.get("up_ver", 0) + 1   # fresh ticks
    st.session_state["up_msg"] = (
        f"{cfg['label']} booked for {len(log)} sub-assemblies ({sum(n):,} parts, "
        f"{picked['Weight (kg)'].sum():,.1f} kg) on {d}" + (f", D.O. {do_no}" if do_no else "")
        + ("" if len(log) == len(picked) else
           f". {len(picked) - len(log)} were already booked by someone else."))
    st.rerun()


def _undo_ui(s: str, df: pd.DataFrame) -> None:
    cfg = STAGES[s]
    later = {"fitup": ["welding"], "welding": ["paint", "site"], "paint": ["site"], "site": []}[s]
    m = done(df, s)
    for x in later:
        m &= ~done(df, x)
    x = df[m]
    if x.empty:
        st.caption("Nothing that can be cleared (a date can only be cleared while the "
                   "next stage is still open).")
        return
    g = (x.groupby(["work_order", "assembly_mark", "unit", cfg["date"]]
                   + ([cfg["do"]] if cfg["do"] else []))["weight_kg"]
         .agg(["size", "sum"]).reset_index())
    g["label"] = g.apply(lambda r: f"{r['unit']}  ·  {r[cfg['date']]}"
                         + (f"  ·  D.O. {r[cfg['do']]}" if cfg["do"] and r[cfg["do"]] else "")
                         + f"  ·  {r['size']} parts / {r['sum']:,.1f} kg", axis=1)
    g = g.sort_values(cfg["date"], ascending=False)
    pick = st.multiselect("Sub-assemblies to clear", g.index.tolist(),
                          format_func=lambda i: g.loc[i, "label"], key=f"undo_{s}")
    if st.button("Clear selected dates", key=f"undo_btn_{s}", disabled=not pick):
        ts = _now_ts()
        sets = f"{cfg['date']} = NULL, {cfg['by']} = NULL, updated_at = :t"
        if cfg["do"]:
            sets += f", {cfg['do']} = NULL"
        steps = []
        for i in pick:
            r = g.loc[i]
            p = {"wo": r["work_order"], "am": r["assembly_mark"], "unit": r["unit"], "t": ts}
            steps.append((f"UPDATE parts SET {sets} WHERE {_UNIT_SQL} AND {_NEXT_SQL[s]}", p))
            steps.append(("""INSERT INTO progress_log (ts, username, stage, action, work_order,
                             assembly_mark, unit, date_value, do_no, parts, weight_kg)
                             VALUES (:t, :u, :s, 'clear', :wo, :am, :unit, :d, :do, :n, :w)""",
                          {**p, "u": st.session_state["user"], "s": cfg["label"],
                           "d": r[cfg["date"]], "do": r[cfg["do"]] if cfg["do"] else None,
                           "n": int(r["size"]), "w": float(r["sum"])}))
        db.transaction(steps)
        _refresh()
        st.session_state["up_msg"] = f"Cleared {cfg['label'].lower()} on {len(pick)} sub-assemblies."
        st.rerun()


# --------------------------------------------------------------- Assemblies
def page_assemblies() -> None:
    st.header("Assemblies")
    if _no_data_hint():
        return
    df = load_parts()
    df = df.assign(_lv=part_level(df))
    f1, f2, f3, f4 = st.columns([1, 1.2, 1.4, 1.6])
    wo = f1.selectbox("Work order", ["All"] + sorted(df["work_order"].unique()), key="as_wo")
    pri = f2.multiselect("Priority", sorted(df["priority"].dropna().astype(int).unique()),
                         key="as_pri", placeholder="All")
    stt = f3.multiselect("Status", LEVELS, key="as_st", placeholder="All")
    q = f4.text_input("Search assembly mark", key="as_q")
    if wo != "All":
        df = df[df["work_order"] == wo]

    keys = ["work_order", "assembly_mark"]
    a = stage_pcts(df, keys)
    a["status"] = df.groupby(keys)["_lv"].apply(group_status).values
    a["priority"] = df.groupby(keys)["priority"].min().values
    a["sub assemblies"] = df.groupby(keys)["unit"].nunique().values
    last = df[[c["date"] for c in STAGES.values()]].max(axis=1)
    a["last update"] = last.groupby([df[k] for k in keys]).max().values
    if pri:
        hit = set(df[df["priority"].isin(pri)]["assembly_mark"])
        a = a[a["assembly_mark"].isin(hit)]
    if stt:
        a = a[a["status"].isin(stt)]
    if q.strip():
        a = a[a["assembly_mark"].str.lower().str.contains(q.strip().lower(), regex=False)]
    a = a[["work_order", "assembly_mark", "priority", "status", "sub assemblies", "Parts",
           "Weight (kg)"] + PCT_COLS + ["last update"]]
    a = a.sort_values(["priority", "assembly_mark"], na_position="last")
    st.caption(f"{len(a):,} assemblies · {a['Weight (kg)'].sum() / 1000:,.2f} t")
    show_table(a.rename(columns={"work_order": "WO", "assembly_mark": "Assembly",
                                 "priority": "Priority", "status": "Status",
                                 "sub assemblies": "Sub assemblies",
                                 "last update": "Last update"}), "assemblies", height=420)

    st.subheader("Assembly detail")
    if a.empty:
        return
    pick = st.selectbox("Assembly", a["assembly_mark"].tolist(), key="as_pick")
    x = df[df["assembly_mark"] == pick]
    u = stage_pcts(x, ["unit"])
    u["status"] = x.groupby("unit")["_lv"].apply(group_status).values
    st.markdown("**Sub-assemblies**")
    show_table(u.rename(columns={"unit": "Sub assembly", "status": "Status"}),
               f"subassemblies_{pick}")
    st.markdown("**Parts**")
    cols = ["unit", "part_mark", "qty", "name", "profile", "length_mm", "weight_kg", "grade",
            "painting_req", "fitup_date", "welding_date", "paint_date", "paint_do",
            "site_date", "site_do"]
    show_table(x[cols].rename(columns={
        "unit": "Sub assembly", "part_mark": "Part", "qty": "No.", "length_mm": "Length (mm)",
        "weight_kg": "Weight (kg)", "painting_req": "Paint", "fitup_date": "Fit-up",
        "welding_date": "Welding", "paint_date": "Painting", "paint_do": "Paint D.O.",
        "site_date": "To site", "site_do": "Site D.O."}), f"parts_{pick}")


# --------------------------------------------------------------- Reports
def page_reports() -> None:
    st.header("Reports")
    if _no_data_hint():
        return
    df = load_parts()
    t1, t2, t3 = st.tabs(["Daily / period report", "D.O. register", "Export master database"])

    with t1:
        c1, c2, c3 = st.columns([1.3, 1, 1])
        today = _today()
        mode = c1.segmented_control("Period", ["Today", "This week", "Last week", "This month",
                                               "Custom"], default="This week", key="rp_mode") \
            or "This week"
        mon = today - timedelta(days=today.weekday())
        start, end = {
            "Today": (today, today),
            "This week": (mon, today),
            "Last week": (mon - timedelta(days=7), mon - timedelta(days=1)),
            "This month": (today.replace(day=1), today),
        }.get(mode, (today - timedelta(days=6), today))
        if mode == "Custom":
            start = c2.date_input("From", value=start, format="DD/MM/YYYY", key="rp_from")
            end = c3.date_input("To", value=end, format="DD/MM/YYYY", key="rp_to")
        s_, e_ = start.isoformat(), end.isoformat()
        st.caption(f"{start:%a %d %b %Y} → {end:%a %d %b %Y}")

        k = st.columns(4)
        daily = {}
        detail = []
        for col, (s, cfg) in zip(k, STAGES.items()):
            x = df[applies(df, s) & (df[cfg["date"]] >= s_) & (df[cfg["date"]] <= e_)]
            with col.container(key=f"sg_{s}"):
                st.metric(cfg["label"], f"{x['weight_kg'].sum():,.1f} kg",
                           f"{x.groupby(['work_order', 'unit']).ngroups:,} sub-assemblies",
                           delta_color="off")
            daily[cfg["label"] + " (kg)"] = x.groupby(cfg["date"])["weight_kg"].sum()
            if not x.empty:
                g = x.groupby(["work_order", "assembly_mark", "unit", cfg["date"]]
                              + ([cfg["do"]] if cfg["do"] else [])).agg(
                    parts=("weight_kg", "size"), kg=("weight_kg", "sum")).reset_index()
                g = g.rename(columns={cfg["date"]: "Date", "work_order": "WO",
                                      "assembly_mark": "Assembly", "unit": "Sub assembly",
                                      (cfg["do"] or "_"): "D.O. No.", "parts": "Parts",
                                      "kg": "Weight (kg)"})
                g.insert(0, "Stage", cfg["label"])
                detail.append(g)
        dd = pd.DataFrame(daily).fillna(0.0)
        dd.index.name = "Date"
        dd = dd.reindex(sorted(dd.index)).reset_index()
        if not dd.empty:
            tot = {"Date": "TOTAL", **{c: dd[c].sum() for c in dd.columns if c != "Date"}}
            dd = pd.concat([dd, pd.DataFrame([tot])], ignore_index=True)
        st.markdown("**Daily output (kg)**")
        show_table(dd, f"daily_{s_}_{e_}", download=False)
        det = (pd.concat(detail, ignore_index=True) if detail else pd.DataFrame())
        if not det.empty:
            det = det.sort_values(["Date", "Stage", "Assembly", "Sub assembly"])
            st.markdown("**What was done**")
            show_table(det, f"detail_{s_}_{e_}", download=False, height=360)
        heads = [f"SteelTrack — progress report{(' — ' + PROJECT) if PROJECT else ''}",
                 f"Period: {start:%d %b %Y} to {end:%d %b %Y}",
                 f"Generated {_now_ts()} MYT by {st.session_state['user']}"]
        st.download_button(
            "⬇ Download report (.xlsx)", type="primary",
            data=reports.workbook({"Daily": dd, "Detail": det},
                                  {c: "#,##0.00" for c in list(dd.columns) + ["Weight (kg)"]},
                                  {"Daily": heads, "Detail": heads}),
            file_name=f"SteelTrack_report_{s_}_{e_}.xlsx", key="rp_xlsx")

    with t2:
        kind = st.segmented_control("Delivery", ["paint", "site"], default="paint", key="do_kind",
                                    format_func=lambda k: STAGES[k]["label"]) or "paint"
        cfg = STAGES[kind]
        x = df[df[cfg["do"]] != ""]
        if x.empty:
            st.caption("No D.O. recorded yet.")
        else:
            reg = (x.groupby(cfg["do"]).agg(
                date=(cfg["date"], "max"), assemblies=("assembly_mark", "nunique"),
                sub_assemblies=("unit", "nunique"), parts=("weight_kg", "size"),
                weight_kg=("weight_kg", "sum")).reset_index()
                .rename(columns={cfg["do"]: "D.O. No.", "date": "Date",
                                 "sub_assemblies": "Sub assemblies", "assemblies": "Assemblies",
                                 "parts": "Parts", "weight_kg": "Weight (kg)"})
                .sort_values(["Date", "D.O. No."], ascending=False))
            show_table(reg, f"do_register_{kind}")
            pick = st.selectbox("D.O. contents", reg["D.O. No."].tolist(), key="do_pick")
            y = x[x[cfg["do"]] == pick]
            lst = (y.groupby(["work_order", "assembly_mark", "unit"]).agg(
                parts=("weight_kg", "size"), pcs=("qty", "sum"), kg=("weight_kg", "sum"))
                .reset_index().rename(columns={"work_order": "WO", "assembly_mark": "Assembly",
                                               "unit": "Sub assembly", "parts": "Parts",
                                               "pcs": "Pcs", "kg": "Weight (kg)"}))
            show_table(lst, f"do_{kind}_{pick}", download=False)
            heads = [f"Delivery list — {cfg['label']} — D.O. {pick}",
                     f"Date: {y[cfg['date']].max()}   ·   {len(lst)} sub-assemblies   ·   "
                     f"{y['weight_kg'].sum():,.2f} kg" + (f"   ·   {PROJECT}" if PROJECT else "")]
            st.download_button("⬇ Delivery list (.xlsx)",
                               reports.workbook({"Delivery list": lst}, {"Weight (kg)": "#,##0.00"},
                                                {"Delivery list": heads}),
                               file_name=f"DO_{kind}_{pick}.xlsx", key="do_xlsx")

    with t3:
        st.write("Download the whole database in the same 25-column layout as "
                 "`master_database_final.xlsx` (stage kg = part weight once a date is booked).")
        if st.button("Prepare Excel file", key="mx_prep"):
            st.session_state["mx_bytes"] = reports.master_xlsx(load_parts())
        if st.session_state.get("mx_bytes"):
            st.download_button("⬇ master_database.xlsx", st.session_state["mx_bytes"],
                               type="primary",
                               file_name=f"master_database_{_today():%Y%m%d}.xlsx", key="mx_dl")


# --------------------------------------------------------------- Activity
def page_activity() -> None:
    st.header("Activity")
    t1, t2 = st.tabs(["Progress log", "Sign-ins"])
    with t1:
        lg = db.query("""SELECT ts AS "Time", username AS "User", stage AS "Stage",
                                action AS "Action", work_order AS "WO",
                                assembly_mark AS "Assembly", unit AS "Sub assembly",
                                date_value AS "Date", do_no AS "D.O.", parts AS "Parts",
                                weight_kg AS "Weight (kg)"
                         FROM progress_log ORDER BY id DESC LIMIT 2000""", ttl=0)
        u = st.multiselect("User", sorted(lg["User"].dropna().unique()), key="ac_u")
        if u:
            lg = lg[lg["User"].isin(u)]
        show_table(lg, "progress_log", height=480)
    with t2:
        show_table(db.query("""SELECT ts AS "Time", username AS "User" FROM user_log
                               ORDER BY id DESC LIMIT 1000""", ttl=0), "sign_ins")


# --------------------------------------------------------------- Data admin
def page_admin() -> None:
    st.header("Data admin")
    t1, t2, t3 = st.tabs(["Import master database", "Backups", "Project settings"])
    with t1:
        st.write("Upload the master database Excel (same layout as "
                 "`master_database_final.xlsx`, sheet **Master Database**).")
        up = st.file_uploader("Excel file", type=["xlsx", "xlsm"], key="imp_file")
        if up is not None:
            try:
                new = reports.parse_master(up)
            except Exception as e:
                st.error(f"Can't read that file: {e}")
                return
            k = st.columns(4)
            k[0].metric("Rows", f"{len(new):,}")
            k[1].metric("Assemblies", f"{new['assembly_mark'].nunique():,}")
            k[2].metric("Total weight", t_(new["weight_kg"].sum()))
            k[3].metric("Rows with fit-up date", f"{new['fitup_date'].notna().sum():,}")
            with st.expander("Preview (first 50 rows)"):
                st.dataframe(new.head(50), use_container_width=True, hide_index=True)
            existing = len(load_parts())
            mode = st.radio(
                "Import mode",
                ["merge", "replace"],
                format_func=lambda m: {
                    "merge": "Add new parts & update details (keeps progress booked in the app)",
                    "replace": "Replace everything with this file (progress comes from the file)",
                }[m], key="imp_mode", index=0 if existing else 1)
            ok = True
            if mode == "replace" and existing:
                st.warning(f"This deletes all {existing:,} parts now in the database "
                           "(a backup copy is made first).")
                ok = st.text_input("Type REPLACE to confirm", key="imp_conf") == "REPLACE"
            if st.button("Import", type="primary", disabled=not ok, key="imp_go"):
                with st.spinner("Importing…"):
                    msg = db.import_parts(db.engine(), new, mode, _now_ts(), _stamp())
                _refresh()
                st.success(msg)

    with t2:
        st.caption("A backup copy of the parts table is taken automatically before every import.")
        if st.button("Take a backup now", key="bk_now"):
            with db.engine().begin() as cx:
                name = db.snapshot(cx, _stamp())
            st.success(f"Saved `{name}`.")
        bks = db.backups()
        if not bks:
            st.caption("No backups yet.")
        else:
            b = st.selectbox("Backup", bks, key="bk_pick")
            c1, c2 = st.columns(2)
            if c1.button("Restore this backup", key="bk_restore"):
                st.session_state["bk_confirm"] = b
            if c2.button("Delete this backup", key="bk_drop"):
                db.execute(f"DROP TABLE {b}")
                st.rerun()
            if st.session_state.get("bk_confirm") == b:
                st.warning(f"Restore `{b}`? Current parts are backed up first, then replaced.")
                if st.button("Yes, restore", type="primary", key="bk_yes"):
                    cur = db.restore(db.engine(), b, _stamp())
                    st.session_state.pop("bk_confirm", None)
                    _refresh()
                    st.success(f"Restored `{b}` (previous data saved as `{cur}`).")

    with t3:
        with st.form("proj"):
            name = st.text_input("Project name (shown in the sidebar, login and reports)",
                                 value=PROJECT)
            if st.form_submit_button("Save"):
                db.set_settings({"project_name": name.strip()}, _now_ts())
                st.rerun()


# --------------------------------------------------------------- Users
def page_users() -> None:
    st.header("Users & access")
    users = db.query("SELECT username, permission, created_at FROM user_credentials "
                     "ORDER BY username", ttl=0)
    st.dataframe(users.rename(columns={"username": "Username", "permission": "Access",
                                       "created_at": "Created"}),
                 use_container_width=True, hide_index=True)
    st.caption("Everyone signed in can view Overview, Assemblies and Reports. "
               "Stage tokens allow booking that stage; **Full access** adds Data admin, "
               "Users and Activity.")

    def _perm_ui(key: str, cur: str = "") -> str:
        full = st.checkbox("Full access (admin)", value=cur == "all", key=f"{key}_full")
        if full:
            return "all"
        have = [t for t in KNOWN_TOKENS if t in _toks(cur)]
        return ",".join(st.multiselect("Can update", KNOWN_TOKENS, default=have,
                                       key=f"{key}_toks"))

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Add user")
        u = st.text_input("Username", key="nu_u")
        p = st.text_input("Password", type="password", key="nu_p")
        perm = _perm_ui("nu")
        if st.button("Add user", type="primary", key="nu_go"):
            if not u.strip() or len(p) < 6:
                st.error("Username required; password at least 6 characters.")
            elif u.strip() in set(users["username"]):
                st.error("That username exists.")
            else:
                db.execute("""INSERT INTO user_credentials (username, password_hash, permission,
                              created_at) VALUES (:u, :h, :p, :t)""",
                           {"u": u.strip(), "h": db.hash_password(p), "p": perm, "t": _now_ts()})
                st.success(f"Added {u.strip()}.")
                st.rerun()
    with c2:
        st.subheader("Edit user")
        who = st.selectbox("User", users["username"].tolist(), key="eu_u")
        cur = str(users.set_index("username").loc[who, "permission"] or "")
        perm2 = _perm_ui(f"eu_{who}", cur)
        newpw = st.text_input("New password (leave blank to keep)", type="password",
                              key=f"eu_p_{who}")
        b1, b2 = st.columns(2)
        if b1.button("Save changes", type="primary", key="eu_go"):
            if who == st.session_state["user"] and perm2 != "all":
                st.error("You can't remove your own admin access.")
            elif newpw and len(newpw) < 6:
                st.error("Password at least 6 characters.")
            else:
                db.execute("UPDATE user_credentials SET permission = :p WHERE username = :u",
                           {"p": perm2, "u": who})
                if newpw:
                    db.set_password(who, newpw)
                st.success("Saved.")
                st.rerun()
        if b2.button("Delete user", key="eu_del", disabled=who == st.session_state["user"]):
            db.execute("DELETE FROM user_credentials WHERE username = :u", {"u": who})
            st.rerun()


# --------------------------------------------------------------- navigation
_PAGE_FUNCS = {
    "Overview": page_overview,
    "Update progress": page_update,
    "Assemblies": page_assemblies,
    "Reports": page_reports,
    "Activity": page_activity,
    "Data admin": page_admin,
    "Users": page_users,
}
_NAV_GROUPS = {
    "Dashboard": ["Overview", "Assemblies", "Reports"],
    "Shop floor": ["Update progress"],
    "Admin": ["Activity", "Data admin", "Users"],
}
_NAV_ICONS = {
    "Overview": ":material/space_dashboard:",
    "Update progress": ":material/edit_note:",
    "Assemblies": ":material/view_list:",
    "Reports": ":material/summarize:",
    "Activity": ":material/history:",
    "Data admin": ":material/database:",
    "Users": ":material/manage_accounts:",
}
assert set(_PAGE_FUNCS) == set(PAGE_PERMS) == set(_PAGE_NAMES)

_visible = [p for p in _PAGE_NAMES if can_see(p)]
_pages = {n: st.Page(_PAGE_FUNCS[n], title=n, icon=_NAV_ICONS[n], url_path=_slug(n),
                     default=(n == _visible[0])) for n in _visible}
pg = st.navigation(list(_pages.values()), position="hidden")

with st.sidebar:
    for _grp, _names in _NAV_GROUPS.items():
        _shown = [n for n in _names if n in _pages]
        if not _shown:
            continue
        st.markdown(f"<div class='tp-navgrp'>{_grp}</div>", unsafe_allow_html=True)
        for n in _shown:
            (st.container(key="tp_nav_on") if n == pg.title else st).page_link(
                _pages[n], label=n, icon=_NAV_ICONS[n])

    _u = str(st.session_state.get("user", ""))
    _role = "Admin" if is_admin() else (", ".join(sorted(_toks(st.session_state.get(
        "permission", "")))) or "View only")
    st.markdown(
        f"<div class='tp-acct'><div class='tp-av'>{html.escape(_u[:1].upper() or '?')}</div>"
        f"<div class='tp-who'><b>{html.escape(_u)}</b><span>{html.escape(_role)}</span>"
        "</div></div>", unsafe_allow_html=True)
    _c1, _c2 = st.columns([1, 1.25], vertical_alignment="center", gap="small")
    with _c1:
        _cur = "Dark" if is_dark() else "Light"
        _pick = st.segmented_control(
            "Appearance", ["Light", "Dark"], default=_cur, selection_mode="single",
            format_func=lambda o: ":material/light_mode:" if o == "Light" else ":material/dark_mode:",
            key="theme_seg", label_visibility="collapsed") or _cur
        if _pick.lower() != _cur.lower():
            st.session_state["theme_choice"] = _pick.lower()
            st.query_params["theme"] = _pick.lower()
            st.rerun()
    with _c2:
        if st.button("Sign out", use_container_width=True, key="signout"):
            _sign_out()
            st.rerun()
    with st.popover("Change password", icon=":material/key:", use_container_width=True):
        with st.form("chpw", clear_on_submit=True):
            o = st.text_input("Current password", type="password")
            n1 = st.text_input("New password", type="password")
            n2 = st.text_input("Repeat new password", type="password")
            if st.form_submit_button("Change"):
                if db.check_login(_u, o) is None:
                    st.error("Current password is wrong.")
                elif len(n1) < 6 or n1 != n2:
                    st.error("New passwords must match and be at least 6 characters.")
                else:
                    db.set_password(_u, n1)
                    st.success("Password changed.")

pg.run()

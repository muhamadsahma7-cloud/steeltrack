"""SteelTrack - steel structure fabrication progress, online.

Sister app of Team Piping: same sign-in, look and layout, built around the
steel master database (MASTER DATABASE 722-1.xlsx, sheet Progress).
Progress is booked as PIECES per SUB ASSEMBLY (a sub assembly can be made
many times over) through four stages:

    Fit-up -> Welding -> Blasting & painting (only if Painting status = Yes)
           -> Send to site (D.O.)

A stage's kg = pieces done x weight of one copy. NAEC Scope = NO rows are
kept but left out of progress.

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
.st-sow{background:var(--panel);border:1px solid var(--line);border-left:4px solid var(--accent);
  border-radius:12px;padding:14px 18px;margin:0 0 1.1rem}
.st-sow-h{font-size:.78rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;
  color:var(--accent);margin-bottom:.35rem}
.st-sow-b{white-space:pre-wrap;color:var(--text);line-height:1.55;font-size:.95rem}
.st-stage{font-size:.78rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;
  margin-bottom:.25rem}
"""
# stage-coloured metric edges + progress fills: containers keyed
# sg_/td_/wo_<stage> (stage cards, today's progress, work order figures)
_STATIC_CSS += "".join(
    f'.st-key-{pre}_{k} [data-testid="stMetric"]{{border-left-color:var(--st_{k})}}'
    f'.st-key-{pre}_{k} [data-testid="stProgress"] [role="progressbar"] div:empty{{background:var(--st_{k})}}'
    f'.st-key-{pre}_{k} .st-stage{{color:var(--st_{k})}}'
    for pre in ("sg", "td", "wo")
    for k in ("fitup", "welding", "paint", "site", "scope", "hold")
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


def donut(done: float, total: float, color: str):
    """Done-vs-balance ring with the % in the middle (as in Team Piping)."""
    d = is_dark()
    pct = done / total * 100 if total else 0
    src = pd.DataFrame({"part": ["Done", "Balance"],
                        "t": [done / 1000, max(total - done, 0.0) / 1000]})
    if not total:                                   # stage not needed: grey ring
        src = pd.DataFrame({"part": ["Balance"], "t": [1.0]})
    arc = (alt.Chart(src).mark_arc(innerRadius=52, outerRadius=76, cornerRadius=3)
           .encode(theta=alt.Theta("t:Q", stack=True),
                   color=alt.Color("part:N", legend=None,
                                   scale=alt.Scale(domain=["Done", "Balance"],
                                                   range=[color, "#2b3a52" if d else "#e5e9f0"])),
                   order=alt.Order("part:N", sort="descending"),
                   tooltip=[alt.Tooltip("part:N", title=""),
                            alt.Tooltip("t:Q", title="tonnes", format=",.2f")]))
    txt = (alt.Chart(pd.DataFrame({"t": [f"{pct:.1f}%" if total else "N/A"]}))
           .mark_text(size=21, fontWeight="bold", color="#e2e8f0" if d else "#1f2933")
           .encode(text="t:N"))
    return (arc + txt).properties(height=175)


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
    _seeded = db.ensure_schema(db.schema_version())
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
    "fitup":   dict(label="Fit-up", token="Update Fit-Up", do=False),
    "welding": dict(label="Welding", token="Update Welding", do=False),
    "paint":   dict(label="Blasting & painting", token="Update Painting", do=False),
    "site":    dict(label="Send to site", token="Update Site Delivery", do=True),
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
UK = ["work_order", "assembly_mark", "unit"]        # a sub assembly (progress unit)
LEVELS = ["Not started", "In progress", "Fit-up done", "Welded", "Ready for site", "At site"]


@st.cache_data(ttl=30, show_spinner=False)
def load_parts() -> pd.DataFrame:
    df = db.query("SELECT * FROM parts ORDER BY id", ttl=0)
    for c in db.PART_COLS:                  # tolerate a table not migrated yet
        if c not in df.columns:
            df[c] = None
    for c in ("work_order", "scope", "area", "assembly_mark", "sub_assembly", "part_mark",
              "name", "profile", "grade", "painting_req", "remark"):
        df[c] = df[c].fillna("").astype(str).str.strip()
    for c in ("weight_kg", "unit_weight", "length_mm", "surface_area", "total_weight"):
        df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    df["weight_kg"] = df["weight_kg"].fillna(0.0)
    df["priority"] = pd.to_numeric(df["priority"], errors="coerce").astype("Int64")
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(1).astype(int)
    df["asm_qty"] = pd.to_numeric(df["asm_qty"], errors="coerce").fillna(1).astype(int)
    df["unit"] = df["sub_assembly"].where(df["sub_assembly"] != "", df["assembly_mark"])
    df["paint_req"] = df["painting_req"].str.upper().isin(["YES", "Y"])
    # Welded? = NO: fitted up but not welded (purlin, plain beam, angle) - skips welding
    df["weld_req"] = df["welded"].fillna("").astype(str).str.strip().str.upper() != "NO"
    df["in_scope"] = df["scope"].str.upper() != "NO"
    return df


@st.cache_data(ttl=30, show_spinner=False)
def load_progress() -> pd.DataFrame:
    df = db.query("""SELECT id, work_order, assembly_mark, unit, stage, qty, date, do_no,
                            done_by, username, ts FROM progress ORDER BY id""", ttl=0)
    for c in ("work_order", "assembly_mark", "unit", "stage", "date", "do_no", "done_by",
              "username", "ts"):
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["date"] = df["date"].str[:10]
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce").fillna(0).astype(int)
    return df


def route(paint_req: bool, weld_req: bool) -> list[str]:
    """The stages a sub assembly goes through, in order."""
    return (["fitup"] + (["welding"] if weld_req else []) + (["paint"] if paint_req else [])
            + ["site"])


def progress_eff() -> pd.DataFrame:
    """Progress as the app counts it. A sub assembly with Welded? = NO is never
    booked for welding: it gets an automatic welding booking mirroring each
    fit-up booking (same pieces, same date), so welding %, daily output and
    the Excel export include it. Undoing the fit-up removes it too."""
    g = load_progress()
    p = load_parts()
    if g.empty or p.empty:
        return g
    w = p.groupby(UK)["weld_req"].any()
    nw = set(w[~w].index)
    if not nw:
        return g
    is_nw = pd.Series([k in nw for k in zip(g["work_order"], g["assembly_mark"], g["unit"])],
                      index=g.index)
    auto = g[is_nw & (g["stage"] == "fitup")].assign(
        stage="welding", done_by="auto (Welded? = NO)")
    auto["id"] = -auto["id"]
    return pd.concat([g[~(is_nw & (g["stage"] == "welding"))], auto], ignore_index=True)


def prev_stage(s: str, paint_req: bool, weld_req: bool) -> str | None:
    r = route(paint_req, weld_req)
    return r[r.index(s) - 1] if s in r and r.index(s) > 0 else None


def next_stage(s: str, paint_req: bool, weld_req: bool) -> str | None:
    r = route(paint_req, weld_req)
    return r[r.index(s) + 1] if s in r and r.index(s) + 1 < len(r) else None


def units_df(asof: str | None = None) -> pd.DataFrame:
    """One row per sub assembly: copies, weight, and copies done per stage
    (as of the end of `asof` when given; undated Excel bookings always count)."""
    p = load_parts()
    if p.empty:
        return pd.DataFrame()
    heavy = p.sort_values("weight_kg", ascending=False).drop_duplicates(UK).set_index(UK)
    u = p.groupby(UK, sort=False).agg(
        priority=("priority", "min"), area=("area", "first"), in_scope=("in_scope", "all"),
        paint_req=("paint_req", "any"), weld_req=("weld_req", "any"),
        parts=("part_mark", "size"), pcs=("qty", "sum"),
        copy_kg=("weight_kg", "sum"), copies=("asm_qty", "max"))
    u["main"] = (heavy["name"] + " " + heavy["profile"]).reindex(u.index)
    u["total_kg"] = u["copy_kg"] * u["copies"]
    g = progress_eff()
    if asof:
        g = g[(g["date"] == "") | (g["date"] <= asof)]
    done = g.groupby(UK + ["stage"])["qty"].sum().unstack("stage")
    last = g[g["date"] != ""].groupby(UK + ["stage"])["date"].max().unstack("stage")
    for s in STAGES:
        u[f"d_{s}"] = (done[s].reindex(u.index) if s in done else pd.Series(0, index=u.index)
                       ).fillna(0).astype(int)
        u[f"last_{s}"] = (last[s].reindex(u.index) if s in last else pd.Series("", index=u.index)
                          ).fillna("")
    for s in STAGES:
        u.loc[~applies(u, s), f"d_{s}"] = 0
    for s in STAGES:
        av = pd.Series(0, index=u.index)
        for pr in (False, True):
            for wr in (False, True):
                m = (u["paint_req"] == pr) & (u["weld_req"] == wr)
                if not m.any() or s not in route(pr, wr):
                    continue
                pv = prev_stage(s, pr, wr)
                prev = u.loc[m, "copies"] if pv is None else u.loc[m, f"d_{pv}"]
                av[m] = (prev - u.loc[m, f"d_{s}"]).clip(lower=0)
        u[f"av_{s}"] = av.astype(int)
    # a stage the unit doesn't need counts as complete
    full = lambda s: (u[f"d_{s}"] >= u["copies"]) | ~applies(u, s)
    lv = pd.Series(0, index=u.index)
    lv = lv.mask(sum(u[f"d_{s}"] for s in STAGES) > 0, 1)
    lv = lv.mask(full("fitup"), 2)
    lv = lv.mask(full("fitup") & full("welding"), 3)
    lv = lv.mask(full("fitup") & full("welding") & full("paint"), 4)
    lv = lv.mask(full("site"), 5)
    u["level"] = lv
    u["status"] = lv.map(dict(enumerate(LEVELS)))
    return u.reset_index()


def applies(u: pd.DataFrame, s: str) -> pd.Series:
    """Does each row (sub assembly or part) need stage s?"""
    if s == "paint":
        return u["paint_req"]
    return pd.Series(True, index=u.index)


def group_status(lv: pd.Series) -> str:
    lo, hi = int(lv.min()), int(lv.max())
    return LEVELS[1] if lo == 0 and hi > 0 else LEVELS[lo]


def events(u: pd.DataFrame) -> pd.DataFrame:
    """progress bookings with their weight (copies x one-copy weight)."""
    g = progress_eff()
    e = g.merge(u[UK + ["copy_kg", "in_scope", "area", "priority"]], on=UK, how="inner")
    e = e[e["in_scope"]]
    e["kg"] = e["qty"] * e["copy_kg"]
    e["label"] = e["stage"].map({k: v["label"] for k, v in STAGES.items()})
    return e


def part_frame(u: pd.DataFrame) -> pd.DataFrame:
    """parts carrying their sub assembly's copies + copies done per stage."""
    cols = UK + ["copies", "paint_req", "weld_req"] + [f"d_{s}" for s in STAGES]
    p = (load_parts().drop(columns=["paint_req", "weld_req"])
         .merge(u[cols], on=UK, how="inner"))
    return p


def stage_pcts(p: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    """weight + % complete per stage, grouped by `by` (p = part_frame)."""
    t = p.assign(_w=p["weight_kg"] * p["copies"])
    agg = {"Parts": ("weight_kg", "size"), "Weight (kg)": ("_w", "sum")}
    for s in STAGES:
        ap = applies(t, s)
        t[f"_a_{s}"] = t["_w"].where(ap, 0.0)
        t[f"_d_{s}"] = (t["weight_kg"] * t[f"d_{s}"]).where(ap, 0.0)
        agg[f"_a_{s}"] = (f"_a_{s}", "sum")
        agg[f"_d_{s}"] = (f"_d_{s}", "sum")
    g = t.groupby(by, dropna=False).agg(**agg)
    for s, cfg in STAGES.items():
        # blank (not 0%) where nothing in the group needs this stage
        g[cfg["label"] + " %"] = (g[f"_d_{s}"] / g[f"_a_{s}"].where(g[f"_a_{s}"] > 0) * 100
                                  ).clip(upper=100)
    return g.drop(columns=[c for c in g.columns if c.startswith("_")]).reset_index()


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
    if load_parts().empty:
        st.info("No parts loaded yet. An admin can import the master database "
                "Excel in **Data admin**.")
        return True
    return False


def _filters(u: pd.DataFrame, key: str, *, search: bool = True) -> pd.DataFrame:
    """Work order / area / priority / search filters, shared by the pages."""
    c = st.columns([1, 1.3, 1.2, 1.6] if search else [1, 1.3, 1.2])
    wo = c[0].multiselect("Work order", sorted(u["work_order"].unique()), key=f"{key}_wo",
                          placeholder="All")
    ar = c[1].multiselect("Area", sorted(u["area"].unique()), key=f"{key}_ar",
                          placeholder="All")
    pr = c[2].multiselect("Priority", sorted(u["priority"].dropna().astype(int).unique()),
                          key=f"{key}_pr", placeholder="All")
    q = c[3].text_input("Search assembly / sub assembly", key=f"{key}_q",
                        placeholder="e.g. CL1 or RFT5-2") if search else ""
    if wo:
        u = u[u["work_order"].isin(wo)]
    if ar:
        u = u[u["area"].isin(ar)]
    if pr:
        u = u[u["priority"].isin(pr)]
    if q.strip():
        ql = q.strip().lower()
        u = u[u["assembly_mark"].str.lower().str.contains(ql, regex=False)
              | u["unit"].str.lower().str.contains(ql, regex=False)]
    return u


# --------------------------------------------------------------- Overview
def page_overview() -> None:
    st.header("Project dashboard")
    sow = (SETTINGS.get("scope_of_work") or "").strip()
    if sow:
        st.markdown(
            "<div class='st-sow'><div class='st-sow-h'>Scope of work</div>"
            f"<div class='st-sow-b'>{html.escape(sow)}</div></div>", unsafe_allow_html=True)
    elif is_admin():
        st.caption("Tip: add the project's scope of work in Data admin → Project settings.")
    if _no_data_hint():
        return
    c1, _ = st.columns([1, 3])
    asof = c1.date_input("As of", value=_today(), max_value=_today(), format="DD/MM/YYYY",
                         help="See the project as it stood at the end of this day.")
    a = asof.isoformat()
    uall = _filters(units_df(a), "ov", search=False)
    u = uall[uall["in_scope"]]
    out = uall[~uall["in_scope"]]
    total = u["total_kg"].sum()
    ev = events(u)
    ev = ev[ev[UK].apply(tuple, axis=1).isin(set(u[UK].apply(tuple, axis=1)))] if len(ev) else ev
    dated = ev[(ev["date"] != "") & (ev["date"] <= a)]

    with st.container(key="sg_scope"):
        k = st.columns(4)
        k[0].metric("NAEC scope", t_(total), f"{u['assembly_mark'].nunique():,} assemblies",
                    delta_color="off")
        k[1].metric("Sub assemblies", f"{len(u):,}", f"{int(u['copies'].sum()):,} pieces to make",
                    delta_color="off")
        k[2].metric("Needs painting", t_(u.loc[u["paint_req"], "total_kg"].sum()),
                    delta_color="off")
        k[3].metric("Not NAEC scope", t_(out["total_kg"].sum()),
                    f"{len(out):,} sub assemblies · not counted", delta_color="off")

    st.subheader("Overall progress")
    week0 = (asof - timedelta(days=6)).isoformat()
    cols = st.columns(4)
    for col, (s, cfg) in zip(cols, STAGES.items()):
        ap = applies(u, s)
        scope = u.loc[ap, "total_kg"].sum()
        dn = (u.loc[ap, f"d_{s}"] * u.loc[ap, "copy_kg"]).sum()
        es = dated[dated["stage"] == s]
        wk = es.loc[es["date"] >= week0, "kg"].sum()
        with col.container(key=f"sg_{s}"):
            st.markdown(f"<div class='st-stage'>{cfg['label']}</div>", unsafe_allow_html=True)
            st.altair_chart(dark_alt(donut(dn, scope, tok(f"st_{s}"))),
                            use_container_width=True)
            st.metric("Done", t_(dn), f"balance {t_(scope - dn)}", delta_color="off")
            st.caption(f"Last 7 days **{wk / 1000:,.2f} t**" if scope
                       else "No parts need this stage.")

    _overview_today(asof, dated)
    _overview_work_orders(u)

    st.subheader("Assembly status")
    stat = u.groupby(["work_order", "assembly_mark"])["level"].apply(group_status)
    w = u.groupby(["work_order", "assembly_mark"])["total_kg"].sum()
    sdf = (pd.DataFrame({"status": stat, "w": w}).groupby("status")
           .agg(assemblies=("w", "size"), weight=("w", "sum"))
           .reindex(LEVELS).fillna(0).reset_index())
    m = st.columns(len(LEVELS))
    for col, r in zip(m, sdf.itertuples()):
        col.metric(r.status, f"{int(r.assemblies):,}", t_(r.weight), delta_color="off")

    order = [c["label"] for c in STAGES.values()]
    colors = [tok(f"st_{s}") for s in STAGES]
    st.subheader("Daily output")
    days = st.segmented_control("Window", [14, 30, 60, 90], default=30,
                                format_func=lambda n: f"{n} days", key="ov_win") or 30
    start = (asof - timedelta(days=days - 1)).isoformat()
    cd = (dated[dated["date"] >= start].groupby(["date", "label"])["kg"].sum().reset_index())
    if not cd.empty:
        cd["date"] = pd.to_datetime(cd["date"])
        ch = (alt.Chart(cd).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2)
              .encode(x=alt.X("yearmonthdate(date):T", title=None,
                              axis=alt.Axis(format="%d %b", labelAngle=0)),
                      xOffset=alt.XOffset("label:N", sort=order),
                      y=alt.Y("kg:Q", title="kg"),
                      color=alt.Color("label:N", sort=order, title=None,
                                      scale=alt.Scale(domain=order, range=colors)),
                      tooltip=[alt.Tooltip("yearmonthdate(date):T", title="Date",
                                           format="%a %d %b"),
                               alt.Tooltip("label:N", title="Stage"),
                               alt.Tooltip("kg:Q", format=",.1f")])
              .properties(height=300))
        st.altair_chart(dark_alt(ch), use_container_width=True)
    else:
        st.caption("No progress booked in this window.")

    st.subheader("Cumulative progress")
    rows = []
    for s, cfg in STAGES.items():
        base = ev.loc[(ev["stage"] == s) & (ev["date"] == ""), "kg"].sum()
        g = dated[dated["stage"] == s].groupby("date")["kg"].sum().sort_index().cumsum() + base
        rows += [{"date": d, "stage": cfg["label"], "t": v / 1000} for d, v in g.items()]
    if rows:
        cd = pd.DataFrame(rows)
        cd["date"] = pd.to_datetime(cd["date"])
        ch = (alt.Chart(cd).mark_line(point=True, interpolate="step-after")
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
        st.caption(f"Dashed line = NAEC scope {t_(total)}. Bookings imported from Excel "
                   "without a date are counted from the start.")

    pf = part_frame(u)
    t1, t2, t3 = st.tabs(["By area", "By priority", "By member type"])
    with t1:
        show_table(stage_pcts(pf, ["work_order", "area"]).rename(
            columns={"work_order": "WO", "area": "Area"}), "progress_by_area")
    with t2:
        p = stage_pcts(pf, ["priority"]).sort_values("priority", na_position="last")
        p["priority"] = p["priority"].astype("string").fillna("(none)")
        show_table(p.rename(columns={"priority": "Priority"}), "progress_by_priority")
    with t3:
        p = stage_pcts(pf, ["name"]).sort_values("Weight (kg)", ascending=False)
        show_table(p.rename(columns={"name": "Member type"}), "progress_by_member")


def _overview_today(asof: date, dated: pd.DataFrame) -> None:
    """What was booked on the As-of day (today by default), per stage."""
    a = asof.isoformat()
    st.subheader("Today's progress" if asof == _today() else f"Progress on {asof:%a %d %b %Y}")
    day = dated[dated["date"] == a]
    cols = st.columns(4)
    for col, (s, cfg) in zip(cols, STAGES.items()):
        x = day[day["stage"] == s]
        with col.container(key=f"td_{s}"):
            st.metric(cfg["label"], f"{x['kg'].sum():,.1f} kg",
                      f"{int(x['qty'].sum()):,} pcs · {x.groupby(UK).ngroups:,} sub assemblies",
                      delta_color="off")
    if day.empty:
        st.caption("Nothing booked on this day.")


def _overview_work_orders(u: pd.DataFrame) -> None:
    """Issued work orders and what is left to fit up / weld. A row with a Work
    order no counts as issued. Welding backlog = fitted up, not yet welded
    (Welded? = NO items are left out: their welding follows fit-up)."""
    st.subheader("Work orders")
    iss = u[u["work_order"] != ""].copy()
    nis = u[u["work_order"] == ""]
    if iss.empty:
        st.caption("No work order issued yet.")
        return
    iss["fit_bal"] = (iss["copies"] - iss["d_fitup"]).clip(lower=0)
    iss["weld_bal"] = (iss["copies"] - iss["d_welding"]).clip(lower=0)
    iss["backlog"] = (iss["d_fitup"] - iss["d_welding"]).clip(lower=0).where(iss["weld_req"], 0)
    for c in ("fit_bal", "weld_bal", "backlog"):
        iss[f"{c}_kg"] = iss[c] * iss["copy_kg"]
    n_wo = iss["work_order"].nunique()
    sub = f"{n_wo:,} WO{'s' if n_wo != 1 else ''} · {int(iss['copies'].sum()):,} pcs"
    if not nis.empty:
        sub += f" · not issued {t_(nis['total_kg'].sum())}"
    k = st.columns(4)
    for col, key, label, val, d in (
        (k[0], "wo_scope", "Work orders issued", t_(iss["total_kg"].sum()), sub),
        (k[1], "wo_fitup", "Balance to fit-up", t_(iss["fit_bal_kg"].sum()),
         f"{int(iss['fit_bal'].sum()):,} pcs"),
        (k[2], "wo_welding", "Balance to welding", t_(iss["weld_bal_kg"].sum()),
         f"{int(iss['weld_bal'].sum()):,} pcs"),
        (k[3], "wo_hold", "Welding backlog", t_(iss["backlog_kg"].sum()),
         f"{int(iss['backlog'].sum()):,} pcs fitted up, not welded"),
    ):
        with col.container(key=key):
            st.metric(label, val, d, delta_color="off")


# --------------------------------------------------------------- Update
_UNIT_SQL = "work_order = :wo AND assembly_mark = :am AND unit = :unit"
_PARTS_UNIT_SQL = ("coalesce(work_order,'') = :wo AND assembly_mark = :am "
                   "AND coalesce(nullif(trim(sub_assembly),''), assembly_mark) = :unit")


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
    u = units_df()
    u = u[u["in_scope"]]
    u = _filters(u, "up")
    u = u[u[f"av_{s}"] > 0]
    prereq = {"fitup": "", "welding": " (fit-up done, Welded? = YES; Welded? = NO items get "
                                     "their welding date from fit-up automatically)",
              "paint": " (Painting status = Yes; fit-up done, and welded unless Welded? = NO)",
              "site": " (finished every stage it needs)"}[s]
    st.caption(f"Sub assemblies with copies ready for **{cfg['label'].lower()}**{prereq}: "
               f"**{len(u):,}** · {int(u[f'av_{s}'].sum()):,} pieces · "
               f"{(u[f'av_{s}'] * u['copy_kg']).sum():,.1f} kg")

    if u.empty:
        st.success(f"Nothing is waiting for {cfg['label'].lower()} in this selection.")
    else:
        u = u.sort_values(["priority", "assembly_mark", "unit"], na_position="last")
        ver = st.session_state.setdefault("up_ver", 0)
        fill = st.checkbox("Book all available", key=f"up_fill_{ver}",
                           help="Fill every row's Book column with all copies available.")
        tbl = pd.DataFrame({
            "Book": (u[f"av_{s}"] if fill else u[f"av_{s}"] * 0).astype(int).values,
            "WO": u["work_order"].values, "Area": u["area"].values,
            "Assembly": u["assembly_mark"].values, "Sub assembly": u["unit"].values,
            "Priority": u["priority"].values, "Main member": u["main"].values,
            "Copies": u["copies"].values, "Done": u[f"d_{s}"].values,
            "Available": u[f"av_{s}"].values, "kg / copy": u["copy_kg"].round(2).values,
        })
        ed = st.data_editor(
            tbl, hide_index=True, use_container_width=True, height=420,
            disabled=[c for c in tbl.columns if c != "Book"],
            column_config={
                "Book": st.column_config.NumberColumn(
                    "Book", min_value=0, step=1, format="%d",
                    help="How many copies of this sub assembly to book now"),
                "kg / copy": st.column_config.NumberColumn(format="%.2f")},
            key=f"up_ed_{s}_{ver}_{fill}_{len(tbl)}")
        ed["Book"] = pd.to_numeric(ed["Book"], errors="coerce").fillna(0).astype(int)
        over = ed[ed["Book"] > ed["Available"]]
        picked = ed[ed["Book"] > 0]
        kg = (picked["Book"] * picked["kg / copy"]).sum()
        st.markdown(f"**To book: {len(picked):,} sub assemblies · {int(picked['Book'].sum()):,} "
                    f"pieces · {kg:,.1f} kg**")
        if not over.empty:
            st.error("More than available on: " + ", ".join(over["Sub assembly"].head(8))
                     + (" …" if len(over) > 8 else ""))

        with st.form("up_form"):
            c1, c2, c3 = st.columns(3)
            d = c1.date_input(f"{cfg['label']} date", value=_today(), max_value=_today(),
                              format="DD/MM/YYYY")
            do_no = (c2.text_input("D.O. No.", placeholder="e.g. NAEC-722-DHEWORKSHOP-002")
                     if cfg["do"] else None)
            who = c3.text_input("Done by / supervisor", value=st.session_state["user"])
            go = st.form_submit_button(f"Confirm {cfg['label'].lower()}", type="primary",
                                       use_container_width=True)
        if go:
            if picked.empty:
                st.error("Enter how many copies to book in the Book column.")
            elif not over.empty:
                st.error("Fix the rows booked above what is available.")
            elif cfg["do"] and not (do_no or "").strip():
                st.error("Enter the D.O. number.")
            else:
                _book(s, picked, u, d.isoformat(), (do_no or "").strip() or None, who.strip())

    if st.session_state.get("up_msg"):
        st.success(st.session_state.pop("up_msg"))

    with st.expander(f"Undo — remove a {cfg['label'].lower()} booking"):
        _undo_ui(s)


def _book(s: str, picked: pd.DataFrame, u: pd.DataFrame, d: str, do_no: str | None,
          who: str) -> None:
    cfg = STAGES[s]
    ts = _now_ts()
    req = u.set_index(UK)[["paint_req", "weld_req"]]
    steps, logs = [], []
    for r in picked.to_dict("records"):
        k = (r["WO"], r["Assembly"], r["Sub assembly"])
        pv = prev_stage(s, bool(req.loc[k, "paint_req"]), bool(req.loc[k, "weld_req"]))
        prev_sql = (f"(SELECT coalesce(max(asm_qty),1) FROM parts WHERE {_PARTS_UNIT_SQL})"
                    if pv is None else
                    f"(SELECT coalesce(sum(qty),0) FROM progress WHERE {_UNIT_SQL} AND stage = :pv)")
        p = {"wo": k[0], "am": k[1], "unit": k[2], "s": s, "pv": pv, "q": int(r["Book"]),
             "d": d, "do": do_no, "by": who or None, "u": st.session_state["user"], "t": ts}
        # guarded: only books what the previous stage still has available, so
        # two people booking the same sub assembly can't overshoot
        steps.append((f"""INSERT INTO progress (work_order, assembly_mark, unit, stage, qty,
                                                date, do_no, done_by, username, ts)
                          SELECT :wo, :am, :unit, :s, :q, :d, :do, :by, :u, :t
                          WHERE {prev_sql} - (SELECT coalesce(sum(qty),0) FROM progress
                                              WHERE {_UNIT_SQL} AND stage = :s) >= :q""", p))
        logs.append((r, p))
    n = db.transaction(steps)
    log = [("""INSERT INTO progress_log (ts, username, stage, action, work_order, assembly_mark,
                                         unit, date_value, do_no, parts, weight_kg)
               VALUES (:t, :u, :sl, 'book', :wo, :am, :unit, :d, :do, :q, :w)""",
            {**p, "sl": cfg["label"], "w": float(r["Book"] * r["kg / copy"])})
           for (r, p), k in zip(logs, n) if k]
    if log:
        db.transaction(log)
    ok = [r for (r, _), k in zip(logs, n) if k]
    _refresh()
    st.session_state["up_ver"] = st.session_state.get("up_ver", 0) + 1     # fresh table
    st.session_state["up_msg"] = (
        f"{cfg['label']} booked: {len(ok)} sub assemblies, "
        f"{sum(r['Book'] for r in ok):,} pieces, "
        f"{sum(r['Book'] * r['kg / copy'] for r in ok):,.1f} kg on {d}"
        + (f", D.O. {do_no}" if do_no else "")
        + ("" if len(ok) == len(picked) else
           f". {len(picked) - len(ok)} not booked: someone else booked them first — check again."))
    st.rerun()


def _undo_ui(s: str) -> None:
    cfg = STAGES[s]
    g = load_progress()
    g = g[g["stage"] == s]
    if g.empty:
        st.caption("No bookings for this stage yet.")
        return
    u = units_df().set_index(UK)
    g = g.join(u[["paint_req", "weld_req"] + [f"d_{x}" for x in STAGES]], on=UK)
    if s == "welding":                      # Welded? = NO welding is automatic
        g = g[g["weld_req"].fillna(True).astype(bool)]
    g["next"] = [next_stage(s, bool(pr), bool(wr)) for pr, wr in
                 zip(g["paint_req"].fillna(False), g["weld_req"].fillna(True))]
    g["room"] = [r[f"d_{s}"] - (r[f"d_{r['next']}"] if r["next"] else 0)
                 for _, r in g.iterrows()]
    g = g[g["qty"] <= g["room"]].sort_values("id", ascending=False)
    if g.empty:
        st.caption("Nothing can be removed: the next stage already uses these pieces. "
                   "Remove that stage's booking first.")
        return
    g["label"] = g.apply(lambda r: f"{r['unit']}  ·  {r['qty']} pcs  ·  "
                         f"{r['date'] or 'no date'}"
                         + (f"  ·  D.O. {r['do_no']}" if r["do_no"] else "")
                         + f"  ·  by {r['done_by'] or r['username']}", axis=1)
    pick = st.multiselect("Bookings to remove", g["id"].tolist(),
                          format_func=lambda i: g.set_index("id").loc[i, "label"],
                          key=f"undo_{s}")
    if st.button("Remove selected bookings", key=f"undo_btn_{s}", disabled=not pick):
        ts = _now_ts()
        steps = []
        for i in pick:
            r = g.set_index("id").loc[i]
            p = {"id": int(i), "wo": r["work_order"], "am": r["assembly_mark"],
                 "unit": r["unit"], "s": s, "n": r["next"] or "-", "t": ts,
                 "u": st.session_state["user"]}
            steps.append((f"""DELETE FROM progress WHERE id = :id
                AND (SELECT coalesce(sum(qty),0) FROM progress WHERE {_UNIT_SQL} AND stage = :s)
                    - qty >= (SELECT coalesce(sum(qty),0) FROM progress
                              WHERE {_UNIT_SQL} AND stage = :n)""", p))
            steps.append(("""INSERT INTO progress_log (ts, username, stage, action, work_order,
                             assembly_mark, unit, date_value, do_no, parts, weight_kg)
                             VALUES (:t, :u, :sl, 'remove', :wo, :am, :unit, :d, :do, :q, NULL)""",
                          {**p, "sl": cfg["label"], "d": r["date"] or None,
                           "do": r["do_no"] or None, "q": int(r["qty"])}))
        db.transaction(steps)
        _refresh()
        st.session_state["up_msg"] = f"Removed {len(pick)} {cfg['label'].lower()} booking(s)."
        st.rerun()


# --------------------------------------------------------------- Assemblies
def page_assemblies() -> None:
    st.header("Assemblies")
    if _no_data_hint():
        return
    u = units_df()
    u = _filters(u, "as")
    c1, c2 = st.columns([2, 1])
    stt = c1.multiselect("Status", LEVELS + ["Not NAEC scope"], key="as_st", placeholder="All")
    keys = ["work_order", "area", "assembly_mark"]
    pf = part_frame(u)
    a = stage_pcts(pf, keys)
    grp = u.groupby(keys)
    a["Status"] = grp["level"].apply(group_status).values
    a.loc[~grp["in_scope"].all().values, "Status"] = "Not NAEC scope"
    a["Priority"] = grp["priority"].min().values
    a["Sub assemblies"] = grp["unit"].size().values
    a["Pieces"] = grp["copies"].sum().values
    last = u[[f"last_{s}" for s in STAGES]].max(axis=1)
    a["Last update"] = last.groupby([u[k] for k in keys]).max().values
    if stt:
        a = a[a["Status"].isin(stt)]
    a = a[keys + ["Priority", "Status", "Sub assemblies", "Pieces", "Weight (kg)"] + PCT_COLS
          + ["Last update"]].sort_values(["Priority", "assembly_mark"], na_position="last")
    st.caption(f"{len(a):,} assemblies · {a['Weight (kg)'].sum() / 1000:,.2f} t")
    show_table(a.rename(columns={"work_order": "WO", "area": "Area",
                                 "assembly_mark": "Assembly"}), "assemblies", height=420)

    st.subheader("Assembly detail")
    if a.empty:
        return
    pick = st.selectbox("Assembly", a["assembly_mark"].tolist(), key="as_pick")
    x = u[u["assembly_mark"] == pick]
    det = pd.DataFrame({"Sub assembly": x["unit"], "Main member": x["main"],
                        "Copies": x["copies"], "kg / copy": x["copy_kg"].round(2),
                        "Total (kg)": x["total_kg"].round(2), "Status": x["status"]})
    for s, cfg in STAGES.items():
        det[cfg["label"]] = [f"{d}/{c}" if ap else "—"
                             for d, c, ap in zip(x[f"d_{s}"], x["copies"], applies(x, s))]
    st.markdown("**Sub assemblies** (pieces done / pieces to make)")
    show_table(det, f"subassemblies_{pick}")
    st.markdown("**Parts (one copy)**")
    p = load_parts()
    p = p[p["assembly_mark"] == pick]
    show_table(p[["unit", "part_mark", "qty", "name", "profile", "grade", "length_mm",
                  "unit_weight", "weight_kg", "asm_qty", "painting_req", "scope"]].rename(
        columns={"unit": "Sub assembly", "part_mark": "Part", "qty": "No.", "name": "Name",
                 "profile": "Profile", "grade": "Material", "length_mm": "Length (mm)",
                 "unit_weight": "kg / pc", "weight_kg": "kg / copy", "asm_qty": "Copies",
                 "painting_req": "Painting", "scope": "NAEC scope"}), f"parts_{pick}")


# --------------------------------------------------------------- Reports
def page_reports() -> None:
    st.header("Reports")
    if _no_data_hint():
        return
    u = units_df()
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
        ev = events(u)
        ev = ev[(ev["date"] >= s_) & (ev["date"] <= e_)]
        k = st.columns(4)
        for col, (s, cfg) in zip(k, STAGES.items()):
            x = ev[ev["stage"] == s]
            with col.container(key=f"sg_{s}"):
                st.metric(cfg["label"], f"{x['kg'].sum():,.1f} kg",
                          f"{int(x['qty'].sum()):,} pieces", delta_color="off")
        dd = (ev.pivot_table(index="date", columns="label", values="kg", aggfunc="sum")
              .reindex(columns=[c["label"] for c in STAGES.values()]).fillna(0.0))
        dd.columns = [f"{c} (kg)" for c in dd.columns]
        dd = dd.reset_index().rename(columns={"date": "Date"})
        if not dd.empty:
            tot = {"Date": "TOTAL", **{c: dd[c].sum() for c in dd.columns if c != "Date"}}
            dd = pd.concat([dd, pd.DataFrame([tot])], ignore_index=True)
        st.markdown("**Daily output (kg)**")
        show_table(dd, f"daily_{s_}_{e_}", download=False)
        det = (ev.sort_values(["date", "stage", "assembly_mark", "unit"])
               [["label", "date", "work_order", "area", "assembly_mark", "unit", "qty", "kg",
                 "do_no", "done_by"]]
               .rename(columns={"label": "Stage", "date": "Date", "work_order": "WO",
                                "area": "Area", "assembly_mark": "Assembly",
                                "unit": "Sub assembly", "qty": "Pieces", "kg": "Weight (kg)",
                                "do_no": "D.O. No.", "done_by": "Done by"}))
        if not det.empty:
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
        ev = events(u)
        x = ev[(ev["stage"] == "site") & (ev["do_no"] != "")]
        if x.empty:
            st.caption("No delivery order recorded yet.")
        else:
            reg = (x.groupby("do_no").agg(
                date=("date", "max"), assemblies=("assembly_mark", "nunique"),
                subs=("unit", "nunique"), pcs=("qty", "sum"), kg=("kg", "sum")).reset_index()
                .rename(columns={"do_no": "D.O. No.", "date": "Date", "assemblies": "Assemblies",
                                 "subs": "Sub assemblies", "pcs": "Pieces", "kg": "Weight (kg)"})
                .sort_values(["Date", "D.O. No."], ascending=False))
            show_table(reg, "do_register")
            pick = st.selectbox("D.O. contents", reg["D.O. No."].tolist(), key="do_pick")
            y = x[x["do_no"] == pick]
            lst = (y.groupby(["work_order", "area", "assembly_mark", "unit"])
                   .agg(pcs=("qty", "sum"), kg=("kg", "sum")).reset_index()
                   .rename(columns={"work_order": "WO", "area": "Area", "assembly_mark": "Assembly",
                                    "unit": "Sub assembly", "pcs": "Pieces",
                                    "kg": "Weight (kg)"}))
            show_table(lst, f"do_{pick}", download=False)
            heads = [f"Delivery list — D.O. {pick}",
                     f"Date: {y['date'].max()}   ·   {int(lst['Pieces'].sum())} pieces   ·   "
                     f"{y['kg'].sum():,.2f} kg" + (f"   ·   {PROJECT}" if PROJECT else "")]
            st.download_button("⬇ Delivery list (.xlsx)",
                               reports.workbook({"Delivery list": lst}, {"Weight (kg)": "#,##0.00"},
                                                {"Delivery list": heads}),
                               file_name=f"DO_{pick}.xlsx", key="do_xlsx")

    with t3:
        st.write("Download the whole database in the same layout as the master Excel "
                 "(sheet *Progress*). Stage kg = row weight × pieces done; date = last booking.")
        if st.button("Prepare Excel file", key="mx_prep"):
            st.session_state["mx_bytes"] = reports.master_xlsx(_export_frame(u))
        if st.session_state.get("mx_bytes"):
            st.download_button("⬇ MASTER DATABASE.xlsx", st.session_state["mx_bytes"],
                               type="primary",
                               file_name=f"MASTER DATABASE {_today():%Y%m%d}.xlsx", key="mx_dl")


def _export_frame(u: pd.DataFrame) -> pd.DataFrame:
    g = progress_eff()
    dos = (g[(g["stage"] == "site") & (g["do_no"] != "")].groupby(UK)["do_no"]
           .agg(lambda s: ", ".join(dict.fromkeys(s))).rename("site_dos"))
    cols = UK + [f"d_{s}" for s in STAGES] + [f"last_{s}" for s in STAGES]
    p = load_parts().merge(u[cols], on=UK, how="left").merge(
        dos.reset_index(), on=UK, how="left")
    p["site_dos"] = p["site_dos"].fillna("")
    return p


# --------------------------------------------------------------- Activity
def page_activity() -> None:
    st.header("Activity")
    t1, t2 = st.tabs(["Progress log", "Sign-ins"])
    with t1:
        lg = db.query("""SELECT ts AS "Time", username AS "User", stage AS "Stage",
                                action AS "Action", work_order AS "WO",
                                assembly_mark AS "Assembly", unit AS "Sub assembly",
                                date_value AS "Date", do_no AS "D.O.", parts AS "Pieces",
                                weight_kg AS "Weight (kg)"
                         FROM progress_log ORDER BY id DESC LIMIT 2000""", ttl=0)
        who = st.multiselect("User", sorted(lg["User"].dropna().unique()), key="ac_u")
        if who:
            lg = lg[lg["User"].isin(who)]
        show_table(lg, "progress_log", height=480)
    with t2:
        show_table(db.query("""SELECT ts AS "Time", username AS "User" FROM user_log
                               ORDER BY id DESC LIMIT 1000""", ttl=0), "sign_ins")


# --------------------------------------------------------------- Data admin
def page_admin() -> None:
    st.header("Data admin")
    t1, t2, t3 = st.tabs(["Import master database", "Backups", "Project settings"])
    with t1:
        st.write("Upload the master database Excel (sheet **Progress**). "
                 "**No. of Set** = how many of each sub assembly to make; when it's empty "
                 "the app works it out from TOTAL WEIGHT.")
        up = st.file_uploader("Excel file", type=["xlsx", "xlsm"], key="imp_file")
        if up is not None:
            try:
                parts, prog, warn = reports.parse_master(up)
            except Exception as e:
                st.error(f"Can't read that file: {e}")
                return
            pu = parts.assign(unit=parts["sub_assembly"].fillna(parts["assembly_mark"]))
            un = pu.groupby(["work_order", "assembly_mark", "unit"]).agg(
                kg=("weight_kg", "sum"), q=("asm_qty", "first"), sc=("scope", "first"))
            un["t"] = un["kg"] * un["q"]
            k = st.columns(4)
            k[0].metric("Rows", f"{len(parts):,}",
                        f"{parts['assembly_mark'].nunique():,} assemblies", delta_color="off")
            k[1].metric("Sub assemblies", f"{len(un):,}", f"{int(un['q'].sum()):,} pieces",
                        delta_color="off")
            k[2].metric("NAEC scope", t_(un.loc[un["sc"] != "NO", "t"].sum()),
                        f"not NAEC: {t_(un.loc[un['sc'] == 'NO', 't'].sum())}",
                        delta_color="off")
            k[3].metric("Progress in the file", f"{len(prog):,} records",
                        f"{int((prog['date'].isna()).sum()) if len(prog) else 0} without date",
                        delta_color="off")
            if warn:
                with st.expander(f"⚠ {len(warn)} thing(s) to check in the file", expanded=True):
                    for w_ in warn:
                        st.write("• " + w_)
            with st.expander("Preview"):
                st.dataframe(parts.head(100), use_container_width=True, hide_index=True)
                if len(prog):
                    st.dataframe(prog, use_container_width=True, hide_index=True)
            existing = len(load_parts())
            if existing:
                # sets lowered below what is already booked -> shows over 100%
                cur = units_df()
                cur["booked"] = cur[[f"d_{x}" for x in STAGES]].max(axis=1)
                nq = un["q"].rename("new").reset_index()
                chk = cur[UK + ["copies", "booked"]].merge(nq, on=UK, how="inner")
                low = chk[chk["new"] < chk["booked"]]
                if not low.empty:
                    st.error("No. of Set in this file is lower than the pieces already booked "
                             "for: " + ", ".join(f"{r.unit} ({r.booked} booked, file says "
                                                 f"{r.new})" for r in low.itertuples())
                             + ". Fix the file, or remove those bookings first (Update progress → Undo).")
            mode = st.radio(
                "Import mode", ["merge", "replace"],
                format_func=lambda m: {
                    "merge": "Add new parts & update details (keeps progress booked in the app)",
                    "replace": "Replace everything with this file (progress comes from the file)",
                }[m], key="imp_mode", index=0 if existing else 1)
            ok = True
            if mode == "replace" and existing:
                st.warning(f"This deletes all {existing:,} parts and all progress bookings now "
                           "in the database (a backup copy is made first).")
                ok = st.text_input("Type REPLACE to confirm", key="imp_conf") == "REPLACE"
            if st.button("Import", type="primary", disabled=not ok, key="imp_go"):
                with st.spinner("Importing…"):
                    msg = db.import_parts(db.engine(), parts, prog, mode, _now_ts(), _stamp())
                _refresh()
                st.success(msg)

    with t2:
        st.caption("A backup copy of parts + progress is taken automatically before every "
                   "import and restore.")
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
                db.drop_backup(b)
                st.rerun()
            if st.session_state.get("bk_confirm") == b:
                st.warning(f"Restore `{b}`? Current data is backed up first, then replaced.")
                if st.button("Yes, restore", type="primary", key="bk_yes"):
                    cur = db.restore(db.engine(), b, _stamp())
                    st.session_state.pop("bk_confirm", None)
                    _refresh()
                    st.success(f"Restored `{b}` (previous data saved as `{cur}`).")

    with t3:
        with st.form("proj"):
            name = st.text_input("Project name (shown in the sidebar, login and reports)",
                                 value=PROJECT)
            scope = st.text_area(
                "Scope of work (shown on the Overview page)",
                value=SETTINGS.get("scope_of_work", ""), height=220,
                placeholder="e.g. Fabrication, blasting & painting and delivery to site of "
                            "the CLS structure and lower roof steelwork. "
                            "One item per line works well.")
            if st.form_submit_button("Save"):
                db.set_settings({"project_name": name.strip(),
                                 "scope_of_work": scope.strip()}, _now_ts())
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

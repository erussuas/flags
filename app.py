import hashlib, io
from collections import Counter

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from ec_parser import (
    # Flag metadata
    FLAG_META, PRIORITY_ORDER, PRIORITY_COLOR, CATEGORY_COLOR,
    # Energy / emissions
    ENERGY_COMMODITIES, GHG_SCOPE, DISPLAY_UNITS, from_kwh,
    # Configuration health
    CONFIG_CHECK_META, SEVERITY_ORDER, run_config_checks,
    # Continuity audit
    CONTINUITY_CHECK_META, run_continuity_checks,
    # Flag file parsing (auto-detects R27 or custom flat format)
    detect_and_parse_flags,
    # Utilities
    period_sort_key,
)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="EnergyCAP Data Quality",
    page_icon="⚑", layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# CACHE LAYER — each file parsed once, keyed by content hash
# No nested @st.cache_data calls (Streamlit limitation)
# ══════════════════════════════════════════════════════════════════════════════
def _fhash(b): return hashlib.md5(b).hexdigest()

@st.cache_data(show_spinner=False)
def cached_r03(fhk, fb):
    import tempfile, os
    from ec_parser import parse_report03
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        f.write(fb); p = f.name
    try:    return parse_report03(p)
    finally: os.unlink(p)

@st.cache_data(show_spinner=False)
def cached_cfg(fhk, fb):
    import tempfile, os
    from ec_parser import parse_report03, run_config_checks
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        f.write(fb); p = f.name
    try:    return run_config_checks(parse_report03(p))
    finally: os.unlink(p)

@st.cache_data(show_spinner=False)
def cached_r18(fhk, fb, shk, sb):
    import tempfile, os
    from ec_parser import parse_report18, parse_report03
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        f.write(fb); p18 = f.name
    sdf = None
    if sb:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
            f.write(sb); p03 = f.name
        try:    sdf = parse_report03(p03)
        finally: os.unlink(p03)
    try:    return parse_report18(p18, sdf)
    finally: os.unlink(p18)

@st.cache_data(show_spinner=False)
def cached_r27(fhk, fb):
    """
    Auto-detects and parses flag files — supports both Report-27 (hierarchical)
    and the flat custom format (one row per issue).
    Returns (DataFrame, format_name) tuple.
    """
    import tempfile, os
    from ec_parser import detect_and_parse_flags
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        f.write(fb); p = f.name
    try:    return detect_and_parse_flags(p)
    finally: os.unlink(p)

@st.cache_data(show_spinner=False)
def cached_r19(fhk, fb, shk, sb):
    import tempfile, os
    from ec_parser import parse_report19, parse_report03
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        f.write(fb); p19 = f.name
    sdf = None
    if sb:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
            f.write(sb); p03 = f.name
        try:    sdf = parse_report03(p03)
        finally: os.unlink(p03)
    try:    return parse_report19(p19, sdf)
    finally: os.unlink(p19)

@st.cache_data(show_spinner=False)
def cached_continuity(r19_hashes_key, r19_bytes_list, shk, sb):
    """Parse and merge multiple R19 files, then run continuity checks."""
    import tempfile, os
    from ec_parser import parse_report19, parse_report03, run_continuity_checks
    sdf = None
    if sb:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
            f.write(sb); p03 = f.name
        try:    sdf = parse_report03(p03)
        finally: os.unlink(p03)
    frames = []
    for fb in r19_bytes_list:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
            f.write(fb); p = f.name
        try:    frames.append(parse_report19(p, sdf))
        finally: os.unlink(p)
    df19 = (pd.concat(frames, ignore_index=True).drop_duplicates(['sheet','period'])
            if frames else pd.DataFrame())
    return run_continuity_checks(df19), df19


# ══════════════════════════════════════════════════════════════════════════════
# UI HELPERS
# ══════════════════════════════════════════════════════════════════════════════
PT = dict(font_family="Inter,system-ui,sans-serif",
          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")

SORT_OPTS = ["Cost (high→low)","Cost (low→high)",
             "Usage kWh (high→low)","Days to resolve (longest)","Bill ID","Vendor A→Z"]
SORT_MAP  = {
    "Cost (high→low)":           ("cost", False),
    "Cost (low→high)":           ("cost", True),
    "Usage kWh (high→low)":      ("total_kwh_equivalent", False),
    "Days to resolve (longest)": ("days_to_resolve", False),
    "Bill ID":                   ("bill_id", True),
    "Vendor A→Z":                ("vendor", True),
}
SEV_ICON   = {"High":"🔴","Medium":"🟠","Low":"🟡"}
SEV_COLOR  = {"High":"red","Medium":"orange","Low":"blue"}
ISSUE_PRI_COLOR = {"High":"red","Medium":"orange","Low":"green"}

def priority_icon(p):
    return {"High":"🔴","Medium":"🟠","Low":"🟢"}.get(p,"⚪")

def bar_h(data, title, color="#0d6efd", height=320, highlight=None):
    labels, values = list(data.keys()), list(data.values())
    colors = ["#ff6b35" if highlight and l==highlight else color for l in labels]
    fig = go.Figure(go.Bar(x=values, y=labels, orientation="h", marker_color=colors,
                           text=values, textposition="outside",
                           hovertemplate="%{y}: %{x}<extra></extra>"))
    fig.update_layout(title=title, height=height, yaxis_autorange="reversed",
                      xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                      margin=dict(l=10,r=50,t=40,b=10),
                      showlegend=False, clickmode="event+select", **PT)
    return fig

def bar_v(data, title, color="#0d6efd", height=320, highlight=None):
    labels, values = list(data.keys()), list(data.values())
    colors = ["#ff6b35" if highlight and l==highlight else color for l in labels]
    fig = go.Figure(go.Bar(x=labels, y=values, marker_color=colors,
                           text=values, textposition="outside",
                           hovertemplate="%{x}: %{y}<extra></extra>"))
    fig.update_layout(title=title, height=height,
                      yaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                      margin=dict(l=10,r=10,t=40,b=10),
                      showlegend=False, clickmode="event+select", **PT)
    return fig

def donut(data, title, colors=None, height=280):
    colors = colors or px.colors.qualitative.Set2
    fig = go.Figure(go.Pie(labels=list(data.keys()), values=list(data.values()),
                           hole=0.55, marker_colors=colors,
                           textinfo="percent+label", textfont_size=11,
                           hovertemplate="%{label}: %{value}<extra></extra>"))
    fig.update_layout(title=title, height=height,
                      margin=dict(l=10,r=10,t=40,b=10), showlegend=False, **PT)
    return fig

def extract_click(ev):
    if not ev: return None
    pts = ev.get("selection", {}).get("points", [])
    if not pts: return None
    return pts[0].get("y") or pts[0].get("x") or None

def issue_badges(issues_list):
    if not issues_list: return
    cols = st.columns(min(len(issues_list), 4))
    for i, issue in enumerate(issues_list):
        p = FLAG_META.get(issue, {}).get("priority", "Medium")
        with cols[i % 4]:
            st.badge(issue, color=ISSUE_PRI_COLOR.get(p, "gray"))

def render_bill_cards(sub_df, display_unit="kWh", max_cards=50):
    if sub_df.empty:
        st.info("No bills match this selection.")
        return
    st.caption(f"Showing {min(len(sub_df), max_cards)} of {len(sub_df)} bills")
    for _, row in sub_df.head(max_cards).iterrows():
        status_ok = row["status"] == "Resolved"
        days      = f"{int(row['days_to_resolve'])}d" if pd.notna(row.get("days_to_resolve")) else "—"
        resolver  = row["resolvers"][-1] if row["resolvers"] else "—"
        with st.container(border=True):
            h1, h2 = st.columns([4,1])
            with h1: st.markdown(f"**Bill {row['bill_id']}** &nbsp; {row['vendor']}")
            with h2: st.markdown(f"**${row['cost']:,.2f}**")
            m1,m2,m3,m4 = st.columns(4)
            m1.caption(f"📅 {row.get('billing_period_label','—')}")
            m2.caption(f"👤 {row.get('assigned_to','—')}")
            m3.caption(f"✓ {resolver} · {days}")
            m4.caption("✅ Resolved" if status_ok else "🔴 Unresolved")
            if pd.notna(row.get("total_kwh_equivalent")):
                val   = from_kwh(float(row["total_kwh_equivalent"]), display_unit)
                cat   = row.get("energy_category", "")
                scope = row.get("ghg_scope", "")
                parts = [f"⚡ {val:,.1f} {display_unit}"]
                if cat:   parts.append(cat)
                if scope: parts.append(f"({scope})")
                st.caption(" · ".join(parts))
            issue_badges(row["issues_list"])

def drill_banner(label, key):
    c1, c2 = st.columns([11,1])
    with c1: st.info(f"🔍 Drill-down active: **{label}**")
    with c2:
        st.write("")
        if st.button("✕", key=f"clr_{key}"):
            st.session_state[key] = None
            st.rerun()

def bill_detail_panel(row):
    st.subheader(f"Bill {row['bill_id']} — {row['vendor']}")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Account:** {row['account']}")
        st.markdown(f"**Period:** {row.get('billing_period_label','—')}")
        st.markdown(f"**Cost:** ${row['cost']:,.2f}  |  **Recovery:** ${row['cost_recovery']:,.2f}")
        st.markdown(f"**Assigned to:** {row.get('assigned_to','—')}")
        if row.get("assignee_org") and str(row.get("assignee_org")) not in ("","None","nan"):
            st.markdown(f"**Assignee org:** {row['assignee_org']}")
        if pd.notna(row.get("pending_days")):
            bucket = row.get("bucket_days","")
            st.markdown(f"**Days open:** {int(row['pending_days'])} days"
                        f"{' · ' + bucket if bucket and bucket not in ('None','nan') else ''}")
        if row.get("responsibility") and str(row.get("responsibility")) not in ("","None","nan"):
            st.markdown(f"**Responsibility:** {row['responsibility']}")
        if pd.notna(row.get("total_kwh_equivalent")):
            st.markdown(f"**Energy usage:** {float(row['total_kwh_equivalent']):,.0f} kWh equiv.")
            if row.get("energy_category"):
                st.markdown(f"**Energy type:** {row['energy_category']} ({row.get('ghg_scope','')})")
    with c2:
        st.markdown(f"**Status:** {row['status']}")
        days = f"{int(row['days_to_resolve'])}d" if pd.notna(row.get('days_to_resolve')) else "Pending"
        st.markdown(f"**Days to resolve:** {days}")
        if row.get("resolvers"):
            st.markdown(f"**Resolved by:** {', '.join(row['resolvers'])}")
    # Show EnergyCAP system notes if available (custom format)
    if row.get("notes") and str(row.get("notes")) not in ("","None","nan"):
        with st.expander("📋 EnergyCAP audit notes"):
            for note in str(row["notes"]).split(";"):
                note = note.strip()
                if note: st.caption(f"• {note}")

    st.markdown("**Flag issues & guidance:**")
    for issue in row["issues_list"]:
        m = FLAG_META.get(issue, {})
        with st.expander(f"{priority_icon(m.get('priority',''))} {issue} — {m.get('category','')}"):
            if m:
                st.markdown(f"**Cause:** {m['cause']}")
                st.markdown(f"**Action:** {m['action']}")
                st.markdown(f"**In EnergyCAP:** {m['in_energycap']}")

def check_card(check, key, dismiss_set):
    """Render a configuration/continuity check card."""
    sev  = check['severity']
    icon = SEV_ICON.get(sev, "⚪")
    cnt  = check['count']
    dismissed = key in dismiss_set

    with st.expander(
        f"{icon} **{check['title']}** — {cnt} {'item' if cnt==1 else 'items'}"
        f" {'(reviewed)' if dismissed else f'· {sev}'}",
        expanded=(not dismissed and sev == "High")
    ):
        if dismissed:
            st.caption("✅ Marked as reviewed / intentional")
            if st.button("Reopen", key=f"reopen_{key}"):
                dismiss_set.discard(key)
                st.rerun()
            return

        c_left, c_right = st.columns([3,1])
        with c_left:
            st.markdown(f"**Why it matters for emissions:** {check['why']}")
            st.markdown(f"**Recommended action:** {check['action']}")
            st.caption(f"**In EnergyCAP:** {check['in_energycap']}")
        with c_right:
            st.badge(sev, color=SEV_COLOR.get(sev, "gray"))
            if st.button("✓ Mark as reviewed", key=f"dismiss_{key}",
                         help="Suppress — intentional or already addressed"):
                dismiss_set.add(key)
                st.rerun()

        rows_df = check.get('rows', pd.DataFrame())
        if not rows_df.empty:
            return rows_df   # caller handles display
    return None


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
_SS = {
    "drill_vendor":None,"drill_issue":None,"drill_period":None,
    "drill_assignee":None,"active_tab":"overview",
    "ft_issues":[],"ft_vendors":[],"ft_assignees":[],
    "ft_status":["Resolved","Unresolved"],"ft_priority":["High","Medium","Low"],
    "ft_energy_cat":[],"ft_sort":"Cost (high→low)",
    "cfg_dismissed": set(),
    "cont_dismissed": set(),
    "has_custom_format": False,
}
for k, v in _SS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("⚑ EnergyCAP Data Quality")
    st.caption("Bill Flags · Emissions Priority · Configuration · Continuity")
    st.divider()
    st.markdown("### Upload Files")

    with st.expander("📂 Bill Flag Files (required)", expanded=True):
        r27_files = st.file_uploader(
            "Report-27 exports OR custom flag extracts", type=["xlsx"],
            accept_multiple_files=True, key="r27_up",
            help="Accepts two formats automatically:\n"
                 "• Report-27 Bill Flags (Bills → Menu (≡) → Report-27 → Export to Excel)\n"
                 "• Custom flat flag export (one row per issue, tabular format)")
        if r27_files:
            for f in r27_files: st.success(f"✓ {f.name}", icon="📄")

    with st.expander("📂 Report-18 — Bill Line Items (optional)", expanded=True):
        r18_files = st.file_uploader(
            "One or more Report-18 exports", type=["xlsx"],
            accept_multiple_files=True, key="r18_up",
            help="Bills → Report-18, filter Line Type = Use")
        if r18_files:
            for f in r18_files: st.success(f"✓ {f.name}", icon="📄")
        else:
            st.warning("Without Report-18, emissions priority unavailable", icon="ℹ️")

    with st.expander("📂 Report-03 — Setup / Master Data (recommended)", expanded=True):
        r03_file = st.file_uploader(
            "Report-03 Setup Report", type=["xlsx"],
            accept_multiple_files=False, key="r03_up",
            help="All Reports → Setup Report for Accounts, Vendors, Meters, Sites")
        if r03_file: st.success(f"✓ {r03_file.name}", icon="📄")
        else:        st.warning("Without Report-03, config health unavailable", icon="ℹ️")

    with st.expander("📂 Report-19 — Monthly Use & Cost (optional)", expanded=True):
        r19_files = st.file_uploader(
            "One or more Report-19 exports", type=["xlsx"],
            accept_multiple_files=True, key="r19_up",
            help="Reports → Report-19 Monthly Utility Use and Cost (Excel: data only)\n"
                 "Settings: Actual data · 12 months · 2 years · Group by: Meter")
        if r19_files:
            for f in r19_files: st.success(f"✓ {f.name}", icon="📄")
        else:
            st.warning("Without Report-19, continuity audit unavailable", icon="ℹ️")

    st.divider()
    st.markdown("### Global Filters")
    df_all_ss = st.session_state.get("df_master", pd.DataFrame())
    if not df_all_ss.empty:
        _av = sorted(df_all_ss["vendor"].unique())
        status_filter   = st.multiselect("Status",["Resolved","Unresolved"],
                                         default=["Resolved","Unresolved"])
        vendor_filter   = st.multiselect("Vendor", _av, default=_av)
        priority_filter = st.multiselect("Priority",["High","Medium","Low"],
                                         default=["High","Medium","Low"])
    else:
        status_filter   = ["Resolved","Unresolved"]
        vendor_filter   = []
        priority_filter = ["High","Medium","Low"]

    st.divider()
    st.caption("Click any chart bar to drill into matching bills. Click ✕ to clear.")


# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════
if not r27_files:
    st.title("EnergyCAP Data Quality Analyzer")
    st.info(
        "👈 Upload at least one **Report-27 Bill Flags** export to get started.\n\n"
        "**Optional but recommended:**\n"
        "- **Report-18** — adds emissions priority queue\n"
        "- **Report-03** — enables configuration health checks\n"
        "- **Report-19** — enables continuity audit (missing bills, gaps)\n\n"
        "Multiple files of the same type are merged automatically."
    )
    st.stop()

r27_bytes_list = [f.read() for f in r27_files]
r18_bytes_list = [f.read() for f in r18_files] if r18_files else []
r03_bytes      = r03_file.read() if r03_file else None
r19_bytes_list = [f.read() for f in r19_files] if r19_files else []
r03_hash       = _fhash(r03_bytes) if r03_bytes else ""

with st.spinner("Parsing files… (cached — only re-runs when files change)"):
    # Report-03
    df_setup = cached_r03(r03_hash, r03_bytes) if r03_bytes else pd.DataFrame()
    config_results = cached_cfg(r03_hash, r03_bytes) if r03_bytes else {}

    # Report-27 / custom flag files (auto-detected)
    r27_results  = [cached_r27(_fhash(b), b) for b in r27_bytes_list]
    flag_formats = [fmt for _, fmt in r27_results]
    r27_frames   = [df for df, _ in r27_results if not df.empty]
    df27 = (pd.concat(r27_frames, ignore_index=True).drop_duplicates("bill_id")
            if r27_frames else pd.DataFrame())
    # Track whether any custom-format files were uploaded
    has_custom_format = 'custom' in flag_formats
    # Add flag_source to R27 rows if not already present
    if not df27.empty and 'flag_source' not in df27.columns:
        df27['flag_source'] = 'r27'
    st.session_state.has_custom_format = has_custom_format

    # Report-18
    r18_frames = [cached_r18(_fhash(b), b, r03_hash, r03_bytes) for b in r18_bytes_list]
    r18_frames = [f for f in r18_frames if not f.empty]
    df18 = (pd.concat(r18_frames, ignore_index=True).drop_duplicates("bill_id")
            if r18_frames else pd.DataFrame())

    # Merge R27 + R18
    if not df27.empty and not df18.empty:
        df_master = df27.merge(df18, on="bill_id", how="left")
    else:
        df_master = df27.copy() if not df27.empty else pd.DataFrame()

    # Report-19 + continuity checks
    cont_results = {}
    df19_full    = pd.DataFrame()
    if r19_bytes_list:
        r19_hashes_key = "|".join(_fhash(b) for b in r19_bytes_list)
        cont_results, df19_full = cached_continuity(
            r19_hashes_key, r19_bytes_list, r03_hash, r03_bytes)

st.session_state.df_master = df_master

if df_master.empty:
    st.warning("No bill records found in the uploaded Report-27 file(s).")
    st.stop()

has_usage    = ("total_kwh_equivalent" in df_master.columns
                and df_master["total_kwh_equivalent"].notna().any())
has_config   = bool(config_results)
has_cont     = bool(cont_results) and not df19_full.empty
all_vendors  = sorted(df_master["vendor"].unique())
if not vendor_filter: vendor_filter = all_vendors

# Apply global filters
df = df_master.copy()
if status_filter:   df = df[df["status"].isin(status_filter)]
if vendor_filter:   df = df[df["vendor"].isin(vendor_filter)]
if priority_filter: df = df[df["primary_priority"].isin(priority_filter)]

if df.empty:
    st.warning("No records match the current filters.")
    st.stop()

# Derived
issues_exp    = (df.explode("issues_list").rename(columns={"issues_list":"issue"})
                  .query("issue.notna() and issue != ''"))
issue_counts  = Counter(issues_exp["issue"].tolist())
vendor_counts = df["vendor"].value_counts().head(12).to_dict()
period_counts = (df[df["billing_period_label"].notna()]
                 .groupby("billing_period_label").size().sort_index().to_dict())
assignee_rows = [a.strip() for _,r in df.iterrows()
                 for a in str(r["assigned_to"]).split(",") if a.strip()]
all_assignees = sorted(set(assignee_rows))
all_issues    = sorted(issue_counts.keys())

total_bills     = len(df)
resolved_ct     = (df["status"]=="Resolved").sum()
unresolved_ct   = (df["status"]=="Unresolved").sum()
resolution_rate = resolved_ct/total_bills*100 if total_bills else 0
total_cost      = df["cost"].sum()
total_recovery  = df["cost_recovery"].sum()
avg_resolve     = df["days_to_resolve"].dropna().mean()
hp_open         = df[(df["primary_priority"]=="High")&(df["status"]=="Unresolved")].shape[0]
total_kwh       = df["total_kwh_equivalent"].sum() if has_usage else 0

# Config scorecard
cfg_high   = sum(1 for v in config_results.values() if v['severity']=='High'   and v['count']>0)
cfg_medium = sum(1 for v in config_results.values() if v['severity']=='Medium' and v['count']>0)
cfg_low    = sum(1 for v in config_results.values() if v['severity']=='Low'    and v['count']>0)

# Continuity scorecard
cont_high   = sum(1 for v in cont_results.values() if v['severity']=='High'   and v['count']>0)
cont_medium = sum(1 for v in cont_results.values() if v['severity']=='Medium' and v['count']>0)
cont_low    = sum(1 for v in cont_results.values() if v['severity']=='Low'    and v['count']>0)


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "📊 Overview",
    "⚑ Flag Analysis",
    "🏢 Vendors",
    "🌿 Emissions Priority",
    "🔧 Configuration Health",
    "📈 Continuity Audit",
    "✅ Action Guide",
    "📋 Bill Detail",
])
(t_overview, t_flags, t_vendors,
 t_emissions, t_config, t_cont,
 t_actions, t_detail) = tabs


# ── TAB 1: OVERVIEW ──────────────────────────────────────────────────────────
with t_overview:
    st.subheader("Data Quality Summary")

    st.markdown("##### Bill Flags")
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Total flagged bills", total_bills)
    c2.metric("Resolution rate",     f"{resolution_rate:.0f}%")
    c3.metric("Total bill value",    f"${total_cost:,.0f}")
    c4.metric("Cost recovered",      f"${total_recovery:,.0f}")
    c5.metric("Energy bills w/ usage",
              str(df["total_kwh_equivalent"].notna().sum()) if has_usage else "—",
              help="Upload Report-18 to enable" if not has_usage else None)
    c6.metric("High-priority open", hp_open,
              delta="need action" if hp_open else "all clear",
              delta_color="inverse" if hp_open else "normal")

    if has_config or has_cont:
        st.markdown("##### Configuration & Continuity")
        cs = st.columns(6)
        if has_config:
            cs[0].metric("Config 🔴 High",   cfg_high)
            cs[1].metric("Config 🟠 Medium", cfg_medium)
            cs[2].metric("Config 🟡 Low",    cfg_low)
        if has_cont:
            cs[3].metric("Continuity 🔴 High",   cont_high)
            cs[4].metric("Continuity 🟠 Medium",  cont_medium)
            cs[5].metric("Continuity 🟡 Low",     cont_low)

    st.divider()

    col1, col2 = st.columns([3,2])
    with col1:
        top12 = dict(sorted(issue_counts.items(), key=lambda x:-x[1])[:12])
        colors_i = [PRIORITY_COLOR.get(FLAG_META.get(k,{}).get("priority","Medium"),"#6c757d")
                    for k in top12]
        fig_iss = go.Figure(go.Bar(
            x=list(top12.values()), y=list(top12.keys()), orientation="h",
            marker_color=colors_i, text=list(top12.values()),
            textposition="outside",
            hovertemplate="%{y}: %{x}<extra></extra>"))
        fig_iss.update_layout(
            title="Flag issues by frequency — click to drill",
            yaxis_autorange="reversed",
            xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
            height=400, margin=dict(l=10,r=50,t=40,b=10),
            showlegend=False, clickmode="event+select", **PT)
        ev1 = st.plotly_chart(fig_iss, use_container_width=True,
                              on_select="rerun", key="ov_issues")
        ci = extract_click(ev1)
        if ci and ci in issue_counts:
            st.session_state.drill_issue = ci
            st.session_state.active_tab  = "overview"

    with col2:
        if period_counts:
            ev2 = st.plotly_chart(
                bar_v(period_counts, "Bills by billing period", color="#0d6efd",
                      height=400, highlight=st.session_state.drill_period),
                use_container_width=True, on_select="rerun", key="ov_period")
            cp = extract_click(ev2)
            if cp and cp in period_counts:
                st.session_state.drill_period = cp
                st.session_state.active_tab   = "overview"

    # Drill panels
    if st.session_state.drill_issue and st.session_state.active_tab == "overview":
        di = st.session_state.drill_issue
        drill_banner(f"Issue: {di}", "drill_issue")
        m = FLAG_META.get(di, {})
        if m:
            with st.expander(f"📖 About: {di}"):
                st.write(f"**Cause:** {m['cause']}")
                st.write(f"**Action:** {m['action']}")
        render_bill_cards(
            df[df["issues_list"].apply(lambda l: di in l)].sort_values("cost", ascending=False))

    if st.session_state.drill_period and st.session_state.active_tab == "overview":
        dp = st.session_state.drill_period
        drill_banner(f"Period: {dp}", "drill_period")
        render_bill_cards(
            df[df["billing_period_label"] == dp].sort_values("cost", ascending=False))

    st.divider()
    c3, c4, c5 = st.columns(3)
    with c3:
        cat_ct = (issues_exp["issue"]
                  .apply(lambda x: FLAG_META.get(x,{}).get("category","Other"))
                  .value_counts().to_dict())
        st.plotly_chart(donut(cat_ct, "Issues by category",
                              [CATEGORY_COLOR.get(k,"#888") for k in cat_ct]),
                        use_container_width=True)
    with c4:
        st.plotly_chart(donut(df["status"].value_counts().to_dict(), "Flag status",
                              ["#198754","#dc3545"]), use_container_width=True)
    with c5:
        pri_ct = (issues_exp["issue"]
                  .apply(lambda x: FLAG_META.get(x,{}).get("priority","Medium"))
                  .value_counts().reindex(["High","Medium","Low"]).dropna().to_dict())
        st.plotly_chart(donut(pri_ct, "Issues by priority",
                              ["#dc3545","#fd7e14","#198754"]), use_container_width=True)

    res_df = df[df["days_to_resolve"].notna() & (df["days_to_resolve"] >= 0)]
    if not res_df.empty:
        fig_r = px.histogram(res_df, x="days_to_resolve", nbins=20,
                             labels={"days_to_resolve":"Days to resolve"},
                             color_discrete_sequence=["#0d6efd"],
                             title=f"Resolution time — Avg {avg_resolve:.1f} d")
        fig_r.update_layout(height=240, margin=dict(l=10,r=10,t=40,b=10),
                            showlegend=False, **PT)
        st.plotly_chart(fig_r, use_container_width=True)


# ── TAB 2: FLAG ANALYSIS ─────────────────────────────────────────────────────
with t_flags:
    st.subheader("Flag Analysis")

    # Show format info badge and custom-format bonus filters
    has_pending = 'pending_days' in df.columns and df['pending_days'].notna().any()
    has_bucket  = 'bucket_days'  in df.columns
    has_resp    = 'responsibility' in df.columns
    has_notes   = 'notes' in df.columns

    if st.session_state.get('has_custom_format'):
        fmt_parts = []
        sources   = df.get('flag_source', pd.Series()).value_counts().to_dict()
        if sources.get('r27'):    fmt_parts.append(f"Report-27: {sources['r27']} bills")
        if sources.get('custom'): fmt_parts.append(f"Custom extract: {sources['custom']} bills")
        st.info(f"📋 Mixed sources — {' · '.join(fmt_parts)}. "
                f"Custom-format fields (aging, responsibility) shown where available.")

    with st.expander("🔽 Filter flags", expanded=True):
        fc1,fc2,fc3 = st.columns(3)
        with fc1:
            ft_issues = st.multiselect("Flag issue type", all_issues,
                                       default=st.session_state.ft_issues,
                                       placeholder="All issues", key="ft_issues_w")
            ft_status = st.multiselect("Status", ["Resolved","Unresolved"],
                                       default=st.session_state.ft_status, key="ft_status_w")
        with fc2:
            ft_vendors = st.multiselect("Vendor", all_vendors,
                                        default=st.session_state.ft_vendors,
                                        placeholder="All vendors", key="ft_vendors_w")
            ft_priority = st.multiselect("Priority", ["High","Medium","Low"],
                                         default=st.session_state.ft_priority, key="ft_pri_w")
        with fc3:
            ft_assignees = st.multiselect("Assignee", all_assignees,
                                          default=st.session_state.ft_assignees,
                                          placeholder="All assignees", key="ft_assign_w")
            energy_cats_avail = []
            if has_usage:
                energy_cats_avail = sorted(
                    [g for g in df["energy_category"].dropna().unique() if g])
                ft_energy = st.multiselect("Energy type", energy_cats_avail,
                                           default=st.session_state.ft_energy_cat,
                                           placeholder="All commodities", key="ft_energy_w")
            else:
                ft_energy = []
            ft_sort = st.selectbox("Sort bills by", SORT_OPTS,
                                   index=SORT_OPTS.index(st.session_state.ft_sort),
                                   key="ft_sort_w")

        # Custom-format bonus filters (only shown when relevant data present)
        if has_bucket:
            buckets_avail = sorted([b for b in df['bucket_days'].unique()
                                    if b and b not in ('','None','nan')])
            ft_bucket = st.multiselect("Aging bucket", buckets_avail,
                                       placeholder="All aging buckets", key="ft_bucket_w")
        else:
            ft_bucket = []
        if has_resp:
            resp_avail = sorted([r for r in df['responsibility'].unique()
                                 if r and r not in ('','None','nan')])
            ft_resp = st.multiselect("Responsibility", resp_avail,
                                     placeholder="All owners", key="ft_resp_w")
        else:
            ft_resp = []

        st.session_state.ft_issues     = ft_issues
        st.session_state.ft_vendors    = ft_vendors
        st.session_state.ft_assignees  = ft_assignees
        st.session_state.ft_status     = ft_status
        st.session_state.ft_priority   = ft_priority
        st.session_state.ft_energy_cat = ft_energy
        st.session_state.ft_sort       = ft_sort

        _,rb = st.columns([8,2])
        with rb:
            if st.button("↺ Reset filters", use_container_width=True):
                for k in ["ft_issues","ft_vendors","ft_assignees","ft_energy_cat"]:
                    st.session_state[k] = []
                st.session_state.ft_status   = ["Resolved","Unresolved"]
                st.session_state.ft_priority = ["High","Medium","Low"]
                st.session_state.ft_sort     = "Cost (high→low)"
                st.rerun()

    fdf = df.copy()
    if ft_issues:
        fdf = fdf[fdf["issues_list"].apply(lambda l: any(i in l for i in ft_issues))]
    if ft_vendors:   fdf = fdf[fdf["vendor"].isin(ft_vendors)]
    if ft_status:    fdf = fdf[fdf["status"].isin(ft_status)]
    if ft_priority:  fdf = fdf[fdf["primary_priority"].isin(ft_priority)]
    if ft_assignees:
        fdf = fdf[fdf["assigned_to"].apply(
            lambda a: any(x.strip() in str(a) for x in ft_assignees))]
    if ft_energy and has_usage:
        fdf = fdf[fdf["energy_category"].isin(ft_energy)]
    if ft_bucket and has_bucket:
        fdf = fdf[fdf["bucket_days"].isin(ft_bucket)]
    if ft_resp and has_resp:
        fdf = fdf[fdf["responsibility"].isin(ft_resp)]
    sc_f, sa_f = SORT_MAP[ft_sort]
    fdf = fdf.sort_values(sc_f, ascending=sa_f, na_position="last")
    fi_exp = (fdf.explode("issues_list").rename(columns={"issues_list":"issue"})
               .query("issue.notna() and issue != ''"))

    fm1,fm2,fm3,fm4 = st.columns(4)
    fm1.metric("Bills shown",  len(fdf))
    fm2.metric("Total cost",   f"${fdf['cost'].sum():,.0f}")
    fm3.metric("Unresolved",   (fdf["status"]=="Unresolved").sum())
    fm4.metric("Unique issues",fi_exp["issue"].nunique())

    ch1,ch2 = st.columns(2)
    with ch1:
        fi_ct = fi_exp["issue"].value_counts().head(12).to_dict()
        if fi_ct:
            ev_fi = st.plotly_chart(
                bar_h(fi_ct, "Issues — click to drill", color="#0d6efd", height=320,
                      highlight=st.session_state.drill_issue),
                use_container_width=True, on_select="rerun", key="ft_issues_chart")
            cfi = extract_click(ev_fi)
            if cfi and cfi in fi_ct:
                st.session_state.drill_issue = cfi
                st.session_state.active_tab  = "flags"
    with ch2:
        fv_ct = fdf["vendor"].value_counts().head(12).to_dict()
        if fv_ct:
            ev_fv = st.plotly_chart(
                bar_h(fv_ct, "Vendors — click to drill", color="#6f42c1", height=320,
                      highlight=st.session_state.drill_vendor),
                use_container_width=True, on_select="rerun", key="ft_vendors_chart")
            cfv = extract_click(ev_fv)
            if cfv and cfv in fv_ct:
                st.session_state.drill_vendor = cfv
                st.session_state.active_tab   = "flags"

    ch3,ch4 = st.columns(2)
    with ch3:
        vis_a = [a.strip() for _,r in fdf.iterrows()
                 for a in str(r["assigned_to"]).split(",") if a.strip()]
        fa_ct = dict(Counter(vis_a).most_common(10))
        if fa_ct:
            ev_fa = st.plotly_chart(
                bar_h(fa_ct, "Assignees — click to drill", color="#198754", height=300,
                      highlight=st.session_state.drill_assignee),
                use_container_width=True, on_select="rerun", key="ft_assign_chart")
            cfa = extract_click(ev_fa)
            if cfa and cfa in fa_ct:
                st.session_state.drill_assignee = cfa
                st.session_state.active_tab     = "flags"
    with ch4:
        fp_ct = (fi_exp["issue"]
                 .apply(lambda x: FLAG_META.get(x,{}).get("priority","Medium"))
                 .value_counts().reindex(["High","Medium","Low"]).dropna().to_dict())
        if fp_ct:
            st.plotly_chart(donut(fp_ct, "Priority breakdown",
                                  ["#dc3545","#fd7e14","#198754"], height=300),
                            use_container_width=True)

    # Drill panels
    if st.session_state.drill_issue and st.session_state.active_tab == "flags":
        di = st.session_state.drill_issue
        drill_banner(f"Issue: {di}", "drill_issue")
        m = FLAG_META.get(di, {})
        if m:
            with st.expander(f"📖 About: {di}"):
                st.write(f"**Cause:** {m['cause']}")
                st.write(f"**Action:** {m['action']}")
                st.write(f"**In EnergyCAP:** {m['in_energycap']}")
        render_bill_cards(fdf[fdf["issues_list"].apply(lambda l: di in l)])

    if st.session_state.drill_vendor and st.session_state.active_tab == "flags":
        dv = st.session_state.drill_vendor
        drill_banner(f"Vendor: {dv}", "drill_vendor")
        render_bill_cards(fdf[fdf["vendor"] == dv])

    if st.session_state.drill_assignee and st.session_state.active_tab == "flags":
        da = st.session_state.drill_assignee
        drill_banner(f"Assignee: {da}", "drill_assignee")
        render_bill_cards(fdf[fdf["assigned_to"].str.contains(da, na=False)])

    st.divider()
    st.subheader(f"All matching bills ({len(fdf)})")
    _dc = {"bill_id":"Bill ID","vendor":"Vendor","billing_period_label":"Period",
           "cost":"Cost ($)","status":"Status","flag_issues":"Flag Issues",
           "assigned_to":"Assigned To","days_to_resolve":"Days",
           "pending_days":"Days Open","bucket_days":"Aging",
           "responsibility":"Responsibility","assignee_org":"Org",
           "energy_category":"Energy Type","total_kwh_equivalent":"kWh Equiv."}
    _sc = [c for c in _dc if c in fdf.columns]
    disp = fdf[_sc].copy().rename(columns=_dc)
    disp["Cost ($)"] = disp["Cost ($)"].map("${:,.2f}".format)
    if "kWh Equiv." in disp.columns:
        disp["kWh Equiv."] = disp["kWh Equiv."].apply(
            lambda x: f"{x:,.0f}" if pd.notna(x) else "—")
    disp["Days"] = disp["Days"].apply(lambda x: f"{int(x)}d" if pd.notna(x) else "—")
    tbl = st.dataframe(disp, use_container_width=True, hide_index=True, height=380,
                       on_select="rerun", selection_mode="single-row", key="flags_table")
    sel = tbl.get("selection",{}).get("rows",[]) if tbl else []
    if sel:
        st.divider(); bill_detail_panel(fdf.iloc[sel[0]])
    buf = io.StringIO(); disp.to_csv(buf, index=False)
    st.download_button("⬇ Download as CSV", buf.getvalue(), "bill_flags.csv", "text/csv")


# ── TAB 3: VENDORS ───────────────────────────────────────────────────────────
with t_vendors:
    st.subheader("Vendor Analysis")
    col1,col2 = st.columns(2)
    with col1:
        ev_vc = st.plotly_chart(
            bar_h(vendor_counts, "Flags by vendor — click to drill",
                  color="#6f42c1", height=380, highlight=st.session_state.drill_vendor),
            use_container_width=True, on_select="rerun", key="vend_count")
        cv1 = extract_click(ev_vc)
        if cv1 and cv1 in vendor_counts:
            st.session_state.drill_vendor = cv1; st.session_state.active_tab = "vendors"
    with col2:
        vcost = df.groupby("vendor")["cost"].sum().sort_values(ascending=False).head(12).to_dict()
        ev_vco = st.plotly_chart(
            bar_h(vcost, "Total flagged cost — click to drill",
                  color="#dc3545", height=380, highlight=st.session_state.drill_vendor),
            use_container_width=True, on_select="rerun", key="vend_cost")
        cv2 = extract_click(ev_vco)
        if cv2 and cv2 in vcost:
            st.session_state.drill_vendor = cv2; st.session_state.active_tab = "vendors"

    if st.session_state.drill_vendor and st.session_state.active_tab == "vendors":
        dv = st.session_state.drill_vendor
        drill_banner(f"Vendor: {dv}", "drill_vendor")
        vdf   = df[df["vendor"] == dv]
        vdf_i = (vdf.explode("issues_list").rename(columns={"issues_list":"issue"})
                 .query("issue.notna() and issue != ''"))
        vc1,vc2,vc3,vc4 = st.columns(4)
        vc1.metric("Bills",         len(vdf))
        vc2.metric("Unresolved",    (vdf["status"]=="Unresolved").sum())
        vc3.metric("Total cost",    f"${vdf['cost'].sum():,.0f}")
        vc4.metric("Unique issues", vdf_i["issue"].nunique())
        v2a,v2b = st.columns(2)
        with v2a:
            vi_ct = vdf_i["issue"].value_counts().to_dict()
            if vi_ct:
                st.plotly_chart(bar_h(vi_ct, f"Issues — {dv}", color="#fd7e14",
                                      height=max(200, len(vi_ct)*36+60)),
                                use_container_width=True)
        with v2b:
            vp_ct = (vdf[vdf["billing_period_label"].notna()]
                     .groupby("billing_period_label").size().sort_index().to_dict())
            if vp_ct:
                st.plotly_chart(bar_v(vp_ct, "Bills by period",
                                      color="#0d6efd", height=300),
                                use_container_width=True)
        render_bill_cards(vdf.sort_values("cost", ascending=False))

    st.divider()
    st.subheader("All vendors — summary")
    vsumm = (df.groupby("vendor")
               .agg(bills=("bill_id","count"),
                    unresolved=("status",lambda x:(x=="Unresolved").sum()),
                    total_cost=("cost","sum"),
                    unique_issues=("issues_list",lambda x:len(set(i for l in x for i in l))))
               .sort_values("bills", ascending=False).reset_index())
    vsumm.columns = ["Vendor","Bills","Unresolved","Total Cost ($)","Unique Issue Types"]
    vsumm["Total Cost ($)"] = vsumm["Total Cost ($)"].map("${:,.0f}".format)
    vs = st.dataframe(vsumm, use_container_width=True, hide_index=True,
                      on_select="rerun", selection_mode="single-row", key="vendor_table")
    vsr = vs.get("selection",{}).get("rows",[]) if vs else []
    if vsr:
        st.session_state.drill_vendor = vsumm.iloc[vsr[0]]["Vendor"]
        st.session_state.active_tab   = "vendors"
        st.rerun()


# ── TAB 4: EMISSIONS PRIORITY ─────────────────────────────────────────────────
with t_emissions:
    st.subheader("🌿 Emissions Priority Queue")
    if not has_usage:
        st.info("Upload **Report-18** (Bill Line Item Report, Use sheet) to enable this tab.")
    else:
        cu,cs,cc = st.columns([2,2,4])
        with cu: disp_unit = st.selectbox("Display unit", DISPLAY_UNITS, index=0, key="em_unit")
        with cs:
            em_status = st.multiselect("Status", ["Resolved","Unresolved"],
                                       default=["Resolved","Unresolved"], key="em_status")
        with cc:
            st.caption("Score = Energy (kWh equiv.) × Flag risk (High=3, Med=2, Low=1) "
                       "× Status multiplier (Unresolved=3×, Resolved=1×)")

        em_df = df[df["total_kwh_equivalent"].notna()].copy()
        if em_status: em_df = em_df[em_df["status"].isin(em_status)]

        if em_df.empty:
            st.warning("No energy usage data matches the current filters.")
        else:
            fr = {"High":3,"Medium":2,"Low":1}
            em_df["flag_risk"]        = em_df["primary_priority"].map(fr).fillna(1)
            em_df["unresolved_mult"]  = (em_df["status"]=="Unresolved").astype(int)*2+1
            em_df["priority_score"]   = (em_df["total_kwh_equivalent"]
                                          * em_df["flag_risk"] * em_df["unresolved_mult"])
            em_df["display_usage"]    = em_df["total_kwh_equivalent"].apply(
                                          lambda x: from_kwh(float(x), disp_unit))
            em_df = em_df.sort_values("priority_score", ascending=False)

            st.markdown("### Usage by energy type")
            cat_sum = (em_df.groupby("energy_category")
                       .agg(bills=("bill_id","count"),
                            total_usage=("display_usage","sum"),
                            unresolved=("status",lambda x:(x=="Unresolved").sum()))
                       .sort_values("total_usage", ascending=False).reset_index())
            cat_sum.columns = ["Energy Type","Bills","Total Usage","Unresolved"]
            cat_sum["Total Usage"] = cat_sum["Total Usage"].apply(
                lambda x: f"{x:,.1f} {disp_unit}")
            scope_map = {v: GHG_SCOPE.get(v,"") for v in em_df["energy_category"].unique()}
            cat_sum.insert(1,"Scope", cat_sum["Energy Type"].map(scope_map))
            st.dataframe(cat_sum, use_container_width=True, hide_index=True)

            cat_bar = (em_df.groupby("energy_category")["display_usage"]
                       .sum().sort_values(ascending=False).to_dict())
            if cat_bar:
                sc_col = {"Electricity":"#0d6efd","Natural Gas":"#fd7e14",
                          "LPG / Propane":"#e67e22","Biomass":"#198754",
                          "District Heat / Steam":"#6f42c1","Diesel / Fuel Oil":"#dc3545",
                          "Gasoline":"#e74c3c","Coal":"#495057"}
                fig_cat = go.Figure(go.Bar(
                    x=list(cat_bar.values()), y=list(cat_bar.keys()),
                    orientation="h",
                    marker_color=[sc_col.get(k,"#888") for k in cat_bar],
                    text=[f"{v:,.0f}" for v in cat_bar.values()],
                    textposition="outside"))
                fig_cat.update_layout(
                    title=f"Flagged bill energy by type ({disp_unit})",
                    yaxis_autorange="reversed",
                    xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                    height=max(200, len(cat_bar)*40+60),
                    margin=dict(l=10,r=80,t=40,b=10), showlegend=False, **PT)
                st.plotly_chart(fig_cat, use_container_width=True)

            st.markdown("### Priority queue")
            st.caption("⚠️ Usage quantities are for prioritisation only. "
                       "Actual CO₂e requires emissions factors applied in your accounting system.")
            pq = em_df[["bill_id","vendor","billing_period_label","status",
                        "energy_category","ghg_scope","display_usage","cost",
                        "primary_priority","flag_issues","assigned_to",
                        "priority_score"]].copy()
            pq.columns = ["Bill ID","Vendor","Period","Status","Energy Type","Scope",
                          f"Usage ({disp_unit})","Cost ($)","Priority","Flag Issues",
                          "Assigned To","Score"]
            pq[f"Usage ({disp_unit})"] = pq[f"Usage ({disp_unit})"].apply(
                lambda x: f"{x:,.1f}")
            pq["Cost ($)"] = pq["Cost ($)"].map("${:,.2f}".format)
            pq["Score"]    = pq["Score"].apply(lambda x: f"{x:,.0f}")
            pq_sel = st.dataframe(pq, use_container_width=True, hide_index=True,
                                  height=480, on_select="rerun",
                                  selection_mode="single-row", key="pq_table")
            pq_rows = pq_sel.get("selection",{}).get("rows",[]) if pq_sel else []
            if pq_rows:
                st.divider(); bill_detail_panel(em_df.iloc[pq_rows[0]])
            buf2 = io.StringIO(); pq.to_csv(buf2, index=False)
            st.download_button("⬇ Download priority queue CSV",
                               buf2.getvalue(), "emissions_priority_queue.csv", "text/csv")


# ── TAB 5: CONFIGURATION HEALTH ──────────────────────────────────────────────
with t_config:
    st.subheader("🔧 Configuration Health")
    if not has_config:
        st.info("Upload **Report-03** (Setup Report) to enable configuration health checks.\n\n"
                "**How to export:** All Reports → Setup Report for Accounts, Vendors, "
                "Cost Centers, Meters, and Sites (Excel only)")
    else:
        st.warning(
            "⚠️ **Findings reflect your Report-03 export date, not the current state of "
            "EnergyCAP.** Refresh Report-03 monthly or after bulk configuration changes.",
            icon="📅")

        active_checks = {k:v for k,v in config_results.items()
                         if v['count']>0 and k not in st.session_state.cfg_dismissed}

        st.markdown("### Summary")
        s1,s2,s3,s4 = st.columns(4)
        s1.metric("Active findings", len(active_checks))
        s2.metric("🔴 High",  sum(1 for v in active_checks.values() if v['severity']=='High'))
        s3.metric("🟠 Medium",sum(1 for v in active_checks.values() if v['severity']=='Medium'))
        s4.metric("🟡 Low",   sum(1 for v in active_checks.values() if v['severity']=='Low'))

        if active_checks:
            item_ct = {k: v['count'] for k,v in
                       sorted(active_checks.items(),
                              key=lambda x:(SEVERITY_ORDER.get(x[1]['severity'],9),-x[1]['count']))}
            colors_c = [{"High":"#dc3545","Medium":"#fd7e14","Low":"#3d9be9"}
                        .get(active_checks[k]['severity'],"#888") for k in item_ct]
            fig_cfg = go.Figure(go.Bar(
                x=list(item_ct.values()),
                y=[active_checks[k]['title'] for k in item_ct],
                orientation="h", marker_color=colors_c,
                text=list(item_ct.values()), textposition="outside",
                hovertemplate="%{y}: %{x}<extra></extra>"))
            fig_cfg.update_layout(
                title="Configuration findings by check",
                yaxis_autorange="reversed",
                xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                height=max(250, len(item_ct)*42+60),
                margin=dict(l=10,r=60,t=40,b=10), showlegend=False, **PT)
            st.plotly_chart(fig_cfg, use_container_width=True)

        st.divider()
        st.markdown("### Findings detail")
        st.caption(
            "Each finding is a **candidate issue** requiring human review — not a confirmed error. "
            "Mark findings as **Reviewed / Intentional** to suppress them.")

        sorted_cfg = sorted(
            [(k,v) for k,v in config_results.items() if v['count']>0],
            key=lambda x:(SEVERITY_ORDER.get(x[1]['severity'],9),-x[1]['count']))

        for check_key, check in sorted_cfg:
            is_dismissed = check_key in st.session_state.cfg_dismissed
            sev  = check['severity']
            cnt  = check['count']
            with st.expander(
                f"{SEV_ICON.get(sev,'⚪')} **{check['title']}** — "
                f"{cnt} {'item' if cnt==1 else 'items'} "
                f"({'reviewed' if is_dismissed else sev})",
                expanded=(not is_dismissed and sev=="High")
            ):
                if is_dismissed:
                    st.caption("✅ Marked as reviewed / intentional")
                    if st.button("Reopen", key=f"cfg_reopen_{check_key}"):
                        st.session_state.cfg_dismissed.discard(check_key); st.rerun()
                    continue

                c_l,c_r = st.columns([3,1])
                with c_l:
                    st.markdown(f"**Why it matters:** {check['why']}")
                    st.markdown(f"**Action:** {check['action']}")
                    st.caption(f"**In EnergyCAP:** {check['in_energycap']}")
                with c_r:
                    st.badge(sev, color=SEV_COLOR.get(sev,"gray"))
                    if st.button("✓ Mark as reviewed", key=f"cfg_dismiss_{check_key}"):
                        st.session_state.cfg_dismissed.add(check_key); st.rerun()

                rows_df = check.get('rows', pd.DataFrame())
                if not rows_df.empty:
                    # Choose display columns per check
                    col_sets = {
                        "duplicate_serial":
                            ['serial_number','meter_code','meter_name','commodity',
                             'building_name','meter_status'],
                        "deregulated_double_count_risk":
                            ['account_name','account_number','vendor_name',
                             'vendor_role','commodity','building_name'],
                        "buildings_no_energy_meters":
                            ['building_name','building_code','building_country','primary_use'],
                        "billing_freq_typos":
                            ['meter_code','meter_name','billing_frequency',
                             'commodity','building_name'],
                        "non_usd_currency":
                            ['account_name','account_number','currency_code',
                             'vendor_name','commodity','building_name'],
                    }
                    default_cols = ['meter_code','meter_name','commodity',
                                    'building_name','building_code','meter_status']
                    show_cols = col_sets.get(check_key, default_cols)
                    show_cols = [c for c in show_cols if c in rows_df.columns]
                    tbl_df = (rows_df[show_cols].drop_duplicates('building_code')
                              if check_key == "buildings_no_energy_meters" and 'building_code' in rows_df.columns
                              else rows_df[show_cols])
                    st.markdown(f"**Affected records ({len(tbl_df)}):**")
                    st.dataframe(tbl_df, use_container_width=True, hide_index=True,
                                 height=min(400, len(tbl_df)*35+40))
                    buf_c = io.StringIO(); tbl_df.to_csv(buf_c, index=False)
                    st.download_button(f"⬇ Download list",
                                       buf_c.getvalue(),
                                       f"config_{check_key}.csv","text/csv",
                                       key=f"cfg_dl_{check_key}")

        # Portfolio composition
        if not df_setup.empty:
            st.divider()
            st.markdown("### Energy meter portfolio composition")
            ca,cb = st.columns(2)
            with ca:
                comm_ct = df_setup['commodity'].value_counts().head(15).to_dict()
                fig_cm = go.Figure(go.Bar(
                    x=list(comm_ct.values()), y=list(comm_ct.keys()),
                    orientation="h",
                    marker_color=["#0d6efd" if ENERGY_COMMODITIES.get(k) else "#adb5bd"
                                  for k in comm_ct],
                    text=list(comm_ct.values()), textposition="outside"))
                fig_cm.update_layout(
                    title="Meters by commodity (blue = energy)",
                    yaxis_autorange="reversed",
                    xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                    height=max(300, len(comm_ct)*36+60),
                    margin=dict(l=10,r=60,t=40,b=10), showlegend=False, **PT)
                st.plotly_chart(fig_cm, use_container_width=True)
            with cb:
                scope_ct = (df_setup[df_setup['is_energy']]['ghg_scope']
                            .value_counts().to_dict())
                if scope_ct:
                    st.plotly_chart(donut(scope_ct, "Energy meters by scope",
                                         ["#0d6efd","#fd7e14","#6f42c1","#198754"],
                                         height=280), use_container_width=True)
                curr_ct = {k:v for k,v in
                           df_setup['currency_code'].value_counts().to_dict().items()
                           if k not in ('','None','nan')}
                if curr_ct:
                    st.plotly_chart(donut(curr_ct,"Accounts by currency",height=250),
                                    use_container_width=True)

        # Dismissed items
        dismissed_cfg = {k:v for k,v in config_results.items()
                         if v['count']>0 and k in st.session_state.cfg_dismissed}
        if dismissed_cfg:
            st.divider()
            st.markdown("### Reviewed / Intentional findings")
            for ck, cv in dismissed_cfg.items():
                c1,c2 = st.columns([8,2])
                with c1:
                    st.caption(f"✅ {cv['title']} — {cv['count']} item(s) — reviewed")
                with c2:
                    if st.button("Reopen", key=f"cfg_reopen2_{ck}"):
                        st.session_state.cfg_dismissed.discard(ck); st.rerun()


# ── TAB 6: CONTINUITY AUDIT ──────────────────────────────────────────────────
with t_cont:
    st.subheader("📈 Continuity Audit")
    if not has_cont:
        st.info(
            "Upload **Report-19** (Monthly Utility Use and Cost — Excel: data only) "
            "to enable the continuity audit.\n\n"
            "**Recommended export settings:**\n"
            "- Data type: Actual\n"
            "- Number of months: 12 · Number of years: 2\n"
            "- End period: most recent completed month\n"
            "- Group by: Meter\n"
            "- Bill is from external vendor: From external vendor only\n"
            "- Void bills: Not void"
        )
    else:
        st.warning(
            "⚠️ **These findings reflect your Report-19 export date.** "
            "Missing months at the end of the period may simply be bills not yet received. "
            "Cross-reference billing frequency from Report-03 before flagging gaps as confirmed errors.",
            icon="📅")

        # Coverage summary
        all_p = sorted(df19_full['period'].unique(), key=period_sort_key)
        st.markdown(
            f"**Coverage:** {all_p[0].replace('-',' ')} → {all_p[-1].replace('-',' ')} "
            f"({len(all_p)} months) · **{df19_full['sheet'].nunique():,}** energy meters")

        # Energy type distribution in R19
        cat_dist = df19_full.drop_duplicates('sheet')['energy_category'].value_counts().to_dict()
        if cat_dist:
            cd1,cd2 = st.columns(2)
            with cd1:
                st.plotly_chart(donut(cat_dist,"Meters by energy type",
                                      height=260), use_container_width=True)
            with cd2:
                # Summary scorecard
                active_cont = {k:v for k,v in cont_results.items()
                               if v['count']>0 and k not in st.session_state.cont_dismissed}
                st.markdown("**Continuity findings:**")
                for k,v in sorted(active_cont.items(),
                                  key=lambda x: SEVERITY_ORDER.get(x[1]['severity'],9)):
                    icon = SEV_ICON.get(v['severity'],'⚪')
                    st.write(f"{icon} {v['title']}: **{v['count']}**")

        st.divider()
        st.markdown("### Findings detail")
        st.caption(
            "Gaps may reflect quarterly/annual billing cycles, not true data issues. "
            "Cross-reference with billing frequency (visible in Report-03) before acting. "
            "Mark findings as reviewed to suppress them.")

        sorted_cont = sorted(
            [(k,v) for k,v in cont_results.items() if v['count']>0],
            key=lambda x:(SEVERITY_ORDER.get(x[1]['severity'],9),-x[1]['count']))

        for ck, cv in sorted_cont:
            is_dismissed = ck in st.session_state.cont_dismissed
            sev = cv['severity']
            cnt = cv['count']

            with st.expander(
                f"{SEV_ICON.get(sev,'⚪')} **{cv['title']}** — "
                f"{cnt} {'meter' if cnt==1 else 'meters/instances'} "
                f"({'reviewed' if is_dismissed else sev})",
                expanded=(not is_dismissed and sev=="High")
            ):
                if is_dismissed:
                    st.caption("✅ Marked as reviewed / intentional")
                    if st.button("Reopen", key=f"cont_reopen_{ck}"):
                        st.session_state.cont_dismissed.discard(ck); st.rerun()
                    continue

                c_l,c_r = st.columns([3,1])
                with c_l:
                    st.markdown(f"**Why it matters:** {cv['why']}")
                    st.markdown(f"**Action:** {cv['action']}")
                    st.caption(f"**In EnergyCAP:** {cv['in_energycap']}")
                with c_r:
                    st.badge(sev, color=SEV_COLOR.get(sev,"gray"))
                    if st.button("✓ Mark as reviewed", key=f"cont_dismiss_{ck}"):
                        st.session_state.cont_dismissed.add(ck); st.rerun()

                rows_df = cv.get('rows', pd.DataFrame())
                if not rows_df.empty:
                    # Column sets per check type
                    gap_cols = ['meter_name','commodity','energy_category','ghg_scope',
                                'building_name','gap_after','gap_before',
                                'months_missing','billing_frequency']
                    col_sets = {
                        "large_gaps":            gap_cols,
                        "medium_gaps":           gap_cols,
                        "small_gaps":            gap_cols,
                        "zero_use_nonzero_cost": ['meter_name','commodity',
                                                   'energy_category','building_name',
                                                   'period','cost'],
                        "negative_use":          ['meter_name','commodity',
                                                   'energy_category','building_name',
                                                   'period','use','use_unit','cost'],
                        "new_meters":            ['meter_name','commodity',
                                                   'energy_category','ghg_scope',
                                                   'building_name','first_period',
                                                   'n_periods','billing_frequency'],
                        "single_year_only":      ['meter_name','commodity',
                                                   'energy_category','ghg_scope',
                                                   'building_name','last_period',
                                                   'n_periods','billing_frequency'],
                    }
                    show_cols = [c for c in col_sets.get(ck, rows_df.columns[:6])
                                 if c in rows_df.columns]
                    show_df = rows_df[show_cols].copy()

                    # Format cost/use columns
                    if 'cost' in show_df.columns:
                        show_df['cost'] = show_df['cost'].apply(
                            lambda x: f"${x:,.2f}" if pd.notna(x) else "—")
                    if 'use' in show_df.columns:
                        show_df['use'] = show_df['use'].apply(
                            lambda x: f"{x:,.2f}" if pd.notna(x) else "—")

                    # Sort meaningfully
                    if 'months_missing' in show_df.columns:
                        show_df = show_df.sort_values('months_missing', ascending=False)
                    elif 'cost' in show_df.columns:
                        pass  # already formatted as string

                    st.markdown(f"**Affected records ({len(show_df)}):**")
                    st.dataframe(show_df, use_container_width=True, hide_index=True,
                                 height=min(500, len(show_df)*35+40))

                    buf_ct = io.StringIO(); show_df.to_csv(buf_ct, index=False)
                    st.download_button(f"⬇ Download list",
                                       buf_ct.getvalue(),
                                       f"continuity_{ck}.csv", "text/csv",
                                       key=f"cont_dl_{ck}")

        # Usage trends chart (from R19)
        st.divider()
        st.markdown("### Monthly energy consumption trend")
        st.caption("Total consumption across all energy meters by period. "
                   "Significant dips may indicate missing bills.")

        trend = (df19_full[df19_full['kwh_equivalent'].notna()]
                 .groupby('period')
                 .agg(total_kwh=('kwh_equivalent','sum'),
                      meter_count=('sheet','nunique'))
                 .reset_index())
        trend['period_sort'] = trend['period'].apply(period_sort_key)
        trend = trend.sort_values('period_sort')

        if not trend.empty:
            fig_trend = go.Figure()
            fig_trend.add_trace(go.Bar(
                x=trend['period'], y=trend['total_kwh'],
                name="Total kWh equiv.",
                marker_color="#0d6efd",
                hovertemplate="%{x}: %{y:,.0f} kWh<extra></extra>"))
            fig_trend.update_layout(
                title="Monthly energy consumption — all energy meters",
                xaxis=dict(showgrid=False),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0",
                           title="kWh equivalent"),
                height=320, margin=dict(l=10,r=10,t=40,b=10),
                showlegend=False, **PT)
            st.plotly_chart(fig_trend, use_container_width=True)

            # By energy type
            trend_by_type = (df19_full[df19_full['kwh_equivalent'].notna()]
                             .groupby(['period','energy_category'])['kwh_equivalent']
                             .sum().reset_index())
            trend_by_type['period_sort'] = trend_by_type['period'].apply(period_sort_key)
            trend_by_type = trend_by_type.sort_values('period_sort')

            type_colors = {"Electricity":"#0d6efd","Natural Gas":"#fd7e14",
                           "LPG / Propane":"#e67e22","Biomass":"#198754",
                           "District Heat / Steam":"#6f42c1",
                           "Diesel / Fuel Oil":"#dc3545"}
            fig_type = go.Figure()
            for cat, grp in trend_by_type.groupby('energy_category'):
                fig_type.add_trace(go.Scatter(
                    x=grp['period'], y=grp['kwh_equivalent'],
                    name=cat, mode='lines+markers',
                    line=dict(color=type_colors.get(cat,"#888"), width=2),
                    hovertemplate=f"{cat}<br>%{{x}}: %{{y:,.0f}} kWh<extra></extra>"))
            fig_type.update_layout(
                title="Monthly consumption by energy type",
                xaxis=dict(showgrid=False),
                yaxis=dict(showgrid=True, gridcolor="#f0f0f0",
                           title="kWh equivalent"),
                height=360, margin=dict(l=10,r=10,t=40,b=10),
                legend=dict(orientation="h", y=-0.2), **PT)
            st.plotly_chart(fig_type, use_container_width=True)


# ── TAB 7: ACTION GUIDE ──────────────────────────────────────────────────────
with t_actions:
    st.subheader("Actionable Insights & Recommended Next Steps")

    open_df = df[df["status"]=="Unresolved"]
    if not open_df.empty:
        st.markdown("### 🔴 Open / Unresolved Flags")
        for _, row in open_df.iterrows():
            with st.container(border=True):
                h1,h2 = st.columns([4,1])
                with h1: st.markdown(f"**Bill {row['bill_id']} — {row['vendor']}**")
                with h2: st.markdown(f"**${row['cost']:,.2f}**")
                st.caption(f"Assigned: {row.get('assigned_to','—') or '—'}")
                issue_badges(row["issues_list"])
        st.divider()

    sorted_issues = sorted(
        [(i,c) for i,c in issue_counts.items() if i in FLAG_META],
        key=lambda x:(PRIORITY_ORDER.get(FLAG_META[x[0]]["priority"],1),-x[1]))
    seen_cats: set = set()
    for issue, cnt in sorted_issues:
        meta = FLAG_META[issue]
        cat  = meta["category"]
        if cat not in seen_cats:
            seen_cats.add(cat); st.markdown(f"### {cat}")
        ib   = issues_exp[issues_exp["issue"]==issue]
        unr  = (ib["status"]=="Unresolved").sum()
        topv = ib["vendor"].value_counts().idxmax() if not ib.empty else "—"
        p    = meta["priority"]
        with st.container(border=True):
            h1,h2 = st.columns([5,1])
            with h1:
                parts = [f"{priority_icon(p)} **{issue}**",
                         f"· {cnt} occurrence{'s' if cnt>1 else ''}",
                         f"· Top vendor: {topv}"]
                if unr: parts.append(f"· ⚠️ {unr} unresolved")
                st.markdown(" ".join(parts))
            with h2: st.badge(p, color=ISSUE_PRI_COLOR.get(p,"gray"))
            st.caption(f"*{meta['cause']}*")
            st.markdown(f"**Action:** {meta['action']}")
            st.caption(f"**In EnergyCAP:** {meta['in_energycap']}")
        matching = df[df["issues_list"].apply(lambda l: issue in l)]
        with st.expander(f"Show {len(matching)} bill(s) with this issue"):
            render_bill_cards(matching.sort_values("cost", ascending=False))

    st.divider()
    st.markdown("### 💡 Systemic Observations")
    rs = issue_counts.get("Rate schedule mismatch",0)
    sn = issue_counts.get("Serial number mismatch",0)
    if rs+sn > total_bills*0.4:
        st.info(f"**High volume of import/configuration mismatches** — "
                f"{rs} rate schedule + {sn} serial number mismatches. "
                "Consider a bulk meter configuration update rather than resolving bill-by-bill.")
    dup = (issue_counts.get("Duplicate bill",0)+issue_counts.get("Overlapping bill",0)
           +issue_counts.get("Multiple bills in period",0))
    if dup:
        st.error(f"**Potential duplicate payments — {dup} overlap/duplicate flags.** "
                 "Verify each before releasing to AP.")
    if vendor_counts:
        topvn = max(vendor_counts,key=vendor_counts.get)
        topvp = vendor_counts[topvn]/total_bills*100
        if topvp>20:
            st.warning(f"**High concentration: {topvn} ({topvp:.0f}% of flagged bills).** "
                       "Review import template and meter mappings.")
    if total_recovery==0 and total_bills>10:
        st.info("**No cost recovery tracked ($0.00).** Log Cost Recovery in EnergyCAP "
                "when billing errors are corrected to track savings ROI.")


# ── TAB 8: BILL DETAIL ───────────────────────────────────────────────────────
with t_detail:
    st.subheader("Bill Detail")
    c1,c2,c3,c4 = st.columns(4)
    with c1: search   = st.text_input("Search bill ID or account","")
    with c2: issue_f  = st.multiselect("Flag issue", sorted(issue_counts), key="det_issues")
    with c3: vendor_f = st.multiselect("Vendor", all_vendors, key="det_vendor")
    with c4: det_sort = st.selectbox("Sort by", SORT_OPTS, key="det_sort")

    ddf = df.copy()
    if search:
        ddf = ddf[ddf["bill_id"].str.contains(search, case=False) |
                  ddf["account"].str.contains(search, case=False, na=False)]
    if issue_f:
        ddf = ddf[ddf["issues_list"].apply(lambda l: any(i in l for i in issue_f))]
    if vendor_f: ddf = ddf[ddf["vendor"].isin(vendor_f)]
    dsc,dsa = SORT_MAP[det_sort]
    ddf = ddf.sort_values(dsc, ascending=dsa, na_position="last")

    _dc2 = {"bill_id":"Bill ID","vendor":"Vendor","billing_period_label":"Period",
            "cost":"Cost ($)","status":"Status","flag_issues":"Flag Issues",
            "assigned_to":"Assigned To","days_to_resolve":"Days",
            "energy_category":"Energy Type","total_kwh_equivalent":"kWh Equiv.",
            "cost_recovery":"Recovery ($)"}
    _sc2 = [c for c in _dc2 if c in ddf.columns]
    disp2 = ddf[_sc2].copy().rename(columns=_dc2)
    disp2["Cost ($)"] = disp2["Cost ($)"].map("${:,.2f}".format)
    if "Recovery ($)" in disp2.columns:
        disp2["Recovery ($)"] = disp2["Recovery ($)"].map("${:,.2f}".format)
    if "kWh Equiv." in disp2.columns:
        disp2["kWh Equiv."] = disp2["kWh Equiv."].apply(
            lambda x: f"{x:,.0f}" if pd.notna(x) else "—")
    disp2["Days"] = disp2["Days"].apply(lambda x: f"{int(x)}d" if pd.notna(x) else "—")

    tbl2 = st.dataframe(disp2, use_container_width=True, hide_index=True, height=440,
                        on_select="rerun", selection_mode="single-row", key="detail_table")
    sel2 = tbl2.get("selection",{}).get("rows",[]) if tbl2 else []
    if sel2:
        st.divider(); bill_detail_panel(ddf.iloc[sel2[0]])
    buf3 = io.StringIO(); disp2.to_csv(buf3, index=False)
    st.download_button("⬇ Download as CSV", buf3.getvalue(), "bill_detail.csv", "text/csv")

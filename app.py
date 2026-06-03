import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from collections import Counter

from parser import (
    load_uploaded_file,
    FLAG_META, PRIORITY_ORDER, PRIORITY_COLOR, CATEGORY_COLOR,
)

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="EnergyCAP Bill Flag Analyzer",
    page_icon="⚑",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .metric-card {
        background: #f8f9fa; border-radius: 10px;
        padding: 1rem 1.25rem; border: 1px solid #e9ecef;
    }
    .metric-label { font-size: 12px; color: #6c757d; margin: 0 0 4px 0; }
    .metric-value { font-size: 26px; font-weight: 600; color: #212529; margin: 0; }
    .metric-sub   { font-size: 12px; color: #6c757d; margin: 4px 0 0 0; }

    .action-card {
        border-left: 4px solid; border-radius: 6px;
        padding: 0.75rem 1rem; margin-bottom: 0.6rem; background: #fafafa;
    }
    .action-high   { border-color: #dc3545; background: #fff5f5; }
    .action-medium { border-color: #fd7e14; background: #fff8f0; }
    .action-low    { border-color: #198754; background: #f0fff4; }
    .action-info   { border-color: #0d6efd; background: #f0f6ff; }

    .bill-card {
        background: #fff; border: 1px solid #dee2e6; border-radius: 8px;
        padding: 0.85rem 1rem; margin-bottom: 0.5rem;
    }
    .bill-card:hover { border-color: #0d6efd; box-shadow: 0 0 0 2px rgba(13,110,253,.12); }
    .bill-card-header { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px; }
    .bill-id   { font-weight:600; font-size:14px; color:#212529; }
    .bill-cost { font-weight:600; font-size:15px; color:#0d6efd; }
    .bill-meta { font-size:12px; color:#6c757d; margin:2px 0; }

    .flag-badge {
        display: inline-block; font-size: 11px; padding: 2px 8px;
        border-radius: 12px; font-weight: 500; margin: 2px;
    }
    .badge-red    { background: #ffe0e0; color: #c0392b; }
    .badge-orange { background: #fff0d6; color: #9b5504; }
    .badge-green  { background: #e0f7ea; color: #1a7a4a; }
    .badge-blue   { background: #ddeeff; color: #1a4fa0; }
    .badge-gray   { background: #e9ecef; color: #495057; }
    .badge-purple { background: #f0e8ff; color: #5a1fa0; }

    .drill-banner {
        background: #e8f0fe; border: 1px solid #4285f4; border-radius: 8px;
        padding: 0.6rem 1rem; margin-bottom: 1rem;
        display: flex; align-items: center; justify-content: space-between;
    }
    .drill-label { font-size: 13px; font-weight: 500; color: #1a56db; }

    .filter-row { background: #f8f9fa; border-radius: 8px; padding: 0.75rem 1rem; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE KEYS
# ══════════════════════════════════════════════════════════════════════════════
# drill_* keys are set by clicking a chart bar and used to pre-filter the
# Bill Cards panel that appears below any chart on click.
def ss_init():
    defaults = {
        "df_all": pd.DataFrame(),
        "drill_vendor": None,       # str or None
        "drill_issue": None,        # str or None
        "drill_period": None,       # str or None
        "drill_assignee": None,     # str or None
        "active_tab": "overview",   # which tab last set a drill
        # flags tab in-tab filters
        "ft_issues": [],
        "ft_vendors": [],
        "ft_assignees": [],
        "ft_status": ["Resolved", "Unresolved"],
        "ft_priority": ["High", "Medium", "Low"],
        "ft_sort": "Cost (high→low)",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

ss_init()


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
PT = dict(font_family="Inter, system-ui, sans-serif",
          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")

SORT_OPTS = ["Cost (high→low)", "Cost (low→high)", "Days to resolve (longest first)",
             "Bill ID", "Vendor A→Z"]
SORT_MAP  = {
    "Cost (high→low)":              ("cost", False),
    "Cost (low→high)":              ("cost", True),
    "Days to resolve (longest first)": ("days_to_resolve", False),
    "Bill ID":                      ("bill_id", True),
    "Vendor A→Z":                   ("vendor", True),
}

def priority_icon(p):
    return {"High": "🔴", "Medium": "🟠", "Low": "🟢"}.get(p, "⚪")

def badge(text, color="gray"):
    return f'<span class="flag-badge badge-{color}">{text}</span>'

def metric_html(label, value, sub="", color="#212529"):
    return f"""<div class="metric-card">
        <p class="metric-label">{label}</p>
        <p class="metric-value" style="color:{color}">{value}</p>
        <p class="metric-sub">{sub}</p></div>"""

def make_bar_h(data, title, color="#0d6efd", height=320, highlight=None):
    """Horizontal bar chart. highlight = label to colour differently."""
    labels = list(data.keys())
    values = list(data.values())
    colors = []
    for lbl in labels:
        if highlight and lbl == highlight:
            colors.append("#ff6b35")
        else:
            colors.append(color)
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=colors,
        text=values, textposition="outside",
        hovertemplate="%{y}: %{x}<extra></extra>",
    ))
    fig.update_layout(
        title=title, height=height,
        yaxis=dict(autorange="reversed"),
        xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
        margin=dict(l=10, r=50, t=40, b=10),
        showlegend=False,
        clickmode="event+select",
        **PT,
    )
    return fig

def make_bar_v(data, title, color="#0d6efd", height=320, highlight=None):
    labels = list(data.keys())
    values = list(data.values())
    colors = [("#ff6b35" if highlight and lbl == highlight else color) for lbl in labels]
    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker_color=colors,
        text=values, textposition="outside",
        hovertemplate="%{x}: %{y}<extra></extra>",
    ))
    fig.update_layout(
        title=title, height=height,
        yaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
        margin=dict(l=10, r=10, t=40, b=10),
        showlegend=False,
        clickmode="event+select",
        **PT,
    )
    return fig

def make_donut(data, title, colors=None, height=280):
    colors = colors or px.colors.qualitative.Set2
    fig = go.Figure(go.Pie(
        labels=list(data.keys()), values=list(data.values()),
        hole=0.55, marker_colors=colors,
        textinfo="percent+label", textfont_size=11,
        hovertemplate="%{label}: %{value}<extra></extra>",
    ))
    fig.update_layout(title=title, height=height,
                      margin=dict(l=10, r=10, t=40, b=10),
                      showlegend=False, **PT)
    return fig

def extract_click(event_state):
    """Return the clicked label from a plotly selection event, or None."""
    if not event_state:
        return None
    pts = event_state.get("selection", {}).get("points", [])
    if not pts:
        return None
    pt = pts[0]
    # horizontal bar → label is on y axis
    return pt.get("y") or pt.get("x") or None

def render_bill_cards(sub_df, max_cards=50):
    """Render a compact card for each bill in sub_df."""
    if sub_df.empty:
        st.info("No bills match this selection.")
        return
    st.caption(f"Showing {min(len(sub_df), max_cards)} of {len(sub_df)} bills")
    for _, row in sub_df.head(max_cards).iterrows():
        p = row.get("primary_priority", "Medium")
        bc = {"High":"red","Medium":"orange","Low":"green"}.get(p, "blue")
        status_color = "#198754" if row["status"] == "Resolved" else "#dc3545"
        issue_badges = " ".join(badge(i, bc) for i in row["issues_list"])
        days = f"{int(row['days_to_resolve'])}d" if pd.notna(row.get("days_to_resolve")) else "—"
        resolver = row["resolvers"][-1] if row["resolvers"] else "—"
        st.markdown(f"""
        <div class="bill-card">
          <div class="bill-card-header">
            <span class="bill-id">Bill {row['bill_id']} &nbsp;·&nbsp; {row['vendor']}</span>
            <span class="bill-cost">${row['cost']:,.2f}</span>
          </div>
          <div class="bill-meta">
            {row.get('account','')}<br>
            Period: {row.get('billing_period_label','—')} &nbsp;·&nbsp;
            Assigned: {row.get('assigned_to','—')} &nbsp;·&nbsp;
            Resolved by: {resolver} &nbsp;·&nbsp;
            Days to resolve: {days} &nbsp;·&nbsp;
            <span style="color:{status_color};font-weight:500">{row['status']}</span>
          </div>
          <div style="margin-top:5px">{issue_badges}</div>
        </div>""", unsafe_allow_html=True)

def drill_banner(label, key):
    """Show an active drill-down banner with a clear button."""
    cols = st.columns([10, 1])
    with cols[0]:
        st.markdown(f"""<div class="drill-banner">
            <span class="drill-label">🔍 Drill-down active: <strong>{label}</strong></span>
        </div>""", unsafe_allow_html=True)
    with cols[1]:
        if st.button("✕ Clear", key=f"clear_{key}", use_container_width=True):
            st.session_state[key] = None
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.title("⚑ Bill Flag Analyzer")
    st.caption("EnergyCAP Report-27")
    st.divider()

    uploaded_files = st.file_uploader(
        "Upload Report-27 Excel file(s)", type=["xlsx"],
        accept_multiple_files=True,
        help="Bills module → Menu (≡) → Report-27 Bill Flags → Export to Excel",
    )

    st.divider()
    st.markdown("### Global Filters")
    st.caption("Apply to all tabs. Use in-tab filters on the Flags tab for finer control.")

    df_all = st.session_state.df_all
    vendor_opts = sorted(df_all["vendor"].unique()) if not df_all.empty else []

    status_filter   = st.multiselect("Flag status",
                        ["Resolved","Unresolved"], default=["Resolved","Unresolved"])
    vendor_filter   = st.multiselect("Vendor", vendor_opts, default=vendor_opts)
    priority_filter = st.multiselect("Priority",
                        ["High","Medium","Low"], default=["High","Medium","Low"])

    st.divider()
    st.caption("Click any bar in a chart to drill into the matching bills. Click ✕ Clear to reset.")


# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════
if uploaded_files:
    dfs = []
    for f in uploaded_files:
        with st.spinner(f"Parsing {f.name}…"):
            dfs.append(load_uploaded_file(f))
    if dfs:
        combined = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=["bill_id"])
        st.session_state.df_all = combined
        df_all = combined
        vendor_opts = sorted(df_all["vendor"].unique())

if df_all.empty:
    st.title("EnergyCAP Bill Flag Analyzer")
    st.info(
        "👈 Upload one or more **Report-27 Excel files** from EnergyCAP to get started.\n\n"
        "**How to export:** Bills module → Menu (≡) → Report-27 Bill Flags → Export to Excel"
    )
    st.stop()

# Apply global sidebar filters
df = df_all.copy()
if status_filter:   df = df[df["status"].isin(status_filter)]
if vendor_filter:   df = df[df["vendor"].isin(vendor_filter)]
if priority_filter: df = df[df["primary_priority"].isin(priority_filter)]

if df.empty:
    st.warning("No records match the current filters.")
    st.stop()

# Pre-compute shared derived data
issues_exp = (df.explode("issues_list")
               .rename(columns={"issues_list": "issue"})
               .query("issue.notna() and issue != ''"))

issue_counts  = Counter(issues_exp["issue"].tolist())
vendor_counts = df["vendor"].value_counts().head(12).to_dict()

period_counts = (df[df["billing_period_label"].notna()]
                 .groupby("billing_period_label").size()
                 .sort_index().to_dict())

assignee_rows = [
    a.strip()
    for _, row in df.iterrows()
    for a in str(row["assigned_to"]).split(",") if a.strip()
]
resolver_rows = [r for _, row in df.iterrows() for r in row["resolvers"]]

all_assignees = sorted(set(assignee_rows))
all_issues    = sorted(issue_counts.keys())
all_vendors   = sorted(df["vendor"].unique())

total_bills      = len(df)
resolved_count   = (df["status"] == "Resolved").sum()
unresolved_count = (df["status"] == "Unresolved").sum()
resolution_rate  = resolved_count / total_bills * 100 if total_bills else 0
total_cost       = df["cost"].sum()
total_recovery   = df["cost_recovery"].sum()
multi_issue_cnt  = (df["num_issues"] >= 3).sum()
avg_resolve      = df["days_to_resolve"].dropna().mean()
hp_open          = df[(df["primary_priority"]=="High") & (df["status"]=="Unresolved")].shape[0]


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
t_overview, t_flags, t_vendors, t_actions, t_detail = st.tabs([
    "📊 Overview", "⚑ Flag Analysis", "🏢 Vendors", "✅ Action Guide", "📋 Bill Detail",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with t_overview:
    st.subheader("Summary")

    cols = st.columns(6)
    metrics = [
        ("Total flagged bills",    str(total_bills),
         f"{resolved_count} resolved · {unresolved_count} open", "#212529"),
        ("Resolution rate",        f"{resolution_rate:.0f}%",
         f"{resolved_count} of {total_bills}",
         "#198754" if resolution_rate >= 90 else "#fd7e14"),
        ("Total bill value",       f"${total_cost:,.0f}", "Under review", "#212529"),
        ("Cost recovered",         f"${total_recovery:,.0f}", "Tracked savings", "#198754"),
        ("Multi-issue bills (3+)", str(multi_issue_cnt),
         f"{multi_issue_cnt/total_bills*100:.0f}% of total", "#6f42c1"),
        ("High-priority open",     str(hp_open), "Need immediate action",
         "#dc3545" if hp_open > 0 else "#198754"),
    ]
    for col, (label, val, sub, color) in zip(cols, metrics):
        col.markdown(metric_html(label, val, sub, color), unsafe_allow_html=True)

    st.markdown("")

    # ── Flag issues bar (clickable) ───────────────────────────────────────────
    col1, col2 = st.columns([3, 2])
    with col1:
        top12 = dict(sorted(issue_counts.items(), key=lambda x: -x[1])[:12])
        colors_i = [PRIORITY_COLOR.get(FLAG_META.get(k,{}).get("priority","Medium"),"#6c757d")
                    for k in top12]
        fig_issues = go.Figure(go.Bar(
            x=list(top12.values()), y=list(top12.keys()),
            orientation="h", marker_color=colors_i,
            text=list(top12.values()), textposition="outside",
            hovertemplate="%{y}: %{x}<extra></extra>",
        ))
        fig_issues.update_layout(
            title="Flag issues by frequency — click a bar to drill down",
            yaxis_autorange="reversed",
            xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
            height=400, margin=dict(l=10,r=50,t=40,b=10),
            showlegend=False, clickmode="event+select", **PT,
        )
        ev1 = st.plotly_chart(fig_issues, use_container_width=True,
                              on_select="rerun", key="ov_issues_chart")
        clicked_issue = extract_click(ev1)
        if clicked_issue and clicked_issue in issue_counts:
            st.session_state.drill_issue = clicked_issue
            st.session_state.active_tab = "overview"

    with col2:
        if period_counts:
            fig_period = make_bar_v(period_counts, "Bills by billing period — click to drill",
                                    color="#0d6efd", height=400)
            ev_period = st.plotly_chart(fig_period, use_container_width=True,
                                        on_select="rerun", key="ov_period_chart")
            clicked_period = extract_click(ev_period)
            if clicked_period and clicked_period in period_counts:
                st.session_state.drill_period = clicked_period
                st.session_state.active_tab = "overview"

    # ── Drill-down panel for Overview issue click ─────────────────────────────
    if st.session_state.drill_issue and st.session_state.active_tab == "overview":
        diss = st.session_state.drill_issue
        drill_banner(f"Issue: {diss}", "drill_issue")
        sub = df[df["issues_list"].apply(lambda lst: diss in lst)]
        sc, sa = SORT_MAP["Cost (high→low)"]
        render_bill_cards(sub.sort_values(sc, ascending=sa))

    if st.session_state.drill_period and st.session_state.active_tab == "overview":
        dp = st.session_state.drill_period
        drill_banner(f"Period: {dp}", "drill_period")
        sub = df[df["billing_period_label"] == dp]
        sc, sa = SORT_MAP["Cost (high→low)"]
        render_bill_cards(sub.sort_values(sc, ascending=sa))

    # ── Category / status / priority donuts ──────────────────────────────────
    c3, c4, c5 = st.columns(3)
    with c3:
        cat_ct = (issues_exp["issue"]
                  .apply(lambda x: FLAG_META.get(x,{}).get("category","Other"))
                  .value_counts().to_dict())
        st.plotly_chart(make_donut(cat_ct, "Issues by category",
                                   [CATEGORY_COLOR.get(k,"#888") for k in cat_ct]),
                        use_container_width=True)
    with c4:
        st.plotly_chart(make_donut(df["status"].value_counts().to_dict(),
                                   "Flag status", ["#198754","#dc3545"]),
                        use_container_width=True)
    with c5:
        pri_ct = (issues_exp["issue"]
                  .apply(lambda x: FLAG_META.get(x,{}).get("priority","Medium"))
                  .value_counts().reindex(["High","Medium","Low"]).dropna().to_dict())
        st.plotly_chart(make_donut(pri_ct, "Issues by priority",
                                   ["#dc3545","#fd7e14","#198754"]),
                        use_container_width=True)

    resolve_df = df[df["days_to_resolve"].notna() & (df["days_to_resolve"] >= 0)]
    if not resolve_df.empty:
        st.subheader("Resolution time distribution")
        fig_res = px.histogram(resolve_df, x="days_to_resolve", nbins=20,
                               labels={"days_to_resolve":"Days to resolve"},
                               color_discrete_sequence=["#0d6efd"],
                               title=f"Avg {avg_resolve:.1f} d · Median {resolve_df['days_to_resolve'].median():.0f} d")
        fig_res.update_layout(height=240, margin=dict(l=10,r=10,t=40,b=10),
                              showlegend=False, **PT)
        st.plotly_chart(fig_res, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — FLAG ANALYSIS  (with rich multi-filter panel + click-through)
# ══════════════════════════════════════════════════════════════════════════════
with t_flags:
    st.subheader("Flag Analysis")

    # ── Multi-filter panel ────────────────────────────────────────────────────
    with st.expander("🔽 Filter flags", expanded=True):
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            ft_issues = st.multiselect(
                "Flag issue type", all_issues,
                default=st.session_state.ft_issues,
                placeholder="All issues",
                key="ft_issues_widget",
            )
            ft_status = st.multiselect(
                "Status", ["Resolved","Unresolved"],
                default=st.session_state.ft_status,
                key="ft_status_widget",
            )
        with f_col2:
            ft_vendors = st.multiselect(
                "Vendor", all_vendors,
                default=st.session_state.ft_vendors,
                placeholder="All vendors",
                key="ft_vendors_widget",
            )
            ft_priority = st.multiselect(
                "Priority", ["High","Medium","Low"],
                default=st.session_state.ft_priority,
                key="ft_priority_widget",
            )
        with f_col3:
            ft_assignees = st.multiselect(
                "Assignee", all_assignees,
                default=st.session_state.ft_assignees,
                placeholder="All assignees",
                key="ft_assignees_widget",
            )
            ft_sort = st.selectbox(
                "Sort bills by", SORT_OPTS,
                index=SORT_OPTS.index(st.session_state.ft_sort),
                key="ft_sort_widget",
            )
        # persist to session state
        st.session_state.ft_issues    = ft_issues
        st.session_state.ft_vendors   = ft_vendors
        st.session_state.ft_assignees = ft_assignees
        st.session_state.ft_status    = ft_status
        st.session_state.ft_priority  = ft_priority
        st.session_state.ft_sort      = ft_sort

        rc1, rc2 = st.columns([8,2])
        with rc2:
            if st.button("↺ Reset all filters", use_container_width=True):
                for k in ["ft_issues","ft_vendors","ft_assignees"]:
                    st.session_state[k] = []
                st.session_state.ft_status   = ["Resolved","Unresolved"]
                st.session_state.ft_priority = ["High","Medium","Low"]
                st.session_state.ft_sort     = "Cost (high→low)"
                st.rerun()

    # ── Apply in-tab filters ──────────────────────────────────────────────────
    fdf = df.copy()
    if ft_issues:
        fdf = fdf[fdf["issues_list"].apply(lambda lst: any(i in lst for i in ft_issues))]
    if ft_vendors:
        fdf = fdf[fdf["vendor"].isin(ft_vendors)]
    if ft_assignees:
        fdf = fdf[fdf["assigned_to"].apply(
            lambda a: any(x.strip() in str(a) for x in ft_assignees)
        )]
    if ft_status:
        fdf = fdf[fdf["status"].isin(ft_status)]
    if ft_priority:
        fdf = fdf[fdf["primary_priority"].isin(ft_priority)]

    sc_f, sa_f = SORT_MAP[ft_sort]
    fdf = fdf.sort_values(sc_f, ascending=sa_f, na_position="last")

    fissues_exp = (fdf.explode("issues_list")
                   .rename(columns={"issues_list":"issue"})
                   .query("issue.notna() and issue != ''"))

    # ── Summary metrics for current filter ───────────────────────────────────
    fm1, fm2, fm3, fm4 = st.columns(4)
    fm1.metric("Bills shown",  len(fdf))
    fm2.metric("Total cost",   f"${fdf['cost'].sum():,.0f}")
    fm3.metric("Unresolved",   (fdf["status"]=="Unresolved").sum())
    fm4.metric("Unique issues",fissues_exp["issue"].nunique())

    # ── Charts (clickable) ────────────────────────────────────────────────────
    ch1, ch2 = st.columns(2)

    with ch1:
        # Issues chart — click sets drill_issue
        fi_ct = fissues_exp["issue"].value_counts().head(12).to_dict()
        if fi_ct:
            fig_fi = make_bar_h(
                fi_ct, "Issues — click to drill",
                color="#0d6efd", height=320,
                highlight=st.session_state.drill_issue,
            )
            ev_fi = st.plotly_chart(fig_fi, use_container_width=True,
                                    on_select="rerun", key="ft_issues_chart")
            clicked_fi = extract_click(ev_fi)
            if clicked_fi and clicked_fi in fi_ct:
                st.session_state.drill_issue = clicked_fi
                st.session_state.active_tab = "flags"

    with ch2:
        # Vendors chart — click sets drill_vendor
        fv_ct = fdf["vendor"].value_counts().head(12).to_dict()
        if fv_ct:
            fig_fv = make_bar_h(
                fv_ct, "Vendors — click to drill",
                color="#6f42c1", height=320,
                highlight=st.session_state.drill_vendor,
            )
            ev_fv = st.plotly_chart(fig_fv, use_container_width=True,
                                    on_select="rerun", key="ft_vendors_chart")
            clicked_fv = extract_click(ev_fv)
            if clicked_fv and clicked_fv in fv_ct:
                st.session_state.drill_vendor = clicked_fv
                st.session_state.active_tab = "flags"

    ch3, ch4 = st.columns(2)

    with ch3:
        # Assignees chart — click sets drill_assignee
        fa_ct = Counter(assignee_rows)
        # filter to visible bills
        vis_assignees: list[str] = []
        for _, row in fdf.iterrows():
            for a in str(row["assigned_to"]).split(","):
                a = a.strip()
                if a:
                    vis_assignees.append(a)
        fa_ct_vis = dict(Counter(vis_assignees).most_common(10))
        if fa_ct_vis:
            fig_fa = make_bar_h(
                fa_ct_vis, "Assignees — click to drill",
                color="#198754", height=300,
                highlight=st.session_state.drill_assignee,
            )
            ev_fa = st.plotly_chart(fig_fa, use_container_width=True,
                                    on_select="rerun", key="ft_assignees_chart")
            clicked_fa = extract_click(ev_fa)
            if clicked_fa and clicked_fa in fa_ct_vis:
                st.session_state.drill_assignee = clicked_fa
                st.session_state.active_tab = "flags"

    with ch4:
        # Priority breakdown
        fp_ct = (fissues_exp["issue"]
                 .apply(lambda x: FLAG_META.get(x,{}).get("priority","Medium"))
                 .value_counts().reindex(["High","Medium","Low"]).dropna().to_dict())
        if fp_ct:
            st.plotly_chart(make_donut(fp_ct, "Priority breakdown",
                                       ["#dc3545","#fd7e14","#198754"], height=300),
                            use_container_width=True)

    # ── Drill-down bill cards (shown below charts) ────────────────────────────
    drill_applied = False

    if st.session_state.drill_issue and st.session_state.active_tab == "flags":
        diss = st.session_state.drill_issue
        drill_banner(f"Issue: {diss}", "drill_issue")
        sub = fdf[fdf["issues_list"].apply(lambda lst: diss in lst)]
        meta = FLAG_META.get(diss, {})
        if meta:
            with st.expander(f"📖 About this flag: {diss}", expanded=False):
                st.markdown(f"**Cause:** {meta.get('cause','')}")
                st.markdown(f"**Action:** {meta.get('action','')}")
                st.markdown(f"**In EnergyCAP:** {meta.get('in_energycap','')}")
        render_bill_cards(sub)
        drill_applied = True

    if st.session_state.drill_vendor and st.session_state.active_tab == "flags":
        dv = st.session_state.drill_vendor
        drill_banner(f"Vendor: {dv}", "drill_vendor")
        sub = fdf[fdf["vendor"] == dv]
        render_bill_cards(sub)
        drill_applied = True

    if st.session_state.drill_assignee and st.session_state.active_tab == "flags":
        da = st.session_state.drill_assignee
        drill_banner(f"Assignee: {da}", "drill_assignee")
        sub = fdf[fdf["assigned_to"].str.contains(da, na=False)]
        render_bill_cards(sub)
        drill_applied = True

    # ── Full bill list (always shown, respects all filters) ───────────────────
    st.divider()
    st.subheader(f"All matching bills ({len(fdf)})")

    disp = fdf[["bill_id","vendor","billing_period_label","cost","status",
                "flag_issues","assigned_to","days_to_resolve","cost_recovery"]].copy()
    disp.columns = ["Bill ID","Vendor","Period","Cost ($)","Status",
                    "Flag Issues","Assigned To","Days to Resolve","Cost Recovery ($)"]
    disp["Cost ($)"]          = disp["Cost ($)"].map("${:,.2f}".format)
    disp["Cost Recovery ($)"] = disp["Cost Recovery ($)"].map("${:,.2f}".format)
    disp["Days to Resolve"]   = disp["Days to Resolve"].apply(
        lambda x: f"{int(x)}d" if pd.notna(x) else "—")

    sel = st.dataframe(disp, use_container_width=True, hide_index=True,
                       height=400, on_select="rerun",
                       selection_mode="single-row", key="flags_table")

    # Click in table → show bill detail
    selected_rows = sel.get("selection", {}).get("rows", []) if sel else []
    if selected_rows:
        row = fdf.iloc[selected_rows[0]]
        st.markdown("---")
        st.subheader(f"Bill {row['bill_id']} — {row['vendor']}")
        dc1, dc2 = st.columns(2)
        with dc1:
            st.markdown(f"**Account:** {row['account']}")
            st.markdown(f"**Period:** {row['billing_period_label']}")
            st.markdown(f"**Cost:** ${row['cost']:,.2f}  |  **Recovery:** ${row['cost_recovery']:,.2f}")
            st.markdown(f"**Assigned to:** {row['assigned_to'] or '—'}")
        with dc2:
            st.markdown(f"**Status:** {row['status']}")
            days = f"{int(row['days_to_resolve'])}d" if pd.notna(row['days_to_resolve']) else "Pending"
            st.markdown(f"**Days to resolve:** {days}")
            if row["resolvers"]:
                st.markdown(f"**Resolved by:** {', '.join(row['resolvers'])}")
        st.markdown("**Issues and guidance:**")
        for issue in row["issues_list"]:
            m = FLAG_META.get(issue, {})
            icon = priority_icon(m.get("priority",""))
            with st.expander(f"{icon} {issue} — {m.get('category','')}"):
                if m:
                    st.markdown(f"**Cause:** {m['cause']}")
                    st.markdown(f"**Action:** {m['action']}")
                    st.markdown(f"**In EnergyCAP:** {m['in_energycap']}")

    buf = io.StringIO()
    disp.to_csv(buf, index=False)
    st.download_button("⬇ Download filtered list as CSV", buf.getvalue(),
                       "bill_flags_filtered.csv", "text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — VENDORS  (fully clickable)
# ══════════════════════════════════════════════════════════════════════════════
with t_vendors:
    st.subheader("Vendor Analysis")

    col1, col2 = st.columns(2)

    with col1:
        fig_vc = make_bar_h(vendor_counts, "Flags by vendor — click to drill",
                            color="#6f42c1", height=380,
                            highlight=st.session_state.drill_vendor)
        ev_vc = st.plotly_chart(fig_vc, use_container_width=True,
                                on_select="rerun", key="vend_count_chart")
        cv1 = extract_click(ev_vc)
        if cv1 and cv1 in vendor_counts:
            st.session_state.drill_vendor = cv1
            st.session_state.active_tab = "vendors"

    with col2:
        vcost = df.groupby("vendor")["cost"].sum().sort_values(ascending=False).head(12)
        vcost_d = vcost.to_dict()
        fig_vco = make_bar_h(vcost_d, "Total flagged cost by vendor — click to drill",
                             color="#dc3545", height=380,
                             highlight=st.session_state.drill_vendor)
        ev_vco = st.plotly_chart(fig_vco, use_container_width=True,
                                 on_select="rerun", key="vend_cost_chart")
        cv2 = extract_click(ev_vco)
        if cv2 and cv2 in vcost_d:
            st.session_state.drill_vendor = cv2
            st.session_state.active_tab = "vendors"

    # ── Vendor drill-down panel ───────────────────────────────────────────────
    if st.session_state.drill_vendor and st.session_state.active_tab == "vendors":
        dv = st.session_state.drill_vendor
        drill_banner(f"Vendor: {dv}", "drill_vendor")

        vdf = df[df["vendor"] == dv]
        vdf_i = (vdf.explode("issues_list").rename(columns={"issues_list":"issue"})
                 .query("issue.notna() and issue != ''"))

        vc1, vc2, vc3, vc4 = st.columns(4)
        vc1.metric("Bills flagged",  len(vdf))
        vc2.metric("Unresolved",     (vdf["status"]=="Unresolved").sum())
        vc3.metric("Total cost",     f"${vdf['cost'].sum():,.0f}")
        vc4.metric("Unique issues",  vdf_i["issue"].nunique())

        v2a, v2b = st.columns(2)
        with v2a:
            vi_ct = vdf_i["issue"].value_counts().to_dict()
            if vi_ct:
                fig_vi = make_bar_h(vi_ct, f"Issues for {dv}",
                                    color="#fd7e14",
                                    height=max(200, len(vi_ct)*36+60))
                ev_vi = st.plotly_chart(fig_vi, use_container_width=True,
                                        on_select="rerun", key="vd_issue_chart")
                cv3 = extract_click(ev_vi)
                if cv3:
                    st.session_state.drill_issue  = cv3
                    st.session_state.drill_vendor = dv
                    # show below

        with v2b:
            # period breakdown for this vendor
            vp_ct = (vdf[vdf["billing_period_label"].notna()]
                     .groupby("billing_period_label").size()
                     .sort_index().to_dict())
            if vp_ct:
                st.plotly_chart(make_bar_v(vp_ct, f"Bills by period — {dv}",
                                           color="#0d6efd", height=300),
                                use_container_width=True)

        render_bill_cards(vdf.sort_values("cost", ascending=False))

    # ── Vendor summary table ──────────────────────────────────────────────────
    st.divider()
    st.subheader("All vendors — summary")
    vsumm = (df.groupby("vendor")
               .agg(bills=("bill_id","count"),
                    unresolved=("status", lambda x: (x=="Unresolved").sum()),
                    total_cost=("cost","sum"),
                    unique_issues=("issues_list", lambda x: len(set(i for lst in x for i in lst))))
               .sort_values("bills", ascending=False).reset_index())
    vsumm.columns = ["Vendor","Bills","Unresolved","Total Cost ($)","Unique Issue Types"]
    vsumm["Total Cost ($)"] = vsumm["Total Cost ($)"].map("${:,.0f}".format)

    vs_sel = st.dataframe(vsumm, use_container_width=True, hide_index=True,
                          on_select="rerun", selection_mode="single-row",
                          key="vendor_table")
    vs_rows = vs_sel.get("selection",{}).get("rows",[]) if vs_sel else []
    if vs_rows:
        picked_vendor = vsumm.iloc[vs_rows[0]]["Vendor"]
        st.session_state.drill_vendor = picked_vendor
        st.session_state.active_tab = "vendors"
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — ACTION GUIDE
# ══════════════════════════════════════════════════════════════════════════════
with t_actions:
    st.subheader("Actionable Insights & Recommended Next Steps")
    st.caption("Sorted by priority (High → Medium → Low) then occurrence count. "
               "Click any issue card's 'Show bills' button to inspect matching bills.")

    open_df = df[df["status"] == "Unresolved"]
    if not open_df.empty:
        st.markdown("### 🔴 Open / Unresolved Flags")
        for _, row in open_df.iterrows():
            p   = row["primary_priority"]
            css = {"High":"action-high","Medium":"action-medium","Low":"action-low"}.get(p,"action-info")
            bc  = {"High":"red","Medium":"orange","Low":"green"}.get(p,"blue")
            badges = " ".join(badge(i, bc) for i in row["issues_list"])
            st.markdown(f"""<div class="action-card {css}">
                <strong>Bill {row['bill_id']} — {row['vendor']}</strong>
                &nbsp;|&nbsp; ${row['cost']:,.2f}
                &nbsp;|&nbsp; Assigned: {row['assigned_to'] or 'Unassigned'}<br>{badges}
            </div>""", unsafe_allow_html=True)
        st.divider()

    sorted_issues = sorted(
        [(i, c) for i, c in issue_counts.items() if i in FLAG_META],
        key=lambda x: (PRIORITY_ORDER.get(FLAG_META[x[0]]["priority"],1), -x[1])
    )
    seen_cats: set = set()
    for issue, cnt in sorted_issues:
        meta  = FLAG_META[issue]
        cat   = meta["category"]
        if cat not in seen_cats:
            seen_cats.add(cat)
            st.markdown(f"### {cat}")
        p    = meta["priority"]
        css  = {"High":"action-high","Medium":"action-medium","Low":"action-low"}.get(p,"action-info")
        bc   = {"High":"red","Medium":"orange","Low":"green"}.get(p,"blue")
        ib   = issues_exp[issues_exp["issue"]==issue]
        unr  = (ib["status"]=="Unresolved").sum()
        icost= ib["cost"].sum()
        topv = ib["vendor"].value_counts().idxmax() if not ib.empty else "—"
        unr_html = f'&nbsp;{badge("⚠ "+str(unr)+" unresolved","red")}' if unr else ""

        with st.container():
            st.markdown(f"""<div class="action-card {css}">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;">
                    <div>
                        <strong>{issue}</strong>
                        &nbsp;{badge(p, bc)}
                        &nbsp;{badge(str(cnt)+" occurrence"+("s" if cnt>1 else ""), "gray")}
                        {unr_html}
                    </div>
                    <div style="font-size:12px;color:#6c757d;text-align:right;">
                        ${icost:,.0f} total · Top vendor: {topv}
                    </div>
                </div>
                <p style="margin:8px 0 4px;font-size:13px;color:#555;"><em>{meta['cause']}</em></p>
                <p style="margin:4px 0 2px;font-size:13px;"><strong>Action:</strong> {meta['action']}</p>
                <p style="margin:2px 0 0;font-size:12px;color:#6c757d;"><strong>In EnergyCAP:</strong> {meta['in_energycap']}</p>
            </div>""", unsafe_allow_html=True)

            # Expandable bill list for each issue
            matching = df[df["issues_list"].apply(lambda lst: issue in lst)]
            with st.expander(f"Show {len(matching)} bill{'s' if len(matching)!=1 else ''} with this issue"):
                render_bill_cards(matching.sort_values("cost", ascending=False))

    st.divider()
    st.markdown("### 💡 Systemic Observations")
    insights = []

    rs = issue_counts.get("Rate schedule mismatch",0)
    sn = issue_counts.get("Serial number mismatch",0)
    if rs + sn > total_bills * 0.4:
        insights.append(("action-info",
            "High volume of import/configuration mismatches",
            f"{rs} rate schedule + {sn} serial number mismatches across {total_bills} bills "
            f"({(rs+sn)/total_bills*100:.0f}% of occurrences). "
            "Consider a bulk meter configuration update in EnergyCAP rather than resolving bill-by-bill."))

    dup = (issue_counts.get("Duplicate bill",0)+issue_counts.get("Overlapping bill",0)
           +issue_counts.get("Multiple bills in period",0))
    if dup > 0:
        insights.append(("action-high",
            f"Potential duplicate payments — {dup} overlap/duplicate flags",
            "These carry the highest financial risk. Verify each before releasing to AP. "
            "Cross-reference with your payment system to confirm no duplicates were paid."))

    if vendor_counts:
        topv_n = max(vendor_counts, key=vendor_counts.get)
        topv_p = vendor_counts[topv_n] / total_bills * 100
        if topv_p > 20:
            insights.append(("action-medium",
                f"High concentration: {topv_n} ({topv_p:.0f}% of flagged bills)",
                f"{topv_n} is disproportionately represented. Review import template and meter mappings."))

    if total_recovery == 0 and total_bills > 10:
        insights.append(("action-info",
            "No cost recovery tracked ($0.00)",
            "Start logging Cost Recovery in EnergyCAP when billing errors are corrected. "
            "This quantifies the ROI of your flag review process."))

    for css, title, body in insights:
        st.markdown(f"""<div class="action-card {css}">
            <strong>{title}</strong>
            <p style="margin:6px 0 0;font-size:13px;">{body}</p>
        </div>""", unsafe_allow_html=True)
    if not insights:
        st.success("No major systemic issues detected with current filters.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — BILL DETAIL
# ══════════════════════════════════════════════════════════════════════════════
with t_detail:
    st.subheader("Bill Detail")
    c1, c2, c3, c4 = st.columns(4)
    with c1: search = st.text_input("Search bill ID or account", "")
    with c2: issue_f  = st.multiselect("Flag issue", sorted(issue_counts), key="det_issues")
    with c3: vendor_f = st.multiselect("Vendor", all_vendors, key="det_vendor")
    with c4: det_sort = st.selectbox("Sort by", SORT_OPTS, key="det_sort")

    ddf = df.copy()
    if search:
        ddf = ddf[ddf["bill_id"].str.contains(search, case=False) |
                  ddf["account"].str.contains(search, case=False, na=False)]
    if issue_f:
        ddf = ddf[ddf["issues_list"].apply(lambda lst: any(i in lst for i in issue_f))]
    if vendor_f:
        ddf = ddf[ddf["vendor"].isin(vendor_f)]
    dsc, dsa = SORT_MAP[det_sort]
    ddf = ddf.sort_values(dsc, ascending=dsa, na_position="last")

    disp2 = ddf[["bill_id","vendor","billing_period_label","cost","status",
                 "flag_issues","assigned_to","days_to_resolve","cost_recovery"]].copy()
    disp2.columns = ["Bill ID","Vendor","Period","Cost ($)","Status",
                     "Flag Issues","Assigned To","Days to Resolve","Cost Recovery ($)"]
    disp2["Cost ($)"]          = disp2["Cost ($)"].map("${:,.2f}".format)
    disp2["Cost Recovery ($)"] = disp2["Cost Recovery ($)"].map("${:,.2f}".format)
    disp2["Days to Resolve"]   = disp2["Days to Resolve"].apply(
        lambda x: f"{int(x)}d" if pd.notna(x) else "—")

    tbl_sel = st.dataframe(disp2, use_container_width=True, hide_index=True,
                           height=440, on_select="rerun",
                           selection_mode="single-row", key="detail_table")

    sel_rows = tbl_sel.get("selection",{}).get("rows",[]) if tbl_sel else []
    if sel_rows:
        row = ddf.iloc[sel_rows[0]]
        st.markdown("---")
        st.subheader(f"Bill {row['bill_id']} — {row['vendor']}")
        dc1, dc2 = st.columns(2)
        with dc1:
            st.markdown(f"**Account:** {row['account']}")
            st.markdown(f"**Period:** {row['billing_period_label']}")
            st.markdown(f"**Cost:** ${row['cost']:,.2f}  |  **Recovery:** ${row['cost_recovery']:,.2f}")
            st.markdown(f"**Assigned to:** {row['assigned_to'] or '—'}")
        with dc2:
            st.markdown(f"**Status:** {row['status']}")
            days = f"{int(row['days_to_resolve'])}d" if pd.notna(row['days_to_resolve']) else "Pending"
            st.markdown(f"**Days to resolve:** {days}")
            if row["resolvers"]:
                st.markdown(f"**Resolved by:** {', '.join(row['resolvers'])}")
        st.markdown("**Issues and guidance:**")
        for issue in row["issues_list"]:
            m = FLAG_META.get(issue, {})
            icon = priority_icon(m.get("priority",""))
            with st.expander(f"{icon} {issue} — {m.get('category','')}"):
                if m:
                    st.markdown(f"**Cause:** {m['cause']}")
                    st.markdown(f"**Action:** {m['action']}")
                    st.markdown(f"**In EnergyCAP:** {m['in_energycap']}")

    buf = io.StringIO()
    disp2.to_csv(buf, index=False)
    st.download_button("⬇ Download filtered list as CSV", buf.getvalue(),
                       "bill_flags_detail.csv", "text/csv")

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
from collections import Counter

from parser import (
    load_uploaded_file, parse_report27_text,
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
    .action-card  {
        border-left: 4px solid; border-radius: 6px;
        padding: 0.75rem 1rem; margin-bottom: 0.6rem; background: #fafafa;
    }
    .action-high   { border-color: #dc3545; background: #fff5f5; }
    .action-medium { border-color: #fd7e14; background: #fff8f0; }
    .action-low    { border-color: #198754; background: #f0fff4; }
    .action-info   { border-color: #0d6efd; background: #f0f6ff; }
    .flag-badge {
        display: inline-block; font-size: 11px; padding: 2px 8px;
        border-radius: 12px; font-weight: 500; margin: 2px;
    }
    .badge-red    { background: #ffe0e0; color: #c0392b; }
    .badge-orange { background: #fff0d6; color: #9b5504; }
    .badge-green  { background: #e0f7ea; color: #1a7a4a; }
    .badge-blue   { background: #ddeeff; color: #1a4fa0; }
    .badge-gray   { background: #e9ecef; color: #495057; }
</style>
""", unsafe_allow_html=True)

# ── Helpers ────────────────────────────────────────────────────────────────────
PT = dict(font_family="Inter, system-ui, sans-serif",
          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")

def bar_h(data, title, color="#0d6efd", height=320):
    labels, values = list(data.keys()), list(data.values())
    fig = go.Figure(go.Bar(x=values, y=labels, orientation="h",
                           marker_color=color, text=values, textposition="outside"))
    fig.update_layout(title=title, height=height, yaxis_autorange="reversed",
                      xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                      margin=dict(l=10,r=40,t=40,b=10), showlegend=False, **PT)
    return fig

def bar_v(data, title, color="#0d6efd", height=320):
    labels, values = list(data.keys()), list(data.values())
    fig = go.Figure(go.Bar(x=labels, y=values,
                           marker_color=color, text=values, textposition="outside"))
    fig.update_layout(title=title, height=height,
                      yaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                      margin=dict(l=10,r=10,t=40,b=10), showlegend=False, **PT)
    return fig

def donut(data, title, colors=None, height=280):
    colors = colors or px.colors.qualitative.Set2
    fig = go.Figure(go.Pie(labels=list(data.keys()), values=list(data.values()),
                           hole=0.55, marker_colors=colors,
                           textinfo="percent+label", textfont_size=11))
    fig.update_layout(title=title, height=height,
                      margin=dict(l=10,r=10,t=40,b=10), showlegend=False, **PT)
    return fig

def metric_html(label, value, sub="", color="#212529"):
    return f"""<div class="metric-card">
        <p class="metric-label">{label}</p>
        <p class="metric-value" style="color:{color}">{value}</p>
        <p class="metric-sub">{sub}</p></div>"""

# ── Sidebar ────────────────────────────────────────────────────────────────────
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
    st.markdown("### Filters")

    if "df_all" not in st.session_state:
        st.session_state.df_all = pd.DataFrame()
    df_all = st.session_state.df_all

    status_filter = st.multiselect("Flag status",
        ["Resolved","Unresolved"], default=["Resolved","Unresolved"])
    vendor_opts = sorted(df_all["vendor"].unique()) if not df_all.empty else []
    vendor_filter = st.multiselect("Vendor", vendor_opts, default=vendor_opts)
    priority_filter = st.multiselect("Priority",
        ["High","Medium","Low"], default=["High","Medium","Low"])

    st.divider()
    st.caption("Upload one or more .xlsx Report-27 exports. Multiple files are merged and deduplicated by Bill ID.")

# ── Load data ──────────────────────────────────────────────────────────────────
if uploaded_files:
    dfs = []
    for f in uploaded_files:
        with st.spinner(f"Parsing {f.name}…"):
            dfs.append(load_uploaded_file(f))
    if dfs:
        combined = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=["bill_id"])
        st.session_state.df_all = combined
        df_all = combined
        # refresh sidebar filter options
        vendor_opts = sorted(df_all["vendor"].unique())

if df_all.empty:
    st.title("EnergyCAP Bill Flag Analyzer")
    st.info(
        "👈 Upload one or more **Report-27 Excel files** from EnergyCAP to get started.\n\n"
        "**How to export:** Bills module → Menu (≡) → Report-27 Bill Flags → Export to Excel"
    )
    st.stop()

# ── Apply filters ──────────────────────────────────────────────────────────────
df = df_all.copy()
if status_filter:   df = df[df["status"].isin(status_filter)]
if vendor_filter:   df = df[df["vendor"].isin(vendor_filter)]
if priority_filter: df = df[df["primary_priority"].isin(priority_filter)]

if df.empty:
    st.warning("No records match the current filters.")
    st.stop()

# ── Derived data ───────────────────────────────────────────────────────────────
issues_exp = (df.explode("issues_list")
               .rename(columns={"issues_list": "issue"})
               .query("issue.notna() and issue != ''"))

issue_counts  = Counter(issues_exp["issue"].tolist())
vendor_counts = df["vendor"].value_counts().head(10).to_dict()

period_counts = (df[df["billing_period_label"].notna()]
                 .groupby("billing_period_label").size()
                 .sort_index().to_dict())

assignee_rows = [
    a.strip()
    for _, row in df.iterrows()
    for a in str(row["assigned_to"]).split(",")
    if a.strip()
]
resolver_rows = [r for _, row in df.iterrows() for r in row["resolvers"]]

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

# ──────────────────────────────────────────────────────────────────────────────
# TAB 1 — OVERVIEW
# ──────────────────────────────────────────────────────────────────────────────
with t_overview:
    st.subheader("Summary")
    cols = st.columns(6)
    metrics = [
        ("Total flagged bills",    str(total_bills),
         f"{resolved_count} resolved · {unresolved_count} open", "#212529"),
        ("Resolution rate",        f"{resolution_rate:.0f}%",
         f"{resolved_count} of {total_bills}",
         "#198754" if resolution_rate >= 90 else "#fd7e14"),
        ("Total bill value",       f"${total_cost:,.0f}", "Under review",    "#212529"),
        ("Cost recovered",         f"${total_recovery:,.0f}", "Tracked savings", "#198754"),
        ("Multi-issue bills (3+)", str(multi_issue_cnt),
         f"{multi_issue_cnt/total_bills*100:.0f}% of total", "#6f42c1"),
        ("High-priority open",     str(hp_open), "Need immediate action",
         "#dc3545" if hp_open > 0 else "#198754"),
    ]
    for col, (label, val, sub, color) in zip(cols, metrics):
        col.markdown(metric_html(label, val, sub, color), unsafe_allow_html=True)

    st.markdown("")
    col1, col2 = st.columns([3, 2])
    with col1:
        top12 = dict(sorted(issue_counts.items(), key=lambda x: -x[1])[:12])
        colors_i = [PRIORITY_COLOR.get(FLAG_META.get(k,{}).get("priority","Medium"),"#6c757d")
                    for k in top12]
        fig = go.Figure(go.Bar(x=list(top12.values()), y=list(top12.keys()),
                               orientation="h", marker_color=colors_i,
                               text=list(top12.values()), textposition="outside"))
        fig.update_layout(title="Flag issues by frequency", yaxis_autorange="reversed",
                          xaxis=dict(showgrid=True, gridcolor="#f0f0f0"),
                          height=400, margin=dict(l=10,r=40,t=40,b=10),
                          showlegend=False, **PT)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        if period_counts:
            st.plotly_chart(bar_v(period_counts, "Bills by billing period",
                                  color="#0d6efd", height=400), use_container_width=True)

    c3, c4, c5 = st.columns(3)
    with c3:
        cat_ct = (issues_exp["issue"]
                  .apply(lambda x: FLAG_META.get(x,{}).get("category","Other"))
                  .value_counts().to_dict())
        cat_colors = [CATEGORY_COLOR.get(k,"#888") for k in cat_ct]
        st.plotly_chart(donut(cat_ct, "Issues by category", cat_colors), use_container_width=True)
    with c4:
        st.plotly_chart(donut(df["status"].value_counts().to_dict(),
                              "Flag status", ["#198754","#dc3545"]), use_container_width=True)
    with c5:
        pri_ct = (issues_exp["issue"]
                  .apply(lambda x: FLAG_META.get(x,{}).get("priority","Medium"))
                  .value_counts().reindex(["High","Medium","Low"]).dropna().to_dict())
        st.plotly_chart(donut(pri_ct, "Issues by priority",
                              ["#dc3545","#fd7e14","#198754"]), use_container_width=True)

    resolve_df = df[df["days_to_resolve"].notna() & (df["days_to_resolve"] >= 0)]
    if not resolve_df.empty:
        st.subheader("Resolution time")
        fig_r = px.histogram(resolve_df, x="days_to_resolve", nbins=20,
                             labels={"days_to_resolve": "Days to resolve"},
                             color_discrete_sequence=["#0d6efd"],
                             title=f"Avg {avg_resolve:.1f} days · Median {resolve_df['days_to_resolve'].median():.0f} days")
        fig_r.update_layout(height=240, margin=dict(l=10,r=10,t=40,b=10),
                            showlegend=False, **PT)
        st.plotly_chart(fig_r, use_container_width=True)

# ──────────────────────────────────────────────────────────────────────────────
# TAB 2 — FLAG ANALYSIS
# ──────────────────────────────────────────────────────────────────────────────
with t_flags:
    st.subheader("Flag Issue Deep Dive")

    col1, col2 = st.columns(2)
    with col1:
        selected_issue = st.selectbox(
            "Select a flag issue to analyze",
            options=sorted(issue_counts, key=lambda x: -issue_counts[x]),
        )
    with col2:
        meta = FLAG_META.get(selected_issue, {})
        st.markdown(f"**Category:** {meta.get('category','—')} &nbsp;|&nbsp; "
                    f"**Priority:** {meta.get('priority','—')}")

    issue_df = issues_exp[issues_exp["issue"] == selected_issue]
    ca, cb, cc = st.columns(3)
    ca.metric("Bills affected", len(issue_df))
    cb.metric("Total cost ($)", f"${issue_df['cost'].sum():,.0f}")
    cc.metric("Unresolved", (issue_df["status"] == "Unresolved").sum())

    col1, col2 = st.columns(2)
    with col1:
        v_ct = issue_df["vendor"].value_counts().head(8).to_dict()
        st.plotly_chart(bar_h(v_ct, f"Vendors — '{selected_issue}'",
                              color=CATEGORY_COLOR.get(meta.get("category",""),"#888"), height=300),
                        use_container_width=True)
    with col2:
        a_ct = Counter(
            a for row in issue_df["assigned_to"]
            for a in str(row).split(",") if a.strip()
        )
        st.plotly_chart(bar_h(dict(a_ct.most_common(8)),
                              f"Assignees — '{selected_issue}'",
                              color="#6f42c1", height=300), use_container_width=True)

    if meta:
        with st.expander("📖 Audit rule details", expanded=True):
            st.markdown(f"**What caused this flag:** {meta.get('cause','')}")
            st.markdown(f"**Recommended action:** {meta.get('action','')}")
            st.markdown(f"**In EnergyCAP:** {meta.get('in_energycap','')}")

    st.subheader("Assignee workload")
    st.plotly_chart(bar_h(dict(Counter(assignee_rows).most_common(10)),
                          "Flag assignments per person", color="#0d6efd", height=280),
                    use_container_width=True)

    if resolver_rows:
        st.subheader("Resolver activity")
        st.plotly_chart(bar_h(dict(Counter(resolver_rows).most_common(10)),
                              "Flag resolutions per person", color="#198754", height=280),
                        use_container_width=True)

# ──────────────────────────────────────────────────────────────────────────────
# TAB 3 — VENDORS
# ──────────────────────────────────────────────────────────────────────────────
with t_vendors:
    st.subheader("Vendor Analysis")
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(bar_h(vendor_counts, "Top vendors by flag count",
                              color="#6f42c1", height=360), use_container_width=True)
    with col2:
        vcost = df.groupby("vendor")["cost"].sum().sort_values(ascending=False).head(10).to_dict()
        st.plotly_chart(bar_h(vcost, "Top vendors by total flagged bill value ($)",
                              color="#dc3545", height=360), use_container_width=True)

    st.subheader("Vendor flag profile")
    sel_vendor = st.selectbox("Select vendor", sorted(df["vendor"].unique()))
    vdf = df[df["vendor"] == sel_vendor]
    vdf_i = vdf.explode("issues_list").rename(columns={"issues_list":"issue"})
    vdf_i = vdf_i[vdf_i["issue"].notna() & (vdf_i["issue"] != "")]

    vc1, vc2, vc3, vc4 = st.columns(4)
    vc1.metric("Bills flagged", len(vdf))
    vc2.metric("Unresolved", (vdf["status"]=="Unresolved").sum())
    vc3.metric("Total cost", f"${vdf['cost'].sum():,.0f}")
    vc4.metric("Unique issues", vdf_i["issue"].nunique())

    if not vdf_i.empty:
        vi_ct = vdf_i["issue"].value_counts().to_dict()
        st.plotly_chart(bar_h(vi_ct, f"Flag issues — {sel_vendor}",
                              color="#fd7e14", height=max(200, len(vi_ct)*36+60)),
                        use_container_width=True)

    st.subheader("All vendors — summary table")
    vsumm = (df.groupby("vendor")
               .agg(bills=("bill_id","count"),
                    unresolved=("status", lambda x: (x=="Unresolved").sum()),
                    total_cost=("cost","sum"),
                    unique_issues=("issues_list", lambda x: len(set(i for lst in x for i in lst))))
               .sort_values("bills", ascending=False).reset_index())
    vsumm.columns = ["Vendor","Bills","Unresolved","Total Cost ($)","Unique Issue Types"]
    vsumm["Total Cost ($)"] = vsumm["Total Cost ($)"].map("${:,.0f}".format)
    st.dataframe(vsumm, use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────────────────────────────────────
# TAB 4 — ACTION GUIDE
# ──────────────────────────────────────────────────────────────────────────────
with t_actions:
    st.subheader("Actionable Insights & Recommended Next Steps")
    st.caption("Each flag type is mapped to a specific action based on EnergyCAP's audit rule documentation. "
               "Sorted by priority (High → Medium → Low) then occurrence count.")

    # Open flags first
    open_df = df[df["status"] == "Unresolved"]
    if not open_df.empty:
        st.markdown("### 🔴 Open / Unresolved Flags")
        for _, row in open_df.iterrows():
            p = row["primary_priority"]
            css = {"High":"action-high","Medium":"action-medium","Low":"action-low"}.get(p,"action-info")
            bc  = {"High":"red","Medium":"orange","Low":"green"}.get(p,"blue")
            badges = " ".join(f'<span class="flag-badge badge-{bc}">{i}</span>'
                              for i in row["issues_list"])
            st.markdown(f"""<div class="action-card {css}">
                <strong>Bill {row['bill_id']} — {row['vendor']}</strong>
                &nbsp;|&nbsp; ${row['cost']:,.2f}
                &nbsp;|&nbsp; Assigned: {row['assigned_to'] or 'Unassigned'}<br>{badges}
            </div>""", unsafe_allow_html=True)
        st.divider()

    # Per-flag guidance
    sorted_issues = sorted(
        [(i, c) for i, c in issue_counts.items() if i in FLAG_META],
        key=lambda x: (PRIORITY_ORDER.get(FLAG_META[x[0]]["priority"],1), -x[1])
    )
    seen_cats: set = set()
    for issue, cnt in sorted_issues:
        meta = FLAG_META[issue]
        cat = meta["category"]
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
        unr_html = f'&nbsp;<span class="flag-badge badge-red">⚠ {unr} unresolved</span>' if unr else ""
        st.markdown(f"""<div class="action-card {css}">
            <div style="display:flex;justify-content:space-between;align-items:flex-start;">
                <div>
                    <strong>{issue}</strong>
                    &nbsp;<span class="flag-badge badge-{bc}">{p}</span>
                    &nbsp;<span class="flag-badge badge-gray">{cnt} occurrence{'s' if cnt>1 else ''}</span>
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

    # Systemic insights
    st.divider()
    st.markdown("### 💡 Systemic Observations")
    insights = []

    rs = issue_counts.get("Rate schedule mismatch",0)
    sn = issue_counts.get("Serial number mismatch",0)
    if rs + sn > total_bills * 0.4:
        insights.append(("action-info",
            "High volume of import/configuration mismatches",
            f"{rs} rate schedule + {sn} serial number mismatches across {total_bills} bills "
            f"({(rs+sn)/total_bills*100:.0f}% of flag occurrences). "
            "Consider a bulk meter configuration update in EnergyCAP rather than resolving bill-by-bill."))

    dup = (issue_counts.get("Duplicate bill",0) + issue_counts.get("Overlapping bill",0)
           + issue_counts.get("Multiple bills in period",0))
    if dup > 0:
        insights.append(("action-high",
            f"Potential duplicate payments — {dup} overlap/duplicate flags",
            "These carry the highest financial risk. Verify each before releasing to AP. "
            "Cross-reference with your payment system to confirm no duplicates were paid."))

    if vendor_counts:
        topv = max(vendor_counts, key=vendor_counts.get)
        topv_pct = vendor_counts[topv] / total_bills * 100
        if topv_pct > 20:
            insights.append(("action-medium",
                f"High flag concentration: {topv} ({topv_pct:.0f}% of flagged bills)",
                f"{topv} is disproportionately represented. This may indicate a batch import issue, "
                "meter configuration problem, or data quality issue with the vendor's EDI format. "
                "Review the import template and meter mappings for this vendor."))

    if total_recovery == 0 and total_bills > 10:
        insights.append(("action-info",
            "No cost recovery tracked ($0.00 across all flags)",
            "EnergyCAP lets you log Cost Recovery when a billing error is corrected. "
            "Start tracking to quantify the ROI of your flag review process. "
            "This data appears in Dashboard widgets and management reports."))

    for css, title, body in insights:
        st.markdown(f"""<div class="action-card {css}">
            <strong>{title}</strong>
            <p style="margin:6px 0 0;font-size:13px;">{body}</p>
        </div>""", unsafe_allow_html=True)
    if not insights:
        st.success("No major systemic issues detected with current filters.")

# ──────────────────────────────────────────────────────────────────────────────
# TAB 5 — BILL DETAIL
# ──────────────────────────────────────────────────────────────────────────────
with t_detail:
    st.subheader("Bill-level detail")
    c1, c2, c3 = st.columns(3)
    with c1: search = st.text_input("Search bill ID or account", "")
    with c2: issue_f = st.multiselect("Filter by flag issue", sorted(issue_counts))
    with c3: sort_by = st.selectbox("Sort by",
                  ["Cost (high–low)","Cost (low–high)","Days to resolve","Bill ID"])

    ddf = df.copy()
    if search:
        ddf = ddf[ddf["bill_id"].str.contains(search, case=False) |
                  ddf["account"].str.contains(search, case=False, na=False)]
    if issue_f:
        ddf = ddf[ddf["issues_list"].apply(lambda lst: any(i in lst for i in issue_f))]
    sort_map = {"Cost (high–low)":("cost",False),"Cost (low–high)":("cost",True),
                "Days to resolve":("days_to_resolve",False),"Bill ID":("bill_id",True)}
    sc, sa = sort_map[sort_by]
    ddf = ddf.sort_values(sc, ascending=sa, na_position="last")

    disp = ddf[["bill_id","vendor","billing_period_label","cost","status",
                "flag_issues","assigned_to","days_to_resolve","cost_recovery"]].copy()
    disp.columns = ["Bill ID","Vendor","Period","Cost ($)","Status",
                    "Flag Issues","Assigned To","Days to Resolve","Cost Recovery ($)"]
    disp["Cost ($)"]          = disp["Cost ($)"].map("${:,.2f}".format)
    disp["Cost Recovery ($)"] = disp["Cost Recovery ($)"].map("${:,.2f}".format)
    disp["Days to Resolve"]   = disp["Days to Resolve"].apply(
        lambda x: f"{int(x)}d" if pd.notna(x) else "—")
    st.dataframe(disp, use_container_width=True, hide_index=True, height=480)

    buf = io.StringIO()
    disp.to_csv(buf, index=False)
    st.download_button("⬇ Download filtered data as CSV", buf.getvalue(),
                       "bill_flags_export.csv", "text/csv")

    # Bill detail view
    st.subheader("Individual bill inspector")
    bill_ids = ddf["bill_id"].tolist()
    if bill_ids:
        sel = st.selectbox("Select a bill", bill_ids)
        row = ddf[ddf["bill_id"]==sel].iloc[0]
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**Bill ID:** {row['bill_id']}")
            st.markdown(f"**Account:** {row['account']}")
            st.markdown(f"**Vendor:** {row['vendor']}")
            st.markdown(f"**Period:** {row['billing_period_label']}")
            st.markdown(f"**Cost:** ${row['cost']:,.2f}")
            st.markdown(f"**Cost Recovery:** ${row['cost_recovery']:,.2f}")
        with c2:
            st.markdown(f"**Status:** {row['status']}")
            st.markdown(f"**Assigned to:** {row['assigned_to'] or '—'}")
            days = f"{int(row['days_to_resolve'])}d" if pd.notna(row['days_to_resolve']) else "Pending"
            st.markdown(f"**Days to resolve:** {days}")
            if row["resolvers"]:
                st.markdown(f"**Resolved by:** {', '.join(row['resolvers'])}")
        st.markdown("**Flag Issues:**")
        for issue in row["issues_list"]:
            m = FLAG_META.get(issue, {})
            icon = {"High":"🔴","Medium":"🟠","Low":"🟢"}.get(m.get("priority",""), "⚪")
            with st.expander(f"{icon} {issue} — {m.get('category','')}"):
                if m:
                    st.markdown(f"**Cause:** {m.get('cause','')}")
                    st.markdown(f"**Action:** {m.get('action','')}")
                    st.markdown(f"**In EnergyCAP:** {m.get('in_energycap','')}")
                else:
                    st.info("No specific guidance available for this flag type.")

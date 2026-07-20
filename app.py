"""
TA Dashboard Generator - Streamlit App
======================================

Regular Hires only. Three dashboard sections:
  1. Filled Roles       (Hiring Outcomes, Efficiency - Portfolio TTF only)
  2. Open Roles         (Pure Open vs Advanced Stage, Aging, Ownership, Positions w/o Req)
  3. Recruiter & Team   (Onboards, Load, Active Candidates, Portfolio TTF)

Changes in this build:
  - STRICT Regular filter via Col AB (Joined) and Col X (Open Reqs) - Option A
  - Median TTF REMOVED everywhere
  - Data labels added to every bar chart
  - Col W = Job_App_for_Hire_1: Source used as Hiring Source
  - Col X = Job_App_for_Hire_1: Hiring Manager used as Hiring Manager (text)
  - Conversion / Quality section REMOVED
  - Workload vs Onboarding Output scatter REMOVED
"""
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

from ta_engine import (
    load_workbook,
    prepare_joined, prepare_open_reqs, prepare_funnel, prepare_dispositions,
    portfolio_ttf, ttf_by_dimension,
    unique_count, group_count,
    apply_dimension_filters, build_audit,
    aging_vs_declines, trend_insights, recruiter_ttf_outliers,
    aged_reqs_with_dispositions,
    STAGE_DETAIL_ORDER, AGING_ORDER,
)

st.set_page_config(page_title="TA Recruiting Intelligence Dashboard",
                   layout="wide", page_icon="📊")
st.title("Talent Acquisition Dashboard Generator")
st.caption("Regular Hires only (strict Col AB filter) • Portfolio TTF = SUM Col AK ÷ Regular Onboards")

INPUT_DIR = Path("input")
INPUT_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Sidebar - Load workbook
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("1. Load workbook")
    uploaded = st.file_uploader("Upload Recruiting Dashboard.xlsx", type=["xlsx"])
    if uploaded:
        (INPUT_DIR / uploaded.name).write_bytes(uploaded.getvalue())
        st.success(f"Loaded {uploaded.name}")
    files = sorted(INPUT_DIR.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True)
    selected = st.selectbox("Workbook", files,
                            format_func=lambda x: x.name if x else "") if files else None

if not selected:
    st.info("Upload Recruiting Dashboard.xlsx from the left sidebar to begin.")
    st.stop()

bundle = load_workbook(selected)

# ---------------------------------------------------------------------------
# Prepare tables
# ---------------------------------------------------------------------------
joined,    joined_notes    = prepare_joined(bundle.raw.get("joined", pd.DataFrame()))
open_reqs, open_notes      = prepare_open_reqs(bundle.raw.get("open_reqs", pd.DataFrame()))
funnel,    funnel_notes    = prepare_funnel(bundle.raw.get("candidate_status", pd.DataFrame()))
disp,      disp_notes      = prepare_dispositions(bundle.raw.get("dispositions", pd.DataFrame()))
notes = bundle.warnings + joined_notes + open_notes + funnel_notes + disp_notes

# ---------------------------------------------------------------------------
# Sidebar - Filters
# ---------------------------------------------------------------------------
def opts(df: pd.DataFrame, col: str) -> list:
    if df.empty or col not in df.columns:
        return []
    vals = df[col].dropna().astype(str).str.strip()
    vals = vals[(vals != "") & (vals.str.lower() != "unknown") & (vals.str.lower() != "nan")]
    return sorted(vals.unique().tolist())


with st.sidebar:
    st.header("2. Filters")

    st.markdown("**Time Period**")
    year_sel    = st.multiselect("Year",
                                  sorted(joined["_year"].dropna().astype(int).unique()) if "_year" in joined else [])
    quarter_sel = st.multiselect("Quarter", opts(joined, "_quarter"))
    month_sel   = st.multiselect("Month",   opts(joined, "_month"))
    ytd_only    = st.checkbox("YTD only (Jan 1 to today)")

    st.markdown("**Organization**")
    region_sel   = st.multiselect("Region",
                                   sorted(set(opts(joined, "_region") + opts(open_reqs, "_region"))))
    business_sel = st.multiselect("Head of Dept / Business Owner",
                                   sorted(set(opts(joined, "_head_of_department") + opts(open_reqs, "_business_head"))))
    dept_sel     = st.multiselect("Department",
                                   sorted(set(opts(joined, "_department") + opts(open_reqs, "_department"))))
    jobfam_sel   = st.multiselect("Job Family",
                                   sorted(set(opts(joined, "_job_family") + opts(open_reqs, "_job_family"))))

    st.markdown("**People**")
    recruiter_sel = st.multiselect("Recruiter",
                                    sorted(set(opts(joined, "_primary_recruiter") + opts(open_reqs, "_primary_recruiter"))))
    hm_sel        = st.multiselect("Hiring Manager",
                                    sorted(set(opts(joined, "_hiring_manager") + opts(open_reqs, "_hiring_manager"))))
    tam_sel       = st.multiselect("Recruiting Manager", opts(joined, "_recruiting_manager"))

    st.markdown("**Source**")
    source_sel = st.multiselect("Hiring Source (Col W)", opts(joined, "_hiring_source"))


def filter_joined(df: pd.DataFrame) -> pd.DataFrame:
    d = apply_dimension_filters(df, {
        "_year": year_sel, "_quarter": quarter_sel, "_month": month_sel,
        "_region": region_sel, "_department": dept_sel,
        "_head_of_department": business_sel, "_job_family": jobfam_sel,
        "_primary_recruiter": recruiter_sel, "_hiring_manager": hm_sel,
        "_recruiting_manager": tam_sel, "_hiring_source": source_sel,
    })
    if ytd_only and "_hire_date" in d.columns:
        year_start = pd.Timestamp(pd.Timestamp.today().year, 1, 1)
        d = d[(d["_hire_date"] >= year_start) & (d["_hire_date"] <= pd.Timestamp.today())]
    return d


def filter_open(df: pd.DataFrame) -> pd.DataFrame:
    return apply_dimension_filters(df, {
        "_region": region_sel, "_business_head": business_sel,
        "_department": dept_sel, "_job_family": jobfam_sel,
        "_primary_recruiter": recruiter_sel, "_hiring_manager": hm_sel,
        "_recruiting_manager": tam_sel,
    })


def filter_funnel(df: pd.DataFrame) -> pd.DataFrame:
    return apply_dimension_filters(df, {
        "_region": region_sel, "_department": dept_sel,
        "_hiring_manager": hm_sel, "_recruiter": recruiter_sel,
        "_hiring_source": source_sel,
    })


j  = filter_joined(joined)
o  = filter_open(open_reqs)
fn = filter_funnel(funnel)

# ---------------------------------------------------------------------------
# Data label helpers
# ---------------------------------------------------------------------------
def bar_h(df, x, y, title, key, hover_data=None):
    """Horizontal bar with data labels."""
    fig = px.bar(df, x=x, y=y, orientation="h", text=x,
                 hover_data=hover_data or [], title=title)
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(margin=dict(t=50, r=80, l=10, b=10))
    st.plotly_chart(fig, key=key, use_container_width=True)


def bar_v(df, x, y, title, key, hover_data=None):
    """Vertical bar with data labels."""
    fig = px.bar(df, x=x, y=y, text=y,
                 hover_data=hover_data or [], title=title)
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(margin=dict(t=50, r=10, l=10, b=10))
    st.plotly_chart(fig, key=key, use_container_width=True)


def line_v(df, x, y, title, key):
    fig = px.line(df, x=x, y=y, markers=True, text=y if isinstance(y, str) else None,
                  title=title)
    fig.update_traces(textposition="top center")
    st.plotly_chart(fig, key=key, use_container_width=True)


def pie(df, names, values, title, key):
    fig = px.pie(df, names=names, values=values, hole=0.4, title=title)
    fig.update_traces(textposition="inside", textinfo="value+percent+label")
    st.plotly_chart(fig, key=key, use_container_width=True)


# ---------------------------------------------------------------------------
# Executive summary — Portfolio TTF only (no Median)
# ---------------------------------------------------------------------------
ttf_stats = portfolio_ttf(j)

st.subheader("Executive Summary")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Regular Onboards", len(j))
c2.metric("Total Days (Col AK)", f"{int(ttf_stats['Total Days']):,}" if ttf_stats["Onboards"] else "0")
c3.metric("Portfolio TTF (Avg days)",
          f"{ttf_stats['TTF']:.1f}" if pd.notna(ttf_stats["TTF"]) else "N/A",
          help="SUM(Total Days to Onboard) / COUNT(Regular Onboards)")
c4.metric("Open Requisitions (Regular)", unique_count(o, "_req_id"))
if not o.empty:
    ou = o.drop_duplicates("_req_id")
    ou = ou[ou["_req_id"].astype(str).str.strip() != ""]
    pure_open_count = int((ou["_stage_class"] == "Pure Open").sum())
    advanced_count  = int((ou["_stage_class"] == "Advanced Stage").sum())
    c5.metric("Pure Open / Advanced", f"{pure_open_count} / {advanced_count}")
else:
    c5.metric("Pure Open / Advanced", "0 / 0")

# Second KPI row
c6, c7, c8, c9, c10 = st.columns(5)
if not o.empty:
    positions_wo_req = int(o["_position_without_req"].sum())
    c6.metric("Positions without Req", positions_wo_req,
              help="Column B populated but Column A blank")
    aging_180 = int(o.drop_duplicates("_req_id")["_aging_bucket"]
                    .astype(str).str.contains("180", na=False).sum())
    c7.metric("180+ Day Reqs", aging_180)
else:
    c6.metric("Positions without Req", 0)
    c7.metric("180+ Day Reqs", 0)

c8.metric("Advanced Candidates", len(fn))
c9.metric("Did-not-join Records",
          int(disp.get("_did_not_join_after_offer", pd.Series(dtype=bool)).sum())
          if not disp.empty else 0)
c10.metric("Sheets Detected", len([k for k, v in bundle.detected.items() if v]))

with st.expander("Data audit & mapping notes", expanded=False):
    st.dataframe(build_audit(bundle, joined, open_reqs, funnel, disp),
                 use_container_width=True)
    for n in notes:
        st.write("-", n)

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "1 · Filled Roles", "2 · Open Roles", "3 · Recruiter & Team Performance", "4 · Insights"
])


# ============================================================================
# 1. FILLED ROLES
# ============================================================================
with tab1:
    st.header("Filled Roles - Regular Onboards Only")
    if j.empty:
        st.warning("No Regular onboards available after filters.")
    else:
        st.markdown("### Hiring Outcomes")
        top = st.columns(4)
        top[0].metric("Total Regular Onboards", len(j))
        if "_hiring_source" in j.columns:
            top[1].metric("Distinct Sources", j["_hiring_source"].nunique())
        if "_primary_recruiter" in j.columns:
            top[2].metric("Distinct Recruiters", j["_primary_recruiter"].nunique())
        if "_department" in j.columns:
            top[3].metric("Distinct Departments", j["_department"].nunique())

        # Trend
        if "_month" in j.columns:
            m = j.groupby("_month").size().reset_index(name="Onboards").sort_values("_month")
            line_v(m, "_month", "Onboards", "Monthly Onboard Trend", "f_month")
        if "_quarter" in j.columns:
            q = j.groupby("_quarter").size().reset_index(name="Onboards").sort_values("_quarter")
            bar_v(q, "_quarter", "Onboards", "Quarter-over-Quarter Onboards", "f_qoq")

        # Onboards by dimensions
        cols = st.columns(3)
        with cols[0]:
            g = group_count(j, "_primary_recruiter", "Recruiter")
            if not g.empty:
                bar_h(g, "Count", "Recruiter", "Onboards by Recruiter", "f_rec")
        with cols[1]:
            g = group_count(j, "_head_of_department", "Head of Dept")
            if not g.empty:
                bar_h(g, "Count", "Head of Dept", "Onboards by Head of Department", "f_hod")
        with cols[2]:
            g = group_count(j, "_department", "Department")
            if not g.empty:
                bar_h(g, "Count", "Department", "Onboards by Department", "f_dept")

        cols2 = st.columns(3)
        with cols2[0]:
            g = group_count(j, "_region", "Region")
            if not g.empty:
                bar_h(g, "Count", "Region", "Onboards by Region", "f_reg")
        with cols2[1]:
            g = group_count(j, "_hiring_source", "Source")
            if not g.empty:
                bar_h(g, "Count", "Source", "Onboards by Source (Col W)", "f_src")
        with cols2[2]:
            g = group_count(j, "_hiring_manager", "Hiring Manager")
            if not g.empty:
                bar_h(g, "Count", "Hiring Manager", "Onboards by Hiring Manager (Col X)", "f_hm")

        # Job Family (if present)
        g_jf = group_count(j, "_job_family", "Job Family")
        if not g_jf.empty:
            bar_h(g_jf, "Count", "Job Family", "Onboards by Job Family", "f_jf")

        # ---- Efficiency (Portfolio TTF only - Median REMOVED) ------------
        st.markdown("### Efficiency — Portfolio TTF")
        st.caption("TTF = SUM(Total Days to Onboard from Col AK) ÷ COUNT(Regular onboards). "
                   "Outliers (blank / <0 / >365) excluded.")

        eff = st.columns(3)
        eff[0].metric("Total Days (Col AK sum)", f"{int(ttf_stats['Total Days']):,}")
        eff[1].metric("Regular Onboards used", int(ttf_stats["Onboards"]))
        eff[2].metric("Portfolio TTF (Avg days)",
                       f"{ttf_stats['TTF']:.1f}" if pd.notna(ttf_stats["TTF"]) else "N/A")

        # TTF cuts by dimensions (no Median column)
        st.markdown("#### TTF Cuts")
        colst = st.columns(2)
        with colst[0]:
            t = ttf_by_dimension(j, "_region", "Region")
            if not t.empty:
                bar_v(t, "Region", "TTF", "Portfolio TTF by Region",
                      "ttf_region", hover_data=["Onboards", "Total Days"])
                st.dataframe(t[["Region", "Onboards", "Total Days", "TTF"]],
                             use_container_width=True)
        with colst[1]:
            t = ttf_by_dimension(j, "_department", "Department")
            if not t.empty:
                bar_v(t, "Department", "TTF", "Portfolio TTF by Department",
                      "ttf_dept", hover_data=["Onboards", "Total Days"])
                st.dataframe(t[["Department", "Onboards", "Total Days", "TTF"]],
                             use_container_width=True)

        colst2 = st.columns(2)
        with colst2[0]:
            t = ttf_by_dimension(j, "_primary_recruiter", "Recruiter")
            if not t.empty:
                bar_v(t.head(20), "Recruiter", "TTF", "Portfolio TTF by Recruiter",
                      "ttf_rec", hover_data=["Onboards", "Total Days"])
                st.dataframe(t[["Recruiter", "Onboards", "Total Days", "TTF"]],
                             use_container_width=True)
        with colst2[1]:
            t = ttf_by_dimension(j, "_head_of_department", "Head of Department")
            if not t.empty:
                bar_v(t.head(20), "Head of Department", "TTF",
                      "Portfolio TTF by Head of Department", "ttf_hod",
                      hover_data=["Onboards", "Total Days"])
                st.dataframe(t[["Head of Department", "Onboards", "Total Days", "TTF"]],
                             use_container_width=True)

        t_jf = ttf_by_dimension(j, "_job_family", "Job Family")
        if not t_jf.empty:
            bar_v(t_jf, "Job Family", "TTF", "Portfolio TTF by Job Family",
                  "ttf_jf", hover_data=["Onboards", "Total Days"])

        # QoQ Portfolio TTF trend (no Median)
        if "_quarter" in j.columns and "_ttf_clean" in j.columns:
            qeff = j.dropna(subset=["_ttf_clean"]).groupby("_quarter").agg(
                Onboards=("_ttf_clean", "count"),
                Total_Days=("_ttf_clean", "sum"),
            ).reset_index()
            qeff["Portfolio TTF"] = (qeff["Total_Days"] / qeff["Onboards"]).round(1)
            fig = px.line(qeff, x="_quarter", y="Portfolio TTF", markers=True,
                          text="Portfolio TTF",
                          title="Portfolio TTF Trend (Quarter over Quarter)")
            fig.update_traces(textposition="top center")
            st.plotly_chart(fig, key="f_ttf_qoq", use_container_width=True)


# ============================================================================
# 2. OPEN ROLES
# ============================================================================
with tab2:
    st.header("Open Roles - Regular Reqs Only (Job Req & Open Pos)")
    if o.empty:
        st.warning("No open requisition data after filters.")
    else:
        ou = o.drop_duplicates("_req_id")
        ou = ou[ou["_req_id"].astype(str).str.strip() != ""]

        st.markdown("### Requisition Health")
        rh = st.columns(4)
        rh[0].metric("Open Requisitions", unique_count(ou, "_req_id"))
        rh[1].metric("Pure Open", int((ou["_stage_class"] == "Pure Open").sum()),
                     help="No candidate identified (Col AK)")
        rh[2].metric("Advanced Stage", int((ou["_stage_class"] == "Advanced Stage").sum()),
                     help="Reference / Background / Offer / EA / Ready for Hire")
        rh[3].metric("Positions without Req", int(o["_position_without_req"].sum()),
                     help="Column B populated but Column A blank")

        # Pure Open vs Advanced Stage pie
        stage_cnt = ou["_stage_class"].value_counts().reset_index()
        stage_cnt.columns = ["Stage Class", "Reqs"]
        pie(stage_cnt, "Stage Class", "Reqs",
            "Open Reqs: Pure Open vs Advanced Stage", "o_stagepie")

        # Full stage funnel (Open -> Screening -> Reference -> Background -> Offer -> EA -> Ready for Hire)
        st.markdown("### Stage Funnel (Column AK, detailed)")
        detail_cnt = ou["_stage_detail"].value_counts().reindex(STAGE_DETAIL_ORDER).dropna().reset_index()
        detail_cnt.columns = ["Stage", "Reqs"]
        if not detail_cnt.empty:
            bar_v(detail_cnt, "Stage", "Reqs", "Open Reqs by Detailed Stage", "o_stagefunnel")

        st.markdown("### Aging & Risk")
        aging = group_count(ou, "_aging_bucket", "Aging Bucket", unique_col="_req_id")
        if not aging.empty:
            order = AGING_ORDER
            aging["order"] = aging["Aging Bucket"].map({v: i for i, v in enumerate(order)}).fillna(99)
            aging = aging.sort_values("order").drop(columns="order")
            bar_v(aging, "Aging Bucket", "Count", "Open Reqs by Aging Bucket", "o_aging")

        # Aging × Stage Class heatmap
        heat = pd.crosstab(ou["_aging_bucket"], ou["_stage_class"])
        if not heat.empty:
            fig = px.imshow(heat, text_auto=True, aspect="auto",
                            title="Aging × Stage Class Heatmap",
                            color_continuous_scale="Reds")
            st.plotly_chart(fig, key="o_agingstage", use_container_width=True)

        st.markdown("### Ownership Views")
        cols = st.columns(3)
        with cols[0]:
            g = group_count(ou, "_primary_recruiter", "Recruiter", unique_col="_req_id")
            if not g.empty:
                bar_h(g, "Count", "Recruiter", "Open Roles by Recruiter", "o_rec")
        with cols[1]:
            g = group_count(ou, "_hiring_manager", "Hiring Manager", unique_col="_req_id")
            if not g.empty:
                bar_h(g, "Count", "Hiring Manager", "Open Roles by Hiring Manager", "o_hm")
        with cols[2]:
            g = group_count(ou, "_department", "Function", unique_col="_req_id")
            if not g.empty:
                bar_h(g, "Count", "Function", "Open Roles by Function", "o_dept")

        cols2 = st.columns(3)
        with cols2[0]:
            g = group_count(ou, "_business_head", "Business Owner", unique_col="_req_id")
            if not g.empty:
                bar_h(g, "Count", "Business Owner", "Open Roles by Business Owner", "o_bh")
        with cols2[1]:
            g = group_count(ou, "_region", "Region", unique_col="_req_id")
            if not g.empty:
                bar_h(g, "Count", "Region", "Open Roles by Region", "o_reg")
        with cols2[2]:
            g = group_count(ou, "_candidate_stage", "Stage", unique_col="_req_id")
            if not g.empty:
                bar_h(g.head(15), "Count", "Stage", "Open Roles by Candidate Stage (raw)", "o_stg")

        # Positions without Requisition
        st.markdown("### Positions Open Without Requisition (Col B populated, Col A blank)")
        pwr = o[o["_position_without_req"]]
        if not pwr.empty:
            colsp = st.columns(3)
            colsp[0].metric("Total Positions w/o Req", len(pwr))
            colsp[1].metric("Distinct Business Owners", pwr["_business_head"].nunique())
            colsp[2].metric("Distinct Regions", pwr["_region"].nunique())

            colsp2 = st.columns(3)
            for slot, (colname, title, key) in zip(colsp2, [
                ("_business_head", "Positions w/o Req by Business Owner", "p_bh"),
                ("_region",        "Positions w/o Req by Region",         "p_reg"),
                ("_department",    "Positions w/o Req by Department",     "p_dept"),
            ]):
                g = pwr[colname].fillna("Unknown").astype(str).value_counts().head(15).reset_index()
                g.columns = [colname.strip("_").replace("_", " ").title(), "Positions"]
                if not g.empty:
                    with slot:
                        bar_h(g, "Positions", g.columns[0], title, key)
        else:
            st.info("No positions found where Column B is populated and Column A is blank.")


# ============================================================================
# 3. RECRUITER & TEAM PERFORMANCE
# ============================================================================
with tab3:
    st.header("Recruiter & Team Performance")
    st.caption("Operational workload and cycle-time view. Portfolio TTF = SUM(Col AK) / COUNT(Regular onboards).")

    if j.empty:
        st.warning("No Regular onboards available for recruiter performance analysis.")
    else:
        rec_hires  = group_count(j, "_primary_recruiter", "Recruiter", top=100)
        rec_load   = group_count(o, "_primary_recruiter", "Recruiter",
                                  unique_col="_req_id", top=100) \
                     if not o.empty else pd.DataFrame(columns=["Recruiter", "Count"])
        rec_active = group_count(fn, "_recruiter", "Recruiter", top=100) \
                     if not fn.empty else pd.DataFrame(columns=["Recruiter", "Count"])
        rec_ttf    = ttf_by_dimension(j, "_primary_recruiter", "Recruiter", top=100)

        summary = rec_hires.rename(columns={"Count": "Onboards"}).merge(
            rec_load.rename(columns={"Count": "Open Req Load"}),
            on="Recruiter", how="outer"
        ).merge(
            rec_active.rename(columns={"Count": "Active Candidates"}),
            on="Recruiter", how="outer"
        ).merge(
            rec_ttf[["Recruiter", "Total Days", "TTF"]],
            on="Recruiter", how="left"
        ).fillna(0)

        summary["Recruiter"] = summary["Recruiter"].astype(str)
        summary = summary[summary["Recruiter"].str.lower() != "unknown"]
        summary = summary.sort_values("Onboards", ascending=False)

        st.markdown("### Recruiter Summary")
        st.dataframe(summary, use_container_width=True)

        cols = st.columns(3)
        with cols[0]:
            bar_h(summary.head(20), "Onboards", "Recruiter",
                  "Onboards by Recruiter", "r_hires")
        with cols[1]:
            bar_h(summary.sort_values("Open Req Load", ascending=False).head(20),
                  "Open Req Load", "Recruiter",
                  "Open Requisition Load", "r_load")
        with cols[2]:
            bar_h(summary.sort_values("Active Candidates", ascending=False).head(20),
                  "Active Candidates", "Recruiter",
                  "Active Candidate Load", "r_active")

        # Portfolio TTF by Recruiter (Workload vs Output scatter REMOVED)
        bar_v(summary.sort_values("TTF", ascending=False).head(20),
              "Recruiter", "TTF", "Portfolio TTF by Recruiter", "r_ttf",
              hover_data=["Onboards", "Total Days"])

        if "_recruiting_manager" in j.columns and \
                j["_recruiting_manager"].astype(str).str.lower().nunique() > 1:
            st.markdown("### Metrics by Recruiting Manager")
            t = ttf_by_dimension(j, "_recruiting_manager", "Recruiting Manager")
            if not t.empty:
                bar_v(t.head(20), "Recruiting Manager", "TTF",
                      "Portfolio TTF by Recruiting Manager", "r_mgr_ttf",
                      hover_data=["Onboards", "Total Days"])
                st.dataframe(t[["Recruiting Manager", "Onboards", "Total Days", "TTF"]],
                             use_container_width=True)


# ============================================================================
# 4. INSIGHTS  (aging × decline correlation, QoQ trend narratives, TTF outliers)
# ============================================================================
with tab4:
    st.header("Insights")
    st.caption("Only insights that clear a materiality bar are shown here — "
               "no activity-metric noise, per the 'So What?' filter (impact, "
               "business risk, trend strength, leadership relevance).")

    DIMS = [
        ("_department", "Department"),
        ("_region", "Region"),
        ("_hiring_manager", "Hiring Manager"),
        ("_primary_recruiter", "Recruiter"),
        ("_hiring_source", "Hiring Source"),
    ]

    st.markdown("### Narrative Insights")
    narratives = trend_insights(j, o, disp, DIMS)
    if narratives:
        for n in narratives:
            st.markdown(f"- {n}")
    else:
        st.info("No cuts currently clear the materiality threshold "
                 "(sample size ≥ 5, QoQ shift ≥ 25%, or aged-share ≥ 35% combined with declines). "
                 "This is expected on small or single-period datasets.")

    st.markdown("### Aging vs. Offer-Decline Correlation")
    st.caption("For each cut: how many open reqs are aged 90+ days, native per-req decline counts "
               "(Col AF 'Number of declines from advanced stage'), and post-offer declines from the "
               "Disposition Reasons Report in the same cut. High values on both sides indicate aging "
               "and pipeline breakage are compounding.")
    corr_dim = st.selectbox("Cut by", [label for _, label in DIMS], key="corr_dim")
    dim_key = dict((label, dim) for dim, label in DIMS)[corr_dim]
    corr_tbl = aging_vs_declines(o, disp, dim_key, corr_dim)
    if not corr_tbl.empty:
        st.dataframe(corr_tbl, use_container_width=True)
        risky = corr_tbl[(corr_tbl["Aged Share"] >= 0.35) & (corr_tbl["Total Declines"] > 0)]
        if not risky.empty:
            bar_h(risky, "Aged Share", corr_dim,
                  f"{corr_dim} cuts with high aging + offer declines", "insight_corr")
    else:
        st.info("Not enough data in this cut to compute a reliable correlation "
                "(minimum 5 open reqs per group).")

    st.markdown("### Aged Reqs With a Matching Decline Record (Req ID join)")
    st.caption("Direct match on Requisition ID between aged (90+ day) Open Reqs and the "
               "Disposition Reasons Report — the specific requisitions where this req already "
               "saw a candidate decline or drop out.")
    aged_disp = aged_reqs_with_dispositions(o, disp)
    if not aged_disp.empty:
        st.dataframe(aged_disp, use_container_width=True)
    else:
        st.info("No aged open reqs currently have a matching record in the Disposition Reasons Report.")

    st.markdown("### Recruiter TTF Outliers")
    st.caption("Recruiters whose Portfolio TTF is at least 10 days worse than the team average, "
               "with at least 5 onboards (so it's a real signal, not a small-sample fluke).")
    outliers = recruiter_ttf_outliers(j)
    if not outliers.empty:
        st.dataframe(outliers, use_container_width=True)
        bar_v(outliers, "Recruiter", "Vs Team Avg (days)",
              "Recruiters Slower Than Team Average TTF", "insight_ttf_outlier")
    else:
        st.info("No recruiter is meaningfully slower than the team average right now.")

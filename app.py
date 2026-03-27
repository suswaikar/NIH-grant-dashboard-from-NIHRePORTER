"""
NIH Grant Dashboard v2 — BU / Boston Medical Center
All grant types, multi-PI support, interactive analytics.

Data source: NIH RePORTER API v2 (public, no key needed)

Run:
    streamlit run app.py
"""

import streamlit as st
import requests
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import time
import hmac

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="NIH Grants — BU / BMC",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CUSTOM CSS ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background-color: #0e1117; }

section[data-testid="stSidebar"] {
    background-color: #141820; border-right: 1px solid #2a2f3e;
}
section[data-testid="stSidebar"] * { color: #c8d0e0 !important; }

[data-testid="metric-container"] {
    background: #1a1f2e; border: 1px solid #2a3050;
    border-radius: 8px; padding: 1rem;
}
[data-testid="metric-container"] label {
    color: #7a8aaa !important; font-size: 0.75rem !important;
    letter-spacing: 0.08em; text-transform: uppercase;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #e8f0ff !important; font-family: 'IBM Plex Mono', monospace !important;
    font-size: 1.8rem !important;
}

h1 { color: #e8f0ff !important; font-weight: 600 !important; letter-spacing: -0.02em; }
h2, h3 { color: #c8d0e0 !important; font-weight: 500 !important; }

[data-testid="stDataFrame"] { border: 1px solid #2a3050; border-radius: 8px; overflow: hidden; }

[data-testid="stTabs"] button {
    color: #7a8aaa !important; font-size: 0.85rem !important;
    letter-spacing: 0.05em; text-transform: uppercase;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    color: #60a5fa !important; border-bottom: 2px solid #60a5fa !important;
}

[data-baseweb="select"] { background: #1a1f2e !important; border-color: #2a3050 !important; }

.stButton button {
    background: #1e3a5f !important; color: #60a5fa !important;
    border: 1px solid #2a5080 !important; border-radius: 6px !important;
    font-family: 'IBM Plex Mono', monospace !important; font-size: 0.8rem !important;
}
.stButton button:hover { background: #254a70 !important; }

.stAlert { background: #1a2535 !important; border-color: #2a4060 !important; color: #8ab4d8 !important; }
</style>
""", unsafe_allow_html=True)


# ── PASSWORD GATE ────────────────────────────────────────────────────────────

def check_password():
    """Return True if the user has entered the correct password."""
    # If no password is configured in secrets, skip the gate (local dev)
    if "password" not in st.secrets:
        return True

    if st.session_state.get("authenticated"):
        return True

    st.markdown("### 🔒 Access restricted")
    st.markdown("This dashboard is for BU/BMC faculty and staff.")
    pwd = st.text_input("Enter password", type="password", key="pwd_input")
    if pwd:
        if hmac.compare_digest(pwd, st.secrets["password"]):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not check_password():
    st.stop()


# ── CONSTANTS ────────────────────────────────────────────────────────────────

API_URL = "https://api.reporter.nih.gov/v2/projects/search"
ORG_NAMES = ["BOSTON UNIVERSITY", "BOSTON MEDICAL CENTER"]

FIELDS = [
    "ProjectTitle", "ContactPiName", "PrincipalInvestigators",
    "Organization", "FiscalYear", "AwardAmount", "TotalCost",
    "ProjectStartDate", "ProjectEndDate", "ActivityCode",
    "CoreProjectNum", "AgencyIcAdmin", "ApplId", "IsActive",
    "FundingMechanism", "ProjectNum",
]

K_CODES = {"K01", "K02", "K05", "K06", "K07", "K08", "K12",
           "K22", "K23", "K24", "K25", "K26", "K43", "K76", "K99", "K00"}
R_CODES = {"R00", "R01", "R03", "R13", "R15", "R21", "R33", "R34",
           "R35", "R36", "R37", "R38", "R41", "R42", "R43", "R44",
           "R56", "R61", "R50"}
U_CODES = {"U01", "U10", "U19", "U24", "U34", "U54", "U56",
           "UG1", "UG3", "UH2", "UH3", "UM1", "UM2"}
T_CODES = {"T01", "T02", "T09", "T14", "T15", "T32", "T34", "T35", "T36", "T37", "TL1", "TL4"}
F_CODES = {"F30", "F31", "F32", "F33", "F99"}
P_CODES = {"P01", "P20", "P30", "P40", "P41", "P42", "P50", "P51", "P60", "PL1"}
DP_CODES = {"DP1", "DP2", "DP3", "DP4", "DP5", "DP7"}

CATEGORY_COLORS = {
    "K": "#60a5fa",   # blue
    "R": "#4ade80",   # green
    "U": "#a78bfa",   # purple
    "T": "#fbbf24",   # amber
    "F": "#f59e0b",   # orange
    "P": "#2dd4bf",   # teal
    "DP": "#f472b6",  # pink
    "Other": "#94a3b8",  # gray
}

DARK_LAYOUT = dict(
    plot_bgcolor="#1a1f2e",
    paper_bgcolor="#1a1f2e",
    font_color="#c8d0e0",
)
GRID_COLOR = "#2a3050"


def dark_layout(**overrides):
    """Build a layout dict with dark theme + grid colors, merging any overrides."""
    layout = {**DARK_LAYOUT}
    # Merge xaxis/yaxis with grid color defaults
    xaxis = {"gridcolor": GRID_COLOR, **overrides.pop("xaxis", {})}
    yaxis = {"gridcolor": GRID_COLOR, **overrides.pop("yaxis", {})}
    layout["xaxis"] = xaxis
    layout["yaxis"] = yaxis
    layout.update(overrides)
    return layout


def grant_category(code):
    """Map activity code to broad grant category."""
    code = str(code).upper()
    if code in K_CODES:
        return "K"
    if code in R_CODES:
        return "R"
    if code in U_CODES or code.startswith("U"):
        return "U"
    if code in T_CODES or code.startswith("T"):
        return "T"
    if code in F_CODES or code.startswith("F"):
        return "F"
    if code in P_CODES or code.startswith("P"):
        return "P"
    if code in DP_CODES or code.startswith("DP"):
        return "DP"
    return "Other"


# ── DATA FETCHING ────────────────────────────────────────────────────────────

# Year ranges to query separately (each must stay under 10,000 results)
YEAR_RANGES = [
    list(range(2006, 2015)),
    list(range(2015, 2020)),
    list(range(2020, 2030)),
]


def fetch_paginated(criteria, limit=500):
    """Fetch all records matching criteria, paginating in chunks."""
    results = []
    offset = 0
    while True:
        payload = {
            "criteria": criteria,
            "include_fields": FIELDS,
            "offset": offset,
            "limit": limit,
            "sort_field": "fiscal_year",
            "sort_order": "desc",
        }
        r = requests.post(API_URL, json=payload, timeout=60)
        r.raise_for_status()
        data = r.json()
        results.extend(data["results"])
        total = data["meta"]["total"]
        offset += limit
        if offset >= total or offset >= 10000:
            break
        time.sleep(0.15)
    return results


def fetch_all_grants():
    """Fetch BU/BMC grants via two strategies and merge:
    1) All grants awarded TO BU/BMC (by org name, split by year range)
    2) External grants where a BU/BMC contact PI is any PI (captures co-PI roles elsewhere)
    """
    seen_appl_ids = set()
    all_results = []

    def add_results(results):
        for rec in results:
            appl_id = rec.get("appl_id")
            if appl_id and appl_id not in seen_appl_ids:
                seen_appl_ids.add(appl_id)
                all_results.append(rec)

    # Strategy 1: grants awarded to BU/BMC (split by year to stay under 10K)
    for yr_range in YEAR_RANGES:
        chunk = fetch_paginated({"org_names": ORG_NAMES, "fiscal_years": yr_range})
        add_results(chunk)

    # Strategy 2: find external grants for BU/BMC contact PIs
    # Collect profile_ids of investigators who are contact PI on a BU/BMC grant
    contact_pids = set()
    for rec in all_results:
        for pi in (rec.get("principal_investigators") or []):
            if pi.get("is_contact_pi") and pi.get("profile_id"):
                contact_pids.add(pi["profile_id"])
    # Also include non-contact PIs who appear on multiple BU/BMC grants (likely faculty)
    from collections import Counter
    all_pid_counts = Counter()
    for rec in all_results:
        for pi in (rec.get("principal_investigators") or []):
            pid = pi.get("profile_id")
            if pid:
                all_pid_counts[pid] += 1
    for pid, count in all_pid_counts.items():
        if count >= 3:
            contact_pids.add(pid)

    # Query external grants in batches of 200 profile_ids, split by year range
    profile_list = sorted(contact_pids)
    batch_size = 200
    n_external_before = len(all_results)
    for i in range(0, len(profile_list), batch_size):
        batch = profile_list[i : i + batch_size]
        for yr_range in YEAR_RANGES:
            try:
                chunk = fetch_paginated({"pi_profile_ids": batch, "fiscal_years": yr_range})
                add_results(chunk)
            except Exception as exc:
                import traceback
                print(f"[WARN] Strategy 2 batch failed (profiles {i}-{i+len(batch)-1}, "
                      f"years {yr_range[0]}-{yr_range[-1]}): {exc}")
                traceback.print_exc()
            time.sleep(0.2)

    n_external = len(all_results) - n_external_before
    print(f"[INFO] Strategy 2 added {n_external} external grants "
          f"(total {len(all_results)} awards, {len(contact_pids)} PI profiles queried)")

    return all_results


def parse_grants(results):
    """Parse API results into grants_df and pi_grants_df."""
    grant_rows = []
    pi_rows = []

    for r in results:
        org = r.get("organization") or {}
        ic = r.get("agency_ic_admin") or {}
        pis = r.get("principal_investigators") or []
        code = r.get("activity_code", "") or ""
        appl_id = r.get("appl_id", "")

        # All PI names joined
        pi_names = []
        contact_pi = (r.get("contact_pi_name") or "").strip().title()
        for pi in pis:
            name = (pi.get("full_name") or "").strip().title()
            if name:
                pi_names.append(name)

        org_name = org.get("org_name") or ""
        is_bu_bmc = any(o.lower() in org_name.lower() for o in ORG_NAMES) if org_name else False

        grant_row = {
            "appl_id": appl_id,
            "core_project_num": r.get("core_project_num", ""),
            "project_num": r.get("project_num", ""),
            "fiscal_year": r.get("fiscal_year"),
            "activity_code": code,
            "grant_category": grant_category(code),
            "project_title": r.get("project_title", ""),
            "contact_pi": contact_pi,
            "org_name": org_name,
            "is_bu_bmc": is_bu_bmc,
            "department": org.get("dept_type", "") or "Unknown",
            "ic": ic.get("abbreviation", "") if ic else "",
            "award_amount": r.get("award_amount") or 0,
            "total_cost": r.get("total_cost") or 0,
            "start_date": (r.get("project_start_date") or "")[:10],
            "end_date": (r.get("project_end_date") or "")[:10],
            "is_active": r.get("is_active", False),
            "n_pis": len(pis),
            "is_multi_pi": len(pis) > 1,
            "all_pi_names": "; ".join(pi_names) if pi_names else contact_pi,
            "funding_mechanism": r.get("funding_mechanism", ""),
        }
        grant_rows.append(grant_row)

        # One row per PI per grant for the PI-level view
        if pis:
            for pi in pis:
                pi_rows.append({
                    "appl_id": appl_id,
                    "profile_id": pi.get("profile_id"),
                    "pi_name": (pi.get("full_name") or "").strip().title(),
                    "is_contact_pi": pi.get("is_contact_pi", False),
                })
        else:
            # Fallback: use contact_pi_name
            pi_rows.append({
                "appl_id": appl_id,
                "profile_id": None,
                "pi_name": contact_pi,
                "is_contact_pi": True,
            })

    grants_df = pd.DataFrame(grant_rows)
    grants_df = grants_df.drop_duplicates(subset=["appl_id"])
    grants_df["fiscal_year"] = pd.to_numeric(grants_df["fiscal_year"], errors="coerce")

    pi_df = pd.DataFrame(pi_rows)

    # Build PI-level grants by joining
    pi_grants_df = pi_df.merge(grants_df, on="appl_id", how="left")

    return grants_df, pi_grants_df


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def load_data():
    """Fetch and parse all BU/BMC NIH grants."""
    results = fetch_all_grants()
    return parse_grants(results)


# ── FILTER HELPER ────────────────────────────────────────────────────────────

def apply_filters(df, filters, is_pi_view=False):
    """Apply sidebar filters to a dataframe."""
    fdf = df.copy()
    pi_col = "pi_name" if is_pi_view else "contact_pi"

    if filters["categories"]:
        fdf = fdf[fdf["grant_category"].isin(filters["categories"])]
    if filters["codes"]:
        fdf = fdf[fdf["activity_code"].isin(filters["codes"])]
    if filters["departments"]:
        fdf = fdf[fdf["department"].isin(filters["departments"])]
    if filters["ics"]:
        fdf = fdf[fdf["ic"].isin(filters["ics"])]
    if filters["fy_range"]:
        fdf = fdf[fdf["fiscal_year"].between(filters["fy_range"][0], filters["fy_range"][1])]
    if filters["active_only"]:
        fdf = fdf[fdf["is_active"]]
    if filters.get("bu_bmc_only"):
        fdf = fdf[fdf["is_bu_bmc"]]
    if filters["pi_search"]:
        mask = fdf[pi_col].str.contains(filters["pi_search"], case=False, na=False)
        # In grant-level view, also search the full PI list so co-PIs are found
        if not is_pi_view and "all_pi_names" in fdf.columns:
            mask = mask | fdf["all_pi_names"].str.contains(filters["pi_search"], case=False, na=False)
        fdf = fdf[mask]
    return fdf


# ── LOAD DATA ────────────────────────────────────────────────────────────────

with st.spinner("Loading grants from NIH RePORTER API... (first load may take ~60 seconds, then cached for 6 hours)"):
    try:
        grants_df, pi_grants_df = load_data()
    except Exception as e:
        st.error(f"Failed to fetch data from NIH RePORTER: {e}")
        st.stop()


# ── SIDEBAR ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## NIH Grant Explorer")
    st.markdown("**BU / Boston Medical Center**")
    st.markdown("---")

    if st.button("Refresh data from NIH", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown("### View mode")
    view_mode = st.radio(
        "How to count grants",
        ["Grant-level (no double-counting)", "PI-level (each PI gets credit)"],
        index=0,
        help="Grant-level counts each award once. PI-level counts it once per PI on the grant.",
    )
    is_pi_view = "PI-level" in view_mode

    st.markdown("### Filters")

    all_cats = sorted(grants_df["grant_category"].unique().tolist())
    sel_categories = st.multiselect("Grant category", all_cats, default=[])

    all_codes = sorted(grants_df["activity_code"].dropna().unique().tolist())
    sel_codes = st.multiselect("Activity code", all_codes, default=[])

    all_depts = sorted(grants_df["department"].dropna().unique().tolist())
    sel_depts = st.multiselect("Department", all_depts, default=[])

    all_ics = sorted(grants_df["ic"].dropna().unique().tolist())
    sel_ics = st.multiselect("NIH Institute (IC)", all_ics, default=[])

    fy_min = int(grants_df["fiscal_year"].min())
    fy_max = int(grants_df["fiscal_year"].max())
    fy_range = st.slider("Fiscal year range", fy_min, fy_max, (fy_min, fy_max))

    active_only = st.checkbox("Active awards only", value=False)
    org_scope = st.radio(
        "Award scope",
        ["All (BU/BMC + external co-PI)", "BU/BMC awarded only"],
        index=0,
        help="'All' includes grants at other institutions where a BU/BMC PI is co-PI.",
    )
    pi_search = st.text_input("Search PI name", "")

    st.markdown("---")
    st.markdown(
        f"<small style='color:#4a5568'>Data: NIH RePORTER API v2<br>"
        f"{len(grants_df):,} award records<br>"
        f"{grants_df['org_name'].nunique()} institutions<br>"
        f"{grants_df['activity_code'].nunique()} activity codes<br>"
        f"FY {fy_min}–{fy_max}</small>",
        unsafe_allow_html=True,
    )

filters = {
    "categories": sel_categories,
    "codes": sel_codes,
    "departments": sel_depts,
    "ics": sel_ics,
    "fy_range": fy_range,
    "active_only": active_only,
    "bu_bmc_only": "BU/BMC awarded" in org_scope,
    "pi_search": pi_search,
}

# Apply filters to both views
gdf = apply_filters(grants_df, filters, is_pi_view=False)
pdf = apply_filters(pi_grants_df, filters, is_pi_view=True)

# Choose active dataframe based on view mode
active_df = pdf if is_pi_view else gdf
pi_col = "pi_name" if is_pi_view else "contact_pi"


# ── HEADER ───────────────────────────────────────────────────────────────────

st.markdown("# NIH Grant Dashboard")
view_label = "PI-level view (each PI counted)" if is_pi_view else "Grant-level view (unique awards)"
st.markdown(
    f"**Boston University / Boston Medical Center** &nbsp;|&nbsp; "
    f"<span style='color:#7a8aaa;font-size:0.85rem'>{view_label}</span>",
    unsafe_allow_html=True,
)
st.markdown("---")

# ── KPI METRICS ──────────────────────────────────────────────────────────────

total_awards = len(gdf)
total_funding = gdf["award_amount"].sum()
unique_pis_count = pdf["pi_name"].nunique() if not pdf.empty else 0
active_count = gdf["is_active"].sum()
multi_pi_count = gdf["is_multi_pi"].sum()
n_categories = gdf["grant_category"].nunique()

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total awards", f"{total_awards:,}")
c2.metric("Total funding", f"${total_funding / 1e6:.1f}M")
c3.metric("Unique PIs", f"{unique_pis_count:,}")
c4.metric("Active awards", f"{int(active_count):,}")
c5.metric("Multi-PI awards", f"{int(multi_pi_count):,}")
c6.metric("Grant categories", f"{n_categories}")

st.markdown("---")


# ── TABS ─────────────────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Overview", "Trends", "By Investigator",
    "By Department", "K-to-R Pipeline", "Full Data",
])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1: OVERVIEW
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.subheader("Awards overview")

    col1, col2 = st.columns(2)

    with col1:
        by_year_cat = (
            gdf.groupby(["fiscal_year", "grant_category"])
            .size().reset_index(name="count")
        )
        fig = px.bar(
            by_year_cat, x="fiscal_year", y="count", color="grant_category",
            title="Awards by fiscal year and category",
            labels={"fiscal_year": "Fiscal Year", "count": "Awards", "grant_category": "Category"},
            color_discrete_map=CATEGORY_COLORS,
        )
        fig.update_layout(**dark_layout(), legend_title_text="Category", barmode="stack")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fund_year = gdf.groupby("fiscal_year")["award_amount"].sum().reset_index()
        fig2 = px.area(
            fund_year, x="fiscal_year", y="award_amount",
            title="Total funding by fiscal year",
            labels={"fiscal_year": "Fiscal Year", "award_amount": "Total Award ($)"},
            color_discrete_sequence=["#60a5fa"],
        )
        fig2.update_traces(fill="tozeroy", fillcolor="rgba(96,165,250,0.15)")
        fig2.update_layout(**dark_layout())
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        ic_counts = (
            gdf.groupby("ic").size().reset_index(name="count")
            .sort_values("count", ascending=False).head(15)
        )
        fig3 = px.bar(
            ic_counts, x="count", y="ic", orientation="h",
            title="Top NIH Institutes",
            labels={"count": "Awards", "ic": "Institute"},
            color_discrete_sequence=["#a78bfa"],
        )
        fig3.update_layout(
            **dark_layout(yaxis=dict(categoryorder="total ascending")),
            margin=dict(l=60),
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        cat_summary = (
            gdf.groupby("grant_category")
            .agg(awards=("appl_id", "count"), funding=("award_amount", "sum"))
            .reset_index().sort_values("funding", ascending=False)
        )
        fig4 = px.bar(
            cat_summary, x="grant_category", y="funding",
            title="Funding by grant category",
            labels={"grant_category": "Category", "funding": "Total Funding ($)"},
            color="grant_category",
            color_discrete_map=CATEGORY_COLORS,
            text="awards",
        )
        fig4.update_traces(texttemplate="%{text} awards", textposition="outside")
        fig4.update_layout(**dark_layout(), showlegend=False)
        st.plotly_chart(fig4, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2: TRENDS
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.subheader("Trends over time")

    col1, col2 = st.columns(2)

    with col1:
        cat_year = (
            gdf.groupby(["fiscal_year", "grant_category"])
            .size().reset_index(name="count")
        )
        fig = px.line(
            cat_year, x="fiscal_year", y="count", color="grant_category",
            title="Awards by category over time",
            markers=True,
            color_discrete_map=CATEGORY_COLORS,
            labels={"fiscal_year": "Fiscal Year", "count": "Awards", "grant_category": "Category"},
        )
        fig.update_layout(**dark_layout(), legend_title_text="")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fund_cat_year = (
            gdf.groupby(["fiscal_year", "grant_category"])["award_amount"]
            .sum().reset_index()
        )
        fig2 = px.area(
            fund_cat_year, x="fiscal_year", y="award_amount", color="grant_category",
            title="Funding by category over time",
            color_discrete_map=CATEGORY_COLORS,
            labels={"fiscal_year": "Fiscal Year", "award_amount": "Funding ($)", "grant_category": "Category"},
        )
        fig2.update_layout(**dark_layout(), legend_title_text="")
        st.plotly_chart(fig2, use_container_width=True)

    # Top activity codes over time
    st.markdown("#### Top activity codes over time")
    top_codes = gdf["activity_code"].value_counts().head(10).index.tolist()
    code_year = (
        gdf[gdf["activity_code"].isin(top_codes)]
        .groupby(["fiscal_year", "activity_code"]).size().reset_index(name="count")
    )
    fig3 = px.line(
        code_year, x="fiscal_year", y="count", color="activity_code",
        markers=True,
        title="Top 10 activity codes over time",
        labels={"fiscal_year": "Fiscal Year", "count": "Awards", "activity_code": "Code"},
    )
    fig3.update_layout(**dark_layout(), legend_title_text="Code")
    st.plotly_chart(fig3, use_container_width=True)

    # Multi-PI trend
    col3, col4 = st.columns(2)
    with col3:
        multi_year = (
            gdf.groupby("fiscal_year")["is_multi_pi"]
            .agg(["sum", "count"]).reset_index()
        )
        multi_year.columns = ["fiscal_year", "multi_pi", "total"]
        multi_year["pct_multi"] = (multi_year["multi_pi"] / multi_year["total"] * 100).round(1)
        fig4 = px.bar(
            multi_year, x="fiscal_year", y="pct_multi",
            title="% Multi-PI awards over time",
            labels={"fiscal_year": "Fiscal Year", "pct_multi": "% Multi-PI"},
            color_discrete_sequence=["#a78bfa"],
        )
        fig4.update_layout(**dark_layout())
        st.plotly_chart(fig4, use_container_width=True)

    with col4:
        new_pi_year = (
            pdf.groupby("pi_name")["fiscal_year"].min().reset_index()
        )
        new_pi_year.columns = ["pi_name", "first_fy"]
        new_pi_counts = new_pi_year.groupby("first_fy").size().reset_index(name="new_pis")
        fig5 = px.bar(
            new_pi_counts, x="first_fy", y="new_pis",
            title="New PIs per year (first NIH award at BU/BMC)",
            labels={"first_fy": "Fiscal Year", "new_pis": "New PIs"},
            color_discrete_sequence=["#4ade80"],
        )
        fig5.update_layout(**dark_layout())
        st.plotly_chart(fig5, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3: BY INVESTIGATOR
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    st.subheader("Investigator-level view")

    # Build PI summary from pi_grants_df (captures all PIs including multi-PI)
    pi_summary = (
        pdf.groupby("pi_name")
        .agg(
            total_awards=("appl_id", "nunique"),
            total_funding=("award_amount", "sum"),
            department=("department", lambda x: x.mode().iloc[0] if len(x) else ""),
            categories=("grant_category", lambda x: ", ".join(sorted(x.unique()))),
            codes=("activity_code", lambda x: ", ".join(sorted(x.unique()))),
            fy_range=("fiscal_year", lambda x: f"{int(x.min())}–{int(x.max())}"),
            active=("is_active", "any"),
            n_as_contact=("is_contact_pi", "sum"),
            n_as_co_pi=("is_contact_pi", lambda x: (~x).sum()),
        )
        .reset_index()
        .sort_values("total_funding", ascending=False)
    )

    col1, col2 = st.columns([2, 1])
    with col1:
        top_n = st.slider("Show top N PIs by funding", 10, 50, 20, key="pi_topn")
    with col2:
        sort_by = st.selectbox("Sort by", ["total_funding", "total_awards"], key="pi_sort")

    top_pis = pi_summary.sort_values(sort_by, ascending=False).head(top_n)

    fig = px.bar(
        top_pis.sort_values(sort_by),
        x=sort_by, y="pi_name", orientation="h",
        color="department",
        title=f"Top {top_n} PIs — {sort_by.replace('_', ' ').title()}",
        hover_data=["codes", "fy_range", "n_as_contact", "n_as_co_pi"],
        color_discrete_sequence=px.colors.qualitative.Pastel,
        labels={"pi_name": "PI", "total_funding": "Total Funding ($)", "total_awards": "Awards"},
    )
    fig.update_layout(
        **dark_layout(yaxis=dict(categoryorder="total ascending")),
        height=max(400, top_n * 22),
        margin=dict(l=200),
    )
    st.plotly_chart(fig, use_container_width=True)

    # PI drill-down
    st.markdown("#### Investigator detail")
    all_pi_names = sorted(pdf["pi_name"].dropna().unique().tolist())
    sel_pi = st.selectbox("Select investigator", all_pi_names, key="pi_select")

    if sel_pi:
        pi_data = pdf[pdf["pi_name"] == sel_pi].sort_values("fiscal_year")
        pi_unique_grants = pi_data.drop_duplicates(subset=["appl_id"])

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Awards", pi_unique_grants["appl_id"].nunique())
        m2.metric("As contact PI", int(pi_data["is_contact_pi"].sum()))
        m3.metric("As co-PI", int((~pi_data["is_contact_pi"]).sum()))
        m4.metric("Total funding", f"${pi_unique_grants['award_amount'].sum():,.0f}")
        m5.metric("Categories", ", ".join(sorted(pi_unique_grants["grant_category"].unique())))

        fig_pi = px.scatter(
            pi_unique_grants, x="fiscal_year", y="award_amount",
            color="grant_category", size="award_amount",
            hover_data=["project_title", "ic", "activity_code", "is_multi_pi"],
            title=f"Grant timeline — {sel_pi}",
            color_discrete_map=CATEGORY_COLORS,
            labels={"fiscal_year": "Fiscal Year", "award_amount": "Award ($)", "grant_category": "Category"},
        )
        fig_pi.update_layout(**dark_layout())
        st.plotly_chart(fig_pi, use_container_width=True)

        show_cols = ["fiscal_year", "activity_code", "grant_category", "project_title",
                     "ic", "award_amount", "is_active", "is_contact_pi", "is_multi_pi",
                     "all_pi_names", "core_project_num"]
        st.dataframe(
            pi_unique_grants[show_cols].rename(columns={
                "fiscal_year": "FY", "activity_code": "Code", "grant_category": "Category",
                "project_title": "Title", "ic": "IC", "award_amount": "Award ($)",
                "is_active": "Active", "is_contact_pi": "Contact PI", "is_multi_pi": "Multi-PI",
                "all_pi_names": "All PIs", "core_project_num": "Project #",
            }),
            use_container_width=True, hide_index=True,
        )

        # Co-PI network
        if pi_unique_grants["is_multi_pi"].any():
            st.markdown("##### Co-investigators")
            co_pis = []
            for _, row in pi_unique_grants[pi_unique_grants["is_multi_pi"]].iterrows():
                names = [n.strip() for n in row["all_pi_names"].split(";")]
                for name in names:
                    if name and name != sel_pi:
                        co_pis.append({"Co-PI": name, "Grant": row["core_project_num"],
                                       "FY": row["fiscal_year"], "Code": row["activity_code"]})
            if co_pis:
                st.dataframe(pd.DataFrame(co_pis), use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4: BY DEPARTMENT
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    st.subheader("Department-level analysis")

    dept_summary = (
        gdf.groupby("department")
        .agg(
            total_awards=("appl_id", "count"),
            total_funding=("award_amount", "sum"),
            unique_pis=("contact_pi", "nunique"),
            active=("is_active", "sum"),
            categories=("grant_category", lambda x: ", ".join(sorted(x.unique()))),
        )
        .reset_index()
        .sort_values("total_funding", ascending=False)
    )

    col1, col2 = st.columns(2)
    with col1:
        fig = px.treemap(
            dept_summary.head(20), path=["department"],
            values="total_funding", color="total_awards",
            title="Funding by department (treemap)",
            color_continuous_scale="Blues",
            labels={"department": "Department", "total_funding": "Funding ($)", "total_awards": "Awards"},
        )
        fig.update_layout(paper_bgcolor="#1a1f2e", font_color="#c8d0e0")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = px.bar(
            dept_summary.head(15).sort_values("total_awards"),
            x="total_awards", y="department", orientation="h",
            color="total_funding",
            title="Awards by department (top 15)",
            color_continuous_scale="Blues",
            labels={"total_awards": "Awards", "department": "Department", "total_funding": "Funding ($)"},
        )
        fig2.update_layout(
            **dark_layout(yaxis=dict(categoryorder="total ascending")),
            height=500,
            margin=dict(l=160),
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### Department trends over time")
    top_depts = dept_summary.head(8)["department"].tolist()
    sel_dept_trend = st.multiselect(
        "Departments to compare", dept_summary["department"].tolist(),
        default=top_depts[:5], key="dept_trend",
    )
    if sel_dept_trend:
        dept_year = (
            gdf[gdf["department"].isin(sel_dept_trend)]
            .groupby(["fiscal_year", "department"]).size().reset_index(name="count")
        )
        fig3 = px.line(
            dept_year, x="fiscal_year", y="count", color="department",
            markers=True, title="Awards per year by department",
            labels={"fiscal_year": "FY", "count": "Awards", "department": "Department"},
        )
        fig3.update_layout(**dark_layout())
        st.plotly_chart(fig3, use_container_width=True)

    st.dataframe(dept_summary, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 5: K-TO-R PIPELINE
# ─────────────────────────────────────────────────────────────────────────────
with tab5:
    st.subheader("K-award to R-award career pipeline")
    st.markdown(
        "Tracks PIs who received mentored career development awards (K) and their "
        "subsequent independent research awards (R, U01). Includes K99-to-R00 transitions."
    )

    # Use pi_grants_df to capture all PIs (not just contact)
    k_mask = pdf["grant_category"] == "K"
    r_mask = pdf["grant_category"] == "R"
    u01_mask = pdf["activity_code"] == "U01"
    independent_mask = r_mask | u01_mask

    # Build career table
    career_rows = []
    for pi_name, grp in pdf.groupby("pi_name"):
        if not pi_name:
            continue
        k_grants = grp[k_mask.reindex(grp.index, fill_value=False)].drop_duplicates(subset=["appl_id"])
        ind_grants = grp[independent_mask.reindex(grp.index, fill_value=False)].drop_duplicates(subset=["appl_id"])

        if k_grants.empty:
            continue

        k_codes = ", ".join(sorted(k_grants["activity_code"].unique()))
        k_fy = int(k_grants["fiscal_year"].min())
        dept = grp["department"].mode().iloc[0] if not grp["department"].empty else ""

        r_codes = ", ".join(sorted(ind_grants["activity_code"].unique())) if not ind_grants.empty else ""
        r_fy = int(ind_grants["fiscal_year"].min()) if not ind_grants.empty else None
        lag = (r_fy - k_fy) if r_fy else None

        # K99 -> R00 specific tracking
        has_k99 = "K99" in k_grants["activity_code"].values
        has_r00 = "R00" in ind_grants["activity_code"].values if not ind_grants.empty else False

        career_rows.append({
            "PI": pi_name,
            "Department": dept,
            "K awards": k_codes,
            "First K FY": k_fy,
            "R/U01 awards": r_codes,
            "First R FY": r_fy,
            "K-to-R lag (yrs)": lag,
            "Converted": not ind_grants.empty,
            "K99-R00": has_k99 and has_r00,
            "Total grants": grp["appl_id"].nunique(),
            "Total funding ($)": grp.drop_duplicates(subset=["appl_id"])["award_amount"].sum(),
            "Active": grp["is_active"].any(),
        })

    career_df = pd.DataFrame(career_rows)

    if career_df.empty:
        st.warning("No K-awardees found in the current filtered data.")
    else:
        converted = career_df[career_df["Converted"]]
        not_converted = career_df[~career_df["Converted"]]
        conversion_rate = len(converted) / len(career_df) * 100

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("PIs with K awards", len(career_df))
        m2.metric("Converted to R/U01", len(converted))
        m3.metric("Conversion rate", f"{conversion_rate:.0f}%")
        avg_lag = converted["K-to-R lag (yrs)"].dropna()
        m4.metric("Avg K-to-R lag", f"{avg_lag.mean():.1f} yrs" if not avg_lag.empty else "N/A")
        k99_r00 = career_df["K99-R00"].sum()
        m5.metric("K99-to-R00", f"{int(k99_r00)}")

        col1, col2 = st.columns(2)
        with col1:
            conv_data = pd.DataFrame({
                "Status": ["Converted to R/U01", "K only (no R yet)"],
                "Count": [len(converted), len(not_converted)],
            })
            fig = px.pie(
                conv_data, names="Status", values="Count",
                title="K-award conversion",
                color_discrete_sequence=["#4ade80", "#60a5fa"],
            )
            fig.update_layout(paper_bgcolor="#1a1f2e", font_color="#c8d0e0")
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            if not converted.empty:
                lag_valid = converted[converted["K-to-R lag (yrs)"].notna() &
                                     (converted["K-to-R lag (yrs)"] >= 0)]
                if not lag_valid.empty:
                    fig2 = px.histogram(
                        lag_valid, x="K-to-R lag (yrs)", nbins=15,
                        title="K-to-R transition lag (years)",
                        labels={"K-to-R lag (yrs)": "Years from first K to first R"},
                        color_discrete_sequence=["#a78bfa"],
                    )
                    fig2.update_layout(**dark_layout())
                    st.plotly_chart(fig2, use_container_width=True)

        # Conversion rate over time
        st.markdown("#### Conversion rate by K-award cohort year")
        if not career_df.empty:
            cohort = career_df.groupby("First K FY").agg(
                total=("PI", "count"),
                converted=("Converted", "sum"),
            ).reset_index()
            cohort["pct"] = (cohort["converted"] / cohort["total"] * 100).round(1)
            # Only show cohorts with enough time to convert (>3 years old)
            recent_cutoff = fy_max - 3
            cohort_mature = cohort[cohort["First K FY"] <= recent_cutoff]
            if not cohort_mature.empty:
                fig3 = px.bar(
                    cohort_mature, x="First K FY", y="pct",
                    title="K-to-R conversion rate by cohort (excludes last 3 years)",
                    labels={"First K FY": "K-Award Cohort Year", "pct": "Conversion Rate (%)"},
                    text="total",
                    color_discrete_sequence=["#4ade80"],
                )
                fig3.update_traces(texttemplate="n=%{text}", textposition="outside")
                fig3.update_layout(**dark_layout())
                st.plotly_chart(fig3, use_container_width=True)

        st.markdown("#### Converted PIs")
        if not converted.empty:
            disp = converted.sort_values("Total funding ($)", ascending=False).copy()
            disp["Total funding ($)"] = disp["Total funding ($)"].apply(lambda x: f"${x:,.0f}")
            disp["Active"] = disp["Active"].map({True: "Yes", False: "No"})
            disp["Converted"] = "Yes"
            disp["K99-R00"] = disp["K99-R00"].map({True: "Yes", False: "No"})
            st.dataframe(disp, use_container_width=True, hide_index=True)

        st.markdown("#### K awardees not yet converted")
        if not not_converted.empty:
            nc = not_converted[["PI", "Department", "K awards", "First K FY",
                                "Active", "Total funding ($)"]].copy()
            nc["Total funding ($)"] = nc["Total funding ($)"].apply(lambda x: f"${x:,.0f}")
            nc["Active"] = nc["Active"].map({True: "Active", False: "Inactive"})
            st.dataframe(
                nc.sort_values("First K FY", ascending=False),
                use_container_width=True, hide_index=True,
            )


# ─────────────────────────────────────────────────────────────────────────────
# TAB 6: FULL DATA
# ─────────────────────────────────────────────────────────────────────────────
with tab6:
    st.subheader(f"All awards ({len(gdf):,} grants, {len(pdf):,} PI-grant records)")

    data_view = st.radio(
        "Show", ["Grant-level data", "PI-level data (one row per PI per grant)"],
        horizontal=True, key="data_view",
    )

    if "Grant-level" in data_view:
        show_cols = [
            "fiscal_year", "contact_pi", "all_pi_names", "n_pis", "department",
            "activity_code", "grant_category", "project_title", "ic",
            "award_amount", "is_active", "start_date", "end_date",
            "core_project_num", "org_name",
        ]
        rename_map = {
            "fiscal_year": "FY", "contact_pi": "Contact PI", "all_pi_names": "All PIs",
            "n_pis": "# PIs", "department": "Department", "activity_code": "Code",
            "grant_category": "Category", "project_title": "Title", "ic": "IC",
            "award_amount": "Award ($)", "is_active": "Active",
            "start_date": "Start", "end_date": "End",
            "core_project_num": "Project #", "org_name": "Organization",
        }
        display_df = (
            gdf[show_cols].rename(columns=rename_map)
            .sort_values(["FY", "Contact PI"], ascending=[False, True])
        )
        st.dataframe(display_df, use_container_width=True, hide_index=True, height=600)

        csv = gdf.to_csv(index=False)
        st.download_button(
            "Download grant-level CSV", data=csv,
            file_name="BU_BMC_NIH_grants.csv", mime="text/csv",
        )
    else:
        show_cols = [
            "fiscal_year", "pi_name", "is_contact_pi", "department",
            "activity_code", "grant_category", "project_title", "ic",
            "award_amount", "is_active", "core_project_num",
        ]
        rename_map = {
            "fiscal_year": "FY", "pi_name": "PI Name", "is_contact_pi": "Contact PI",
            "department": "Department", "activity_code": "Code",
            "grant_category": "Category", "project_title": "Title", "ic": "IC",
            "award_amount": "Award ($)", "is_active": "Active",
            "core_project_num": "Project #",
        }
        display_df = (
            pdf[show_cols].rename(columns=rename_map)
            .sort_values(["FY", "PI Name"], ascending=[False, True])
        )
        st.dataframe(display_df, use_container_width=True, hide_index=True, height=600)

        csv = pdf.to_csv(index=False)
        st.download_button(
            "Download PI-level CSV", data=csv,
            file_name="BU_BMC_NIH_grants_by_PI.csv", mime="text/csv",
        )

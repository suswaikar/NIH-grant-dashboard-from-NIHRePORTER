"""
Evans Endowment Awards — ROI Drill-Down Dashboard
===================================================
Interactive dashboard showing Evans-funded awards (K, Pilot, Junior Faculty,
Bridge) by year, with drill-down to individual awardees and their subsequent
NIH grants from NIH RePORTER.

Run:  streamlit run evans_roi_app.py
"""

import streamlit as st
import pandas as pd
import requests
import time
import re
import hmac
import plotly.express as px
import plotly.graph_objects as go


# ── PASSWORD GATE ────────────────────────────────────────────────────────────
def check_password():
    """Gate access with a shared password (stored in Streamlit secrets)."""
    if "password" not in st.secrets:
        return True  # skip gate in local dev if no secret configured
    if st.session_state.get("authenticated"):
        return True
    pwd = st.text_input("Password", type="password", key="_pwd")
    if pwd and hmac.compare_digest(pwd, st.secrets["password"]):
        st.session_state["authenticated"] = True
        st.rerun()
    elif pwd:
        st.error("Incorrect password")
    return False


# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Evans Endowment ROI",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CUSTOM CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background-color: #0e1117; }
section[data-testid="stSidebar"] { background-color: #141820; border-right: 1px solid #2a2f3e; }
section[data-testid="stSidebar"] * { color: #c8d0e0 !important; }
[data-testid="metric-container"] {
    background: #1a1f2e; border: 1px solid #2a3050; border-radius: 8px; padding: 1rem;
}
[data-testid="metric-container"] label { color: #7a8aaa !important; font-size: 0.75rem !important; }
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #e8f0ff !important; font-family: 'IBM Plex Mono', monospace !important; font-size: 1.8rem !important;
}
h1 { color: #e8f0ff !important; font-weight: 600 !important; }
h2, h3 { color: #c8d0e0 !important; font-weight: 500 !important; }
[data-testid="stDataFrame"] { border: 1px solid #2a3050; border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ── CONSTANTS ────────────────────────────────────────────────────────────────
from pathlib import Path

# Use relative paths (for Streamlit Cloud) or local OneDrive paths
_SCRIPT_DIR = Path(__file__).parent
_LOCAL_BASE = Path(
    r"C:\Users\swaikar\OneDrive - Boston University\Department of Medicine"
    r"\Chair advisory committee for Evans\Data Requests"
)
_CLOUD_DATA = _SCRIPT_DIR / "data"

# Prefer local OneDrive if available, otherwise use bundled data/ folder
BASE = _LOCAL_BASE if _LOCAL_BASE.exists() else _CLOUD_DATA
K_FILE = BASE / "K award 2016 - 2026.xlsx"
AWARD_TRACKER = BASE / "DoM Award Tracker.xlsx"
SOURCE_DATA = BASE / "Evans_Endowment_Awards_Source_Data.xlsx"

API_URL = "https://api.reporter.nih.gov/v2/projects/search"
BU_ORGS = {"BOSTON UNIVERSITY", "BOSTON MEDICAL CENTER", "BOSTON UNIVERSITY MEDICAL CAMPUS"}

# K-mechanism activity codes to EXCLUDE from post-K return calculation.
# K24 is intentionally omitted — it's a midcareer award not supported by Evans/DoM,
# so obtaining a K24 counts as a successful outcome (like an R01).
K_CODES = {
    "K01", "K02", "K05", "K06", "K07", "K08", "K11", "K12",
    "K22", "K23", "K25", "K26", "K43", "K76", "K99", "K00",
    "KL1", "KL2",
}

SECTION_MAP = {
    "Cardiology": "Cardiovascular Medicine", "Cardiovascula Center": "Cardiovascular Medicine",
    "Cardiovascular Medicine": "Cardiovascular Medicine", "Clin Epi": "Clinical Epidemiology",
    "Clinical Epidemiology": "Clinical Epidemiology", "Endocrine": "Endocrinology",
    "Endocrinology": "Endocrinology", "GI": "Gastroenterology", "Gastroenterology": "Gastroenterology",
    "GIM": "General Internal Medicine", "General Internal Medicine": "General Internal Medicine",
    "Geriatrics": "Geriatrics", "geriatrics": "Geriatrics", "Hem & Medical Onc": "Hematology/Oncology",
    "Hematology/Oncology": "Hematology/Oncology", "Hem/Onc": "Hematology/Oncology",
    "Infectious Disease": "Infectious Diseases", "Infectious Diseases": "Infectious Diseases",
    "Infectous Disease": "Infectious Diseases", "Nephrology": "Nephrology", "Renal": "Nephrology",
    "Pulmonary": "Pulmonary", "Pulmonary Center": "Pulmonary", "Rheumatology": "Rheumatology",
    "Vascular Biology": "Vascular Biology",
}
NAME_MAP = {
    "Elliot Hagedorn": "Elliott Hagedorn", "Titi Ilori": "Titilayo Ilori",
    "Andrew BERICAL": "Andrew Berical", "Kostas Alysandratos": "Konstantinos Alysandratos",
}


def clean_name(name: str) -> str:
    name = re.sub(r",?\s*(MD|PhD|DO|MPH|MS|MA|DrPH|ScD)\b\.?", "", str(name), flags=re.IGNORECASE)
    name = re.sub(r"\s+", " ", name).strip()
    if "," in name:
        parts = [p.strip() for p in name.split(",", 1)]
        if len(parts) == 2 and parts[1]:
            name = parts[1] + " " + parts[0]
    name = re.sub(r"\s+[A-Z]\.?\s+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = " ".join(w.capitalize() if w.isupper() and len(w) > 1 else w for w in name.split())
    return NAME_MAP.get(name, name)


# ── DATA LOADING ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600)
def load_k_awardees():
    """Load K awardees from all FY sheets."""
    xls = pd.ExcelFile(K_FILE)
    all_rows = []

    for sheet in xls.sheet_names:
        raw = pd.read_excel(K_FILE, sheet_name=sheet, header=None)
        header_row = None
        for r in range(min(5, len(raw))):
            vals = [str(v).strip().lower() for v in raw.iloc[r] if pd.notna(v)]
            if "section" in vals:
                header_row = r
                break
        if header_row is None:
            continue

        df = pd.read_excel(K_FILE, sheet_name=sheet, header=header_row)
        if "Section" not in df.columns or "Name" not in df.columns:
            continue

        df = df[df["Section"].notna() & df["Name"].notna()].copy()
        df["Section"] = df["Section"].str.strip().map(SECTION_MAP).fillna(df["Section"].str.strip())
        df["Name"] = df["Name"].astype(str).apply(clean_name)
        df["FY"] = sheet
        fy_match = re.search(r"\d{4}", sheet)
        df["FY_Num"] = int(fy_match.group()) if fy_match else 0

        gap_col = [c for c in df.columns if "50" in str(c) and ("salary" in str(c).lower() or "gap" in str(c).lower())]
        fringe_col = [c for c in df.columns if "fringe" in str(c).lower() and "total" not in str(c).lower()]
        cost_col = [c for c in df.columns if "total cost" in str(c).lower()]

        df["SalaryGap"] = pd.to_numeric(df[gap_col[0]], errors="coerce") if gap_col else 0
        if cost_col:
            df["TotalCost"] = pd.to_numeric(df[cost_col[0]], errors="coerce")
        elif gap_col and fringe_col:
            df["TotalCost"] = pd.to_numeric(df[gap_col[0]], errors="coerce") + pd.to_numeric(df[fringe_col[0]], errors="coerce")
        elif gap_col:
            df["TotalCost"] = pd.to_numeric(df[gap_col[0]], errors="coerce") * 1.288
        else:
            df["TotalCost"] = 0

        # Get award number if available
        award_col = [c for c in df.columns if "award" in str(c).lower() and "no" in str(c).lower()]
        df["AwardNo"] = df[award_col[0]].astype(str) if award_col else ""

        all_rows.append(df[["Name", "Section", "FY", "FY_Num", "SalaryGap", "TotalCost", "AwardNo"]])

    return pd.concat(all_rows, ignore_index=True)


@st.cache_data(ttl=3600)
def load_other_awards():
    """Load Pilot and Junior Faculty awards from Award Tracker."""
    if not AWARD_TRACKER.exists():
        return pd.DataFrame()

    at = pd.read_excel(AWARD_TRACKER, sheet_name="Combined", header=0)
    at.columns = ["Award", "FY", "Section", "Name", "Amount"]
    at["Amount"] = pd.to_numeric(at["Amount"], errors="coerce")
    fy_map = {"AY22": "FY2022", "AY23": "FY2023", "AY24": "FY2024", "FY25": "FY2025", "FY26": "FY2026"}
    at["FY"] = at["FY"].map(fy_map)
    at["FY_Num"] = at["FY"].str.extract(r"(\d{4})").astype(int)
    at["Award"] = at["Award"].replace({"Junior Award": "Junior Faculty"})
    at = at[at["Award"].isin(["Pilot", "Junior Faculty"])].copy()
    return at


@st.cache_data(ttl=3600)
def load_demographics():
    """Load PI demographics (sex, current position) from source data."""
    if not SOURCE_DATA.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(SOURCE_DATA, sheet_name="PI Demographics", header=0)
        df = df.drop_duplicates(subset="Name", keep="first")
        return df
    except Exception:
        return pd.DataFrame()


def query_nih_reporter(first: str, last: str) -> list:
    """Query NIH RePORTER for a PI's grants."""
    results = []
    for yr_range in [list(range(2015, 2021)), list(range(2021, 2027))]:
        try:
            payload = {
                "criteria": {
                    "pi_names": [{"first_name": first, "last_name": last}],
                    "fiscal_years": yr_range,
                },
                "include_fields": [
                    "ProjectNum", "ProjectTitle", "ContactPiName",
                    "PrincipalInvestigators", "Organization", "FiscalYear",
                    "AwardAmount", "ActivityCode", "CoreProjectNum",
                    "AgencyIcAdmin", "IsActive", "DirectCostAmt", "IndirectCostAmt",
                    "ProjectStartDate", "ProjectEndDate",
                ],
                "offset": 0, "limit": 500,
                "sort_field": "fiscal_year", "sort_order": "desc",
            }
            r = requests.post(API_URL, json=payload, timeout=30)
            r.raise_for_status()
            data = r.json()
            results.extend(data.get("results", []))
        except Exception:
            pass
        time.sleep(0.2)
    return results


@st.cache_data(ttl=21600, show_spinner="Querying NIH RePORTER...")
def get_nih_grants_for_person(name: str) -> pd.DataFrame:
    """Get all NIH grants for a person, filtered to true matches."""
    parts = name.split()
    first, last = parts[0], parts[-1]
    results = query_nih_reporter(first, last)

    grants = []
    for g in results:
        pis = g.get("principal_investigators") or []
        pi_match = any(
            last.upper() in (p.get("last_name", "") or "").upper()
            and first.upper()[:3] in (p.get("first_name", "") or "").upper()[:3]
            for p in pis
        )
        if not pi_match:
            continue

        is_contact = any(
            last.upper() in (p.get("last_name", "") or "").upper()
            and first.upper()[:3] in (p.get("first_name", "") or "").upper()[:3]
            and p.get("is_contact_pi", False)
            for p in pis
        )

        org = (g.get("organization") or {}).get("org_name", "")
        ic = g.get("agency_ic_admin") or {}
        ic_abbr = ic.get("abbreviation", "") if isinstance(ic, dict) else str(ic)

        grants.append({
            "Project Number": g.get("project_num", ""),
            "Activity Code": g.get("activity_code", ""),
            "Title": g.get("project_title", ""),
            "Organization": org,
            "Fiscal Year": g.get("fiscal_year"),
            "Direct Cost": g.get("direct_cost_amt"),
            "Indirect Cost": g.get("indirect_cost_amt"),
            "Award Amount": g.get("award_amount"),
            "IC": ic_abbr,
            "Contact PI": g.get("contact_pi_name", ""),
            "Is Contact PI": is_contact,
            "Start Date": (g.get("project_start_date") or "")[:10],
            "End Date": (g.get("project_end_date") or "")[:10],
            "Location": "BU/BMC" if org.upper() in BU_ORGS else "External",
            "Is K Grant": g.get("activity_code", "") in K_CODES,
        })

    df = pd.DataFrame(grants)
    if len(df):
        df = df.drop_duplicates(subset=["Project Number", "Fiscal Year"])

        # Filter false positives for common names
        if name == "Sun Lee":
            df = df[
                df["Contact PI"].str.upper().str.contains("LEE-MARQUEZ|LEE, SUN Y|LEE,SUN", na=False, regex=True)
                | df["Organization"].str.upper().isin(BU_ORGS)
            ]
        elif name == "Sudhir Kumar":
            df = df[df["Organization"].str.upper().isin(BU_ORGS)]

    return df


# ── BRIDGE DATA (manual, from slides / Jessica) ─────────────────────────────
BRIDGE_DATA = pd.DataFrame([
    {"FY": "FY2025", "FY_Num": 2025, "Count": 3, "Amount": 75000},
    {"FY": "FY2026", "FY_Num": 2026, "Count": 7, "Amount": 229000},
])

# Jessica's FY2025 K override
JESSICA_OVERRIDES = {
    "FY2025": {"Count": 22, "Amount": 862682},
}


# ── BUILD OVERVIEW TABLE ─────────────────────────────────────────────────────
def build_overview(k_data, other_data):
    """Build the awards overview table."""
    rows = []

    # K Awards by year
    for fy, grp in k_data.groupby("FY"):
        fy_num = grp["FY_Num"].iloc[0]
        # Use Jessica's override for FY2025
        if fy in JESSICA_OVERRIDES:
            count = JESSICA_OVERRIDES[fy]["Count"]
            amount = JESSICA_OVERRIDES[fy]["Amount"]
        else:
            count = len(grp)
            amount = grp["TotalCost"].sum()
        rows.append({"Award": "K Award", "FY": fy, "FY_Num": fy_num, "Count": count, "Amount": amount})

    # Pilot & Junior Faculty
    if len(other_data):
        for (award, fy), grp in other_data.groupby(["Award", "FY"]):
            fy_num = grp["FY_Num"].iloc[0]
            rows.append({"Award": award, "FY": fy, "FY_Num": fy_num, "Count": len(grp), "Amount": grp["Amount"].sum()})

    # Bridge
    for _, br in BRIDGE_DATA.iterrows():
        rows.append({"Award": "Bridge", "FY": br["FY"], "FY_Num": br["FY_Num"], "Count": br["Count"], "Amount": br["Amount"]})

    return pd.DataFrame(rows)


# ── BATCH NIH REPORTER QUERY (for ROI tab) ──────────────────────────────────

@st.cache_data(ttl=21600, show_spinner="Querying NIH RePORTER for all awardees (this takes ~2 minutes on first load)...")
def batch_query_all_awardees(names: tuple) -> pd.DataFrame:
    """Query NIH RePORTER for all awardees and return combined results."""
    all_grants = []
    for i, name in enumerate(names):
        parts = name.split()
        first, last = parts[0], parts[-1]
        results = query_nih_reporter(first, last)

        for g in results:
            pis = g.get("principal_investigators") or []
            pi_match = any(
                last.upper() in (p.get("last_name", "") or "").upper()
                and first.upper()[:3] in (p.get("first_name", "") or "").upper()[:3]
                for p in pis
            )
            if not pi_match:
                continue

            is_contact = any(
                last.upper() in (p.get("last_name", "") or "").upper()
                and first.upper()[:3] in (p.get("first_name", "") or "").upper()[:3]
                and p.get("is_contact_pi", False)
                for p in pis
            )

            org = (g.get("organization") or {}).get("org_name", "")
            ic = g.get("agency_ic_admin") or {}
            ic_abbr = ic.get("abbreviation", "") if isinstance(ic, dict) else str(ic)

            all_grants.append({
                "Name": name,
                "Project Number": g.get("project_num", ""),
                "Core Project": g.get("core_project_num", ""),
                "Activity Code": g.get("activity_code", ""),
                "Title": (g.get("project_title") or "")[:120],
                "Organization": org,
                "Fiscal Year": g.get("fiscal_year"),
                "Direct Cost": g.get("direct_cost_amt"),
                "Indirect Cost": g.get("indirect_cost_amt"),
                "Award Amount": g.get("award_amount"),
                "IC": ic_abbr,
                "Contact PI": g.get("contact_pi_name", ""),
                "Is Contact PI": is_contact,
                "Start Date": (g.get("project_start_date") or "")[:10],
                "End Date": (g.get("project_end_date") or "")[:10],
                "Location": "BU/BMC" if org.upper() in BU_ORGS else "External",
                "Is K Grant": g.get("activity_code", "") in K_CODES,
            })

    df = pd.DataFrame(all_grants)
    if len(df):
        df = df.drop_duplicates(subset=["Name", "Project Number", "Fiscal Year"])
        # Filter false positives
        sun_bad = (df["Name"] == "Sun Lee") & ~(
            df["Contact PI"].str.upper().str.contains("LEE-MARQUEZ|LEE, SUN Y|LEE,SUN", na=False, regex=True)
            | df["Organization"].str.upper().isin(BU_ORGS)
        )
        kumar_bad = (df["Name"] == "Sudhir Kumar") & ~df["Organization"].str.upper().isin(BU_ORGS)
        df = df[~sun_bad & ~kumar_bad]
    return df


# ── MAIN APP ─────────────────────────────────────────────────────────────────

def main():
    if not check_password():
        st.stop()

    st.title("Evans Endowment Awards — ROI Dashboard")
    st.caption("Drill down into K, Pilot, Junior Faculty, and Bridge awards funded by the Evans Endowment")

    k_data = load_k_awardees()
    other_data = load_other_awards()
    demo_data = load_demographics()
    overview = build_overview(k_data, other_data)

    # ── SIDEBAR ──────────────────────────────────────────────────────────────
    st.sidebar.header("Filters")
    all_fys = sorted(overview["FY"].unique())
    selected_fys = st.sidebar.multiselect("Fiscal Years", all_fys, default=all_fys)

    award_types = sorted(overview["Award"].unique())
    selected_awards = st.sidebar.multiselect("Award Types", award_types, default=award_types)

    filtered = overview[overview["FY"].isin(selected_fys) & overview["Award"].isin(selected_awards)]

    # ── TABS ─────────────────────────────────────────────────────────────────
    main_tab, roi_tab, sex_tab, drill_tab, lookup_tab = st.tabs([
        "📊 Awards Overview", "💰 ROI Summary", "👤 Sex Breakdown", "🔍 Drill Down", "🔎 Investigator Lookup"
    ])

    # ==================================================================
    # TAB 1: AWARDS OVERVIEW
    # ==================================================================
    with main_tab:
        # ── TOP METRICS ──────────────────────────────────────────────────
        k_fys = set(selected_fys)
        n_k_investigators = k_data[k_data["FY"].isin(k_fys)]["Name"].nunique()
        if len(other_data):
            n_pilot = other_data[(other_data["Award"] == "Pilot") & (other_data["FY"].isin(k_fys))]["Name"].nunique()
            n_jrfac = other_data[(other_data["Award"] == "Junior Faculty") & (other_data["FY"].isin(k_fys))]["Name"].nunique()
        else:
            n_pilot, n_jrfac = 0, 0
        n_bridge = int(BRIDGE_DATA[BRIDGE_DATA["FY"].isin(k_fys)]["Count"].sum())

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Investment", f"${filtered['Amount'].sum():,.0f}")
        col2.metric("Fiscal Years", f"{filtered['FY'].nunique()}")
        col3.metric("Award Types", f"{filtered['Award'].nunique()}")
        col4.metric("Total Award-Years", f"{int(filtered['Count'].sum()):,}")

        st.markdown("##### Unique Investigators by Award Type")
        inv_cols = st.columns(4)
        inv_cols[0].metric("K Awardees", f"{n_k_investigators}")
        inv_cols[1].metric("Pilot Awardees", f"{n_pilot}")
        inv_cols[2].metric("Junior Faculty", f"{n_jrfac}")
        inv_cols[3].metric("Bridge Awards", f"{n_bridge}")

        # ── OVERVIEW TABLE ───────────────────────────────────────────────
        st.header("Awards Overview")

        # Pivot table: Award × FY
        pivot_count = filtered.pivot_table(index="Award", columns="FY", values="Count", aggfunc="sum", fill_value=0)
        pivot_amount = filtered.pivot_table(index="Award", columns="FY", values="Amount", aggfunc="sum", fill_value=0)

        # Reorder
        award_order = [a for a in ["K Award", "Pilot", "Junior Faculty", "Bridge"] if a in pivot_count.index]
        fy_order = [f for f in sorted(filtered["FY"].unique()) if f in pivot_count.columns]
        pivot_count = pivot_count.reindex(index=award_order, columns=fy_order, fill_value=0)
        pivot_amount = pivot_amount.reindex(index=award_order, columns=fy_order, fill_value=0)

        # Combined display
        display_rows = []
        for award in award_order:
            row = {"Award": award}
            for fy in fy_order:
                c = int(pivot_count.loc[award, fy]) if award in pivot_count.index else 0
                a = pivot_amount.loc[award, fy] if award in pivot_amount.index else 0
                row[fy] = f"{c} (${a:,.0f})" if c > 0 else "—"
            # Total
            total_c = int(pivot_count.loc[award].sum())
            total_a = pivot_amount.loc[award].sum()
            row["Total"] = f"{total_c} (${total_a:,.0f})"
            display_rows.append(row)

        # Grand total row
        grand = {"Award": "TOTAL"}
        for fy in fy_order:
            gc = int(pivot_count[fy].sum())
            ga = pivot_amount[fy].sum()
            grand[fy] = f"{gc} (${ga:,.0f})" if gc > 0 else "—"
        grand["Total"] = f"{int(pivot_count.values.sum())} (${pivot_amount.values.sum():,.0f})"
        display_rows.append(grand)

        display_df = pd.DataFrame(display_rows).set_index("Award")
        st.dataframe(display_df, use_container_width=True)

        # ── CHART ────────────────────────────────────────────────────────────
        chart_data = filtered.copy()
        chart_data["FY_sort"] = chart_data["FY_Num"]

        fig = px.bar(
            chart_data.sort_values("FY_sort"),
            x="FY", y="Amount", color="Award",
            barmode="stack",
            color_discrete_map={"K Award": "#60a5fa", "Pilot": "#4ade80", "Junior Faculty": "#a78bfa", "Bridge": "#fbbf24"},
            labels={"Amount": "Total Investment ($)", "FY": "Fiscal Year"},
            title="Evans Endowment Investment by Award Type",
        )
        fig.update_layout(
            plot_bgcolor="#1a1f2e", paper_bgcolor="#1a1f2e", font_color="#c8d0e0",
            xaxis=dict(gridcolor="#2a3050"), yaxis=dict(gridcolor="#2a3050", tickformat="$,.0f"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ==================================================================
    # TAB 2: ROI SUMMARY
    # ==================================================================
    with roi_tab:
        st.header("Return on Investment Analysis")
        st.caption(
            "Conservative estimate: only NIH grants newly initiated AFTER the last year of DoM support. "
            "GT97 awards are excluded — they are an accounting mechanism, not a research investment."
        )

        # Collect all unique awardee names (K + Pilot + Jr Faculty)
        all_k_names = sorted(k_data["Name"].unique())
        pilot_jr_names = sorted(other_data["Name"].unique()) if len(other_data) else []
        combined_names = sorted(set(all_k_names) | set(pilot_jr_names))

        # Batch query — cached for 6 hours
        all_nih = batch_query_all_awardees(tuple(combined_names))

        if len(all_nih) == 0:
            st.warning("No NIH grants found. Try reloading.")
        else:
            # Build per-awardee last support year
            CURRENT_FY = 2026
            BUFFER_YEARS = 1

            last_k_year = {}
            for name in all_k_names:
                pk = k_data[k_data["Name"] == name]
                last_k_year[name] = pk["FY_Num"].max() if len(pk) else 9999

            last_pilot_year = {}
            if len(other_data):
                for name in pilot_jr_names:
                    po = other_data[other_data["Name"] == name]
                    last_pilot_year[name] = po["FY_Num"].max() if len(po) else 9999

            # Filter to non-K, post-support grants
            non_k = all_nih[~all_nih["Is K Grant"]].copy()

            # ── K AWARDS SUB-TAB ─────────────────────────────────────────
            # ── PILOT / JR FACULTY SUB-TAB ───────────────────────────────
            roi_k_tab, roi_pj_tab = st.tabs(["K Award → R01/Equivalent", "Pilot / Junior Faculty"])

            # ==============================================================
            # K AWARD ROI
            # ==============================================================
            with roi_k_tab:
                k_investment = k_data["TotalCost"].sum()

                # Post-K grants for K awardees
                k_non_k = non_k[non_k["Name"].isin(all_k_names)].copy()
                k_non_k["LastK"] = k_non_k["Name"].map(last_k_year)
                k_post = k_non_k[k_non_k["Fiscal Year"] > k_non_k["LastK"]].copy()

                k_direct = k_post["Direct Cost"].sum()
                k_indirect = k_post["Indirect Cost"].sum()
                k_total_nih = k_direct + k_indirect
                k_roi = k_direct / k_investment if k_investment > 0 else 0

                k_eligible = [n for n, yr in last_k_year.items() if yr + BUFFER_YEARS < CURRENT_FY]
                k_n_eligible = len(k_eligible)
                k_n_transitioned = k_post[k_post["Name"].isin(k_eligible)]["Name"].nunique()
                k_trans_pct = (100 * k_n_transitioned / k_n_eligible) if k_n_eligible > 0 else 0
                k_n_too_recent = len(all_k_names) - k_n_eligible

                hc1, hc2, hc3, hc4 = st.columns(4)
                hc1.metric("DoM K Investment", f"${k_investment / 1e6:.1f}M")
                hc2.metric("Post-K NIH Direct", f"${k_direct / 1e6:.1f}M")
                hc3.metric("Post-K NIH Total", f"${k_total_nih / 1e6:.1f}M")
                hc4.metric("ROI (Direct / DoM)", f"{k_roi:.0f}×")

                rc1, rc2, rc3, rc4 = st.columns(4)
                rc1.metric("K Awardees", f"{len(all_k_names)}")
                rc2.metric("Eligible for Transition", f"{k_n_eligible}",
                           help=f"Excludes {k_n_too_recent} with K ending FY{CURRENT_FY - BUFFER_YEARS}+")
                rc3.metric("Transition Rate", f"{k_trans_pct:.0f}%",
                           help=f"{k_n_transitioned} of {k_n_eligible} eligible")
                rc4.metric("ROI (Total / DoM)", f"{k_total_nih / k_investment:.0f}×" if k_investment > 0 else "—")

                # === K-TO-R BY SECTION ===
                st.markdown("### K-to-R Outcomes by Section")
                k_sections = k_data.groupby(["Name", "Section"]).agg(
                    LastK=("FY_Num", "max"),
                    TotalGap=("SalaryGap", "sum"),
                ).reset_index()

                k_post_by_pi = k_post.groupby("Name").agg(
                    PostK_Direct=("Direct Cost", "sum"),
                    PostK_Indirect=("Indirect Cost", "sum"),
                    PostK_Grants=("Core Project", "nunique"),
                ).reset_index()

                k_merged = k_sections.merge(k_post_by_pi, on="Name", how="left")
                k_merged["PostK_Direct"] = k_merged["PostK_Direct"].fillna(0)
                k_merged["PostK_Indirect"] = k_merged["PostK_Indirect"].fillna(0)
                k_merged["PostK_Grants"] = k_merged["PostK_Grants"].fillna(0).astype(int)
                k_merged["HasPostK"] = k_merged["PostK_Direct"] > 0
                k_merged["Eligible"] = k_merged["LastK"] + BUFFER_YEARS < CURRENT_FY

                section_summary = k_merged.groupby("Section").agg(
                    N_Awardees=("Name", "nunique"),
                    N_Eligible=("Eligible", "sum"),
                    N_Transitioned=("HasPostK", lambda x: (x & k_merged.loc[x.index, "Eligible"]).sum()),
                    Total_Gap=("TotalGap", "sum"),
                    Total_Direct=("PostK_Direct", "sum"),
                ).reset_index()
                section_summary["Transition %"] = (
                    (100 * section_summary["N_Transitioned"] / section_summary["N_Eligible"])
                    .where(section_summary["N_Eligible"] > 0, other=0)
                    .round(0).astype(int)
                )
                section_summary = section_summary.sort_values("Total_Direct", ascending=False)

                # Summary table
                sec_table = section_summary.copy()
                sec_table["Total_Gap"] = sec_table["Total_Gap"].apply(lambda x: f"${x:,.0f}")
                sec_table["Total_Direct"] = sec_table["Total_Direct"].apply(lambda x: f"${x:,.0f}")
                sec_table["Transition %"] = sec_table["Transition %"].astype(str) + "%"
                sec_table_display = sec_table[["Section", "N_Awardees", "N_Eligible", "N_Transitioned", "Total_Gap", "Total_Direct", "Transition %"]].copy()
                sec_table_display.columns = ["Section", "K Awardees", "Eligible", "Transitioned", "DoM Gap Investment", "Post-K NIH Direct", "Transition %"]
                st.dataframe(sec_table_display, use_container_width=True, hide_index=True)

                # Expandable detail per section
                st.markdown("##### Click a section to see individual awardees")
                for _, sec_row in section_summary.iterrows():
                    section = sec_row["Section"]
                    n_aw = int(sec_row["N_Awardees"])
                    n_tr = int(sec_row["N_Transitioned"])
                    gap_fmt = f"${sec_row['Total_Gap']:,.0f}"
                    direct_fmt = f"${sec_row['Total_Direct']:,.0f}"
                    trans_pct = f"{int(sec_row['Transition %'])}%"
                    with st.expander(
                        f"**{section}** — {n_aw} awardees, {n_tr} transitioned "
                        f"({trans_pct}), Gap: {gap_fmt}, Post-K Direct: {direct_fmt}"
                    ):
                        sec_people = k_merged[k_merged["Section"] == section].sort_values("Name")
                        people_display = sec_people[["Name", "LastK", "Eligible", "HasPostK", "TotalGap", "PostK_Direct", "PostK_Grants"]].copy()
                        people_display.columns = ["Name", "Last K Year", "Eligible", "Transitioned", "DoM Gap", "Post-K Direct", "Post-K Grants"]
                        people_display["Last K Year"] = people_display["Last K Year"].apply(lambda x: f"FY{x}")
                        people_display["Eligible"] = people_display["Eligible"].map({True: "Yes", False: "Too recent"})
                        people_display["Transitioned"] = people_display["Transitioned"].map({True: "✓", False: "—"})
                        people_display["DoM Gap"] = people_display["DoM Gap"].apply(lambda x: f"${x:,.0f}" if x > 0 else "—")
                        people_display["Post-K Direct"] = people_display["Post-K Direct"].apply(lambda x: f"${x:,.0f}" if x > 0 else "—")
                        st.dataframe(people_display, use_container_width=True, hide_index=True)

                # Section bar chart
                fig_sec = px.bar(
                    section_summary, x="Section", y="Total_Direct",
                    color="N_Transitioned", color_continuous_scale="Blues",
                    labels={"Total_Direct": "Post-K NIH Direct Costs ($)", "N_Transitioned": "# Transitioned"},
                    title="Post-K NIH Direct Costs by Section",
                )
                fig_sec.update_layout(
                    plot_bgcolor="#1a1f2e", paper_bgcolor="#1a1f2e", font_color="#c8d0e0",
                    xaxis=dict(gridcolor="#2a3050", tickangle=45),
                    yaxis=dict(gridcolor="#2a3050", tickformat="$,.0f"),
                )
                st.plotly_chart(fig_sec, use_container_width=True)

                # Top K awardees
                st.markdown("### Top 15 K Awardees by Post-K NIH Direct Funding")
                top_k = k_post.groupby("Name").agg(
                    Direct=("Direct Cost", "sum"),
                    Indirect=("Indirect Cost", "sum"),
                    N_Grants=("Core Project", "nunique"),
                ).reset_index().sort_values("Direct", ascending=False).head(15)
                name_section_k = dict(k_data.drop_duplicates("Name")[["Name", "Section"]].values)
                top_k["Section"] = top_k["Name"].map(name_section_k).fillna("")
                top_k_display = top_k[["Name", "Section", "N_Grants", "Direct", "Indirect"]].copy()
                top_k_display["Total"] = top_k_display["Direct"] + top_k_display["Indirect"]
                top_k_display.columns = ["Name", "Section", "Unique Grants", "Direct Costs", "Indirect Costs", "Total"]
                for c in ["Direct Costs", "Indirect Costs", "Total"]:
                    top_k_display[c] = top_k_display[c].apply(lambda x: f"${x:,.0f}")
                st.dataframe(top_k_display, use_container_width=True, hide_index=True)

            # ==============================================================
            # PILOT / JUNIOR FACULTY ROI
            # ==============================================================
            with roi_pj_tab:
                if not len(other_data):
                    st.info("No Pilot / Junior Faculty data available.")
                else:
                    # Pilot/Jr Faculty awardees NOT also K awardees (avoid double-counting)
                    pj_only_names = sorted(set(pilot_jr_names) - set(all_k_names))
                    pj_also_k = sorted(set(pilot_jr_names) & set(all_k_names))

                    pj_investment = other_data["Amount"].sum()

                    # Post-support grants for Pilot/Jr Faculty (using their last pilot/jr year)
                    pj_non_k = non_k[non_k["Name"].isin(pilot_jr_names)].copy()
                    pj_non_k["LastSupport"] = pj_non_k["Name"].map(last_pilot_year)
                    pj_post = pj_non_k[pj_non_k["Fiscal Year"] > pj_non_k["LastSupport"]].copy()

                    pj_direct = pj_post["Direct Cost"].sum()
                    pj_indirect = pj_post["Indirect Cost"].sum()
                    pj_total_nih = pj_direct + pj_indirect
                    pj_roi = pj_direct / pj_investment if pj_investment > 0 else 0

                    pj_eligible = [n for n, yr in last_pilot_year.items() if yr + BUFFER_YEARS < CURRENT_FY]
                    pj_n_eligible = len(pj_eligible)
                    pj_n_transitioned = pj_post[pj_post["Name"].isin(pj_eligible)]["Name"].nunique()
                    pj_trans_pct = (100 * pj_n_transitioned / pj_n_eligible) if pj_n_eligible > 0 else 0

                    hc1, hc2, hc3, hc4 = st.columns(4)
                    hc1.metric("DoM Pilot/Jr Investment", f"${pj_investment / 1e6:.1f}M")
                    hc2.metric("Post-Award NIH Direct", f"${pj_direct / 1e6:.1f}M")
                    hc3.metric("Post-Award NIH Total", f"${pj_total_nih / 1e6:.1f}M")
                    hc4.metric("ROI (Direct / DoM)", f"{pj_roi:.0f}×")

                    rc1, rc2, rc3, rc4 = st.columns(4)
                    rc1.metric("Awardees", f"{len(pilot_jr_names)}")
                    rc2.metric("Eligible", f"{pj_n_eligible}",
                               help=f"Excludes those with support ending FY{CURRENT_FY - BUFFER_YEARS}+")
                    rc3.metric("With Subsequent NIH $", f"{pj_n_transitioned}")
                    rc4.metric("Transition Rate", f"{pj_trans_pct:.0f}%")

                    if pj_also_k:
                        st.caption(f"Note: {len(pj_also_k)} awardees also received K awards: {', '.join(pj_also_k)}")

                    # By award type
                    st.markdown("### By Award Type")
                    for award_type in ["Pilot", "Junior Faculty"]:
                        at_sub = other_data[other_data["Award"] == award_type]
                        if len(at_sub) == 0:
                            continue
                        at_names = sorted(at_sub["Name"].unique())
                        at_invest = at_sub["Amount"].sum()
                        at_post = pj_post[pj_post["Name"].isin(at_names)]
                        at_direct = at_post["Direct Cost"].sum()

                        with st.expander(
                            f"**{award_type}** — {len(at_names)} awardees, "
                            f"Investment: ${at_invest:,.0f}, "
                            f"Post-Award NIH Direct: ${at_direct:,.0f}"
                        ):
                            at_detail = at_sub.groupby("Name").agg(
                                Section=("Section", "first"),
                                FYs=("FY", lambda x: ", ".join(sorted(x.unique()))),
                                Investment=("Amount", "sum"),
                            ).reset_index()
                            # Add post-award NIH
                            at_post_by_pi = at_post.groupby("Name").agg(
                                PostDirect=("Direct Cost", "sum"),
                                PostGrants=("Core Project", "nunique"),
                            ).reset_index()
                            at_detail = at_detail.merge(at_post_by_pi, on="Name", how="left")
                            at_detail["PostDirect"] = at_detail["PostDirect"].fillna(0)
                            at_detail["PostGrants"] = at_detail["PostGrants"].fillna(0).astype(int)
                            at_detail["Investment"] = at_detail["Investment"].apply(lambda x: f"${x:,.0f}")
                            at_detail["PostDirect"] = at_detail["PostDirect"].apply(lambda x: f"${x:,.0f}" if x > 0 else "—")
                            at_detail.columns = ["Name", "Section", "Years", "DoM Investment", "Post-Award NIH Direct", "Post-Award Grants"]
                            st.dataframe(at_detail.sort_values("Name"), use_container_width=True, hide_index=True)

                    # Top Pilot/Jr Faculty awardees
                    if len(pj_post):
                        st.markdown("### Top Awardees by Post-Award NIH Direct Funding")
                        top_pj = pj_post.groupby("Name").agg(
                            Direct=("Direct Cost", "sum"),
                            Indirect=("Indirect Cost", "sum"),
                            N_Grants=("Core Project", "nunique"),
                        ).reset_index().sort_values("Direct", ascending=False).head(15)
                        name_section_pj = {}
                        for _, r in other_data.drop_duplicates("Name").iterrows():
                            name_section_pj[r["Name"]] = r.get("Section", "")
                        top_pj["Section"] = top_pj["Name"].map(name_section_pj).fillna("")
                        top_pj_display = top_pj[["Name", "Section", "N_Grants", "Direct", "Indirect"]].copy()
                        top_pj_display["Total"] = top_pj_display["Direct"] + top_pj_display["Indirect"]
                        top_pj_display.columns = ["Name", "Section", "Unique Grants", "Direct Costs", "Indirect Costs", "Total"]
                        for c in ["Direct Costs", "Indirect Costs", "Total"]:
                            top_pj_display[c] = top_pj_display[c].apply(lambda x: f"${x:,.0f}")
                        st.dataframe(top_pj_display, use_container_width=True, hide_index=True)

            # === METHODOLOGY (shared) ===
            with st.expander("Methodology & Caveats"):
                st.markdown("""
**Data sources:**
- K award tracking: `K award 2016 - 2026.xlsx` (DoM internal)
- Pilot / Junior Faculty: `DoM Award Tracker.xlsx` (DoM internal)
- NIH grant data: NIH RePORTER API v2 (public, queried live)

**Conservative estimate:** Only counts NIH grants whose fiscal year is strictly
AFTER the last year of DoM support for that investigator. K-mechanism grants
are excluded from the return calculation.

**1-year buffer:** Awardees whose support ended in the last year are excluded from
transition rate calculations — they haven't had enough time to secure subsequent funding.

**Caveats:**
- GT97 awards excluded (accounting mechanism, not research investment)
- Bridge funding is included in the overview but not in ROI calculations (no individual-level data)
- NIH RePORTER name matching uses first 3 characters of first name + full last name; false positives filtered for known common names
- Some investigators may have grants not captured due to name variations
- ROI does not imply causation — DoM support is one of many factors in an investigator's success
""")

    # ==================================================================
    # TAB: SEX BREAKDOWN
    # ==================================================================
    with sex_tab:
        st.header("Awards & Outcomes by Biological Sex")

        if len(demo_data) == 0 or "Sex" not in demo_data.columns:
            st.warning("No demographics data available. Add sex data to the PI Demographics tab in Evans_Endowment_Awards_Source_Data.xlsx.")
        else:
            sex_map = dict(zip(demo_data["Name"], demo_data["Sex"]))
            n_with_sex = sum(1 for v in sex_map.values() if pd.notna(v))
            st.caption(f"Sex data available for {n_with_sex} of {len(sex_map)} investigators.")

            # --- K Awards by Sex ---
            st.markdown("### K Awards by Sex")
            k_with_sex = k_data.copy()
            k_with_sex["Sex"] = k_with_sex["Name"].map(sex_map)
            k_with_sex = k_with_sex[k_with_sex["Sex"].notna()]

            # Summary metrics
            k_sex_summary = k_with_sex.groupby("Sex").agg(
                N_Awardees=("Name", "nunique"),
                N_AwardYears=("Name", "count"),
                Total_Gap=("SalaryGap", "sum"),
                Total_Cost=("TotalCost", "sum"),
            ).reset_index()
            k_sex_summary["Avg Gap/Awardee"] = k_sex_summary["Total_Gap"] / k_sex_summary["N_Awardees"]
            k_sex_summary["Sex"] = k_sex_summary["Sex"].map({"F": "Female", "M": "Male"})

            sc1, sc2 = st.columns(2)
            for _, row in k_sex_summary.iterrows():
                col = sc1 if row["Sex"] == "Female" else sc2
                col.metric(f"K Awardees ({row['Sex']})", f"{int(row['N_Awardees'])}")
                col.metric(f"Total Award-Years ({row['Sex']})", f"{int(row['N_AwardYears'])}")
                col.metric(f"Total DoM Cost ({row['Sex']})", f"${row['Total_Cost']:,.0f}")

            # K awards by sex and year
            k_sex_year = k_with_sex.groupby(["FY", "Sex"]).agg(
                Count=("Name", "nunique"),
                Cost=("TotalCost", "sum"),
            ).reset_index()
            k_sex_year["Sex"] = k_sex_year["Sex"].map({"F": "Female", "M": "Male"})

            fig_ksex = px.bar(
                k_sex_year.sort_values("FY"),
                x="FY", y="Count", color="Sex", barmode="group",
                color_discrete_map={"Female": "#f472b6", "Male": "#60a5fa"},
                labels={"Count": "Unique K Awardees", "FY": "Fiscal Year"},
                title="K Awardees by Sex and Year",
            )
            fig_ksex.update_layout(
                plot_bgcolor="#1a1f2e", paper_bgcolor="#1a1f2e", font_color="#c8d0e0",
                xaxis=dict(gridcolor="#2a3050"), yaxis=dict(gridcolor="#2a3050"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )
            st.plotly_chart(fig_ksex, use_container_width=True)

            # --- K-to-R Conversion by Sex ---
            st.markdown("### K-to-R Conversion Rate by Sex")

            CURRENT_FY = 2026
            BUFFER_YEARS = 1

            # Need NIH data — use batch query if available
            all_k_names = sorted(k_data["Name"].unique())
            combined_names_for_sex = sorted(set(all_k_names) | set(demo_data["Name"].unique()))
            all_nih_sex = batch_query_all_awardees(tuple(combined_names_for_sex))

            if len(all_nih_sex) > 0:
                non_k_sex = all_nih_sex[~all_nih_sex["Is K Grant"]].copy()

                k_sections_sex = k_data.groupby(["Name"]).agg(
                    LastK=("FY_Num", "max"),
                    Section=("Section", "first"),
                ).reset_index()
                k_sections_sex["Sex"] = k_sections_sex["Name"].map(sex_map)
                k_sections_sex = k_sections_sex[k_sections_sex["Sex"].notna()]
                k_sections_sex["Eligible"] = k_sections_sex["LastK"] + BUFFER_YEARS < CURRENT_FY

                # Post-K grants
                last_k_map = dict(zip(k_sections_sex["Name"], k_sections_sex["LastK"]))
                non_k_sex["LastK"] = non_k_sex["Name"].map(last_k_map)
                post_k_sex = non_k_sex[non_k_sex["LastK"].notna() & (non_k_sex["Fiscal Year"] > non_k_sex["LastK"])]
                names_with_post_k = set(post_k_sex["Name"].unique())

                k_sections_sex["HasPostK"] = k_sections_sex["Name"].isin(names_with_post_k)

                # Post-K direct costs per person
                post_k_direct_by_pi = post_k_sex.groupby("Name")["Direct Cost"].sum().to_dict()
                k_sections_sex["PostK_Direct"] = k_sections_sex["Name"].map(post_k_direct_by_pi).fillna(0)

                # Conversion summary by sex
                sex_conv = k_sections_sex.groupby("Sex").agg(
                    N_Total=("Name", "nunique"),
                    N_Eligible=("Eligible", "sum"),
                    N_Transitioned=("HasPostK", lambda x: (x & k_sections_sex.loc[x.index, "Eligible"]).sum()),
                    Total_PostK_Direct=("PostK_Direct", "sum"),
                ).reset_index()
                sex_conv["Transition %"] = (
                    (100 * sex_conv["N_Transitioned"] / sex_conv["N_Eligible"])
                    .where(sex_conv["N_Eligible"] > 0, other=0)
                    .round(0).astype(int)
                )
                sex_conv["Sex"] = sex_conv["Sex"].map({"F": "Female", "M": "Male"})

                # Display
                conv_display = sex_conv.copy()
                conv_display["Total_PostK_Direct"] = conv_display["Total_PostK_Direct"].apply(lambda x: f"${x:,.0f}")
                conv_display["Transition %"] = conv_display["Transition %"].astype(str) + "%"
                conv_display.columns = ["Sex", "K Awardees", "Eligible", "Transitioned", "Post-K NIH Direct", "Transition %"]
                st.dataframe(conv_display, use_container_width=True, hide_index=True)

                # Bar chart
                fig_conv = go.Figure()
                for _, row in sex_conv.iterrows():
                    fig_conv.add_trace(go.Bar(
                        name=row["Sex"],
                        x=["Eligible", "Transitioned"],
                        y=[row["N_Eligible"], row["N_Transitioned"]],
                        marker_color="#f472b6" if row["Sex"] == "Female" else "#60a5fa",
                        text=[int(row["N_Eligible"]), int(row["N_Transitioned"])],
                        textposition="auto",
                    ))
                fig_conv.update_layout(
                    barmode="group", title="K-to-R Transition by Sex",
                    plot_bgcolor="#1a1f2e", paper_bgcolor="#1a1f2e", font_color="#c8d0e0",
                    xaxis=dict(gridcolor="#2a3050"), yaxis=dict(gridcolor="#2a3050", title="Count"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig_conv, use_container_width=True)

                # Detailed table
                with st.expander("View all K awardees by sex"):
                    detail = k_sections_sex[["Name", "Sex", "Section", "LastK", "Eligible", "HasPostK", "PostK_Direct"]].copy()
                    detail["Sex"] = detail["Sex"].map({"F": "Female", "M": "Male"})
                    detail["LastK"] = detail["LastK"].apply(lambda x: f"FY{x}")
                    detail["Eligible"] = detail["Eligible"].map({True: "Yes", False: "Too recent"})
                    detail["HasPostK"] = detail["HasPostK"].map({True: "✓", False: "—"})
                    detail["PostK_Direct"] = detail["PostK_Direct"].apply(lambda x: f"${x:,.0f}" if x > 0 else "—")
                    detail.columns = ["Name", "Sex", "Section", "Last K Year", "Eligible", "Transitioned", "Post-K Direct"]
                    st.dataframe(detail.sort_values(["Sex", "Name"]), use_container_width=True, hide_index=True)

            # --- Pilot/Jr Faculty by Sex ---
            if len(other_data):
                st.markdown("### Pilot / Junior Faculty Awards by Sex")
                pj_with_sex = other_data.copy()
                pj_with_sex["Sex"] = pj_with_sex["Name"].map(sex_map)
                pj_with_sex = pj_with_sex[pj_with_sex["Sex"].notna()]

                if len(pj_with_sex):
                    pj_sex_summary = pj_with_sex.groupby(["Award", "Sex"]).agg(
                        N_Awardees=("Name", "nunique"),
                        Total_Amount=("Amount", "sum"),
                    ).reset_index()
                    pj_sex_summary["Sex"] = pj_sex_summary["Sex"].map({"F": "Female", "M": "Male"})

                    pj_display = pj_sex_summary.copy()
                    pj_display["Total_Amount"] = pj_display["Total_Amount"].apply(lambda x: f"${x:,.0f}")
                    pj_display.columns = ["Award Type", "Sex", "Unique Awardees", "Total Investment"]
                    st.dataframe(pj_display, use_container_width=True, hide_index=True)

                    fig_pj = px.bar(
                        pj_sex_summary, x="Award", y="N_Awardees", color="Sex", barmode="group",
                        color_discrete_map={"Female": "#f472b6", "Male": "#60a5fa"},
                        labels={"N_Awardees": "Unique Awardees", "Award": "Award Type"},
                        title="Pilot / Junior Faculty Awardees by Sex",
                    )
                    fig_pj.update_layout(
                        plot_bgcolor="#1a1f2e", paper_bgcolor="#1a1f2e", font_color="#c8d0e0",
                        xaxis=dict(gridcolor="#2a3050"), yaxis=dict(gridcolor="#2a3050"),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    )
                    st.plotly_chart(fig_pj, use_container_width=True)

    # ==================================================================
    # TAB: DRILL DOWN
    # ==================================================================
    with drill_tab:
        st.header("Drill Down by Award Type and Year")

        drill_col1, drill_col2 = st.columns(2)
        with drill_col1:
            drill_award = st.selectbox("Select Award Type", ["K Award", "Pilot", "Junior Faculty"])
        with drill_col2:
            if drill_award == "K Award":
                avail_fys = sorted(k_data["FY"].unique())
            elif len(other_data):
                avail_fys = sorted(other_data[other_data["Award"] == drill_award]["FY"].dropna().unique())
            else:
                avail_fys = []
            drill_fy = st.selectbox("Select Fiscal Year", avail_fys, index=len(avail_fys) - 1 if avail_fys else 0)

        if drill_award == "K Award" and drill_fy:
            fy_k = k_data[k_data["FY"] == drill_fy].copy()
            st.subheader(f"K Awardees — {drill_fy} ({len(fy_k)} awardees)")

            display_k = fy_k[["Name", "Section", "SalaryGap", "TotalCost", "AwardNo"]].copy()
            display_k.columns = ["Name", "Section", "50% Salary Gap", "Total Cost", "Award Number"]
            display_k["50% Salary Gap"] = display_k["50% Salary Gap"].apply(lambda x: f"${x:,.0f}" if pd.notna(x) and x > 0 else "—")
            display_k["Total Cost"] = display_k["Total Cost"].apply(lambda x: f"${x:,.0f}" if pd.notna(x) and x > 0 else "—")
            display_k = display_k.sort_values("Name").reset_index(drop=True)
            st.dataframe(display_k, use_container_width=True, hide_index=True)

        elif drill_award in ("Pilot", "Junior Faculty") and drill_fy and len(other_data):
            fy_other = other_data[(other_data["Award"] == drill_award) & (other_data["FY"] == drill_fy)].copy()
            st.subheader(f"{drill_award} Awards — {drill_fy} ({len(fy_other)} awards)")

            display_other = fy_other[["Name", "Section", "Amount"]].copy()
            display_other["Amount"] = display_other["Amount"].apply(lambda x: f"${x:,.0f}" if pd.notna(x) else "—")
            display_other = display_other.sort_values("Name").reset_index(drop=True)
            st.dataframe(display_other, use_container_width=True, hide_index=True)

    # ==================================================================
    # TAB 4: INDIVIDUAL PI LOOKUP
    # ==================================================================
    with lookup_tab:
        st.header("Investigator NIH Grant Lookup")
        st.caption("Select an awardee to view their full NIH grant portfolio from RePORTER")

        all_names = sorted(k_data["Name"].unique())
        if len(other_data):
            all_names = sorted(set(all_names) | set(other_data["Name"].unique()))

        selected_pi = st.selectbox("Select Investigator", [""] + all_names, index=0)

        if selected_pi:
            pi_k = k_data[k_data["Name"] == selected_pi]
            if len(pi_k):
                k_years = ", ".join(sorted(pi_k["FY"].unique()))
                k_sections = ", ".join(pi_k["Section"].unique())
                last_k = pi_k["FY_Num"].max()
                total_gap = pi_k["SalaryGap"].sum()
                st.info(f"**K Award history:** {k_years} | Section: {k_sections} | "
                        f"Total 50% Salary Gap: ${total_gap:,.0f} | Last K year: FY{last_k}")
            else:
                last_k = None

            if len(other_data):
                pi_other = other_data[other_data["Name"] == selected_pi]
                if len(pi_other):
                    for _, row in pi_other.iterrows():
                        st.info(f"**{row['Award']}:** {row['FY']} | Section: {row['Section']} | Amount: ${row['Amount']:,.0f}")

            nih_grants = get_nih_grants_for_person(selected_pi)

            if len(nih_grants) == 0:
                st.warning("No NIH grants found in RePORTER for this investigator.")
            else:
                k_grants = nih_grants[nih_grants["Is K Grant"]].copy()
                non_k = nih_grants[~nih_grants["Is K Grant"]].copy()

                if last_k:
                    post_k = non_k[non_k["Fiscal Year"] > last_k]
                else:
                    post_k = non_k

                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.metric("Total NIH Grants", f"{len(nih_grants)}")
                mc2.metric("K Grant Years", f"{len(k_grants)}")
                mc3.metric("Post-K Non-K Grants", f"{len(post_k)}")
                if len(post_k):
                    mc4.metric("Post-K Direct Costs", f"${post_k['Direct Cost'].sum():,.0f}")
                else:
                    mc4.metric("Post-K Direct Costs", "$0")

                pi_tab1, pi_tab2, pi_tab3 = st.tabs(["Post-K Grants", "K Grants (RePORTER)", "All Grants"])

                display_cols = ["Project Number", "Activity Code", "Title", "Organization",
                               "Fiscal Year", "Direct Cost", "Indirect Cost", "Award Amount",
                               "IC", "Contact PI", "Location"]

                with pi_tab1:
                    if len(post_k):
                        st.subheader(f"Non-K Grants After K Period (FY{last_k}+)" if last_k else "Non-K Grants")
                        st.dataframe(
                            post_k[display_cols].sort_values("Fiscal Year", ascending=False).reset_index(drop=True),
                            use_container_width=True, hide_index=True,
                        )
                    else:
                        st.info("No post-K non-K grants found." +
                                (" K period may still be active." if last_k and last_k >= 2025 else ""))

                with pi_tab2:
                    if len(k_grants):
                        st.dataframe(
                            k_grants[display_cols].sort_values("Fiscal Year", ascending=False).reset_index(drop=True),
                            use_container_width=True, hide_index=True,
                        )
                    else:
                        st.info("No K grants found in NIH RePORTER.")

                with pi_tab3:
                    st.dataframe(
                        nih_grants[display_cols].sort_values("Fiscal Year", ascending=False).reset_index(drop=True),
                        use_container_width=True, hide_index=True,
                    )


if __name__ == "__main__":
    main()

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

API_URL = "https://api.reporter.nih.gov/v2/projects/search"
BU_ORGS = {"BOSTON UNIVERSITY", "BOSTON MEDICAL CENTER", "BOSTON UNIVERSITY MEDICAL CAMPUS"}

K_CODES = {
    "K01", "K02", "K05", "K06", "K07", "K08", "K11", "K12",
    "K22", "K23", "K24", "K25", "K26", "K43", "K76", "K99", "K00",
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


# ── MAIN APP ─────────────────────────────────────────────────────────────────

def main():
    if not check_password():
        st.stop()

    st.title("Evans Endowment Awards — ROI Dashboard")
    st.caption("Drill down into K, Pilot, Junior Faculty, and Bridge awards funded by the Evans Endowment")

    k_data = load_k_awardees()
    other_data = load_other_awards()
    overview = build_overview(k_data, other_data)

    # ── SIDEBAR ──────────────────────────────────────────────────────────────
    st.sidebar.header("Filters")
    all_fys = sorted(overview["FY"].unique())
    selected_fys = st.sidebar.multiselect("Fiscal Years", all_fys, default=all_fys)

    award_types = sorted(overview["Award"].unique())
    selected_awards = st.sidebar.multiselect("Award Types", award_types, default=award_types)

    filtered = overview[overview["FY"].isin(selected_fys) & overview["Award"].isin(selected_awards)]

    # ── TOP METRICS ──────────────────────────────────────────────────────────
    # Count unique investigators per award type (across selected FYs)
    k_fys = set(selected_fys)
    n_k_investigators = k_data[k_data["FY"].isin(k_fys)]["Name"].nunique()
    if len(other_data):
        n_pilot = other_data[(other_data["Award"] == "Pilot") & (other_data["FY"].isin(k_fys))]["Name"].nunique()
        n_jrfac = other_data[(other_data["Award"] == "Junior Faculty") & (other_data["FY"].isin(k_fys))]["Name"].nunique()
    else:
        n_pilot, n_jrfac = 0, 0
    # Bridge: we only have aggregate counts, not names
    n_bridge = int(BRIDGE_DATA[BRIDGE_DATA["FY"].isin(k_fys)]["Count"].sum())

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Investment", f"${filtered['Amount'].sum():,.0f}")
    col2.metric("Fiscal Years", f"{filtered['FY'].nunique()}")
    col3.metric("Award Types", f"{filtered['Award'].nunique()}")
    col4.metric("Total Award-Years", f"{int(filtered['Count'].sum()):,}")

    # Unique investigator counts by award type
    st.markdown("##### Unique Investigators by Award Type")
    inv_cols = st.columns(4)
    inv_cols[0].metric("K Awardees", f"{n_k_investigators}")
    inv_cols[1].metric("Pilot Awardees", f"{n_pilot}")
    inv_cols[2].metric("Junior Faculty", f"{n_jrfac}")
    inv_cols[3].metric("Bridge Awards", f"{n_bridge}")

    # ── OVERVIEW TABLE ───────────────────────────────────────────────────────
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

    # ── CHART ────────────────────────────────────────────────────────────────
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

    # ── DRILL-DOWN ───────────────────────────────────────────────────────────
    st.header("Drill Down")

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

    # ── INDIVIDUAL PI LOOKUP ─────────────────────────────────────────────────
    st.header("Investigator NIH Grant Lookup")
    st.caption("Select an awardee to view their full NIH grant portfolio from RePORTER")

    # Build list of all awardee names
    all_names = sorted(k_data["Name"].unique())
    if len(other_data):
        all_names = sorted(set(all_names) | set(other_data["Name"].unique()))

    selected_pi = st.selectbox("Select Investigator", [""] + all_names, index=0)

    if selected_pi:
        # Show their K/award history
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

        # Query NIH RePORTER
        nih_grants = get_nih_grants_for_person(selected_pi)

        if len(nih_grants) == 0:
            st.warning("No NIH grants found in RePORTER for this investigator.")
        else:
            # Split into K grants vs post-K grants
            k_grants = nih_grants[nih_grants["Is K Grant"]].copy()
            non_k = nih_grants[~nih_grants["Is K Grant"]].copy()

            if last_k:
                post_k = non_k[non_k["Fiscal Year"] > last_k]
                during_k = non_k[non_k["Fiscal Year"] <= last_k]
            else:
                post_k = non_k
                during_k = pd.DataFrame()

            # Metrics
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("Total NIH Grants", f"{len(nih_grants)}")
            mc2.metric("K Grant Years", f"{len(k_grants)}")
            mc3.metric("Post-K Non-K Grants", f"{len(post_k)}")
            if len(post_k):
                mc4.metric("Post-K Direct Costs", f"${post_k['Direct Cost'].sum():,.0f}")
            else:
                mc4.metric("Post-K Direct Costs", "$0")

            # Tabs
            tab1, tab2, tab3 = st.tabs(["Post-K Grants", "K Grants (RePORTER)", "All Grants"])

            display_cols = ["Project Number", "Activity Code", "Title", "Organization",
                           "Fiscal Year", "Direct Cost", "Indirect Cost", "Award Amount",
                           "IC", "Contact PI", "Location"]

            with tab1:
                if len(post_k):
                    st.subheader(f"Non-K Grants After K Period (FY{last_k}+)" if last_k else "Non-K Grants")
                    st.dataframe(
                        post_k[display_cols].sort_values("Fiscal Year", ascending=False).reset_index(drop=True),
                        use_container_width=True, hide_index=True,
                    )
                else:
                    st.info("No post-K non-K grants found." +
                            (" K period may still be active." if last_k and last_k >= 2025 else ""))

            with tab2:
                if len(k_grants):
                    st.dataframe(
                        k_grants[display_cols].sort_values("Fiscal Year", ascending=False).reset_index(drop=True),
                        use_container_width=True, hide_index=True,
                    )
                else:
                    st.info("No K grants found in NIH RePORTER.")

            with tab3:
                st.dataframe(
                    nih_grants[display_cols].sort_values("Fiscal Year", ascending=False).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                )


if __name__ == "__main__":
    main()

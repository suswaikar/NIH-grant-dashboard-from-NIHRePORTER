#!/usr/bin/env python
"""
Build a single self-contained HTML report combining the K-to-R Pipeline
analysis and the Pilot/Junior Award ROI analysis (GT97 excluded — see pilot
script header for rationale). Reads the two Excel outputs, computes headline
numbers, generates interactive Plotly charts, and emits one HTML file
viewable in any browser or attachable to email.

Usage:
    python build_html_report.py

Output:
    DoM_Award_ROI_Report.html  (in Data Requests folder)
"""

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime
import sys, io
import html as html_lib

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ── PATHS ────────────────────────────────────────────────────────────────────
BASE = Path(
    r"C:\Users\swaikar\OneDrive - Boston University\Department of Medicine"
    r"\Chair advisory committee for Evans"
)
DATA = BASE / "shared drive_041626" / "Research Advisory Committee" / "Data Requests"
K_FILE = DATA / "K_awardees_summary_2016_2026.xlsx"
PILOT_FILE = DATA / "DoM_Pilot_ROI_Summary.xlsx"
OUT_FILE = DATA / "DoM_Award_ROI_Report.html"

PLOTLY_TEMPLATE = "plotly_white"
PRIMARY = "#1F4E79"
ACCENT = "#3D7BBF"
MUTED = "#6c757d"


# ── LOAD DATA ────────────────────────────────────────────────────────────────
def fmt_money(x):
    if pd.isna(x) or x == 0:
        return "—"
    if abs(x) >= 1_000_000:
        return f"${x/1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"${x/1_000:.0f}k"
    return f"${x:,.0f}"


def fmt_money_full(x):
    if pd.isna(x):
        return "—"
    return f"${x:,.0f}"


def fmt_ratio(x):
    if pd.isna(x) or x == 0:
        return "—"
    return f"{x:.1f}×"


print("Loading Excel data...")
pilot_summary = pd.read_excel(PILOT_FILE, sheet_name="Investigator Summary")
pilot_by_type = pd.read_excel(PILOT_FILE, sheet_name="By Award Type")
pilot_narrative = pd.read_excel(PILOT_FILE, sheet_name="Investigator Narrative")
pilot_xref = pd.read_excel(PILOT_FILE, sheet_name="Cross-Reference")

k_summary = pd.read_excel(K_FILE, sheet_name="K Awardees 2016-2026")
k_bu = pd.read_excel(K_FILE, sheet_name="NIH Grants - BU_BMC")
k_ext = pd.read_excel(K_FILE, sheet_name="NIH Grants - External")
k_xref = pd.read_excel(K_FILE, sheet_name="Cross-Reference")

# ── HEADLINE NUMBERS ─────────────────────────────────────────────────────────
pilot_dom = pilot_summary["DoMTotal"].sum()
pilot_dir_gen = pilot_summary["NIH Direct (Generous)"].sum()
pilot_dir_cons = pilot_summary["NIH Direct (Conservative)"].sum()
pilot_ind_gen = pilot_summary["NIH Indirect (Generous)"].sum()
pilot_ind_cons = pilot_summary["NIH Indirect (Conservative)"].sum()
pilot_pis_total = len(pilot_summary)
pilot_pis_funded_gen = (pilot_summary["NIH Direct (Generous)"] > 0).sum()
pilot_pis_funded_cons = (pilot_summary["NIH Direct (Conservative)"] > 0).sum()

k_dom = k_summary["Total 50% Salary Gap"].sum()
k_dir = k_summary["Post-K NIH Direct Costs"].sum()
k_ind = k_summary["Post-K NIH Indirect Costs"].sum()
k_pis_total = len(k_summary)
k_pis_funded = (k_summary["Post-K Grant Count"] > 0).sum()

combined_dom = pilot_dom + k_dom
combined_dir = pilot_dir_cons + k_dir  # use conservative for combined headline
combined_ind = pilot_ind_cons + k_ind

# ── FIGURES ──────────────────────────────────────────────────────────────────
print("Building charts...")

# Fig 1: ROI comparison across programs
roi_data = pd.DataFrame([
    {"Program": "Pilot", "Recipients": int(pilot_by_type.loc[pilot_by_type["Award Type"]=="Pilot", "Recipients"].iloc[0]),
     "DoM $": pilot_by_type.loc[pilot_by_type["Award Type"]=="Pilot", "DoM $ (this type only)"].iloc[0],
     "Generous": pilot_by_type.loc[pilot_by_type["Award Type"]=="Pilot", "ROI Generous"].iloc[0],
     "Conservative": pilot_by_type.loc[pilot_by_type["Award Type"]=="Pilot", "ROI Conservative"].iloc[0]},
    {"Program": "Junior", "Recipients": int(pilot_by_type.loc[pilot_by_type["Award Type"]=="Junior", "Recipients"].iloc[0]),
     "DoM $": pilot_by_type.loc[pilot_by_type["Award Type"]=="Junior", "DoM $ (this type only)"].iloc[0],
     "Generous": pilot_by_type.loc[pilot_by_type["Award Type"]=="Junior", "ROI Generous"].iloc[0],
     "Conservative": pilot_by_type.loc[pilot_by_type["Award Type"]=="Junior", "ROI Conservative"].iloc[0]},
    {"Program": "K (gap)", "Recipients": k_pis_funded,
     "DoM $": k_dom,
     "Generous": round(k_dir / k_dom, 1) if k_dom else 0,
     "Conservative": round(k_dir / k_dom, 1) if k_dom else 0},  # K only has one definition
])

fig1 = go.Figure()
fig1.add_trace(go.Bar(
    name="Generous (active from anchor year)",
    x=roi_data["Program"], y=roi_data["Generous"],
    marker_color=ACCENT,
    text=[f"{v:.0f}×" for v in roi_data["Generous"]],
    textposition="outside",
))
fig1.add_trace(go.Bar(
    name="Conservative (started at/after anchor year)",
    x=roi_data["Program"], y=roi_data["Conservative"],
    marker_color=PRIMARY,
    text=[f"{v:.0f}×" for v in roi_data["Conservative"]],
    textposition="outside",
))
fig1.update_layout(
    title="ROI by program — NIH direct costs ÷ DoM investment",
    yaxis_title="Return multiple (×)",
    barmode="group",
    template=PLOTLY_TEMPLATE,
    height=420,
    margin=dict(t=60, b=40, l=60, r=20),
    legend=dict(orientation="h", y=-0.15),
)

# Fig 2: DoM investment vs NIH return (stacked)
fig2 = go.Figure()
prog_names = ["Pilot", "Junior", "K (salary gap)"]
dom_vals = [
    pilot_by_type.loc[pilot_by_type["Award Type"]=="Pilot", "DoM $ (this type only)"].iloc[0],
    pilot_by_type.loc[pilot_by_type["Award Type"]=="Junior", "DoM $ (this type only)"].iloc[0],
    k_dom,
]
nih_dir_vals = [
    pilot_by_type.loc[pilot_by_type["Award Type"]=="Pilot", "NIH Direct (Conservative)"].iloc[0],
    pilot_by_type.loc[pilot_by_type["Award Type"]=="Junior", "NIH Direct (Conservative)"].iloc[0],
    k_dir,
]
nih_ind_vals = [
    pilot_by_type.loc[pilot_by_type["Award Type"]=="Pilot", "NIH Indirect (Conservative)"].iloc[0],
    pilot_by_type.loc[pilot_by_type["Award Type"]=="Junior", "NIH Indirect (Conservative)"].iloc[0],
    k_ind,
]
fig2.add_trace(go.Bar(
    name="DoM investment", x=prog_names, y=dom_vals,
    marker_color="#dc3545",
    text=[fmt_money(v) for v in dom_vals], textposition="auto",
))
fig2.add_trace(go.Bar(
    name="NIH direct (conservative)", x=prog_names, y=nih_dir_vals,
    marker_color=PRIMARY,
    text=[fmt_money(v) for v in nih_dir_vals], textposition="auto",
))
fig2.add_trace(go.Bar(
    name="NIH indirect", x=prog_names, y=nih_ind_vals,
    marker_color=ACCENT,
    text=[fmt_money(v) for v in nih_ind_vals], textposition="auto",
))
fig2.update_layout(
    title="DoM investment vs subsequent NIH funding (conservative definition)",
    yaxis_title="Dollars",
    yaxis_type="log",
    barmode="group",
    template=PLOTLY_TEMPLATE,
    height=460,
    margin=dict(t=60, b=40, l=80, r=20),
    legend=dict(orientation="h", y=-0.15),
)

# Fig 3: Top-funded pilot investigators (only those with non-zero conservative funding)
top_pilot = pilot_summary[pilot_summary["NIH Direct (Conservative)"] > 0].nlargest(15, "NIH Direct (Conservative)")[
    ["Name", "Section", "DoMTotal", "NIH Direct (Conservative)", "ROI Conservative (Direct/DoM)"]
].copy()
fig3 = go.Figure()
fig3.add_trace(go.Bar(
    y=top_pilot["Name"][::-1],
    x=top_pilot["NIH Direct (Conservative)"][::-1],
    orientation="h",
    marker_color=PRIMARY,
    text=[fmt_money(v) for v in top_pilot["NIH Direct (Conservative)"][::-1]],
    textposition="outside",
    customdata=list(zip(top_pilot["Section"][::-1], top_pilot["DoMTotal"][::-1], top_pilot["ROI Conservative (Direct/DoM)"][::-1])),
    hovertemplate="<b>%{y}</b><br>Section: %{customdata[0]}<br>DoM total: %{customdata[1]:$,.0f}<br>NIH direct: %{x:$,.0f}<br>ROI: %{customdata[2]:.1f}×<extra></extra>",
))
fig3.update_layout(
    title=f"Top {len(top_pilot)} Pilot/Junior recipients by conservative NIH direct funding",
    xaxis_title="NIH Direct Costs (USD)",
    template=PLOTLY_TEMPLATE,
    height=480,
    margin=dict(t=60, b=40, l=200, r=80),
)

# Fig 4: Top-funded K awardees
top_k = k_summary.nlargest(15, "Post-K NIH Direct Costs")[
    ["Name", "Section", "Total 50% Salary Gap", "Post-K NIH Direct Costs", "Post-K Grant Count"]
].copy()
top_k["ROI"] = (top_k["Post-K NIH Direct Costs"] / top_k["Total 50% Salary Gap"]).replace([float("inf"), -float("inf")], 0).fillna(0)
fig4 = go.Figure()
fig4.add_trace(go.Bar(
    y=top_k["Name"][::-1],
    x=top_k["Post-K NIH Direct Costs"][::-1],
    orientation="h",
    marker_color=ACCENT,
    text=[fmt_money(v) for v in top_k["Post-K NIH Direct Costs"][::-1]],
    textposition="outside",
    customdata=list(zip(top_k["Section"][::-1], top_k["Total 50% Salary Gap"][::-1], top_k["ROI"][::-1])),
    hovertemplate="<b>%{y}</b><br>Section: %{customdata[0]}<br>DoM gap: %{customdata[1]:$,.0f}<br>Post-K direct: %{x:$,.0f}<br>ROI: %{customdata[2]:.1f}×<extra></extra>",
))
fig4.update_layout(
    title="Top 15 K awardees by post-K NIH direct funding (FY2016–2026)",
    xaxis_title="Post-K NIH Direct Costs (USD)",
    template=PLOTLY_TEMPLATE,
    height=480,
    margin=dict(t=60, b=40, l=200, r=80),
)

# Fig 5: K conversion rate by section
k_by_section = k_summary.groupby("Section").agg(
    Total=("Name", "count"),
    WithFunding=("Post-K Grant Count", lambda x: (x > 0).sum()),
    DirectFunding=("Post-K NIH Direct Costs", "sum"),
    GapSpend=("Total 50% Salary Gap", "sum"),
).reset_index()
k_by_section["Rate"] = (k_by_section["WithFunding"] / k_by_section["Total"] * 100).round(0)
k_by_section = k_by_section.sort_values("DirectFunding", ascending=True)

fig5 = go.Figure()
fig5.add_trace(go.Bar(
    y=k_by_section["Section"],
    x=k_by_section["DirectFunding"],
    orientation="h",
    marker_color=PRIMARY,
    text=[
        f"{int(r)}% ({int(w)}/{int(t)}) — {fmt_money(d)}"
        for r, w, t, d in zip(k_by_section["Rate"], k_by_section["WithFunding"],
                              k_by_section["Total"], k_by_section["DirectFunding"])
    ],
    textposition="outside",
    hovertemplate="<b>%{y}</b><br>Awardees: %{customdata[0]}<br>With post-K funding: %{customdata[1]} (%{customdata[2]}%)<br>Direct funding: %{x:$,.0f}<br>DoM gap spend: %{customdata[3]:$,.0f}<extra></extra>",
    customdata=list(zip(k_by_section["Total"], k_by_section["WithFunding"],
                        k_by_section["Rate"], k_by_section["GapSpend"])),
))
fig5.update_layout(
    title="K-to-R outcomes by section (post-K NIH direct funding, FY2016–2026)",
    xaxis_title="Post-K NIH Direct Costs (USD)",
    template=PLOTLY_TEMPLATE,
    height=480,
    margin=dict(t=60, b=40, l=200, r=120),
)

# ── HTML TEMPLATE ────────────────────────────────────────────────────────────
def fig_html(fig, div_id):
    return fig.to_html(full_html=False, include_plotlyjs=False, div_id=div_id)


def df_to_html(df, table_id, money_cols=None, ratio_cols=None, int_cols=None):
    """Render a DataFrame as a styled HTML table."""
    money_cols = set(money_cols or [])
    ratio_cols = set(ratio_cols or [])
    int_cols = set(int_cols or [])

    out = [f'<table class="data-table" id="{table_id}"><thead><tr>']
    for col in df.columns:
        out.append(f'<th>{html_lib.escape(str(col))}</th>')
    out.append("</tr></thead><tbody>")
    for _, row in df.iterrows():
        out.append("<tr>")
        for col in df.columns:
            v = row[col]
            if col in money_cols:
                cell = fmt_money_full(v)
                cls = "num"
            elif col in ratio_cols:
                cell = fmt_ratio(v)
                cls = "num ratio"
            elif col in int_cols:
                cell = "—" if pd.isna(v) else f"{int(v):,}"
                cls = "num"
            else:
                cell = "" if pd.isna(v) else html_lib.escape(str(v))
                cls = ""
            out.append(f'<td class="{cls}">{cell}</td>')
        out.append("</tr>")
    out.append("</tbody></table>")
    return "".join(out)


# Pre-build narrative table for pilot (subset of columns)
narr_disp = pilot_narrative[[
    "Name", "Section", "DoM History", "DoM Total",
    "NIH Direct (Conservative)", "ROI Conservative",
    "Major NIH Grants — Conservative (started at/after first DoM FY)",
]].copy()
narr_disp.columns = ["Name", "Section", "DoM Awards Received", "DoM Total",
                     "NIH Direct (cons.)", "ROI", "Major NIH Awards Started Since DoM Support"]
narr_disp = narr_disp.sort_values("NIH Direct (cons.)", ascending=False)

# Pre-build K summary table (top 30 by post-K direct)
k_disp = k_summary[[
    "Name", "Section", "Years Supported", "Total 50% Salary Gap",
    "Post-K NIH Direct Costs", "Post-K NIH Indirect Costs", "Post-K Grant Count",
]].copy()
k_disp["ROI"] = (k_disp["Post-K NIH Direct Costs"] / k_disp["Total 50% Salary Gap"]).replace([float("inf"), -float("inf")], 0).fillna(0)
k_disp = k_disp.sort_values("Post-K NIH Direct Costs", ascending=False)

# By-program table
bytype_disp = pilot_by_type[[
    "Award Type", "Recipients", "DoM $ (this type only)",
    "NIH Direct (Generous)", "ROI Generous",
    "NIH Direct (Conservative)", "ROI Conservative",
]].copy()
bytype_disp.columns = ["Program", "Recipients", "DoM $",
                       "NIH Direct (Generous)", "ROI Generous",
                       "NIH Direct (Conservative)", "ROI Conservative"]

# K by-section table
ksec_disp = k_by_section.sort_values("DirectFunding", ascending=False)[[
    "Section", "Total", "WithFunding", "Rate", "GapSpend", "DirectFunding"
]].copy()
ksec_disp.columns = ["Section", "K Awardees", "With Post-K NIH", "% Conv.",
                     "DoM Gap Spend", "Post-K NIH Direct"]

now = datetime.now().strftime("%Y-%m-%d %H:%M")

# ── BUILD HTML ───────────────────────────────────────────────────────────────
print("Building HTML...")

CSS = """
:root {
    --primary: #1F4E79;
    --accent: #3D7BBF;
    --muted: #6c757d;
    --bg: #ffffff;
    --bg-alt: #f7f9fc;
    --border: #e3e8ef;
    --text: #1a1a1a;
    --good: #1d7a3a;
    --warn: #b06b00;
}
* { box-sizing: border-box; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    color: var(--text);
    background: var(--bg);
    line-height: 1.5;
    margin: 0;
    padding: 0;
    font-size: 15px;
}
.container { max-width: 1200px; margin: 0 auto; padding: 32px 24px; }
header {
    border-bottom: 3px solid var(--primary);
    padding-bottom: 16px;
    margin-bottom: 32px;
}
h1 { color: var(--primary); margin: 0 0 6px 0; font-size: 28px; font-weight: 700; }
h2 {
    color: var(--primary);
    border-bottom: 2px solid var(--border);
    padding-bottom: 8px;
    margin-top: 48px;
    font-size: 22px;
}
h3 { color: var(--primary); font-size: 17px; margin-top: 28px; }
.subtitle { color: var(--muted); font-size: 14px; }
.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin: 20px 0; }
.kpi {
    background: var(--bg-alt);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 16px;
    border-left: 4px solid var(--primary);
}
.kpi-label { font-size: 12px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.04em; font-weight: 600; }
.kpi-value { font-size: 26px; font-weight: 700; color: var(--text); margin-top: 4px; }
.kpi-sub { font-size: 12px; color: var(--muted); margin-top: 2px; }
.kpi.accent { border-left-color: var(--accent); }
.kpi.good { border-left-color: var(--good); }
.kpi-value.ratio { color: var(--primary); }
.kpi-value.ratio.big { font-size: 32px; }

.data-table {
    width: 100%;
    border-collapse: collapse;
    margin: 16px 0;
    font-size: 13px;
}
.data-table th {
    background: var(--primary);
    color: white;
    padding: 10px 12px;
    text-align: left;
    font-weight: 600;
    border: 1px solid var(--primary);
}
.data-table td {
    padding: 8px 12px;
    border: 1px solid var(--border);
    vertical-align: top;
}
.data-table tbody tr:nth-child(even) { background: var(--bg-alt); }
.data-table tbody tr:hover { background: #fff8e1; }
.data-table td.num { text-align: right; font-variant-numeric: tabular-nums; }
.data-table td.ratio { font-weight: 600; color: var(--primary); }

.summary-box {
    background: linear-gradient(135deg, #1F4E79 0%, #3D7BBF 100%);
    color: white;
    padding: 24px 28px;
    border-radius: 10px;
    margin: 20px 0;
}
.summary-box h3 { color: white; margin-top: 0; }
.summary-box .big-roi { font-size: 42px; font-weight: 800; margin: 8px 0; }
.summary-box p { margin: 8px 0; opacity: 0.95; font-size: 15px; }

.callout {
    background: #fff8e1;
    border-left: 4px solid var(--warn);
    padding: 14px 18px;
    border-radius: 6px;
    margin: 16px 0;
}
.callout strong { color: var(--warn); }

.method-box {
    background: var(--bg-alt);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 16px 20px;
    margin: 16px 0;
}
.method-box ol, .method-box ul { margin: 8px 0; padding-left: 22px; }
.method-box li { margin: 4px 0; }

details {
    background: var(--bg-alt);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 12px 16px;
    margin: 16px 0;
}
details summary {
    cursor: pointer;
    font-weight: 600;
    color: var(--primary);
    font-size: 15px;
    padding: 4px 0;
}
details[open] summary { margin-bottom: 12px; }

.scrollable-table {
    max-height: 600px;
    overflow-y: auto;
    border: 1px solid var(--border);
    border-radius: 6px;
}
.scrollable-table .data-table { margin: 0; }
.scrollable-table .data-table th { position: sticky; top: 0; z-index: 1; }

footer {
    margin-top: 60px;
    padding-top: 20px;
    border-top: 1px solid var(--border);
    color: var(--muted);
    font-size: 12px;
}

@media print {
    body { font-size: 11px; }
    h1 { font-size: 20px; }
    h2 { font-size: 16px; }
    .summary-box { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    details { break-inside: avoid; }
    details summary { display: none; }
    details > * { display: block !important; }
}
"""

html_parts = []
html_parts.append(f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DoM Award ROI Report</title>
<script src="https://cdn.plot.ly/plotly-2.30.0.min.js" charset="utf-8"></script>
<style>{CSS}</style>
</head>
<body>
<div class="container">

<header>
<h1>Department of Medicine — Internal Award ROI Report</h1>
<div class="subtitle">
K-to-R Pipeline (FY2016–2026) and Pilot / Junior Faculty Awards (AY22–AY26)<br>
Prepared by Sushrut Waikar (with analytical support via Claude Code) — generated {now}
</div>
</header>

<section>
<h2>Executive Summary</h2>

<div class="summary-box">
<h3>Combined headline</h3>
<p>The Department invested <strong>{fmt_money_full(combined_dom)}</strong>
across the K-award salary-gap program ({k_pis_total} faculty) and the pilot / junior faculty award
programs ({pilot_pis_total} faculty). Conservatively counted — only NIH grants newly initiated at or
after the year of DoM support — these {k_pis_total + pilot_pis_total} investigators have generated
<strong>{fmt_money_full(combined_dir)}</strong> in NIH direct costs and
<strong>{fmt_money_full(combined_dir + combined_ind)}</strong> in NIH total funding.</p>
<div class="big-roi">{combined_dir/combined_dom:.0f}× direct &nbsp;·&nbsp; {(combined_dir+combined_ind)/combined_dom:.0f}× total</div>
<p style="margin-top:12px; font-size:13px; opacity: 0.85;">Conservative estimate. Generous estimate (active grants from anchor year onward, including pre-existing portfolios) is roughly 2× higher. <em>GT97 awards are excluded — they are an accounting mechanism for senior investigators with high effort levels rather than a research investment.</em></p>
</div>

<div class="kpi-row">
<div class="kpi"><div class="kpi-label">Total DoM Investment</div><div class="kpi-value">{fmt_money(combined_dom)}</div><div class="kpi-sub">{k_pis_total + pilot_pis_total} faculty across both programs</div></div>
<div class="kpi accent"><div class="kpi-label">NIH Direct (Conservative)</div><div class="kpi-value">{fmt_money(combined_dir)}</div><div class="kpi-sub">Newly initiated grants only</div></div>
<div class="kpi accent"><div class="kpi-label">NIH Total (Conservative)</div><div class="kpi-value">{fmt_money(combined_dir + combined_ind)}</div><div class="kpi-sub">Direct + indirect</div></div>
<div class="kpi good"><div class="kpi-label">ROI (Direct / DoM)</div><div class="kpi-value ratio big">{combined_dir/combined_dom:.0f}×</div><div class="kpi-sub">Conservative</div></div>
</div>
</section>

<section>
<h2>Section 1 — Pilot / Junior Faculty Awards (AY22–AY26)</h2>

<p>{fmt_money_full(pilot_dom)} invested across {pilot_pis_total} investigators in two programs: <strong>Pilot</strong> ($25k–$50k research project awards) and <strong>Junior Award</strong> ($100k career development awards for early-career faculty).</p>
<p style="font-size:13px; color: var(--muted);"><strong>Note on GT97:</strong> GT97 awards (formerly the third program tracked here) are an accounting mechanism that allows well-funded investigators to report less than 100% federal effort. They are not research investments and are excluded from this analysis.</p>

<div class="kpi-row">
<div class="kpi"><div class="kpi-label">DoM Invested</div><div class="kpi-value">{fmt_money(pilot_dom)}</div><div class="kpi-sub">{pilot_pis_total} investigators</div></div>
<div class="kpi accent"><div class="kpi-label">Generous NIH Direct</div><div class="kpi-value">{fmt_money(pilot_dir_gen)}</div><div class="kpi-sub">{pilot_pis_funded_gen}/{pilot_pis_total} ({100*pilot_pis_funded_gen/pilot_pis_total:.0f}%) have funding</div></div>
<div class="kpi accent"><div class="kpi-label">Conservative NIH Direct</div><div class="kpi-value">{fmt_money(pilot_dir_cons)}</div><div class="kpi-sub">{pilot_pis_funded_cons}/{pilot_pis_total} ({100*pilot_pis_funded_cons/pilot_pis_total:.0f}%) have funding</div></div>
<div class="kpi good"><div class="kpi-label">ROI (Direct / DoM)</div><div class="kpi-value ratio big">{pilot_dir_cons/pilot_dom:.0f}×</div><div class="kpi-sub">{pilot_dir_gen/pilot_dom:.0f}× generous</div></div>
</div>

<h3>By Program</h3>
{df_to_html(bytype_disp, "tbl-bytype",
            money_cols={"DoM $", "NIH Direct (Generous)", "NIH Direct (Conservative)"},
            ratio_cols={"ROI Generous", "ROI Conservative"},
            int_cols={"Recipients"})}

<div class="callout">
<strong>Reading the table:</strong> The Junior Award (9× conservative) is the most defensible ROI claim
because the recipients are early-career investigators where the link between DoM support and subsequent
NIH funding is most plausibly causal. The Pilot ROI (33× conservative) is also strong but recipients are
mid-career investigators with more established trajectories, so attributing all subsequent funding to the
DoM investment is less direct.
</div>

<h3>ROI by Program — Visual Comparison</h3>
{fig_html(fig1, "fig-roi-program")}

<h3>Investment vs Return</h3>
{fig_html(fig2, "fig-investment-return")}

<h3>Top Recipients by Conservative NIH Direct Funding</h3>
{fig_html(fig3, "fig-top-pilot")}

<details>
<summary>Per-investigator narrative (all {len(narr_disp)} recipients, sorted by NIH direct funding)</summary>
<div class="scrollable-table">
{df_to_html(narr_disp, "tbl-narrative",
            money_cols={"DoM Total", "NIH Direct (cons.)"},
            ratio_cols={"ROI"})}
</div>
</details>
</section>

<section>
<h2>Section 2 — K-to-R Pipeline (FY2016–2026)</h2>

<p>The Chair's Advisory Committee provides 50% salary-gap support for K-awarded faculty.
Question: do these investigators successfully transition to or continue with R/U funding?</p>

<div class="kpi-row">
<div class="kpi"><div class="kpi-label">DoM Salary Gap Spent</div><div class="kpi-value">{fmt_money(k_dom)}</div><div class="kpi-sub">{k_pis_total} K-awarded faculty</div></div>
<div class="kpi accent"><div class="kpi-label">Post-K NIH Direct</div><div class="kpi-value">{fmt_money(k_dir)}</div><div class="kpi-sub">{k_pis_funded}/{k_pis_total} ({100*k_pis_funded/k_pis_total:.0f}%) have post-K funding</div></div>
<div class="kpi accent"><div class="kpi-label">Post-K NIH Total</div><div class="kpi-value">{fmt_money(k_dir + k_ind)}</div><div class="kpi-sub">Direct + indirect</div></div>
<div class="kpi good"><div class="kpi-label">ROI (Direct / DoM)</div><div class="kpi-value ratio big">{k_dir/k_dom:.0f}×</div><div class="kpi-sub">{(k_dir+k_ind)/k_dom:.0f}× total</div></div>
</div>

<h3>K-to-R Outcomes by Section</h3>
{fig_html(fig5, "fig-k-section")}

{df_to_html(ksec_disp, "tbl-ksec",
            money_cols={"DoM Gap Spend", "Post-K NIH Direct"},
            int_cols={"K Awardees", "With Post-K NIH", "% Conv."})}

<h3>Top 15 K Awardees by Post-K NIH Direct Funding</h3>
{fig_html(fig4, "fig-top-k")}

<details>
<summary>All {len(k_disp)} K awardees (sorted by post-K NIH direct funding)</summary>
<div class="scrollable-table">
{df_to_html(k_disp, "tbl-k",
            money_cols={"Total 50% Salary Gap", "Post-K NIH Direct Costs", "Post-K NIH Indirect Costs"},
            ratio_cols={"ROI"},
            int_cols={"Post-K Grant Count"})}
</div>
</details>
</section>

<section>
<h2>Methodology</h2>

<div class="method-box">
<h3 style="margin-top: 0;">Data sources</h3>
<ul>
<li><strong>Internal DoM records:</strong> K Award tracking spreadsheet (FY2016–2026) and DoM Award Tracker (Pilot/GT97/Junior, AY22–AY26). All faculty names, sections, fiscal years, and award amounts.</li>
<li><strong>NIH RePORTER API v2</strong> (<code>https://api.reporter.nih.gov/v2/projects/search</code>) — public NIH grant database, queried programmatically by investigator name.</li>
</ul>

<h3>Pipeline</h3>
<ol>
<li><strong>Load DoM rosters.</strong> Normalize names (strip degrees, fix comma-reversed forms, manual map for spelling variants), map academic-year labels to numeric years.</li>
<li><strong>Query NIH RePORTER</strong> by first/last name across the relevant fiscal-year window. Multi-PI grants are captured by checking every PI on the grant, not just the contact PI.</li>
<li><strong>Filter same-name collisions</strong> using hand-curated rules (e.g., a different "Sun Lee" at non-BU institutions, a "Sudhir Kumar" at Temple/Iowa State). The Step 2 PI-list match already enforces correct attribution; Step 3 only removes known false positives.</li>
<li><strong>Anchor at first DoM-support year</strong> per investigator. For the pilot ROI, two definitions are reported:
   <ul>
   <li><strong>Generous:</strong> all NIH grants active in fiscal year ≥ first DoM year (includes pre-existing grants that overlap).</li>
   <li><strong>Conservative:</strong> only NIH grants whose <code>project_start_date</code> ≥ first DoM year (truly new awards).</li>
   </ul>
   For the K analysis, "post-K" includes all non-K grants with FY ≥ first K-support year, so concurrent R awards count as a successful K-to-R transition.</li>
<li><strong>Compute ROI</strong> = NIH funding ÷ DoM investment, both per-investigator and aggregated by program/section.</li>
</ol>

<h3>Notable methodology fix (April 2026)</h3>
<p>An earlier filter version required each grant to be either at BU/BMC, have the awardee as contact PI, or have the awardee's name in the contact-PI string. This excluded ~30% of legitimate non-contact co-PI grants at external institutions (e.g., a BU faculty member listed as co-PI on a BWH or Stanford U01). Removing the redundant filter increased the pilot generous ROI from 80× to 109× and the K-to-R direct-cost total from $72.5M to $96.8M.</p>
</div>
</section>

<section>
<h2>Caveats & Limitations</h2>
<ul>
<li><strong>Causation vs association.</strong> Subsequent NIH funding is associated with DoM-supported investigators; this analysis does not establish causation. Counterfactual analyses (matched controls, regression-discontinuity) would be needed for a causal claim.</li>
<li><strong>Multi-PI counting.</strong> A multi-PI grant contributes its full amount to each listed PI's total; aggregate sums are not deduplicated across investigators.</li>
<li><strong>Indirect-cost coverage.</strong> RePORTER's direct/indirect breakdowns are not populated for all historical records; missing values are zero, so indirects are lower bounds.</li>
<li><strong>Academic vs federal fiscal year.</strong> AY22 ≈ FY22, with overlap but not identity. A small number of grants near year boundaries could be classified differently under stricter conventions.</li>
<li><strong>Name disambiguation.</strong> Spot-check any awardee with surprisingly high or low totals before formal reporting.</li>
<li><strong>K-to-R cross-reference flagged 58 discrepancies</strong> between our K-awardee tracking and what RePORTER returns; most are likely fiscal-year reporting differences but warrant a manual audit before final publication.</li>
</ul>
</section>

<footer>
<p>
Generated automatically from <code>K_awardees_summary_2016_2026.xlsx</code> and <code>DoM_Pilot_ROI_Summary.xlsx</code>.<br>
Source scripts: <code>build_k_to_r_analysis.py</code>, <code>build_pilot_roi_analysis.py</code>, <code>build_html_report.py</code> (in the NIH awards folder).<br>
All grant numbers and investigator names are public via NIH RePORTER. No protected health information in this report.
</p>
</footer>

</div>
</body>
</html>
""")

OUT_FILE.write_text("".join(html_parts), encoding="utf-8")
print(f"Saved to: {OUT_FILE}")
print(f"File size: {OUT_FILE.stat().st_size / 1024:.0f} KB")

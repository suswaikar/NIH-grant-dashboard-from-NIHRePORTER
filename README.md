# Evans Endowment Research Awards — Dashboard & Analysis Tools

Tools for the **BU Department of Medicine Research Advisory Committee** to track
Evans Endowment-funded awards (K, Pilot, Junior Faculty, Bridge) and evaluate
their return on investment through subsequent NIH funding.

## Two Dashboards

### 1. NIH Grant Dashboard (`app.py`)

Interactive Streamlit dashboard for exploring **all** NIH grants associated with
Boston University and Boston Medical Center investigators. Pulls live data from the
NIH RePORTER API v2 — no local data files needed.

**Features:**
- All grant types (K, R, U, T, F, P, DP) with multi-PI support
- External grants where BU/BMC investigators are co-PIs at other institutions
- Grant-level and PI-level view modes
- K-to-R pipeline analysis
- Sidebar filters: grant category, department, IC, fiscal year range, PI name search
- Password-protected, dark theme

### 2. Evans Endowment ROI Dashboard (`evans_roi_app.py`)

Interactive dashboard for tracking Evans-funded awards and evaluating ROI through
subsequent NIH grants. Reads from bundled DoM data files + live NIH RePORTER queries.

**Tabs:**

| Tab | Content |
|-----|---------|
| **Awards Overview** | Investment table (K, Pilot, Jr Faculty, Bridge × FY), stacked bar chart, unique investigator counts |
| **ROI Summary** | Two sub-tabs: **K Award → R01/Equivalent** (transition rates by section, expandable awardee detail, top 15 by post-K funding) and **Pilot / Junior Faculty** (ROI by program, awardee detail) |
| **Sex Breakdown** | Awards and K-to-R conversion rates by biological sex |
| **Drill Down** | Individual awardees for any award type and fiscal year |
| **Investigator Lookup** | Select any awardee → live NIH RePORTER query showing all their grants (post-K, K, all) |

**Key methodological choices:**
- Conservative ROI: only non-K NIH grants with fiscal year *after* last year of DoM support
- 1-year buffer: awardees whose K ended FY2025+ excluded from transition rate denominator
- K24 (midcareer) treated as a successful outcome, not filtered out
- GT97 excluded from ROI (accounting mechanism, not research investment)
- False positive filtering for common names (Sun Lee, Sudhir Kumar)

## Analysis Scripts

| Script | Purpose |
|--------|---------|
| `build_k_to_r_analysis.py` | Standalone K-to-R transition analysis. Reads K award data, queries NIH RePORTER, produces `K_awardees_summary_2016_2026.xlsx` with 4 tabs |
| `build_pilot_roi_analysis.py` | Pilot/GT97/Junior Faculty ROI analysis → `DoM_Pilot_ROI_Summary.xlsx` |
| `build_html_report.py` | Generates `DoM_Award_ROI_Report.html` — static HTML report with all ROI metrics |
| `export_source_data.py` | Exports all departmental award data to `Evans_Endowment_Awards_Source_Data.xlsx` for institutional data sharing |

## Data Files

### Bundled (`data/` folder, deployed to Streamlit Cloud)

| File | Content |
|------|---------|
| `K award 2016 - 2026.xlsx` | K award tracking: one sheet per FY (FY2016–FY2026) with Section, Name, Award Number, 50% Salary Gap, Fringe, Total Cost |
| `DoM Award Tracker.xlsx` | Pilot, Junior Faculty, and GT97 awards (FY2022–FY2026) |
| `Evans_Endowment_Awards_Source_Data.xlsx` | Combined export with tabs: K Awards, Pilot, Jr Faculty, GT97, Bridge, PI Demographics (sex, current position), README |

### Data path resolution

The apps check for local OneDrive paths first (for development), then fall back to
the `data/` folder (for Streamlit Cloud deployment). No code changes needed between
local and cloud.

## Deployment

Both apps are deployed from this repo on **Streamlit Community Cloud**:

- **NIH Dashboard**: Main file = `app.py`
- **Evans ROI Dashboard**: Main file = `evans_roi_app.py`

### Deploy steps

1. Go to [share.streamlit.io](https://share.streamlit.io) → sign in with GitHub
2. New app → select this repo → branch `main` → set main file
3. Advanced settings → Secrets: `password = "evans2026"`
4. Deploy

### Password protection

Both apps use a shared password gate. Password is stored in Streamlit Cloud Secrets
(not in code). If no `password` key exists in secrets, the gate is skipped (local dev).

## Local development

```bash
# Install dependencies
pip install -r requirements.txt

# Run the NIH dashboard
streamlit run app.py --server.port 8501

# Run the Evans ROI dashboard
streamlit run evans_roi_app.py --server.port 8502

# Run analysis scripts
python build_k_to_r_analysis.py
python build_pilot_roi_analysis.py
python export_source_data.py
```

## Folder structure

```
R01 for K_Pilot_JrFaculty/
├── app.py                          # NIH Grant Dashboard (all BU/BMC grants)
├── evans_roi_app.py                # Evans Endowment ROI Dashboard
├── build_k_to_r_analysis.py        # K-to-R transition analysis script
├── build_pilot_roi_analysis.py     # Pilot/Jr Faculty ROI script
├── build_html_report.py            # HTML report generator
├── export_source_data.py           # Source data export script
├── _build_overview_doc.py          # Word document generator (committee report)
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── .gitignore
├── .streamlit/
│   ├── config.toml                 # Dark theme
│   └── secrets.toml                # Local password (gitignored)
├── data/                           # Bundled data for Streamlit Cloud
│   ├── K award 2016 - 2026.xlsx
│   ├── DoM Award Tracker.xlsx
│   └── Evans_Endowment_Awards_Source_Data.xlsx
└── archive/                        # Old/superseded files
```

## Adding data

To add historical Pilot, Junior Faculty, or Bridge data:

1. Open `data/Evans_Endowment_Awards_Source_Data.xlsx`
2. Add rows to the appropriate tab (same column format: Name, Section, Fiscal Year, Amount)
3. Commit and push — Streamlit Cloud auto-redeploys
4. The dashboard picks up new data on next cache refresh (6 hours or manual reload)

To update PI demographics (sex, current position):
1. Edit the **PI Demographics** tab in the same file
2. Commit and push

## Key metrics (as of April 2026)

- **Total DoM investment**: $6.4M across 89 investigators
- **K award program**: 59 investigators, $4.4M in salary gap support (FY2016–2026)
- **Retention**: 83% of K awardees still at CAMED/BMC
- **Sex balance**: 51% female, 49% male among K awardees
- **K-to-R transition**: ~59% of eligible K awardees (K ended before FY2025) secured post-K NIH funding
- **Post-K NIH direct costs**: $60M+ (conservative estimate)
- **ROI**: ~14× direct costs returned per dollar invested

## Technical notes

- NIH RePORTER API v2 (public, no key needed): max 10K records per query, split by year ranges
- Batch PI queries cached for 6 hours on Streamlit Cloud
- Name matching: first 3 chars of first name + full last name
- All dollar figures from NIH RePORTER are nominal (not inflation-adjusted)

# NIH Grant Dashboard — BU / Boston Medical Center

Interactive Streamlit dashboard for exploring NIH grants awarded to
Boston University and Boston Medical Center faculty.

## Grant types covered
- **K-awards**: K01, K08, K23, K24, K99
- **R-awards**: R00, R01, R03, R21, R34, R61

## Features

| Tab | What you can do |
|-----|----------------|
| 📈 Trends | Awards and funding by year, grant type, NIH institute |
| 🔬 By Investigator | Rank PIs by funding/awards, drill into individual timelines |
| 🏥 By Department | Treemap and trend charts broken down by department |
| 🔗 K → R Pipeline | Conversion rate from K to R, lag-time histogram, career tables |
| 📋 Full Data | Filterable table with CSV export |

## Sidebar filters
- Department
- Activity code (K01, K23, R01, etc.)
- NIH Institute (IC)
- Fiscal year range
- Active awards only
- PI name search

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
streamlit run app.py
```

The app will open at http://localhost:8501

## Notes
- Data is cached after first load — click **Refresh data** in the sidebar to re-fetch
- NIH RePORTER API returns up to 10,000 records per query; BU/BMC K+R awards
  are well within this limit
- Department field comes from NIH's org_dept_type — may be blank for some awards;
  cross-reference with your own faculty list for finer filtering

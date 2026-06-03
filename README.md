# EnergyCAP Bill Flag Analyzer

Streamlit app for analyzing EnergyCAP flag exports with GHG emissions prioritization.

## Files

| File | Purpose |
|---|---|
| `app.py` | Main Streamlit app |
| `parser.py` | Report parsers + GHG conversion logic |
| `requirements.txt` | Python dependencies |

## Supported Uploads

| Report | Required? | Purpose |
|---|---|---|
| **Report-27** (Bill Flags) | ✅ Required | Flag data, vendors, assignees, status |
| **Report-18** (Bill Line Items) | Recommended | GHG usage quantities per bill |
| **Report-03** (Setup Report) | Recommended | Authoritative commodity classification |

- **Multiple Report-27 and Report-18 files** can be uploaded at once and are merged automatically
- **Report-03** is a master data file — upload once, covers all periods

## How to Export from EnergyCAP

**Report-27:** Bills → Menu (≡) → Report-27 Bill Flags → Export to Excel

**Report-18:** Bills → Report-18 Bill Line Item Report → Filter: Line Item Type = Use → Export to Excel

**Report-03:** All Reports → Setup Report for Accounts, Vendors, Cost Centers, Meters, and Sites (Excel only)

## Features

- 📊 **Overview** — KPI metrics, flag frequency, billing period trends, resolution time
- ⚑ **Flag Analysis** — Multi-filter panel (issue, vendor, assignee, status, priority, GHG category) with clickable drill-down charts
- 🏢 **Vendors** — Flag and cost concentration by vendor, clickable drill-down
- 🌿 **GHG Priority Queue** — Bills ranked by emissions impact score (usage × flag risk × unresolved multiplier); user-selectable display unit (kWh/MWh/GJ/MMBtu/THERM)
- ✅ **Action Guide** — Every flag type mapped to specific remediation steps with EnergyCAP navigation paths
- 📋 **Bill Detail** — Searchable/filterable bill table with per-issue guidance

## Performance

All heavy parsing is wrapped in `@st.cache_data` — files are only re-parsed when the uploaded content changes. Filtering and drill-down interactions are instant.

## GHG Commodity Classification

When Report-03 is uploaded, commodities are classified authoritatively from EnergyCAP's own meter master data (87% coverage in a typical portfolio). The remaining 13% fall back to caption-based heuristics.

**GHG categories tracked:** Electricity (Scope 2), Natural Gas (Scope 1), LPG/Propane (Scope 1), Biomass (Scope 1 biogenic), District Heat/Steam (Scope 2), Diesel/Fuel Oil (Scope 1), Gasoline, Aviation Fuel, Coal

**Unit harmonization:** kWh, MWh, GJ, MJ, MMBtu, THERM, DKTHM, MCF, CCF, CF, m³ are all converted to kWh internally. Users select their preferred display unit.

## Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push all three files to a GitHub repo
2. Go to share.streamlit.io → New app
3. Set Main file path to `app.py`
4. Deploy — no secrets or environment variables needed

# EnergyCAP Bill Flag Analyzer

A Streamlit app for analyzing **EnergyCAP Report-27 (Bill Flags)** exports. Upload one or more `.xlsx` files and get an instant interactive dashboard with flag breakdowns, vendor analysis, resolution tracking, and actionable next steps mapped to EnergyCAP's audit rule documentation.

---

## Features

| Tab | What you get |
|-----|-------------|
| **Overview** | KPI cards, flag frequency chart, billing period trend, category/priority/status donuts, resolution time histogram |
| **Flag Analysis** | Drill into any flag type — vendors affected, assignee workload, audit rule explanation |
| **Vendors** | Vendor league table by flag count and cost, per-vendor flag profile |
| **Action Guide** | Every flag type mapped to a specific recommended action + EnergyCAP navigation path. Systemic insights auto-detected from your data |
| **Bill Detail** | Searchable/filterable bill table, expandable per-bill flag explanations, CSV export |

---

## Quick Start

### Run locally

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_ORG/energycap-flag-analyzer.git
cd energycap-flag-analyzer

# 2. Create a virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

### Deploy to Streamlit Community Cloud

1. Push this repo to GitHub (public or private).
2. Go to [share.streamlit.io](https://share.streamlit.io) and click **New app**.
3. Select your repo, branch `main`, and set **Main file path** to `app.py`.
4. Click **Deploy**. No secrets or environment variables needed.

---

## How to Export Report-27 from EnergyCAP

1. In EnergyCAP, go to the **Bills** module.
2. Open the module menu (≡) and select **Report-27 Bill Flags**.
3. Set your date filters (Bill Entry Date range is recommended).
4. Export to **Excel (.xlsx)**.
5. Upload the file in the app sidebar.

You can upload **multiple files** (e.g., different date ranges) — the app combines and deduplicates them automatically.

---

## Supported Flag Types

The app includes built-in guidance for all standard EnergyCAP audit flag types:

**Import / Configuration**
- Rate schedule mismatch
- Serial number mismatch
- Conflicting use units

**Duplicate / Overlap**
- Duplicate bill
- Overlapping bill
- Multiple bills in period

**Statistical Outlier** (uses EnergyCAP quadratic regression)
- Abnormal cost
- Abnormal use
- Abnormal demand
- High use per day
- High cost per day

**Date / Timeline**
- Gap between bills
- Shorter or longer bill
- Late statement date
- Late due date
- Unexpected billing period

**Line Item Review**
- Flagged line item type found
- Flagged line item description found

---

## Requirements

```
streamlit>=1.32.0
pandas>=2.0.0
plotly>=5.18.0
openpyxl>=3.1.0
xlrd>=2.0.1
```

Python 3.9+ recommended.

> **Note:** The app uses the `extract-text` CLI tool when running on Streamlit Cloud (pre-installed in the Streamlit environment). When running locally, it falls back to `openpyxl` for `.xlsx` parsing. If `extract-text` is not available locally, the app will use the openpyxl fallback automatically.

---

## Local Parsing (no extract-text)

If you're running locally and don't have `extract-text`, the app automatically falls back to reading the Excel file with `openpyxl`. The Report-27 layout is parsed by reading merged cell patterns and tab-separated content from each sheet.

---

## License

MIT — free to use, modify, and distribute.

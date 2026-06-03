"""
parser.py — EnergyCAP Report-27, Report-18, Report-03 parsers + GHG logic
All heavy parsing is designed to be called inside @st.cache_data functions.
"""
import re, shutil, subprocess, tempfile, os
from datetime import datetime
from collections import Counter

import pandas as pd
from openpyxl import load_workbook


# ══════════════════════════════════════════════════════════════════════════════
# FLAG METADATA
# ══════════════════════════════════════════════════════════════════════════════
FLAG_META = {
    "Rate schedule mismatch":              {"category":"Import / Configuration","priority":"Medium","cause":"Rate schedule in import file doesn't match the rate schedule assigned to the meter in EnergyCAP.","action":"Update the meter's rate schedule in EnergyCAP to match the utility's current rate, or correct the import file.","in_energycap":"Meter record → Rate Schedule field."},
    "Serial number mismatch":              {"category":"Import / Configuration","priority":"Medium","cause":"Serial number in the import file doesn't match EnergyCAP. May indicate the vendor swapped the physical meter.","action":"Verify with the vendor whether the meter was replaced. If replaced, update the serial number in EnergyCAP.","in_energycap":"Meter record → Serial Number field. If meter was swapped, retire old meter and create a new one."},
    "Conflicting use units":               {"category":"Import / Configuration","priority":"Medium","cause":"The use unit on the bill conflicts with the use unit on the meter in EnergyCAP.","action":"Determine the correct unit of measure. Update the meter configuration in EnergyCAP or correct the import file.","in_energycap":"Meter record → Use Unit field."},
    "Duplicate bill":                      {"category":"Duplicate / Overlap","priority":"High","cause":"Total bill cost equals a prior bill's cost AND start/end dates match — likely a data entry or billing error.","action":"Compare with the prior bill. If truly duplicate, void one. If it's a legitimate re-bill, confirm with the vendor and document.","in_energycap":"More Actions → Void Bill."},
    "Overlapping bill":                    {"category":"Duplicate / Overlap","priority":"High","cause":"One or more bills have overlapping start/end dates on the same account.","action":"Review dates of all overlapping bills. Correct dates if it's a data entry error.","in_energycap":"Account History tab → view adjacent bills and identify the overlap."},
    "Multiple bills in period":            {"category":"Duplicate / Overlap","priority":"High","cause":"More than one bill exists for the same account and billing period.","action":"Determine if these are split bills (valid) or accidental duplicates. Void true duplicates.","in_energycap":"Filter bill list by account and period. More Actions → Void Bill."},
    "Abnormal cost":                       {"category":"Statistical Outlier","priority":"High","cause":"Total cost is a severe statistical outlier (>3.0 std dev) vs last 12 bills using quadratic regression.","action":"Review bill line items for unexpected charges. Check if a rate change occurred. Consider cost recovery if an error is found.","in_energycap":"Open bill → check line items. Account History tab to compare to prior bills."},
    "Abnormal use":                        {"category":"Statistical Outlier","priority":"High","cause":"Usage is a severe statistical outlier compared to the last 12 bills. Could indicate meter issue, leak, or billing error.","action":"Check for operational changes at the site. Contact vendor if usage spike is unexplained. If meter issue suspected, request re-read.","in_energycap":"Open bill → review meter use values. Compare to Account History."},
    "Abnormal demand":                     {"category":"Statistical Outlier","priority":"High","cause":"Demand reading is a severe statistical outlier compared to historical bills.","action":"Investigate if new high-load equipment was added. Verify demand meter readings with the vendor.","in_energycap":"Review demand meter readings via Account History."},
    "Gap between bills":                   {"category":"Date / Timeline","priority":"Medium","cause":"Gap of 2+ days between this bill and the preceding bill — could mean a missing bill.","action":"Check if a bill was skipped or not yet received. Contact vendor or check the vendor portal for missing invoices.","in_energycap":"Account History tab to see the gap. If a bill is missing, manually enter it or re-import."},
    "Shorter or longer bill":              {"category":"Date / Timeline","priority":"Low","cause":"Bill period is significantly shorter or longer than the average of prior bills.","action":"Verify with the vendor that the billing period is correct. If dates are off, correct them.","in_energycap":"Review Start and End Date on the bill header."},
    "Late statement date":                 {"category":"Date / Timeline","priority":"Low","cause":"Statement date is too many days after the bill's end date.","action":"Verify the statement date is correct. If the vendor consistently sends invoices late, consider adjusting the audit threshold.","in_energycap":"Correct the Statement Date in the bill header."},
    "Late due date":                       {"category":"Date / Timeline","priority":"Low","cause":"Due date is too many days after the bill's end date.","action":"Verify due date accuracy to avoid late payment penalties.","in_energycap":"Correct the Due Date in the bill header."},
    "Unexpected billing period":           {"category":"Date / Timeline","priority":"Medium","cause":"Billing period is outside the bill's start and end dates — likely a data entry error.","action":"Correct the billing period to align with the start and end dates of the bill.","in_energycap":"Edit the bill header → set the Billing Period to match the start/end dates."},
    "Flagged line item type found":        {"category":"Line Item Review","priority":"Medium","cause":"Bill contains a line item type configured for monitoring (e.g., late fees, taxes, penalties).","action":"Review the specific line item. If it's a late fee, confirm the prior bill was paid on time. If it's an unexpected charge, dispute with vendor.","in_energycap":"Open bill → scroll to line items section → review the flagged line item type."},
    "Flagged line item description found": {"category":"Line Item Review","priority":"Medium","cause":"Bill contains a line item with a description matching a monitored keyword (case-insensitive).","action":"Review the line item description and value. Common triggers: 'balance forward', 'prior balance', 'penalty'. Confirm whether charge is legitimate.","in_energycap":"Open bill → check line items for the matching description."},
    "High use per day":                    {"category":"Statistical Outlier","priority":"Medium","cause":"Use per day exceeds 3× the highest use-per-day from the top 90% of prior bills (last 4 years).","action":"Investigate operational changes or equipment issues at the site.","in_energycap":"Compare current use/day to historical via Account History tab."},
    "High cost per day":                   {"category":"Statistical Outlier","priority":"Medium","cause":"Cost per day exceeds 3× the highest cost-per-day from the top 90% of prior bills.","action":"Check for rate changes, new charges, or usage spikes driving up cost.","in_energycap":"Open bill and compare line-by-line with a prior period bill."},
}

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}
PRIORITY_COLOR = {"High": "#dc3545", "Medium": "#fd7e14", "Low": "#198754"}
CATEGORY_COLOR = {
    "Import / Configuration": "#0d6efd",
    "Duplicate / Overlap":    "#dc3545",
    "Statistical Outlier":    "#6f42c1",
    "Date / Timeline":        "#fd7e14",
    "Line Item Review":       "#198754",
}


# ══════════════════════════════════════════════════════════════════════════════
# GHG COMMODITY CLASSIFICATION
# ══════════════════════════════════════════════════════════════════════════════
# From Report-03 Commodity field — authoritative classification
GHG_COMMODITIES = {
    "Electric":                 "Electricity",
    "Lighting":                 "Electricity",
    "Solar PV":                 "Electricity",
    "Natural Gas":              "Natural Gas",
    "Methane":                  "Natural Gas",
    "Biogas":                   "Natural Gas",
    "Other Gaseous Fuels":      "Natural Gas",
    "Liquified Petroleum Gases":"LPG / Propane",
    "Propane":                  "LPG / Propane",
    "Butane":                   "LPG / Propane",
    "Biomass":                  "Biomass",
    "Delivered Heat":           "District Heat / Steam",
    "Steam":                    "District Heat / Steam",
    "Diesel Fuel":              "Diesel / Fuel Oil",
    "Fuel Oil":                 "Diesel / Fuel Oil",
    "Gasoline":                 "Gasoline",
    "Aviation Fuel":            "Aviation Fuel",
    "Coal":                     "Coal",
}

NON_GHG_COMMODITIES = {
    "Water", "Water and Sewer", "Sewer", "Fixed Facility Charge",
    "Non-Hazardous Waste - Not Recycl", "Fire Protection", "Solid Waste",
    "Refuse", "Internet", "Telecom", "Transportation",
    "Miscellaneous", "Miscellaneous Supplies",
}

# GHG scope per energy category
GHG_SCOPE = {
    "Electricity":            "Scope 2",
    "Natural Gas":            "Scope 1",
    "LPG / Propane":          "Scope 1",
    "Biomass":                "Scope 1 (biogenic)",
    "District Heat / Steam":  "Scope 2",
    "Diesel / Fuel Oil":      "Scope 1",
    "Gasoline":               "Scope 1",
    "Aviation Fuel":          "Scope 1",
    "Coal":                   "Scope 1",
}

# Unit conversion to kWh (internal base)
TO_KWH = {
    "kWh":   1.0,
    "MWh":   1_000.0,
    "GJ":    277.778,
    "MJ":    0.277778,
    "MMBtu": 293.071,
    "THERM": 29.3071,
    "DKTHM": 293.071,
    "MCF":   293.071,   # 1 MCF nat gas ≈ 1 MMBtu
    "CCF":   29.3071,   # 1 CCF ≈ 1 therm
    "CF":    0.293071,
    "m³":    10.55,     # nat gas, EU avg
    "DKTH":  293.071,
    "DKTHM": 293.071,
}

# Display units (from kWh)
FROM_KWH = {
    "kWh":   1.0,
    "MWh":   1e-3,
    "GJ":    3.6e-3,
    "MMBtu": 3.41214e-3,
    "THERM": 3.41214e-2,
}
DISPLAY_UNITS = ["kWh", "MWh", "GJ", "MMBtu", "THERM"]

# Caption fallback for ambiguous units when R03 not available
WATER_CAPTION_RE = re.compile(
    r'water|sewer|sewage|irrigation|\(water\)|\(sewer\)|h2o|fire line|fire service',
    re.IGNORECASE)
GHG_UNITS_FALLBACK = {"kWh","MWh","GJ","MJ","MMBtu","THERM","DKTHM","MCF","CCF","CF","m³","DKTH"}
ALWAYS_WATER_UNITS = {"gal","Kgal","Hgal","Mgal","Mgal","L","Hgal"}


def kwh(value: float, unit: str) -> float:
    return value * TO_KWH.get(unit, 0.0)

def from_kwh(kwh_val: float, display_unit: str) -> float:
    return kwh_val * FROM_KWH.get(display_unit, 1.0)


# ══════════════════════════════════════════════════════════════════════════════
# REPORT-03 PARSER  (setup/master data)
# ══════════════════════════════════════════════════════════════════════════════
def parse_report03(path: str) -> pd.DataFrame:
    """
    Parse Report-03 Setup Report.
    Returns DataFrame keyed by meter_code with commodity, ghg_category, scope,
    vendor_role, account_number, cost_center_code, building_code.
    """
    wb = load_workbook(path, data_only=True)
    ws = wb['Sheet1']

    rows = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        meter_code = str(row[26]).strip() if row[26] else ''
        if not meter_code or meter_code == 'None':
            continue
        commodity = str(row[51]).strip() if row[51] else ''
        rows.append({
            'meter_code':          meter_code,
            'account_number':      str(row[1]).strip()  if row[1]  else '',
            'vendor_name':         str(row[19]).strip() if row[19] else '',
            'vendor_code':         str(row[20]).strip() if row[20] else '',
            'vendor_role':         str(row[21]).strip() if row[21] else '',
            'cost_center_code':    str(row[18]).strip() if row[18] else '',
            'building_code':       str(row[53]).strip() if row[53] else '',
            'building_name':       str(row[52]).strip() if row[52] else '',
            'commodity':           commodity,
            'ghg_category':        GHG_COMMODITIES.get(commodity, ''),
            'is_ghg':              commodity in GHG_COMMODITIES,
            'ghg_scope':           GHG_SCOPE.get(GHG_COMMODITIES.get(commodity, ''), ''),
        })

    df = pd.DataFrame(rows).drop_duplicates('meter_code')
    return df


# ══════════════════════════════════════════════════════════════════════════════
# REPORT-18 PARSER  (bill line item usage)
# ══════════════════════════════════════════════════════════════════════════════
def parse_report18(path: str, setup_df: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Parse Report-18 Bill Line Item Report (Use sheet).
    Joins with setup_df on meter_code for authoritative commodity classification.
    Falls back to caption heuristics when setup_df not available.
    Returns one row per bill_id with total_kwh_equivalent and primary_unit.
    """
    wb = load_workbook(path, data_only=True)

    # Build meter→commodity lookup from setup if available
    meter_lookup = {}
    if setup_df is not None and not setup_df.empty:
        meter_lookup = setup_df.set_index('meter_code')[
            ['commodity','ghg_category','is_ghg','ghg_scope']
        ].to_dict('index')

    all_rows = []

    for sheet_name in wb.sheetnames:
        if sheet_name.lower() in {'report overview', 'overview', 'demand', 'info_use'}:
            continue
        ws = wb[sheet_name]
        header_idx = {}
        data_started = False

        for row in ws.iter_rows(values_only=True):
            vals = list(row)
            str_vals = [str(v).strip() if v is not None else '' for v in vals]

            if 'Bill ID' in str_vals and 'Unit' in str_vals and not data_started:
                for j, v in enumerate(str_vals):
                    if v: header_idx[v] = j
                data_started = True
                continue

            if not data_started:
                continue

            bill_id_raw = vals[header_idx.get('Bill ID', 9)]
            if bill_id_raw is None:
                continue
            try:
                bill_id = str(int(float(str(bill_id_raw)))).strip()
            except (ValueError, TypeError):
                continue

            meter_code     = str(vals[header_idx.get('Meter Code', 1)] or '').strip()
            caption        = str(vals[header_idx.get('Caption', 12)] or '').strip()
            unit           = str(vals[header_idx.get('Unit', 14)] or '').strip()
            billing_period = vals[header_idx.get('Billing Period', 8)]
            account_code   = str(vals[header_idx.get('Account Code', 0)] or '').strip()

            try:
                value = float(vals[header_idx.get('Value', 13)] or 0)
            except (ValueError, TypeError):
                value = 0.0

            if value <= 0:
                continue

            # ── Commodity classification ──────────────────────────────────────
            if meter_code in meter_lookup:
                meta = meter_lookup[meter_code]
                is_ghg      = meta['is_ghg']
                ghg_category = meta['ghg_category']
                ghg_scope    = meta['ghg_scope']
                commodity    = meta['commodity']
            else:
                # Caption fallback
                is_ghg = (
                    unit not in ALWAYS_WATER_UNITS
                    and unit in GHG_UNITS_FALLBACK
                    and not WATER_CAPTION_RE.search(caption)
                )
                ghg_category = ''
                ghg_scope    = ''
                commodity    = ''

            if not is_ghg:
                continue

            kwh_eq = kwh(value, unit)
            if kwh_eq == 0:
                continue

            all_rows.append({
                'bill_id':        bill_id,
                'meter_code':     meter_code,
                'account_code':   account_code,
                'billing_period': str(billing_period) if billing_period else '',
                'caption':        caption,
                'value':          value,
                'unit':           unit,
                'commodity':      commodity,
                'ghg_category':   ghg_category,
                'ghg_scope':      ghg_scope,
                'kwh_equivalent': kwh_eq,
            })

    if not all_rows:
        return pd.DataFrame()

    df = pd.DataFrame(all_rows)

    # Sum kWh per bill_id across all GHG commodities
    total_kwh = (df.groupby('bill_id')['kwh_equivalent']
                   .sum().rename('total_kwh_equivalent'))

    # Primary commodity = highest kWh per bill (handles multi-commodity bills)
    primary = (df.sort_values('kwh_equivalent', ascending=False)
                 .drop_duplicates('bill_id')
                 [['bill_id','unit','value','kwh_equivalent',
                   'commodity','ghg_category','ghg_scope','meter_code']]
                 .rename(columns={
                     'unit':          'primary_unit',
                     'value':         'primary_value',
                     'kwh_equivalent':'primary_kwh',
                 }))

    result = primary.merge(total_kwh, on='bill_id', how='left')
    return result


# ══════════════════════════════════════════════════════════════════════════════
# REPORT-27 PARSER  (bill flags)
# ══════════════════════════════════════════════════════════════════════════════
def parse_report27_text(text: str) -> pd.DataFrame:
    lines = text.split("\n")
    records, current = [], {}

    for line in lines:
        stripped = line.strip()

        if re.search(r'Account:\s{2,}', stripped):
            if current.get("bill_id"):
                records.append(current)
            current = {
                "account":"","address":"","vendor":"","bill_id":"",
                "billing_period":"","cost":0.0,"status":"",
                "flag_issues":"","assigned_to":"","cost_recovery":0.0,
                "flagged_date":None,"resolved_date":None,
                "num_issues":0,"resolvers":[],"flag_events":[],
            }
            current["account"] = re.sub(r'Account:\s+','',stripped).strip()
            continue

        if (current.get("account") and not current.get("address")
                and not current.get("vendor") and stripped
                and "Vendor:" not in stripped
                and not re.match(r'\d{6}', stripped)):
            current["address"] = stripped; continue

        if "Vendor:" in stripped:
            vm = re.search(r'Vendor:\s+(.+?)(?:\s*\[|$)', stripped)
            if vm: current["vendor"] = vm.group(1).strip()
            continue

        if re.match(r'\d{6}', stripped) and "Billing Period" not in stripped:
            parts = stripped.split()
            if parts and parts[0].isdigit() and len(parts[0]) == 6:
                current["bill_id"] = parts[0]
                cm = re.search(r'(\d[\d,]*\.\d{4})\s*$', line.strip())
                if cm: current["cost"] = float(cm.group(1).replace(",",""))
                pm = re.search(r'\t(20\d{4})\t', line)
                if pm: current["billing_period"] = pm.group(1)
            continue

        if "Flag Type:" in stripped:
            sm = re.search(r'Flag Status:\s*(\w+)', stripped)
            if sm: current["status"] = sm.group(1)
            am = re.search(r'Assigned to:\s*\t+(.+?)(?:\t{4,}|Cost Recovery)', stripped)
            if am: current["assigned_to"] = am.group(1).strip()
            rm = re.search(r'Cost Recovery:\s*\$([0-9,.]+)', stripped)
            if rm: current["cost_recovery"] = float(rm.group(1).replace(",",""))
            continue

        if stripped.startswith("Flag Issue:"):
            im = re.search(r'Flag Issue:\s*(.+?)(?:\t{2,}|$)', stripped)
            if im:
                raw = im.group(1).strip()
                current["flag_issues"] = raw
                current["num_issues"]  = len([x for x in raw.split(",") if x.strip()])
            continue

        if re.match(r'\d{2}/\d{2}/\d{4}', stripped):
            dtm = re.match(r'(\d{2}/\d{2}/\d{4} \d{2}:\d{2} (?:AM|PM))', stripped)
            if dtm:
                try:    dt = datetime.strptime(dtm.group(1), "%m/%d/%Y %I:%M %p")
                except: dt = None
                if "Bill flagged" in stripped or "flagged as Audit" in stripped:
                    if not current.get("flagged_date") and dt:
                        current["flagged_date"] = dt
                elif "Flag resolved" in stripped and dt:
                    current["resolved_date"] = dt
                    am2 = re.search(r'\d{2}:\d{2} (?:AM|PM) ([\w.@]+) Flag resolved', stripped)
                    current["resolvers"].append(am2.group(1) if am2 else "SYSTEM")

    if current.get("bill_id"):
        records.append(current)
    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["billing_period_dt"]    = pd.to_datetime(df["billing_period"], format="%Y%m", errors="coerce")
    df["billing_period_label"] = df["billing_period_dt"].dt.strftime("%b %Y")
    df["days_to_resolve"] = df.apply(
        lambda r: (r["resolved_date"] - r["flagged_date"]).days
        if r["resolved_date"] and r["flagged_date"] else None, axis=1)
    df["primary_resolver"] = df["resolvers"].apply(lambda x: x[-1] if x else "Unresolved")
    df["issues_list"] = df["flag_issues"].apply(
        lambda x: [i.strip() for i in x.split(",") if i.strip()] if x else [])
    df["primary_issue"]    = df["issues_list"].apply(lambda x: x[0] if x else "Unknown")
    df["primary_category"] = df["primary_issue"].apply(
        lambda x: FLAG_META.get(x,{}).get("category","Other"))
    df["primary_priority"] = df["primary_issue"].apply(
        lambda x: FLAG_META.get(x,{}).get("priority","Medium"))
    return df


# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS  (temp-file wrappers for Streamlit UploadedFile objects)
# ══════════════════════════════════════════════════════════════════════════════
def _tmp_save(uploaded_file) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
        f.write(uploaded_file.read())
        return f.name

def load_report03(uploaded_file) -> pd.DataFrame:
    path = _tmp_save(uploaded_file)
    try:    return parse_report03(path)
    finally: os.unlink(path)

def load_report18(uploaded_file, setup_df=None) -> pd.DataFrame:
    path = _tmp_save(uploaded_file)
    try:    return parse_report18(path, setup_df)
    finally: os.unlink(path)

def load_report27(uploaded_file) -> pd.DataFrame:
    path = _tmp_save(uploaded_file)
    try:
        text = ""
        if shutil.which("extract-text"):
            r = subprocess.run(["extract-text", path], capture_output=True, text=True, timeout=60)
            text = r.stdout
        if not text.strip():
            wb = load_workbook(path, data_only=True)
            lines = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    lines.append("\t" + "\t".join(str(c) if c is not None else "" for c in row))
            text = "\n".join(lines)
        return parse_report27_text(text) if text.strip() else pd.DataFrame()
    finally:
        os.unlink(path)

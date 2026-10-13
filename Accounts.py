import os
import glob
import re
import pandas as pd
import numpy as np
import openpyxl

# ==============================================================================
# FOLDER PATHS
# ==============================================================================
if os.path.exists("data"):
    WATCH_FOLDER_PATH = "data"
    OUTPUT_FOLDER_PATH = "."
elif os.path.exists(r"G:\My Drive\Accounts Data Automation"):
    WATCH_FOLDER_PATH = r"G:\My Drive\Accounts Data Automation"
    OUTPUT_FOLDER_PATH = r"G:\My Drive\Accounts Data Output"
else:
    WATCH_FOLDER_PATH = "data"
    OUTPUT_FOLDER_PATH = "."

os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)
OUTPUT_FILE = os.path.join(OUTPUT_FOLDER_PATH, "CFO_Daily_Master_Report.xlsx")

# ==============================================================================
# 1. DATE PARSER
# ==============================================================================
def parse_date_robust(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return np.nan
    val_str = str(val).strip()
    try:
        parsed = pd.to_datetime(val_str, errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%Y-%m-%d")
    except Exception:
        pass

    match = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", val_str)
    if match:
        y, m, d = match.groups()
        return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"

    match = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", val_str)
    if match:
        m, d, y = match.groups()
        return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"

    return val_str

# ==============================================================================
# 2. NUMERIC CLEANER
# ==============================================================================
def clean_amount(val):
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).strip()
    if val_str.upper() in ["N/A", "NAN", "NULL", "NONE", ""]:
        return 0.0
    cleaned = re.sub(r"[^\d.-]", "", val_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

# ==============================================================================
# 3. LOAD DATA
# ==============================================================================
def load_and_clean_data(folder_path):
    all_files = glob.glob(os.path.join(folder_path, "**", "*.xlsx"), recursive=True) + \
                glob.glob(os.path.join(folder_path, "**", "*.csv"), recursive=True)
    valid_files = [f for f in all_files if not os.path.basename(f).startswith("~$")]

    if not valid_files:
        print(f"No files found in {folder_path} to process.")
        return pd.DataFrame()

    print(f"Found {len(valid_files)} files to process: {[os.path.basename(f) for f in valid_files]}")

    dataframes = []
    for filepath in sorted(valid_files):
        try:
            if filepath.endswith(".csv"):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)

            df.columns = [c.strip() for c in df.columns]

            fname_match = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(filepath))
            fallback_date = fname_match.group(1) if fname_match else np.nan

            if "Posting_Date" in df.columns:
                df["Posting_Date"] = df["Posting_Date"].apply(parse_date_robust)
                if fallback_date:
                    df["Posting_Date"] = df["Posting_Date"].fillna(fallback_date)

            dataframes.append(df)
        except Exception as e:
            print(f"Error reading {filepath}: {e}")

    if not dataframes:
        return pd.DataFrame()

    consolidated_df = pd.concat(dataframes, ignore_index=True)

    string_cols = ["Entity", "Department", "Account_Name", "Account_Type", "Payment_Method", "Status"]
    for col in string_cols:
        if col in consolidated_df.columns:
            consolidated_df[col] = consolidated_df[col].astype(str).str.strip().str.title()

    if "Amount" in consolidated_df.columns:
        consolidated_df["Amount"] = consolidated_df["Amount"].apply(clean_amount)

    def get_signed_amount(row):
        acc_type = str(row.get("Account_Type", "")).lower()
        amt = row.get("Amount", 0.0)
        if "revenue" in acc_type:
            return amt
        elif any(k in acc_type for k in ["cogs", "opex", "capex", "cost"]):
            return -amt
        return amt

    consolidated_df["Signed_Amount"] = consolidated_df.apply(get_signed_amount, axis=1)
    return consolidated_df

# ==============================================================================
# 4. BUILD FINANCIAL SUMMARIES
# ==============================================================================
def generate_cfo_daily_summary(df):
    if df.empty:
        return pd.DataFrame()

    records = []
    for date, group in df.groupby("Posting_Date"):
        if pd.isna(date) or str(date).lower() in ["nan", "none"]:
            continue

        rev = group[group["Account_Type"].str.contains("Revenue", case=False, na=False)]["Amount"].sum()
        cogs = group[group["Account_Type"].str.contains("Cogs|Cost", case=False, na=False)]["Amount"].sum()
        opex = group[group["Account_Type"].str.contains("Opex", case=False, na=False)]["Amount"].sum()
        capex = group[group["Account_Type"].str.contains("Capex", case=False, na=False)]["Amount"].sum()

        ebitda = rev - (cogs + opex)
        net_cash_flow = rev - (cogs + opex + capex)
        pending_held = group[group["Status"].str.contains("Hold|Pending", case=False, na=False)]["Amount"].sum()
        total_entries = len(group)

        records.append({
            "Posting_Date": date,
            "Daily_Revenue_Inflow ($)": round(rev, 2),
            "COGS ($)": round(cogs, 2),
            "OpEx ($)": round(opex, 2),
            "Operating_Profit_EBITDA ($)": round(ebitda, 2),
            "CapEx ($)": round(capex, 2),
            "Net_Daily_Cash_Flow ($)": round(net_cash_flow, 2),
            "Pending_or_Held_Exposure ($)": round(pending_held, 2),
            "Total_Processed_Entries": total_entries
        })

    summary_df = pd.DataFrame(records)
    if not summary_df.empty:
        summary_df = summary_df.sort_values("Posting_Date").reset_index(drop=True)
    return summary_df

def generate_dept_breakdown(df):
    if df.empty:
        return pd.DataFrame()
    spend_df = df[df["Account_Type"].str.contains("Cogs|Opex|Cost", case=False, na=False)]
    dept_group = spend_df.groupby(["Department", "Account_Type"])["Amount"].sum().reset_index()
    dept_group = dept_group.rename(columns={"Amount": "Total_Spend_USD"})
    dept_group["Total_Spend_USD"] = dept_group["Total_Spend_USD"].round(2)
    return dept_group.sort_values(["Department", "Account_Type"]).reset_index(drop=True)

# ==============================================================================
# 5. EXECUTION & EXCEL EXPORT
# ==============================================================================
def main():
    print(f"Reading data from: {WATCH_FOLDER_PATH}")
    raw_df = load_and_clean_data(WATCH_FOLDER_PATH)

    if raw_df.empty:
        print("No data processed. Exiting.")
        return

    cfo_summary = generate_cfo_daily_summary(raw_df)
    dept_breakdown = generate_dept_breakdown(raw_df)

    print(f"Writing master report to: {OUTPUT_FILE}")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        cfo_summary.to_excel(writer, sheet_name="CFO_Daily_Summary", index=False)
        dept_breakdown.to_excel(writer, sheet_name="Department_Breakdown", index=False)
        raw_df.to_excel(writer, sheet_name="Consolidated_Clean_Data", index=False)

    print("Report generated successfully with all available dates.")

if __name__ == "__main__":
    main()

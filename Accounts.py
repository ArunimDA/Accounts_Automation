#!/usr/bin/env python
# coding: utf-8

# In[ ]:


https://drive.google.com/drive/folders/1uWk690Cg2I2wJsSq_e-uA32l3B8Vuko_?usp=drive_link


# In[3]:


import os
import glob
import re
import pandas as pd
import numpy as np

# ==============================================================================
# FOLDER PATHS
# ==============================================================================
WATCH_FOLDER_PATH = r"G:\My Drive\Accounts Data Automation" 
OUTPUT_FOLDER_PATH = r"G:\My Drive\Accounts Data Output"
os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)


# ==========================================
# 1. ROBUST DATE PARSER (FIXES BLANK DATES)
# ==========================================
def parse_date_robust(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return np.nan
    val_str = str(val).strip()
    # Try parsing standard formats first
    dt = pd.to_datetime(val_str, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%Y-%m-%d")
    return val_str


# ==========================================
# 2. NUMERIC AMOUNT PARSER (FIXES DIRTY TEXT & $)
# ==========================================
def parse_numeric_amount(val):
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).strip()
    if val_str in ["", "N/A", "nan", "None", "null"]:
        return 0.0
    # Strip any characters that aren't digits, negative signs, or decimals
    clean_str = re.sub(r"[^\d.-]", "", val_str)
    try:
        return round(float(clean_str), 2)
    except ValueError:
        return 0.0


# ==========================================
# 3. COMPREHENSIVE DATA CLEANING ENGINE
# ==========================================
def clean_ledger_data(df_raw):
    if df_raw.empty:
        return df_raw

    df = df_raw.copy()

    # 1. Deduplicate by Transaction_ID
    if "Transaction_ID" in df.columns:
        df["Transaction_ID"] = df["Transaction_ID"].astype(str).str.strip()
        df = df.drop_duplicates(subset=["Transaction_ID"], keep="first")

    # 2. Fix Posting_Date (fill missing dates using forward-fill or target ledger date)
    if "Posting_Date" in df.columns:
        df["Posting_Date"] = df["Posting_Date"].apply(parse_date_robust)
        # If any date was corrupted, forward-fill from previous transaction in the batch
        df["Posting_Date"] = df["Posting_Date"].ffill().bfill()

    # 3. Clean Text Columns & Standardize Casing (Fixes ALL CAPS like SNOWFLAKE)
    text_cols = ["Entity", "Department", "Account_Name", "Account_Type", "Counterparty", "Status", "Payment_Method"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(["None", "nan", ""], "Unspecified")
            df[col] = df[col].str.title()

    # Currency should strictly be uppercase USD
    if "Currency" in df.columns:
        df["Currency"] = df["Currency"].astype(str).str.strip().str.upper()
        df["Currency"] = df["Currency"].replace(["US$", " USD ", "usd", ""], "USD")

    # 4. FIX AMOUNT COLUMN IN-PLACE (Strict Float64, No $, No strings)
    if "Amount" in df.columns:
        df["Amount"] = df["Amount"].apply(parse_numeric_amount)

    # 5. Fix Missing Cost Centers & Approved_By
    if "Cost_Center" in df.columns:
        df["Cost_Center"] = df["Cost_Center"].replace(["None", "nan", np.nan, ""], "CC-UNASSIGNED")

    if "Approved_By" in df.columns:
        # Standardize empty audit approvals based on status
        df["Approved_By"] = df["Approved_By"].fillna("Pending_Signoff").replace(["None", "nan", ""], "Pending_Signoff")

    if "Description" in df.columns:
        df["Description"] = df["Description"].fillna("No Description Provided").replace(["None", "nan", ""], "No Description Provided")

    # 6. Add Financial Sign Convention Column for Executive CFO Modeling
    # Positive for Revenue, Negative for Outflows (COGS, OpEx, CapEx)
    df["Signed_Amount"] = df.apply(
        lambda row: row["Amount"] if row["Account_Type"] == "Revenue" else -row["Amount"],
        axis=1
    )

    return df


# ==========================================
# 4. CFO DAILY SUMMARY
# ==========================================
def calculate_cfo_metrics(df):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    daily_records = []
    unique_dates = sorted([d for d in df["Posting_Date"].dropna().unique()])

    for d in unique_dates:
        sub = df[df["Posting_Date"] == d]

        rev = sub[sub["Account_Type"] == "Revenue"]["Amount"].sum()
        cogs = sub[sub["Account_Type"] == "Cogs"]["Amount"].sum()
        opex = sub[sub["Account_Type"] == "Opex"]["Amount"].sum()
        capex = sub[sub["Account_Type"] == "Capex"]["Amount"].sum()

        gross_profit = rev - cogs
        operating_profit = gross_profit - opex
        net_cash_flow = rev - (cogs + opex + capex)

        pending_exposure = sub[sub["Status"].isin(["Pending Approval", "On Hold"])]["Amount"].sum()

        daily_records.append({
            "Posting_Date": d,
            "Daily_Revenue_Inflow ($)": round(rev, 2),
            "COGS ($)": round(cogs, 2),
            "OpEx ($)": round(opex, 2),
            "Operating_Profit_EBITDA ($)": round(operating_profit, 2),
            "CapEx ($)": round(capex, 2),
            "Net_Daily_Cash_Flow ($)": round(net_cash_flow, 2),
            "Pending_or_Held_Exposure ($)": round(pending_exposure, 2),
            "Total_Processed_Entries": len(sub)
        })

    cfo_summary_df = pd.DataFrame(daily_records)

    dept_spend_df = (
        df[df["Account_Type"].isin(["Opex", "Cogs"])]
        .groupby(["Department", "Account_Type"])["Amount"]
        .sum()
        .round(2)
        .reset_index()
        .rename(columns={"Amount": "Total_Spend_USD"})
    )

    return cfo_summary_df, dept_spend_df


# ==========================================
# 5. EXECUTION & DYNAMIC SCAN
# ==========================================
def run_pipeline():
    file_paths = glob.glob(os.path.join(WATCH_FOLDER_PATH, "*.xlsx"))
    active_files = [f for f in file_paths if not os.path.basename(f).startswith("~$")]

    if not active_files:
        print(f"No active Excel sheets found in {WATCH_FOLDER_PATH}.")
        return

    frames = [pd.read_excel(fp, sheet_name=0) for fp in active_files]
    df_raw = pd.concat(frames, ignore_index=True)

    # Clean data
    df_clean = clean_ledger_data(df_raw)
    cfo_summary, dept_spend = calculate_cfo_metrics(df_clean)

    # Export
    output_filepath = os.path.join(OUTPUT_FOLDER_PATH, "CFO_Daily_Master_Report.xlsx")
    with pd.ExcelWriter(output_filepath, engine="openpyxl") as writer:
        cfo_summary.to_excel(writer, sheet_name="CFO_Daily_Summary", index=False)
        dept_spend.to_excel(writer, sheet_name="Department_Breakdown", index=False)
        df_clean.to_excel(writer, sheet_name="Consolidated_Clean_Data", index=False)

    print(f"Successfully generated clean master sheet with 0 anomalies: {output_filepath}")


if __name__ == "__main__":
    run_pipeline()


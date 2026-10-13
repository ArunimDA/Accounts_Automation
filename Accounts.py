import glob
import os
import re
import numpy as np
import openpyxl
import pandas as pd

# ==============================================================================
# FOLDER PATHS (Relative paths for GitHub Actions & Local execution)
# ==============================================================================
# If running on GitHub, it reads from 'data' downloaded via gdown.
# If running locally without 'data', it can fallback to your local Drive path.
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
# 1. ROBUST DATE PARSER
# ==============================================================================
def parse_date_robust(val):
  if pd.isna(val) or val is None or str(val).strip() == "":
    return np.nan
  val_str = str(val).strip()
  # Standard pandas parse attempt
  try:
    parsed = pd.to_datetime(val_str, errors="coerce")
    if pd.notna(parsed):
      return parsed.strftime("%Y-%m-%d")
  except Exception:
    pass

  # Fallback regex parsing (YYYY-MM-DD, MM/DD/YYYY, etc.)
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
# 2. NUMERIC AMOUNT CLEANER
# ==============================================================================
def clean_amount(val):
  if pd.isna(val) or val is None:
    return 0.0
  val_str = str(val).strip()
  if val_str.upper() in ["N/A", "NAN", "NULL", "NONE", ""]:
    return 0.0
  # Strip currency signs, commas, and letters like USD/US$
  cleaned = re.sub(r"[^\d.-]", "", val_str)
  try:
    return float(cleaned)
  except ValueError:
    return 0.0


# ==============================================================================
# 3. LOAD AND CONSOLIDATE ALL RAW FILES
# ==============================================================================
def load_and_clean_data(folder_path):
  all_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(
      os.path.join(folder_path, "*.csv")
  )
  valid_files = [
      f for f in all_files if not os.path.basename(f).startswith("~$")
  ]

  if not valid_files:
    print(f"No files found in {folder_path} to process.")
    return pd.DataFrame()

  print(f"Found {len(valid_files)} files: {[os.path.basename(f) for f in valid_files]}")

  dataframes = []
  for filepath in sorted(valid_files):
    try:
      if filepath.endswith(".csv"):
        df = pd.read_csv(filepath)
      else:
        df = pd.read_excel(filepath)

      # Standardize column names
      df.columns = [c.strip() for c in df.columns]

      # Extract date from filename if Posting_Date is blank
      fname_date_match = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(filepath))
      fallback_date = fname_date_match.group(1) if fname_date_match else np.nan

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

  # Clean string fields
  string_cols = ["Entity", "Department", "Account_Name", "Account_Type", "Payment_Method", "Status"]
  for col in string_cols:
    if col in consolidated_df.columns:
      consolidated_df[col] = consolidated_df[col].astype(str).str.strip().str.title()

  # Clean Amount
  if "Amount" in consolidated_df.columns:
    consolidated_df["Amount"] = consolidated_df["Amount"].apply(clean_amount)

  # Signed Amount mapping
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
        "Net_Daily_Cash_Flow ($

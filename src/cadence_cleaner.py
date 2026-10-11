import pandas as pd
from pathlib import Path
import datetime
import openpyxl
import numpy as np
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================================
# CONFIG
# =========================================================

BASE_FOLDER = Path(__file__).resolve().parent.parent
INPUT_FILE = BASE_FOLDER / "raw" / "raw.xlsx"

# Dynamic Desktop Output Directory (Works on Windows, macOS, and Linux)
OUTPUT_DIR = Path.home() / "Desktop" / "Cleaned Data"

REQUIRED_COLUMNS = ["STORE_NAME", "BRAND", "EXT_PRICE_W_TAX", "RECEIPT_NO", "CUSTOMER_NAME"]

COLUMNS_TO_DROP = [
    "SERIAL_NO",
    "MOBILE_NUMBER",
    "DATABASE_NAME",
    "CARD_TYPE",
    "TYPE_OF_PAYMENT",
    "TERMS",
    "AUTHORIZATION_CODE",
    "USE_CASE",
    "UNIT_ORIG_PRICE_W_TAX",
    "UNIT_DISC_AMT_W_TAX",
    "UNIT_FINAL_PRICE_W_TAX",
]

STORE_TYPO_FIXES = {
    "ALTITUDE PIONER": "ALTITUDE PIONEER",
}

DIRECT_STORES = ["ECOFLOW", "OURA", "RAPTOR", "INSTA360"]

BRAND_RULES = {
    "ZENITH": [
        "OURA", "PLAUDAI", "BOOX", "COROS", "GARMIN",
        "LOOP", "SHOKZ", "RAYNEO", "RINGCONN", "POLAR",
    ],
    "ALTITUDE PIONEER": [
        "DJI", "INSTA360", "ECOFLOW", "RAPTOR", "YABER", "ENABOT",
    ],
    "ALTITUDE": [
        "ECOFLOW", "INSTA360", "DJI", "RAPTOR", "YABER", "ENABOT",
        "RINGCONN", "OURA", "PLAUDAI", "BOOX", "SHOKZ", "COROS",
    ],
    "UG": [
        "ECOFLOW", "INSTA360", "DJI", "ANKER", "YABER",
        "RINGCONN", "OURA", "PLAUDAI", "SHOKZ",
    ],
}


# =========================================================
# HELPERS
# =========================================================

def clean_text(series: pd.Series) -> pd.Series:
    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )

def create_channel(store: str) -> str:
    if store == "":
        return ""
    if store == "ALTITUDE PIONEER":
        return store
    return store.split()[0]

def create_category(store: str, brand: str) -> str:
    for name in DIRECT_STORES:
        if name in store:
            return name

    for store_key, allowed_brands in BRAND_RULES.items():
        if store_key in store:
            for item in allowed_brands:
                if item in brand:
                    return item
            return "OTHER BRANDS"

    return ""


def remove_sale_return_pairs(df: pd.DataFrame) -> pd.DataFrame:
    """Removes offsetting sale and return pairs on the same receipt for the same customer."""
    if not {"RECEIPT_NO", "CUSTOMER_NAME", "EXT_PRICE_W_TAX"}.issubset(df.columns):
        return df

    indices_to_drop = []
    group_cols = ["RECEIPT_NO", "CUSTOMER_NAME"]
    
    for _, group in df.groupby(group_cols, dropna=False):
        if len(group) < 2:
            continue
        
        prices = group["EXT_PRICE_W_TAX"].values
        index_labels = group.index.values
        dropped_in_group = set()
        
        for i in range(len(prices)):
            if i in dropped_in_group:
                continue
            for j in range(i + 1, len(prices)):
                if j in dropped_in_group:
                    continue
                # Check if values are exact opposites (e.g., 499.00 and -499.00)
                if np.isclose(prices[i], -prices[j], atol=1e-5) and prices[i] != 0:
                    dropped_in_group.add(i)
                    dropped_in_group.add(j)
                    indices_to_drop.extend([index_labels[i], index_labels[j]])
                    break
                    
    if indices_to_drop:
        df = df.drop(index=indices_to_drop).reset_index(drop=True)
    return df

# =========================================================
# MAIN PROCESSING FUNCTION (Called by GUI)
# =========================================================

def process_cadence(input_file_path):
    input_path = Path(input_file_path)

    # 1. Read file and skip top title/merged header row (header=1)
    if input_path.suffix.lower() in [".xls", ".xlsx"]:
        df = pd.read_excel(input_path, header=1)
    else:
        df = pd.read_csv(input_path, header=1)

    df.columns = df.columns.astype(str).str.strip().str.upper()

    # Drop ghost columns from merged cells
    df = df.loc[:, ~df.columns.str.contains('UNNAMED', case=False)]

    # Check required columns
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}")

    # Ensure numeric price before manipulations
    df["EXT_PRICE_W_TAX"] = pd.to_numeric(df["EXT_PRICE_W_TAX"], errors="coerce").fillna(0)

    # =========================================================
    # INTEREST / FINANCING ABSORPTION LOGIC
    # =========================================================
    # Identify interest rows where brand or description indicates interest charges
    interest_mask = (
        df["BRAND"].isin(["INTEREST"]) | 
        df["LONG_DESCRIPTION"].str.contains("INTEREST", case=False, na=False)
    )

    if interest_mask.any():
        # Isolate interest records and sum them per RECEIPT_NO and CUSTOMER_NAME
        interest_df = df[interest_mask].copy()
        interest_grouped = interest_df.groupby(["RECEIPT_NO", "CUSTOMER_NAME"], as_index=False)["EXT_PRICE_W_TAX"].sum()
        interest_grouped = interest_grouped.rename(columns={"EXT_PRICE_W_TAX": "TOTAL_INTEREST"})

        # Keep only non-interest rows as our main dataset
        df = df[~interest_mask].copy()

        # Merge interest back to the matching receipt/customer records
        df = df.merge(interest_grouped, on=["RECEIPT_NO", "CUSTOMER_NAME"], how="left")
        df["TOTAL_INTEREST"] = df["TOTAL_INTEREST"].fillna(0)

        # Distribute interest proportionally across items in the same receipt (safeguards multi-item receipts)
        receipt_totals = df.groupby(["RECEIPT_NO", "CUSTOMER_NAME"])["EXT_PRICE_W_TAX"].transform("sum")
        proportion = df["EXT_PRICE_W_TAX"] / receipt_totals.replace(0, 1)

        df["EXT_PRICE_W_TAX"] = df["EXT_PRICE_W_TAX"] + (df["TOTAL_INTEREST"] * proportion)
        df = df.drop(columns=["TOTAL_INTEREST"])

    # Drop unwanted columns after processing
    to_drop = [c for c in COLUMNS_TO_DROP if c in df.columns]
    df = df.drop(columns=to_drop)

    # Filter zero prices
    zero_mask = df["EXT_PRICE_W_TAX"] == 0
    df = df[~zero_mask].copy()

    # Remove offsetting sale-return pairs
    remove_sale_return_pairs(df)

    # Clean text fields
    df["STORE_NAME"] = clean_text(df["STORE_NAME"]).replace(STORE_TYPO_FIXES)
    df["BRAND"] = clean_text(df["BRAND"])

    # Channel & Category Assignment
    df["CHANNEL"] = df["STORE_NAME"].apply(create_channel)
    df["CATEGORY"] = [
        create_category(store, brand)
        for store, brand in zip(df["STORE_NAME"], df["BRAND"])
    ]

    tail = ["CATEGORY", "CHANNEL"]
    df = df[[c for c in df.columns if c not in tail] + tail]

    # 2. Setup Destination Path with Dynamic Date Filename
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    current_date = datetime.date.today().strftime("%Y-%m-%d")
    output_file = OUTPUT_DIR / f"cleaned_cadence_{current_date}.xlsx"

    # 3. Save to Excel normally first
    df.to_excel(output_file, index=False, engine='openpyxl')

    # 4. Convert Range into an Official Excel Table
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    
    max_row = ws.max_row
    max_col = ws.max_column
    col_letter = openpyxl.utils.get_column_letter(max_col)
    table_range = f"A1:{col_letter}{max_row}"

    tab = Table(displayName="CadenceCleanTable", ref=table_range)
    style = TableStyleInfo(
        name="TableStyleMedium2", 
        showFirstColumn=False, 
        showLastColumn=False, 
        showRowStripes=True, 
        showColumnStripes=False
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)
    
    wb.save(output_file)

    
    return len(df), str(output_file)


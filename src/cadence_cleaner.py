import pandas as pd
from pathlib import Path
import datetime
import openpyxl
import numpy as np
import re
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================================
# CONFIG
# =========================================================

BASE_FOLDER = Path(__file__).resolve().parent.parent

OUTPUT_DIR = Path.home() / "Desktop" / "Cleaned Data"

REQUIRED_COLUMNS = [
    "STORE_NAME",
    "BRAND",
    "EXT_PRICE_W_TAX",
    "RECEIPT_NO",
    "CUSTOMER_NAME",
]

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

DIRECT_STORES = [
    "ECOFLOW",
    "OURA",
    "RAPTOR",
    "INSTA360",
]

# EF TAG rules (ECOFLOW rows only). Checked in order; first match wins.
EF_TAG_RULES = [
    ("DELTA", "Delta Series"),
    ("RIVER", "River Series"),
    ("SOLAR PANEL", "Solar Panel"),
]
EF_TAG_DEFAULT = "EF Others"

# I360 TAG rules (INSTA360 rows only). Checked in order; first match wins.
I360_TAG_RULES = [
    ("X6", "X6"),
    ("X5", "X5"),
    ("GO ULTRA", "Go Ultra"),
    ("ACE PRO 2", "Ace Pro 2"),
    ("LUNA", "Luna"),
    ("GO 3S", "Go 3s"),
]
I360_TAG_DEFAULT = "I360 Others"

# EcoFlow accessories and cables (adaptor, adapter, cables, etc.) 
# are excluded from main series and assigned to "EF Others".
EF_ACCESSORY_TERMS = [
    "CABLE",
    "CABLES",
    "ADAPTER",
    "ADAPTOR",
    "CHARGER",
    "BAG",
    "COVER",
    "TROLLEY",
]

BRAND_RULES = {
    "ALTITUDE PIONEER": [
        "DJI",
        "INSTA360",
        "ECOFLOW",
        "RAPTOR",
        "YABER",
        "ENABOT",
    ],
    "ZENITH": [
        "OURA",
        "PLAUDAI",
        "BOOX",
        "COROS",
        "GARMIN",
        "LOOP",
        "SHOKZ",
        "RAYNEO",
        "RINGCONN",
        "POLAR",
    ],
    "ALTITUDE": [
        "ECOFLOW",
        "INSTA360",
        "DJI",
        "RAPTOR",
        "YABER",
        "ENABOT",
        "RINGCONN",
        "OURA",
        "PLAUDAI",
        "BOOX",
        "SHOKZ",
        "COROS",
        "RAYNEO",
    ],
    "UG": [
        "ECOFLOW",
        "INSTA360",
        "DJI",
        "YABER",
        "RINGCONN",
        "OURA",
        "PLAUDAI",
        "SHOKZ",
        "COROS",
        "RAYNEO",
        "BOOX",
        "RAPTOR",
        "ENABOT",
    ],
}


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def normalize_header(value):
    """Normalize an Excel/CSV header value[cite: 8]."""
    if pd.isna(value):
        return ""

    return (
        str(value)
        .strip()
        .upper()
        .replace("\n", " ")
    )


def clean_text(series: pd.Series) -> pd.Series:
    """Trim, uppercase, collapse repeated spaces, blank for NaN[cite: 8, 9]."""
    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )


def contains_term(text: str, term: str) -> bool:
    """Match a whole word or phrase rather than a partial word[cite: 8]."""
    pattern = (
        rf"(?<![A-Z0-9])"
        rf"{re.escape(term)}"
        rf"(?![A-Z0-9])"
    )

    return bool(re.search(pattern, text))


def create_channel(store: str) -> str:
    """Create CHANNEL from STORE_NAME[cite: 8, 9]."""
    if not store:
        return ""

    if store.startswith("ALTITUDE PIONEER"):
        return "ALTITUDE PIONEER"

    return store.split()[0]


def create_category(store: str, brand: str) -> str:
    """Assign CATEGORY using direct-store and brand rules[cite: 8, 9]."""
    if not store:
        return ""

    for name in DIRECT_STORES:
        if contains_term(store, name):
            return name

    for store_key, allowed_brands in BRAND_RULES.items():
        if contains_term(store, store_key):
            if not brand:
                return "OTHER BRANDS"

            for item in allowed_brands:
                if contains_term(brand, item):
                    return item

            return "OTHER BRANDS"

    return ""


def create_ef_tag(category: str, long_description: str) -> str:
    """Tag EcoFlow products by series, ensuring batteries roll up into Delta/River series[cite: 8]."""
    if category != "ECOFLOW":
        return ""

    # Accessories and cable terms are excluded from main series tags
    if any(contains_term(long_description, t) for t in EF_ACCESSORY_TERMS):
        return EF_TAG_DEFAULT

    for term, tag in EF_TAG_RULES:
        if contains_term(long_description, term):
            return tag

    return EF_TAG_DEFAULT


def create_i360_tag(category: str, long_description: str) -> str:
    """Tag Insta360 products by model. Blank for non-Insta360 rows."""
    if category != "INSTA360":
        return ""

    for term, tag in I360_TAG_RULES:
        if contains_term(long_description, term):
            return tag

    return I360_TAG_DEFAULT


def create_product_tag(category: str, ef_tag: str, i360_tag: str) -> str:
    """Combines EF TAG and I360 TAG into a single column to prevent pivot table blank nesting."""
    if category == "ECOFLOW":
        return ef_tag
    elif category == "INSTA360":
        return i360_tag
    return ""


def fix_store_typo(store: str) -> str:
    """Correct known store-name typos, including names with suffixes[cite: 8]."""
    for wrong, correct in STORE_TYPO_FIXES.items():
        if store == wrong or store.startswith(wrong + " "):
            return correct + store[len(wrong):]

    return store


def remove_sale_return_pairs(df: pd.DataFrame) -> pd.DataFrame:
    """Removes offsetting sale and return pairs sharing customer_name, receipt_number, or general_notes[cite: 9]."""
    if "EXT_PRICE_W_TAX" not in df.columns:
        return df

    cols_to_check = [c for c in ["RECEIPT_NO", "CUSTOMER_NAME", "GENERAL_NOTES"] if c in df.columns]
    if not cols_to_check:
        return df

    indices_to_drop = set()

    for col in cols_to_check:
        for _, group in df.groupby(col, dropna=False):
            if len(group) < 2:
                continue
            key_val = group[col].iloc[0]
            if pd.isna(key_val) or str(key_val).strip() == "":
                continue

            prices = group["EXT_PRICE_W_TAX"].values
            index_labels = group.index.values
            dropped_in_subgroup = set()

            for i in range(len(prices)):
                if index_labels[i] in indices_to_drop or i in dropped_in_subgroup:
                    continue
                for j in range(i + 1, len(prices)):
                    if index_labels[j] in indices_to_drop or j in dropped_in_subgroup:
                        continue
                    if np.isclose(prices[i], -prices[j], atol=1e-5) and prices[i] != 0:
                        dropped_in_subgroup.add(i)
                        dropped_in_subgroup.add(j)
                        indices_to_drop.add(index_labels[i])
                        indices_to_drop.add(index_labels[j])
                        break

    if indices_to_drop:
        df = df.drop(index=list(indices_to_drop)).reset_index(drop=True)
    return df


# =========================================================
# MAIN PROCESSING FUNCTION (Called by GUI)
# =========================================================

def process_cadence(input_file_path):
    input_path = Path(input_file_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found:\n{input_path}")

    # Preview first 100 rows for automatic header detection
    if input_path.suffix.lower() in [".xls", ".xlsx"]:
        preview = pd.read_excel(input_path, header=None, nrows=100)
    else:
        preview = pd.read_csv(input_path, header=None, nrows=100)

    required = set(REQUIRED_COLUMNS)
    header_row = None

    for row_index, row in preview.iterrows():
        row_values = {
            normalize_header(value)
            for value in row.tolist()
        }
        if required.issubset(row_values):
            header_row = row_index
            break

    if header_row is None:
        raise ValueError(
            f"Could not find all required columns in the first 100 rows:\n{REQUIRED_COLUMNS}"
        )

    # Load file using detected header
    if input_path.suffix.lower() in [".xls", ".xlsx"]:
        df = pd.read_excel(input_path, header=header_row)
    else:
        df = pd.read_csv(input_path, header=header_row)

    # Normalize column names
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )

    df = df.loc[:, ~df.columns.str.contains('UNNAMED', case=False)]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}")

    df["EXT_PRICE_W_TAX"] = pd.to_numeric(df["EXT_PRICE_W_TAX"], errors="coerce").fillna(0)

    # Interest handling logic from original cadence_cleaner.py
    interest_mask = (
        df["BRAND"].eq("INTEREST") | 
        clean_text(df["LONG_DESCRIPTION"]).str.contains("INTEREST", case=False, na=False)
    ) if "LONG_DESCRIPTION" in df.columns else df["BRAND"].eq("INTEREST")

    if interest_mask.any():
        interest_df = df[interest_mask].copy()
        interest_grouped = interest_df.groupby(["RECEIPT_NO", "CUSTOMER_NAME"], as_index=False)["EXT_PRICE_W_TAX"].sum()
        interest_grouped = interest_grouped.rename(columns={"EXT_PRICE_W_TAX": "TOTAL_INTEREST"})

        df = df[~interest_mask].copy()

        df = df.merge(interest_grouped, on=["RECEIPT_NO", "CUSTOMER_NAME"], how="left")
        df["TOTAL_INTEREST"] = df["TOTAL_INTEREST"].fillna(0)

        receipt_totals = df.groupby(["RECEIPT_NO", "CUSTOMER_NAME"])["EXT_PRICE_W_TAX"].transform("sum")
        proportion = df["EXT_PRICE_W_TAX"] / receipt_totals.replace(0, 1)

        df["EXT_PRICE_W_TAX"] = df["EXT_PRICE_W_TAX"] + (df["TOTAL_INTEREST"] * proportion)
        df = df.drop(columns=["TOTAL_INTEREST"])

    to_drop = [c for c in COLUMNS_TO_DROP if c in df.columns]
    df = df.drop(columns=to_drop)

    zero_mask = df["EXT_PRICE_W_TAX"] == 0
    df = df[~zero_mask].copy()

    df = remove_sale_return_pairs(df)

    df["STORE_NAME"] = clean_text(df["STORE_NAME"]).apply(fix_store_typo)
    df["BRAND"] = clean_text(df["BRAND"])

    has_long_desc = "LONG_DESCRIPTION" in df.columns
    if has_long_desc:
        long_desc_clean = clean_text(df["LONG_DESCRIPTION"])
    else:
        long_desc_clean = pd.Series("", index=df.index)

    df["CHANNEL"] = df["STORE_NAME"].apply(create_channel)
    df["CATEGORY"] = [
        create_category(store, brand)
        for store, brand in zip(df["STORE_NAME"], df["BRAND"])
    ]
    
    # Generate individual tags internally to build the unified product tag
    ef_tags = [
        create_ef_tag(category, desc)
        for category, desc in zip(df["CATEGORY"], long_desc_clean)
    ]
    i360_tags = [
        create_i360_tag(category, desc)
        for category, desc in zip(df["CATEGORY"], long_desc_clean)
    ]
    
    df["PRODUCT_TAG"] = [
        create_product_tag(cat, ef, i360)
        for cat, ef, i360 in zip(df["CATEGORY"], ef_tags, i360_tags)
    ]

    tail = ["CATEGORY", "CHANNEL", "PRODUCT_TAG"]
    df = df[[c for c in df.columns if c not in tail] + tail]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    current_date = datetime.date.today().strftime("%Y-%m-%d")
    output_file = OUTPUT_DIR / f"cleaned_cadence_{current_date}.xlsx"

    df.to_excel(output_file, index=False, engine='openpyxl')

    # Format as an official Excel Table named 'CadenceCleanTable'
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    ws.title = "Cleaned Data"
    
    max_row = ws.max_row
    max_column = ws.max_column
    col_letter = openpyxl.utils.get_column_letter(max_column)
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
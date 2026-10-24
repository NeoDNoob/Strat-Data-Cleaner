import pandas as pd
from pathlib import Path
import datetime
import openpyxl
import numpy as np
import re
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================================
# CONFIGURATION & CONSTANTS
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

EF_TAG_RULES = [
    ("DELTA", "Delta Series"),
    ("RIVER", "River Series"),
    ("SOLAR PANEL", "Solar Panel"),
]
EF_TAG_DEFAULT = "EF Others"

I360_TAG_RULES = [
    ("X6", "X6"),
    ("X5", "X5"),
    ("GO ULTRA", "Go Ultra"),
    ("ACE PRO 2", "Ace Pro 2"),
    ("LUNA", "Luna"),
    ("GO 3S", "Go 3s"),
]
I360_TAG_DEFAULT = "I360 Others"

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
    "ALTITUDE PIONEER": ["DJI", "INSTA360", "ECOFLOW", "RAPTOR", "YABER", "ENABOT"],
    "ZENITH": ["OURA", "PLAUDAI", "BOOX", "COROS", "GARMIN", "LOOP", "SHOKZ", "RAYNEO", "RINGCONN", "POLAR"],
    "ALTITUDE": ["ECOFLOW", "INSTA360", "DJI", "RAPTOR", "YABER", "ENABOT", "RINGCONN", "OURA", "PLAUDAI", "BOOX", "SHOKZ", "COROS", "RAYNEO"],
    "UG": ["ECOFLOW", "INSTA360", "DJI", "YABER", "RINGCONN", "OURA", "PLAUDAI", "SHOKZ", "COROS", "RAYNEO", "BOOX", "RAPTOR", "ENABOT"],
}


# =========================================================
# HELPER FUNCTIONS: TEXT & HEADER PROCESSING
# =========================================================

def normalize_header(value) -> str:
    """Normalize an Excel/CSV header value."""
    if pd.isna(value):
        return ""
    return str(value).strip().upper().replace("\n", " ")


def clean_text(series: pd.Series) -> pd.Series:
    """Trim, uppercase, collapse repeated spaces, blank for NaN."""
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )


def contains_term(text: str, term: str) -> bool:
    """Match a whole word or phrase rather than a partial word."""
    pattern = rf"(?<![A-Z0-9]){re.escape(term)}(?![A-Z0-9])"
    return bool(re.search(pattern, text))


def fix_store_typo(store: str) -> str:
    """Correct known store-name typos, including names with suffixes."""
    for wrong, correct in STORE_TYPO_FIXES.items():
        if store == wrong or store.startswith(wrong + " "):
            return correct + store[len(wrong):]
    return store


def read_data_file(file_path: Path, header=None, nrows=None) -> pd.DataFrame:
    """Helper to load Excel or CSV files dynamically based on extension."""
    if file_path.suffix.lower() in [".xls", ".xlsx"]:
        return pd.read_excel(file_path, header=header, nrows=nrows)
    return pd.read_csv(file_path, header=header, nrows=nrows)


# =========================================================
# HELPER FUNCTIONS: CATEGORIZATION & TAGGING
# =========================================================

def create_channel(store: str) -> str:
    """Create CHANNEL from STORE_NAME."""
    if not store:
        return ""
    if store.startswith("ALTITUDE PIONEER"):
        return "ALTITUDE PIONEER"
    return store.split()[0]


def create_category(store: str, brand: str) -> str:
    """Assign CATEGORY using direct-store and brand rules."""
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
    """Tag EcoFlow products by series, ensuring batteries roll up into Delta/River series."""
    if category != "ECOFLOW":
        return ""

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


def create_oura_tag(category: str, brand: str, long_description: str) -> str:
    """Tag Oura products by model."""
    cat, bnd, desc = str(category).upper(), str(brand).upper(), str(long_description).upper()
    
    if "OURA" in cat or "OURA" in bnd:
        if any(k in desc for k in ["RING 4", "RING4"]):
            return "OURA RING 4"
        elif any(k in desc for k in ["RING 5", "RING5"]):
            return "OURA RING 5"
        elif any(k in desc for k in ["CHARGER", "SIZING KIT"]):
            return "OURA ACCESSORIES"
        else:
            return "OURA OTHERS"
    return ""


def create_raptor_tag(category: str, brand: str, long_description: str) -> str:
    """Tag Raptor products by category/description."""
    cat, bnd, desc = str(category).upper(), str(brand).upper(), str(long_description).upper()
    
    if "RAPTOR" in cat or "RAPTOR" in bnd:
        if any(k in desc for k in ["TOOLBOX", "TOOL BOX"]):
            return "TOOLBOX"
        elif "COOLER" in desc:
            return "COOLERS"
        elif any(k in desc for k in ["DRINKWARE", "BOTTLE", "MUG", "TUMBLER", "FLASK", "CUP"]):
            return "DRINKWARE"
        elif any(k in desc for k in ["APPAREL", "SHIRT", "HOODIE", "HAT", "CAP", "TEE", "JACKET", "POLO", "SHORTS", "PANTS"]):
            return "APPAREL"
        elif any(k in desc for k in ["FOODWARE", "LUNCH", "CUTLERY", "PLATE", "BOWL", "SPORK", "MESS KIT"]):
            return "FOODWARE"
        elif any(k in desc for k in ["CAMPING", "OUTING", "TENT", "CHAIR", "TABLE", "LANTERN", "HAMMOCK"]):
            return "OUTING"
        elif any(k in desc for k in ["BAG", "BACKPACK", "POUCH", "SLING", "DUFFLE", "PACK", "CARRY", "SACK"]):
            return "BAGS"
        elif any(k in desc for k in ["STORAGE", "BOX", "CONTAINER", "BIN"]):
            return "STORAGES"
        elif any(k in desc for k in ["CASE", "HARDWARE"]):
            return "HARDWARE"
        else:
            return "RAPTOR OTHERS"
    return ""


def create_product_tag(category: str, ef_tag: str, i360_tag: str, oura_tag: str, raptor_tag: str) -> str:
    """Combines all brand tags into a single column to prevent pivot table blank nesting."""
    return {
        "ECOFLOW": ef_tag,
        "INSTA360": i360_tag,
        "OURA": oura_tag,
        "RAPTOR": raptor_tag,
    }.get(category, "")


# =========================================================
# HELPER FUNCTIONS: DATA CLEANING & FILTERING
# =========================================================

def clean_special_brands(df: pd.DataFrame) -> pd.DataFrame:
    """Applies specific brand renaming and filtering logic for '*', 'ZZ', and 'Plaud'."""
    if "BRAND" not in df.columns or "LONG_DESCRIPTION" not in df.columns:
        return df

    brand = df["BRAND"].fillna("").astype(str).str.strip()
    long_desc = df["LONG_DESCRIPTION"].fillna("").astype(str)

    # Akaso fix
    df.loc[(brand == "*") & long_desc.str.contains("Akaso", case=False, na=False), "BRAND"] = "Akaso"
    brand = df["BRAND"].fillna("").astype(str).str.strip()

    # Plaud fix
    df.loc[brand.str.upper() == "PLAUD", "BRAND"] = "PlaudAi"
    brand = df["BRAND"].fillna("").astype(str).str.strip()

    # ZZ brand rules
    mask_zz = brand == "ZZ"
    if mask_zz.any():
        has_paper_bag = long_desc.str.contains("Paper Bag", case=False, na=False)
        
        if "EXT_PRICE_W_TAX" in df.columns:
            zero_price = df["EXT_PRICE_W_TAX"].fillna(0) == 0
            df = df[~(mask_zz & has_paper_bag & zero_price)].reset_index(drop=True)
            
            brand = df["BRAND"].fillna("").astype(str).str.strip()
            long_desc = df["LONG_DESCRIPTION"].fillna("").astype(str)
            mask_zz = brand == "ZZ"
            has_paper_bag = long_desc.str.contains("Paper Bag", case=False, na=False)

        mask_rename = mask_zz & ~has_paper_bag
        if mask_rename.any():
            def get_first_word(desc: str) -> str:
                words = desc.strip().split()
                return words[0].capitalize() if words else "ZZ"
            df.loc[mask_rename, "BRAND"] = long_desc[mask_rename].apply(get_first_word)

    return df


def remove_negative_daily_stores(df: pd.DataFrame) -> pd.DataFrame:
    """Removes all entries for stores that have a net negative total EXT_PRICE_W_TAX for a given date."""
    if not all(col in df.columns for col in ["STORE_NAME", "RCPT_DATE", "EXT_PRICE_W_TAX"]):
        return df
    
    daily_totals = df.groupby(["STORE_NAME", "RCPT_DATE"], dropna=False)["EXT_PRICE_W_TAX"].sum()
    negative_store_dates = daily_totals[daily_totals < 0].index
    
    if len(negative_store_dates) > 0:
        mask = df.set_index(["STORE_NAME", "RCPT_DATE"]).index.isin(negative_store_dates)
        df = df[~mask].reset_index(drop=True)
        
    return df


def remove_sale_return_pairs(df: pd.DataFrame) -> pd.DataFrame:
    """Removes offsetting sale and return pairs sharing customer_name, receipt_number, or general_notes."""
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
                        dropped_in_subgroup.update({i, j})
                        indices_to_drop.update({index_labels[i], index_labels[j]})
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
    preview = read_data_file(input_path, header=None, nrows=100)

    required = set(REQUIRED_COLUMNS)
    header_row = None

    for row_index, row in preview.iterrows():
        row_values = {normalize_header(value) for value in row.tolist()}
        if required.issubset(row_values):
            header_row = row_index
            break

    if header_row is None:
        raise ValueError(f"Could not find all required columns in the first 100 rows:\n{REQUIRED_COLUMNS}")

    # Load full file using detected header
    df = read_data_file(input_path, header=header_row)

    # Normalize column names and drop unnamed columns
    df.columns = df.columns.astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
    df = df.loc[:, ~df.columns.str.contains('UNNAMED', case=False)]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}")

    df["EXT_PRICE_W_TAX"] = pd.to_numeric(df["EXT_PRICE_W_TAX"], errors="coerce").fillna(0)
    df = clean_special_brands(df)

    # Interest handling logic
    long_desc_clean = clean_text(df["LONG_DESCRIPTION"]) if "LONG_DESCRIPTION" in df.columns else pd.Series("", index=df.index)
    interest_mask = df["BRAND"].eq("INTEREST") | long_desc_clean.str.contains("INTEREST", case=False, na=False)

    if interest_mask.any():
        interest_df = df[interest_mask].copy()
        interest_grouped = interest_df.groupby(["RECEIPT_NO", "CUSTOMER_NAME"], as_index=False)["EXT_PRICE_W_TAX"].sum()
        interest_grouped = interest_grouped.rename(columns={"EXT_PRICE_W_TAX": "TOTAL_INTEREST"})

        df = df[~interest_mask].copy()
        df = df.merge(interest_grouped, on=["RECEIPT_NO", "CUSTOMER_NAME"], how="left")
        df["TOTAL_INTEREST"] = df["TOTAL_INTEREST"].fillna(0)

        receipt_totals = df.groupby(["RECEIPT_NO", "CUSTOMER_NAME"])["EXT_PRICE_W_TAX"].transform("sum")
        proportion = df["EXT_PRICE_W_TAX"] / receipt_totals.replace(0, 1)

        df["EXT_PRICE_W_TAX"] += df["TOTAL_INTEREST"] * proportion
        df = df.drop(columns=["TOTAL_INTEREST"])

    # Drop unneeded columns and zero-price rows
    to_drop = [c for c in COLUMNS_TO_DROP if c in df.columns]
    df = df.drop(columns=to_drop)
    df = df[df["EXT_PRICE_W_TAX"] != 0].copy()

    # Apply row-level filters and standardizations
    df = remove_sale_return_pairs(df)
    df = remove_negative_daily_stores(df)
    df["STORE_NAME"] = clean_text(df["STORE_NAME"]).apply(fix_store_typo)
    df["BRAND"] = clean_text(df["BRAND"])

    # Generate Channel, Category, and Tags
    df["CHANNEL"] = df["STORE_NAME"].apply(create_channel)
    df["CATEGORY"] = [create_category(s, b) for s, b in zip(df["STORE_NAME"], df["BRAND"])]
    
    ef_tags = [create_ef_tag(cat, desc) for cat, desc in zip(df["CATEGORY"], long_desc_clean)]
    i360_tags = [create_i360_tag(cat, desc) for cat, desc in zip(df["CATEGORY"], long_desc_clean)]
    oura_tags = [create_oura_tag(cat, bnd, desc) for cat, bnd, desc in zip(df["CATEGORY"], df["BRAND"], long_desc_clean)]
    raptor_tags = [create_raptor_tag(cat, bnd, desc) for cat, bnd, desc in zip(df["CATEGORY"], df["BRAND"], long_desc_clean)]
    
    df["PRODUCT_TAG"] = [
        create_product_tag(cat, ef, i360, oura, raptor)
        for cat, ef, i360, oura, raptor in zip(df["CATEGORY"], ef_tags, i360_tags, oura_tags, raptor_tags)
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

    # Reorder columns
    tail = ["CATEGORY", "CHANNEL", "PRODUCT_TAG"]
    df = df[[c for c in df.columns if c not in tail] + tail]

    # Output generation with collision handling
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yesterday_date = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    base_filename = f"cleaned_cadence_{yesterday_date}"
    output_file = OUTPUT_DIR / f"{base_filename}.xlsx"

    counter = 1
    while output_file.exists():
        output_file = OUTPUT_DIR / f"{base_filename} ({counter}).xlsx"
        counter += 1

    df.to_excel(output_file, index=False, engine='openpyxl')

    # Format as an official Excel Table named 'CadenceCleanTable'
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    ws.title = "Cleaned Cadence Data"
    
    max_row = ws.max_row
    max_column = ws.max_column
    col_letter = openpyxl.utils.get_column_letter(max_column)
    table_range = f"A1:{col_letter}{max_row}"

    tab = Table(displayName="CadenceCleanTable", ref=table_range)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", 
        showFirstColumn=False, 
        showLastColumn=False, 
        showRowStripes=True, 
        showColumnStripes=False
    )
    ws.add_table(tab)
    
    wb.save(output_file)

    return len(df), str(output_file)
import pandas as pd
from pathlib import Path
import datetime
import openpyxl
import numpy as np
import re
import difflib
from itertools import combinations
from typing import Any, Optional
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================================
# CONFIGURATION & CONSTANTS
# =========================================================

BASE_FOLDER = Path(__file__).resolve().parent.parent
OUTPUT_DIR = Path.home() / "Desktop" / "Cleaned Data"

# --- Sale/return pair offsetting ---
PAIR_MAX_SPLIT_LINES = 4        # max lines on the split side (e.g. main + accessory + freebie + 1)
PAIR_MAX_DAYS = 31              # max days between rows of one pair; None = no limit
PAIR_SAME_EVENT_DAYS = 3        # rows this close in date count as one exchange/re-ring event and are paired first
PAIR_TOLERANCE_CENTS = 1        # allowed rounding difference (1 = 0.01) between the two sides
PAIR_POOL_LIMIT = 20            # per customer, only the N closest-dated rows are combined (speed guard)
PAIR_ITEM_COLUMNS = ("DESCRIPTION_1", "ITEM_CODE", "LONG_DESCRIPTION")  # item identity used to confirm same-basket returns/re-sales
PAIR_GROUP_POOL_LIMIT = 10      # max lines per side considered for multi-line (N-to-M) basket matches
IGNORED_CUSTOMER_NAMES = {      # shared/placeholder customers: never paired (edit to match your data)
    "", "NAN", "NONE", "N/A", "NA", "UNKNOWN", "GUEST", "WALK-IN", "WALK IN", "WALKIN",
    "CASH", "CASH CUSTOMER", "POS CUSTOMER",
}

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

DIRECT_STORES = ["ECOFLOW", "OURA", "RAPTOR", "INSTA360"]

EF_TAG_DEFAULT = "EF Others"
I360_TAG_DEFAULT = "I360 Others"
RAPTOR_TAG_DEFAULT = "RAPTOR OTHERS"
OURA_TAG_DEFAULT = "OURA OTHERS"

BRAND_RULES = {
    "ALTITUDE PIONEER": ["DJI", "INSTA360", "ECOFLOW", "RAPTOR", "YABER", "ENABOT"],
    "ZENITH": ["RINGCONN","OURA","BOOX","COROS","PLAUDAI","GARMIN","SHOKZ","LOOP","POLAR","RAYNEO","EARQUIS"],
    "ALTITUDE": ["ECOFLOW", "INSTA360", "DJI","RINGCONN", "OURA", "BOOX", "COROS","PLAUDAI", "RAPTOR","SHOKZ","YABER","FUJIFILM", "RAYNEO","ENABOT","POLAR","GARMIN","OCCULI","EARQUIS","VULCROX","SONY"],
    "UG": ["ECOFLOW", "INSTA360", "DJI", "RINGCONN","ANKER","OURA","SHOKZ","YABER","GARMIN","PLAUDAI","RAYNEO","BOOX","COROS","POLAR","RAPTOR","OCCULLI","EARQUIS","SONY"],
}

# =========================================================
# HELPER FUNCTIONS: TEXT & HEADER PROCESSING
# =========================================================

def normalize_header(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper().replace("\n", " ")

def clean_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\s+", " ", regex=True)
    )

def contains_term(text: str, term: str) -> bool:
    pattern = rf"(?<![A-Z0-9]){re.escape(term)}(?![A-Z0-9])"
    return bool(re.search(pattern, text))

def fix_store_typos_vectorized(series: pd.Series) -> pd.Series:
    res = series.copy()
    for wrong, correct in STORE_TYPO_FIXES.items():
        mask = res.str.startswith(wrong)
        if mask.any():
            res.loc[mask] = res.loc[mask].str.replace(wrong, correct, regex=False)
    return res

def read_data_file(file_path: Path, header=None, nrows=None) -> pd.DataFrame:
    if file_path.suffix.lower() in [".xls", ".xlsx"]:
        return pd.read_excel(file_path, header=header, nrows=nrows)
    return pd.read_csv(file_path, header=header, nrows=nrows)

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_FULL = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
MONTH_LOOKUP = {m.upper(): i for i, m in enumerate(MONTH_ABBR, start=1)}

RE_DMY_TEXT = re.compile(r"^(\d{1,2})[-/. ]([A-Za-z]{3,9})[-/. ,]+(\d{2}|\d{4})$")
RE_ISO = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$")
RE_SLASH = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4}|\d{2})$")
RE_SERIAL_TEXT = re.compile(r"^\d{5}(?:\.\d+)?$")
RE_TIME_SUFFIX = re.compile(r"[T\s]+\d{1,2}:\d{2}.*$")

def parse_date_value(value: Any) -> Optional[pd.Timestamp]:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    if isinstance(value, (datetime.date)):
        ts = pd.Timestamp(value)
        return ts.tz_localize(None).normalize() if ts.tzinfo else ts.normalize()
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        if 1 <= float(value) < 2958466:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))).normalize()
        return None
    text = RE_TIME_SUFFIX.sub("", str(value).strip()).strip()  # keep the date only, drop any time-of-day
    if RE_SERIAL_TEXT.match(text):
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(text))).normalize()
    match = RE_DMY_TEXT.match(text)
    if match:
        day, month_word, year_text = match.groups()
        word = month_word.upper()
        month = MONTH_LOOKUP.get(word[:3])
        if month is None or not MONTH_FULL[month - 1].startswith(word):
            return None
        year = int(year_text)
        year = 2000 + year if len(year_text) <= 2 else year
        try:
            return pd.Timestamp(year=year, month=month, day=int(day))
        except ValueError:
            return None
    match = RE_ISO.match(text)
    if match:
        try:
            return pd.Timestamp(year=int(match.group(1)), month=int(match.group(2)), day=int(match.group(3)))
        except ValueError:
            return None
    match = RE_SLASH.match(text)
    if match:
        first, second, year_text = int(match.group(1)), int(match.group(2)), match.group(3)
        year = int(year_text)
        year = 2000 + year if len(year_text) <= 2 else year
        try:
            return pd.Timestamp(year=year, month=second, day=first)
        except ValueError:
            return None
    return None

def to_date_only(value: Any) -> Optional[pd.Timestamp]:
    """Receipt date without time-of-day. Falls back to pandas for formats parse_date_value does not know."""
    ts = parse_date_value(value)
    if ts is None and value is not None and not (not isinstance(value, str) and pd.isna(value)) and str(value).strip():
        fallback = pd.to_datetime(RE_TIME_SUFFIX.sub("", str(value).strip()), errors="coerce", dayfirst=True)
        ts = fallback.normalize() if pd.notna(fallback) else None
    return ts

def dates_for_excel(frame: pd.DataFrame) -> pd.DataFrame:
    """Return a copy whose date columns are formatted as MM/DD/YYYY strings for Excel output."""
    frame = frame.copy()
    for col in frame.columns:
        if col == "RCPT_DATE" or pd.api.types.is_datetime64_any_dtype(frame[col]):
            frame[col] = pd.to_datetime(frame[col], errors="coerce").dt.strftime("%m/%d/%Y")
    return frame

def fuzzy_normalize_customers(df: pd.DataFrame, customer_col="CUSTOMER_NAME", threshold=0.85) -> pd.DataFrame:
    """Standardizes customer names with minor typos using fuzzy matching."""
    if customer_col not in df.columns:
        return df
    
    df[customer_col] = df[customer_col].fillna("").astype(str).str.strip().str.upper()
    unique_names = [name for name in df[customer_col].unique() if name and name != "NAN" and name != ""]
    
    name_mapping = {}
    canonical_list = []
    
    for name in unique_names:
        if not canonical_list:
            canonical_list.append(name)
            name_mapping[name] = name
            continue
            
        matches = difflib.get_close_matches(name, canonical_list, n=1, cutoff=threshold)
        if matches:
            name_mapping[name] = matches[0]
        else:
            canonical_list.append(name)
            name_mapping[name] = name
            
    df[customer_col] = df[customer_col].map(lambda x: name_mapping.get(x, x))
    return df

def roll_up_accessories(df: pd.DataFrame, price_col="EXT_PRICE_W_TAX") -> pd.DataFrame:
    """Rolls up adapter, socket, or accessory values into the main high-value unit for the customer."""
    if "CUSTOMER_NAME" not in df.columns or "LONG_DESCRIPTION" not in df.columns:
        return df
        
    desc_clean = df["LONG_DESCRIPTION"].fillna("").astype(str).str.upper()
    
    is_accessory = desc_clean.str.contains("ADAPTER|ADAPTOR|ACCESSORY|CABLES|CABLE|POUCH|CHARGER|SOCKET", regex=True, na=False)
    if not is_accessory.any():
        return df
        
    accessories_df = df[is_accessory]
    main_df = df[~is_accessory].copy()
    
    rows_to_drop = []
    for idx, acc_row in accessories_df.iterrows():
        cust = acc_row["CUSTOMER_NAME"]
        acc_val = acc_row.get(price_col, 0)
        
        if not cust or cust == "NAN" or cust == "" or pd.isna(acc_val) or acc_val < 0:
            continue
            
        cust_mask = (main_df["CUSTOMER_NAME"] == cust) & (main_df[price_col] > 0)
        if cust_mask.any():
            target_idx = main_df.loc[cust_mask, price_col].idxmax()
            main_df.loc[target_idx, price_col] += acc_val
            rows_to_drop.append(idx)
            
    df_rolled = pd.concat([main_df, accessories_df.drop(index=rows_to_drop)], ignore_index=True) if rows_to_drop else df
    return df_rolled.reset_index(drop=True)

# =========================================================
# MONOBRAND TAGGING LOGIC
# =========================================================

def create_raptor_tag(desc: str) -> str:
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
    elif any(k in desc for k in ["CASE", "LITE", "EXTREME", "CARRY"]):
        return "HARDWARE"
    elif any(k in desc for k in ["BAG", "BACKPACK", "POUCH", "SLING", "DUFFLE", "PACK", "CARRY", "SACK"]):
        return "BAGS"
    elif any(k in desc for k in ["STORAGE", "BOX", "CONTAINER", "BIN"]):
        return "STORAGES"
    return RAPTOR_TAG_DEFAULT

def create_ef_tag(desc: str) -> str:
    accessory_terms = ["CABLE", "CABLES", "ADAPTOR", "ADAPTORS", "ADAPTER", "ADAPTERS", "FLASHLIGHT", "FLASHLIGHTS"]
    if any(term in desc for term in accessory_terms):
        return EF_TAG_DEFAULT
    if "SOLAR" in desc:
        return "Solar Panel"
    elif "DELTA" in desc:
        return "Delta Series"
    elif "RIVER" in desc:
        return "River Series"
    return EF_TAG_DEFAULT

def create_i360_tag(desc: str) -> str:
    if "BUNDLE" not in desc and "LIMITED EDITION" not in desc:
        return I360_TAG_DEFAULT
    if "ACE PRO 2" in desc:
        return "Ace Pro 2"
    elif "GO ULTRA" in desc:
        return "Go Ultra"
    elif "GO 3S" in desc:
        return "Go 3s"
    elif "LUNA" in desc:
        return "Luna"
    elif "X5" in desc:
        return "X5"
    elif "X6" in desc:
        return "X6"
    return I360_TAG_DEFAULT

def create_oura_tag(desc: str) -> str:
    if any(k in desc for k in ["CHARGER", "PROTECTOR", "SIZING KIT", "SIZING", "ACCESSORY", "ACCESSORIES"]):
        return "OURA ACCESSORIES"
    elif any(k in desc for k in ["RING 4", "RING4"]):
        return "OURA RING 4"
    elif any(k in desc for k in ["RING 5", "RING5"]):
        return "OURA RING 5"
    return OURA_TAG_DEFAULT

def resolve_product_tag(brand: str, long_description: str) -> str:
    bnd = str(brand).strip().upper()
    desc = str(long_description).upper()
    if bnd == "RAPTOR":
        return create_raptor_tag(desc)
    elif bnd == "ECOFLOW":
        return create_ef_tag(desc)
    elif bnd == "INSTA360":
        return create_i360_tag(desc)
    elif bnd == "OURA":
        return create_oura_tag(desc)
    return ""

# =========================================================
# DATA CLEANING & OPTIMIZED FILTERING WITH AUDIT LOGGING
# =========================================================

def create_channel(store: str) -> str:
    if not store:
        return ""
    if store.startswith("ALTITUDE PIONEER"):
        return "ALTITUDE PIONEER"
    return store.split()[0]

def create_category(store: str, brand: str) -> str:
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

def clean_special_brands(df: pd.DataFrame, audit_log: list) -> pd.DataFrame:
    if "BRAND" not in df.columns or "LONG_DESCRIPTION" not in df.columns:
        return df
    brand = df["BRAND"].fillna("").astype(str).str.strip()
    long_desc = df["LONG_DESCRIPTION"].fillna("").astype(str)

    jinko_raptor_mask = (
        brand.str.upper().eq("JINKO")
        & long_desc.str.contains("Raptor", case=False, na=False)
    )
    df.loc[jinko_raptor_mask, "BRAND"] = "RAPTOR"

    df.loc[(brand == "*") & long_desc.str.contains("Akaso", case=False, na=False), "BRAND"] = "Akaso"
    df.loc[brand.str.upper() == "PLAUD", "BRAND"] = "PlaudAi"
    
    brand = df["BRAND"].fillna("").astype(str).str.strip()
    mask_zz = brand == "ZZ"
    if mask_zz.any():
        has_paper_bag = long_desc.str.contains("Paper Bag", case=False, na=False)
        if "EXT_PRICE_W_TAX" in df.columns:
            zero_price = df["EXT_PRICE_W_TAX"].fillna(0) == 0
            to_remove_mask = mask_zz & has_paper_bag & zero_price
            if to_remove_mask.any():
                removed_rows = df[to_remove_mask].copy()
                removed_rows["REMOVAL_REASON"] = "ZZ Brand Paper Bag with Zero Price"
                audit_log.append(removed_rows)
                df = df[~to_remove_mask].reset_index(drop=True)

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

def _date_span_days(rows: list, date_list: list) -> int:
    """Days between the earliest and latest receipt date among `rows` (0 if dates are unknown)."""
    stamps = [date_list[i] for i in rows if pd.notna(date_list[i])]
    return (max(stamps) - min(stamps)).days if len(stamps) >= 2 else 0


def _nearest_rows(pool: list, anchor: int, date_list: list) -> list:
    """Rows from `pool` closest in date to `anchor`, inside the matching window and capped in size."""
    anchor_date = date_list[anchor]
    scored = []
    for i in pool:
        gap = 0
        if pd.notna(anchor_date) and pd.notna(date_list[i]):
            gap = abs((date_list[i] - anchor_date).days)
        if PAIR_MAX_DAYS is not None and gap > PAIR_MAX_DAYS:
            continue
        scored.append((gap, i))
    scored.sort()
    return [i for _, i in scored[:PAIR_POOL_LIMIT]]


def _clean_item_key(value) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    if isinstance(value, (float, np.floating)) and float(value).is_integer():
        value = int(value)
    return " ".join(str(value).split()).upper()


def _item_keys_and_qtys(df: pd.DataFrame) -> tuple:
    """Per-row item identity (first non-empty of PAIR_ITEM_COLUMNS) and absolute quantity."""
    cols = [c for c in PAIR_ITEM_COLUMNS if c in df.columns]
    if cols:
        per_col = [[_clean_item_key(v) for v in df[c].tolist()] for c in cols]
        keys = [next((k for k in vals if k), "") for vals in zip(*per_col)]
    else:
        keys = [""] * len(df)
    if "SOLD_QTY" in df.columns:
        qtys = pd.to_numeric(df["SOLD_QTY"], errors="coerce").abs().round().fillna(1).astype(int).clip(lower=1).tolist()
    else:
        qtys = [1] * len(df)
    return keys, qtys


def _signature(rows, item_keys: list, qtys: list) -> tuple:
    """What a set of lines contains: {item: total quantity}, order-independent."""
    contents: dict = {}
    for i in rows:
        contents[item_keys[i]] = contents.get(item_keys[i], 0) + qtys[i]
    return tuple(sorted(contents.items()))


def _find_pair_candidates(neg: list, pos: list, cents: list, date_list: list, item_keys: list, qtys: list) -> list:
    """
    Builds every possible (returns <-> sales) match for one customer where the amounts cancel out:
      - 1 sale line   <-> 1..N return lines  (basket returned as main + accessory/freebie lines,
                                              repurchased as a single full-price line)
      - 1 return line <-> 2..N sale lines    (the reverse split)
      - 2..N returns  <-> 2..N sales         (same items and quantities returned and re-sold, with the
                                              price re-allocated between the lines; totals must cancel)
    Row ORDER in the file is deliberately ignored: a return may be listed before its sale.
    """
    found = []

    def consider(neg_rows: tuple, pos_rows: tuple) -> None:
        members = list(neg_rows) + list(pos_rows)
        span = _date_span_days(members, date_list)
        if PAIR_MAX_DAYS is not None and span > PAIR_MAX_DAYS:
            return
        found.append((span, len(members), min(members), neg_rows, pos_rows))

    for p in pos:
        pool = _nearest_rows(neg, p, date_list)
        for k in range(1, PAIR_MAX_SPLIT_LINES + 1):
            for combo in combinations(pool, k):
                if abs(sum(cents[i] for i in combo) + cents[p]) <= PAIR_TOLERANCE_CENTS:
                    consider(combo, (p,))

    for n in neg:
        pool = _nearest_rows(pos, n, date_list)
        for k in range(2, PAIR_MAX_SPLIT_LINES + 1):
            for combo in combinations(pool, k):
                if abs(sum(cents[i] for i in combo) + cents[n]) <= PAIR_TOLERANCE_CENTS:
                    consider((n,), combo)

    # Multi-line baskets: identical items/quantities on both sides, totals cancel, per-line prices may differ.
    shared = {item_keys[i] for i in neg if item_keys[i]} & {item_keys[j] for j in pos if item_keys[j]}
    neg_pool = [i for i in neg if item_keys[i] in shared][:PAIR_GROUP_POOL_LIMIT]
    pos_pool = [j for j in pos if item_keys[j] in shared][:PAIR_GROUP_POOL_LIMIT]
    if len(neg_pool) >= 2 and len(pos_pool) >= 2:
        for kn in range(2, min(PAIR_MAX_SPLIT_LINES, len(neg_pool)) + 1):
            for neg_combo in combinations(neg_pool, kn):
                contents = _signature(neg_combo, item_keys, qtys)
                wanted = {item for item, _ in contents}
                neg_total = sum(cents[i] for i in neg_combo)
                pool = [j for j in pos_pool if item_keys[j] in wanted]
                for kp in range(2, min(PAIR_MAX_SPLIT_LINES, len(pool)) + 1):
                    for pos_combo in combinations(pool, kp):
                        if (abs(neg_total + sum(cents[j] for j in pos_combo)) <= PAIR_TOLERANCE_CENTS
                                and _signature(pos_combo, item_keys, qtys) == contents):
                            consider(neg_combo, pos_combo)

    return found


def _pair_priority(span: int, size: int, first_row: int) -> tuple:
    if span <= PAIR_SAME_EVENT_DAYS:
        return (0, span, size, first_row)   # same exchange event: tightest in time first
    return (1, size, span, first_row)       # otherwise: simplest line-for-line match first


def remove_customer_sale_return_pairs(df: pd.DataFrame, audit_log: list, price_col="EXT_PRICE_W_TAX") -> pd.DataFrame:
    """
    Offsets returns against the sale that cancels them, per customer, so only unmatched
    (genuine) returns remain.

    - Handles split baskets: several return lines (main + accessory + freebie) that add up to one
      repurchased line, the reverse, and the same items returned and re-sold with the price
      re-allocated between lines (multi-line to multi-line, confirmed by item code and quantity).
    - Independent of row order (return-sale-sale or sale-return-sale give the same result) and of
      the associate / receipt number.
    - Several separate pairs per customer are matched independently (not "all returns vs one sale").
    - Return + repurchase within PAIR_SAME_EVENT_DAYS are paired first (exchange / re-ring);
      after that the simplest line-for-line matches win. A row is used at most once.
    - Every removed row goes to the audit log with a PAIR_ID so each match can be reviewed.
    """
    if price_col not in df.columns or "CUSTOMER_NAME" not in df.columns:
        return df

    df = df.reset_index(drop=True)
    cents = (pd.to_numeric(df[price_col], errors="coerce").fillna(0) * 100).round().astype("int64").tolist()
    if not any(c < 0 for c in cents):
        return df

    if "RCPT_DATE" in df.columns:
        date_list = pd.to_datetime(df["RCPT_DATE"], errors="coerce").tolist()
    else:
        date_list = [pd.NaT] * len(df)
    item_keys, qtys = _item_keys_and_qtys(df)

    customers = df["CUSTOMER_NAME"].fillna("").astype(str).str.strip().str.upper()
    chosen = []

    for cust, rows in customers.groupby(customers, sort=False).indices.items():
        if cust in IGNORED_CUSTOMER_NAMES:
            continue
        rows = [int(i) for i in rows]
        neg = [i for i in rows if cents[i] < 0]
        pos = [i for i in rows if cents[i] > 0]
        if not neg or not pos:
            continue

        candidates = _find_pair_candidates(neg, pos, cents, date_list, item_keys, qtys)
        candidates.sort(key=lambda c: _pair_priority(c[0], c[1], c[2]))

        used: set = set()
        for span, _, _, neg_rows, pos_rows in candidates:
            members = set(neg_rows) | set(pos_rows)
            if members & used:
                continue
            used |= members
            chosen.append((cust, span, neg_rows, pos_rows))

    if not chosen:
        return df

    drop_positions = []
    for pair_id, (cust, span, neg_rows, pos_rows) in enumerate(chosen, start=1):
        members = list(neg_rows) + list(pos_rows)
        block = df.iloc[members].copy()
        block["REMOVAL_REASON"] = [
            f"Customer-Level Matched Return Set ({cust})" if cents[i] < 0
            else f"Customer-Level Matched Sale Counterpart ({cust})"
            for i in members
        ]
        block["PAIR_ID"] = pair_id
        block["PAIR_DATE_SPAN_DAYS"] = span
        audit_log.append(block)
        drop_positions.extend(members)

    return df.drop(index=drop_positions).reset_index(drop=True)


# =========================================================
# MAIN PROCESSING FUNCTION
# =========================================================

def process_cadence(input_file_path):
    input_path = Path(input_file_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found:\n{input_path}")

    audit_log = []

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

    df = read_data_file(input_path, header=header_row)
    df.columns = df.columns.astype(str).str.strip().str.upper().str.replace(r"\s+", " ", regex=True)
    df = df.loc[:, ~df.columns.str.contains('UNNAMED', case=False)]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}")

    # Convert to numeric early
    df["EXT_PRICE_W_TAX"] = pd.to_numeric(
        df["EXT_PRICE_W_TAX"].astype(str).str.replace(",", "").str.strip(), 
        errors="coerce"
    ).fillna(0)
    
    if "SOLD_QTY" in df.columns:
        df["SOLD_QTY"] = pd.to_numeric(
            df["SOLD_QTY"].astype(str).str.replace(",", "").str.strip(), 
            errors="coerce"
        ).fillna(0)

    # Receipt dates: calendar date only (no time-of-day), so day-level grouping and pair matching are exact
    if "RCPT_DATE" in df.columns:
        df["RCPT_DATE"] = pd.to_datetime([parse_date_value(v) for v in df["RCPT_DATE"]], errors="coerce")

    # Run preprocessing steps
    df = fuzzy_normalize_customers(df, "CUSTOMER_NAME")
    df = clean_special_brands(df, audit_log)

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

    to_drop = [c for c in COLUMNS_TO_DROP if c in df.columns]
    df = df.drop(columns=to_drop)

    # Log and filter zero-price rows
    zero_price_mask = df["EXT_PRICE_W_TAX"] == 0
    if zero_price_mask.any():
        zero_df = df[zero_price_mask].copy()
        zero_df["REMOVAL_REASON"] = "Zero Sale Amount / Price"
        audit_log.append(zero_df)
        df = df[~zero_price_mask].copy().reset_index(drop=True)

    df = remove_customer_sale_return_pairs(df, audit_log, "EXT_PRICE_W_TAX")

    # Accessories are rolled up AFTER pairing so split baskets stay matchable line by line,
    # and leftover (genuine) returns are not swallowed into a sale.
    df = roll_up_accessories(df, "EXT_PRICE_W_TAX")
    df["STORE_NAME"] = fix_store_typos_vectorized(clean_text(df["STORE_NAME"]))
    df["BRAND"] = clean_text(df["BRAND"])

    df["CHANNEL"] = df["STORE_NAME"].apply(create_channel)
    df["CATEGORY"] = [create_category(s, b) for s, b in zip(df["STORE_NAME"], df["BRAND"])]
    
    current_long_desc = clean_text(df["LONG_DESCRIPTION"]) if "LONG_DESCRIPTION" in df.columns else pd.Series("", index=df.index)
    df["PRODUCT_TAG"] = [
        resolve_product_tag(bnd, desc)
        for bnd, desc in zip(df["BRAND"], current_long_desc)
    ]
    
    tail = ["CATEGORY", "CHANNEL", "PRODUCT_TAG"]
    df = df[[c for c in df.columns if c not in tail] + tail]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yesterday_date = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    base_filename = f"cleaned_cadence_{yesterday_date}"
    output_file = OUTPUT_DIR / f"{base_filename}.xlsx"

    counter = 1
    while output_file.exists():
        output_file = OUTPUT_DIR / f"{base_filename} ({counter}).xlsx"
        counter += 1

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        dates_for_excel(df).to_excel(writer, sheet_name="Cleaned Cadence Data", index=False)
        if audit_log:
            audit_df = pd.concat(audit_log, ignore_index=True)
            front = ["REMOVAL_REASON"] + [c for c in ("PAIR_ID", "PAIR_DATE_SPAN_DAYS") if c in audit_df.columns]
            for c in ("PAIR_ID", "PAIR_DATE_SPAN_DAYS"):
                if c in audit_df.columns:
                    audit_df[c] = audit_df[c].astype("Int64")
            cols = front + [c for c in audit_df.columns if c not in front]
            audit_df = audit_df[cols]
            dates_for_excel(audit_df).to_excel(writer, sheet_name="Removal Audit Log", index=False)
        else:
            pd.DataFrame(columns=["REMOVAL_REASON", "NOTE"]).to_excel(writer, sheet_name="Removal Audit Log", index=False)

    wb = openpyxl.load_workbook(output_file)
    ws_main = wb["Cleaned Cadence Data"]
    table_range = f"A1:{openpyxl.utils.get_column_letter(ws_main.max_column)}{ws_main.max_row}"
    tab_main = Table(displayName="CadenceCleanTable", ref=table_range)
    tab_main.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_main.add_table(tab_main)

    ws_audit = wb["Removal Audit Log"]
    audit_table_range = f"A1:{openpyxl.utils.get_column_letter(ws_audit.max_column)}{ws_audit.max_row}"
    tab_audit = Table(displayName="RemovalAuditTable", ref=audit_table_range)
    tab_audit.tableStyleInfo = TableStyleInfo(name="TableStyleMedium7", showRowStripes=True)
    ws_audit.add_table(tab_audit)

    wb.save(output_file)
    return len(df), str(output_file)
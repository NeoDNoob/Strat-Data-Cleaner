from __future__ import annotations

import datetime
import re
from pathlib import Path
from typing import Any, Optional

import numpy as np
import pandas as pd
import openpyxl
import difflib
from itertools import combinations
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================================
# CONFIG & MAPPINGS
# =========================================================

OUTPUT_DIR = Path.home() / "Desktop" / "Cleaned Data"

HEADER_SCAN_ROWS = 30
MIN_HEADER_MATCHES = 5

EXPECTED_COLUMNS = [
    "DATABASE_NAME", "STORE_CODE", "STORE_NAME", "RCPT_DATE", "RECEIPT_NO",
    "CUSTOMER_NAME", "MOBILE_NUMBER", "ADDRESS_1", "TRANSACTION_TYPE", "SERIAL_NO",
    "DESCRIPTION_1", "BRAND", "LONG_DESCRIPTION", "SOLD_QTY", "TENDER_TYPE",
    "CARD_TYPE", "TYPE_OF_PAYMENT", "TERMS", "AUTHORIZATION_CODE", "USE_CASE", "TAG",
    "ASSOCIATE", "UNIT_ORIG_PRICE_W_TAX", "UNIT_DISC_AMT_W_TAX",
    "UNIT_FINAL_PRICE_W_TAX", "EXT_PRICE_W_TAX", "GENERAL_NOTES",
]

COLUMN_ALIASES = {"UNIT_ORGI_PRICE_W_TAX": "UNIT_ORIG_PRICE_W_TAX"}

COLUMNS_TO_REMOVE = [
    "DATABASE_NAME", "MOBILE_NUMBER", "SERIAL_NO", "CARD_TYPE", "AUTHORIZATION_CODE",
    "USE_CASE", "UNIT_ORIG_PRICE_W_TAX", "UNIT_DISC_AMT_W_TAX", "UNIT_FINAL_PRICE_W_TAX",
    "TERMS", "TYPE_OF_PAYMENT",
]

REQUIRED_COLUMNS = ["STORE_NAME", "RCPT_DATE", "SOLD_QTY", "EXT_PRICE_W_TAX"]
TEXT_ID_COLUMNS = ["STORE_CODE", "RECEIPT_NO", "DESCRIPTION_1"]
NUMERIC_COLUMNS = ["SOLD_QTY", "EXT_PRICE_W_TAX"]

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

STORE_NAME_MAP = {
    "ALTITUDE ALABANG TOWN CENTER": "Altitude Alabang Town Center",
    "ALTITUDE FESTIVAL MALL": "Altitude Festival Mall",
    "ALTITUDE GATEWAY MALL": "Altitude Gateway Mall",
    "ALTITUDE GLORIETTA": "Altitude Glorietta",
    "ALTITUDE GREENBELT 5": "Altitude Greenbelt",
    "ALTITUDE MEGAMALL": "Altitude Megamall",
    "ALTITUDE PIONEER": "Altitude Pioneer",
    "ALTITUDE RPM": "Altitude Robinsons Place Manila",
    "ALTITUDE TRINOMA": "Altitude Trinoma",
    "ECOFLOW SM CITY CALAMBA": "Ecoflow SM City Calamba",
    "ECOFLOW SM CITY CLARK": "Ecoflow SM City Clark",
    "ECOFLOW SM CITY FAIRVIEW": "Ecoflow SM City Fairview",
    "ECOFLOW SM CITY TAYTAY": "Ecoflow SM City Taytay",
    "ECOFLOW SM DASMA": "Ecoflow SM Dasma",
    "ECOFLOW SM GRAND CENTRAL": "Ecoflow SM Grand Central",
    "ECOFLOW SM MALL OF ASIA": "Ecoflow SM Mall Of Asia",
    "ECOFLOW SM MARILAO": "Ecoflow SM Marilao",
    "ECOFLOW SM MEGAMALL": "Ecoflow SM Megamall",
    "ECOFLOW SM NORTH EDSA": "Ecoflow SM North Edsa",
    "ECOFLOW SM PAMPANGA": "Ecoflow SM Pampanga",
    "INSTA360 FESTIVAL MALL": "Insta360 Festival Mall",
    "INSTA360 ROBINSON MANILA": "Insta360 Robinson Manila",
    "INSTA360 SM MALL OF ASIA": "Insta360 SM Mall Of Asia",
    "INSTA360 SM MEGAMALL": "Insta360 SM Megamall",
    "INSTA360 SM NORTH EDSA": "Insta360 SM North Edsa",
    "OURA GLORIETTA 1": "Oura Glorietta",
    "OURARW": "Oura Rockwell",
    "OURA SM AURA": "Oura SM Aura",
    "OURA SM MOA": "Oura SM Mall Of Asia",
    "OURA SM NORTH EDSA": "Oura SM North Edsa",
    "RAPTOR FESTIVAL MALL": "Raptor Festival Mall",
    "RAPTOR GATEWAY": "Raptor Gateway",
    "RAPTOR GRAND CENTRAL": "Raptor Grand Central",
    "RAPTOR SM NORTH EDSA": "Raptor SM North Edsa",
    "UG ALABANG": "Urban Gadgets Alabang Town Center",
    "UG BGC": "Urban Gadgets Bonifacio Global City",
    "UG GREENHILLS": "Urban Gadgets Greenhills",
    "UG MOA": "Urban Gadgets Mall Of Asia",
    "UG VERTISNORTH": "Urban Gadgets Vertis North",
    "ZENITH EASTWOOD": "Zenith Eastwood",
    "ZENITH GLORIETTA 4": "Zenith Glorietta",
    "ZENITH PIONEER": "Zenith Pioneer",
    "ZENITH SHANGRILA": "Zenith Shangrila",
}

CHANNEL_PREFIXES = {
    "Altitude": "Altitude",
    "Insta360": "Insta360",
    "EcoFlow": "EcoFlow",
    "Oura": "Oura",
    "Raptor": "Raptor",
    "Urban Gadgets": "Urban Gadgets",
    "Zenith": "Zenith",
}

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_FULL = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
MONTH_LOOKUP = {m.upper(): i for i, m in enumerate(MONTH_ABBR, start=1)}

ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
RE_DMY_TEXT = re.compile(r"^(\d{1,2})[-/. ]([A-Za-z]{3,9})[-/. ,]+(\d{2}|\d{4})$")
RE_ISO = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$")
RE_SLASH = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4}|\d{2})$")
RE_SERIAL_TEXT = re.compile(r"^\d{5}(?:\.\d+)?$")
RE_TIME_SUFFIX = re.compile(r"[T\s]+\d{1,2}:\d{2}.*$")

# =========================================================
# HELPERS
# =========================================================

def normalize_col_name(name: Any) -> str:
    return re.sub(r"[\s_]+", "_", str(name).strip()).strip("_").upper()

def canonical_col_name(name: Any) -> str:
    key = normalize_col_name(name)
    return COLUMN_ALIASES.get(key, key)

def normalize_store_key(name: Any) -> str:
    return " ".join(str(name).split()).upper()

def clean_cell(value: Any) -> Any:
    if isinstance(value, str):
        value = ILLEGAL_XML_CHARS.sub("", value).strip()
        return value if value != "" else None
    return value

def to_text_id(value: Any) -> Any:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    if isinstance(value, (bool, np.bool_)):
        return str(value)
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        return str(int(value)) if float(value).is_integer() else str(value)
    return str(value).strip()

def clean_associate(value: Any) -> Optional[str]:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    return " ".join(word.capitalize() for word in text.split())

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

def find_header_row(raw: pd.DataFrame) -> Optional[int]:
    known = set(EXPECTED_COLUMNS) | set(COLUMN_ALIASES)
    best_idx, best_score = None, 0
    for i in range(min(HEADER_SCAN_ROWS, len(raw))):
        score = sum(1 for v in raw.iloc[i].tolist() if isinstance(v, str) and normalize_col_name(v) in known)
        if score > best_score:
            best_idx, best_score = i, score
    return best_idx if best_score >= MIN_HEADER_MATCHES else None

def prepare_sheet(raw: pd.DataFrame, sheet_name: str) -> Optional[pd.DataFrame]:
    header_idx = find_header_row(raw)
    if header_idx is None:
        return None
    names = [
        canonical_col_name(v) if isinstance(v, str) and v.strip() else f"UNNAMED_{i + 1}"
        for i, v in enumerate(raw.iloc[header_idx].tolist())
    ]
    data = raw.iloc[header_idx + 1:].copy()
    data.columns = names
    return data.reset_index(drop=True)

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

def clean_special_brands(df: pd.DataFrame, audit_log: list) -> pd.DataFrame:
    if "BRAND" not in df.columns or "LONG_DESCRIPTION" not in df.columns:
        return df

    brand = df["BRAND"].fillna("").astype(str).str.strip()
    long_desc = df["LONG_DESCRIPTION"].fillna("").astype(str)

    mask_jinko = brand.str.upper() == "JINKO"
    if mask_jinko.any():
        def first_word_as_brand(desc: str) -> str:
            words = str(desc).strip().split()
            return words[0].upper() if words else "JINKO"
        df.loc[mask_jinko, "BRAND"] = long_desc[mask_jinko].apply(first_word_as_brand)

    brand = df["BRAND"].fillna("").astype(str).str.strip()
    mask_akaso = (brand == "*") & long_desc.str.contains("Akaso", case=False, na=False)
    df.loc[mask_akaso, "BRAND"] = "Akaso"
    
    brand = df["BRAND"].fillna("").astype(str).str.strip()
    mask_plaud = brand.str.upper() == "PLAUD"
    df.loc[mask_plaud, "BRAND"] = "PlaudAi"
    
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
# MAIN PROCESSOR
# =========================================================

def process_database(input_file_path: str) -> tuple[int, str]:
    input_path = Path(input_file_path)
    audit_log = []

    if input_path.suffix.lower() in [".xls", ".xlsx"]:
        frames = pd.read_excel(input_path, sheet_name=None, header=None, dtype=object, engine="openpyxl", keep_default_na=False, na_values=[""])
    else:
        with open(input_path, newline="", encoding="utf-8-sig", errors="replace") as f:
            rows = list(__import__('csv').reader(f))
        width = max((len(r) for r in rows), default=0)
        padded = [r + [""] * (width - len(r)) for r in rows]
        frames = {"csv": pd.DataFrame(padded, dtype=object)}

    prepared = [prepare_sheet(raw, sname) for sname, raw in frames.items()]
    prepared = [p for p in prepared if p is not None]
    if not prepared:
        raise ValueError("No valid header row found in the database export file.")

    df = pd.concat(prepared, ignore_index=True) if len(prepared) > 1 else prepared[0]

    data_cols = [c for c in df.columns if not c.startswith("UNNAMED_")]
    for col in data_cols:
        df[col] = df[col].map(clean_cell)
    df = df.dropna(how="all", subset=data_cols).reset_index(drop=True)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}")

    to_drop = [c for c in COLUMNS_TO_REMOVE if c in df.columns]
    df = df.drop(columns=to_drop)

    # Convert numeric columns early so price comparisons don't throw TypeErrors
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            stripped = df[col].map(lambda v: v.replace(",", "").strip() if isinstance(v, str) else v)
            df[col] = pd.to_numeric(stripped, errors="coerce").fillna(0)

    df = fuzzy_normalize_customers(df, "CUSTOMER_NAME")
    df = clean_special_brands(df, audit_log)

    lookup = {normalize_store_key(k): v for k, v in STORE_NAME_MAP.items()}
    standard = {normalize_store_key(v): v for v in STORE_NAME_MAP.values()}
    original_stores = df["STORE_NAME"]
    keys = original_stores.map(lambda v: normalize_store_key(v) if isinstance(v, str) else None)
    mapped = keys.map(lookup).where(lambda s: s.notna(), keys.map(standard))
    df["STORE_NAME"] = mapped.where(mapped.notna(), original_stores)

    def channel_of(name: Any) -> Optional[str]:
        if not isinstance(name, str):
            return None
        key = normalize_store_key(name)
        return next((label for prefix, label in CHANNEL_PREFIXES.items() if key.startswith(normalize_store_key(prefix))), None)

    df.insert(0, "CHANNEL", df["STORE_NAME"].map(channel_of))

    df["RCPT_DATE"] = pd.to_datetime([parse_date_value(v) for v in df["RCPT_DATE"]], errors="coerce")
    
    dates = df["RCPT_DATE"]
    month_no = dates.dt.month.astype("Int64")
    month_name = month_no.map(lambda m: MONTH_ABBR[int(m) - 1] if pd.notna(m) else None)
    iso_weeks = dates.dt.isocalendar().week
    week_no = pd.Series([f"W{int(w):02d}" if pd.notna(w) else None for w in iso_weeks], index=df.index)
    quarter_year = pd.Series([f"Q{d.quarter} {d.year}" if pd.notna(d) else None for d in dates], index=df.index)

    pos = df.columns.get_loc("RCPT_DATE") + 1
    df.insert(pos, "MONTH_NAME", month_name)
    df.insert(pos + 1, "MONTH_NO", month_no)
    df.insert(pos + 2, "WEEK_NO", week_no)
    df.insert(pos + 3, "QUARTER_YEAR", quarter_year)

    for col in TEXT_ID_COLUMNS:
        if col in df.columns:
            df[col] = df[col].map(to_text_id)

    if "ASSOCIATE" in df.columns:
        df["ASSOCIATE"] = df["ASSOCIATE"].map(clean_associate)

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            stripped = df[col].map(lambda v: v.replace(",", "").strip() if isinstance(v, str) else v)
            df[col] = pd.to_numeric(stripped, errors="coerce")

    df = clean_special_brands(df, audit_log)

    interest_mask = (
        df["BRAND"].str.contains("INTEREST", case=False, na=False) |
        df["LONG_DESCRIPTION"].str.contains("INTEREST", case=False, na=False) |
        df["DESCRIPTION_1"].str.contains("INTEREST", case=False, na=False)
    )

    if interest_mask.any() and "RECEIPT_NO" in df.columns and "CUSTOMER_NAME" in df.columns:
        keys = ["RECEIPT_NO", "CUSTOMER_NAME"]
        df["CUSTOMER_NAME"] = df["CUSTOMER_NAME"].fillna("")

        interest = df[interest_mask]
        main = df[~interest_mask].copy()

        totals = interest.groupby(keys)["EXT_PRICE_W_TAX"].sum()
        amounts = main["EXT_PRICE_W_TAX"].fillna(0)
        target = amounts.groupby([main["RECEIPT_NO"], main["CUSTOMER_NAME"]]).idxmax()

        matched = totals.index.intersection(target.index)
        rows = target.loc[matched].values
        main.loc[rows, "EXT_PRICE_W_TAX"] = (
            main.loc[rows, "EXT_PRICE_W_TAX"].fillna(0).values + totals.loc[matched].values
        )

        is_matched = interest.set_index(keys).index.isin(matched)
        df = pd.concat([main, interest[~is_matched]], ignore_index=True)

    df = remove_customer_sale_return_pairs(df, audit_log, "EXT_PRICE_W_TAX")

    # Accessories are rolled up AFTER pairing, so split baskets are still matchable line by line
    # and leftover (genuine) returns are not swallowed into a sale.
    df = roll_up_accessories(df, "EXT_PRICE_W_TAX")

    zero_price_mask = df["EXT_PRICE_W_TAX"] == 0
    if zero_price_mask.any():
        zero_df = df[zero_price_mask].copy()
        zero_df["REMOVAL_REASON"] = "Zero Sale Amount / Price"
        audit_log.append(zero_df)
        df = df[~zero_price_mask].copy()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yesterday_date = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    base_filename = f"cleaned_database_{yesterday_date}"
    output_file = OUTPUT_DIR / f"{base_filename}.xlsx"
    
    counter = 1
    while output_file.exists():
        output_file = OUTPUT_DIR / f"{base_filename} ({counter}).xlsx"
        counter += 1

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        dates_for_excel(df).to_excel(writer, sheet_name="Cleaned Database Data", index=False)
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
    ws_main = wb["Cleaned Database Data"]
    table_range = f"A1:{get_column_letter(ws_main.max_column)}{ws_main.max_row}"
    tab_main = Table(displayName="DatabaseCleanTable", ref=table_range)
    tab_main.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_main.add_table(tab_main)

    ws_audit = wb["Removal Audit Log"]
    audit_table_range = f"A1:{get_column_letter(ws_audit.max_column)}{ws_audit.max_row}"
    tab_audit = Table(displayName="RemovalAuditTable", ref=audit_table_range)
    tab_audit.tableStyleInfo = TableStyleInfo(name="TableStyleMedium7", showRowStripes=True)
    ws_audit.add_table(tab_audit)

    wb.save(output_file)
    return len(df), str(output_file)
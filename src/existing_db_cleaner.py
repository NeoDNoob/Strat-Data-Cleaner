from __future__ import annotations

import datetime
import re
from pathlib import Path
from typing import Any, Optional

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================================
# CONFIG & MAPPINGS
# =========================================================

OUTPUT_DIR = Path.home() / "Desktop" / "Cleaned Data"

HEADER_SCAN_ROWS = 50
MIN_HEADER_MATCHES = 4

EXPECTED_COLUMNS = [
    "CHANNEL", "STORE_CODE", "STORE_NAME", "RCPT_DATE", "MONTH_NAME", 
    "MONTH_NO", "WEEK_NO", "QUARTER_YEAR", "RECEIPT_NO", "CUSTOMER_NAME", 
    "ADDRESS", "TRANSACTION_TYPE", "ITEM_CODE", "BRAND", "LONG_DESCRIPTION", 
    "SOLD_QTY", "TENDER_TYPE", "TAG", "ASSOCIATE", "SALE_AMOUNT", 
    "GENERAL_NOTES", "EMPLOYMENT_STATUS", "ASSOCIATE_POSITION", 
    "ASSIGNMENT_CATEGORY", "SALES_DIVISION", "SPECIALIST"
]

REQUIRED_COLUMNS = ["STORE_NAME", "RCPT_DATE", "SOLD_QTY", "SALE_AMOUNT"]
TEXT_ID_COLUMNS = ["STORE_CODE", "RECEIPT_NO", "ITEM_CODE"]
NUMERIC_COLUMNS = ["SOLD_QTY", "SALE_AMOUNT"]

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_FULL = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
MONTH_LOOKUP = {m.upper(): i for i, m in enumerate(MONTH_ABBR, start=1)}

ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
RE_DMY_TEXT = re.compile(r"^(\d{1,2})[-/. ]([A-Za-z]{3,9})[-/. ,]+(\d{2}|\d{4})$")
RE_ISO = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$")
RE_SLASH = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4}|\d{2})$")
RE_SERIAL_TEXT = re.compile(r"^\d{5}(?:\.\d+)?$")

# =========================================================
# HELPERS
# =========================================================

def normalize_col_name(name: Any) -> str:
    return re.sub(r"[\s_]+", "_", str(name).strip()).strip("_").upper()

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

def extract_target_table_excel(input_path: Path, target_table_name: str = "DatabaseCleanTable") -> Optional[pd.DataFrame]:
    try:
        wb = openpyxl.load_workbook(input_path, data_only=True)
        for sheet in wb.worksheets:
            for table_name, table in sheet.tables.items():
                if table_name.lower() == target_table_name.lower():
                    data_range = sheet[table.ref]
                    rows = [[cell.value for cell in row] for row in data_range]
                    if not rows:
                        continue
                    headers = [normalize_col_name(h) for h in rows[0]]
                    df = pd.DataFrame(rows[1:], columns=headers)
                    return df
    except Exception as e:
        print(f"[Info] Excel openpyxl table scan bypassed: {e}")
    return None

def find_header_row(raw: pd.DataFrame) -> Optional[int]:
    known = set(EXPECTED_COLUMNS)
    best_idx, best_score = None, 0
    for i in range(min(HEADER_SCAN_ROWS, len(raw))):
        score = sum(1 for v in raw.iloc[i].tolist() if isinstance(v, str) and normalize_col_name(v) in known)
        if score > best_score:
            best_idx, best_score = i, score
    return best_idx if best_score >= MIN_HEADER_MATCHES else None

def prepare_sheet(raw: pd.DataFrame) -> Optional[pd.DataFrame]:
    header_idx = find_header_row(raw)
    if header_idx is None:
        return None
    names = [
        normalize_col_name(v) if isinstance(v, str) and v.strip() else f"UNNAMED_{i + 1}"
        for i, v in enumerate(raw.iloc[header_idx].tolist())
    ]
    data = raw.iloc[header_idx + 1:].copy()
    data.columns = names
    return data.reset_index(drop=True)

def parse_date_value(value: Any) -> Optional[pd.Timestamp]:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    if isinstance(value, (datetime.date, datetime.datetime, pd.Timestamp)):
        ts = pd.Timestamp(value)
        return ts.tz_localize(None).normalize() if ts.tzinfo else ts.normalize()
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        if 1 <= float(value) < 2958466:
            return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))).normalize()
        return None
    text = str(value).strip()
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
    mask_plaud = brand.str.upper().str.contains("PLAUD", na=False)
    df.loc[mask_plaud, "BRAND"] = "PlaudAi"
    
    brand = df["BRAND"].fillna("").astype(str).str.strip()
    mask_zz = brand == "ZZ"
    if mask_zz.any():
        has_paper_bag = long_desc.str.contains("Paper Bag", case=False, na=False)
        
        if "SALE_AMOUNT" in df.columns:
            zero_price = pd.to_numeric(df["SALE_AMOUNT"], errors="coerce").fillna(0) == 0
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

def remove_sale_return_pairs(df: pd.DataFrame, audit_log: list) -> pd.DataFrame:
    if "SALE_AMOUNT" not in df.columns:
        return df

    cols_to_check = [c for c in ["RECEIPT_NO", "CUSTOMER_NAME", "GENERAL_NOTES"] if c in df.columns]
    if not cols_to_check:
        return df

    if not (df["SALE_AMOUNT"] < 0).any():
        return df

    drop_reasons = {}

    for col in cols_to_check:
        groups_with_returns = df[df["SALE_AMOUNT"] < 0][col].dropna().unique()
        if len(groups_with_returns) == 0:
            continue

        relevant_df = df[df[col].isin(groups_with_returns)]

        for key_val, group in relevant_df.groupby(col, dropna=True):
            if str(key_val).strip() == "":
                continue

            prices = group["SALE_AMOUNT"].values
            index_labels = group.index.values

            neg_indices = np.where(prices < 0)[0]
            if len(neg_indices) == 0:
                continue

            pos_indices = np.where(prices > 0)[0]
            qtys = group["SOLD_QTY"].values if "SOLD_QTY" in group.columns else None
            brands = group["BRAND"].fillna("").astype(str).str.strip().values if "BRAND" in group.columns else None
            descs = group["LONG_DESCRIPTION"].fillna("").astype(str).str.strip().values if "LONG_DESCRIPTION" in group.columns else None

            matched_pos = set()

            for i in neg_indices:
                if index_labels[i] in drop_reasons:
                    continue
                for j in pos_indices:
                    if j in matched_pos or index_labels[j] in drop_reasons:
                        continue

                    brand_match = (brands is None) or (brands[i] == brands[j])
                    desc_match = (descs is None) or (descs[i] == descs[j])
                    price_match = abs(prices[i] + prices[j]) < 1e-5

                    qty_match = True
                    if qtys is not None:
                        qty_match = abs(qtys[i] + qtys[j]) < 1e-5

                    if brand_match and desc_match and price_match and qty_match:
                        matched_pos.add(j)
                        drop_reasons[index_labels[i]] = "Matched Sale-Return Pair (Return Row)"
                        drop_reasons[index_labels[j]] = "Matched Sale-Return Pair (Sale Counterpart)"
                        break

    if drop_reasons:
        indices_to_drop = list(drop_reasons.keys())
        removed_df = df.loc[indices_to_drop].copy()
        removed_df["REMOVAL_REASON"] = removed_df.index.map(drop_reasons)
        audit_log.append(removed_df)
        df = df.drop(index=indices_to_drop).reset_index(drop=True)

    return df

# =========================================================
# MAIN PROCESSOR
# =========================================================

def process_existing_database(input_file_path: str) -> tuple[int, str]:
    input_path = Path(input_file_path)
    audit_log = []

    df = None
    if input_path.suffix.lower() in [".xls", ".xlsx"]:
        df = extract_target_table_excel(input_path, "DatabaseCleanTable")

    if df is None:
        if input_path.suffix.lower() in [".xls", ".xlsx"]:
            frames = pd.read_excel(input_path, sheet_name=None, header=None, dtype=object, engine="openpyxl", keep_default_na=False, na_values=[""])
        else:
            with open(input_path, newline="", encoding="utf-8-sig", errors="replace") as f:
                rows = list(__import__('csv').reader(f))
            width = max((len(r) for r in rows), default=0)
            padded = [r + [""] * (width - len(r)) for r in rows]
            frames = {"csv": pd.DataFrame(padded, dtype=object)}

        prepared = [prepare_sheet(raw) for raw in frames.values()]
        prepared = [p for p in prepared if p is not None]
        if not prepared:
            raise ValueError("No valid DatabaseCleanTable header found across the file sheets.")

        df = pd.concat(prepared, ignore_index=True) if len(prepared) > 1 else prepared[0]

    data_cols = [c for c in df.columns if not str(c).startswith("UNNAMED_")]
    for col in data_cols:
        df[col] = df[col].map(clean_cell)
    df = df.dropna(how="all", subset=data_cols).reset_index(drop=True)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}")

    df = clean_special_brands(df, audit_log)

    if "RCPT_DATE" in df.columns:
        df["RCPT_DATE"] = pd.to_datetime([parse_date_value(v) for v in df["RCPT_DATE"]], errors="coerce")
        
        dates = df["RCPT_DATE"]
        month_no = dates.dt.month.astype("Int64")
        month_name = month_no.map(lambda m: MONTH_ABBR[int(m) - 1] if pd.notna(m) else None)
        iso_weeks = dates.dt.isocalendar().week
        week_no = pd.Series([f"W{int(w):02d}" if pd.notna(w) else None for w in iso_weeks], index=df.index)
        quarter_year = pd.Series([f"Q{d.quarter} {d.year}" if pd.notna(d) else None for d in dates], index=df.index)

        if "MONTH_NAME" in df.columns: df["MONTH_NAME"] = month_name
        if "MONTH_NO" in df.columns: df["MONTH_NO"] = month_no
        if "WEEK_NO" in df.columns: df["WEEK_NO"] = week_no
        if "QUARTER_YEAR" in df.columns: df["QUARTER_YEAR"] = quarter_year

    for col in TEXT_ID_COLUMNS:
        if col in df.columns:
            df[col] = df[col].map(to_text_id)

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            stripped = df[col].map(lambda v: v.replace(",", "").strip() if isinstance(v, str) else v)
            df[col] = pd.to_numeric(stripped, errors="coerce")

    df = clean_special_brands(df, audit_log)

    interest_mask = (
        df["BRAND"].str.contains("INTEREST", case=False, na=False) |
        df["LONG_DESCRIPTION"].str.contains("INTEREST", case=False, na=False)
    )

    if interest_mask.any() and "RECEIPT_NO" in df.columns and "CUSTOMER_NAME" in df.columns:
        keys = ["RECEIPT_NO", "CUSTOMER_NAME"]
        df["CUSTOMER_NAME"] = df["CUSTOMER_NAME"].fillna("")

        interest = df[interest_mask]
        main = df[~interest_mask].copy()

        totals = interest.groupby(keys)["SALE_AMOUNT"].sum()
        amounts = main["SALE_AMOUNT"].fillna(0)
        target = amounts.groupby([main["RECEIPT_NO"], main["CUSTOMER_NAME"]]).idxmax()

        matched = totals.index.intersection(target.index)
        rows = target.loc[matched].values
        main.loc[rows, "SALE_AMOUNT"] = (
            main.loc[rows, "SALE_AMOUNT"].fillna(0).values + totals.loc[matched].values
        )

        is_matched = interest.set_index(keys).index.isin(matched)
        df = pd.concat([main, interest[~is_matched]], ignore_index=True)

    df = remove_sale_return_pairs(df, audit_log)

    zero_price_mask = df["SALE_AMOUNT"] == 0
    if zero_price_mask.any():
        zero_df = df[zero_price_mask].copy()
        zero_df["REMOVAL_REASON"] = "Zero Sale Amount / Price"
        audit_log.append(zero_df)
        df = df[~zero_price_mask].copy()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    yesterday_date = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    base_filename = f"cleaned_existing_database_{yesterday_date}"
    output_file = OUTPUT_DIR / f"{base_filename}.xlsx"
    
    counter = 1
    while output_file.exists():
        output_file = OUTPUT_DIR / f"{base_filename} ({counter}).xlsx"
        counter += 1

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned Database Data", index=False)
        if audit_log:
            audit_df = pd.concat(audit_log, ignore_index=True)
            cols = ["REMOVAL_REASON"] + [c for c in audit_df.columns if c != "REMOVAL_REASON"]
            audit_df = audit_df[cols]
            audit_df.to_excel(writer, sheet_name="Removal Audit Log", index=False)
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
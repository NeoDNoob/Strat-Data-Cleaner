"""Meta Ads Cleaner for Strat Data Cleaner.

Entry point: process_meta_ads(file_path, retail_path, progress=None) -> (rows_out, output_path, stats)
Input : TWO files
        1. Supermetrics Meta Ads export (Date, Campaign name, Ad set name, Cost, Messaging conversations started)
        2. Retail sales export (the same file the other cleaners take). Used: STORE_NAME, RCPT_DATE, BRAND, TAG,
           RECEIPT_NO, TRANSACTION_TYPE, plus LONG_DESCRIPTION, EXT_PRICE_W_TAX, CUSTOMER_NAME and STORE_CODE for the
           brand masking and the interest-line matching (ported from the Brand Performance cleaner)
Output: Excel workbook in Desktop/Cleaned Data with four sheets: Meta Cleaned Data, Retail Data (Date, Store, Brand, Tag,
customer, EXT_PRICE_W_TAX and the distinct online transaction count on Online rows), Reconciliation (Meta Ads + Retail sections) and Audit
(removed Meta Ads rows and set-aside retail lines, told apart by a Source column).
The two datasets are NOT joined: Meta Ads and retail rows stay on their own sheets.
"""
import datetime
import re
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- configuration
RECOGNIZED_STORES = set("""OURA SMNE|ALTITUDE PARK TRIANGLE|RAPTOR UPTOWN MALL|ZENITH SHANGRI-LA|OURA SM MOA|UG UPTOWN MALL|OURA GLORIETTA|ZENITH EASTWOOD|ZENITH GLORIETTA|ALTITUDE CLARK|UG ALABANG|OURA SM AURA|UG BGC|INSTA360 ROBINSON MANILA|ALTITUDE GREENBELT 5|ALTITUDE RPM|UG VERTIS NORTH|INSTA360 FESTIVAL MALL|RAPTOR SM NORTH EDSA|ALTITUDE GLORIETTA|UG GREENHILLS|UG MALL OF ASIA|INSTA360 SM MALL OF ASIA|INSTA360 SM MEGAMALL|RAPTOR FESTIVAL MALL|INSTA360 SM NORTH EDSA|ALTITUDE TRINOMA|ALTITUDE PIONEER|RAPTOR GRAND CENTRAL|ALTITUDE MEGAMALL|ALTITUDE GATEWAY MALL|ECOFLOW SM PAMPANGA|ZENITH PIONEER|RAPTOR GATEWAY|ALTITUDE ALABANG TOWN CENTER|ALTITUDE FESTIVAL MALL|ECOFLOW SM MEGAMALL|ECOFLOW SM DASMA|ECOFLOW SM CITY CLARK|ECOFLOW SM MALL OF ASIA|ECOFLOW SM CITY CALAMBA|ECOFLOW SM MARILAO|ECOFLOW SM NORTH EDSA|ECOFLOW SM GRAND CENTRAL|ECOFLOW SM CITY FAIRVIEW|ECOFLOW SM CITY TAYTAY""".split('|'))

# (regex on lower-cased, whitespace-collapsed campaign name, store). First match wins.
STORE_RULES = [
    (r'^ars altitude - park triangle', 'ALTITUDE PARK TRIANGLE'),
    (r'^ars altitude - clark cityfront', 'ALTITUDE CLARK'),
    (r'^ars altitude - greenbelt 5', 'ALTITUDE GREENBELT 5'),
    (r'^ars altitude - glorietta 2', 'ALTITUDE GLORIETTA'),
    (r'^ars altitude - trinoma', 'ALTITUDE TRINOMA'),
    (r'^ars altitude - sm megamall', 'ALTITUDE MEGAMALL'),
    (r'^ars altitude - gateway mall 1', 'ALTITUDE GATEWAY MALL'),
    (r'^ars altitude - alabang town center', 'ALTITUDE ALABANG TOWN CENTER'),
    (r'^ars altitude - festival mall', 'ALTITUDE FESTIVAL MALL'),
    (r'^ars altitude - robinsons', 'ALTITUDE RPM'),
    (r'^ug ars - uptown mall', 'UG UPTOWN MALL'),
    (r'^ug ars - alabang town center', 'UG ALABANG'),
    (r'^ug ars - bhs', 'UG BGC'),
    (r'^ug ars - vertis north', 'UG VERTIS NORTH'),
    (r'^ug ars - greenhills', 'UG GREENHILLS'),
    (r'^ug ars - sm moa', 'UG MALL OF ASIA'),
    (r'^ug - pioneer', 'UG PIONEER'),                       # no such store on our end -> removed
    (r'^zenith ars - shangri', 'ZENITH SHANGRI-LA'),
    (r'^zenith ars - eastwood', 'ZENITH EASTWOOD'),
    (r'^zenith ars - glorietta 4', 'ZENITH GLORIETTA'),
    (r'^zenith - .*pioneer', 'ZENITH PIONEER'),
    (r'^oura ars - sm moa', 'OURA SM MOA'),
    (r'^oura ars - glorietta 1', 'OURA GLORIETTA'),
    (r'^oura ars - sm aura premier', 'OURA SM AURA'),
    (r'^oura ars - smne', 'OURA SMNE'),
    (r'^raptor - festival mall', 'RAPTOR FESTIVAL MALL'),
    (r'^raptor - gateway', 'RAPTOR GATEWAY'),
    (r'^raptor - sm grand central', 'RAPTOR GRAND CENTRAL'),
    (r'^raptor - sm north edsa', 'RAPTOR SM NORTH EDSA'),
    (r'^raptor - uptown bgc', 'RAPTOR UPTOWN MALL'),
    (r'^insta360 - altitude ho', 'ALTITUDE PIONEER'),
    (r'^insta360 - head office', 'ALTITUDE PIONEER'),
    (r'^insta360 - robinsons manila', 'INSTA360 ROBINSON MANILA'),
    (r'^insta360 - festival mall', 'INSTA360 FESTIVAL MALL'),
    (r'^insta360 - sm mall of asia', 'INSTA360 SM MALL OF ASIA'),
    (r'^insta360 - sm megamall', 'INSTA360 SM MEGAMALL'),
    (r'^insta360 - sm north edsa', 'INSTA360 SM NORTH EDSA'),
    (r'^ecoflow - altitude ho', 'ALTITUDE PIONEER'),
    (r'^ecoflow - head office', 'ALTITUDE PIONEER'),
    (r'^ecoflow pampanga', 'ECOFLOW SM PAMPANGA'),
    (r'^ecoflow - sm clark', 'ECOFLOW SM CITY CLARK'),
    (r'^ecoflow - sm dasmari', 'ECOFLOW SM DASMA'),
    (r'^ecoflow - sm calamba', 'ECOFLOW SM CITY CALAMBA'),
    (r'^ecoflow - sm marilao', 'ECOFLOW SM MARILAO'),
    (r'^ecoflow sm north edsa', 'ECOFLOW SM NORTH EDSA'),
    (r'^ecoflow - sm grand central', 'ECOFLOW SM GRAND CENTRAL'),
    (r'^ecoflow - sm fairview', 'ECOFLOW SM CITY FAIRVIEW'),
    (r'^ecoflow - sm taytay', 'ECOFLOW SM CITY TAYTAY'),
    (r'^ecoflow - sm mall of asia', 'ECOFLOW SM MALL OF ASIA'),
    (r'^ecoflow - sm megamall', 'ECOFLOW SM MEGAMALL'),
    (r'^altitude - ho', 'ALTITUDE PIONEER'),
    (r'^enabot page', 'ALTITUDE PIONEER'),
    (r'^yaber - messaging', 'ALTITUDE PIONEER'),
    (r'^ho - occulli', 'ALTITUDE PIONEER'),
    (r'^(boox|coros|earquis|loop|plaud|polar|ringconn|shokz) - pioneer', 'ZENITH PIONEER'),
    # "Promoting" / "Engagement" campaigns that name a store
    (r'promoting raptor - uptown mall bgc', 'RAPTOR UPTOWN MALL'),
    (r'promoting oura( philippines)? - glorietta 1', 'OURA GLORIETTA'),
    (r'promoting oura( philippines)? - sm aura', 'OURA SM AURA'),
    (r'promoting oura( philippines)? - sm mall of asia', 'OURA SM MOA'),
    (r'promoting oura( philippines)? - sm north edsa', 'OURA SMNE'),
    (r'^earquis - ho - engagement', 'ALTITUDE PIONEER'),
]
# Store campaigns we KNOW are outside the recognized list (so they do not raise an "unmapped" warning).
KNOWN_EXCLUDED_STORE_RULES = [
    r'^insta360 - sm city davao', r'^insta360 - yashano',
    r'^ecoflow - sm (bacolod|butuan|cdo|city (iloilo|lipa|lucena|puerto|yashano|zamboanga)|davao)',
    r'^ecoflow - southwoods', r'^ecoflow powerocean',
    r'promoting raptor - (glorietta|sm dasmari)', r'promoting ecoflow - sm city (mindpro|baguio|iloilo)',
]
NON_STORE_RULES = [r'powerocean']

ALTITUDE_PIONEER_BRANDS = {'DJI', 'Insta360', 'EcoFlow', 'Raptor', 'Yaber', 'Enabot', 'Occulli'}
EXCLUDED_BRANDS = {'Aero', 'Skydio', 'TekMage'}

BRAND_RULES = [
    ('EcoFlow', r'eco\s*flow|power\s*ocean|delta|river series|dpu'), ('Insta360', r'inst[a]?\s*360|inta360'),
    ('DJI', r'\bdji\b'), ('Raptor', r'raptor'), ('Oura', r'\boura\b'), ('RingConn', r'ring\s*conn'),
    ('PlaudAi', r'plaud'), ('Boox', r'\bboox\b'), ('Coros', r'coros'), ('Garmin', r'garmin'),
    ('Polar', r'\bpolar\b'), ('Shokz', r'shokz'), ('Loop', r'\bloop\b'), ('Earquis', r'earquis'),
    ('Occulli', r'occulli'), ('Enabot', r'enabot'), ('Yaber', r'yaber'), ('RayNeo', r'rayneo'),
    ('Skydio', r'skydio'), ('TekMage', r'tekmage'), ('Aero', r'\baero\b'), ('Vulcrox', r'vulcrox'),
]

REQUIRED = {'date': 'Date', 'campaign': 'Campaign name', 'adset': 'Ad set name', 'cost': 'Cost'}

# ---------------------------------------------------------------- helpers
def _norm(s):
    return re.sub(r'\s+', ' ', str(s)).strip().lower()


def _brand_of(s):
    s = _norm(s)
    for brand, pat in BRAND_RULES:
        if re.search(pat, s):
            return brand
    return None


def _store_of(campaign):
    c = _norm(campaign)
    for pat, store in STORE_RULES:
        if re.search(pat, c):
            return store
    return None


def _read_source(path, progress):
    ext = Path(path).suffix.lower()
    if ext in ('.csv', '.txt'):
        df = pd.read_csv(path, sep=None, engine='python')
    else:
        sheets = pd.read_excel(path, sheet_name=None)
        df = None
        for name, sh in sheets.items():
            cols = {str(c).strip().lower() for c in sh.columns}
            if {'date', 'campaign name', 'ad set name', 'cost'} <= cols:
                df = sh
                progress(f'Reading sheet "{name}"')
                break
        if df is None:
            raise ValueError('No sheet has the expected columns: Date, Campaign name, Ad set name, Cost, '
                             'Messaging conversations started. Is this a Meta Ads (Supermetrics) export?')
    df.columns = [str(c).strip() for c in df.columns]
    lower = {c.lower(): c for c in df.columns}
    missing = [v for v in REQUIRED.values() if v.lower() not in lower]
    if missing:
        raise ValueError(f'Missing column(s): {", ".join(missing)}. Is this a Meta Ads (Supermetrics) export?')
    conv_col = next((c for c in df.columns if 'conversation' in c.lower()), None)
    if conv_col is None:
        raise ValueError('Missing the "Messaging conversations started" column.')
    out = pd.DataFrame({
        'Date': pd.to_datetime(df[lower['date']], errors='coerce'),
        'Campaign name': df[lower['campaign name']].astype(str),
        'Ad set name': df[lower['ad set name']].astype(str),
        'Cost': pd.to_numeric(df[lower['cost']], errors='coerce').fillna(0.0),
        'Conversations': pd.to_numeric(df[conv_col], errors='coerce').fillna(0.0),
    })
    return out


def _campaign_type(name):
    n = _norm(name)
    if 'messaging' in n:
        return 'messaging'
    if 'promoting' in n or 'engagement' in n:
        return 'promo'
    return None


def _removal_reason(r):
    if r.Brand_camp in EXCLUDED_BRANDS or r.Brand_adset in EXCLUDED_BRANDS:
        b = r.Brand_camp if r.Brand_camp in EXCLUDED_BRANDS else r.Brand_adset
        return f'Excluded brand ({b})'
    ctype = _campaign_type(r['Campaign name'])
    if ctype is None:
        return 'Non-messaging campaign'
    if ctype == 'promo' and r.Camp_conv <= 0:
        return 'Promoting/Engagement campaign with no conversations'
    if pd.isna(r.Store):
        c = _norm(r['Campaign name'])
        if any(re.search(p, c) for p in NON_STORE_RULES):
            return 'Non-store campaign'
        if any(re.search(p, c) for p in KNOWN_EXCLUDED_STORE_RULES):
            return 'Store not in recognized list'
        return 'Non-store campaign' if ctype == 'promo' else 'Store not in recognized list'
    if r.Store not in RECOGNIZED_STORES:
        return 'Store not in recognized list'
    if pd.isna(r.Brand):
        return 'Multi-brand / no specific brand'
    if r.Store == 'ALTITUDE PIONEER' and r.Brand not in ALTITUDE_PIONEER_BRANDS:
        return 'Brand not allowed in Altitude Pioneer'
    return None


# ---------------------------------------------------------------- retail sales (second file)
META_HEADER = {'date', 'campaign name', 'ad set name', 'cost'}
RETAIL_HEADER = {'store_name', 'rcpt_date', 'receipt_no', 'brand', 'tag'}
HEADER_SCAN_ROWS = 15            # the retail export has a title line ("Exported data") above its headers

# --- Ported from brand_performance_cleaner.py (keep in sync) ---
STORE_TYPO_FIXES = {"ALTITUDE PIONER": "ALTITUDE PIONEER"}
# Retail-only store name fixes: the WHOLE name must match (ignoring case and extra spaces), so nothing else is touched.
STORE_EXACT_FIXES = {
    "OURA GLORIETTA 1": "OURA GLORIETTA",
    "OURA SM NORTH EDSA": "OURA SMNE",
    "UG MOA": "UG MALL OF ASIA",
    "UG VERTISNORTH": "UG VERTIS NORTH",
    "ZENITH GLORIETTA 4": "ZENITH GLORIETTA",
    "ZENITH SHANGRILA": "ZENITH SHANGRI-LA",
}
# Store masking: every row of a masked store moves to its target store (matched on the start of the name,
# ignoring case and extra spaces). Other stores are never touched.
STORE_MASK = {
    "RAP FESTIVAL MALL": "RAPTOR FESTIVAL MALL",
    "RAP GATEWAY MALL": "RAPTOR GATEWAY",
    "RAP SM NORTH EDSA": "RAPTOR SM NORTH EDSA",
}
# Shared/placeholder customers: their interest lines are only matched by receipt or same store + same day.
IGNORED_CUSTOMER_NAMES = {"", "NAN", "NONE", "N/A", "NA", "UNKNOWN", "GUEST", "WALK-IN", "WALK IN", "WALKIN",
                          "CASH", "CASH CUSTOMER", "POS CUSTOMER"}
INTEREST_MATCH_LABELS = {
    1: "Receipt + Customer",
    2: "Store + Customer + Same Date",
    3: "Store + Customer, Nearest Date",
    4: "Customer Only, Nearest Date (Other Store)",
}
# Brands that exist in the retail data (Brand Performance list) but are not Meta ad brands: spelled consistently,
# and not reported as a brand mismatch.
RETAIL_ONLY_BRANDS = {b.upper(): b for b in ("Akaso", "Fujifilm", "Sony", "Anker")}


def _read_raw(path, nrows=None):
    """Every sheet of an Excel file (or the CSV) as a header-less frame of raw cells."""
    ext = Path(path).suffix.lower()
    if ext in ('.csv', '.txt'):
        return {'file': pd.read_csv(path, sep=None, engine='python', header=None, dtype=object, nrows=nrows)}
    return pd.read_excel(path, sheet_name=None, header=None, dtype=object, nrows=nrows)


def _cell_text(v):
    return '' if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT else str(v).strip()


def _detect_kind(path):
    """'meta', 'retail' or None, from the header row of any sheet. Used to catch files picked in the wrong slot."""
    for sh in _read_raw(path, nrows=HEADER_SCAN_ROWS).values():
        for row in sh.values.tolist():
            cells = {_cell_text(c).lower() for c in row if _cell_text(c)}
            if META_HEADER <= cells:
                return 'meta'
            if RETAIL_HEADER <= cells:
                return 'retail'
    return None


def _one_date(v):
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, (datetime.datetime, datetime.date, pd.Timestamp)):
        return pd.Timestamp(v).normalize()
    if isinstance(v, (int, float)):                                   # Excel serial number
        return (pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(v))).normalize() if 20000 < v < 80000 else pd.NaT
    s = str(v).strip()
    if not s:
        return pd.NaT
    for fmt in ('%d-%b-%y', '%d-%b-%Y', '%Y-%m-%d', '%Y-%m-%d %H:%M:%S', '%d/%m/%Y', '%m/%d/%Y'):
        try:
            return pd.Timestamp(datetime.datetime.strptime(s, fmt))
        except ValueError:
            pass
    return pd.to_datetime(s, errors='coerce')


def _map_unique(series, fn):
    """Applies fn once per distinct value (a retail file has hundreds of thousands of rows, few distinct values)."""
    cache = {}
    def go(v):
        k = v if isinstance(v, (str, int, float, datetime.date)) else repr(v)
        if k not in cache:
            cache[k] = fn(v)
        return cache[k]
    return series.map(go)


def _clean_receipt(v):
    s = _cell_text(v)
    return s[:-2] if s.endswith('.0') and s[:-2].isdigit() else s


def _canon_brand(v):
    """Brand spelling shared with the Meta side: both come from BRAND_RULES, so they can never differ.
    A retail brand that matches no rule is kept as written (and reported)."""
    s = _cell_text(v)
    return _brand_of(s) or RETAIL_ONLY_BRANDS.get(s.upper()) or s


def _plain_store(v):
    return re.sub(r'\s+', ' ', _cell_text(v)).upper()


def _canon_store(v):
    """Store name as the Meta side writes it: upper case, single spaces, then the Brand Performance store masking
    (RAP ... -> RAPTOR ...) and typo fixes."""
    s = re.sub(r'\s+', ' ', _cell_text(v)).upper()
    for wrong, correct in STORE_MASK.items():
        if s.startswith(wrong):
            s = correct + s[len(wrong):]
            break
    for wrong, correct in STORE_TYPO_FIXES.items():
        if s.startswith(wrong):
            s = correct + s[len(wrong):]
    return STORE_EXACT_FIXES.get(s, s)


def _canon_tag(v):
    s = _cell_text(v)
    return {'online': 'Online', 'offline': 'Offline'}.get(s.lower(), s or '(blank)')


def _read_retail(path, progress):
    raw = _read_raw(path)
    for name, sh in raw.items():
        for i in range(min(len(sh), HEADER_SCAN_ROWS)):
            names = [_cell_text(c).lower() for c in sh.iloc[i].tolist()]
            if RETAIL_HEADER <= set(names):
                progress(f'Reading retail sheet "{name}"' if len(raw) > 1 else 'Reading retail sales file…')
                body = sh.iloc[i + 1:]

                def col(header):
                    return body.iloc[:, names.index(header)]

                def opt(header):                 # optional column: blank when the file does not have it
                    return col(header) if header in names else pd.Series('', index=body.index, dtype=object)

                out = pd.DataFrame({
                    'Store_raw': col('store_name'), 'Date_raw': col('rcpt_date'), 'Receipt_raw': col('receipt_no'),
                    'Brand_raw': col('brand'), 'Tag_raw': col('tag'),
                    'Type_raw': opt('transaction_type'), 'Desc_raw': opt('long_description'),
                    'Price_raw': opt('ext_price_w_tax'), 'Cust_raw': opt('customer_name'), 'Code_raw': opt('store_code'),
                })
                blank = out.apply(lambda s: s.map(lambda v: _cell_text(v) == '')).all(axis=1)
                have = {'price': 'ext_price_w_tax' in names, 'customer': 'customer_name' in names,
                        'desc': 'long_description' in names}
                return out[~blank].reset_index(drop=True), have
    raise ValueError('No sheet has the expected retail columns: STORE_NAME, RCPT_DATE, RECEIPT_NO, BRAND, TAG. '
                     'Is this the retail sales export?')


def _mask_brands(r, have_price):
    """Brand masking from the Brand Performance cleaner: Jinko + Raptor -> Raptor, '*' + Akaso -> Akaso,
    Plaud / PlaudAI -> PlaudAi, brand ZZ -> first word of the description (ZZ paper bags with a zero price are
    removed). Returns (rows kept, rows removed, number of rows whose brand changed)."""
    before = r['Brand_raw'].map(_cell_text)
    desc = r['Desc_raw'].map(_cell_text)
    up = before.str.upper()
    r = r.copy()
    r.loc[up.eq('JINKO') & desc.str.contains('raptor', case=False, na=False), 'Brand_raw'] = 'RAPTOR'
    r.loc[before.eq('*') & desc.str.contains('akaso', case=False, na=False), 'Brand_raw'] = 'Akaso'
    r.loc[up.isin(['PLAUD', 'PLAUDAI']), 'Brand_raw'] = 'PlaudAi'
    zz = r['Brand_raw'].map(_cell_text).eq('ZZ')
    removed = r.iloc[0:0]
    if zz.any():
        bag = desc.str.contains('paper bag', case=False, na=False)
        if have_price:
            drop = zz & bag & r['Price'].eq(0)
            removed = r[drop]
            r, desc, zz, bag = r[~drop].copy(), desc[~drop], zz[~drop], bag[~drop]
        rename = zz & ~bag
        if rename.any():
            r.loc[rename, 'Brand_raw'] = desc[rename].map(lambda d: d.split()[0].capitalize() if d.split() else 'ZZ')
    changed = int((r['Brand_raw'].map(_cell_text) != before.reindex(r.index)).sum())
    return r, removed, changed


def _matching_frame(r):
    return pd.DataFrame({'RECEIPT_NO': r['Receipt'],
                         'CUSTOMER_NAME': r['Cust_raw'].map(lambda v: _cell_text(v).upper()),
                         'STORE_NAME': r['Store'], 'STORE_CODE': r['Code_raw'].map(_cell_text),
                         'RCPT_DATE': r['Date'], 'EXT_PRICE_W_TAX': r['Price']}, index=r.index)


def _interest_store_key(frame):
    """Store identifier per row: STORE_CODE when filled in, otherwise STORE_NAME."""
    out = frame['STORE_NAME'].fillna('').astype(str).str.strip().str.upper()
    code = frame['STORE_CODE'].fillna('').astype(str).str.strip().str.upper()
    return out.mask(code.ne('') & code.ne('NAN'), code)


def match_interest_to_sales(interest, main, price_col='EXT_PRICE_W_TAX'):
    """Ported from the Brand Performance cleaner. Finds the sale receipt each interest line belongs to. Tiers are
    tried in order; the first hit wins:
      1  same RECEIPT_NO + CUSTOMER_NAME
      2  same store + customer + receipt date
      3  same store + customer, nearest receipt date
      4  same customer, nearest receipt date, any store
    Shared/placeholder customers only use tier 1 or 2. Returns TARGET_RECEIPT, TARGET_CUSTOMER, TIER (0 = no match)."""
    def keys(frame):
        k = pd.DataFrame(index=frame.index)
        k['rk'] = frame['RECEIPT_NO'].fillna('').astype(str)
        k['ck'] = frame['CUSTOMER_NAME'].fillna('').astype(str)
        k['sk'] = _interest_store_key(frame)
        k['dt'] = pd.to_datetime(frame['RCPT_DATE'], errors='coerce')
        return k

    mk, ik = keys(main), keys(interest)
    mk['rcpt'] = main['RECEIPT_NO']
    mk['cust'] = main['CUSTOMER_NAME'].fillna('')
    mk['amt'] = pd.to_numeric(main[price_col], errors='coerce').fillna(0)
    rec = (mk.groupby(['rk', 'ck'], sort=False)
             .agg(rcpt=('rcpt', 'first'), cust=('cust', 'first'), sk=('sk', 'first'),
                  dt=('dt', 'first'), total=('amt', 'sum'))
             .reset_index())
    rec['cn'] = rec['ck'].astype(str).str.strip().str.upper()
    exact = {(r_, c_): i for i, (r_, c_) in enumerate(zip(rec['rk'], rec['ck']))}
    by_customer = {cn: grp for cn, grp in rec.groupby('cn', sort=False)}

    def pick(cands, dt):
        if pd.notna(dt):
            gap = (pd.to_datetime(cands['dt']) - dt).abs().dt.days.fillna(10 ** 6)
        else:
            gap = pd.Series(0, index=cands.index)
        order = pd.DataFrame({'gap': gap, 'neg_total': -cands['total']}).sort_values(['gap', 'neg_total'], kind='stable')
        return order.index[0]

    n = len(interest)
    tgt_rcpt, tgt_cust, tiers = [None] * n, [None] * n, [0] * n
    cache = {}
    for pos, (rk, ck, sk, dt) in enumerate(zip(ik['rk'], ik['ck'], ik['sk'], ik['dt'])):
        hit, tier = exact.get((rk, ck)), 1
        if hit is None:
            cn = ck.strip().upper()
            ckey = (cn, sk, dt if pd.notna(dt) else None)
            if ckey not in cache:
                found, found_tier = None, 0
                cands = by_customer.get(cn)
                if cands is not None and not cands.empty:
                    placeholder = cn in IGNORED_CUSTOMER_NAMES
                    same_store = cands[cands['sk'] == sk]
                    if pd.notna(dt):
                        same_day = same_store[pd.to_datetime(same_store['dt']) == dt]
                        if not same_day.empty:
                            found, found_tier = pick(same_day, dt), 2
                    if found is None and not placeholder and not same_store.empty:
                        found, found_tier = pick(same_store, dt), 3
                    if found is None and not placeholder:
                        found, found_tier = pick(cands, dt), 4
                cache[ckey] = (found, found_tier)
            hit, tier = cache[ckey]
        if hit is not None:
            tgt_rcpt[pos], tgt_cust[pos], tiers[pos] = rec.at[hit, 'rcpt'], rec.at[hit, 'cust'], tier
    return pd.DataFrame({'TARGET_RECEIPT': pd.Series(tgt_rcpt, dtype=object).to_numpy(),
                         'TARGET_CUSTOMER': pd.Series(tgt_cust, dtype=object).to_numpy(),
                         'TIER': tiers}, index=interest.index)


AUDIT_COLS = ['Removal Reason', 'Date', 'Store Name', 'Brand', 'Tag', 'Receipt No', 'Customer', 'Amount',
              'Matched Receipt No', 'Match Method']


def _audit_rows(frame, reasons, matched=None, method=None):
    n = len(frame)
    return pd.DataFrame({
        'Removal Reason': reasons, 'Date': frame['Date'].to_numpy(), 'Store Name': frame['Store'].to_numpy(),
        'Brand': frame['Brand_raw'].map(_cell_text).to_numpy(), 'Tag': frame['Tag'].to_numpy(),
        'Receipt No': frame['Receipt'].to_numpy(), 'Customer': frame['Cust_raw'].map(_cell_text).to_numpy(),
        'Amount': frame['Price'].round(2).to_numpy(),
        'Matched Receipt No': matched if matched is not None else [None] * n,
        'Match Method': method if method is not None else [''] * n,
    })[AUDIT_COLS]


def _clean_retail(path, progress):
    """Returns (retail_data, retail_audit, info).
    retail_data = Date, Store Name, Brand, Tag, Customer Name, EXT_PRICE_W_TAX (summed), Distinct Online
    Transactions (receipt no. + date + customer name, Online tag only, returns excluded). Before counting: store masking, brand masking and the interest
    lines are handled the way the Brand Performance cleaner does; removed lines go to retail_audit."""
    r, have = _read_retail(path, progress)
    rows_in = len(r)
    r['Date'] = _map_unique(r['Date_raw'], _one_date)
    bad = int(r['Date'].isna().sum())
    if bad:
        raise ValueError(f'{bad} row(s) in the retail file have an unreadable RCPT_DATE. Please check that column.')

    progress('Masking retail stores and brands…')
    r['Store'] = _map_unique(r['Store_raw'], _canon_store)
    stores_masked = sorted({f"{_plain_store(v)} → {_canon_store(v)}" for v in r['Store_raw'].unique()
                            if _canon_store(v) != _plain_store(v)})
    r['Receipt'] = _map_unique(r['Receipt_raw'], _clean_receipt)
    r['Tag'] = _map_unique(r['Tag_raw'], _canon_tag)
    r['IsReturn'] = _map_unique(r['Type_raw'], lambda v: 'return' in _cell_text(v).lower())
    r['Price'] = pd.to_numeric(r['Price_raw'].astype(str).str.replace(',', '').str.strip(), errors='coerce').fillna(0.0)

    audit_parts = []
    r, zz_removed, brand_masked = _mask_brands(r, have['price'])
    if len(zz_removed):
        audit_parts.append(_audit_rows(zz_removed, ['ZZ Brand Paper Bag with Zero Price'] * len(zz_removed)))

    progress('Setting aside interest lines…')
    brand_up = r['Brand_raw'].map(lambda v: _cell_text(v).upper())
    desc_up = r['Desc_raw'].map(lambda v: re.sub(r'\s+', ' ', _cell_text(v)).upper())
    is_interest = brand_up.eq('INTEREST') | desc_up.str.contains('INTEREST', na=False)
    interest, r = r[is_interest], r[~is_interest].copy()
    interest_unmatched = 0
    if len(interest):
        if len(r):
            matches = match_interest_to_sales(_matching_frame(interest), _matching_frame(r))
            tier = matches['TIER'].to_numpy().astype(int)
            matched = matches['TARGET_RECEIPT'].to_numpy()
        else:
            tier, matched = np.zeros(len(interest), dtype=int), [None] * len(interest)
        interest_unmatched = int((tier == 0).sum())
        reasons = ['Interest line - set aside (matched to a sale receipt)' if t > 0
                   else 'Interest line - NO MATCHING SALE FOUND' for t in tier]
        method = [INTEREST_MATCH_LABELS.get(int(t), 'No Matching Sale Found') for t in tier]
        audit_parts.append(_audit_rows(interest, reasons, matched=list(matched), method=method))
    retail_audit = (pd.concat(audit_parts, ignore_index=True) if audit_parts else pd.DataFrame(columns=AUDIT_COLS))
    if len(retail_audit):
        retail_audit = retail_audit.sort_values(['Removal Reason', 'Date', 'Store Name', 'Receipt No']).reset_index(drop=True)

    progress('Standardizing retail brands and tags…')
    r['Brand'] = _map_unique(r['Brand_raw'], _canon_brand)

    known_brands = {b for b, _ in BRAND_RULES} | set(RETAIL_ONLY_BRANDS.values())
    info = {
        'rows_in': rows_in, 'rows_used': len(r),
        'zz_removed': len(zz_removed), 'interest_lines': len(interest), 'interest_unmatched': interest_unmatched,
        'brand_masked_rows': brand_masked, 'stores_masked': stores_masked,
        'unmapped_stores': sorted({s for s in r['Store'].unique() if s not in RECOGNIZED_STORES}),
        'unmapped_brands': sorted({b for b in r['Brand'].unique() if b not in known_brands}),
        'other_tags': sorted({t for t in r['Tag'].unique() if t not in ('Online', 'Offline')}),
    }

    r['Customer'] = r['Cust_raw'].map(lambda v: re.sub(r'\s+', ' ', _cell_text(v)).upper())
    keys = ['Date', 'Store', 'Brand', 'Tag', 'Customer']

    progress('Counting distinct online transactions…')
    on = r[r['Tag'].eq('Online')]
    info['online_lines'] = len(on)
    info['online_returns_excluded'] = int(on['IsReturn'].sum())
    on = on[~on['IsReturn']]
    info['online_no_receipt'] = int(on['Receipt'].eq('').sum())
    on = on[on['Receipt'].ne('')].copy()
    # one online transaction = same receipt number + same date + same customer name (store added so that two
    # stores that happen to reuse a receipt number are never merged)
    on['Key'] = on['Store'] + '|' + on['Date'].dt.strftime('%Y-%m-%d') + '|' + on['Receipt'] + '|' + on['Customer']
    info['online_receipts'] = int(on['Key'].nunique())
    counts = on.groupby(keys)['Key'].nunique()

    retail_data = (r.groupby(keys, as_index=False)['Price'].sum()
                   .sort_values(keys).reset_index(drop=True))
    retail_data['Price'] = retail_data['Price'].round(2)
    idx = pd.MultiIndex.from_frame(retail_data[keys])
    online_col = counts.reindex(idx).fillna(0).astype(int).to_numpy()
    is_online = retail_data['Tag'].eq('Online').to_numpy()
    retail_data['Online'] = [int(n) if o else None for n, o in zip(online_col, is_online)]
    retail_data.columns = ['Date', 'Store Name', 'Brand', 'Tag', 'Customer Name', 'EXT_PRICE_W_TAX',
                           'Distinct Online Transactions']
    info['retail_rows_out'] = len(retail_data)
    info['online_rows_out'] = int(is_online.sum())
    info['online_sum'] = int(np.nansum([n for n in retail_data['Distinct Online Transactions'] if n is not None]))
    info['amount_used'] = round(float(r['Price'].sum()), 2)
    info['amount_out'] = round(float(retail_data['EXT_PRICE_W_TAX'].sum()), 2)
    info['date_min'], info['date_max'] = r['Date'].min(), r['Date'].max()
    return retail_data, retail_audit, info


# ---------------------------------------------------------------- main entry
def process_meta_ads(file_path, retail_path=None, progress=None):
    progress = progress or (lambda text: None)
    if not retail_path:
        raise ValueError('The Meta Ads Cleaner needs two files: the Meta Ads (Supermetrics) export and the retail '
                         'sales file.')
    if Path(file_path).resolve() == Path(retail_path).resolve():
        raise ValueError('The same file was picked twice. Pick the Meta Ads export and the retail sales file.')
    progress('Checking the two files…')
    if (_detect_kind(file_path), _detect_kind(retail_path)) == ('retail', 'meta'):
        file_path, retail_path = retail_path, file_path          # picked in the wrong slots: just swap them
        progress('The two files were picked the other way round; using them as Meta Ads + retail.')
    progress('Reading Meta Ads export…')
    df = _read_source(file_path, progress)
    rows_in = len(df)
    bad_dates = int(df['Date'].isna().sum())
    if bad_dates:
        raise ValueError(f'{bad_dates} row(s) have an unreadable Date. Please check the Date column.')

    progress('Mapping stores and brands…')
    df['Store'] = df['Campaign name'].map(_store_of)
    df['Brand_adset'] = df['Ad set name'].map(_brand_of)
    df['Brand_camp'] = df['Campaign name'].map(_brand_of)
    df['Brand'] = df['Brand_adset'].fillna(df['Brand_camp'])
    df['Camp_conv'] = df.groupby('Campaign name')['Conversations'].transform('sum')
    df['Reason'] = df.apply(_removal_reason, axis=1)

    kept = df[df['Reason'].isna()]
    removed = df[df['Reason'].notna()].sort_values(['Reason', 'Campaign name', 'Ad set name', 'Date'])

    progress('Consolidating per date, store and brand…')
    daily = (kept.groupby(['Date', 'Store', 'Brand'], as_index=False)[['Cost', 'Conversations']].sum()
             .sort_values(['Store', 'Brand', 'Date']).reset_index(drop=True))
    daily['Cost'] = daily['Cost'].round(2)
    daily['Conversations'] = daily['Conversations'].round().astype(int)
    summary = (daily.groupby(['Store', 'Brand'], as_index=False)[['Cost', 'Conversations']].sum()
               .sort_values(['Store', 'Brand']).reset_index(drop=True))
    summary['Cost'] = summary['Cost'].round(2)

    # Messaging campaigns that matched no store rule and are not known exclusions -> warn (new store?).
    msg = df[df['Reason'].eq('Store not in recognized list') & df['Store'].isna()]  # eligible campaigns with no store rule
    unmapped = sorted({c for c in msg['Campaign name'].unique()
                       if not any(re.search(p, _norm(c)) for p in KNOWN_EXCLUDED_STORE_RULES)})

    retail_data, retail_audit, retail_info = _clean_retail(retail_path, progress)

    progress('Writing workbook…')
    out_dir = Path.home() / 'Desktop' / 'Cleaned Data'
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f'{Path(file_path).stem}_Meta_Ads_Cleaned.xlsx'
    _write_workbook(out_path, summary, daily, removed, kept, df, retail_data, retail_audit, retail_info)

    removed_stats = []
    for reason, g in removed.groupby('Reason'):
        removed_stats.append((f'{reason} (cost {g["Cost"].sum():,.0f})', len(g)))
    stats = {'rows_in': rows_in, 'rows_out': len(daily), 'removed': removed_stats,
             'unmapped_stores': unmapped, 'retail': retail_info}
    return len(daily), str(out_path), stats


# ---------------------------------------------------------------- workbook
def _write_workbook(out_path, summary, daily, removed, kept, df, retail_data, retail_audit, retail_info):  # noqa: C901
    font = 'Arial'
    hf, hfill = Font(name=font, bold=True, color='FFFFFF'), PatternFill('solid', fgColor='1F3864')
    base, bold = Font(name=font), Font(name=font, bold=True)

    def sheet(ws, cols, widths):
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill, c.alignment = hf, hfill, Alignment(horizontal='center')
        ws.freeze_panes = 'A2'
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def style(ws, fmts):
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = base
            for idx, f in fmts.items():
                row[idx].number_format = f

    wb = Workbook()

    # ---- 1. Meta Cleaned Data: row-level rows that passed every check, tagged with Store and Brand
    wc = wb.active
    wc.title = 'Meta Cleaned Data'
    sheet(wc, ['Date', 'Campaign name', 'Ad set name', 'Cost', 'Messaging conversations started', 'Store', 'Brand'],
          [12, 58, 58, 14, 18, 34, 14])
    tagged = kept.sort_values(['Store', 'Brand', 'Date', 'Campaign name', 'Ad set name'])
    for date, camp, adset, cost, conv, store, brand in zip(tagged['Date'], tagged['Campaign name'], tagged['Ad set name'],
                                                           tagged['Cost'], tagged['Conversations'], tagged['Store'],
                                                           tagged['Brand']):
        wc.append([date.to_pydatetime(), camp, adset, float(cost), float(conv), store, brand])
    style(wc, {0: 'yyyy-mm-dd', 3: '#,##0.00', 4: '#,##0'})
    wc.auto_filter.ref = f'A1:G{len(tagged) + 1}'

    # ---- 2. Retail Data: Date + Store + Brand + Tag + Customer, with the summed EXT_PRICE_W_TAX and the distinct
    #         online transaction count (Online rows only; blank on Offline rows). Not joined to the Meta Ads data.
    wrd = wb.create_sheet('Retail Data')
    sheet(wrd, ['Date', 'Store Name', 'Brand', 'Tag', 'Customer Name', 'EXT_PRICE_W_TAX',
                'Distinct Online Transactions'], [12, 34, 14, 12, 34, 18, 30])
    for d, s, b, t, cust, amt, n in retail_data.itertuples(index=False):
        wrd.append([d.to_pydatetime(), s, b, t, cust, float(amt), None if n is None or pd.isna(n) else int(n)])
    style(wrd, {0: 'yyyy-mm-dd', 5: '#,##0.00', 6: '#,##0'})
    wrd.auto_filter.ref = f'A1:G{len(retail_data) + 1}'

    # ---- 3. Reconciliation: Meta Ads section, then Retail section
    wr = wb.create_sheet('Reconciliation')
    sheet(wr, ['Check', 'Rows / Value', 'Cost'], [64, 20, 18])
    section_rows, bold_rows = [], []

    def put(row, bold_row=False):
        wr.append(row)
        if bold_row:
            bold_rows.append(wr.max_row)

    def section(title):
        wr.append([title])
        section_rows.append(wr.max_row)

    section('META ADS')
    put(['Kept (before daily consolidation)', len(kept), round(float(kept.Cost.sum()), 2)])
    for reason, g in removed.groupby('Reason'):
        put([reason, len(g), round(float(g.Cost.sum()), 2)])
    tot_rows, tot_cost = len(kept) + len(removed), round(float(kept.Cost.sum() + removed.Cost.sum()), 2)
    put(['TOTAL (kept + removed)', tot_rows, tot_cost], True)
    put(['Original file', len(df), round(float(df.Cost.sum()), 2)])
    put(['Difference (must be 0)', tot_rows - len(df), round(tot_cost - float(df.Cost.sum()), 2)], True)
    put(['Rows after consolidation per date, store and brand', len(daily), round(float(summary.Cost.sum()), 2)])
    put(['Difference vs kept cost (must be 0)', '', round(float(summary.Cost.sum()) - float(kept.Cost.sum()), 2)], True)
    wr.append([])
    section('RETAIL')
    ri = retail_info
    put(['Retail source rows (blank lines ignored)', ri['rows_in']])
    put(['  less: ZZ paper-bag lines with zero price (removed)', ri['zz_removed']])
    put(['  less: Interest lines (set aside, see Audit)', ri['interest_lines']])
    put(['Retail rows used', ri['rows_used']])
    put(['Check: source - removed - used (must be 0)',
         ri['rows_in'] - ri['zz_removed'] - ri['interest_lines'] - ri['rows_used']], True)
    put(['Interest lines with no matching sale', ri['interest_unmatched']])
    put(['Brand masks applied (rows whose brand changed)', ri['brand_masked_rows']])
    put(['Store masks applied', '; '.join(ri['stores_masked']) or 'none'])
    put(['Retail Data rows (distinct Date + Store + Brand + Tag + Customer)', ri['retail_rows_out']])
    put(['Amount (EXT_PRICE_W_TAX) of the retail rows used', ri['amount_used']])
    put(['Amount on the Retail Data sheet', ri['amount_out']])
    put(['Difference (must be 0)', round(ri['amount_out'] - ri['amount_used'], 2)], True)
    put(['Lines tagged Online', ri['online_lines']])
    put(['  less: Online return lines (not counted as transactions)', ri['online_returns_excluded']])
    put(['  less: Online lines with no receipt number', ri['online_no_receipt']])
    put(['Distinct online transactions overall (receipt no. + date + customer)', ri['online_receipts']])
    put(['Sum of Distinct Online Transactions on the Retail Data sheet', ri['online_sum']])
    put(['Difference (a receipt holding several brands counts once per brand)', ri['online_sum'] - ri['online_receipts']])
    dates = ('n/a' if pd.isna(ri['date_min']) else f"{ri['date_min']:%Y-%m-%d} to {ri['date_max']:%Y-%m-%d}")
    put(['Retail dates covered', dates])
    put(['Stores not in the recognized list (kept as-is)', '; '.join(ri['unmapped_stores']) or 'none'])
    put(['Brands not matching a known brand (kept as-is)', '; '.join(ri['unmapped_brands']) or 'none'])
    put(['Tags other than Online / Offline', '; '.join(ri['other_tags']) or 'none'])
    style(wr, {1: '#,##0', 2: '#,##0.00'})
    for r_ in bold_rows:
        for c in wr[r_]:
            c.font = bold
    sec_fill = PatternFill('solid', fgColor='D9E1F2')
    for r_ in section_rows:
        for c in wr[r_][:3]:
            c.font, c.fill = bold, sec_fill
    for row in wr.iter_rows(min_row=2):                      # text results (e.g. store masks) read left-aligned
        if isinstance(row[1].value, str):
            row[1].alignment = Alignment(horizontal='left')

    # ---- 4. Audit: everything that was taken out, Meta Ads and retail together (Source column tells them apart)
    wa = wb.create_sheet('Audit')
    sheet(wa, ['Source', 'Removal Reason', 'Date', 'Store Name', 'Brand', 'Tag', 'Campaign name', 'Ad set name',
               'Receipt No', 'Customer', 'Cost / Amount', 'Messaging conversations started', 'Matched Receipt No',
               'Match Method'], [12, 46, 12, 34, 14, 10, 58, 58, 14, 26, 14, 18, 18, 40])

    def blank(v):
        return None if v is None or (not isinstance(v, str) and pd.isna(v)) else v

    for date, camp, adset, cost, conv, store, brand, reason in zip(
            removed['Date'], removed['Campaign name'], removed['Ad set name'], removed['Cost'],
            removed['Conversations'], removed['Store'], removed['Brand'], removed['Reason']):
        wa.append(['Meta Ads', reason, date.to_pydatetime(), blank(store), blank(brand), None, camp, adset,
                   None, None, float(cost), float(conv), None, None])
    for reason, date, store, brand, tag, receipt, cust, amount, matched, method in retail_audit.itertuples(index=False):
        wa.append(['Retail', reason, date.to_pydatetime(), store, brand, tag, None, None,
                   blank(receipt) or None, blank(cust) or None, float(amount),
                   None, None if blank(matched) is None else str(matched), method or None])
    style(wa, {2: 'yyyy-mm-dd', 10: '#,##0.00', 11: '#,##0'})
    wa.auto_filter.ref = f'A1:N{wa.max_row}'

    try:
        wb.save(out_path)
    except PermissionError:  # the previous output is open in Excel
        alt = out_path.with_name(f'{out_path.stem}_{datetime.datetime.now():%H%M%S}.xlsx')
        wb.save(alt)
        raise PermissionError(f'Close "{out_path.name}" in Excel and run again (a copy was saved as {alt.name}).')
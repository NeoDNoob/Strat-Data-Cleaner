"""Meta Ads Cleaner for Strat Data Cleaner.

Entry point: process_meta_ads(file_path, progress=None) -> (rows_out, output_path, stats)
Input : Supermetrics Meta Ads export (Date, Campaign name, Ad set name, Cost, Messaging conversations started)
Output: Excel workbook in Desktop/Cleaned Data with Summary, Daily, Removal Audit Log and Reconciliation sheets.
"""
import datetime
import re
from pathlib import Path

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


# ---------------------------------------------------------------- main entry
def process_meta_ads(file_path, progress=None):
    progress = progress or (lambda text: None)
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

    progress('Writing workbook…')
    out_dir = Path.home() / 'Desktop' / 'Cleaned Data'
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f'{Path(file_path).stem}_Meta_Ads_Cleaned.xlsx'
    _write_workbook(out_path, summary, daily, removed, kept, df)

    removed_stats = []
    for reason, g in removed.groupby('Reason'):
        removed_stats.append((f'{reason} (cost {g["Cost"].sum():,.0f})', len(g)))
    stats = {'rows_in': rows_in, 'rows_out': len(daily), 'removed': removed_stats,
             'unmapped_stores': unmapped}
    return len(daily), str(out_path), stats


# ---------------------------------------------------------------- workbook
def _write_workbook(out_path, summary, daily, removed, kept, df):  # noqa: C901
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
    ws = wb.active
    ws.title = 'Summary'
    sheet(ws, ['Store', 'Brand', 'Cost', 'Messaging Conversations Started'], [34, 14, 16, 32])
    for r in summary.itertuples(index=False):
        ws.append([r.Store, r.Brand, float(r.Cost), int(r.Conversations)])
    ws.append(['TOTAL', '', round(float(summary.Cost.sum()), 2), int(summary.Conversations.sum())])
    style(ws, {2: '#,##0.00', 3: '#,##0'})
    for c in ws[ws.max_row]:
        c.font = bold
    ws.auto_filter.ref = f'A1:D{len(summary) + 1}'

    wd = wb.create_sheet('Daily')
    sheet(wd, ['Date', 'Store', 'Brand', 'Cost', 'Messaging Conversations Started'], [12, 34, 14, 14, 32])
    for r in daily.itertuples(index=False):
        wd.append([r.Date.to_pydatetime(), r.Store, r.Brand, float(r.Cost), int(r.Conversations)])
    style(wd, {0: 'yyyy-mm-dd', 3: '#,##0.00', 4: '#,##0'})
    wd.auto_filter.ref = f'A1:E{len(daily) + 1}'

    wa = wb.create_sheet('Removal Audit Log')
    sheet(wa, ['Date', 'Campaign name', 'Ad set name', 'Cost', 'Messaging conversations started', 'Removal Reason'],
          [12, 58, 58, 14, 18, 42])
    for r in removed.itertuples(index=False):
        wa.append([r.Date.to_pydatetime(), r._1, r._2, float(r.Cost), float(r.Conversations), r.Reason])
    style(wa, {0: 'yyyy-mm-dd', 3: '#,##0.00', 4: '#,##0'})
    wa.auto_filter.ref = f'A1:F{len(removed) + 1}'

    wr = wb.create_sheet('Reconciliation')
    sheet(wr, ['Bucket', 'Source rows', 'Cost'], [46, 14, 18])
    wr.append(['Kept (before daily consolidation)', len(kept), round(float(kept.Cost.sum()), 2)])
    for reason, g in removed.groupby('Reason'):
        wr.append([reason, len(g), round(float(g.Cost.sum()), 2)])
    tot_rows, tot_cost = len(kept) + len(removed), round(float(kept.Cost.sum() + removed.Cost.sum()), 2)
    wr.append(['TOTAL (kept + removed)', tot_rows, tot_cost])
    wr.append(['Original file', len(df), round(float(df.Cost.sum()), 2)])
    wr.append(['Difference (must be 0)', tot_rows - len(df), round(tot_cost - float(df.Cost.sum()), 2)])
    wr.append(['Summary total (after consolidation)', len(daily), round(float(summary.Cost.sum()), 2)])
    wr.append(['Difference vs kept cost (must be 0)', '', round(float(summary.Cost.sum()) - float(kept.Cost.sum()), 2)])
    style(wr, {1: '#,##0', 2: '#,##0.00'})
    for r_ in (wr.max_row - 4, wr.max_row - 2):
        for c in wr[r_]:
            c.font = bold

    try:
        wb.save(out_path)
    except PermissionError:  # the previous output is open in Excel
        alt = out_path.with_name(f'{out_path.stem}_{datetime.datetime.now():%H%M%S}.xlsx')
        wb.save(alt)
        raise PermissionError(f'Close "{out_path.name}" in Excel and run again (a copy was saved as {alt.name}).')
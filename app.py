import os, json, re
from pathlib import Path
from flask import Flask, request, jsonify, redirect, Response
import pandas as pd
import numpy as np
import base64 as _b64

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

# Use persistent disk on Render, local folder otherwise
DATA_DIR = Path(os.environ.get('DATA_DIR', 'data'))
DATA_DIR.mkdir(exist_ok=True)
PROCESSED_PATH = DATA_DIR / 'processed.json'
PREV_STATUS_PATH  = DATA_DIR / 'prev_statuses.json'
NEW_STORAGE_PATH  = DATA_DIR / 'new_to_storage.json'
DISPOSITIONS_PATH = DATA_DIR / 'dispositions.json'    # order dispositions
BACKORDER_PATH    = DATA_DIR / 'backorder.json'        # storage jobs with incomplete parts
BILLING_PATH      = DATA_DIR / 'billing.json'          # storage billing orders
BILLING_PATH      = DATA_DIR / 'billing.json'          # storage billing file data
BILLING_PATH      = DATA_DIR / 'billing_temp.xlsx'    # storage billing file
BILLING_PATH      = DATA_DIR / 'billing_temp.xlsx'    # storage billing file
META_PATH = DATA_DIR / 'meta.json'

# ── Excel reader ──────────────────────────────────────────────────────────────
def read_excel_fast(path):
    try:
        return pd.read_excel(path, engine='calamine')
    except Exception:
        return pd.read_excel(path, engine='openpyxl')

# ── Age bucket ────────────────────────────────────────────────────────────────
AGE_BUCKET_ORDER = ['0-7d','8-14d','15-30d','31-60d','61-90d','91-180d','181-365d','365d+']
AGE_BUCKET_LABELS = {
    '0-7d':'0 – 7 Days','8-14d':'8 – 14 Days','15-30d':'15 – 30 Days',
    '31-60d':'31 – 60 Days','61-90d':'61 – 90 Days','91-180d':'91 – 180 Days',
    '181-365d':'181 – 365 Days','365d+':'Over 365 Days'
}

def assign_age_buckets(series):
    result = pd.Series('365d+', index=series.index, dtype=str)
    result[series <= 365] = '181-365d'
    result[series <= 180] = '91-180d'
    result[series <= 90]  = '61-90d'
    result[series <= 60]  = '31-60d'
    result[series <= 30]  = '15-30d'
    result[series <= 14]  = '8-14d'
    result[series <= 7]   = '0-7d'
    result[series.isna()] = None
    return result

# ── Process uploaded files ────────────────────────────────────────────────────
def process_backorder(warehouse_path, backorder_path, dispositions=None):
    """Cross-reference storage orders against the ship schedule to find parts
    not yet in warehouse inventory.  Matching is done on ORDER NUMBER, not just
    project name, so a new order for the same project is never incorrectly flagged.

    Primary:  orders flagged 'Storage Charged' in dispositions that also appear
              in the warehouse in a storage location.
    Fallback: any order physically in a storage location (when no dispositions set).
    """
    df_wh = read_excel_fast(warehouse_path)
    try:
        df_ss = pd.read_csv(backorder_path, encoding='latin-1')
    except Exception:
        df_ss = pd.read_csv(backorder_path, encoding='utf-8')

    STORAGE_LOCS = [
        'S-Finished Goods/D1000-D2000',
        'S-Offsite Warehouse - Porter',
        'S-Offsite Warehouse',
        'S-Offsite Warehouse - B',
        'S-Offsite Warehouse \u2013 C',
        'S-Offsite Warehouse - Fabric',
        'S-Offsite Warehouse - Hardware',
        'S-Trailer',
    ]
    storage_inv = df_wh[df_wh['LOCATION_GROUP'].isin(STORAGE_LOCS)].copy()

    # Normalise the ORDERS column to strings for comparison against SS Order No
    storage_inv['ORDER_STR'] = (
        storage_inv['ORDERS']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )
    storage_inv = storage_inv[storage_inv['ORDER_STR'] != '']

    # Determine which ORDER NUMBERS are in scope
    if dispositions:
        charged_orders = {str(o).strip() for o, d in dispositions.items() if d == 'Storage Charged'}
        if charged_orders:
            storage_order_set = set(storage_inv['ORDER_STR'].unique())
            in_scope_orders = charged_orders & storage_order_set
        else:
            in_scope_orders = set(storage_inv['ORDER_STR'].unique())
    else:
        in_scope_orders = set(storage_inv['ORDER_STR'].unique())

    if not in_scope_orders:
        return []

    # Per-order warehouse summary
    order_summary = storage_inv.groupby('ORDER_STR').agg(
        ship_to_val   =('SHIP_TO',       'first'),
        storage_value =('EXTENDED_COST', 'sum'),
        avg_age       =('INV_AGE',       'mean'),
        max_age       =('INV_AGE',       'max'),
        storage_status=('ORDER_STATUS',  lambda x: x.mode().iloc[0] if len(x) else ''),
    ).reset_index()
    order_dict = {r['ORDER_STR']: r for _, r in order_summary.iterrows()}

    # Build parts-in-warehouse lookup per ORDER (any location)
    wh_df_orders = df_wh.copy()
    wh_df_orders['ORDER_STR'] = (
        wh_df_orders['ORDERS'].fillna('').astype(str).str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )
    wh_parts_by_order = {}
    for order_str, grp in wh_df_orders[wh_df_orders['ORDER_STR'] != ''].groupby('ORDER_STR'):
        wh_parts_by_order[order_str] = set(grp['PART_NO'].astype(str).str.strip().unique())

    # Filter ship schedule to in-scope orders only
    JOB_PART_RE = re.compile(r'(.+-\d+-\d+$|.*-FG$)')
    df_ss['SS_ORDER'] = (
        df_ss['Order No'].astype(str).str.strip()
        .str.replace(r'\.0$', '', regex=True)
    )
    df_ss['PART_NORM'] = df_ss['Part No'].astype(str).str.strip()
    df_ss['REQ_DT']    = pd.to_datetime(df_ss['Require Ship Date'], errors='coerce')
    ss_storage = df_ss[df_ss['SS_ORDER'].isin(in_scope_orders)].copy()

    results = []
    for _, row in ss_storage.iterrows():
        order = row['SS_ORDER']
        part  = row['PART_NORM']
        # Only job-specific parts; exclude BE parts
        if not JOB_PART_RE.match(part) or part.upper().startswith('BE'):
            continue
        # Skip if this part is already in the warehouse for THIS specific order
        if part in wh_parts_by_order.get(order, set()):
            continue
        # order_dict values are pandas Series rows; use [] indexing with defaults
        wh_row = order_dict.get(order)
        if wh_row is not None:
            ship_to_raw    = str(wh_row['ship_to_val'])
            storage_value  = round(float(wh_row['storage_value']), 2)
            avg_age        = round(float(wh_row['avg_age']), 1)
            max_age_val    = int(wh_row['max_age'])
            storage_status = str(wh_row['storage_status'])
        else:
            ship_to_raw    = str(row.get('Customer Address Code', ''))
            storage_value  = 0.0
            avg_age        = 0.0
            max_age_val    = 0
            storage_status = ''
        results.append({
            'ship_to':        ship_to_raw.strip(),
            'project':        re.sub(r'^[ST]-', '', ship_to_raw.strip()).strip(),
            'storage_orders': [order],
            'storage_value':  storage_value,
            'avg_age':        avg_age,
            'max_age':        max_age_val,
            'storage_status': storage_status,
            'ss_order':       order,
            'coordinator':    str(row['Project Coordinator']).strip() if pd.notna(row.get('Project Coordinator')) else '—',
            'part_no':        part,
            'revision':       str(row.get('Revision', '')).strip() if pd.notna(row.get('Revision')) else '',
            'quantity':       int(row['Quantity']) if pd.notna(row.get('Quantity')) else 0,
            'qty_shipped':    int(row['Qty Shipped']) if pd.notna(row.get('Qty Shipped')) else 0,
            'line_status':    str(row['Order Line Status']).strip(),
            'order_status':   str(row['Order Status']).strip(),
            'ship_date':      row['REQ_DT'].strftime('%m/%d/%Y') if pd.notna(row['REQ_DT']) else '—',
            'customer_name':  str(row.get('Customer Name (Project Name) *', '')).strip(),
        })

    return results

def process_files(warehouse_path, notes_path, billing_path=None):
    df = read_excel_fast(warehouse_path)
    notes = pd.read_csv(notes_path)

    df['PROJECT_NAME'] = df['SHIP_TO'].str.replace(r'^S-', '', regex=True).str.strip()
    df['AGE_BUCKET'] = assign_age_buckets(df['INV_AGE'])

    coord_map = (notes[['Order No','Project Coordinator']]
                 .drop_duplicates('Order No')
                 .dropna(subset=['Project Coordinator']))
    df = df.merge(coord_map, left_on='ORDERS', right_on='Order No', how='left')
    df.rename(columns={'Project Coordinator':'COORDINATOR'}, inplace=True)
    df['SHIP_DATE'] = pd.to_datetime(df['SHIP_DATE'], errors='coerce')
    # Load billing file if available — build set of billed order numbers
    billed_orders = set()
    billing_data  = {}  # order → {charge_amount, charge_type, note}
    if BILLING_PATH.exists() and BILLING_PATH.stat().st_size > 0:
        try:
            bf = pd.read_excel(BILLING_PATH)
            bf['Order Number'] = bf['Order Number'].astype(str).str.strip()
            billed_orders = set(bf['Order Number'].unique())
            for _, br in bf.iterrows():
                billing_data[str(br['Order Number']).strip()] = {
                    'charge_amount': round(float(br['Charge Amount']), 2) if pd.notna(br.get('Charge Amount')) else 0,
                    'charge_type':   str(br['Charge Type']) if pd.notna(br.get('Charge Type')) else '',
                    'note':          str(br['Note']) if pd.notna(br.get('Note')) else '',
                }
        except Exception:
            pass
    # Ensure optional columns exist safely
    for _col in ['CONTAINERS', 'REVISION', 'PART_NO']:
        if _col not in df.columns:
            df[_col] = None

    # KPIs
    total_value  = float(df['EXTENDED_COST'].sum())
    over90_value = float(df.loc[df['INV_AGE'] > 90, 'EXTENDED_COST'].sum())
    kpis = {
        'total_value':        total_value,
        'over90_value':       over90_value,
        'over90_pct':         over90_value / total_value if total_value else 0,
        'total_containers':   len(df),
        'avg_age':            round(float(df['INV_AGE'].mean()), 1) if len(df) else 0,
        'in_storage_value':   float(df[df['LOCATION_GROUP'].str.contains('Storage', case=False, na=False)]['EXTENDED_COST'].sum()),
        'finance_hold_value': float(df[df['ORDER_STATUS'].str.contains('Finance Hold', case=False, na=False)]['EXTENDED_COST'].sum()),
    }

    # Coordinator table
    coord_tbl = []
    for coord, grp in df[df['COORDINATOR'].notna()].groupby('COORDINATOR'):
        val = float(grp['EXTENDED_COST'].sum())
        o90 = float(grp.loc[grp['INV_AGE'] > 90, 'EXTENDED_COST'].sum())
        coord_tbl.append({
            'coordinator': coord,
            'containers':  len(grp),
            'value':       round(val, 2),
            'over90_value': round(o90, 2),
            'over90_pct':  o90 / val if val else 0,
            'avg_age':     round(float(grp['INV_AGE'].mean()), 1),
        })

    # Age bucket table
    age_tbl = []
    for code in AGE_BUCKET_ORDER:
        mask = df['AGE_BUCKET'] == code
        val  = float(df.loc[mask, 'EXTENDED_COST'].sum())
        age_tbl.append({
            'bucket': code, 'label': AGE_BUCKET_LABELS[code],
            'containers': int(mask.sum()), 'value': round(val, 2),
            'pct': val / total_value if total_value else 0,
        })

    # Orders table
    orders_tbl = []
    valid = df[df['ORDERS'].notna() & (df['ORDERS'].astype(str) != '.')].copy()
    for order, grp in valid.groupby('ORDERS'):
        ship = grp['SHIP_DATE'].dropna()
        sr = ''
        if len(ship):
            mn = ship.min().strftime('%m/%d/%y')
            mx = ship.max().strftime('%m/%d/%y')
            sr = mn if mn == mx else f'{mn} – {mx}'
        coord = str(grp['COORDINATOR'].dropna().iloc[0]) if grp['COORDINATOR'].notna().any() else 'Unassigned'
        orders_tbl.append({
            'order':       str(order),
            'project':     str(grp['PROJECT_NAME'].iloc[0] or ''),
            'coordinator': coord,
            'containers':  len(grp),
            'value':       round(float(grp['EXTENDED_COST'].sum()), 2),
            'avg_age':     round(float(grp['INV_AGE'].mean()), 1),
            'max_age':     int(grp['INV_AGE'].max()),
            'age_bucket':  str(grp['AGE_BUCKET'].mode().iloc[0]) if len(grp) else '',
            'status':      str(grp['ORDER_STATUS'].mode().iloc[0]) if len(grp) else '',
            'ship_range':  sr,
        })
    orders_tbl.sort(key=lambda x: x['project'])

    # S-Drop table
    sdrop_df = df[df['LOCATION'].str.contains('drop', case=False, na=False) & (df['INV_AGE'] > 2)].copy()
    sdrop_df = sdrop_df.sort_values('INV_AGE', ascending=False)
    sdrop_df['SHIP_DATE_STR'] = sdrop_df['SHIP_DATE'].dt.strftime('%m/%d/%Y').where(sdrop_df['SHIP_DATE'].notna(), None)

    loc_sum = (sdrop_df.groupby('LOCATION')
               .agg(items=('LOCATION','count'), value=('EXTENDED_COST','sum'),
                    avg_age=('INV_AGE','mean'), max_age=('INV_AGE','max'))
               .reset_index().sort_values('items', ascending=False))

    sdrop_df['SHIP_DATE_STR'] = sdrop_df['SHIP_DATE'].dt.strftime('%m/%d/%Y').where(sdrop_df['SHIP_DATE'].notna(), None)
    sdrop_items = []
    for _, row in sdrop_df.iterrows():
        age  = int(row['INV_AGE'])
        flag = '>90 Days' if age > 90 else ('>30 Days' if age > 30 else '≤30 Days')
        sdrop_items.append({
            'location':   str(row['LOCATION']),
            'order':      str(row['ORDERS']) if pd.notna(row['ORDERS']) and str(row['ORDERS']) != '.' else '—',
            'project':    str(row['PROJECT_NAME']) if pd.notna(row['PROJECT_NAME']) and str(row['PROJECT_NAME']) != '.' else '—',
            'coordinator': str(row['COORDINATOR']) if pd.notna(row['COORDINATOR']) else 'Unassigned',
            'part_no':    str(row['PART_NO']) if pd.notna(row['PART_NO']) else '—',
            'part_group': str(row['PART_GROUP']) if pd.notna(row['PART_GROUP']) and str(row['PART_GROUP']) != '.' else '—',
            'qty':        float(row['QUANTITY']) if pd.notna(row['QUANTITY']) else 0,
            'age':        age,
            'value':      round(float(row['EXTENDED_COST']), 2) if pd.notna(row['EXTENDED_COST']) else 0,
            'status':     str(row['ORDER_STATUS']) if pd.notna(row['ORDER_STATUS']) and str(row['ORDER_STATUS']) != '.' else '—',
            'flag':       flag,
            'serial':    str(row.get('CONTAINERS', '')).strip() if pd.notna(row.get('CONTAINERS')) and str(row.get('CONTAINERS','')).strip() not in ('', 'nan') else '—',
            'ship_date': str(row['SHIP_DATE_STR']) if pd.notna(row['SHIP_DATE_STR']) else '—',
        })

    uniq = int(sdrop_df[sdrop_df['ORDERS'].notna() & (sdrop_df['ORDERS'].astype(str) != '.')]['ORDERS'].nunique())
    sdrop = {
        'kpis': {
            'total_items':   len(sdrop_df),
            'total_value':   round(float(sdrop_df['EXTENDED_COST'].sum()), 2),
            'unique_orders': uniq,
            'avg_age':       round(float(sdrop_df['INV_AGE'].mean()), 1) if len(sdrop_df) else 0,
            'max_age':       int(sdrop_df['INV_AGE'].max()) if len(sdrop_df) else 0,
        },
        'by_location': [{'location': r['LOCATION'], 'items': int(r['items']),
                         'value': round(float(r['value']), 2),
                         'avg_age': round(float(r['avg_age']), 1),
                         'max_age': int(r['max_age'])} for _, r in loc_sum.iterrows()],
        'items': sdrop_items,
    }

    # ── Closed Orders tab ─────────────────────────────────────────────────────
    closed_df = df[df['ORDER_STATUS'].str.strip().str.lower() == 'closed'].copy()
    closed_items = []
    for _, row in closed_df.iterrows():
        age = int(row['INV_AGE']) if pd.notna(row['INV_AGE']) else 0
        closed_items.append({
            'location':   str(row['LOCATION']),
            'location_group': str(row['LOCATION_GROUP']) if pd.notna(row['LOCATION_GROUP']) else '—',
            'order':      str(row['ORDERS']) if pd.notna(row['ORDERS']) and str(row['ORDERS']) != '.' else '—',
            'project':    str(row['PROJECT_NAME']) if pd.notna(row['PROJECT_NAME']) and str(row['PROJECT_NAME']) != '.' else '—',
            'coordinator': str(row['COORDINATOR']) if pd.notna(row['COORDINATOR']) else 'Unassigned',
            'part_no':    str(row['PART_NO']) if pd.notna(row['PART_NO']) else '—',
            'part_group': str(row['PART_GROUP']) if pd.notna(row['PART_GROUP']) and str(row['PART_GROUP']) != '.' else '—',
            'qty':        float(row['QUANTITY']) if pd.notna(row['QUANTITY']) else 0,
            'age':        age,
            'value':      round(float(row['EXTENDED_COST']), 2) if pd.notna(row['EXTENDED_COST']) else 0,
            'status':     str(row['ORDER_STATUS']) if pd.notna(row['ORDER_STATUS']) else '—',
            'age_bucket': str(row['AGE_BUCKET']) if pd.notna(row['AGE_BUCKET']) else '—',
            'serial':    str(row.get('CONTAINERS', '')).strip() if pd.notna(row.get('CONTAINERS')) and str(row.get('CONTAINERS','')).strip() not in ('', 'nan') else '—',
            'ship_date': row['SHIP_DATE'].strftime('%m/%d/%Y') if pd.notna(row['SHIP_DATE']) else '—',
        })
    closed = {
        'kpis': {
            'total_items':   len(closed_df),
            'total_value':   round(float(closed_df['EXTENDED_COST'].sum()), 2),
            'unique_orders': int(closed_df['ORDERS'].nunique()),
            'avg_age':       round(float(closed_df['INV_AGE'].mean()), 1) if len(closed_df) else 0,
            'max_age':       int(closed_df['INV_AGE'].max()) if len(closed_df) else 0,
        },
        'items': closed_items,
        'coordinators': sorted(set(closed_df['COORDINATOR'].fillna('Unassigned').unique().tolist())),
    }

    # ── Offsite Storage tab ────────────────────────────────────────────────────
    BUILDING_MAP = {
        'Downstairs Storage': ['S-Finished Goods/D1000-D2000'],
        'Porter Storage':     ['S-Offsite Warehouse - Porter'],
        'North Warehouse':    ['S-Offsite Warehouse', 'S-Offsite Warehouse - B',
                               'S-Offsite Warehouse – C', 'S-Offsite Warehouse - Fabric',
                               'S-Offsite Warehouse - Hardware'],
        'Trailers':           ['S-Trailer'],
    }
    ALL_OFFSITE = [lg for lgs in BUILDING_MAP.values() for lg in lgs]
    offsite_df = df[df['LOCATION_GROUP'].isin(ALL_OFFSITE)].copy()
    if 'CONTAINERS' not in offsite_df.columns:
        offsite_df['CONTAINERS'] = None
    if 'CONTAINERS' not in df.columns:
        df['CONTAINERS'] = None

    # Map each row to its building label
    lg_to_building = {lg: b for b, lgs in BUILDING_MAP.items() for lg in lgs}
    offsite_df['BUILDING'] = offsite_df['LOCATION_GROUP'].map(lg_to_building)

    bld_sum = (offsite_df.groupby('BUILDING')
        .agg(items=('BUILDING','count'), value=('EXTENDED_COST','sum'),
             avg_age=('INV_AGE','mean'), max_age=('INV_AGE','max'),
             orders=('ORDERS','nunique'))
        .reset_index())
    # Sort by defined order
    bld_order = list(BUILDING_MAP.keys())
    bld_sum['_sort'] = bld_sum['BUILDING'].map({b:i for i,b in enumerate(bld_order)})
    bld_sum = bld_sum.sort_values('_sort')


    # Group by order + location_group — keeps rows distinct, ~400KB
    offsite_df['SHIP_DATE_STR'] = offsite_df['SHIP_DATE'].dt.strftime('%m/%d/%Y').where(offsite_df['SHIP_DATE'].notna(), None)
    # Give no-order rows a unique key so they don't collapse together
    offsite_df['ORDER_KEY'] = offsite_df.apply(
        lambda r: str(r['ORDERS']) if pd.notna(r['ORDERS']) and str(r['ORDERS']) != '.'
                  else f"__noorder_{r['LOCATION_GROUP']}_{r.get('PART_NO','')}",
        axis=1
    )
    ord_agg = offsite_df.groupby(['ORDER_KEY','LOCATION_GROUP','BUILDING']).agg(
        order=('ORDERS', lambda x: str(x.iloc[0]) if str(x.iloc[0]) != '.' and pd.notna(x.iloc[0]) else '—'),
        project=('PROJECT_NAME','first'),
        coordinator=('COORDINATOR', lambda x: x.dropna().iloc[0] if x.dropna().any() else 'Unassigned'),
        containers=('ORDERS','count'),
        value=('EXTENDED_COST','sum'),
        avg_age=('INV_AGE','mean'),
        max_age=('INV_AGE','max'),
        status=('ORDER_STATUS', lambda x: x.mode().iloc[0] if len(x) else ''),
        location=('LOCATION','first'),
        ship_date=('SHIP_DATE_STR','first'),
        part_no=('PART_NO', 'first'),
        serial=('CONTAINERS', lambda x: str(x.dropna().iloc[0]).strip() if x.dropna().any() and str(x.dropna().iloc[0]).strip() not in ('','nan') else '—'),
    ).reset_index()

    # ── Billing file processing ──────────────────────────────────────────────

    billing_orders = {}  # order_number -> {charge_amount, charge_type, note, pc, ship_date}
    if billing_path and Path(billing_path).exists():
        try:
            bdf = pd.read_excel(billing_path, engine='openpyxl')
            # Print column names to help debug
            print(f"Billing file columns: {list(bdf.columns)}")
            print(f"Billing file rows: {len(bdf)}")
            if len(bdf) > 0:
                print(f"First row sample: {bdf.iloc[0].to_dict()}")
            for _, row in bdf.iterrows():
                raw_order = row.get('Order Number', row.get('Order No', row.get('Order', '')))
                if pd.isna(raw_order):
                    continue
                # Handle numeric order numbers (Excel stores as float like 2013082.0)
                if isinstance(raw_order, (int, float)):
                    order = str(int(raw_order))
                else:
                    order = str(raw_order).strip().rstrip('.0') if str(raw_order).endswith('.0') else str(raw_order).strip()
                if order and order != 'nan':
                    billing_orders[order] = {
                        'charge_amount': round(float(row['Charge Amount']), 2) if pd.notna(row.get('Charge Amount')) else 0,
                        'charge_type':   str(row.get('Charge Type', '')).strip() if pd.notna(row.get('Charge Type')) else '',
                        'note':          str(row.get('Note', '')).strip() if pd.notna(row.get('Note')) else '',
                        'pc':            str(row.get('PC', '')).strip() if pd.notna(row.get('PC')) else '',
                    }
        except Exception as e:
            billing_orders = {}
    BILLING_PATH.write_text(json.dumps(billing_orders))

    offsite_items = []
    for _, row in ord_agg.sort_values('max_age', ascending=False).iterrows():
        age = int(row['max_age'])
        if age > 90:   flag = '>90 Days'
        elif age > 30: flag = '>30 Days'
        else:          flag = '≤30 Days'
        order_str = str(row['order'])
        billing_info = billing_orders.get(order_str, {})
        offsite_items.append({
            'building':       str(row['BUILDING']),
            'location_group': str(row['LOCATION_GROUP']),
            'order':          order_str,
            'project':        str(row['project']) if pd.notna(row['project']) and str(row['project']) != '.' else '—',
            'coordinator':    str(row['coordinator']),
            'containers':     int(row['containers']),
            'value':          round(float(row['value']), 2),
            'avg_age':        round(float(row['avg_age']), 1),
            'max_age':        age,
            'status':         str(row['status']) if str(row['status']) != '.' else '—',
            'flag':           flag,
            'location':       str(row['location']) if pd.notna(row['location']) else '—',
            'ship_date':      str(row['ship_date']) if row['ship_date'] else '—',
            'part_no':        str(row['part_no']) if pd.notna(row['part_no']) else '—',
            'serial':         str(row['serial']),
            'charge_amount':  billing_info.get('charge_amount', 0),
            'charge_type':    billing_info.get('charge_type', ''),
            'billing_note':   billing_info.get('note', ''),
            'billed':         bool(billing_info),
        })

    offsite = {
        'kpis': {
            'total_items':   len(offsite_df),
            'total_value':   round(float(offsite_df['EXTENDED_COST'].sum()), 2),
            'unique_orders': int(offsite_df['ORDERS'].nunique()),
            'avg_age':       round(float(offsite_df['INV_AGE'].mean()), 1) if len(offsite_df) else 0,
            'max_age':       int(offsite_df['INV_AGE'].max()) if len(offsite_df) else 0,
        },
        'by_group': [{'group': r['BUILDING'], 'items': int(r['items']),
                      'orders': int(r['orders']),
                      'value': round(float(r['value']), 2),
                      'avg_age': round(float(r['avg_age']), 1),
                      'max_age': int(r['max_age'])} for _, r in bld_sum.iterrows()],
        'items': offsite_items,
        'coordinators': sorted(set(offsite_df['COORDINATOR'].fillna('Unassigned').unique().tolist())),
        'statuses':     sorted([s for s in offsite_df['ORDER_STATUS'].dropna().unique().tolist() if s != '.']),
    }


    # ── New-to-Storage detection ──────────────────────────────────────────────
    # Build current order→status+location map
    curr_statuses = {}
    curr_locations = {}
    for order, grp in df[df['ORDERS'].notna() & (df['ORDERS'].astype(str) != '.')].groupby('ORDERS'):
        curr_statuses[str(order)]  = str(grp['ORDER_STATUS'].mode().iloc[0]) if len(grp) else ''
        curr_locations[str(order)] = str(grp['LOCATION_GROUP'].mode().iloc[0]) if len(grp) else ''

    new_to_storage = []

    # Load previous statuses if they exist
    prev_data = {}
    if PREV_STATUS_PATH.exists():
        prev_data     = json.loads(PREV_STATUS_PATH.read_text())
        prev_statuses = prev_data.get('statuses', prev_data) if isinstance(prev_data, dict) and 'statuses' in prev_data else prev_data
        prev_locations= prev_data.get('locations', {}) if isinstance(prev_data, dict) else {}
        for order, curr_status in curr_statuses.items():
            prev_status = prev_statuses.get(order, '') if isinstance(prev_statuses, dict) else ''
            was_storage = prev_status.lower().startswith('storage')
            is_storage  = curr_status.lower().startswith('storage')
            if is_storage and not was_storage:
                grp = df[df['ORDERS'].astype(str) == order]
                if len(grp):
                    coord = str(grp['COORDINATOR'].dropna().iloc[0]) if grp['COORDINATOR'].notna().any() else 'Unassigned'
                    proj  = str(grp['PROJECT_NAME'].iloc[0]) if pd.notna(grp['PROJECT_NAME'].iloc[0]) else '—'
                    ship  = pd.to_datetime(grp['SHIP_DATE'].iloc[0], errors='coerce')
                    new_to_storage.append({
                        'order':          order,
                        'project':        proj,
                        'coordinator':    coord,
                        'prev_status':    prev_status or '(not in previous upload)',
                        'new_status':     curr_status,
                        'detected':       pd.Timestamp.now().strftime('%B %d, %Y'),
                        'value':          round(float(grp['EXTENDED_COST'].sum()), 2),
                        'containers':     len(grp),
                        'avg_age':        round(float(grp['INV_AGE'].mean()), 1),
                        'ship_date':      ship.strftime('%m/%d/%Y') if pd.notna(ship) else '—',
                        'orig_location':  curr_locations.get(order, ''),
                    })

    # Build current ship date map for comparison
    curr_ship_dates = {}
    for order, grp in df[df['ORDERS'].notna() & (df['ORDERS'].astype(str) != '.')].groupby('ORDERS'):
        sd = pd.to_datetime(grp['SHIP_DATE'].iloc[0], errors='coerce')
        curr_ship_dates[str(order)] = sd.strftime('%m/%d/%Y') if pd.notna(sd) else ''

    # Load existing log and remove items that have moved to a new location group
    existing_log = []
    if NEW_STORAGE_PATH.exists():
        existing_log = json.loads(NEW_STORAGE_PATH.read_text())

    # Load dispositions — orders with a disposition set are removed from log
    dispositions = {}
    if DISPOSITIONS_PATH.exists():
        dispositions = json.loads(DISPOSITIONS_PATH.read_text())

    def is_resolved(entry):
        order = entry['order']
        # Resolved by disposition selection
        if order in dispositions and dispositions[order]:
            return True
        # Resolved by location change
        if order not in curr_locations: return False
        orig = entry.get('orig_location', '')
        curr = curr_locations.get(order, '')
        return orig and curr and curr != orig

    active_log = [e for e in existing_log if not is_resolved(e)]

    # Update existing entries: flag if ship date pushed further out
    prev_ship_dates = prev_data.get('ship_dates', {}) if isinstance(prev_data, dict) else {}
    for entry in active_log:
        order = entry['order']
        prev_sd = prev_ship_dates.get(order, '')
        curr_sd = curr_ship_dates.get(order, '')
        try:
            if prev_sd and curr_sd and prev_sd != curr_sd:
                from datetime import datetime
                p = datetime.strptime(prev_sd, '%m/%d/%Y')
                c = datetime.strptime(curr_sd, '%m/%d/%Y')
                entry['ship_date_pushed'] = c > p
                if c > p:
                    entry['ship_date']      = curr_sd
                    entry['prev_ship_date'] = prev_sd
            else:
                entry['ship_date_pushed'] = False
        except Exception:
            entry['ship_date_pushed'] = False

    # Add newly detected items not already in log
    existing_orders = {e['order'] for e in active_log}
    merged_log = active_log + [e for e in new_to_storage if e['order'] not in existing_orders]
    NEW_STORAGE_PATH.write_text(json.dumps(merged_log))

    # Save current statuses, locations AND ship dates for next comparison
    PREV_STATUS_PATH.write_text(json.dumps({'statuses': curr_statuses, 'locations': curr_locations, 'ship_dates': curr_ship_dates}))

    return {
        'kpis':         kpis,
        'billing':      billing_orders,
        'coord_table':  coord_tbl,
        'age_table':    age_tbl,
        'orders_table': orders_tbl,
        'sdrop':        sdrop,
        'closed':       closed,
        'offsite':      offsite,
        'projects':     sorted(df['PROJECT_NAME'].dropna().unique().tolist()),
        'coordinators': sorted(set(df['COORDINATOR'].fillna('Unassigned').unique().tolist())),
        'drop_locations': sorted(sdrop_df['LOCATION'].unique().tolist()),
        'uploaded':     pd.Timestamp.now().strftime('%B %d, %Y'),
    }


# ── HTML pages ────────────────────────────────────────────────────────────────
INDEX_HTML = """<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Aged Inventory Dashboard</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center}
.card{background:#fff;border-radius:12px;box-shadow:0 4px 24px rgba(0,0,0,.10);padding:48px 56px;width:100%;max-width:560px}
.logo{display:flex;align-items:center;gap:12px;margin-bottom:32px}
.logo-icon{width:40px;height:40px;background:#1F3864;border-radius:8px;display:flex;align-items:center;justify-content:center}
.logo-icon svg{width:22px;height:22px;fill:#fff}
h1{font-size:22px;color:#1F3864;font-weight:700}
.subtitle{color:#64748b;font-size:13px;margin-top:2px}
.status-banner{background:#e2efda;border:1px solid #a8d08d;border-radius:8px;padding:12px 16px;margin-bottom:24px;font-size:13px;color:#375623}
.status-banner strong{display:block;margin-bottom:2px}
.drop-zone{border:2px dashed #cbd5e1;border-radius:10px;padding:28px 20px;text-align:center;cursor:pointer;transition:all .2s;margin-bottom:16px;background:#f8fafc}
.drop-zone:hover,.drop-zone.drag-over{border-color:#2E75B6;background:#eff6ff}
.drop-zone svg{width:32px;height:32px;stroke:#94a3b8;margin-bottom:8px}
.drop-zone .label{font-size:14px;color:#475569;font-weight:600}
.drop-zone .sublabel{font-size:12px;color:#94a3b8;margin-top:4px}
.drop-zone input{display:none}
.file-chosen{font-size:12px;color:#2E75B6;margin-top:6px;font-weight:600}
.btn{width:100%;padding:13px;background:#1F3864;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;transition:background .2s;margin-top:8px}
.btn:hover{background:#2E75B6}
.btn:disabled{background:#94a3b8;cursor:not-allowed}
.view-link{display:block;text-align:center;margin-top:20px;font-size:13px;color:#2E75B6;text-decoration:none;font-weight:600}
.view-link:hover{text-decoration:underline}
.progress{display:none;margin-top:20px}
.progress-bar{height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden}
.progress-fill{height:100%;background:linear-gradient(90deg,#1F3864,#2E75B6);width:0;transition:width .5s ease;border-radius:4px}
.steps{margin-top:14px;display:flex;flex-direction:column;gap:8px}
.step{display:flex;align-items:center;gap:10px;font-size:12px;color:#94a3b8}
.step.active{color:#1F3864;font-weight:600}
.step.done{color:#375623}
.step-dot{width:16px;height:16px;border-radius:50%;border:2px solid #cbd5e1;flex-shrink:0}
.step.active .step-dot{border-color:#2E75B6;background:#2E75B6}
.step.done .step-dot{border-color:#375623;background:#375623}
.error-msg{background:#fee2e2;border:1px solid #fca5a5;border-radius:8px;padding:10px 14px;font-size:13px;color:#991b1b;margin-top:12px;display:none}
</style></head><body>
<div class="card">
  <div class="logo">
    <div class="logo-icon"><svg viewBox="0 0 24 24"><path d="M20 7H4a2 2 0 00-2 2v9a2 2 0 002 2h16a2 2 0 002-2V9a2 2 0 00-2-2z"/><path d="M16 3H8a2 2 0 00-2 2v2h12V5a2 2 0 00-2-2z"/></svg></div>
    <div><h1>Aged Inventory Dashboard</h1><div class="subtitle">Upload today's files to refresh the dashboard</div></div>
  </div>
  STATUS_BLOCK
  <form id="uploadForm">
    <div class="drop-zone" id="wz" onclick="document.getElementById('wf').click()">
      <input type="file" id="wf" accept=".xlsx" onchange="pick(this,'wz','wl')">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><polyline points="9 15 12 12 15 15"/></svg>
      <div class="label">Warehouse File (.xlsx)</div><div class="sublabel">Click to browse or drag and drop</div>
      <div class="file-chosen" id="wl"></div>
    </div>
    <div class="drop-zone" id="nz" onclick="document.getElementById('nf').click()">
      <input type="file" id="nf" accept=".csv" onchange="pick(this,'nz','nl')">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>
      <div class="label">Internal Notes (.csv)</div><div class="sublabel">Click to browse or drag and drop</div>
      <div class="file-chosen" id="nl"></div>
    </div>
    <div class="drop-zone" id="bz" onclick="document.getElementById('bf').click()">
      <input type="file" id="bf" accept=".xlsx" onchange="pick(this,'bz','bl')">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>
      <div class="label">Storage Billing (.xlsx) <span style="font-size:11px;color:#94a3b8;font-weight:400">optional</span></div>
      <div class="sublabel">Click to browse or drag and drop</div>
      <div class="file-chosen" id="bl"></div>
    </div>
    <div class="drop-zone" id="boz" onclick="document.getElementById('bof').click()">
      <input type="file" id="bof" accept=".csv" onchange="pick(this,'boz','bol')">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>
      <div class="label">Ship Schedule (.csv) <span style="font-size:11px;color:#94a3b8;font-weight:400">optional</span></div>
      <div class="sublabel">Click to browse or drag and drop</div>
      <div class="file-chosen" id="bol"></div>
    </div>
    <div class="progress" id="prog">
      <div class="progress-bar"><div class="progress-fill" id="pf"></div></div>
      <div class="steps">
        <div class="step" id="s1"><div class="step-dot"></div><span>Uploading files…</span></div>
        <div class="step" id="s2"><div class="step-dot"></div><span>Reading warehouse data (~10 sec)…</span></div>
        <div class="step" id="s3"><div class="step-dot"></div><span>Joining coordinator data…</span></div>
        <div class="step" id="s4"><div class="step-dot"></div><span>Building tables…</span></div>
        <div class="step" id="s5"><div class="step-dot"></div><span>Done!</span></div>
      </div>
    </div>
    <div class="error-msg" id="err"></div>
    <button class="btn" type="submit" id="btn" disabled>Upload &amp; Process</button>
  </form>
  LINK_BLOCK
</div>
<script>
function pick(input,z,l){document.getElementById(l).textContent=input.files[0]?'&#10003; '+input.files[0].name:'';chk();}
function chk(){document.getElementById('btn').disabled=!(document.getElementById('wf').files[0]&&document.getElementById('nf').files[0]);}
['wz','nz','bz','boz'].forEach(function(id){
  var z=document.getElementById(id);
  z.addEventListener('dragover',function(e){e.preventDefault();z.classList.add('drag-over');});
  z.addEventListener('dragleave',function(){z.classList.remove('drag-over');});
  z.addEventListener('drop',function(e){
    e.preventDefault();z.classList.remove('drag-over');
    var fi=z.querySelector('input'),lb=z.querySelector('.file-chosen');
    if(e.dataTransfer.files[0]){var dt=new DataTransfer();dt.items.add(e.dataTransfer.files[0]);fi.files=dt.files;lb.textContent='&#10003; '+e.dataTransfer.files[0].name;chk();}
  });
});
function setStep(n){
  for(var i=1;i<=5;i++){var el=document.getElementById('s'+i);el.classList.remove('active','done');if(i<n)el.classList.add('done');if(i===n)el.classList.add('active');}
  document.getElementById('pf').style.width=((n-1)/4*100)+'%';
}
document.getElementById('uploadForm').addEventListener('submit',async function(e){
  e.preventDefault();
  var btn=document.getElementById('btn'),err=document.getElementById('err');
  btn.disabled=true;err.style.display='none';document.getElementById('prog').style.display='block';
  setStep(1);
  var fd=new FormData();
  fd.append('warehouse',document.getElementById('wf').files[0]);
  fd.append('notes',document.getElementById('nf').files[0]);
  var bf=document.getElementById('bf').files[0];
  if(bf) fd.append('billing',bf);
  var bof=document.getElementById('bof').files[0];
  if(bof) fd.append('backorder',bof);
  var bf=document.getElementById('bf').files[0];
  if(bf) fd.append('billing', bf);
  var bf=document.getElementById('bf').files[0];
  if(bf) fd.append('billing',bf);
  var bof=document.getElementById('bof').files[0];
  if(bof) fd.append('backorder',bof);
  if(document.getElementById('bf').files[0]) fd.append('billing',document.getElementById('bf').files[0]);
  // Advance steps while processing runs in background
  [1000,4000,8000,11000].forEach(function(d,i){setTimeout(function(){setStep(i+2);},d);});
  try{
    // 60 second timeout for initial upload (file transfer can be slow)
    var uploadCtrl=new AbortController();
    var uploadTimeout=setTimeout(()=>uploadCtrl.abort(),60000);
    var res=await fetch('/upload',{method:'POST',body:fd,signal:uploadCtrl.signal});
    clearTimeout(uploadTimeout);
    var data=await res.json();
    if(!data.success) throw new Error(data.error||'Upload failed');
    // Poll /api/upload_status until done
    var attempts=0;
    var poll=setInterval(async function(){
      attempts++;
      if(attempts>120){clearInterval(poll);err.textContent='&#9888; Processing timed out. Please try again.';err.style.display='block';document.getElementById('prog').style.display='none';btn.disabled=false;return;}
      try{
        var sr=await fetch('/api/upload_status');
        var st=await sr.json();
        if(st.status==='done'){clearInterval(poll);setStep(5);document.getElementById('pf').style.width='100%';setTimeout(function(){window.location.href='/dashboard';},800);}
        else if(st.status==='error'){clearInterval(poll);throw new Error(st.error||'Processing failed');}
      }catch(pe){clearInterval(poll);document.getElementById('prog').style.display='none';err.textContent='&#9888; '+pe.message;err.style.display='block';btn.disabled=false;}
    },1000);
  }catch(ex){
    document.getElementById('prog').style.display='none';
    err.textContent='&#9888; '+ex.message;err.style.display='block';btn.disabled=false;
  }
});
</script></body></html>"""


DASHBOARD_HTML = _b64.b64decode("PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CiAgPG1ldGEgY2hhcnNldD0iVVRGLTgiPgogIDxtZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsIGluaXRpYWwtc2NhbGU9MS4wIj4KICA8dGl0bGU+QWdlZCBJbnZlbnRvcnkgRGFzaGJvYXJkPC90aXRsZT4KICA8c3R5bGU+CiAgICAqLCAqOjpiZWZvcmUsICo6OmFmdGVyIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgbWFyZ2luOiAwOyBwYWRkaW5nOiAwOyB9CgogICAgOnJvb3QgewogICAgICAtLW5hdnk6ICMxRjM4NjQ7IC0tYmx1ZTogIzJFNzVCNjsgLS1saWdodC1ibHVlOiAjRDZFNEYwOyAtLXNlY3Rpb24tYmc6ICNFQkYzRkI7CiAgICAgIC0tZ3JlZW46ICMzNzU2MjM7IC0tZ3JlZW4tbGlnaHQ6ICNFMkVGREE7IC0tcmVkOiAjQzAwMDAwOyAtLXJlZC1saWdodDogI0ZGQ0NDQzsKICAgICAgLS1hbWJlcjogI0M1NUExMTsgLS1hbWJlci1saWdodDogI0ZGRjJDQzsKICAgICAgLS10ZXh0OiAjMWUyOTNiOyAtLW11dGVkOiAjNjQ3NDhiOyAtLWJvcmRlcjogI2UyZThmMDsgLS13aGl0ZTogI2ZmZmZmZjsKICAgICAgLS1zdXJmYWNlOiAjZjhmYWZjOwogICAgfQoKICAgIGJvZHkgeyBmb250LWZhbWlseTogJ1NlZ29lIFVJJywgQXJpYWwsIHNhbnMtc2VyaWY7IGJhY2tncm91bmQ6ICNmMGY0Zjg7CiAgICAgICAgICAgY29sb3I6IHZhcigtLXRleHQpOyBmb250LXNpemU6IDEzcHg7IH0KCiAgICAvKiDilIDilIAgSGVhZGVyIOKUgOKUgCAqLwogICAgLmhlYWRlciB7IGJhY2tncm91bmQ6IHZhcigtLW5hdnkpOyBjb2xvcjogI2ZmZjsgcGFkZGluZzogMCAzMnB4OwogICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsKICAgICAgICAgICAgICBoZWlnaHQ6IDU2cHg7IHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1pbmRleDogMTAwOwogICAgICAgICAgICAgIGJveC1zaGFkb3c6IDAgMnB4IDhweCByZ2JhKDAsMCwwLC4yKTsgfQogICAgLmhlYWRlci1sZWZ0IHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OyB9CiAgICAuaGVhZGVyIGgxIHsgZm9udC1zaXplOiAxN3B4OyBmb250LXdlaWdodDogNzAwOyB9CiAgICAuaGVhZGVyLXN1YiB7IGZvbnQtc2l6ZTogMTFweDsgb3BhY2l0eTogLjc7IG1hcmdpbi10b3A6IDFweDsgfQogICAgLnVwbG9hZC1saW5rIHsgYmFja2dyb3VuZDogcmdiYSgyNTUsMjU1LDI1NSwuMTUpOyBjb2xvcjogI2ZmZjsgdGV4dC1kZWNvcmF0aW9uOiBub25lOwogICAgICAgICAgICAgICAgICAgcGFkZGluZzogNnB4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDZweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOwogICAgICAgICAgICAgICAgICAgdHJhbnNpdGlvbjogYmFja2dyb3VuZCAuMnM7IGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoMjU1LDI1NSwyNTUsLjI1KTsgfQogICAgLnVwbG9hZC1saW5rOmhvdmVyIHsgYmFja2dyb3VuZDogcmdiYSgyNTUsMjU1LDI1NSwuMjUpOyB9CgogICAgLyog4pSA4pSAIFRhYiBuYXYg4pSA4pSAICovCiAgICAudGFiLW5hdiB7IGJhY2tncm91bmQ6IHZhcigtLW5hdnkpOyBwYWRkaW5nOiAwIDMycHg7CiAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGdhcDogNHB4OyBib3JkZXItdG9wOiAxcHggc29saWQgcmdiYSgyNTUsMjU1LDI1NSwuMSk7IH0KICAgIC50YWItYnRuIHsgcGFkZGluZzogMTBweCAyMHB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiByZ2JhKDI1NSwyNTUsMjU1LC42KTsKICAgICAgICAgICAgICAgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGJvcmRlcjogbm9uZTsgY3Vyc29yOiBwb2ludGVyOyBib3JkZXItYm90dG9tOiAzcHggc29saWQgdHJhbnNwYXJlbnQ7CiAgICAgICAgICAgICAgIHRyYW5zaXRpb246IGFsbCAuMnM7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KICAgIC50YWItYnRuOmhvdmVyIHsgY29sb3I6IHJnYmEoMjU1LDI1NSwyNTUsLjkpOyB9CiAgICAudGFiLWJ0bi5hY3RpdmUgeyBjb2xvcjogI2ZmZjsgYm9yZGVyLWJvdHRvbS1jb2xvcjogdmFyKC0tYW1iZXItbGlnaHQpOyB9CiAgICAudGFiLWJ0biAudGFiLWJhZGdlIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyBiYWNrZ3JvdW5kOiB2YXIoLS1yZWQpOyBjb2xvcjogI2ZmZjsKICAgICAgICAgICAgICAgICAgICAgICAgICBib3JkZXItcmFkaXVzOiAxMHB4OyBwYWRkaW5nOiAxcHggN3B4OyBmb250LXNpemU6IDEwcHg7CiAgICAgICAgICAgICAgICAgICAgICAgICAgbWFyZ2luLWxlZnQ6IDZweDsgZm9udC13ZWlnaHQ6IDcwMDsgfQoKICAgIC8qIOKUgOKUgCBGaWx0ZXJzIGJhciDilIDilIAgKi8KICAgIC5maWx0ZXJzLWJhciB7IGJhY2tncm91bmQ6ICMxNjJkNTQ7IHBhZGRpbmc6IDEwcHggMzJweDsKICAgICAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGdhcDogMTJweDsgYWxpZ24taXRlbXM6IGZsZXgtZW5kOyBmbGV4LXdyYXA6IHdyYXA7CiAgICAgICAgICAgICAgICAgICBib3JkZXItYm90dG9tOiAxcHggc29saWQgcmdiYSgyNTUsMjU1LDI1NSwuMDgpOyB9CiAgICAvKiBzZHJvcCBmaWx0ZXJzIHNob3duL2hpZGRlbiBieSBKUyAqLwogICAgLmZpbHRlci1ncm91cCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogNHB4OyB9CiAgICAuZmlsdGVyLWdyb3VwIGxhYmVsIHsgZm9udC1zaXplOiAxMHB4OyBjb2xvcjogcmdiYSgyNTUsMjU1LDI1NSwuNjUpOyBmb250LXdlaWdodDogNjAwOwogICAgICAgICAgICAgICAgICAgICAgICAgIHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAuMDVlbTsgfQogICAgLmZpbHRlci1ncm91cCBzZWxlY3QgeyBiYWNrZ3JvdW5kOiByZ2JhKDI1NSwyNTUsMjU1LC4xMik7IGNvbG9yOiAjZmZmOyBib3JkZXI6IDFweCBzb2xpZCByZ2JhKDI1NSwyNTUsMjU1LC4yNSk7CiAgICAgICAgICAgICAgICAgICAgICAgICAgIGJvcmRlci1yYWRpdXM6IDZweDsgcGFkZGluZzogNnB4IDI4cHggNnB4IDEwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDUwMDsKICAgICAgICAgICAgICAgICAgICAgICAgICAgY3Vyc29yOiBwb2ludGVyOyBhcHBlYXJhbmNlOiBub25lOyBtaW4td2lkdGg6IDE4MHB4OyBtYXgtd2lkdGg6IDI2MHB4OwogICAgICAgICAgICAgICAgICAgICAgICAgICBiYWNrZ3JvdW5kLWltYWdlOiB1cmwoImRhdGE6aW1hZ2Uvc3ZnK3htbCwlM0NzdmcgeG1sbnM9J2h0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnJyB3aWR0aD0nMTInIGhlaWdodD0nMTInIHZpZXdCb3g9JzAgMCAyNCAyNCcgZmlsbD0nbm9uZScgc3Ryb2tlPSd3aGl0ZScgc3Ryb2tlLXdpZHRoPScyJyUzRSUzQ3BvbHlsaW5lIHBvaW50cz0nNiA5IDEyIDE1IDE4IDknLyUzRSUzQy9zdmclM0UiKTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgYmFja2dyb3VuZC1yZXBlYXQ6IG5vLXJlcGVhdDsgYmFja2dyb3VuZC1wb3NpdGlvbjogcmlnaHQgOHB4IGNlbnRlcjsgfQogICAgLmZpbHRlci1ncm91cCBzZWxlY3Qgb3B0aW9uIHsgYmFja2dyb3VuZDogdmFyKC0tbmF2eSk7IH0KICAgIC5maWx0ZXItZ3JvdXAgc2VsZWN0OmZvY3VzIHsgb3V0bGluZTogbm9uZTsgYm9yZGVyLWNvbG9yOiB2YXIoLS1saWdodC1ibHVlKTsgfQogICAgLnJlc2V0LWJ0biB7IGJhY2tncm91bmQ6IHJnYmEoMjU1LDI1NSwyNTUsLjEpOyBjb2xvcjogcmdiYSgyNTUsMjU1LDI1NSwuOCk7IGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoMjU1LDI1NSwyNTUsLjIpOwogICAgICAgICAgICAgICAgIGJvcmRlci1yYWRpdXM6IDZweDsgcGFkZGluZzogNnB4IDE0cHg7IGZvbnQtc2l6ZTogMTJweDsgY3Vyc29yOiBwb2ludGVyOwogICAgICAgICAgICAgICAgIHRyYW5zaXRpb246IGFsbCAuMnM7IGFsaWduLXNlbGY6IGZsZXgtZW5kOyB9CiAgICAucmVzZXQtYnRuOmhvdmVyIHsgYmFja2dyb3VuZDogcmdiYSgyNTUsMjU1LDI1NSwuMik7IGNvbG9yOiAjZmZmOyB9CgogICAgLyog4pSA4pSAIE1haW4gbGF5b3V0IOKUgOKUgCAqLwogICAgLm1haW4geyBwYWRkaW5nOiAyNHB4IDMycHg7IG1heC13aWR0aDogMTYwMHB4OyBtYXJnaW46IDAgYXV0bzsgfQogICAgLnRhYi1wYW5lbCB7IGRpc3BsYXk6IG5vbmU7IH0KICAgIC50YWItcGFuZWwuYWN0aXZlIHsgZGlzcGxheTogYmxvY2s7IH0KCiAgICAvKiDilIDilIAgS1BJIGNhcmRzIOKUgOKUgCAqLwogICAgLmtwaS1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxNzBweCwgMWZyKSk7CiAgICAgICAgICAgICAgICBnYXA6IDE2cHg7IG1hcmdpbi1ib3R0b206IDI0cHg7IH0KICAgIC5rcGktY2FyZCB7IGJhY2tncm91bmQ6ICNmZmY7IGJvcmRlci1yYWRpdXM6IDEwcHg7IHBhZGRpbmc6IDIwcHggMjJweDsKICAgICAgICAgICAgICAgIGJveC1zaGFkb3c6IDAgMXB4IDRweCByZ2JhKDAsMCwwLC4wOCk7IGJvcmRlci10b3A6IDNweCBzb2xpZCB2YXIoLS1ibHVlKTsgfQogICAgLmtwaS1jYXJkLmFjY2VudC1yZWQgICB7IGJvcmRlci10b3AtY29sb3I6IHZhcigtLXJlZCk7IH0KICAgIC5rcGktY2FyZC5hY2NlbnQtYW1iZXIgeyBib3JkZXItdG9wLWNvbG9yOiB2YXIoLS1hbWJlcik7IH0KICAgIC5rcGktY2FyZC5hY2NlbnQtZ3JlZW4geyBib3JkZXItdG9wLWNvbG9yOiB2YXIoLS1ncmVlbik7IH0KICAgIC5rcGktbGFiZWwgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA2MDA7IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7CiAgICAgICAgICAgICAgICAgbGV0dGVyLXNwYWNpbmc6IC4wNGVtOyBtYXJnaW4tYm90dG9tOiA4cHg7IH0KICAgIC5rcGktdmFsdWUgeyBmb250LXNpemU6IDI2cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS1uYXZ5KTsgbGluZS1oZWlnaHQ6IDE7IH0KICAgIC5rcGktc3ViICAgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IG1hcmdpbi10b3A6IDVweDsgfQoKICAgIC8qIOKUgOKUgCBTZWN0aW9uIGNhcmRzIOKUgOKUgCAqLwogICAgLnNlY3Rpb24tY2FyZCB7IGJhY2tncm91bmQ6ICNmZmY7IGJvcmRlci1yYWRpdXM6IDEwcHg7IGJveC1zaGFkb3c6IDAgMXB4IDRweCByZ2JhKDAsMCwwLC4wOCk7CiAgICAgICAgICAgICAgICAgICAgbWFyZ2luLWJvdHRvbTogMjRweDsgb3ZlcmZsb3c6IGhpZGRlbjsgfQogICAgLnNlY3Rpb24taGVhZGVyIHsgcGFkZGluZzogMTZweCAyMHB4OyBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKICAgICAgICAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgfQogICAgLnNlY3Rpb24tdGl0bGUgeyBmb250LXNpemU6IDE0cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS1uYXZ5KTsgfQogICAgLnNlY3Rpb24tbWV0YSAgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IH0KCiAgICAvKiDilIDilIAgU2VhcmNoIC8gZmlsdGVyIGlucHV0IOKUgOKUgCAqLwogICAgLnNlYXJjaC1ib3ggeyBwYWRkaW5nOiAxMnB4IDIwcHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBkaXNwbGF5OiBmbGV4OyBnYXA6IDEwcHg7IH0KICAgIC5zZWFyY2gtYm94IGlucHV0LCAuc2VhcmNoLWJveCBzZWxlY3QgewogICAgICBwYWRkaW5nOiA3cHggMTJweDsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogNnB4OwogICAgICBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0KTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZSk7IH0KICAgIC5zZWFyY2gtYm94IGlucHV0IHsgZmxleDogMTsgbWF4LXdpZHRoOiAzMjBweDsgfQogICAgLnNlYXJjaC1ib3ggaW5wdXQ6Zm9jdXMsIC5zZWFyY2gtYm94IHNlbGVjdDpmb2N1cyB7IG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tYmx1ZSk7IH0KCiAgICAvKiDilIDilIAgVGFibGVzIOKUgOKUgCAqLwogICAgLnRhYmxlLXdyYXAgeyBvdmVyZmxvdy14OiBhdXRvOyB9CiAgICB0YWJsZSB7IHdpZHRoOiAxMDAlOyBib3JkZXItY29sbGFwc2U6IGNvbGxhcHNlOyBmb250LXNpemU6IDEycHg7IH0KICAgIHRoZWFkIHRoIHsgYmFja2dyb3VuZDogdmFyKC0tbmF2eSk7IGNvbG9yOiAjZmZmOyBwYWRkaW5nOiA5cHggMTJweDsgdGV4dC1hbGlnbjogbGVmdDsKICAgICAgICAgICAgICAgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMXB4OyB3aGl0ZS1zcGFjZTogbm93cmFwOwogICAgICAgICAgICAgICBjdXJzb3I6IHBvaW50ZXI7IHVzZXItc2VsZWN0OiBub25lOyB9CiAgICB0aGVhZCB0aDpob3ZlciB7IGJhY2tncm91bmQ6IHZhcigtLWJsdWUpOyB9CiAgICB0aGVhZCB0aC5zb3J0LWFzYzo6YWZ0ZXIgIHsgY29udGVudDogJyDihpEnOyB9CiAgICB0aGVhZCB0aC5zb3J0LWRlc2M6OmFmdGVyIHsgY29udGVudDogJyDihpMnOyB9CiAgICB0Ym9keSB0ciB7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB0cmFuc2l0aW9uOiBiYWNrZ3JvdW5kIC4xczsgfQogICAgdGJvZHkgdHI6aG92ZXIgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZWN0aW9uLWJnKTsgfQogICAgdGJvZHkgdHI6bnRoLWNoaWxkKGV2ZW4pIHsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZSk7IH0KICAgIHRib2R5IHRyOm50aC1jaGlsZChldmVuKTpob3ZlciB7IGJhY2tncm91bmQ6IHZhcigtLXNlY3Rpb24tYmcpOyB9CiAgICB0ZCB7IHBhZGRpbmc6IDhweCAxMnB4OyB3aGl0ZS1zcGFjZTogbm93cmFwOyB9CiAgICB0ZC53cmFwIHsgd2hpdGUtc3BhY2U6IG5vcm1hbDsgbWF4LXdpZHRoOiAyNDBweDsgfQogICAgdGZvb3QgdGQgeyBmb250LXdlaWdodDogNzAwOyBiYWNrZ3JvdW5kOiB2YXIoLS1saWdodC1ibHVlKTsgY29sb3I6IHZhcigtLW5hdnkpOwogICAgICAgICAgICAgICBwYWRkaW5nOiA5cHggMTJweDsgYm9yZGVyLXRvcDogMnB4IHNvbGlkIHZhcigtLWJsdWUpOyB9CgogICAgLyog4pSA4pSAIENlbGwgaGVscGVycyDilIDilIAgKi8KICAgIC5udW0geyB0ZXh0LWFsaWduOiByaWdodDsgZm9udC12YXJpYW50LW51bWVyaWM6IHRhYnVsYXItbnVtczsgfQogICAgLnRhZyB7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgcGFkZGluZzogMnB4IDhweDsgYm9yZGVyLXJhZGl1czogNHB4OyBmb250LXNpemU6IDEwcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KICAgIC50YWctcmVkICAgIHsgYmFja2dyb3VuZDogdmFyKC0tcmVkLWxpZ2h0KTsgICBjb2xvcjogdmFyKC0tcmVkKTsgfQogICAgLnRhZy1hbWJlciAgeyBiYWNrZ3JvdW5kOiB2YXIoLS1hbWJlci1saWdodCk7IGNvbG9yOiB2YXIoLS1hbWJlcik7IH0KICAgIC50YWctZ3JlZW4gIHsgYmFja2dyb3VuZDogdmFyKC0tZ3JlZW4tbGlnaHQpOyBjb2xvcjogdmFyKC0tZ3JlZW4pOyB9CiAgICAudGFnLWJsdWUgICB7IGJhY2tncm91bmQ6IHZhcigtLWxpZ2h0LWJsdWUpOyAgY29sb3I6IHZhcigtLW5hdnkpOyB9CiAgICAuYmFyLWNlbGwgICB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyBtaW4td2lkdGg6IDEyMHB4OyB9CiAgICAuYmFyICAgICAgICB7IGZsZXg6IDE7IGhlaWdodDogNnB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiAzcHg7IG92ZXJmbG93OiBoaWRkZW47IH0KICAgIC5iYXItZmlsbCAgIHsgaGVpZ2h0OiAxMDAlOyBib3JkZXItcmFkaXVzOiAzcHg7IGJhY2tncm91bmQ6IHZhcigtLWJsdWUpOyB0cmFuc2l0aW9uOiB3aWR0aCAuM3M7IH0KICAgIC5iYXItZmlsbC5yZWQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1yZWQpOyB9CgogICAgLyog4pSA4pSAIFR3by1jb2wgbGF5b3V0IOKUgOKUgCAqLwogICAgLnR3by1jb2wgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IGdhcDogMjRweDsgbWFyZ2luLWJvdHRvbTogMjRweDsgfQogICAgQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC50d28tY29sIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQoKICAgIC8qIOKUgOKUgCBTLURyb3Agc3BlY2lmaWMg4pSA4pSAICovCiAgICAuc2Ryb3AtbG9jLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE2MHB4LDFmcikpOwogICAgICAgICAgICAgICAgICAgICAgZ2FwOiAxMnB4OyBwYWRkaW5nOiAxNnB4IDIwcHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9CiAgICAuc2Ryb3AtbG9jLWNhcmQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogOHB4OwogICAgICAgICAgICAgICAgICAgICAgcGFkZGluZzogMTRweCAxNnB4OyB9CiAgICAuc2Ryb3AtbG9jLW5hbWUgIHsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tbmF2eSk7CiAgICAgICAgICAgICAgICAgICAgICAgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbWFyZ2luLWJvdHRvbTogNnB4OyB9CiAgICAuc2Ryb3AtbG9jLWl0ZW1zIHsgZm9udC1zaXplOiAyMHB4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tbmF2eSk7IH0KICAgIC5zZHJvcC1sb2MtdmFsICAgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IG1hcmdpbi10b3A6IDJweDsgfQoKICAgIC8qIOKUgOKUgCBMb2FkaW5nIG92ZXJsYXkg4pSA4pSAICovCiAgICAjbG9hZGluZyB7IHBvc2l0aW9uOiBmaXhlZDsgaW5zZXQ6IDA7IGJhY2tncm91bmQ6IHJnYmEoMzEsNTYsMTAwLC44NSk7CiAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogICAgICAgICAgICAgICB6LWluZGV4OiA5OTk7IGNvbG9yOiAjZmZmOyBmb250LXNpemU6IDE1cHg7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KICAgIC5zcGlubmVyIHsgd2lkdGg6IDM2cHg7IGhlaWdodDogMzZweDsgYm9yZGVyOiAzcHggc29saWQgcmdiYSgyNTUsMjU1LDI1NSwuMyk7CiAgICAgICAgICAgICAgIGJvcmRlci10b3AtY29sb3I6ICNmZmY7IGJvcmRlci1yYWRpdXM6IDUwJTsgYW5pbWF0aW9uOiBzcGluIC44cyBsaW5lYXIgaW5maW5pdGU7CiAgICAgICAgICAgICAgIG1hcmdpbi1yaWdodDogMTZweDsgfQogICAgQGtleWZyYW1lcyBzcGluIHsgdG8geyB0cmFuc2Zvcm06IHJvdGF0ZSgzNjBkZWcpOyB9IH0KICAKLyog4pSA4pSAIE11bHRpLXNlbGVjdCBkcm9wZG93biDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAgKi8KLm1zLXdyYXAgeyBwb3NpdGlvbjpyZWxhdGl2ZTsgZGlzcGxheTppbmxpbmUtYmxvY2s7IG1pbi13aWR0aDoxODBweDsgfQoubXMtYnRuICB7IHdpZHRoOjEwMCU7IHBhZGRpbmc6N3B4IDEwcHg7IGJhY2tncm91bmQ6IzFhMmY1NDsgY29sb3I6I2ZmZjsKICAgICAgICAgICBib3JkZXI6MXB4IHNvbGlkIHJnYmEoMjU1LDI1NSwyNTUsLjI1KTsgYm9yZGVyLXJhZGl1czo2cHg7CiAgICAgICAgICAgY3Vyc29yOnBvaW50ZXI7IGZvbnQtc2l6ZToxM3B4OyB0ZXh0LWFsaWduOmxlZnQ7CiAgICAgICAgICAgZGlzcGxheTpmbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGp1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuOyBnYXA6NnB4OyB9Ci5tcy1idG46YWZ0ZXIgeyBjb250ZW50Oifilr4nOyBmb250LXNpemU6MTFweDsgb3BhY2l0eTouNzsgZmxleC1zaHJpbms6MDsgfQoubXMtYnRuLm9wZW46YWZ0ZXIgeyBjb250ZW50OifilrQnOyB9Ci5tcy1sYWJlbCB7IHdoaXRlLXNwYWNlOm5vd3JhcDsgb3ZlcmZsb3c6aGlkZGVuOyB0ZXh0LW92ZXJmbG93OmVsbGlwc2lzOyB9Ci5tcy1wYW5lbCB7IGRpc3BsYXk6bm9uZTsgcG9zaXRpb246YWJzb2x1dGU7IHRvcDpjYWxjKDEwMCUgKyA0cHgpOyBsZWZ0OjA7CiAgICAgICAgICAgIG1pbi13aWR0aDoxMDAlOyBtYXgtaGVpZ2h0OjIyMHB4OyBvdmVyZmxvdy15OmF1dG87CiAgICAgICAgICAgIGJhY2tncm91bmQ6I2ZmZjsgYm9yZGVyOjFweCBzb2xpZCAjZGRlOyBib3JkZXItcmFkaXVzOjhweDsKICAgICAgICAgICAgYm94LXNoYWRvdzowIDZweCAyMHB4IHJnYmEoMCwwLDAsLjE1KTsgei1pbmRleDo5OTk7IHBhZGRpbmc6NnB4IDA7IH0KLm1zLXBhbmVsLm9wZW4geyBkaXNwbGF5OmJsb2NrOyB9Ci5tcy1pdGVtIHsgZGlzcGxheTpmbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGdhcDo4cHg7IHBhZGRpbmc6N3B4IDEycHg7CiAgICAgICAgICAgZm9udC1zaXplOjEzcHg7IGNvbG9yOiMxRjM4NjQ7IGN1cnNvcjpwb2ludGVyOyB3aGl0ZS1zcGFjZTpub3dyYXA7IH0KLm1zLWl0ZW06aG92ZXIgeyBiYWNrZ3JvdW5kOiNmMGY0ZmY7IH0KLm1zLWl0ZW0gaW5wdXQgeyB3aWR0aDoxNHB4OyBoZWlnaHQ6MTRweDsgYWNjZW50LWNvbG9yOiMyRTc1QjY7IGN1cnNvcjpwb2ludGVyOyBmbGV4LXNocmluazowOyB9Ci5tcy1kaXZpZGVyIHsgaGVpZ2h0OjFweDsgYmFja2dyb3VuZDojZWVlOyBtYXJnaW46NHB4IDA7IH0KPC9zdHlsZT4KPC9oZWFkPgo8Ym9keT4KCjxkaXYgaWQ9ImxvYWRpbmciPjxkaXYgY2xhc3M9InNwaW5uZXIiPjwvZGl2PiBMb2FkaW5nIGRhc2hib2FyZOKApjwvZGl2PgoKPCEtLSBIZWFkZXIgLS0+CjxkaXYgY2xhc3M9ImhlYWRlciI+CiAgPGRpdiBjbGFzcz0iaGVhZGVyLWxlZnQiPgogICAgPGRpdj4KICAgICAgPGRpdiBzdHlsZT0iZm9udC1zaXplOjE3cHg7Zm9udC13ZWlnaHQ6NzAwOyI+QWdlZCBJbnZlbnRvcnkgRGFzaGJvYXJkPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImhlYWRlci1zdWIiIGlkPSJoZWFkZXJTdWIiPkFzIG9mIHRvZGF5PC9kaXY+CiAgICA8L2Rpdj4KICA8L2Rpdj4KICA8YSBjbGFzcz0idXBsb2FkLWxpbmsiIGhyZWY9Ii8iPuKGkSBVcGxvYWQgTmV3IERhdGE8L2E+CjwvZGl2PgoKPCEtLSBUYWIgbmF2aWdhdGlvbiAtLT4KPGRpdiBjbGFzcz0idGFiLW5hdiI+CiAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biBhY3RpdmUiIG9uY2xpY2s9InN3aXRjaFRhYignbWFpbicsIHRoaXMpIj5NYWluIERhc2hib2FyZDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InRhYi1idG4iIG9uY2xpY2s9InN3aXRjaFRhYignc2Ryb3AnLCB0aGlzKSIgaWQ9InNkcm9wVGFiQnRuIj4KICAgIFMtRHJvcCBSZXZpZXcgPHNwYW4gY2xhc3M9InRhYi1iYWRnZSIgaWQ9InNkcm9wQmFkZ2UiPuKAlDwvc3Bhbj4KICA8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBvbmNsaWNrPSJzd2l0Y2hUYWIoJ2Nsb3NlZCcsIHRoaXMpIiBpZD0iY2xvc2VkVGFiQnRuIj4KICAgIENsb3NlZCBPcmRlcnMgPHNwYW4gY2xhc3M9InRhYi1iYWRnZSIgc3R5bGU9ImJhY2tncm91bmQ6IzM3NTYyMyIgaWQ9ImNsb3NlZEJhZGdlIj7igJQ8L3NwYW4+CiAgPC9idXR0b24+CiAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgb25jbGljaz0ic3dpdGNoVGFiKCdvZmZzaXRlJywgdGhpcykiIGlkPSJvZmZzaXRlVGFiQnRuIj4KICAgIFN0b3JhZ2UgPHNwYW4gY2xhc3M9InRhYi1iYWRnZSIgc3R5bGU9ImJhY2tncm91bmQ6IzJFNzVCNiIgaWQ9Im9mZnNpdGVCYWRnZSI+4oCUPC9zcGFuPgogIDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InRhYi1idG4iIG9uY2xpY2s9InN3aXRjaFRhYignbmV3c3RvcmFnZScsIHRoaXMpIiBpZD0ibmV3c3RvcmFnZVRhYkJ0biI+CiAgICBOZXcgdG8gU3RvcmFnZSA8c3BhbiBjbGFzcz0idGFiLWJhZGdlIiBzdHlsZT0iYmFja2dyb3VuZDojOEIwMDAwIiBpZD0ibmV3c3RvcmFnZUJhZGdlIj7igJQ8L3NwYW4+CiAgPC9idXR0b24+CiAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgb25jbGljaz0ic3dpdGNoVGFiKCdzanBpJywgdGhpcykiIGlkPSJzanBpVGFiQnRuIj4KICAgIFN0b3JhZ2UgSm9icyBQYXJ0cyBJbmNvbXBsZXRlIDxzcGFuIGNsYXNzPSJ0YWItYmFkZ2UiIHN0eWxlPSJiYWNrZ3JvdW5kOiM3QjNGMDAiIGlkPSJzanBpQmFkZ2UiPuKAlDwvc3Bhbj4KICA8L2J1dHRvbj4KPC9kaXY+Cgo8IS0tIE1haW4gZGFzaGJvYXJkIGZpbHRlcnMgLS0+CjxkaXYgY2xhc3M9ImZpbHRlcnMtYmFyIiBpZD0ibWFpbkZpbHRlcnMiPgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IFByb2plY3Q8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLXByb2plY3QiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1wcm9qZWN0JywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgUHJvamVjdHM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1wcm9qZWN0LXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBDb29yZGluYXRvcjwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtY29vcmQiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1jb29yZCcsIGV2ZW50KVwiPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIENvb3JkaW5hdG9yczwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLWNvb3JkLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBBZ2UgQnVja2V0PC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1idWNrZXQiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1idWNrZXQnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBCdWNrZXRzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtYnVja2V0LXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIG9uY2xpY2s9InJlc2V0RmlsdGVycygpIj4mIzEwMDA1OyBSZXNldDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgc3R5bGU9ImJhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMik7Y29sb3I6I2ZmZjtib3JkZXItY29sb3I6cmdiYSgyNTUsMjU1LDI1NSwuNCk7IiBvbmNsaWNrPSJleHBvcnRNYWluQ1NWKCkiPiYjMTEwMTU7IEV4cG9ydCBDU1Y8L2J1dHRvbj4KPC9kaXY+Cgo8IS0tIFMtRHJvcCBmaWx0ZXJzIC0tPgo8ZGl2IGNsYXNzPSJmaWx0ZXJzLWJhciIgaWQ9InNkcm9wRmlsdGVycyIgc3R5bGU9ImRpc3BsYXk6bm9uZSI+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgRHJvcCBMb2NhdGlvbjwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtZHJvcGxvYyI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIiBvbmNsaWNrPVwibXNUb2dnbGUoJ21zLWRyb3Bsb2MnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBMb2NhdGlvbnM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1kcm9wbG9jLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBDb29yZGluYXRvcjwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtZHJvcGNvb3JkIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtZHJvcGNvb3JkJywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgQ29vcmRpbmF0b3JzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtZHJvcGNvb3JkLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBBZ2UgRmxhZzwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtZHJvcGZsYWciPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1kcm9wZmxhZycsIGV2ZW50KVwiPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIEFnZXM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1kcm9wZmxhZy1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGJ1dHRvbiBjbGFzcz0icmVzZXQtYnRuIiBvbmNsaWNrPSJyZXNldFNkcm9wRmlsdGVycygpIj4mIzEwMDA1OyBSZXNldDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgc3R5bGU9ImJhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMik7Y29sb3I6I2ZmZjtib3JkZXItY29sb3I6cmdiYSgyNTUsMjU1LDI1NSwuNCk7IiBvbmNsaWNrPSJleHBvcnRTZHJvcENTVigpIj4mIzExMDE1OyBFeHBvcnQgQ1NWPC9idXR0b24+CjwvZGl2PgoKPCEtLSBDbG9zZWQgT3JkZXJzIGZpbHRlcnMgLS0+CjxkaXYgY2xhc3M9ImZpbHRlcnMtYmFyIiBpZD0iY2xvc2VkRmlsdGVycyIgc3R5bGU9ImRpc3BsYXk6bm9uZSI+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQ29vcmRpbmF0b3I8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLWNsY29vcmQiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1jbGNvb3JkJywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgQ29vcmRpbmF0b3JzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtY2xjb29yZC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQWdlIEJ1Y2tldDwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtY2xidWNrZXQiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1jbGJ1Y2tldCcsIGV2ZW50KVwiPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIEJ1Y2tldHM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1jbGJ1Y2tldC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGJ1dHRvbiBjbGFzcz0icmVzZXQtYnRuIiBvbmNsaWNrPSJyZXNldENsb3NlZEZpbHRlcnMoKSI+JiMxMDAwNTsgUmVzZXQ8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIHN0eWxlPSJiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjIpO2NvbG9yOiNmZmY7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjQpOyIgb25jbGljaz0iZXhwb3J0Q2xvc2VkQ1NWKCkiPiYjMTEwMTU7IEV4cG9ydCBDU1Y8L2J1dHRvbj4KPC9kaXY+Cgo8IS0tIE9mZnNpdGUgU3RvcmFnZSBmaWx0ZXJzIC0tPgo8ZGl2IGNsYXNzPSJmaWx0ZXJzLWJhciIgaWQ9Im9mZnNpdGVGaWx0ZXJzIiBzdHlsZT0iZGlzcGxheTpub25lIj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBCdWlsZGluZzwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtb2ZidWlsZCI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIiBvbmNsaWNrPVwibXNUb2dnbGUoJ21zLW9mYnVpbGQnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBCdWlsZGluZ3M8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1vZmJ1aWxkLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBDb29yZGluYXRvcjwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtb2Zjb29yZCI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIiBvbmNsaWNrPVwibXNUb2dnbGUoJ21zLW9mY29vcmQnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBDb29yZGluYXRvcnM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1vZmNvb3JkLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBPcmRlciBTdGF0dXM8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLW9mc3RhdHVzIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtb2ZzdGF0dXMnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBTdGF0dXNlczwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLW9mc3RhdHVzLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBBZ2UgRmxhZzwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtb2ZmbGFnIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtb2ZmbGFnJywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgQWdlczwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLW9mZmxhZy1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgRGlzcG9zaXRpb248L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLW9mZGlzcCI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBEaXNwb3NpdGlvbnM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1vZmRpc3AtcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IEJpbGxlZCBTdGF0dXM8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLW9mYmlsbGVkIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtb2ZiaWxsZWQtcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgb25jbGljaz0icmVzZXRPZmZzaXRlRmlsdGVycygpIj4mIzEwMDA1OyBSZXNldDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgc3R5bGU9ImJhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMik7Y29sb3I6I2ZmZjtib3JkZXItY29sb3I6cmdiYSgyNTUsMjU1LDI1NSwuNCk7IiBvbmNsaWNrPSJleHBvcnRPZmZzaXRlQ1NWKCkiPiYjMTEwMTU7IEV4cG9ydCBDU1Y8L2J1dHRvbj4KPC9kaXY+Cgo8IS0tIE5ldyB0byBTdG9yYWdlIGZpbHRlcnMgLS0+CjxkaXYgY2xhc3M9ImZpbHRlcnMtYmFyIiBpZD0ibmV3c3RvcmFnZUZpbHRlcnMiIHN0eWxlPSJkaXNwbGF5Om5vbmUiPgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IENvb3JkaW5hdG9yPC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1uc2Nvb3JkIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIENvb3JkaW5hdG9yczwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLW5zY29vcmQtcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgb25jbGljaz0icmVzZXROZXdTdG9yYWdlRmlsdGVycygpIj4mIzEwMDA1OyBSZXNldDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgc3R5bGU9ImJhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMik7Y29sb3I6I2ZmZjtib3JkZXItY29sb3I6cmdiYSgyNTUsMjU1LDI1NSwuNCk7IiBvbmNsaWNrPSJleHBvcnROZXdTdG9yYWdlQ1NWKCkiPiYjMTEwMTU7IEV4cG9ydCBDU1Y8L2J1dHRvbj4KPC9kaXY+Cgo8IS0tIFN0b3JhZ2UgSm9icyBQYXJ0cyBJbmNvbXBsZXRlIGZpbHRlcnMgLS0+CjxkaXYgY2xhc3M9ImZpbHRlcnMtYmFyIiBpZD0ic2pwaUZpbHRlcnMiIHN0eWxlPSJkaXNwbGF5Om5vbmUiPgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IENvb3JkaW5hdG9yPC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1zanBpY29vcmQiPjxidXR0b24gY2xhc3M9Im1zLWJ0biI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgQ29vcmRpbmF0b3JzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtc2pwaWNvb3JkLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBMaW5lIFN0YXR1czwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtc2pwaXN0YXR1cyI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBTdGF0dXNlczwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLXNqcGlzdGF0dXMtcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgb25jbGljaz0icmVzZXRTanBpRmlsdGVycygpIj4mIzEwMDA1OyBSZXNldDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgc3R5bGU9ImJhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMik7Y29sb3I6I2ZmZjtib3JkZXItY29sb3I6cmdiYSgyNTUsMjU1LDI1NSwuNCk7IiBvbmNsaWNrPSJleHBvcnRTanBpQ1NWKCkiPiYjMTEwMTU7IEV4cG9ydCBDU1Y8L2J1dHRvbj4KPC9kaXY+Cgo8ZGl2IGNsYXNzPSJtYWluIj4KCiAgPCEtLSDilZDilZAgTUFJTiBUQUIg4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQIC0tPgogIDxkaXYgY2xhc3M9InRhYi1wYW5lbCBhY3RpdmUiIGlkPSJwYW5lbC1tYWluIj4KCiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCI+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlRvdGFsIEludmVudG9yeSBWYWx1ZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9ImtUb3RhbCI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiIgaWQ9ImtDb250YWluZXJzIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LXJlZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5PdmVyIDkwIERheXMgVmFsdWU8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJrT3ZlcjkwIj7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktc3ViIiBpZD0ia092ZXI5MFBjdCI+4oCUPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1hbWJlciI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5JbiBTdG9yYWdlIFZhbHVlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ia1N0b3JhZ2UiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+RmluYW5jZSBIb2xkIFZhbHVlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ia0ZpbmFuY2UiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtZ3JlZW4iPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+QXZlcmFnZSBBZ2U8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJrQXZnQWdlIj7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktc3ViIj5kYXlzPC9kaXY+PC9kaXY+CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJ0d28tY29sIj4KICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5JbnZlbnRvcnkgYnkgUHJvamVjdCBDb29yZGluYXRvcjwvZGl2PgogICAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1tZXRhIiBpZD0iY29vcmRNZXRhIj48L2Rpdj4KICAgICAgICA8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJ0YWJsZS13cmFwIj4KICAgICAgICAgIDx0YWJsZSBpZD0iY29vcmRUYWJsZSI+CiAgICAgICAgICAgIDx0aGVhZD48dHI+CiAgICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnY29vcmRUYWJsZScsMCkiPkNvb3JkaW5hdG9yPC90aD4KICAgICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdjb29yZFRhYmxlJywxKSI+Q29udGFpbmVyczwvdGg+CiAgICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnY29vcmRUYWJsZScsMikiPlZhbHVlPC90aD4KICAgICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdjb29yZFRhYmxlJywzKSI+T3ZlciA5MGQgVmFsdWU8L3RoPgogICAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nvb3JkVGFibGUnLDQpIj4lIE92ZXIgOTBkPC90aD4KICAgICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdjb29yZFRhYmxlJyw1KSI+QXZnIEFnZTwvdGg+CiAgICAgICAgICAgIDwvdHI+PC90aGVhZD4KICAgICAgICAgICAgPHRib2R5IGlkPSJjb29yZEJvZHkiPjwvdGJvZHk+CiAgICAgICAgICAgIDx0Zm9vdCBpZD0iY29vcmRGb290Ij48L3Rmb290PgogICAgICAgICAgPC90YWJsZT4KICAgICAgICA8L2Rpdj4KICAgICAgPC9kaXY+CgogICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWNhcmQiPgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24taGVhZGVyIj4KICAgICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPkludmVudG9yeSBieSBBZ2UgQnVja2V0PC9kaXY+CiAgICAgICAgPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0idGFibGUtd3JhcCI+CiAgICAgICAgICA8dGFibGUgaWQ9ImFnZVRhYmxlIj4KICAgICAgICAgICAgPHRoZWFkPjx0cj4KICAgICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdhZ2VUYWJsZScsMCkiPkFnZSBCdWNrZXQ8L3RoPgogICAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ2FnZVRhYmxlJywxKSI+Q29udGFpbmVyczwvdGg+CiAgICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnYWdlVGFibGUnLDIpIj5WYWx1ZTwvdGg+CiAgICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnYWdlVGFibGUnLDMpIj4lIG9mIFRvdGFsPC90aD4KICAgICAgICAgICAgPC90cj48L3RoZWFkPgogICAgICAgICAgICA8dGJvZHkgaWQ9ImFnZUJvZHkiPjwvdGJvZHk+CiAgICAgICAgICAgIDx0Zm9vdCBpZD0iYWdlRm9vdCI+PC90Zm9vdD4KICAgICAgICAgIDwvdGFibGU+CiAgICAgICAgPC9kaXY+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIj4KICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1oZWFkZXIiPgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPk9yZGVycyBpbiBWaWV3PC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1tZXRhIiBpZD0ib3JkZXJzTWV0YSI+PC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJzZWFyY2gtYm94Ij4KICAgICAgICA8aW5wdXQgdHlwZT0idGV4dCIgaWQ9Im9yZGVyU2VhcmNoIiBwbGFjZWhvbGRlcj0iU2VhcmNoIGJ5IG9yZGVyICMsIHByb2plY3QsIG9yIGNvb3JkaW5hdG9y4oCmIiBvbmlucHV0PSJyZW5kZXJPcmRlcnMoKSI+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJ0YWJsZS13cmFwIj4KICAgICAgICA8dGFibGUgaWQ9Im9yZGVyc1RhYmxlIj4KICAgICAgICAgIDx0aGVhZD48dHI+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29yZGVyc1RhYmxlJywwKSI+T3JkZXIgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29yZGVyc1RhYmxlJywxKSI+UHJvamVjdDwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29yZGVyc1RhYmxlJywyKSI+Q29vcmRpbmF0b3I8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdvcmRlcnNUYWJsZScsMykiPkNvbnRhaW5lcnM8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdvcmRlcnNUYWJsZScsNCkiPlZhbHVlPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnb3JkZXJzVGFibGUnLDUpIj5BdmcgQWdlPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnb3JkZXJzVGFibGUnLDYpIj5NYXggQWdlPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb3JkZXJzVGFibGUnLDcpIj5BZ2UgQnVja2V0PC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb3JkZXJzVGFibGUnLDgpIj5TaGlwIERhdGUgUmFuZ2U8L3RoPgogICAgICAgICAgPC90cj48L3RoZWFkPgogICAgICAgICAgPHRib2R5IGlkPSJvcmRlcnNCb2R5Ij48L3Rib2R5PgogICAgICAgICAgPHRmb290IGlkPSJvcmRlcnNGb290Ij48L3Rmb290PgogICAgICAgIDwvdGFibGU+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CgogIDwvZGl2PjwhLS0gL3BhbmVsLW1haW4gLS0+CgogIDwhLS0g4pWQ4pWQIFMtRFJPUCBUQUIg4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQIC0tPgogIDxkaXYgY2xhc3M9InRhYi1wYW5lbCIgaWQ9InBhbmVsLXNkcm9wIj4KCiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCIgc3R5bGU9ImdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDUsIDFmcik7Ij4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LXJlZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBJdGVtcyBpbiBEcm9wPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ic2RLcGlJdGVtcyI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiI+b3ZlciAyIGRheXM8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LXJlZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBWYWx1ZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9InNkS3BpVmFsdWUiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5VbmlxdWUgT3JkZXJzPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ic2RLcGlPcmRlcnMiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+QXZnIEFnZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9InNkS3BpQXZnIj7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktc3ViIj5kYXlzPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1hbWJlciI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5NYXggQWdlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ic2RLcGlNYXgiPuKAlDwvZGl2PjxkaXYgY2xhc3M9ImtwaS1zdWIiPmRheXM8L2Rpdj48L2Rpdj4KICAgIDwvZGl2PgoKICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24taGVhZGVyIj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5CeSBEcm9wIExvY2F0aW9uPC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtZ3JpZCIgaWQ9InNkcm9wTG9jR3JpZCI+PC9kaXY+CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+SXRlbSBEZXRhaWw8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLW1ldGEiIGlkPSJzZHJvcE1ldGEiPjwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic2VhcmNoLWJveCI+CiAgICAgICAgPGlucHV0IHR5cGU9InRleHQiIGlkPSJzZHJvcFNlYXJjaCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBieSBvcmRlciAjLCBwcm9qZWN0LCBwYXJ0IG5v4oCmIiBvbmlucHV0PSJyZW5kZXJTZHJvcCgpIj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InRhYmxlLXdyYXAiPgogICAgICAgIDx0YWJsZSBpZD0ic2Ryb3BUYWJsZSI+CiAgICAgICAgICA8dGhlYWQ+PHRyPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzZHJvcFRhYmxlJywwKSI+TG9jYXRpb248L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzZHJvcFRhYmxlJywxKSI+T3JkZXIgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDIpIj5Qcm9qZWN0PC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsMykiPkNvb3JkaW5hdG9yPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsNCkiPlBhcnQgTm8uPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsNSkiPlNlcmlhbCAjPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsNikiPlBhcnQgR3JvdXA8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdzZHJvcFRhYmxlJyw3KSI+UXR5PC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsOCkiPkFnZSAoZGF5cyk8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdzZHJvcFRhYmxlJyw5KSI+VmFsdWU8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzZHJvcFRhYmxlJywxMCkiPlNoaXAgRGF0ZTwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDExKSI+T3JkZXIgU3RhdHVzPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsMTIpIj5BZ2UgRmxhZzwvdGg+CiAgICAgICAgICA8L3RyPjwvdGhlYWQ+CiAgICAgICAgICA8dGJvZHkgaWQ9InNkcm9wQm9keSI+PC90Ym9keT4KICAgICAgICAgIDx0Zm9vdCBpZD0ic2Ryb3BGb290Ij48L3Rmb290PgogICAgICAgIDwvdGFibGU+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CgogIDwvZGl2PjwhLS0gL3BhbmVsLXNkcm9wIC0tPgoKICA8IS0tIOKVkOKVkCBDTE9TRUQgT1JERVJTIFRBQiDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZAgLS0+CiAgPGRpdiBjbGFzcz0idGFiLXBhbmVsIiBpZD0icGFuZWwtY2xvc2VkIj4KICAgIDxkaXYgY2xhc3M9ImtwaS1ncmlkIiBzdHlsZT0iZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOnJlcGVhdCg1LDFmcik7Ij4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWdyZWVuIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPkNsb3NlZCBJdGVtczwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9ImNsS3BpSXRlbXMiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtZ3JlZW4iPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VG90YWwgVmFsdWU8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJjbEtwaVZhbHVlIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VW5pcXVlIE9yZGVyczwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9ImNsS3BpT3JkZXJzIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWFtYmVyIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPkF2ZyBBZ2U8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJjbEtwaUF2ZyI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiI+ZGF5czwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+TWF4IEFnZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9ImNsS3BpTWF4Ij7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktc3ViIj5kYXlzPC9kaXY+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24taGVhZGVyIj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5DbG9zZWQgT3JkZXIgSXRlbXM8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLW1ldGEiIGlkPSJjbG9zZWRNZXRhIj48L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNlYXJjaC1ib3giPgogICAgICAgIDxpbnB1dCB0eXBlPSJ0ZXh0IiBpZD0iY2xvc2VkU2VhcmNoIiBwbGFjZWhvbGRlcj0iU2VhcmNoIGJ5IG9yZGVyICMsIHByb2plY3QsIHBhcnQgbm8uLi4iIG9uaW5wdXQ9InJlbmRlckNsb3NlZCgpIj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InRhYmxlLXdyYXAiPgogICAgICAgIDx0YWJsZSBpZD0iY2xvc2VkVGFibGUiPgogICAgICAgICAgPHRoZWFkPjx0cj4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDApIj5PcmRlciAjPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDEpIj5Qcm9qZWN0PC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDIpIj5Db29yZGluYXRvcjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJywzKSI+UGFydCBOby48L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdjbG9zZWRUYWJsZScsNCkiPlNlcmlhbCAjPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDUpIj5QYXJ0IEdyb3VwPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDYpIj5Mb2NhdGlvbjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJyw3KSI+U2hpcCBEYXRlPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDgpIj5RdHk8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdjbG9zZWRUYWJsZScsOSkiPkFnZSAoZGF5cyk8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdjbG9zZWRUYWJsZScsMTApIj5WYWx1ZTwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJywxMSkiPkFnZSBCdWNrZXQ8L3RoPgogICAgICAgICAgPC90cj48L3RoZWFkPgogICAgICAgICAgPHRib2R5IGlkPSJjbG9zZWRCb2R5Ij48L3Rib2R5PgogICAgICAgICAgPHRmb290IGlkPSJjbG9zZWRGb290Ij48L3Rmb290PgogICAgICAgIDwvdGFibGU+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgPC9kaXY+PCEtLSAvcGFuZWwtY2xvc2VkIC0tPgoKICA8IS0tIOKVkOKVkCBPRkZTSVRFIFNUT1JBR0UgVEFCIOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkCAtLT4KICA8ZGl2IGNsYXNzPSJ0YWItcGFuZWwiIGlkPSJwYW5lbC1vZmZzaXRlIj4KICAgIDxkaXYgY2xhc3M9ImtwaS1ncmlkIiBzdHlsZT0iZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOnJlcGVhdCg1LDFmcik7Ij4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VG90YWwgSXRlbXM8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJvZktwaUl0ZW1zIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LXJlZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBWYWx1ZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpVmFsdWUiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5VbmlxdWUgT3JkZXJzPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ib2ZLcGlPcmRlcnMiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+QXZnIEFnZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpQXZnIj7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktc3ViIj5kYXlzPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1hbWJlciI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5NYXggQWdlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ib2ZLcGlNYXgiPuKAlDwvZGl2PjxkaXYgY2xhc3M9ImtwaS1zdWIiPmRheXM8L2Rpdj48L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0ia3BpLWdyaWQiIHN0eWxlPSJncmlkLXRlbXBsYXRlLWNvbHVtbnM6cmVwZWF0KDMsMWZyKTttYXJnaW4tdG9wOjEycHg7Ij4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiIHN0eWxlPSJib3JkZXItdG9wOjRweCBzb2xpZCAjMkU3NUI2Ij4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlNPUCBCdWlsZCBBaGVhZDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpU09QIj7igJQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIiBpZD0ib2ZLcGlTT1BWYWwiIHN0eWxlPSJmb250LXNpemU6MTNweDtjb2xvcjojMkU3NUI2O2ZvbnQtd2VpZ2h0OjYwMCI+4oCUPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiI+b3JkZXJzIMK3IHZhbHVlPC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCIgc3R5bGU9ImJvcmRlci10b3A6NHB4IHNvbGlkICMzNzU2MjMiPgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+U3RvcmFnZSBDaGFyZ2VkPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ib2ZLcGlDaGFyZ2VkIj7igJQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIiBpZD0ib2ZLcGlDaGFyZ2VkVmFsIiBzdHlsZT0iZm9udC1zaXplOjEzcHg7Y29sb3I6IzM3NTYyMztmb250LXdlaWdodDo2MDAiPuKAlDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1zdWIiPm9yZGVycyDCtyB2YWx1ZTwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiIHN0eWxlPSJib3JkZXItdG9wOjRweCBzb2xpZCAjZGMzNTQ1Ij4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlVuYmlsbGVkIFN0b3JhZ2U8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJvZktwaVVuYmlsbGVkIj7igJQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIiBpZD0ib2ZLcGlVbmJpbGxlZFZhbCIgc3R5bGU9ImZvbnQtc2l6ZToxM3B4O2NvbG9yOiNkYzM1NDU7Zm9udC13ZWlnaHQ6NjAwIj7igJQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIj5vcmRlcnMgbm90IHlldCBjaGFyZ2VkPC9kaXY+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCIgc3R5bGU9ImdyaWQtdGVtcGxhdGUtY29sdW1uczpyZXBlYXQoMiwxZnIpO21hcmdpbi10b3A6MTJweDsiPgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCIgc3R5bGU9ImJvcmRlci10b3A6NHB4IHNvbGlkICM4QjQ1MTMiPgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+UHJvamVjdHMgQmVpbmcgQmlsbGVkPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ib2ZLcGlCaWxsZWRDb3VudCI+4oCUPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiI+c3RvcmFnZSBjaGFyZ2VzIHJhaXNlZDwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiIHN0eWxlPSJib3JkZXItdG9wOjRweCBzb2xpZCAjOEI0NTEzIj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlRvdGFsIEJpbGxlZDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpQmlsbGVkVG90YWwiPuKAlDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1zdWIiPmNoYXJnZSBhbW91bnQ8L2Rpdj4KICAgICAgPC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24taGVhZGVyIj48ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5CeSBMb2NhdGlvbiBHcm91cDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtZ3JpZCIgaWQ9Im9mZnNpdGVMb2NHcmlkIj48L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIj4KICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1oZWFkZXIiPgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPkl0ZW0gRGV0YWlsPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1tZXRhIiBpZD0ib2Zmc2l0ZU1ldGEiPjwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic2VhcmNoLWJveCI+CiAgICAgICAgPGlucHV0IHR5cGU9InRleHQiIGlkPSJvZmZzaXRlU2VhcmNoIiBwbGFjZWhvbGRlcj0iU2VhcmNoIGJ5IG9yZGVyICMsIHByb2plY3QsIHBhcnQgbm8uLi4iIG9uaW5wdXQ9InJlbmRlck9mZnNpdGUoKSI+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJ0YWJsZS13cmFwIj4KICAgICAgICA8dGFibGUgaWQ9Im9mZnNpdGVUYWJsZSI+CiAgICAgICAgICA8dGhlYWQ+PHRyPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDApIj5CdWlsZGluZzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMSkiPkxvY2F0aW9uIEdyb3VwPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywyKSI+T3JkZXIgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMykiPlByb2plY3Q8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDQpIj5Db29yZGluYXRvcjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsNSkiPkxvY2F0aW9uPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJyw2KSI+UGFydCBOby48L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDcpIj5TZXJpYWwgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsOCkiPlNoaXAgRGF0ZTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsOSkiPkNvbnRhaW5lcnM8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDEwKSI+VmFsdWU8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDExKSI+QXZnIEFnZTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTIpIj5NYXggQWdlPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywxMykiPk9yZGVyIFN0YXR1czwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTQpIj5BZ2UgRmxhZzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTUpIj5DaGFyZ2UgQW1vdW50PC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywxNikiPkNoYXJnZSBUeXBlPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywxNykiPkJpbGxpbmcgTm90ZTwvdGg+CiAgICAgICAgICAgIDx0aD5EaXNwb3NpdGlvbjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTYpIj5CaWxsZWQ/PC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywxNykiPkNoYXJnZSBBbXQ8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDE4KSI+Q2hhcmdlIFR5cGU8L3RoPgogICAgICAgICAgPC90cj48L3RoZWFkPgogICAgICAgICAgPHRib2R5IGlkPSJvZmZzaXRlQm9keSI+PC90Ym9keT4KICAgICAgICAgIDx0Zm9vdCBpZD0ib2Zmc2l0ZUZvb3QiPjwvdGZvb3Q+CiAgICAgICAgPC90YWJsZT4KICAgICAgPC9kaXY+CiAgICA8L2Rpdj4KICA8L2Rpdj48IS0tIC9wYW5lbC1vZmZzaXRlIC0tPgoKICA8IS0tIOKVkOKVkCBORVcgVE8gU1RPUkFHRSBUQUIg4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQIC0tPgogIDxkaXYgY2xhc3M9InRhYi1wYW5lbCIgaWQ9InBhbmVsLW5ld3N0b3JhZ2UiPgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIiBzdHlsZT0ibWFyZ2luLWJvdHRvbToxNnB4O2JhY2tncm91bmQ6I2ZmZjhmODtib3JkZXI6MXB4IHNvbGlkICNmNWM2YzY7Ij4KICAgICAgPGRpdiBzdHlsZT0icGFkZGluZzoxNnB4IDIwcHg7Zm9udC1zaXplOjEzcHg7Y29sb3I6IzdCMDAwMDtsaW5lLWhlaWdodDoxLjYiPgogICAgICAgIDxzdHJvbmc+JiMxMjg2ODA7IFdoYXQgaXMgdGhpcyB0YWI/PC9zdHJvbmc+PGJyPgogICAgICAgIFRoaXMgdGFiIHNob3dzIG9yZGVycyB3aG9zZSBzdGF0dXMgPHN0cm9uZz5jaGFuZ2VkIHRvIHN0YXJ0IHdpdGggIlN0b3JhZ2UiPC9zdHJvbmc+IHNpbmNlIHRoZSBwcmV2aW91cyB1cGxvYWQuCiAgICAgICAgSXQgaXMgYSBjdW11bGF0aXZlIGxvZyDigJQgZW50cmllcyByZW1haW4gaGVyZSBhY3Jvc3MgZnV0dXJlIHVwbG9hZHMgc28geW91IGNhbiB0cmFjayB3aGVuIGl0ZW1zIGZpcnN0IG1vdmVkIGludG8gc3RvcmFnZS4KICAgICAgICBVcGxvYWQgbmV3IGZpbGVzIGRhaWx5IHRvIGtlZXAgdGhpcyBsaXN0IGN1cnJlbnQuCiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCIgc3R5bGU9ImdyaWQtdGVtcGxhdGUtY29sdW1uczpyZXBlYXQoMywxZnIpOyI+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIiBzdHlsZT0iYm9yZGVyLXRvcDo0cHggc29saWQgIzhCMDAwMCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5OZXcgVGhpcyBVcGxvYWQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJuc0twaU5ldyI+4oCUPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlRvdGFsIFRyYWNrZWQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJuc0twaVRvdGFsIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LXJlZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBWYWx1ZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im5zS3BpVmFsdWUiPuKAlDwvZGl2PjwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+T3JkZXJzIE1vdmVkIHRvIFN0b3JhZ2U8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLW1ldGEiIGlkPSJuc01ldGEiPjwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic2VhcmNoLWJveCI+CiAgICAgICAgPGlucHV0IHR5cGU9InRleHQiIGlkPSJuc1NlYXJjaCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBieSBvcmRlciAjLCBwcm9qZWN0LCBjb29yZGluYXRvci4uLiIgb25pbnB1dD0icmVuZGVyTmV3U3RvcmFnZSgpIj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InRhYmxlLXdyYXAiPgogICAgICAgIDx0YWJsZSBpZD0ibnNUYWJsZSI+CiAgICAgICAgICA8dGhlYWQ+PHRyPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCduc1RhYmxlJywwKSI+RGV0ZWN0ZWQgRGF0ZTwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ25zVGFibGUnLDEpIj5PcmRlciAjPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsMikiPlByb2plY3Q8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCduc1RhYmxlJywzKSI+Q29vcmRpbmF0b3I8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCduc1RhYmxlJyw0KSI+UHJldmlvdXMgU3RhdHVzPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsNSkiPk5ldyBTdGF0dXM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCduc1RhYmxlJyw2KSI+U2hpcCBEYXRlPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsNykiPkNvbnRhaW5lcnM8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCduc1RhYmxlJyw4KSI+VmFsdWU8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCduc1RhYmxlJyw5KSI+QXZnIEFnZTwvdGg+CiAgICAgICAgICA8L3RyPjwvdGhlYWQ+CiAgICAgICAgICA8dGJvZHkgaWQ9Im5zQm9keSI+PC90Ym9keT4KICAgICAgICAgIDx0Zm9vdCBpZD0ibnNGb290Ij48L3Rmb290PgogICAgICAgIDwvdGFibGU+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgPC9kaXY+PCEtLSAvcGFuZWwtbmV3c3RvcmFnZSAtLT4KCiAgPCEtLSDilZDilZAgU1RPUkFHRSBKT0JTIFBBUlRTIElOQ09NUExFVEUg4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQIC0tPgogIDxkaXYgY2xhc3M9InRhYi1wYW5lbCIgaWQ9InBhbmVsLXNqcGkiPgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIiBzdHlsZT0ibWFyZ2luLWJvdHRvbToxNnB4O2JhY2tncm91bmQ6I2ZmZmFmNTtib3JkZXI6MXB4IHNvbGlkICNlOGM5OWE7Ij4KICAgICAgPGRpdiBzdHlsZT0icGFkZGluZzoxNnB4IDIwcHg7Zm9udC1zaXplOjEzcHg7Y29sb3I6IzdCM0YwMDtsaW5lLWhlaWdodDoxLjYiPgogICAgICAgIDxzdHJvbmc+JiM5ODg4OyBXaGF0IGlzIHRoaXMgdGFiPzwvc3Ryb25nPjxicj4KICAgICAgICBTaG93cyBwcm9qZWN0cyBmbGFnZ2VkIGFzIDxzdHJvbmc+U3RvcmFnZSBDaGFyZ2VkPC9zdHJvbmc+IChzZXQgb24gdGhlIFN0b3JhZ2UgdGFiKSB0aGF0IGhhdmUKICAgICAgICBvcGVuIGxpbmVzIG9uIHRoZSBTaGlwIFNjaGVkdWxlIHdoZXJlIHRoZSBwYXJ0IGlzIDxzdHJvbmc+bm90IHlldCBpbiB3YXJlaG91c2UgaW52ZW50b3J5PC9zdHJvbmc+CiAgICAgICAg4oCUIG1lYW5pbmcgd2UgYXJlIGNoYXJnaW5nIHN0b3JhZ2UgYnV0IG5vdCBhbGwgcGFydHMgYXJlIGNvbXBsZXRlLiBTZXQgZGlzcG9zaXRpb25zIG9uIHRoZQogICAgICAgIFN0b3JhZ2UgdGFiIGZpcnN0LCB0aGVuIHVwbG9hZCB0aGUgU2hpcCBTY2hlZHVsZSBDU1YgdG8gcG9wdWxhdGUgdGhpcyB0YWIuCiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCIgc3R5bGU9ImdyaWQtdGVtcGxhdGUtY29sdW1uczpyZXBlYXQoNCwxZnIpOyI+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIiBzdHlsZT0iYm9yZGVyLXRvcDo0cHggc29saWQgIzdCM0YwMCI+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5JbmNvbXBsZXRlIFN0b3JhZ2UgSm9iczwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9InNqcGlLcGlQcm9qZWN0cyI+4oCUPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiI+dW5pcXVlIHByb2plY3RzPC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtcmVkIj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktbGFiZWwiPk9wZW4gQmFja29yZGVyIExpbmVzPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ic2pwaUtwaUxpbmVzIj7igJQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIj5zaGlwIHNjaGVkdWxlIGxpbmVzIG1pc3NpbmcgZnJvbSBpbnZlbnRvcnk8L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1hbWJlciI+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBTdG9yYWdlIFZhbHVlPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ic2pwaUtwaVN0b3JhZ2VWYWwiPuKAlDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1zdWIiPmludmVudG9yeSBpbiBzdG9yYWdlPC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+T3BlbiBRdHkgKG5vdCBzaGlwcGVkKTwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9InNqcGlLcGlVbnNoaXBwZWQiPuKAlDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1zdWIiPnF0eSBvbiBzY2hlZHVsZSBtaW51cyBzaGlwcGVkPC9kaXY+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+SW5jb21wbGV0ZSBKb2JzIERldGFpbDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tbWV0YSIgaWQ9InNqcGlNZXRhIj48L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNlYXJjaC1ib3giPgogICAgICAgIDxpbnB1dCB0eXBlPSJ0ZXh0IiBpZD0ic2pwaVNlYXJjaCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBieSBwcm9qZWN0LCBwYXJ0IG5vLCBjb29yZGluYXRvci4uLiIgb25pbnB1dD0icmVuZGVyU2pwaSgpIj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InRhYmxlLXdyYXAiPgogICAgICAgIDx0YWJsZSBpZD0ic2pwaVRhYmxlIj4KICAgICAgICAgIDx0aGVhZD48dHI+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3NqcGlUYWJsZScsMCkiPlByb2plY3QgKFN0b3JhZ2UpPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2pwaVRhYmxlJywxKSI+U3RvcmFnZSBTdGF0dXM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzanBpVGFibGUnLDIpIj5Db29yZGluYXRvcjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3NqcGlUYWJsZScsMykiPlNTIE9yZGVyICM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzanBpVGFibGUnLDQpIj5DdXN0b21lciBOYW1lPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2pwaVRhYmxlJyw1KSI+UGFydCBOby48L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzanBpVGFibGUnLDYpIj5SZXZpc2lvbjwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ3NqcGlUYWJsZScsNykiPlF0eTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ3NqcGlUYWJsZScsOCkiPlF0eSBTaGlwcGVkPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2pwaVRhYmxlJyw5KSI+TGluZSBTdGF0dXM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzanBpVGFibGUnLDEwKSI+T3JkZXIgU3RhdHVzPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnc2pwaVRhYmxlJywxMSkiPlN0b3JhZ2UgVmFsdWU8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdzanBpVGFibGUnLDEyKSI+QXZnIEFnZTwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3NqcGlUYWJsZScsMTMpIj5TaGlwIERhdGU8L3RoPgogICAgICAgICAgPC90cj48L3RoZWFkPgogICAgICAgICAgPHRib2R5IGlkPSJzanBpQm9keSI+PC90Ym9keT4KICAgICAgICAgIDx0Zm9vdCBpZD0ic2pwaUZvb3QiPjwvdGZvb3Q+CiAgICAgICAgPC90YWJsZT4KICAgICAgPC9kaXY+CiAgICA8L2Rpdj4KICA8L2Rpdj48IS0tIC9wYW5lbC1zanBpIC0tPgoKPC9kaXY+PCEtLSAvbWFpbiAtLT4KCjxzY3JpcHQ+Ci8vIOKUgOKUgCBHbG9iYWxzIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApsZXQgUkFXID0gbnVsbDsKbGV0IGZpbHRlcmVkT3JkZXJzID0gW107CmNvbnN0IHNvcnRTdGF0ZSA9IHt9OwoKY29uc3QgQUdFX09SREVSICA9IFsnMC03ZCcsJzgtMTRkJywnMTUtMzBkJywnMzEtNjBkJywnNjEtOTBkJywnOTEtMTgwZCcsJzE4MS0zNjVkJywnMzY1ZCsnXTsKY29uc3QgQUdFX0xBQkVMUyA9IHsnMC03ZCc6JzDigJM3ZCcsJzgtMTRkJzonOOKAkzE0ZCcsJzE1LTMwZCc6JzE14oCTMzBkJywnMzEtNjBkJzonMzHigJM2MGQnLAogICc2MS05MGQnOic2MeKAkzkwZCcsJzkxLTE4MGQnOic5MeKAkzE4MGQnLCcxODEtMzY1ZCc6JzE4MeKAkzM2NWQnLCczNjVkKyc6Jz4zNjVkJ307CgovLyDilIDilIAgRm9ybWF0dGVycyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKY29uc3QgZm10JCAgPSB2ID0+IHYgPT0gbnVsbCA/ICfigJQnIDogJyQnICsgTWF0aC5yb3VuZCh2KS50b0xvY2FsZVN0cmluZygpOwpjb25zdCBmbXROICA9IHYgPT4gdiA9PSBudWxsID8gJ+KAlCcgOiBNYXRoLnJvdW5kKHYpLnRvTG9jYWxlU3RyaW5nKCk7CmNvbnN0IGZtdFBjdD0gdiA9PiB2ID09IG51bGwgPyAn4oCUJyA6ICh2KjEwMCkudG9GaXhlZCgxKSsnJSc7CmNvbnN0IGZtdEYgID0gdiA9PiB2ID09IG51bGwgPyAn4oCUJyA6IHYudG9GaXhlZCgxKTsKCmZ1bmN0aW9uIGFnZUJ1Y2tldFRhZyhjb2RlKSB7CiAgaWYgKCFjb2RlKSByZXR1cm4gJyc7CiAgY29uc3QgY2xzID0gKGNvZGU9PT0nMzY1ZCsnfHxjb2RlPT09JzE4MS0zNjVkJykgPyAndGFnLXJlZCcKICAgICAgICAgICAgOiBjb2RlPT09JzkxLTE4MGQnID8gJ3RhZy1hbWJlcicgOiAndGFnLWJsdWUnOwogIHJldHVybiBgPHNwYW4gY2xhc3M9InRhZyAke2Nsc30iPiR7QUdFX0xBQkVMU1tjb2RlXXx8Y29kZX08L3NwYW4+YDsKfQpmdW5jdGlvbiBtYXhBZ2VUYWcoYWdlKSB7CiAgY29uc3QgY2xzID0gYWdlPjkwID8gJ3RhZy1yZWQnIDogYWdlPjMwID8gJ3RhZy1hbWJlcicgOiAndGFnLWdyZWVuJzsKICByZXR1cm4gYDxzcGFuIGNsYXNzPSJ0YWcgJHtjbHN9Ij4ke2FnZX1kPC9zcGFuPmA7Cn0KZnVuY3Rpb24gc2Ryb3BBZ2VGbGFnKGFnZSkgewogIGlmIChhZ2UgPD0gMikgIHJldHVybiAnMiBEYXlzIG9yIExlc3MnOwogIGlmIChhZ2UgPD0gNSkgIHJldHVybiAnMuKAkzUgRGF5cyc7CiAgaWYgKGFnZSA8PSAxMCkgcmV0dXJuICc24oCTMTAgRGF5cyc7CiAgcmV0dXJuICcxMCsgRGF5cyc7Cn0KZnVuY3Rpb24gc2Ryb3BGbGFnVGFnKGFnZSkgewogIGNvbnN0IGYgPSBzZHJvcEFnZUZsYWcoYWdlKTsKICBjb25zdCBjbHMgPSBhZ2UgPiAxMCA/ICd0YWctcmVkJyA6IGFnZSA+IDUgPyAndGFnLWFtYmVyJyA6IGFnZSA+IDIgPyAndGFnLWdyZWVuJyA6ICcnOwogIHJldHVybiBgPHNwYW4gY2xhc3M9InRhZyAke2Nsc30iPiR7Zn08L3NwYW4+YDsKfQpmdW5jdGlvbiBmbGFnVGFnKGZsYWcpIHsKICBjb25zdCBjbHMgPSBmbGFnPT09Jz45MCBEYXlzJyA/ICd0YWctcmVkJyA6IGZsYWc9PT0nPjMwIERheXMnID8gJ3RhZy1hbWJlcicgOiAndGFnLWdyZWVuJzsKICByZXR1cm4gYDxzcGFuIGNsYXNzPSJ0YWcgJHtjbHN9Ij4ke2ZsYWd9PC9zcGFuPmA7Cn0KCi8vIOKUgOKUgCBUYWIgc3dpdGNoaW5nIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiBzd2l0Y2hUYWIobmFtZSwgYnRuKSB7CiAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYi1wYW5lbCcpLmZvckVhY2gocCA9PiBwLmNsYXNzTGlzdC5yZW1vdmUoJ2FjdGl2ZScpKTsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcudGFiLWJ0bicpLmZvckVhY2goYiA9PiBiLmNsYXNzTGlzdC5yZW1vdmUoJ2FjdGl2ZScpKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgncGFuZWwtJyArIG5hbWUpLmNsYXNzTGlzdC5hZGQoJ2FjdGl2ZScpOwogIGJ0bi5jbGFzc0xpc3QuYWRkKCdhY3RpdmUnKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWFpbkZpbHRlcnMnKS5zdHlsZS5kaXNwbGF5ICAgPSBuYW1lPT09J21haW4nICAgID8gJycgOiAnbm9uZSc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Nkcm9wRmlsdGVycycpLnN0eWxlLmRpc3BsYXkgID0gbmFtZT09PSdzZHJvcCcgICA/ICcnIDogJ25vbmUnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbG9zZWRGaWx0ZXJzJykuc3R5bGUuZGlzcGxheSA9IG5hbWU9PT0nY2xvc2VkJyAgPyAnJyA6ICdub25lJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUZpbHRlcnMnKS5zdHlsZS5kaXNwbGF5ICAgID0gbmFtZT09PSdvZmZzaXRlJyAgICAgPyAnJyA6ICdub25lJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbmV3c3RvcmFnZUZpbHRlcnMnKS5zdHlsZS5kaXNwbGF5ID0gbmFtZT09PSduZXdzdG9yYWdlJyAgPyAnJyA6ICdub25lJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2pwaUZpbHRlcnMnKS5zdHlsZS5kaXNwbGF5ICAgICAgID0gbmFtZT09PSdzanBpJyAgICAgICAgPyAnJyA6ICdub25lJzsKICBpZiAobmFtZSA9PT0gJ25ld3N0b3JhZ2UnICYmICFOU19MT0FERUQpIGxvYWROZXdTdG9yYWdlKCk7CiAgaWYgKG5hbWUgPT09ICdzanBpJyAgICAgICAmJiAhU0pQSV9MT0FERUQpIGxvYWRTanBpKCk7Cn0KCi8vIOKUgOKUgCBMb2FkIGRhdGEg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmFzeW5jIGZ1bmN0aW9uIGxvYWREYXRhKCkgewogIGF3YWl0IGxvYWREaXNwb3NpdGlvbnMoKTsKICBhd2FpdCBsb2FkQmlsbGluZygpOwogIHRyeSB7CiAgICAvLyAxMCBzZWNvbmQgdGltZW91dCDigJQgaWYgc2VydmVyIGRvZXNuJ3QgcmVzcG9uZCwgc2hvdyBlcnJvcgogICAgY29uc3QgY29udHJvbGxlciA9IG5ldyBBYm9ydENvbnRyb2xsZXIoKTsKICAgIGNvbnN0IHRpbWVvdXQgPSBzZXRUaW1lb3V0KCgpID0+IGNvbnRyb2xsZXIuYWJvcnQoKSwgMTAwMDApOwogICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goJy9hcGkvZGF0YScsIHsgc2lnbmFsOiBjb250cm9sbGVyLnNpZ25hbCB9KTsKICAgIGNsZWFyVGltZW91dCh0aW1lb3V0KTsKICAgIGlmICghcmVzLm9rKSB0aHJvdyBuZXcgRXJyb3IoJ25vX2RhdGEnKTsKICAgIFJBVyA9IGF3YWl0IHJlcy5qc29uKCk7CiAgfSBjYXRjaChlKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbG9hZGluZycpLmlubmVySFRNTCA9IGAKICAgICAgPGRpdiBzdHlsZT0idGV4dC1hbGlnbjpjZW50ZXI7cGFkZGluZzo0MHB4Ij4KICAgICAgICA8ZGl2IHN0eWxlPSJmb250LXNpemU6NDhweDttYXJnaW4tYm90dG9tOjE2cHgiPvCfk4I8L2Rpdj4KICAgICAgICA8ZGl2IHN0eWxlPSJmb250LXNpemU6MjBweDtmb250LXdlaWdodDo3MDA7bWFyZ2luLWJvdHRvbTo4cHgiPk5vIGRhdGEgbG9hZGVkIHlldDwvZGl2PgogICAgICAgIDxkaXYgc3R5bGU9ImZvbnQtc2l6ZToxNHB4O29wYWNpdHk6Ljg7bWFyZ2luLWJvdHRvbToyNHB4Ij5VcGxvYWQgeW91ciB3YXJlaG91c2UgYW5kIGludGVybmFsIG5vdGVzIGZpbGVzIHRvIGdldCBzdGFydGVkLjwvZGl2PgogICAgICAgIDxhIGhyZWY9Ii8iIHN0eWxlPSJiYWNrZ3JvdW5kOiNmZmY7Y29sb3I6IzFGMzg2NDtwYWRkaW5nOjEycHggMjhweDtib3JkZXItcmFkaXVzOjhweDt0ZXh0LWRlY29yYXRpb246bm9uZTtmb250LXdlaWdodDo3MDA7Zm9udC1zaXplOjE0cHgiPuKGkSBVcGxvYWQgRmlsZXM8L2E+CiAgICAgIDwvZGl2PmA7CiAgICByZXR1cm47CiAgfQoKICBpZiAoIVJBVy5zZHJvcCkgUkFXLnNkcm9wID0geyBrcGlzOnt9LCBieV9sb2NhdGlvbjpbXSwgaXRlbXM6W10gfTsKCiAgLy8gU2V0IHRoZSB1cGxvYWRlZCBkYXRlIGluIHRoZSBoZWFkZXIKICBpZiAoUkFXLnVwbG9hZGVkKSB7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnaGVhZGVyU3ViJykudGV4dENvbnRlbnQgPSAnQXMgb2YgJyArIFJBVy51cGxvYWRlZDsKICB9CgogIC8vIFVwZGF0ZSB0YWIgYmFkZ2VzCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Nkcm9wQmFkZ2UnKS50ZXh0Q29udGVudCAgID0gUkFXLnNkcm9wPy5rcGlzPy50b3RhbF9pdGVtcyAgID8/IDA7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nsb3NlZEJhZGdlJykudGV4dENvbnRlbnQgID0gUkFXLmNsb3NlZD8ua3Bpcz8udG90YWxfaXRlbXMgID8/IDA7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29mZnNpdGVCYWRnZScpLnRleHRDb250ZW50ID0gUkFXLm9mZnNpdGU/LmtwaXM/LnRvdGFsX2l0ZW1zID8/IDA7CgogIC8vIFBvcHVsYXRlIGFsbCBtdWx0aS1zZWxlY3QgZmlsdGVycwogIG1zUG9wdWxhdGUoJ21zLXByb2plY3QnLCBSQVcucHJvamVjdHMgfHwgW10sIGFwcGx5RmlsdGVycyk7CiAgY29uc3QgbWFpbkNvb3JkcyA9IFsuLi5uZXcgU2V0KChSQVcub3JkZXJzX3RhYmxlfHxbXSkubWFwKHIgPT4gci5jb29yZGluYXRvcikpXS5maWx0ZXIoQm9vbGVhbikuc29ydCgpOwogIG1zUG9wdWxhdGUoJ21zLWNvb3JkJywgbWFpbkNvb3JkcywgYXBwbHlGaWx0ZXJzKTsKICBtc1BvcHVsYXRlKCdtcy1idWNrZXQnLCAgWycwLTdkJywnOC0xNGQnLCcxNS0zMGQnLCczMS02MGQnLCc2MS05MGQnLCc5MS0xODBkJywnMTgxLTM2NWQnLCczNjVkKyddLCBhcHBseUZpbHRlcnMpOwoKICBjb25zdCBzZHJvcENvb3JkcyA9IFsuLi5uZXcgU2V0KChSQVcuc2Ryb3A/Lml0ZW1zfHxbXSkubWFwKHI9PnIuY29vcmRpbmF0b3IpKV0uZmlsdGVyKEJvb2xlYW4pLnNvcnQoKTsKICBtc1BvcHVsYXRlKCdtcy1kcm9wbG9jJywgICBSQVcuZHJvcF9sb2NhdGlvbnMgfHwgW10sIHJlbmRlclNkcm9wKTsKICBtc1BvcHVsYXRlKCdtcy1kcm9wY29vcmQnLCBzZHJvcENvb3JkcywgcmVuZGVyU2Ryb3ApOwogIG1zUG9wdWxhdGUoJ21zLWRyb3BmbGFnJywgIFsnMiBEYXlzIG9yIExlc3MnLCcy4oCTNSBEYXlzJywnNuKAkzEwIERheXMnLCcxMCsgRGF5cyddLCByZW5kZXJTZHJvcCk7CgogIGNvbnN0IGNsQ29vcmRzID0gWy4uLm5ldyBTZXQoKFJBVy5jbG9zZWQ/Lml0ZW1zfHxbXSkubWFwKHIgPT4gci5jb29yZGluYXRvcikpXS5maWx0ZXIoQm9vbGVhbikuc29ydCgpOwogIG1zUG9wdWxhdGUoJ21zLWNsY29vcmQnLCBjbENvb3JkcywgcmVuZGVyQ2xvc2VkKTsKICBtc1BvcHVsYXRlKCdtcy1jbGJ1Y2tldCcsIFsnMC03ZCcsJzgtMTRkJywnMTUtMzBkJywnMzEtNjBkJywnNjEtOTBkJywnOTEtMTgwZCcsJzE4MS0zNjVkJywnMzY1ZCsnXSwgcmVuZGVyQ2xvc2VkKTsKCiAgbXNQb3B1bGF0ZSgnbXMtb2ZidWlsZCcsICBbJ0Rvd25zdGFpcnMgU3RvcmFnZScsJ1BvcnRlciBTdG9yYWdlJywnTm9ydGggV2FyZWhvdXNlJywnVHJhaWxlcnMnXSwgcmVuZGVyT2Zmc2l0ZSk7CiAgLy8gRGVyaXZlIGNvb3JkaW5hdG9ycyBmcm9tIGl0ZW1zIGRpcmVjdGx5IHNvIFVuYXNzaWduZWQgaXMgYWx3YXlzIGluY2x1ZGVkCiAgY29uc3Qgb2ZDb29yZHMgPSBbLi4ubmV3IFNldCgoUkFXLm9mZnNpdGU/Lml0ZW1zfHxbXSkubWFwKHIgPT4gci5jb29yZGluYXRvcikpXS5maWx0ZXIoQm9vbGVhbikuc29ydCgpOwogIG1zUG9wdWxhdGUoJ21zLW9mY29vcmQnLCBvZkNvb3JkcywgcmVuZGVyT2Zmc2l0ZSk7CiAgbXNQb3B1bGF0ZSgnbXMtb2ZzdGF0dXMnLCBSQVcub2Zmc2l0ZT8uc3RhdHVzZXMgfHwgW10sIHJlbmRlck9mZnNpdGUpOwogIG1zUG9wdWxhdGUoJ21zLW9mZmxhZycsICAgWyc+OTAgRGF5cycsJz4zMCBEYXlzJywn4omkMzAgRGF5cyddLCByZW5kZXJPZmZzaXRlKTsKICBtc1BvcHVsYXRlKCdtcy1vZmRpc3AnLCAgIFsnU09QIEJ1aWxkIEFoZWFkJywnU3RvcmFnZSBDaGFyZ2VkJywnSW5zdGFsbGF0aW9uIERlbGF5JywnKE5vbmUpJ10sIHJlbmRlck9mZnNpdGUpOwogIG1zUG9wdWxhdGUoJ21zLW9mYmlsbGVkJywgWydCaWxsZWQnLCdOb3QgQmlsbGVkJ10sIHJlbmRlck9mZnNpdGUpOwoKICAvLyBSZW5kZXIgbmV3IHRhYnMKICByZW5kZXJDbG9zZWRLcGlzKCk7CiAgcmVuZGVyT2Zmc2l0ZUtwaXMoKTsKICByZW5kZXJPZmZzaXRlTG9jR3JpZCgpOwoKICBhcHBseUZpbHRlcnMoKTsKICByZW5kZXJTZHJvcEtwaXMoKTsKICByZW5kZXJTZHJvcExvY0dyaWQoKTsKICByZW5kZXJTZHJvcCgpOwogIHJlbmRlckNsb3NlZCgpOwogIHJlbmRlck9mZnNpdGUoKTsKICB1cGRhdGVEaXNwb3NpdGlvbktQSXMoKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbG9hZGluZycpLnN0eWxlLmRpc3BsYXkgPSAnbm9uZSc7Cn0KCi8vIOKUgOKUgCBNYWluIGZpbHRlcnMg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIGdldEZpbHRlcnMoKSB7CiAgcmV0dXJuIHsKICAgIHByb2plY3Q6IG1zR2V0U2VsZWN0ZWQoJ21zLXByb2plY3QnKSwKICAgIGNvb3JkOiAgIG1zR2V0U2VsZWN0ZWQoJ21zLWNvb3JkJyksCiAgICBidWNrZXQ6ICBtc0dldFNlbGVjdGVkKCdtcy1idWNrZXQnKSwKICB9Owp9CmZ1bmN0aW9uIHJlc2V0RmlsdGVycygpIHsKICBbJ2ZpbHRlclByb2plY3QnLCdmaWx0ZXJDb29yZCcsJ2ZpbHRlckJ1Y2tldCddLmZvckVhY2goaWQgPT4gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQpLnZhbHVlPScnKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb3JkZXJTZWFyY2gnKS52YWx1ZSA9ICcnOwogIGFwcGx5RmlsdGVycygpOwp9CgpmdW5jdGlvbiBhcHBseUZpbHRlcnMoKSB7CiAgaWYgKCFSQVcpIHJldHVybjsKICBjb25zdCB7cHJvamVjdCwgY29vcmQsIGJ1Y2tldH0gPSBnZXRGaWx0ZXJzKCk7CiAgZmlsdGVyZWRPcmRlcnMgPSBSQVcub3JkZXJzX3RhYmxlLmZpbHRlcihvID0+CiAgICAoIXByb2plY3QubGVuZ3RoIHx8IHByb2plY3QuaW5jbHVkZXMoby5wcm9qZWN0KSkgJiYKICAgICghY29vcmQubGVuZ3RoICAgfHwgY29vcmQuaW5jbHVkZXMoby5jb29yZGluYXRvcikpICYmCiAgICAoIWJ1Y2tldC5sZW5ndGggIHx8IGJ1Y2tldC5pbmNsdWRlcyhvLmFnZV9idWNrZXQpKQogICk7CiAgcmVuZGVyS1BJcygpOyByZW5kZXJDb29yZFRhYmxlKCk7IHJlbmRlckFnZVRhYmxlKCk7IHJlbmRlck9yZGVycygpOwp9CgovLyDilIDilIAgS1BJcyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyS1BJcygpIHsKICBjb25zdCB7cHJvamVjdCwgY29vcmQsIGJ1Y2tldH0gPSBnZXRGaWx0ZXJzKCk7CiAgY29uc3QgaGFzRmlsdGVyID0gcHJvamVjdC5sZW5ndGggfHwgY29vcmQubGVuZ3RoIHx8IGJ1Y2tldC5sZW5ndGg7CiAgaWYgKCFoYXNGaWx0ZXIpIHsKICAgIGNvbnN0IGsgPSBSQVcua3BpczsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrVG90YWwnKS50ZXh0Q29udGVudCAgICAgID0gZm10JChrLnRvdGFsX3ZhbHVlKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrQ29udGFpbmVycycpLnRleHRDb250ZW50ID0gZm10TihrLnRvdGFsX2NvbnRhaW5lcnMpICsgJyBjb250YWluZXJzJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrT3ZlcjkwJykudGV4dENvbnRlbnQgICAgID0gZm10JChrLm92ZXI5MF92YWx1ZSk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna092ZXI5MFBjdCcpLnRleHRDb250ZW50ICA9IGZtdFBjdChrLm92ZXI5MF9wY3QpICsgJyBvZiB0b3RhbCc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna1N0b3JhZ2UnKS50ZXh0Q29udGVudCAgICA9IGZtdCQoay5pbl9zdG9yYWdlX3ZhbHVlKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrRmluYW5jZScpLnRleHRDb250ZW50ICAgID0gZm10JChrLmZpbmFuY2VfaG9sZF92YWx1ZSk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna0F2Z0FnZScpLnRleHRDb250ZW50ICAgICA9IGZtdEYoay5hdmdfYWdlKTsKICB9IGVsc2UgewogICAgbGV0IHR2PTAsIHRjPTAsIG85MD0wLCB0YT0wOwogICAgZmlsdGVyZWRPcmRlcnMuZm9yRWFjaChvID0+IHsgdHYrPW8udmFsdWU7IHRjKz1vLmNvbnRhaW5lcnM7IGlmKG8ubWF4X2FnZT45MCkgbzkwKz1vLnZhbHVlOyB0YSs9by5hdmdfYWdlKm8uY29udGFpbmVyczsgfSk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna1RvdGFsJykudGV4dENvbnRlbnQgICAgICA9IGZtdCQodHYpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tDb250YWluZXJzJykudGV4dENvbnRlbnQgPSBmbXROKHRjKSArICcgY29udGFpbmVycyc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna092ZXI5MCcpLnRleHRDb250ZW50ICAgICA9IGZtdCQobzkwKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrT3ZlcjkwUGN0JykudGV4dENvbnRlbnQgID0gdHYgPyBmbXRQY3QobzkwL3R2KSsnIG9mIHRvdGFsJyA6ICfigJQnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tTdG9yYWdlJykudGV4dENvbnRlbnQgICAgPSAn4oCUJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrRmluYW5jZScpLnRleHRDb250ZW50ICAgID0gJ+KAlCc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna0F2Z0FnZScpLnRleHRDb250ZW50ICAgICA9IHRjID8gKHRhL3RjKS50b0ZpeGVkKDEpIDogJ+KAlCc7CiAgfQp9CgovLyDilIDilIAgQ29vcmRpbmF0b3IgdGFibGUg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHJlbmRlckNvb3JkVGFibGUoKSB7CiAgY29uc3Qge3Byb2plY3QsIGNvb3JkLCBidWNrZXR9ID0gZ2V0RmlsdGVycygpOwogIGxldCByb3dzOwogIGlmICghcHJvamVjdCAmJiAhY29vcmQgJiYgIWJ1Y2tldCkgewogICAgcm93cyA9IFJBVy5jb29yZF90YWJsZTsKICB9IGVsc2UgewogICAgY29uc3QgbWFwID0ge307CiAgICBmaWx0ZXJlZE9yZGVycy5mb3JFYWNoKG8gPT4gewogICAgICBpZiAoIW1hcFtvLmNvb3JkaW5hdG9yXSkgbWFwW28uY29vcmRpbmF0b3JdPXtjb29yZGluYXRvcjpvLmNvb3JkaW5hdG9yLGNvbnRhaW5lcnM6MCx2YWx1ZTowLG92ZXI5MF92YWx1ZTowLGFnZXM6W119OwogICAgICBtYXBbby5jb29yZGluYXRvcl0uY29udGFpbmVycys9by5jb250YWluZXJzOyBtYXBbby5jb29yZGluYXRvcl0udmFsdWUrPW8udmFsdWU7CiAgICAgIGlmKG8ubWF4X2FnZT45MCkgbWFwW28uY29vcmRpbmF0b3JdLm92ZXI5MF92YWx1ZSs9by52YWx1ZTsKICAgICAgbWFwW28uY29vcmRpbmF0b3JdLmFnZXMucHVzaCguLi5BcnJheShvLmNvbnRhaW5lcnMpLmZpbGwoby5hdmdfYWdlKSk7CiAgICB9KTsKICAgIHJvd3MgPSBPYmplY3QudmFsdWVzKG1hcCkubWFwKHI9Pih7Li4uciwKICAgICAgb3ZlcjkwX3BjdDogci52YWx1ZSA/IHIub3ZlcjkwX3ZhbHVlL3IudmFsdWUgOiAwLAogICAgICBhdmdfYWdlOiByLmFnZXMubGVuZ3RoID8gci5hZ2VzLnJlZHVjZSgoYSxiKT0+YStiLDApL3IuYWdlcy5sZW5ndGggOiAwLAogICAgfSkpLnNvcnQoKGEsYik9PmEuY29vcmRpbmF0b3IubG9jYWxlQ29tcGFyZShiLmNvb3JkaW5hdG9yKSk7CiAgfQogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb29yZE1ldGEnKS50ZXh0Q29udGVudCA9IHJvd3MubGVuZ3RoICsgJyBjb29yZGluYXRvcnMnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjb29yZEJvZHknKS5pbm5lckhUTUwgPSByb3dzLm1hcChyPT5gCiAgICA8dHI+CiAgICAgIDx0ZD4ke3IuY29vcmRpbmF0b3J9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10TihyLmNvbnRhaW5lcnMpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQoci52YWx1ZSl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10JChyLm92ZXI5MF92YWx1ZSl9PC90ZD4KICAgICAgPHRkPjxkaXYgY2xhc3M9ImJhci1jZWxsIj48ZGl2IGNsYXNzPSJiYXIiPjxkaXYgY2xhc3M9ImJhci1maWxsIHJlZCIgc3R5bGU9IndpZHRoOiR7TWF0aC5taW4oci5vdmVyOTBfcGN0KjEwMCwxMDApLnRvRml4ZWQoMSl9JSI+PC9kaXY+PC9kaXY+PHNwYW4+JHtmbXRQY3Qoci5vdmVyOTBfcGN0KX08L3NwYW4+PC9kaXY+PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10RihyLmF2Z19hZ2UpfTwvdGQ+CiAgICA8L3RyPmApLmpvaW4oJycpOwogIGNvbnN0IHRvdCA9IHJvd3MucmVkdWNlKChhLHIpPT4oe2NvbnRhaW5lcnM6YS5jb250YWluZXJzK3IuY29udGFpbmVycyx2YWx1ZTphLnZhbHVlK3IudmFsdWUsb3ZlcjkwX3ZhbHVlOmEub3ZlcjkwX3ZhbHVlK3Iub3ZlcjkwX3ZhbHVlfSkse2NvbnRhaW5lcnM6MCx2YWx1ZTowLG92ZXI5MF92YWx1ZTowfSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nvb3JkRm9vdCcpLmlubmVySFRNTCA9IGA8dHI+PHRkPlRPVEFMPC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXROKHRvdC5jb250YWluZXJzKX08L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQodG90LnZhbHVlKX08L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQodG90Lm92ZXI5MF92YWx1ZSl9PC90ZD48dGQgY2xhc3M9Im51bSI+JHt0b3QudmFsdWU/Zm10UGN0KHRvdC5vdmVyOTBfdmFsdWUvdG90LnZhbHVlKTon4oCUJ308L3RkPjx0ZD48L3RkPjwvdHI+YDsKfQoKLy8g4pSA4pSAIEFnZSBidWNrZXQgdGFibGUg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHJlbmRlckFnZVRhYmxlKCkgewogIGNvbnN0IHtwcm9qZWN0LCBjb29yZCwgYnVja2V0fSA9IGdldEZpbHRlcnMoKTsKICBsZXQgcm93czsKICBpZiAoIXByb2plY3QgJiYgIWNvb3JkICYmICFidWNrZXQpIHsKICAgIHJvd3MgPSBSQVcuYWdlX3RhYmxlOwogIH0gZWxzZSB7CiAgICBjb25zdCBtYXAgPSB7fTsgQUdFX09SREVSLmZvckVhY2goaz0+bWFwW2tdPXtidWNrZXQ6ayxsYWJlbDprLGNvbnRhaW5lcnM6MCx2YWx1ZTowfSk7CiAgICBmaWx0ZXJlZE9yZGVycy5mb3JFYWNoKG89PntpZihtYXBbby5hZ2VfYnVja2V0XSl7bWFwW28uYWdlX2J1Y2tldF0uY29udGFpbmVycys9by5jb250YWluZXJzO21hcFtvLmFnZV9idWNrZXRdLnZhbHVlKz1vLnZhbHVlO319KTsKICAgIGNvbnN0IHR2PU9iamVjdC52YWx1ZXMobWFwKS5yZWR1Y2UoKGEscik9PmErci52YWx1ZSwwKTsKICAgIHJvd3M9QUdFX09SREVSLm1hcChrPT4oey4uLm1hcFtrXSxwY3Q6dHY/bWFwW2tdLnZhbHVlL3R2OjB9KSk7CiAgfQogIGNvbnN0IHR2PXJvd3MucmVkdWNlKChhLHIpPT5hK3IudmFsdWUsMCk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FnZUJvZHknKS5pbm5lckhUTUw9cm93cy5tYXAocj0+YAogICAgPHRyPgogICAgICA8dGQ+JHtyLmxhYmVsfHxBR0VfTEFCRUxTW3IuYnVja2V0XX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXROKHIuY29udGFpbmVycyl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10JChyLnZhbHVlKX08L3RkPgogICAgICA8dGQ+PGRpdiBjbGFzcz0iYmFyLWNlbGwiPjxkaXYgY2xhc3M9ImJhciI+PGRpdiBjbGFzcz0iYmFyLWZpbGwiIHN0eWxlPSJ3aWR0aDoke01hdGgubWluKChyLnBjdHx8MCkqMTAwLDEwMCkudG9GaXhlZCgxKX0lIj48L2Rpdj48L2Rpdj48c3Bhbj4ke2ZtdFBjdChyLnBjdCl9PC9zcGFuPjwvZGl2PjwvdGQ+CiAgICA8L3RyPmApLmpvaW4oJycpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhZ2VGb290JykuaW5uZXJIVE1MPWA8dHI+PHRkPlRPVEFMPC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXROKHJvd3MucmVkdWNlKChhLHIpPT5hK3IuY29udGFpbmVycywwKSl9PC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXQkKHR2KX08L3RkPjx0ZD48L3RkPjwvdHI+YDsKfQoKLy8g4pSA4pSAIE9yZGVycyB0YWJsZSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyT3JkZXJzKCkgewogIGNvbnN0IHEgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb3JkZXJTZWFyY2gnKS52YWx1ZS50b0xvd2VyQ2FzZSgpOwogIGxldCByb3dzID0gZmlsdGVyZWRPcmRlcnM7CiAgaWYgKHEpIHJvd3M9cm93cy5maWx0ZXIobz0+by5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpfHxvLnByb2plY3QudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKXx8by5jb29yZGluYXRvci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb3JkZXJzTWV0YScpLnRleHRDb250ZW50ID0gcm93cy5sZW5ndGgrJyBvcmRlcnMnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvcmRlcnNCb2R5JykuaW5uZXJIVE1MID0gcm93cy5tYXAobz0+YAogICAgPHRyPgogICAgICA8dGQ+PHN0cm9uZz4ke28ub3JkZXJ9PC9zdHJvbmc+PC90ZD4KICAgICAgPHRkIGNsYXNzPSJ3cmFwIj4ke28ucHJvamVjdH08L3RkPgogICAgICA8dGQ+JHtvLmNvb3JkaW5hdG9yfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4oby5jb250YWluZXJzKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXQkKG8udmFsdWUpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdEYoby5hdmdfYWdlKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHttYXhBZ2VUYWcoby5tYXhfYWdlKX08L3RkPgogICAgICA8dGQ+JHthZ2VCdWNrZXRUYWcoby5hZ2VfYnVja2V0KX08L3RkPgogICAgICA8dGQ+JHtvLnNoaXBfcmFuZ2V8fCfigJQnfTwvdGQ+CiAgICA8L3RyPmApLmpvaW4oJycpIHx8ICc8dHI+PHRkIGNvbHNwYW49IjkiIHN0eWxlPSJ0ZXh0LWFsaWduOmNlbnRlcjtwYWRkaW5nOjMycHg7Y29sb3I6dmFyKC0tbXV0ZWQpIj5ObyBvcmRlcnMgbWF0Y2ggZmlsdGVycy48L3RkPjwvdHI+JzsKICBpZiAoIXJvd3MubGVuZ3RoKSB7IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvcmRlcnNGb290JykuaW5uZXJIVE1MPScnOyByZXR1cm47IH0KICBjb25zdCB0b3Q9cm93cy5yZWR1Y2UoKGEsbyk9Pih7Y29udGFpbmVyczphLmNvbnRhaW5lcnMrby5jb250YWluZXJzLHZhbHVlOmEudmFsdWUrby52YWx1ZX0pLHtjb250YWluZXJzOjAsdmFsdWU6MH0pOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvcmRlcnNGb290JykuaW5uZXJIVE1MPWA8dHI+PHRkIGNvbHNwYW49IjMiPlRPVEFMIOKAlCAke3Jvd3MubGVuZ3RofSBvcmRlcnM8L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4odG90LmNvbnRhaW5lcnMpfTwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10JCh0b3QudmFsdWUpfTwvdGQ+PHRkIGNvbHNwYW49IjQiPjwvdGQ+PC90cj5gOwp9CgovLyDilIDilIAgUy1Ecm9wIEtQSXMg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHJlbmRlclNkcm9wS3BpcygpIHsKICBjb25zdCBrID0gUkFXLnNkcm9wPy5rcGlzOwogIGlmICghaykgcmV0dXJuOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZEtwaUl0ZW1zJykudGV4dENvbnRlbnQgID0gZm10TihrLnRvdGFsX2l0ZW1zKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2RLcGlWYWx1ZScpLnRleHRDb250ZW50ICA9IGZtdCQoay50b3RhbF92YWx1ZSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NkS3BpT3JkZXJzJykudGV4dENvbnRlbnQgPSBmbXROKGsudW5pcXVlX29yZGVycyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NkS3BpQXZnJykudGV4dENvbnRlbnQgICAgPSBmbXRGKGsuYXZnX2FnZSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NkS3BpTWF4JykudGV4dENvbnRlbnQgICAgPSBmbXROKGsubWF4X2FnZSk7Cn0KCi8vIOKUgOKUgCBTLURyb3AgbG9jYXRpb24gY2FyZHMg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHJlbmRlclNkcm9wTG9jR3JpZCgpIHsKICBjb25zdCBsb2NzID0gUkFXLnNkcm9wPy5ieV9sb2NhdGlvbiB8fCBbXTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BMb2NHcmlkJykuaW5uZXJIVE1MID0gbG9jcy5tYXAobD0+YAogICAgPGRpdiBjbGFzcz0ic2Ryb3AtbG9jLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtbmFtZSI+JHtsLmxvY2F0aW9ufTwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtaXRlbXMiPiR7Zm10TihsLml0ZW1zKX0gPHNwYW4gc3R5bGU9ImZvbnQtc2l6ZToxM3B4O2ZvbnQtd2VpZ2h0OjQwMDtjb2xvcjp2YXIoLS1tdXRlZCkiPml0ZW1zPC9zcGFuPjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtdmFsIj4ke2ZtdCQobC52YWx1ZSl9IMK3IGF2ZyAke2ZtdEYobC5hdmdfYWdlKX1kIMK3IG1heCAke2ZtdE4obC5tYXhfYWdlKX1kPC9kaXY+CiAgICA8L2Rpdj5gKS5qb2luKCcnKTsKfQoKLy8g4pSA4pSAIFMtRHJvcCBkZXRhaWwgdGFibGUg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHJlc2V0U2Ryb3BGaWx0ZXJzKCkgewogIG1zUmVzZXQoJ21zLWRyb3Bsb2MnKTsgbXNSZXNldCgnbXMtZHJvcGNvb3JkJyk7IG1zUmVzZXQoJ21zLWRyb3BmbGFnJyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Nkcm9wU2VhcmNoJykudmFsdWUgPSAnJzsKICByZW5kZXJTZHJvcCgpOwp9CgpmdW5jdGlvbiByZW5kZXJTZHJvcCgpIHsKICBjb25zdCBpdGVtcyA9IFJBVy5zZHJvcD8uaXRlbXMgfHwgW107CiAgY29uc3QgbG9jICAgPSBtc0dldFNlbGVjdGVkKCdtcy1kcm9wbG9jJyk7CiAgY29uc3QgY29vcmQgPSBtc0dldFNlbGVjdGVkKCdtcy1kcm9wY29vcmQnKTsKICBjb25zdCBmbGFnICA9IG1zR2V0U2VsZWN0ZWQoJ21zLWRyb3BmbGFnJyk7CiAgY29uc3QgcSAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BTZWFyY2gnKS52YWx1ZS50b0xvd2VyQ2FzZSgpOwoKICBsZXQgcm93cyA9IGl0ZW1zOwogIGlmIChsb2MubGVuZ3RoKSAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGxvYy5pbmNsdWRlcyhyLmxvY2F0aW9uKSk7CiAgaWYgKGNvb3JkLmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gY29vcmQuaW5jbHVkZXMoci5jb29yZGluYXRvcikpOwogIGlmIChmbGFnLmxlbmd0aCkgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGZsYWcuaW5jbHVkZXMoc2Ryb3BBZ2VGbGFnKHIuYWdlKSkpOwogIGlmIChxKSAgICByb3dzID0gcm93cy5maWx0ZXIociA9PgogICAgci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICByLnByb2plY3QudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fAogICAgci5wYXJ0X25vLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwKICAgIHIuY29vcmRpbmF0b3IudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKQogICk7CgogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZHJvcE1ldGEnKS50ZXh0Q29udGVudCA9IHJvd3MubGVuZ3RoICsgJyBpdGVtcyc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Nkcm9wQm9keScpLmlubmVySFRNTCA9IHJvd3MubWFwKHI9PmAKICAgIDx0cj4KICAgICAgPHRkPiR7ci5sb2NhdGlvbn08L3RkPgogICAgICA8dGQ+PHN0cm9uZz4ke3Iub3JkZXJ9PC9zdHJvbmc+PC90ZD4KICAgICAgPHRkIGNsYXNzPSJ3cmFwIj4ke3IucHJvamVjdH08L3RkPgogICAgICA8dGQ+JHtyLmNvb3JkaW5hdG9yfTwvdGQ+CiAgICAgIDx0ZD4ke3IucGFydF9ub3x8J+KAlCd9PC90ZD4KICAgICAgPHRkPiR7ci5zZXJpYWx8fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3IucGFydF9ncm91cH08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXROKHIucXR5KX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHttYXhBZ2VUYWcoci5hZ2UpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQoci52YWx1ZSl9PC90ZD4KICAgICAgPHRkPiR7ci5zaGlwX2RhdGV8fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3Iuc3RhdHVzfTwvdGQ+CiAgICAgIDx0ZD4ke3Nkcm9wRmxhZ1RhZyhyLmFnZSl9PC90ZD4KICAgIDwvdHI+YCkuam9pbignJykgfHwgJzx0cj48dGQgY29sc3Bhbj0iMTEiIHN0eWxlPSJ0ZXh0LWFsaWduOmNlbnRlcjtwYWRkaW5nOjMycHg7Y29sb3I6dmFyKC0tbXV0ZWQpIj5ObyBpdGVtcyBtYXRjaCBmaWx0ZXJzLjwvdGQ+PC90cj4nOwoKICBpZiAoIXJvd3MubGVuZ3RoKSB7IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZHJvcEZvb3QnKS5pbm5lckhUTUw9Jyc7IHJldHVybjsgfQogIGNvbnN0IHRvdFF0eT1yb3dzLnJlZHVjZSgoYSxyKT0+YStyLnF0eSwwKSwgdG90VmFsPXJvd3MucmVkdWNlKChhLHIpPT5hK3IudmFsdWUsMCk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Nkcm9wRm9vdCcpLmlubmVySFRNTD1gPHRyPjx0ZCBjb2xzcGFuPSI2Ij5UT1RBTCDigJQgJHtyb3dzLmxlbmd0aH0gaXRlbXM8L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4odG90UXR5KX08L3RkPjx0ZD48L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQodG90VmFsKX08L3RkPjx0ZCBjb2xzcGFuPSIyIj48L3RkPjwvdHI+YDsKfQoKLy8g4pSA4pSAIFNvcnRpbmcg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHNvcnRUYWJsZSh0YWJsZUlkLCBjb2xJZHgpIHsKICBjb25zdCBrZXkgPSB0YWJsZUlkKydfJytjb2xJZHg7CiAgY29uc3QgYXNjID0gc29ydFN0YXRlW2tleV0gIT09ICdhc2MnOwogIHNvcnRTdGF0ZVtrZXldID0gYXNjID8gJ2FzYycgOiAnZGVzYyc7CiAgY29uc3QgdGFibGUgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCh0YWJsZUlkKTsKICB0YWJsZS5xdWVyeVNlbGVjdG9yQWxsKCd0aGVhZCB0aCcpLmZvckVhY2goKHRoLGkpPT57CiAgICB0aC5jbGFzc0xpc3QucmVtb3ZlKCdzb3J0LWFzYycsJ3NvcnQtZGVzYycpOwogICAgaWYoaT09PWNvbElkeCkgdGguY2xhc3NMaXN0LmFkZChhc2M/J3NvcnQtYXNjJzonc29ydC1kZXNjJyk7CiAgfSk7CiAgY29uc3QgdGJvZHkgPSB0YWJsZS5xdWVyeVNlbGVjdG9yKCd0Ym9keScpOwogIEFycmF5LmZyb20odGJvZHkucXVlcnlTZWxlY3RvckFsbCgndHInKSkuc29ydCgoYSxiKT0+ewogICAgbGV0IGF2PShhLmNlbGxzW2NvbElkeF0/LnRleHRDb250ZW50fHwnJykucmVwbGFjZSgvWyQsJVxcc+KGkeKGk10vZywnJykudHJpbSgpOwogICAgbGV0IGJ2PShiLmNlbGxzW2NvbElkeF0/LnRleHRDb250ZW50fHwnJykucmVwbGFjZSgvWyQsJVxcc+KGkeKGk10vZywnJykudHJpbSgpOwogICAgY29uc3QgYW49cGFyc2VGbG9hdChhdiksYm49cGFyc2VGbG9hdChidik7CiAgICBjb25zdCBjbXA9KCFpc05hTihhbikmJiFpc05hTihibikpP2FuLWJuOmF2LmxvY2FsZUNvbXBhcmUoYnYpOwogICAgcmV0dXJuIGFzYz9jbXA6LWNtcDsKICB9KS5mb3JFYWNoKHI9PnRib2R5LmFwcGVuZENoaWxkKHIpKTsKfQoKLy8g4pSA4pSAIE11bHRpLXNlbGVjdCBkcm9wZG93biBlbmdpbmUg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIG1zVG9nZ2xlKGlkKSB7CiAgY29uc3Qgd3JhcCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsKICBjb25zdCBwYW5lbD0gd3JhcC5xdWVyeVNlbGVjdG9yKCcubXMtcGFuZWwnKTsKICBjb25zdCBpc09wZW4gPSBwYW5lbC5jbGFzc0xpc3QuY29udGFpbnMoJ29wZW4nKTsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcubXMtcGFuZWwub3BlbicpLmZvckVhY2gocCA9PiB7CiAgICBwLmNsYXNzTGlzdC5yZW1vdmUoJ29wZW4nKTsKICAgIHAuY2xvc2VzdCgnLm1zLXdyYXAnKS5xdWVyeVNlbGVjdG9yKCcubXMtYnRuJykuY2xhc3NMaXN0LnJlbW92ZSgnb3BlbicpOwogIH0pOwogIGlmICghaXNPcGVuKSB7IHBhbmVsLmNsYXNzTGlzdC5hZGQoJ29wZW4nKTsgd3JhcC5xdWVyeVNlbGVjdG9yKCcubXMtYnRuJykuY2xhc3NMaXN0LmFkZCgnb3BlbicpOyB9Cn0KLy8gVXNlIGV2ZW50IGRlbGVnYXRpb24g4oCUIGNhdGNoZXMgZHluYW1pY2FsbHkgYWRkZWQgYnV0dG9ucyB0b28KZG9jdW1lbnQuYWRkRXZlbnRMaXN0ZW5lcignY2xpY2snLCBlID0+IHsKICBjb25zdCBidG4gPSBlLnRhcmdldC5jbG9zZXN0KCcubXMtYnRuJyk7CiAgaWYgKGJ0bikgewogICAgZS5zdG9wUHJvcGFnYXRpb24oKTsKICAgIGNvbnN0IHdyYXAgPSBidG4uY2xvc2VzdCgnLm1zLXdyYXAnKTsKICAgIGlmICh3cmFwKSBtc1RvZ2dsZSh3cmFwLmlkKTsKICAgIHJldHVybjsKICB9CiAgLy8gQ2xpY2sgb3V0c2lkZSDigJQgY2xvc2UgYWxsCiAgaWYgKCFlLnRhcmdldC5jbG9zZXN0KCcubXMtd3JhcCcpKSB7CiAgICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcubXMtcGFuZWwub3BlbicpLmZvckVhY2gocCA9PiB7CiAgICAgIHAuY2xhc3NMaXN0LnJlbW92ZSgnb3BlbicpOwogICAgICBwLmNsb3Nlc3QoJy5tcy13cmFwJykucXVlcnlTZWxlY3RvcignLm1zLWJ0bicpLmNsYXNzTGlzdC5yZW1vdmUoJ29wZW4nKTsKICAgIH0pOwogIH0KfSk7CgpmdW5jdGlvbiBtc1BvcHVsYXRlKGlkLCBvcHRpb25zLCBvbkNoYW5nZSkgewogIGNvbnN0IHBhbmVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQgKyAnLXBhbmVsJyk7CiAgcGFuZWwuaW5uZXJIVE1MID0gJyc7CgogIC8vIEhlbHBlciB0byBidWlsZCBhIGNoZWNrYm94IHJvdyB1c2luZyBET00gbWV0aG9kcyAoYXZvaWRzIGlubmVySFRNTCBsYWJlbCBpc3N1ZXMpCiAgZnVuY3Rpb24gbWFrZVJvdyh2YWx1ZSwgbGFiZWxUZXh0LCBpc0JvbGQpIHsKICAgIGNvbnN0IHJvdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOwogICAgcm93LmNsYXNzTmFtZSA9ICdtcy1pdGVtJzsKICAgIGNvbnN0IGNiID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnaW5wdXQnKTsKICAgIGNiLnR5cGUgPSAnY2hlY2tib3gnOyBjYi5jaGVja2VkID0gdHJ1ZTsKICAgIGlmICh2YWx1ZSAhPT0gbnVsbCkgY2IudmFsdWUgPSB2YWx1ZTsKICAgIGVsc2UgY2IuaWQgPSBpZCArICctYWxsJzsKICAgIGNvbnN0IGxibCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ3NwYW4nKTsKICAgIGxibC50ZXh0Q29udGVudCA9IGxhYmVsVGV4dDsKICAgIGxibC5zdHlsZS5jc3NUZXh0ID0gJ2N1cnNvcjpwb2ludGVyO2NvbG9yOiMxRjM4NjQ7Zm9udC1zaXplOjEzcHg7dXNlci1zZWxlY3Q6bm9uZTsnICsgKGlzQm9sZCA/ICdmb250LXdlaWdodDo2MDA7JyA6ICcnKTsKICAgIHJvdy5hcHBlbmRDaGlsZChjYik7IHJvdy5hcHBlbmRDaGlsZChsYmwpOwogICAgcmV0dXJuIHsgcm93LCBjYiB9OwogIH0KCiAgLy8gIkFsbCIgcm93CiAgY29uc3QgeyByb3c6IGFsbFJvdywgY2I6IGFsbENiIH0gPSBtYWtlUm93KG51bGwsICdBbGwnLCB0cnVlKTsKICBwYW5lbC5hcHBlbmRDaGlsZChhbGxSb3cpOwogIGNvbnN0IGRpdiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2RpdicpOyBkaXYuY2xhc3NOYW1lID0gJ21zLWRpdmlkZXInOyBwYW5lbC5hcHBlbmRDaGlsZChkaXYpOwoKICAvLyBPcHRpb24gcm93cwogIG9wdGlvbnMuZm9yRWFjaChvcHQgPT4gewogICAgY29uc3QgeyByb3csIGNiIH0gPSBtYWtlUm93KG9wdCwgb3B0LCBmYWxzZSk7CiAgICBwYW5lbC5hcHBlbmRDaGlsZChyb3cpOwogIH0pOwoKICAvLyBFdmVudCBsaXN0ZW5lcnMKICBwYW5lbC5xdWVyeVNlbGVjdG9yQWxsKCdpbnB1dFt2YWx1ZV0nKS5mb3JFYWNoKGNiID0+IHsKICAgIGNiLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgICAgYWxsQ2IuY2hlY2tlZCA9IFsuLi5wYW5lbC5xdWVyeVNlbGVjdG9yQWxsKCdpbnB1dFt2YWx1ZV0nKV0uZXZlcnkoYyA9PiBjLmNoZWNrZWQpOwogICAgICBtc1VwZGF0ZUxhYmVsKGlkLCBvcHRpb25zLmxlbmd0aCk7CiAgICAgIG9uQ2hhbmdlKCk7CiAgICB9KTsKICB9KTsKICBhbGxDYi5hZGRFdmVudExpc3RlbmVyKCdjaGFuZ2UnLCAoKSA9PiB7CiAgICBwYW5lbC5xdWVyeVNlbGVjdG9yQWxsKCdpbnB1dFt2YWx1ZV0nKS5mb3JFYWNoKGNiID0+IGNiLmNoZWNrZWQgPSBhbGxDYi5jaGVja2VkKTsKICAgIG1zVXBkYXRlTGFiZWwoaWQsIG9wdGlvbnMubGVuZ3RoKTsKICAgIG9uQ2hhbmdlKCk7CiAgfSk7Cn0KCmZ1bmN0aW9uIG1zR2V0U2VsZWN0ZWQoaWQpIHsKICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkICsgJy1wYW5lbCcpOwogIGlmICghcGFuZWwpIHJldHVybiBbXTsKICBjb25zdCBjaGVja2VkID0gWy4uLnBhbmVsLnF1ZXJ5U2VsZWN0b3JBbGwoJ2lucHV0W3ZhbHVlXTpjaGVja2VkJyldLm1hcChjID0+IGMudmFsdWUpOwogIGNvbnN0IGFsbCAgICAgPSBbLi4ucGFuZWwucXVlcnlTZWxlY3RvckFsbCgnaW5wdXRbdmFsdWVdJyldLm1hcChjID0+IGMudmFsdWUpOwogIC8vIElmIGFsbCBjaGVja2VkLCByZXR1cm4gZW1wdHkgYXJyYXkgKG1lYW5zICJubyBmaWx0ZXIiKQogIHJldHVybiBjaGVja2VkLmxlbmd0aCA9PT0gYWxsLmxlbmd0aCA/IFtdIDogY2hlY2tlZDsKfQoKZnVuY3Rpb24gbXNVcGRhdGVMYWJlbChpZCwgdG90YWwpIHsKICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkICsgJy1wYW5lbCcpOwogIGNvbnN0IGJ0biAgID0gZG9jdW1lbnQucXVlcnlTZWxlY3RvcignIycgKyBpZCArICcgLm1zLWxhYmVsJyk7CiAgaWYgKCFwYW5lbCB8fCAhYnRuKSByZXR1cm47CiAgY29uc3QgY2hlY2tlZCA9IFsuLi5wYW5lbC5xdWVyeVNlbGVjdG9yQWxsKCdpbnB1dFt2YWx1ZV06Y2hlY2tlZCcpXTsKICBjb25zdCBhbGwgICAgID0gWy4uLnBhbmVsLnF1ZXJ5U2VsZWN0b3JBbGwoJ2lucHV0W3ZhbHVlXScpXTsKICBpZiAoY2hlY2tlZC5sZW5ndGggPT09IDApICAgICAgICAgYnRuLnRleHRDb250ZW50ID0gJ05vbmUgc2VsZWN0ZWQnOwogIGVsc2UgaWYgKGNoZWNrZWQubGVuZ3RoID09PSBhbGwubGVuZ3RoKSBidG4udGV4dENvbnRlbnQgPSAnQWxsJzsKICBlbHNlIGlmIChjaGVja2VkLmxlbmd0aCA9PT0gMSkgICAgYnRuLnRleHRDb250ZW50ID0gY2hlY2tlZFswXS52YWx1ZTsKICBlbHNlICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYnRuLnRleHRDb250ZW50ID0gY2hlY2tlZC5sZW5ndGggKyAnIHNlbGVjdGVkJzsKfQoKZnVuY3Rpb24gbXNSZXNldChpZCkgewogIGNvbnN0IHBhbmVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQgKyAnLXBhbmVsJyk7CiAgaWYgKCFwYW5lbCkgcmV0dXJuOwogIHBhbmVsLnF1ZXJ5U2VsZWN0b3JBbGwoJ2lucHV0JykuZm9yRWFjaChjYiA9PiBjYi5jaGVja2VkID0gdHJ1ZSk7CiAgbXNVcGRhdGVMYWJlbChpZCwgMCk7Cn0KCi8vIOKUgOKUgCBDU1YgRXhwb3J0IOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiB0b0NTVihoZWFkZXJzLCByb3dzKSB7CiAgY29uc3QgZXNjYXBlID0gdiA9PiB7CiAgICBpZiAodiA9PSBudWxsKSByZXR1cm4gJyc7CiAgICBjb25zdCBzID0gU3RyaW5nKHYpOwogICAgcmV0dXJuIHMuaW5jbHVkZXMoJywnKSB8fCBzLmluY2x1ZGVzKCciJykgfHwgcy5pbmNsdWRlcygnXG4nKSA/ICciJyArIHMucmVwbGFjZSgvIi9nLCAnIiInKSArICciJyA6IHM7CiAgfTsKICBjb25zdCBsaW5lcyA9IFtoZWFkZXJzLm1hcChlc2NhcGUpLmpvaW4oJywnKV07CiAgcm93cy5mb3JFYWNoKHIgPT4gbGluZXMucHVzaChyLm1hcChlc2NhcGUpLmpvaW4oJywnKSkpOwogIHJldHVybiBsaW5lcy5qb2luKCdcbicpOwp9CgpmdW5jdGlvbiBkb3dubG9hZENTVihmaWxlbmFtZSwgY3N2KSB7CiAgY29uc3QgYmxvYiA9IG5ldyBCbG9iKFtjc3ZdLCB7dHlwZTogJ3RleHQvY3N2J30pOwogIGNvbnN0IHVybCAgPSBVUkwuY3JlYXRlT2JqZWN0VVJMKGJsb2IpOwogIGNvbnN0IGEgICAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdhJyk7CiAgYS5ocmVmID0gdXJsOyBhLmRvd25sb2FkID0gZmlsZW5hbWU7IGEuY2xpY2soKTsKICBVUkwucmV2b2tlT2JqZWN0VVJMKHVybCk7Cn0KCmZ1bmN0aW9uIGV4cG9ydE1haW5DU1YoKSB7CiAgY29uc3QgaGVhZGVycyA9IFsnT3JkZXIgIycsJ1Byb2plY3QnLCdDb29yZGluYXRvcicsJ0NvbnRhaW5lcnMnLCdWYWx1ZScsJ0F2ZyBBZ2UnLCdNYXggQWdlJywnQWdlIEJ1Y2tldCcsJ1N0YXR1cycsJ1NoaXAgRGF0ZSBSYW5nZSddOwogIGNvbnN0IHJvd3MgPSBmaWx0ZXJlZE9yZGVycy5tYXAobyA9PiBbCiAgICBvLm9yZGVyLCBvLnByb2plY3QsIG8uY29vcmRpbmF0b3IsIG8uY29udGFpbmVycywKICAgIG8udmFsdWUsIG8uYXZnX2FnZSwgby5tYXhfYWdlLCBvLmFnZV9idWNrZXQsIG8uc3RhdHVzLCBvLnNoaXBfcmFuZ2UKICBdKTsKICBjb25zdCBwcm9qICA9IG1zR2V0U2VsZWN0ZWQoJ21zLXByb2plY3QnKTsKICBjb25zdCBjb29yZCA9IG1zR2V0U2VsZWN0ZWQoJ21zLWNvb3JkJyk7CiAgY29uc3QgYnVja2V0PSBtc0dldFNlbGVjdGVkKCdtcy1idWNrZXQnKTsKICBjb25zdCBsYWJlbCA9IFsKICAgIHByb2oubGVuZ3RoICAgPyBwcm9qLmpvaW4oJysnKSAgIDogJycsCiAgICBjb29yZC5sZW5ndGggID8gY29vcmQuam9pbignKycpICA6ICcnLAogICAgYnVja2V0Lmxlbmd0aCA/IGJ1Y2tldC5qb2luKCcrJykgOiAnJywKICBdLmZpbHRlcihCb29sZWFuKS5qb2luKCdfJykgfHwgJ0FsbCc7CiAgZG93bmxvYWRDU1YoJ0FnZWRfSW52ZW50b3J5XycgKyBsYWJlbCArICcuY3N2JywgdG9DU1YoaGVhZGVycywgcm93cykpOwp9CgpmdW5jdGlvbiBleHBvcnRTZHJvcENTVigpIHsKICBjb25zdCBpdGVtcyA9IFJBVy5zZHJvcD8uaXRlbXMgfHwgW107CiAgY29uc3QgbG9jICAgPSBtc0dldFNlbGVjdGVkKCdtcy1kcm9wbG9jJyk7CiAgY29uc3QgY29vcmQgPSBtc0dldFNlbGVjdGVkKCdtcy1kcm9wY29vcmQnKTsKICBjb25zdCBmbGFnICA9IG1zR2V0U2VsZWN0ZWQoJ21zLWRyb3BmbGFnJyk7CiAgY29uc3QgcSAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BTZWFyY2gnKS52YWx1ZS50b0xvd2VyQ2FzZSgpOwogIGxldCByb3dzID0gaXRlbXM7CiAgaWYgKGxvYy5sZW5ndGgpICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gbG9jLmluY2x1ZGVzKHIubG9jYXRpb24pKTsKICBpZiAoY29vcmQubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKGZsYWcubGVuZ3RoKSAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gZmxhZy5pbmNsdWRlcyhzZHJvcEFnZUZsYWcoci5hZ2UpKSk7CiAgaWYgKHEpICAgICByb3dzID0gcm93cy5maWx0ZXIociA9PgogICAgci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICByLnBhcnRfbm8udG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLmNvb3JkaW5hdG9yLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwogIGNvbnN0IGhlYWRlcnMgPSBbJ0xvY2F0aW9uJywnT3JkZXIgIycsJ1Byb2plY3QnLCdDb29yZGluYXRvcicsJ1BhcnQgTm8uJywnU2VyaWFsICMnLCdQYXJ0IEdyb3VwJywnUXR5JywnQWdlIChkYXlzKScsJ1ZhbHVlJywnU2hpcCBEYXRlJywnT3JkZXIgU3RhdHVzJywnQWdlIEZsYWcnXTsKICBjb25zdCBjc3ZSb3dzID0gcm93cy5tYXAociA9PiBbci5sb2NhdGlvbiwgci5vcmRlciwgci5wcm9qZWN0LCByLmNvb3JkaW5hdG9yLCByLnBhcnRfbm8sIHIuc2VyaWFsfHwnJywgci5wYXJ0X2dyb3VwLCByLnF0eSwgci5hZ2UsIHIudmFsdWUsIHIuc2hpcF9kYXRlfHwnJywgci5zdGF0dXMsIHIuZmxhZ10pOwogIGRvd25sb2FkQ1NWKCdTRHJvcF9SZXZpZXcuY3N2JywgdG9DU1YoaGVhZGVycywgY3N2Um93cykpOwp9CgovLyDilIDilIAgQ2xvc2VkIE9yZGVycyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyQ2xvc2VkS3BpcygpIHsKICBjb25zdCBrID0gUkFXLmNsb3NlZD8ua3BpczsKICBpZiAoIWspIHJldHVybjsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xLcGlJdGVtcycpLnRleHRDb250ZW50ICA9IGZtdE4oay50b3RhbF9pdGVtcyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NsS3BpVmFsdWUnKS50ZXh0Q29udGVudCAgPSBmbXQkKGsudG90YWxfdmFsdWUpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbEtwaU9yZGVycycpLnRleHRDb250ZW50ID0gZm10TihrLnVuaXF1ZV9vcmRlcnMpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbEtwaUF2ZycpLnRleHRDb250ZW50ICAgID0gZm10RihrLmF2Z19hZ2UpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbEtwaU1heCcpLnRleHRDb250ZW50ICAgID0gZm10TihrLm1heF9hZ2UpOwp9CgpmdW5jdGlvbiByZXNldENsb3NlZEZpbHRlcnMoKSB7CiAgbXNSZXNldCgnbXMtY2xjb29yZCcpOyBtc1Jlc2V0KCdtcy1jbGJ1Y2tldCcpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbG9zZWRTZWFyY2gnKS52YWx1ZSA9ICcnOwogIHJlbmRlckNsb3NlZCgpOwp9CgpmdW5jdGlvbiByZW5kZXJDbG9zZWQoKSB7CiAgY29uc3QgaXRlbXMgPSBSQVcuY2xvc2VkPy5pdGVtcyB8fCBbXTsKICBjb25zdCBjb29yZCAgPSBtc0dldFNlbGVjdGVkKCdtcy1jbGNvb3JkJyk7CiAgY29uc3QgYnVja2V0ID0gbXNHZXRTZWxlY3RlZCgnbXMtY2xidWNrZXQnKTsKICBjb25zdCBxICAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xvc2VkU2VhcmNoJykudmFsdWUudG9Mb3dlckNhc2UoKTsKICBsZXQgcm93cyA9IGl0ZW1zOwogIGlmIChjb29yZC5sZW5ndGgpICByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKGJ1Y2tldC5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGJ1Y2tldC5pbmNsdWRlcyhyLmFnZV9idWNrZXQpKTsKICBpZiAocSkgICAgICByb3dzID0gcm93cy5maWx0ZXIociA9PgogICAgci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICByLnBhcnRfbm8udG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nsb3NlZE1ldGEnKS50ZXh0Q29udGVudCA9IHJvd3MubGVuZ3RoICsgJyBpdGVtcyc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nsb3NlZEJvZHknKS5pbm5lckhUTUwgPSByb3dzLm1hcChyID0+IGAKICAgIDx0cj4KICAgICAgPHRkPjxzdHJvbmc+JHtyLm9yZGVyfTwvc3Ryb25nPjwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0id3JhcCI+JHtyLnByb2plY3R9PC90ZD4KICAgICAgPHRkPiR7ci5jb29yZGluYXRvcn08L3RkPgogICAgICA8dGQ+JHtyLnBhcnRfbm98fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3Iuc2VyaWFsfHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtyLnBhcnRfZ3JvdXB9PC90ZD4KICAgICAgPHRkPiR7ci5sb2NhdGlvbn08L3RkPgogICAgICA8dGQ+JHtyLnNoaXBfZGF0ZXx8J+KAlCd9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10TihyLnF0eSl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7bWF4QWdlVGFnKHIuYWdlKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXQkKHIudmFsdWUpfTwvdGQ+CiAgICAgIDx0ZD4ke2FnZUJ1Y2tldFRhZyhyLmFnZV9idWNrZXQpfTwvdGQ+CiAgICA8L3RyPmApLmpvaW4oJycpIHx8ICc8dHI+PHRkIGNvbHNwYW49IjEwIiBzdHlsZT0idGV4dC1hbGlnbjpjZW50ZXI7cGFkZGluZzozMnB4O2NvbG9yOnZhcigtLW11dGVkKSI+Tm8gaXRlbXMgbWF0Y2ggZmlsdGVycy48L3RkPjwvdHI+JzsKICBpZiAoIXJvd3MubGVuZ3RoKSB7IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbG9zZWRGb290JykuaW5uZXJIVE1MID0gJyc7IHJldHVybjsgfQogIGNvbnN0IHRvdCA9IHJvd3MucmVkdWNlKChhLHIpID0+ICh7cXR5OmEucXR5K3IucXR5LCB2YWx1ZTphLnZhbHVlK3IudmFsdWV9KSwge3F0eTowLHZhbHVlOjB9KTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xvc2VkRm9vdCcpLmlubmVySFRNTCA9IGA8dHI+PHRkIGNvbHNwYW49IjkiPlRPVEFMIOKAlCAke3Jvd3MubGVuZ3RofSBpdGVtczwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10Tih0b3QucXR5KX08L3RkPjx0ZD48L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQodG90LnZhbHVlKX08L3RkPjx0ZD48L3RkPjwvdHI+YDsKfQoKZnVuY3Rpb24gZXhwb3J0Q2xvc2VkQ1NWKCkgewogIGNvbnN0IGl0ZW1zID0gUkFXLmNsb3NlZD8uaXRlbXMgfHwgW107CiAgY29uc3QgY29vcmQgID0gbXNHZXRTZWxlY3RlZCgnbXMtY2xjb29yZCcpOwogIGNvbnN0IGJ1Y2tldCA9IG1zR2V0U2VsZWN0ZWQoJ21zLWNsYnVja2V0Jyk7CiAgY29uc3QgcSAgICAgID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nsb3NlZFNlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBpdGVtczsKICBpZiAoY29vcmQubGVuZ3RoKSAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gY29vcmQuaW5jbHVkZXMoci5jb29yZGluYXRvcikpOwogIGlmIChidWNrZXQubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIociA9PiBidWNrZXQuaW5jbHVkZXMoci5hZ2VfYnVja2V0KSk7CiAgaWYgKHEpICAgICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucGFydF9uby50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpKTsKICBjb25zdCBoZWFkZXJzID0gWydPcmRlciAjJywnUHJvamVjdCcsJ0Nvb3JkaW5hdG9yJywnUGFydCBOby4nLCdTZXJpYWwgIycsJ1BhcnQgR3JvdXAnLCdMb2NhdGlvbicsJ1NoaXAgRGF0ZScsJ1F0eScsJ0FnZSAoZGF5cyknLCdWYWx1ZScsJ0FnZSBCdWNrZXQnLCdPcmRlciBTdGF0dXMnXTsKICBjb25zdCBjc3ZSb3dzID0gcm93cy5tYXAociA9PiBbci5vcmRlcixyLnByb2plY3Qsci5jb29yZGluYXRvcixyLnBhcnRfbm8sci5zZXJpYWx8fCcnLHIucGFydF9ncm91cCxyLmxvY2F0aW9uLHIuc2hpcF9kYXRlfHwnJyxyLnF0eSxyLmFnZSxyLnZhbHVlLHIuYWdlX2J1Y2tldCxyLnN0YXR1c10pOwogIGRvd25sb2FkQ1NWKCdDbG9zZWRfT3JkZXJzLmNzdicsIHRvQ1NWKGhlYWRlcnMsIGNzdlJvd3MpKTsKfQoKLy8g4pSA4pSAIE9mZnNpdGUgU3RvcmFnZSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyT2Zmc2l0ZUtwaXMoKSB7CiAgY29uc3QgayA9IFJBVy5vZmZzaXRlPy5rcGlzOwogIGlmICghaykgcmV0dXJuOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZktwaUl0ZW1zJykudGV4dENvbnRlbnQgID0gZm10TihrLnRvdGFsX2l0ZW1zKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2ZLcGlWYWx1ZScpLnRleHRDb250ZW50ICA9IGZtdCQoay50b3RhbF92YWx1ZSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29mS3BpT3JkZXJzJykudGV4dENvbnRlbnQgPSBmbXROKGsudW5pcXVlX29yZGVycyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29mS3BpQXZnJykudGV4dENvbnRlbnQgICAgPSBmbXRGKGsuYXZnX2FnZSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29mS3BpTWF4JykudGV4dENvbnRlbnQgICAgPSBmbXROKGsubWF4X2FnZSk7Cn0KCmZ1bmN0aW9uIHJlbmRlck9mZnNpdGVMb2NHcmlkKCkgewogIGNvbnN0IGdyb3VwcyA9IFJBVy5vZmZzaXRlPy5ieV9ncm91cCB8fCBbXTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUxvY0dyaWQnKS5pbm5lckhUTUwgPSBncm91cHMubWFwKGcgPT4gYAogICAgPGRpdiBjbGFzcz0ic2Ryb3AtbG9jLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtbmFtZSI+JHtnLmdyb3VwfTwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtaXRlbXMiPiR7Zm10TihnLm9yZGVycyl9IDxzcGFuIHN0eWxlPSJmb250LXNpemU6MTNweDtmb250LXdlaWdodDo0MDA7Y29sb3I6dmFyKC0tbXV0ZWQpIj5vcmRlcnM8L3NwYW4+ICZuYnNwO8K3Jm5ic3A7ICR7Zm10TihnLml0ZW1zKX0gPHNwYW4gc3R5bGU9ImZvbnQtc2l6ZToxM3B4O2ZvbnQtd2VpZ2h0OjQwMDtjb2xvcjp2YXIoLS1tdXRlZCkiPmNvbnRhaW5lcnM8L3NwYW4+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNkcm9wLWxvYy12YWwiPiR7Zm10JChnLnZhbHVlKX0gwrcgYXZnICR7Zm10RihnLmF2Z19hZ2UpfWQgwrcgbWF4ICR7Zm10TihnLm1heF9hZ2UpfWQ8L2Rpdj4KICAgIDwvZGl2PmApLmpvaW4oJycpOwp9CgpmdW5jdGlvbiByZXNldE9mZnNpdGVGaWx0ZXJzKCkgewogIG1zUmVzZXQoJ21zLW9mYnVpbGQnKTsgbXNSZXNldCgnbXMtb2Zjb29yZCcpOyBtc1Jlc2V0KCdtcy1vZnN0YXR1cycpOyBtc1Jlc2V0KCdtcy1vZmZsYWcnKTsgbXNSZXNldCgnbXMtb2ZkaXNwJyk7IG1zUmVzZXQoJ21zLW9mYmlsbGVkJyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29mZnNpdGVTZWFyY2gnKS52YWx1ZSA9ICcnOwogIHJlbmRlck9mZnNpdGUoKTsKfQoKZnVuY3Rpb24gcmVuZGVyT2Zmc2l0ZSgpIHsKICBjb25zdCBpdGVtcyAgPSBSQVcub2Zmc2l0ZT8uaXRlbXMgfHwgW107CiAgY29uc3QgZ3JvdXAgID0gbXNHZXRTZWxlY3RlZCgnbXMtb2ZidWlsZCcpOwogIGNvbnN0IGNvb3JkICA9IG1zR2V0U2VsZWN0ZWQoJ21zLW9mY29vcmQnKTsKICBjb25zdCBzdGF0dXMgPSBtc0dldFNlbGVjdGVkKCdtcy1vZnN0YXR1cycpOwogIGNvbnN0IGZsYWcgICA9IG1zR2V0U2VsZWN0ZWQoJ21zLW9mZmxhZycpOwogIGNvbnN0IGRpc3AgICA9IG1zR2V0U2VsZWN0ZWQoJ21zLW9mZGlzcCcpOwogIGNvbnN0IHEgICAgICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZmZzaXRlU2VhcmNoJykudmFsdWUudG9Mb3dlckNhc2UoKTsKICBsZXQgcm93cyA9IGl0ZW1zOwogIGlmIChncm91cC5sZW5ndGgpICByb3dzID0gcm93cy5maWx0ZXIociA9PiBncm91cC5pbmNsdWRlcyhyLmJ1aWxkaW5nKSk7CiAgaWYgKGNvb3JkLmxlbmd0aCkgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGNvb3JkLmluY2x1ZGVzKHIuY29vcmRpbmF0b3IpKTsKICBpZiAoc3RhdHVzLmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gc3RhdHVzLmluY2x1ZGVzKHIuc3RhdHVzKSk7CiAgaWYgKGZsYWcubGVuZ3RoKSAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGZsYWcuaW5jbHVkZXMoci5mbGFnKSk7CiAgaWYgKGRpc3AubGVuZ3RoKSAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHsKICAgIGNvbnN0IGQgPSBESVNQT1NJVElPTlNbci5vcmRlcl0gfHwgJyc7CiAgICByZXR1cm4gZGlzcC5pbmNsdWRlcyhkIHx8ICcoTm9uZSknKTsKICB9KTsKICBjb25zdCBiaWxsZWQgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmJpbGxlZCcpOwogIGlmIChiaWxsZWQubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIociA9PiB7CiAgICBjb25zdCBiID0gci5iaWxsZWQgPyAnQmlsbGVkJyA6ICdOb3QgQmlsbGVkJzsKICAgIHJldHVybiBiaWxsZWQuaW5jbHVkZXMoYik7CiAgfSk7CiAgaWYgKHEpICAgICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4KICAgIHIub3JkZXIudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLnByb2plY3QudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fAogICAgci5wYXJ0X25vLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwgci5jb29yZGluYXRvci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZU1ldGEnKS50ZXh0Q29udGVudCA9IHJvd3MubGVuZ3RoLnRvTG9jYWxlU3RyaW5nKCkgKyAnIGl0ZW1zJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUJvZHknKS5pbm5lckhUTUwgPSByb3dzLm1hcChyID0+IHsKICAgIC8vIFBhcnNlIHNoaXAgZGF0ZSBmb3IgY29uZGl0aW9uYWwgZm9ybWF0dGluZwogICAgY29uc3Qgc2QgPSByLnNoaXBfZGF0ZSAmJiByLnNoaXBfZGF0ZSAhPT0gJ+KAlCcgPyBuZXcgRGF0ZShyLnNoaXBfZGF0ZSkgOiBudWxsOwogICAgY29uc3QgdG9kYXkgPSBuZXcgRGF0ZSgpOyB0b2RheS5zZXRIb3VycygwLDAsMCwwKTsKICAgIGNvbnN0IGRheXNPdXQgPSBzZCA/IE1hdGgucm91bmQoKHNkIC0gdG9kYXkpIC8gODY0MDAwMDApIDogbnVsbDsKICAgIGNvbnN0IHN0YXR1c1N0b3JhZ2UgPSByLnN0YXR1cyAmJiByLnN0YXR1cy50b0xvd2VyQ2FzZSgpLnN0YXJ0c1dpdGgoJ3N0b3JhZ2UnKTsKICAgIGNvbnN0IGlzRG93bnN0YWlycyA9IHIuYnVpbGRpbmcgPT09ICdEb3duc3RhaXJzIFN0b3JhZ2UnOwogICAgY29uc3QgaXNOb3J0aCA9IHIuYnVpbGRpbmcgPT09ICdOb3J0aCBXYXJlaG91c2UnOwogICAgY29uc3QgZGF0ZUZsYWcgPSAoaXNEb3duc3RhaXJzICYmIHN0YXR1c1N0b3JhZ2UgJiYgZGF5c091dCAhPT0gbnVsbCAmJiBkYXlzT3V0ID4gNDUpIHx8CiAgICAgICAgICAgICAgICAgICAgIChpc05vcnRoICAgICAmJiBzdGF0dXNTdG9yYWdlICYmIGRheXNPdXQgIT09IG51bGwgJiYgZGF5c091dCA+IDYwKTsKICAgIGNvbnN0IGRpc3AgPSBESVNQT1NJVElPTlNbci5vcmRlcnx8JyddIHx8ICcnOwogICAgLy8gUmVkIGlmOiBzdG9yYWdlIHN0YXR1cywgbm90IFNPUCBCdWlsZCBBaGVhZCwgYW5kIG5vdCBiaWxsZWQKICAgIGNvbnN0IGJpbGwgPSBCSUxMSU5HW3Iub3JkZXJdIHx8IG51bGw7CiAgICBjb25zdCBuZWVkc0JpbGxpbmcgPSBzdGF0dXNTdG9yYWdlICYmIGRpc3AgIT09ICdTT1AgQnVpbGQgQWhlYWQnICYmICFiaWxsOwogICAgY29uc3Qgcm93U3R5bGUgPSBuZWVkc0JpbGxpbmcKICAgICAgPyAnYmFja2dyb3VuZDojZmZmMGYwO2JvcmRlci1sZWZ0OjNweCBzb2xpZCAjZGMzNTQ1OycKICAgICAgOiBkYXRlRmxhZyA/ICdiYWNrZ3JvdW5kOiNmZmYzY2Q7Ym9yZGVyLWxlZnQ6M3B4IHNvbGlkICNmZmMxMDc7JyA6ICcnOwogICAgcmV0dXJuIGA8dHIgc3R5bGU9IiR7cm93U3R5bGV9Ij4KICAgICAgPHRkPiR7ci5idWlsZGluZ3x8J+KAlCd9PC90ZD4KICAgICAgPHRkPiR7ci5sb2NhdGlvbl9ncm91cHx8J+KAlCd9PC90ZD4KICAgICAgPHRkPjxzdHJvbmc+JHtyLm9yZGVyfHwn4oCUJ308L3N0cm9uZz48L3RkPgogICAgICA8dGQgY2xhc3M9IndyYXAiPiR7ci5wcm9qZWN0fHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtyLmNvb3JkaW5hdG9yfHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtyLmxvY2F0aW9ufHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtyLnBhcnRfbm98fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3Iuc2VyaWFsfHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtyLnNoaXBfZGF0ZXx8J+KAlCd9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10TihyLmNvbnRhaW5lcnMpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQoci52YWx1ZSl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10RihyLmF2Z19hZ2UpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke21heEFnZVRhZyhyLm1heF9hZ2UpfTwvdGQ+CiAgICAgIDx0ZD4ke3Iuc3RhdHVzfHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtmbGFnVGFnKHIuZmxhZyl9PC90ZD4KICAgICAgPHRkPiR7YmlsbCA/IGJpbGwuY2hhcmdlX3R5cGUgOiAn4oCUJ308L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtiaWxsID8gZm10JChiaWxsLmNoYXJnZV9hbW91bnQpIDogbmVlZHNCaWxsaW5nID8gJzxzcGFuIHN0eWxlPSJjb2xvcjojZGMzNTQ1O2ZvbnQtd2VpZ2h0OjYwMCI+4pqgIE5vdCBCaWxsZWQ8L3NwYW4+JyA6ICfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke2JpbGwgPyBiaWxsLm5vdGUgOiAn4oCUJ308L3RkPgogICAgICA8dGQ+PHNlbGVjdCBjbGFzcz0iZGlzcC1zZWxlY3QiIGRhdGEtb3JkZXI9IiR7ci5vcmRlcnx8Jyd9IiBvbmNoYW5nZT0ic2F2ZURpc3Bvc2l0aW9uKHRoaXMpIiBzdHlsZT0iZm9udC1zaXplOjExcHg7cGFkZGluZzoycHggNHB4O2JvcmRlcjoxcHggc29saWQgI2NjYztib3JkZXItcmFkaXVzOjRweDt3aWR0aDoxNDBweDsiPgogICAgICAgIDxvcHRpb24gdmFsdWU9IiI+4oCUIFNlbGVjdCDigJQ8L29wdGlvbj4KICAgICAgICA8b3B0aW9uIHZhbHVlPSJTT1AgQnVpbGQgQWhlYWQiICR7ZGlzcD09PSdTT1AgQnVpbGQgQWhlYWQnPydzZWxlY3RlZCc6Jyd9PlNPUCBCdWlsZCBBaGVhZDwvb3B0aW9uPgogICAgICAgIDxvcHRpb24gdmFsdWU9IlN0b3JhZ2UgQ2hhcmdlZCIgJHtkaXNwPT09J1N0b3JhZ2UgQ2hhcmdlZCc/J3NlbGVjdGVkJzonJ30+U3RvcmFnZSBDaGFyZ2VkPC9vcHRpb24+CiAgICAgICAgPG9wdGlvbiB2YWx1ZT0iSW5zdGFsbGF0aW9uIERlbGF5IiAke2Rpc3A9PT0nSW5zdGFsbGF0aW9uIERlbGF5Jz8nc2VsZWN0ZWQnOicnfT5JbnN0YWxsYXRpb24gRGVsYXk8L29wdGlvbj4KICAgICAgPC9zZWxlY3Q+PC90ZD4KICAgIDwvdHI+YDsKICB9KS5qb2luKCcnKSB8fCAnPHRyPjx0ZCBjb2xzcGFuPSIxMCIgc3R5bGU9InRleHQtYWxpZ246Y2VudGVyO3BhZGRpbmc6MzJweDtjb2xvcjp2YXIoLS1tdXRlZCkiPk5vIGl0ZW1zIG1hdGNoIGZpbHRlcnMuPC90ZD48L3RyPic7CiAgaWYgKCFyb3dzLmxlbmd0aCkgeyBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUZvb3QnKS5pbm5lckhUTUwgPSAnJzsgcmV0dXJuOyB9CiAgY29uc3QgdG90ID0gcm93cy5yZWR1Y2UoKGEscikgPT4gKHtjb250YWluZXJzOmEuY29udGFpbmVycytyLmNvbnRhaW5lcnMsIHZhbHVlOmEudmFsdWUrci52YWx1ZX0pLCB7Y29udGFpbmVyczowLHZhbHVlOjB9KTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUZvb3QnKS5pbm5lckhUTUwgPSBgPHRyPjx0ZCBjb2xzcGFuPSIxMiI+VE9UQUwg4oCUICR7cm93cy5sZW5ndGgudG9Mb2NhbGVTdHJpbmcoKX0gcm93czwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10Tih0b3QuY29udGFpbmVycyl9PC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXQkKHRvdC52YWx1ZSl9PC90ZD48dGQgY29sc3Bhbj0iNCI+PC90ZD48L3RyPmA7Cn0KCmZ1bmN0aW9uIGV4cG9ydE9mZnNpdGVDU1YoKSB7CiAgY29uc3QgaXRlbXMgID0gUkFXLm9mZnNpdGU/Lml0ZW1zIHx8IFtdOwogIGNvbnN0IGdyb3VwICA9IG1zR2V0U2VsZWN0ZWQoJ21zLW9mYnVpbGQnKTsKICBjb25zdCBjb29yZCAgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmNvb3JkJyk7CiAgY29uc3Qgc3RhdHVzID0gbXNHZXRTZWxlY3RlZCgnbXMtb2ZzdGF0dXMnKTsKICBjb25zdCBmbGFnICAgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmZsYWcnKTsKICBjb25zdCBkaXNwICAgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmRpc3AnKTsKICBjb25zdCBxICAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZVNlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBpdGVtczsKICBpZiAoZ3JvdXAubGVuZ3RoKSAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gZ3JvdXAuaW5jbHVkZXMoci5idWlsZGluZykpOwogIGlmIChjb29yZC5sZW5ndGgpICByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKHN0YXR1cy5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHN0YXR1cy5pbmNsdWRlcyhyLnN0YXR1cykpOwogIGlmIChmbGFnLmxlbmd0aCkgICByb3dzID0gcm93cy5maWx0ZXIociA9PiBmbGFnLmluY2x1ZGVzKHIuZmxhZykpOwogIGlmIChkaXNwLmxlbmd0aCkgICByb3dzID0gcm93cy5maWx0ZXIociA9PiB7CiAgICBjb25zdCBkID0gRElTUE9TSVRJT05TW3Iub3JkZXJdIHx8ICcnOwogICAgcmV0dXJuIGRpc3AuaW5jbHVkZXMoZCB8fCAnKE5vbmUpJyk7CiAgfSk7CiAgY29uc3QgYmlsbGVkID0gbXNHZXRTZWxlY3RlZCgnbXMtb2ZiaWxsZWQnKTsKICBpZiAoYmlsbGVkLmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gewogICAgY29uc3QgYiA9IHIuYmlsbGVkID8gJ0JpbGxlZCcgOiAnTm90IEJpbGxlZCc7CiAgICByZXR1cm4gYmlsbGVkLmluY2x1ZGVzKGIpOwogIH0pOwogIGlmIChxKSAgICAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHIub3JkZXIudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLnByb2plY3QudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLnBhcnRfbm8udG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSk7CiAgY29uc3QgaGVhZGVycyA9IFsnQnVpbGRpbmcnLCdMb2NhdGlvbiBHcm91cCcsJ09yZGVyICMnLCdQcm9qZWN0JywnQ29vcmRpbmF0b3InLCdMb2NhdGlvbicsJ1BhcnQgTm8uJywnU2VyaWFsICMnLCdTaGlwIERhdGUnLCdDb250YWluZXJzJywnVmFsdWUnLCdBdmcgQWdlJywnTWF4IEFnZScsJ09yZGVyIFN0YXR1cycsJ0FnZSBGbGFnJywnQ2hhcmdlIFR5cGUnLCdDaGFyZ2UgQW1vdW50JywnQmlsbGluZyBOb3RlJywnRGlzcG9zaXRpb24nLCdVbmJpbGxlZD8nXTsKICBjb25zdCBjc3ZSb3dzID0gcm93cy5tYXAociA9PiB7IGNvbnN0IGI9QklMTElOR1tyLm9yZGVyXXx8bnVsbDsgY29uc3QgZD1ESVNQT1NJVElPTlNbci5vcmRlcl18fCcnOyBjb25zdCBuYj1yLnN0YXR1cyYmci5zdGF0dXMudG9Mb3dlckNhc2UoKS5zdGFydHNXaXRoKCdzdG9yYWdlJykmJmQhPT0nU09QIEJ1aWxkIEFoZWFkJyYmIWI7IHJldHVybiBbci5idWlsZGluZ3x8Jycsci5sb2NhdGlvbl9ncm91cHx8Jycsci5vcmRlcnx8Jycsci5wcm9qZWN0fHwnJyxyLmNvb3JkaW5hdG9yfHwnJyxyLmxvY2F0aW9ufHwnJyxyLnBhcnRfbm98fCcnLHIuc2VyaWFsfHwnJyxyLnNoaXBfZGF0ZXx8Jycsci5jb250YWluZXJzLHIudmFsdWUsci5hdmdfYWdlLHIubWF4X2FnZSxyLnN0YXR1c3x8Jycsci5mbGFnLGI/Yi5jaGFyZ2VfdHlwZTonJyxiP2IuY2hhcmdlX2Ftb3VudDonJyxiP2Iubm90ZTonJyxkLG5iPydZZXMnOicnXTsgfSk7CiAgZG93bmxvYWRDU1YoJ09mZnNpdGVfU3RvcmFnZS5jc3YnLCB0b0NTVihoZWFkZXJzLCBjc3ZSb3dzKSk7Cn0KCi8vIOKUgOKUgCBCaWxsaW5nIGRhdGEg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmxldCBCSUxMSU5HID0ge307ICAvLyBvcmRlciAtPiB7Y2hhcmdlX2Ftb3VudCwgY2hhcmdlX3R5cGUsIG5vdGUsIHBjfQoKYXN5bmMgZnVuY3Rpb24gbG9hZEJpbGxpbmcoKSB7CiAgdHJ5IHsKICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKCcvYXBpL2JpbGxpbmcnKTsKICAgIEJJTExJTkcgPSBhd2FpdCByZXMuanNvbigpOyAgLy8ge29yZGVyOiB7Y2hhcmdlX2Ftb3VudCwgY2hhcmdlX3R5cGUsIG5vdGV9fQogICAgLy8gS1BJcyBhcmUgY29tcHV0ZWQgYnkgdXBkYXRlRGlzcG9zaXRpb25LUElzKCkgd2hpY2ggcnVucyBhZnRlciB0aGlzCiAgfSBjYXRjaChlKSB7IEJJTExJTkcgPSB7fTsgfQp9CgovLyDilIDilIAgRGlzcG9zaXRpb25zIChwZXJzaXN0ZWQgcGVyLW9yZGVyIHNlbGVjdGlvbnMpIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApsZXQgRElTUE9TSVRJT05TID0ge307Cgphc3luYyBmdW5jdGlvbiBsb2FkRGlzcG9zaXRpb25zKCkgewogIHRyeSB7CiAgICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaCgnL2FwaS9kaXNwb3NpdGlvbnMnKTsKICAgIERJU1BPU0lUSU9OUyA9IGF3YWl0IHJlcy5qc29uKCk7CiAgfSBjYXRjaChlKSB7IERJU1BPU0lUSU9OUyA9IHt9OyB9Cn0KCmZ1bmN0aW9uIHNhdmVEaXNwb3NpdGlvbihzZWxlY3QpIHsKICBjb25zdCBvcmRlciA9IHNlbGVjdC5kYXRhc2V0Lm9yZGVyOwogIGNvbnN0IGRpc3AgID0gc2VsZWN0LnZhbHVlOwogIGlmICghb3JkZXIpIHJldHVybjsKICBESVNQT1NJVElPTlNbb3JkZXJdID0gZGlzcDsKICAvLyBTeW5jIHNhbWUgb3JkZXIncyBkcm9wZG93biBvbiBvdGhlciB0YWIgaWYgcHJlc2VudAogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoYC5kaXNwLXNlbGVjdFtkYXRhLW9yZGVyPSIke29yZGVyfSJdYCkuZm9yRWFjaChzID0+IHsgcy52YWx1ZSA9IGRpc3A7IH0pOwogIC8vIFVwZGF0ZSBkaXNwb3NpdGlvbiBLUEkgY291bnRzCiAgdXBkYXRlRGlzcG9zaXRpb25LUElzKCk7CiAgLy8gUE9TVCB0byBzZXJ2ZXIKICBmZXRjaCgnL2FwaS9zZXRfZGlzcG9zaXRpb24nLCB7CiAgICBtZXRob2Q6ICdQT1NUJywKICAgIGhlYWRlcnM6IHsnQ29udGVudC1UeXBlJzonYXBwbGljYXRpb24vanNvbid9LAogICAgYm9keTogSlNPTi5zdHJpbmdpZnkoe29yZGVyLCBkaXNwb3NpdGlvbjogZGlzcH0pCiAgfSkuY2F0Y2goZSA9PiBjb25zb2xlLmVycm9yKCdEaXNwb3NpdGlvbiBzYXZlIGZhaWxlZCcsIGUpKTsKfQoKZnVuY3Rpb24gaXNTdG9yYWdlU3RhdHVzKHMpIHsgcmV0dXJuIHMgJiYgcy50b0xvd2VyQ2FzZSgpLnN0YXJ0c1dpdGgoJ3N0b3JhZ2UnKTsgfQoKZnVuY3Rpb24gdXBkYXRlRGlzcG9zaXRpb25LUElzKCkgewogIGNvbnN0IGl0ZW1zID0gUkFXPy5vZmZzaXRlPy5pdGVtcyB8fCBbXTsKICBsZXQgc29wT3JkZXJzPTAsc29wVmFsPTAsY2hhcmdlZE9yZGVycz0wLGNoYXJnZWRWYWw9MCx1bmJpbGxlZD0wLGJpbGxlZENvdW50PTAsYmlsbGVkVG90YWw9MDsKICBpdGVtcy5mb3JFYWNoKHIgPT4gewogICAgY29uc3QgZCAgICA9IERJU1BPU0lUSU9OU1tyLm9yZGVyXSB8fCAnJzsKICAgIGNvbnN0IGJpbGwgPSBCSUxMSU5HW3Iub3JkZXJdIHx8IG51bGw7CiAgICBpZiAoZCA9PT0gJ1NPUCBCdWlsZCBBaGVhZCcpIHsgc29wT3JkZXJzKys7ICAgICBzb3BWYWwgICAgICs9IHIudmFsdWV8fDA7IH0KICAgIGlmIChkID09PSAnU3RvcmFnZSBDaGFyZ2VkJykgeyBjaGFyZ2VkT3JkZXJzKys7IGNoYXJnZWRWYWwgKz0gci52YWx1ZXx8MDsgfQogICAgaWYgKGlzU3RvcmFnZVN0YXR1cyhyLnN0YXR1cykgJiYgIWJpbGwgJiYgZCAhPT0gJ1NPUCBCdWlsZCBBaGVhZCcgJiYgZCAhPT0gJ0luc3RhbGxhdGlvbiBEZWxheScpIHVuYmlsbGVkKys7CiAgICBpZiAoYmlsbCkgeyBiaWxsZWRDb3VudCsrOyBiaWxsZWRUb3RhbCArPSBiaWxsLmNoYXJnZV9hbW91bnR8fDA7IH0KICB9KTsKICBjb25zdCBzZXQgPSAoaWQsdmFsKSA9PiB7IGNvbnN0IGVsPWRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsgaWYoZWwpIGVsLnRleHRDb250ZW50PXZhbDsgfTsKICBzZXQoJ29mS3BpU09QJywgICAgICAgICBzb3BPcmRlcnMudG9Mb2NhbGVTdHJpbmcoKSk7CiAgc2V0KCdvZktwaVNPUFZhbCcsICAgICAgZm10JChzb3BWYWwpKTsKICBzZXQoJ29mS3BpQ2hhcmdlZCcsICAgICBjaGFyZ2VkT3JkZXJzLnRvTG9jYWxlU3RyaW5nKCkpOwogIHNldCgnb2ZLcGlDaGFyZ2VkVmFsJywgIGZtdCQoY2hhcmdlZFZhbCkpOwogIHNldCgnb2ZLcGlVbmJpbGxlZCcsICAgIHVuYmlsbGVkLnRvTG9jYWxlU3RyaW5nKCkpOwogIHNldCgnb2ZLcGlCaWxsZWRDb3VudCcsIGJpbGxlZENvdW50LnRvTG9jYWxlU3RyaW5nKCkpOwogIHNldCgnb2ZLcGlCaWxsZWRUb3RhbCcsIGZtdCQoYmlsbGVkVG90YWwpKTsKfQoKLy8g4pSA4pSAIE5ldyB0byBTdG9yYWdlIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApsZXQgTlNfREFUQSA9IFtdOwpsZXQgTlNfTE9BREVEID0gZmFsc2U7Cgphc3luYyBmdW5jdGlvbiBsb2FkTmV3U3RvcmFnZSgpIHsKICBpZiAoTlNfTE9BREVEKSByZXR1cm47CiAgTlNfTE9BREVEID0gdHJ1ZTsKICB0cnkgewogICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goJy9hcGkvbmV3X3N0b3JhZ2UnKTsKICAgIE5TX0RBVEEgPSBhd2FpdCByZXMuanNvbigpOwogICAgY29uc3QgbnNDb29yZHMgPSBbLi4ubmV3IFNldChOU19EQVRBLm1hcChyID0+IHIuY29vcmRpbmF0b3IpKV0uZmlsdGVyKEJvb2xlYW4pLnNvcnQoKTsKICAgIG1zUG9wdWxhdGUoJ21zLW5zY29vcmQnLCBuc0Nvb3JkcywgcmVuZGVyTmV3U3RvcmFnZSk7CiAgICByZW5kZXJOZXdTdG9yYWdlKCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbmV3c3RvcmFnZUJhZGdlJykudGV4dENvbnRlbnQgPSBOU19EQVRBLmxlbmd0aDsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc0twaVRvdGFsJykudGV4dENvbnRlbnQgPSBOU19EQVRBLmxlbmd0aC50b0xvY2FsZVN0cmluZygpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ25zS3BpVmFsdWUnKS50ZXh0Q29udGVudCA9IGZtdCQoTlNfREFUQS5yZWR1Y2UoKGEscikgPT4gYSArIHIudmFsdWUsIDApKTsKICAgIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKS50b0xvY2FsZURhdGVTdHJpbmcoJ2VuLVVTJywge3llYXI6J251bWVyaWMnLG1vbnRoOidsb25nJyxkYXk6J251bWVyaWMnfSk7CiAgICBjb25zdCBuZXdUb2RheSA9IE5TX0RBVEEuZmlsdGVyKHIgPT4gci5kZXRlY3RlZCA9PT0gdG9kYXkpLmxlbmd0aDsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc0twaU5ldycpLnRleHRDb250ZW50ID0gbmV3VG9kYXkudG9Mb2NhbGVTdHJpbmcoKTsKICB9IGNhdGNoKGUpIHsKICAgIGNvbnNvbGUuZXJyb3IoJ05ldyBzdG9yYWdlIGxvYWQgZmFpbGVkJywgZSk7CiAgfQp9CgpmdW5jdGlvbiByZXNldE5ld1N0b3JhZ2VGaWx0ZXJzKCkgewogIG1zUmVzZXQoJ21zLW5zY29vcmQnKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbnNTZWFyY2gnKS52YWx1ZSA9ICcnOwogIHJlbmRlck5ld1N0b3JhZ2UoKTsKfQoKZnVuY3Rpb24gcmVuZGVyTmV3U3RvcmFnZSgpIHsKICBjb25zdCBjb29yZCA9IG1zR2V0U2VsZWN0ZWQoJ21zLW5zY29vcmQnKTsKICBjb25zdCBxICAgICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc1NlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBOU19EQVRBOwogIGlmIChjb29yZC5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGNvb3JkLmluY2x1ZGVzKHIuY29vcmRpbmF0b3IpKTsKICBpZiAocSkgICAgICAgICAgICByb3dzID0gcm93cy5maWx0ZXIociA9PgogICAgci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICByLmNvb3JkaW5hdG9yLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc01ldGEnKS50ZXh0Q29udGVudCA9IHJvd3MubGVuZ3RoICsgJyBvcmRlcnMnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc0JvZHknKS5pbm5lckhUTUwgPSByb3dzLm1hcChyID0+IHsKICAgIGNvbnN0IHB1c2hlZCA9IHIuc2hpcF9kYXRlX3B1c2hlZDsKICAgIGNvbnN0IHJvd0JnICA9IHB1c2hlZCA/ICdiYWNrZ3JvdW5kOiNmZmYwZjA7Ym9yZGVyLWxlZnQ6M3B4IHNvbGlkICNkYzM1NDU7JyA6ICdiYWNrZ3JvdW5kOiNmZmY4Zjg7JzsKICAgIGNvbnN0IGRpc3BWYWwgPSBESVNQT1NJVElPTlNbci5vcmRlcl0gfHwgJyc7CiAgICByZXR1cm4gYDx0ciBzdHlsZT0iJHtyb3dCZ30iPgogICAgICA8dGQ+JHtyLmRldGVjdGVkfTwvdGQ+CiAgICAgIDx0ZD48c3Ryb25nPiR7ci5vcmRlcn08L3N0cm9uZz48L3RkPgogICAgICA8dGQgY2xhc3M9IndyYXAiPiR7ci5wcm9qZWN0fTwvdGQ+CiAgICAgIDx0ZD4ke3IuY29vcmRpbmF0b3J9PC90ZD4KICAgICAgPHRkIHN0eWxlPSJjb2xvcjojODg4Ij4ke3IucHJldl9zdGF0dXN9PC90ZD4KICAgICAgPHRkIHN0eWxlPSJjb2xvcjojOEIwMDAwO2ZvbnQtd2VpZ2h0OjYwMCI+JHtyLm5ld19zdGF0dXN9PC90ZD4KICAgICAgPHRkPiR7cHVzaGVkID8gJzxzcGFuIHN0eWxlPSJjb2xvcjojZGMzNTQ1O2ZvbnQtd2VpZ2h0OjYwMCIgdGl0bGU9IlNoaXAgZGF0ZSBwdXNoZWQgb3V0IGZyb20gJyArIChyLnByZXZfc2hpcF9kYXRlfHwnPycpICsgJyI+4pqgICcgKyByLnNoaXBfZGF0ZSArICc8L3NwYW4+JyA6IHIuc2hpcF9kYXRlfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4oci5jb250YWluZXJzKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXQkKHIudmFsdWUpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdEYoci5hdmdfYWdlKX08L3RkPgogICAgICA8dGQ+PHNlbGVjdCBjbGFzcz0iZGlzcC1zZWxlY3QiIGRhdGEtb3JkZXI9IiR7ci5vcmRlcn0iIG9uY2hhbmdlPSJzYXZlRGlzcG9zaXRpb24odGhpcykiIHN0eWxlPSJmb250LXNpemU6MTFweDtwYWRkaW5nOjJweCA0cHg7Ym9yZGVyOjFweCBzb2xpZCAjY2NjO2JvcmRlci1yYWRpdXM6NHB4O3dpZHRoOjE0MHB4OyI+CiAgICAgICAgPG9wdGlvbiB2YWx1ZT0iIj7igJQgU2VsZWN0IOKAlDwvb3B0aW9uPgogICAgICAgIDxvcHRpb24gdmFsdWU9IlNPUCBCdWlsZCBBaGVhZCIgJHtkaXNwVmFsPT09J1NPUCBCdWlsZCBBaGVhZCc/J3NlbGVjdGVkJzonJ30+U09QIEJ1aWxkIEFoZWFkPC9vcHRpb24+CiAgICAgICAgPG9wdGlvbiB2YWx1ZT0iU3RvcmFnZSBDaGFyZ2VkIiAke2Rpc3BWYWw9PT0nU3RvcmFnZSBDaGFyZ2VkJz8nc2VsZWN0ZWQnOicnfT5TdG9yYWdlIENoYXJnZWQ8L29wdGlvbj4KICAgICAgICA8b3B0aW9uIHZhbHVlPSJJbnN0YWxsYXRpb24gRGVsYXkiICR7ZGlzcFZhbD09PSdJbnN0YWxsYXRpb24gRGVsYXknPydzZWxlY3RlZCc6Jyd9Pkluc3RhbGxhdGlvbiBEZWxheTwvb3B0aW9uPgogICAgICA8L3NlbGVjdD48L3RkPgogICAgPC90cj5gOwogIH0pLmpvaW4oJycpIHx8ICc8dHI+PHRkIGNvbHNwYW49IjExIiBzdHlsZT0idGV4dC1hbGlnbjpjZW50ZXI7cGFkZGluZzozMnB4O2NvbG9yOnZhcigtLW11dGVkKSI+Tm8gb3JkZXJzIGRldGVjdGVkIHlldC4gVXBsb2FkIGZpbGVzIG9uIGNvbnNlY3V0aXZlIGRheXMgdG8gc3RhcnQgdHJhY2tpbmcgY2hhbmdlcy48L3RkPjwvdHI+JzsKICBpZiAoIXJvd3MubGVuZ3RoKSB7IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc0Zvb3QnKS5pbm5lckhUTUwgPSAnJzsgcmV0dXJuOyB9CiAgY29uc3QgdG90ID0gcm93cy5yZWR1Y2UoKGEscikgPT4gKHtjb250YWluZXJzOiBhLmNvbnRhaW5lcnMrci5jb250YWluZXJzLCB2YWx1ZTogYS52YWx1ZStyLnZhbHVlfSksIHtjb250YWluZXJzOjAsIHZhbHVlOjB9KTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbnNGb290JykuaW5uZXJIVE1MID0gYDx0cj48dGQgY29sc3Bhbj0iOCI+VE9UQUwg4oCUICR7cm93cy5sZW5ndGh9IG9yZGVyczwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10Tih0b3QuY29udGFpbmVycyl9PC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXQkKHRvdC52YWx1ZSl9PC90ZD48dGQ+PC90ZD48L3RyPmA7Cn0KCmZ1bmN0aW9uIGV4cG9ydE5ld1N0b3JhZ2VDU1YoKSB7CiAgY29uc3QgY29vcmQgPSBtc0dldFNlbGVjdGVkKCdtcy1uc2Nvb3JkJyk7CiAgY29uc3QgcSAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbnNTZWFyY2gnKS52YWx1ZS50b0xvd2VyQ2FzZSgpOwogIGxldCByb3dzID0gTlNfREFUQTsKICBpZiAoY29vcmQubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKHEpICAgICAgICAgICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpKTsKICBjb25zdCBoZWFkZXJzID0gWydEZXRlY3RlZCBEYXRlJywnT3JkZXIgIycsJ1Byb2plY3QnLCdDb29yZGluYXRvcicsJ1ByZXZpb3VzIFN0YXR1cycsJ05ldyBTdGF0dXMnLCdTaGlwIERhdGUnLCdTaGlwIERhdGUgUHVzaGVkPycsJ0NvbnRhaW5lcnMnLCdWYWx1ZScsJ0F2ZyBBZ2UnLCdEaXNwb3NpdGlvbiddOwogIGNvbnN0IGNzdlJvd3MgPSByb3dzLm1hcChyID0+IFtyLmRldGVjdGVkLCByLm9yZGVyLCByLnByb2plY3QsIHIuY29vcmRpbmF0b3IsIHIucHJldl9zdGF0dXMsIHIubmV3X3N0YXR1cywgci5zaGlwX2RhdGUsIHIuc2hpcF9kYXRlX3B1c2hlZCA/ICdZZXMgKHdhcyAnICsgKHIucHJldl9zaGlwX2RhdGV8fCc/JykgKyAnKScgOiAnTm8nLCByLmNvbnRhaW5lcnMsIHIudmFsdWUsIHIuYXZnX2FnZSwgRElTUE9TSVRJT05TW3Iub3JkZXJdfHwnJ10pOwogIGRvd25sb2FkQ1NWKCdOZXdfdG9fU3RvcmFnZS5jc3YnLCB0b0NTVihoZWFkZXJzLCBjc3ZSb3dzKSk7Cn0KCi8vIOKUgOKUgCBTdG9yYWdlIEpvYnMgUGFydHMgSW5jb21wbGV0ZSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKbGV0IFNKUElfREFUQSA9IFtdOwpsZXQgU0pQSV9MT0FERUQgPSBmYWxzZTsKCmFzeW5jIGZ1bmN0aW9uIGxvYWRTanBpKCkgewogIGlmIChTSlBJX0xPQURFRCkgcmV0dXJuOwogIFNKUElfTE9BREVEID0gdHJ1ZTsKICB0cnkgewogICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goJy9hcGkvYmFja29yZGVyJyk7CiAgICBTSlBJX0RBVEEgPSBhd2FpdCByZXMuanNvbigpOwogICAgLy8gUG9wdWxhdGUgZmlsdGVycwogICAgY29uc3QgY29vcmRzICAgPSBbLi4ubmV3IFNldChTSlBJX0RBVEEubWFwKHIgPT4gci5jb29yZGluYXRvcikpXS5maWx0ZXIoQm9vbGVhbikuc29ydCgpOwogICAgY29uc3Qgc3RhdHVzZXMgPSBbLi4ubmV3IFNldChTSlBJX0RBVEEubWFwKHIgPT4gci5saW5lX3N0YXR1cykpXS5maWx0ZXIoQm9vbGVhbikuc29ydCgpOwogICAgbXNQb3B1bGF0ZSgnbXMtc2pwaWNvb3JkJywgIGNvb3JkcywgICByZW5kZXJTanBpKTsKICAgIG1zUG9wdWxhdGUoJ21zLXNqcGlzdGF0dXMnLCBzdGF0dXNlcywgcmVuZGVyU2pwaSk7CiAgICAvLyBLUElzCiAgICBjb25zdCBwcm9qZWN0cyA9IG5ldyBTZXQoU0pQSV9EQVRBLm1hcChyID0+IHIuc2hpcF90bykpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NqcGlCYWRnZScpLnRleHRDb250ZW50ICAgICAgICA9IHByb2plY3RzLnNpemUudG9Mb2NhbGVTdHJpbmcoKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzanBpS3BpUHJvamVjdHMnKS50ZXh0Q29udGVudCAgPSBwcm9qZWN0cy5zaXplLnRvTG9jYWxlU3RyaW5nKCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2pwaUtwaUxpbmVzJykudGV4dENvbnRlbnQgICAgID0gU0pQSV9EQVRBLmxlbmd0aC50b0xvY2FsZVN0cmluZygpOwogICAgLy8gU3RvcmFnZSB2YWx1ZSDigJQgb25lIHBlciBwcm9qZWN0IChhdm9pZCBkb3VibGUgY291bnRpbmcpCiAgICBjb25zdCBwcm9qVmFsID0ge307CiAgICBTSlBJX0RBVEEuZm9yRWFjaChyID0+IHsgcHJvalZhbFtyLnNoaXBfdG9dID0gci5zdG9yYWdlX3ZhbHVlOyB9KTsKICAgIGNvbnN0IHRvdGFsU3RvcmFnZVZhbCA9IE9iamVjdC52YWx1ZXMocHJvalZhbCkucmVkdWNlKChhLGIpID0+IGErYiwgMCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2pwaUtwaVN0b3JhZ2VWYWwnKS50ZXh0Q29udGVudCA9IGZtdCQodG90YWxTdG9yYWdlVmFsKTsKICAgIGNvbnN0IHRvdGFsUXR5ID0gU0pQSV9EQVRBLnJlZHVjZSgoYSxyKSA9PiBhICsgKHIucXVhbnRpdHl8fDApLCAwKTsKICAgIGNvbnN0IHRvdGFsU2hpcHBlZCA9IFNKUElfREFUQS5yZWR1Y2UoKGEscikgPT4gYSArIChyLnF0eV9zaGlwcGVkfHwwKSwgMCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2pwaUtwaVVuc2hpcHBlZCcpLnRleHRDb250ZW50ICA9IGZtdE4odG90YWxRdHkgLSB0b3RhbFNoaXBwZWQpICsgJyByZW1haW5pbmcnOwogICAgcmVuZGVyU2pwaSgpOwogIH0gY2F0Y2goZSkgewogICAgY29uc29sZS5lcnJvcignU0pQSSBsb2FkIGZhaWxlZCcsIGUpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NqcGlCYWRnZScpLnRleHRDb250ZW50ID0gJ+KAlCc7CiAgfQp9CgpmdW5jdGlvbiByZXNldFNqcGlGaWx0ZXJzKCkgewogIG1zUmVzZXQoJ21zLXNqcGljb29yZCcpOyBtc1Jlc2V0KCdtcy1zanBpc3RhdHVzJyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NqcGlTZWFyY2gnKS52YWx1ZSA9ICcnOwogIHJlbmRlclNqcGkoKTsKfQoKZnVuY3Rpb24gcmVuZGVyU2pwaSgpIHsKICBjb25zdCBjb29yZCAgPSBtc0dldFNlbGVjdGVkKCdtcy1zanBpY29vcmQnKTsKICBjb25zdCBzdGF0dXMgPSBtc0dldFNlbGVjdGVkKCdtcy1zanBpc3RhdHVzJyk7CiAgY29uc3QgcSAgICAgID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NqcGlTZWFyY2gnKS52YWx1ZS50b0xvd2VyQ2FzZSgpOwogIGxldCByb3dzID0gU0pQSV9EQVRBOwogIGlmIChjb29yZC5sZW5ndGgpICByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKHN0YXR1cy5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHN0YXR1cy5pbmNsdWRlcyhyLmxpbmVfc3RhdHVzKSk7CiAgaWYgKHEpICAgICAgICAgICAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+CiAgICAoci5wcm9qZWN0fHwnJykudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fAogICAgKHIuc2hpcF90b3x8JycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwKICAgIChyLnBhcnRfbm98fCcnKS50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICAoci5jb29yZGluYXRvcnx8JycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwKICAgIChyLnNzX29yZGVyfHwnJykudG9TdHJpbmcoKS5pbmNsdWRlcyhxKSB8fAogICAgKHIuY3VzdG9tZXJfbmFtZXx8JycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwoKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2pwaU1ldGEnKS50ZXh0Q29udGVudCA9IHJvd3MubGVuZ3RoLnRvTG9jYWxlU3RyaW5nKCkgKyAnIGxpbmVzJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2pwaUJvZHknKS5pbm5lckhUTUwgPSByb3dzLm1hcChyID0+IGAKICAgIDx0ciBzdHlsZT0iYmFja2dyb3VuZDojZmZmYWY1OyI+CiAgICAgIDx0ZCBjbGFzcz0id3JhcCI+PHN0cm9uZz4ke3IucHJvamVjdH08L3N0cm9uZz48L3RkPgogICAgICA8dGQgc3R5bGU9ImZvbnQtc2l6ZToxMnB4O2NvbG9yOiM3QjNGMDAiPiR7ci5zdG9yYWdlX3N0YXR1c308L3RkPgogICAgICA8dGQ+JHtyLmNvb3JkaW5hdG9yfHwn4oCUJ308L3RkPgogICAgICA8dGQ+PHN0cm9uZz4ke3Iuc3Nfb3JkZXJ9PC9zdHJvbmc+PC90ZD4KICAgICAgPHRkIGNsYXNzPSJ3cmFwIj4ke3IuY3VzdG9tZXJfbmFtZXx8J+KAlCd9PC90ZD4KICAgICAgPHRkPiR7ci5wYXJ0X25vfTwvdGQ+CiAgICAgIDx0ZD4ke3IucmV2aXNpb258fCfigJQnfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4oci5xdWFudGl0eSl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10TihyLnF0eV9zaGlwcGVkKX08L3RkPgogICAgICA8dGQ+PHNwYW4gY2xhc3M9InRhZyAke3IubGluZV9zdGF0dXM9PT0nT3Blbic/J3RhZy1yZWQnOnIubGluZV9zdGF0dXM9PT0nQWxsb2NhdGVkJz8ndGFnLWFtYmVyJzondGFnLWdyZWVuJ30iPiR7ci5saW5lX3N0YXR1c308L3NwYW4+PC90ZD4KICAgICAgPHRkIHN0eWxlPSJmb250LXNpemU6MTJweCI+JHtyLm9yZGVyX3N0YXR1c3x8J+KAlCd9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10JChyLnN0b3JhZ2VfdmFsdWUpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdEYoci5hdmdfYWdlKX08L3RkPgogICAgICA8dGQ+JHtyLnNoaXBfZGF0ZX08L3RkPgogICAgPC90cj5gKS5qb2luKCcnKSB8fCAnPHRyPjx0ZCBjb2xzcGFuPSIxMiIgc3R5bGU9InRleHQtYWxpZ246Y2VudGVyO3BhZGRpbmc6MzJweDtjb2xvcjp2YXIoLS1tdXRlZCkiPk5vIGluY29tcGxldGUgc3RvcmFnZSBqb2JzIGZvdW5kLiBVcGxvYWQgdGhlIEJhY2tvcmRlciBSZXBvcnQgQ1NWIHRvIHBvcHVsYXRlIHRoaXMgdGFiLjwvdGQ+PC90cj4nOwoKICBpZiAoIXJvd3MubGVuZ3RoKSB7IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzanBpRm9vdCcpLmlubmVySFRNTCA9ICcnOyByZXR1cm47IH0KICBjb25zdCB0b3QgPSByb3dzLnJlZHVjZSgoYSxyKSA9PiAoewogICAgcXR5OiBhLnF0eSArIChyLnF1YW50aXR5fHwwKSwKICAgIHF0eV9zaGlwcGVkOiBhLnF0eV9zaGlwcGVkICsgKHIucXR5X3NoaXBwZWR8fDApCiAgfSksIHtxdHk6MCwgcXR5X3NoaXBwZWQ6MH0pOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzanBpRm9vdCcpLmlubmVySFRNTCA9IGA8dHI+CiAgICA8dGQgY29sc3Bhbj0iNyI+VE9UQUwg4oCUICR7cm93cy5sZW5ndGgudG9Mb2NhbGVTdHJpbmcoKX0gbGluZXMgwrcgJHtuZXcgU2V0KHJvd3MubWFwKHI9PnIuc2hpcF90bykpLnNpemV9IHByb2plY3RzPC90ZD4KICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4odG90LnF0eSl9PC90ZD4KICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4odG90LnF0eV9zaGlwcGVkKX08L3RkPgogICAgPHRkIGNvbHNwYW49IjUiPjwvdGQ+CiAgPC90cj5gOwp9CgpmdW5jdGlvbiBleHBvcnRTanBpQ1NWKCkgewogIGNvbnN0IGNvb3JkICA9IG1zR2V0U2VsZWN0ZWQoJ21zLXNqcGljb29yZCcpOwogIGNvbnN0IHN0YXR1cyA9IG1zR2V0U2VsZWN0ZWQoJ21zLXNqcGlzdGF0dXMnKTsKICBjb25zdCBxICAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2pwaVNlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBTSlBJX0RBVEE7CiAgaWYgKGNvb3JkLmxlbmd0aCkgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGNvb3JkLmluY2x1ZGVzKHIuY29vcmRpbmF0b3IpKTsKICBpZiAoc3RhdHVzLmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gc3RhdHVzLmluY2x1ZGVzKHIubGluZV9zdGF0dXMpKTsKICBpZiAocSkgICAgICAgICAgICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gci5wcm9qZWN0LnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwgci5wYXJ0X25vLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwogIGNvbnN0IGhlYWRlcnMgPSBbJ1Byb2plY3QgKFN0b3JhZ2UpJywnU3RvcmFnZSBTdGF0dXMnLCdDb29yZGluYXRvcicsJ1NTIE9yZGVyICMnLCdDdXN0b21lciBOYW1lJywnUGFydCBOby4nLCdSZXZpc2lvbicsJ1F1YW50aXR5JywnUXR5IFNoaXBwZWQnLCdMaW5lIFN0YXR1cycsJ09yZGVyIFN0YXR1cycsJ1N0b3JhZ2UgVmFsdWUnLCdBdmcgQWdlJywnTWF4IEFnZScsJ1NoaXAgRGF0ZSddOwogIGNvbnN0IGNzdlJvd3MgPSByb3dzLm1hcChyID0+IFtyLnByb2plY3QsIHIuc3RvcmFnZV9zdGF0dXMsIHIuY29vcmRpbmF0b3IsIHIuc3Nfb3JkZXIsIHIuY3VzdG9tZXJfbmFtZSwgci5wYXJ0X25vLCByLnJldmlzaW9uLCByLnF1YW50aXR5LCByLnF0eV9zaGlwcGVkLCByLmxpbmVfc3RhdHVzLCByLm9yZGVyX3N0YXR1cywgci5zdG9yYWdlX3ZhbHVlLCByLmF2Z19hZ2UsIHIubWF4X2FnZSwgci5zaGlwX2RhdGVdKTsKICBkb3dubG9hZENTVignU3RvcmFnZV9Kb2JzX1BhcnRzX0luY29tcGxldGUuY3N2JywgdG9DU1YoaGVhZGVycywgY3N2Um93cykpOwp9CgovLyDilIDilIAgSW5pdCDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKbG9hZERhdGEoKTsKPC9zY3JpcHQ+CjwvYm9keT4KPC9odG1sPg==").decode('utf-8')

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    has_data = PROCESSED_PATH.exists()
    meta = json.loads(META_PATH.read_text()) if META_PATH.exists() else {}
    html = INDEX_HTML
    if has_data:
        status = '<div class="status-banner"><strong>&#10003; Data loaded &mdash; {uploaded}</strong>{filename} &middot; {rows:,} rows{sdrop}</div>'.format(
            uploaded=meta.get('uploaded',''),
            filename=meta.get('warehouse_filename',''),
            rows=meta.get('rows',0),
            sdrop=(' &middot; ' + str(meta.get('sdrop_items','')) + ' S-Drop items') if meta.get('sdrop_items') else ''
        )
        link = '<a class="view-link" href="/dashboard">&rarr; View current dashboard</a>'
    else:
        status = ''
        link = ''
    html = html.replace('STATUS_BLOCK', status).replace('LINK_BLOCK', link)
    return Response(html, mimetype='text/html')

# Processing state for async upload
_processing_state = {'status': 'idle', 'error': None}


@app.route('/upload', methods=['POST'])
def upload():
    import threading
    warehouse = request.files.get('warehouse')
    notes     = request.files.get('notes')
    if not warehouse or not notes:
        return jsonify({'error': 'Both files are required'}), 400
    billing = request.files.get('billing')

    # Read all file bytes in the request handler (must happen before returning)
    backorder   = request.files.get('backorder')
    wh_bytes    = warehouse.read()
    notes_bytes = notes.read()
    bill_bytes  = billing.read() if billing else None
    bo_bytes    = backorder.read() if backorder else None
    wh_name     = warehouse.filename
    nt_name     = notes.filename

    def _save_and_process():
        global _processing_state
        _processing_state = {'status': 'processing', 'error': None}
        try:
            wh_path    = DATA_DIR / 'warehouse_temp.xlsx'
            notes_path = DATA_DIR / 'notes_temp.csv'
            wh_path.write_bytes(wh_bytes)
            notes_path.write_bytes(notes_bytes)
            bl_path = None
            if bill_bytes:
                BILLING_PATH.write_bytes(bill_bytes)
                bl_path = BILLING_PATH
            # Process backorder cross-reference if file was uploaded
            if bo_bytes:
                bo_path = DATA_DIR / 'backorder_temp.csv'
                bo_path.write_bytes(bo_bytes)
                try:
                    # Load current dispositions to filter by Storage Charged
                    disp = json.loads(DISPOSITIONS_PATH.read_text()) if DISPOSITIONS_PATH.exists() else {}
                    bo_items = process_backorder(wh_path, bo_path, disp)
                    BACKORDER_PATH.write_text(json.dumps(bo_items))
                except Exception as be:
                    print(f"Backorder processing error: {be}")
            data = process_files(wh_path, notes_path, bl_path)
            PROCESSED_PATH.write_text(json.dumps(data))
            META_PATH.write_text(json.dumps({
                'warehouse_filename': wh_name,
                'notes_filename':     nt_name,
                'rows':               data['kpis']['total_containers'],
                'sdrop_items':        data['sdrop']['kpis']['total_items'],
                'uploaded':           data['uploaded'],
            }))
            _processing_state = {'status': 'done', 'error': None}
        except Exception as e:
            import traceback
            _processing_state = {'status': 'error', 'error': str(e), 'detail': traceback.format_exc()}

    threading.Thread(target=_save_and_process, daemon=True).start()
    return jsonify({'success': True, 'async': True})

@app.route('/api/upload_status')
def upload_status():
    return jsonify(_processing_state)

@app.route('/dashboard')
def dashboard():
    if not PROCESSED_PATH.exists():
        return redirect('/')
    return Response(DASHBOARD_HTML, mimetype='text/html')

@app.route('/api/set_disposition', methods=['POST'])
def api_set_disposition():
    data = request.get_json()
    order = data.get('order')
    disposition = data.get('disposition', '')
    if not order:
        return jsonify({'error': 'order required'}), 400
    dispositions = {}
    if DISPOSITIONS_PATH.exists():
        dispositions = json.loads(DISPOSITIONS_PATH.read_text())
    if disposition:
        dispositions[order] = disposition
    elif order in dispositions:
        del dispositions[order]
    DISPOSITIONS_PATH.write_text(json.dumps(dispositions))
    return jsonify({'success': True})

@app.route('/api/backorder')
def api_backorder():
    if not BACKORDER_PATH.exists():
        return '[]', 200, {'Content-Type': 'application/json'}
    return BACKORDER_PATH.read_text(), 200, {'Content-Type': 'application/json'}

@app.route('/api/backorder_debug')
def api_backorder_debug():
    result = {
        'backorder_file_exists': BACKORDER_PATH.exists(),
        'backorder_count': 0,
        'sample_projects': [],
        'processed_file_exists': PROCESSED_PATH.exists(),
    }
    if BACKORDER_PATH.exists():
        data = json.loads(BACKORDER_PATH.read_text())
        result['backorder_count'] = len(data)
        projects = list(set(r.get('ship_to','') for r in data))[:10]
        result['sample_projects'] = projects
    return jsonify(result)

@app.route('/api/billing')
def api_billing():
    if not BILLING_PATH.exists():
        return '{}', 200, {'Content-Type': 'application/json'}
    data = BILLING_PATH.read_text()
    parsed = json.loads(data)
    print(f"[DEBUG] /api/billing returning {len(parsed)} orders: {list(parsed.keys())[:5]}")
    return data, 200, {'Content-Type': 'application/json'}

@app.route('/api/billing_debug')
def api_billing_debug():
    """Debug endpoint to check billing data"""
    if not BILLING_PATH.exists():
        return jsonify({'error': 'No billing file uploaded', 'path': str(BILLING_PATH)})
    data = json.loads(BILLING_PATH.read_text())
    # Also check a few offsite orders to see if they match
    offsite_orders = []
    if PROCESSED_PATH.exists():
        processed = json.loads(PROCESSED_PATH.read_text())
        offsite_orders = [r['order'] for r in processed.get('offsite', {}).get('items', [])[:10]]
    matches = [o for o in offsite_orders if o in data]
    return jsonify({
        'billing_count': len(data),
        'billing_sample_keys': list(data.keys())[:10],
        'offsite_sample_orders': offsite_orders[:10],
        'matches': matches
    })

@app.route('/api/dispositions')
def api_dispositions():
    if not DISPOSITIONS_PATH.exists():
        return '{}', 200, {'Content-Type': 'application/json'}
    return DISPOSITIONS_PATH.read_text(), 200, {'Content-Type': 'application/json'}

@app.route('/api/new_storage')
def api_new_storage():
    if not NEW_STORAGE_PATH.exists():
        return '[]', 200, {'Content-Type': 'application/json'}
    return NEW_STORAGE_PATH.read_text(), 200, {'Content-Type': 'application/json'}

@app.route('/api/billing_data')
def api_billing_data():
    if not PROCESSED_PATH.exists():
        return '{}', 200, {'Content-Type': 'application/json'}
    data = json.loads(PROCESSED_PATH.read_text())
    return json.dumps(data.get('billing', {})), 200, {'Content-Type': 'application/json'}

@app.route('/api/data')
def api_data():
    if not PROCESSED_PATH.exists():
        return jsonify({'error': 'No data — please upload files first'}), 404
    return PROCESSED_PATH.read_text(), 200, {'Content-Type': 'application/json'}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

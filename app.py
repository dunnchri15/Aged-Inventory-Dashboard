import os, json
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
['wz','nz','bz'].forEach(function(id){
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
  var bf=document.getElementById('bf').files[0];
  if(bf) fd.append('billing', bf);
  var bf=document.getElementById('bf').files[0];
  if(bf) fd.append('billing',bf);
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


DASHBOARD_HTML = _b64.b64decode("PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CiAgPG1ldGEgY2hhcnNldD0iVVRGLTgiPgogIDxtZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsIGluaXRpYWwtc2NhbGU9MS4wIj4KICA8dGl0bGU+QWdlZCBJbnZlbnRvcnkgRGFzaGJvYXJkPC90aXRsZT4KICA8c3R5bGU+CiAgICAqLCAqOjpiZWZvcmUsICo6OmFmdGVyIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgbWFyZ2luOiAwOyBwYWRkaW5nOiAwOyB9CgogICAgOnJvb3QgewogICAgICAtLW5hdnk6ICMxRjM4NjQ7IC0tYmx1ZTogIzJFNzVCNjsgLS1saWdodC1ibHVlOiAjRDZFNEYwOyAtLXNlY3Rpb24tYmc6ICNFQkYzRkI7CiAgICAgIC0tZ3JlZW46ICMzNzU2MjM7IC0tZ3JlZW4tbGlnaHQ6ICNFMkVGREE7IC0tcmVkOiAjQzAwMDAwOyAtLXJlZC1saWdodDogI0ZGQ0NDQzsKICAgICAgLS1hbWJlcjogI0M1NUExMTsgLS1hbWJlci1saWdodDogI0ZGRjJDQzsKICAgICAgLS10ZXh0OiAjMWUyOTNiOyAtLW11dGVkOiAjNjQ3NDhiOyAtLWJvcmRlcjogI2UyZThmMDsgLS13aGl0ZTogI2ZmZmZmZjsKICAgICAgLS1zdXJmYWNlOiAjZjhmYWZjOwogICAgfQoKICAgIGJvZHkgeyBmb250LWZhbWlseTogJ1NlZ29lIFVJJywgQXJpYWwsIHNhbnMtc2VyaWY7IGJhY2tncm91bmQ6ICNmMGY0Zjg7CiAgICAgICAgICAgY29sb3I6IHZhcigtLXRleHQpOyBmb250LXNpemU6IDEzcHg7IH0KCiAgICAvKiDilIDilIAgSGVhZGVyIOKUgOKUgCAqLwogICAgLmhlYWRlciB7IGJhY2tncm91bmQ6IHZhcigtLW5hdnkpOyBjb2xvcjogI2ZmZjsgcGFkZGluZzogMCAzMnB4OwogICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsKICAgICAgICAgICAgICBoZWlnaHQ6IDU2cHg7IHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1pbmRleDogMTAwOwogICAgICAgICAgICAgIGJveC1zaGFkb3c6IDAgMnB4IDhweCByZ2JhKDAsMCwwLC4yKTsgfQogICAgLmhlYWRlci1sZWZ0IHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMnB4OyB9CiAgICAuaGVhZGVyIGgxIHsgZm9udC1zaXplOiAxN3B4OyBmb250LXdlaWdodDogNzAwOyB9CiAgICAuaGVhZGVyLXN1YiB7IGZvbnQtc2l6ZTogMTFweDsgb3BhY2l0eTogLjc7IG1hcmdpbi10b3A6IDFweDsgfQogICAgLnVwbG9hZC1saW5rIHsgYmFja2dyb3VuZDogcmdiYSgyNTUsMjU1LDI1NSwuMTUpOyBjb2xvcjogI2ZmZjsgdGV4dC1kZWNvcmF0aW9uOiBub25lOwogICAgICAgICAgICAgICAgICAgcGFkZGluZzogNnB4IDE0cHg7IGJvcmRlci1yYWRpdXM6IDZweDsgZm9udC1zaXplOiAxMnB4OyBmb250LXdlaWdodDogNjAwOwogICAgICAgICAgICAgICAgICAgdHJhbnNpdGlvbjogYmFja2dyb3VuZCAuMnM7IGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoMjU1LDI1NSwyNTUsLjI1KTsgfQogICAgLnVwbG9hZC1saW5rOmhvdmVyIHsgYmFja2dyb3VuZDogcmdiYSgyNTUsMjU1LDI1NSwuMjUpOyB9CgogICAgLyog4pSA4pSAIFRhYiBuYXYg4pSA4pSAICovCiAgICAudGFiLW5hdiB7IGJhY2tncm91bmQ6IHZhcigtLW5hdnkpOyBwYWRkaW5nOiAwIDMycHg7CiAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGdhcDogNHB4OyBib3JkZXItdG9wOiAxcHggc29saWQgcmdiYSgyNTUsMjU1LDI1NSwuMSk7IH0KICAgIC50YWItYnRuIHsgcGFkZGluZzogMTBweCAyMHB4OyBmb250LXNpemU6IDEycHg7IGZvbnQtd2VpZ2h0OiA2MDA7IGNvbG9yOiByZ2JhKDI1NSwyNTUsMjU1LC42KTsKICAgICAgICAgICAgICAgYmFja2dyb3VuZDogdHJhbnNwYXJlbnQ7IGJvcmRlcjogbm9uZTsgY3Vyc29yOiBwb2ludGVyOyBib3JkZXItYm90dG9tOiAzcHggc29saWQgdHJhbnNwYXJlbnQ7CiAgICAgICAgICAgICAgIHRyYW5zaXRpb246IGFsbCAuMnM7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KICAgIC50YWItYnRuOmhvdmVyIHsgY29sb3I6IHJnYmEoMjU1LDI1NSwyNTUsLjkpOyB9CiAgICAudGFiLWJ0bi5hY3RpdmUgeyBjb2xvcjogI2ZmZjsgYm9yZGVyLWJvdHRvbS1jb2xvcjogdmFyKC0tYW1iZXItbGlnaHQpOyB9CiAgICAudGFiLWJ0biAudGFiLWJhZGdlIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyBiYWNrZ3JvdW5kOiB2YXIoLS1yZWQpOyBjb2xvcjogI2ZmZjsKICAgICAgICAgICAgICAgICAgICAgICAgICBib3JkZXItcmFkaXVzOiAxMHB4OyBwYWRkaW5nOiAxcHggN3B4OyBmb250LXNpemU6IDEwcHg7CiAgICAgICAgICAgICAgICAgICAgICAgICAgbWFyZ2luLWxlZnQ6IDZweDsgZm9udC13ZWlnaHQ6IDcwMDsgfQoKICAgIC8qIOKUgOKUgCBGaWx0ZXJzIGJhciDilIDilIAgKi8KICAgIC5maWx0ZXJzLWJhciB7IGJhY2tncm91bmQ6ICMxNjJkNTQ7IHBhZGRpbmc6IDEwcHggMzJweDsKICAgICAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGdhcDogMTJweDsgYWxpZ24taXRlbXM6IGZsZXgtZW5kOyBmbGV4LXdyYXA6IHdyYXA7CiAgICAgICAgICAgICAgICAgICBib3JkZXItYm90dG9tOiAxcHggc29saWQgcmdiYSgyNTUsMjU1LDI1NSwuMDgpOyB9CiAgICAvKiBzZHJvcCBmaWx0ZXJzIHNob3duL2hpZGRlbiBieSBKUyAqLwogICAgLmZpbHRlci1ncm91cCB7IGRpc3BsYXk6IGZsZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGdhcDogNHB4OyB9CiAgICAuZmlsdGVyLWdyb3VwIGxhYmVsIHsgZm9udC1zaXplOiAxMHB4OyBjb2xvcjogcmdiYSgyNTUsMjU1LDI1NSwuNjUpOyBmb250LXdlaWdodDogNjAwOwogICAgICAgICAgICAgICAgICAgICAgICAgIHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGxldHRlci1zcGFjaW5nOiAuMDVlbTsgfQogICAgLmZpbHRlci1ncm91cCBzZWxlY3QgeyBiYWNrZ3JvdW5kOiByZ2JhKDI1NSwyNTUsMjU1LC4xMik7IGNvbG9yOiAjZmZmOyBib3JkZXI6IDFweCBzb2xpZCByZ2JhKDI1NSwyNTUsMjU1LC4yNSk7CiAgICAgICAgICAgICAgICAgICAgICAgICAgIGJvcmRlci1yYWRpdXM6IDZweDsgcGFkZGluZzogNnB4IDI4cHggNnB4IDEwcHg7IGZvbnQtc2l6ZTogMTJweDsgZm9udC13ZWlnaHQ6IDUwMDsKICAgICAgICAgICAgICAgICAgICAgICAgICAgY3Vyc29yOiBwb2ludGVyOyBhcHBlYXJhbmNlOiBub25lOyBtaW4td2lkdGg6IDE4MHB4OyBtYXgtd2lkdGg6IDI2MHB4OwogICAgICAgICAgICAgICAgICAgICAgICAgICBiYWNrZ3JvdW5kLWltYWdlOiB1cmwoImRhdGE6aW1hZ2Uvc3ZnK3htbCwlM0NzdmcgeG1sbnM9J2h0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnJyB3aWR0aD0nMTInIGhlaWdodD0nMTInIHZpZXdCb3g9JzAgMCAyNCAyNCcgZmlsbD0nbm9uZScgc3Ryb2tlPSd3aGl0ZScgc3Ryb2tlLXdpZHRoPScyJyUzRSUzQ3BvbHlsaW5lIHBvaW50cz0nNiA5IDEyIDE1IDE4IDknLyUzRSUzQy9zdmclM0UiKTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgYmFja2dyb3VuZC1yZXBlYXQ6IG5vLXJlcGVhdDsgYmFja2dyb3VuZC1wb3NpdGlvbjogcmlnaHQgOHB4IGNlbnRlcjsgfQogICAgLmZpbHRlci1ncm91cCBzZWxlY3Qgb3B0aW9uIHsgYmFja2dyb3VuZDogdmFyKC0tbmF2eSk7IH0KICAgIC5maWx0ZXItZ3JvdXAgc2VsZWN0OmZvY3VzIHsgb3V0bGluZTogbm9uZTsgYm9yZGVyLWNvbG9yOiB2YXIoLS1saWdodC1ibHVlKTsgfQogICAgLnJlc2V0LWJ0biB7IGJhY2tncm91bmQ6IHJnYmEoMjU1LDI1NSwyNTUsLjEpOyBjb2xvcjogcmdiYSgyNTUsMjU1LDI1NSwuOCk7IGJvcmRlcjogMXB4IHNvbGlkIHJnYmEoMjU1LDI1NSwyNTUsLjIpOwogICAgICAgICAgICAgICAgIGJvcmRlci1yYWRpdXM6IDZweDsgcGFkZGluZzogNnB4IDE0cHg7IGZvbnQtc2l6ZTogMTJweDsgY3Vyc29yOiBwb2ludGVyOwogICAgICAgICAgICAgICAgIHRyYW5zaXRpb246IGFsbCAuMnM7IGFsaWduLXNlbGY6IGZsZXgtZW5kOyB9CiAgICAucmVzZXQtYnRuOmhvdmVyIHsgYmFja2dyb3VuZDogcmdiYSgyNTUsMjU1LDI1NSwuMik7IGNvbG9yOiAjZmZmOyB9CgogICAgLyog4pSA4pSAIE1haW4gbGF5b3V0IOKUgOKUgCAqLwogICAgLm1haW4geyBwYWRkaW5nOiAyNHB4IDMycHg7IG1heC13aWR0aDogMTYwMHB4OyBtYXJnaW46IDAgYXV0bzsgfQogICAgLnRhYi1wYW5lbCB7IGRpc3BsYXk6IG5vbmU7IH0KICAgIC50YWItcGFuZWwuYWN0aXZlIHsgZGlzcGxheTogYmxvY2s7IH0KCiAgICAvKiDilIDilIAgS1BJIGNhcmRzIOKUgOKUgCAqLwogICAgLmtwaS1ncmlkIHsgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maXQsIG1pbm1heCgxNzBweCwgMWZyKSk7CiAgICAgICAgICAgICAgICBnYXA6IDE2cHg7IG1hcmdpbi1ib3R0b206IDI0cHg7IH0KICAgIC5rcGktY2FyZCB7IGJhY2tncm91bmQ6ICNmZmY7IGJvcmRlci1yYWRpdXM6IDEwcHg7IHBhZGRpbmc6IDIwcHggMjJweDsKICAgICAgICAgICAgICAgIGJveC1zaGFkb3c6IDAgMXB4IDRweCByZ2JhKDAsMCwwLC4wOCk7IGJvcmRlci10b3A6IDNweCBzb2xpZCB2YXIoLS1ibHVlKTsgfQogICAgLmtwaS1jYXJkLmFjY2VudC1yZWQgICB7IGJvcmRlci10b3AtY29sb3I6IHZhcigtLXJlZCk7IH0KICAgIC5rcGktY2FyZC5hY2NlbnQtYW1iZXIgeyBib3JkZXItdG9wLWNvbG9yOiB2YXIoLS1hbWJlcik7IH0KICAgIC5rcGktY2FyZC5hY2NlbnQtZ3JlZW4geyBib3JkZXItdG9wLWNvbG9yOiB2YXIoLS1ncmVlbik7IH0KICAgIC5rcGktbGFiZWwgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IGZvbnQtd2VpZ2h0OiA2MDA7IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7CiAgICAgICAgICAgICAgICAgbGV0dGVyLXNwYWNpbmc6IC4wNGVtOyBtYXJnaW4tYm90dG9tOiA4cHg7IH0KICAgIC5rcGktdmFsdWUgeyBmb250LXNpemU6IDI2cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS1uYXZ5KTsgbGluZS1oZWlnaHQ6IDE7IH0KICAgIC5rcGktc3ViICAgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IG1hcmdpbi10b3A6IDVweDsgfQoKICAgIC8qIOKUgOKUgCBTZWN0aW9uIGNhcmRzIOKUgOKUgCAqLwogICAgLnNlY3Rpb24tY2FyZCB7IGJhY2tncm91bmQ6ICNmZmY7IGJvcmRlci1yYWRpdXM6IDEwcHg7IGJveC1zaGFkb3c6IDAgMXB4IDRweCByZ2JhKDAsMCwwLC4wOCk7CiAgICAgICAgICAgICAgICAgICAgbWFyZ2luLWJvdHRvbTogMjRweDsgb3ZlcmZsb3c6IGhpZGRlbjsgfQogICAgLnNlY3Rpb24taGVhZGVyIHsgcGFkZGluZzogMTZweCAyMHB4OyBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKICAgICAgICAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgfQogICAgLnNlY3Rpb24tdGl0bGUgeyBmb250LXNpemU6IDE0cHg7IGZvbnQtd2VpZ2h0OiA3MDA7IGNvbG9yOiB2YXIoLS1uYXZ5KTsgfQogICAgLnNlY3Rpb24tbWV0YSAgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IH0KCiAgICAvKiDilIDilIAgU2VhcmNoIC8gZmlsdGVyIGlucHV0IOKUgOKUgCAqLwogICAgLnNlYXJjaC1ib3ggeyBwYWRkaW5nOiAxMnB4IDIwcHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBkaXNwbGF5OiBmbGV4OyBnYXA6IDEwcHg7IH0KICAgIC5zZWFyY2gtYm94IGlucHV0LCAuc2VhcmNoLWJveCBzZWxlY3QgewogICAgICBwYWRkaW5nOiA3cHggMTJweDsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogNnB4OwogICAgICBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS10ZXh0KTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZSk7IH0KICAgIC5zZWFyY2gtYm94IGlucHV0IHsgZmxleDogMTsgbWF4LXdpZHRoOiAzMjBweDsgfQogICAgLnNlYXJjaC1ib3ggaW5wdXQ6Zm9jdXMsIC5zZWFyY2gtYm94IHNlbGVjdDpmb2N1cyB7IG91dGxpbmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tYmx1ZSk7IH0KCiAgICAvKiDilIDilIAgVGFibGVzIOKUgOKUgCAqLwogICAgLnRhYmxlLXdyYXAgeyBvdmVyZmxvdy14OiBhdXRvOyB9CiAgICB0YWJsZSB7IHdpZHRoOiAxMDAlOyBib3JkZXItY29sbGFwc2U6IGNvbGxhcHNlOyBmb250LXNpemU6IDEycHg7IH0KICAgIHRoZWFkIHRoIHsgYmFja2dyb3VuZDogdmFyKC0tbmF2eSk7IGNvbG9yOiAjZmZmOyBwYWRkaW5nOiA5cHggMTJweDsgdGV4dC1hbGlnbjogbGVmdDsKICAgICAgICAgICAgICAgZm9udC13ZWlnaHQ6IDYwMDsgZm9udC1zaXplOiAxMXB4OyB3aGl0ZS1zcGFjZTogbm93cmFwOwogICAgICAgICAgICAgICBjdXJzb3I6IHBvaW50ZXI7IHVzZXItc2VsZWN0OiBub25lOyB9CiAgICB0aGVhZCB0aDpob3ZlciB7IGJhY2tncm91bmQ6IHZhcigtLWJsdWUpOyB9CiAgICB0aGVhZCB0aC5zb3J0LWFzYzo6YWZ0ZXIgIHsgY29udGVudDogJyDihpEnOyB9CiAgICB0aGVhZCB0aC5zb3J0LWRlc2M6OmFmdGVyIHsgY29udGVudDogJyDihpMnOyB9CiAgICB0Ym9keSB0ciB7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB0cmFuc2l0aW9uOiBiYWNrZ3JvdW5kIC4xczsgfQogICAgdGJvZHkgdHI6aG92ZXIgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zZWN0aW9uLWJnKTsgfQogICAgdGJvZHkgdHI6bnRoLWNoaWxkKGV2ZW4pIHsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZSk7IH0KICAgIHRib2R5IHRyOm50aC1jaGlsZChldmVuKTpob3ZlciB7IGJhY2tncm91bmQ6IHZhcigtLXNlY3Rpb24tYmcpOyB9CiAgICB0ZCB7IHBhZGRpbmc6IDhweCAxMnB4OyB3aGl0ZS1zcGFjZTogbm93cmFwOyB9CiAgICB0ZC53cmFwIHsgd2hpdGUtc3BhY2U6IG5vcm1hbDsgbWF4LXdpZHRoOiAyNDBweDsgfQogICAgdGZvb3QgdGQgeyBmb250LXdlaWdodDogNzAwOyBiYWNrZ3JvdW5kOiB2YXIoLS1saWdodC1ibHVlKTsgY29sb3I6IHZhcigtLW5hdnkpOwogICAgICAgICAgICAgICBwYWRkaW5nOiA5cHggMTJweDsgYm9yZGVyLXRvcDogMnB4IHNvbGlkIHZhcigtLWJsdWUpOyB9CgogICAgLyog4pSA4pSAIENlbGwgaGVscGVycyDilIDilIAgKi8KICAgIC5udW0geyB0ZXh0LWFsaWduOiByaWdodDsgZm9udC12YXJpYW50LW51bWVyaWM6IHRhYnVsYXItbnVtczsgfQogICAgLnRhZyB7IGRpc3BsYXk6IGlubGluZS1ibG9jazsgcGFkZGluZzogMnB4IDhweDsgYm9yZGVyLXJhZGl1czogNHB4OyBmb250LXNpemU6IDEwcHg7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KICAgIC50YWctcmVkICAgIHsgYmFja2dyb3VuZDogdmFyKC0tcmVkLWxpZ2h0KTsgICBjb2xvcjogdmFyKC0tcmVkKTsgfQogICAgLnRhZy1hbWJlciAgeyBiYWNrZ3JvdW5kOiB2YXIoLS1hbWJlci1saWdodCk7IGNvbG9yOiB2YXIoLS1hbWJlcik7IH0KICAgIC50YWctZ3JlZW4gIHsgYmFja2dyb3VuZDogdmFyKC0tZ3JlZW4tbGlnaHQpOyBjb2xvcjogdmFyKC0tZ3JlZW4pOyB9CiAgICAudGFnLWJsdWUgICB7IGJhY2tncm91bmQ6IHZhcigtLWxpZ2h0LWJsdWUpOyAgY29sb3I6IHZhcigtLW5hdnkpOyB9CiAgICAuYmFyLWNlbGwgICB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyBtaW4td2lkdGg6IDEyMHB4OyB9CiAgICAuYmFyICAgICAgICB7IGZsZXg6IDE7IGhlaWdodDogNnB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiAzcHg7IG92ZXJmbG93OiBoaWRkZW47IH0KICAgIC5iYXItZmlsbCAgIHsgaGVpZ2h0OiAxMDAlOyBib3JkZXItcmFkaXVzOiAzcHg7IGJhY2tncm91bmQ6IHZhcigtLWJsdWUpOyB0cmFuc2l0aW9uOiB3aWR0aCAuM3M7IH0KICAgIC5iYXItZmlsbC5yZWQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1yZWQpOyB9CgogICAgLyog4pSA4pSAIFR3by1jb2wgbGF5b3V0IOKUgOKUgCAqLwogICAgLnR3by1jb2wgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IGdhcDogMjRweDsgbWFyZ2luLWJvdHRvbTogMjRweDsgfQogICAgQG1lZGlhIChtYXgtd2lkdGg6IDkwMHB4KSB7IC50d28tY29sIHsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiAxZnI7IH0gfQoKICAgIC8qIOKUgOKUgCBTLURyb3Agc3BlY2lmaWMg4pSA4pSAICovCiAgICAuc2Ryb3AtbG9jLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpdCwgbWlubWF4KDE2MHB4LDFmcikpOwogICAgICAgICAgICAgICAgICAgICAgZ2FwOiAxMnB4OyBwYWRkaW5nOiAxNnB4IDIwcHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9CiAgICAuc2Ryb3AtbG9jLWNhcmQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czogOHB4OwogICAgICAgICAgICAgICAgICAgICAgcGFkZGluZzogMTRweCAxNnB4OyB9CiAgICAuc2Ryb3AtbG9jLW5hbWUgIHsgZm9udC1zaXplOiAxMXB4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tbmF2eSk7CiAgICAgICAgICAgICAgICAgICAgICAgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbWFyZ2luLWJvdHRvbTogNnB4OyB9CiAgICAuc2Ryb3AtbG9jLWl0ZW1zIHsgZm9udC1zaXplOiAyMHB4OyBmb250LXdlaWdodDogNzAwOyBjb2xvcjogdmFyKC0tbmF2eSk7IH0KICAgIC5zZHJvcC1sb2MtdmFsICAgeyBmb250LXNpemU6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IG1hcmdpbi10b3A6IDJweDsgfQoKICAgIC8qIOKUgOKUgCBMb2FkaW5nIG92ZXJsYXkg4pSA4pSAICovCiAgICAjbG9hZGluZyB7IHBvc2l0aW9uOiBmaXhlZDsgaW5zZXQ6IDA7IGJhY2tncm91bmQ6IHJnYmEoMzEsNTYsMTAwLC44NSk7CiAgICAgICAgICAgICAgIGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOwogICAgICAgICAgICAgICB6LWluZGV4OiA5OTk7IGNvbG9yOiAjZmZmOyBmb250LXNpemU6IDE1cHg7IGZvbnQtd2VpZ2h0OiA2MDA7IH0KICAgIC5zcGlubmVyIHsgd2lkdGg6IDM2cHg7IGhlaWdodDogMzZweDsgYm9yZGVyOiAzcHggc29saWQgcmdiYSgyNTUsMjU1LDI1NSwuMyk7CiAgICAgICAgICAgICAgIGJvcmRlci10b3AtY29sb3I6ICNmZmY7IGJvcmRlci1yYWRpdXM6IDUwJTsgYW5pbWF0aW9uOiBzcGluIC44cyBsaW5lYXIgaW5maW5pdGU7CiAgICAgICAgICAgICAgIG1hcmdpbi1yaWdodDogMTZweDsgfQogICAgQGtleWZyYW1lcyBzcGluIHsgdG8geyB0cmFuc2Zvcm06IHJvdGF0ZSgzNjBkZWcpOyB9IH0KICAKLyog4pSA4pSAIE11bHRpLXNlbGVjdCBkcm9wZG93biDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAgKi8KLm1zLXdyYXAgeyBwb3NpdGlvbjpyZWxhdGl2ZTsgZGlzcGxheTppbmxpbmUtYmxvY2s7IG1pbi13aWR0aDoxODBweDsgfQoubXMtYnRuICB7IHdpZHRoOjEwMCU7IHBhZGRpbmc6N3B4IDEwcHg7IGJhY2tncm91bmQ6IzFhMmY1NDsgY29sb3I6I2ZmZjsKICAgICAgICAgICBib3JkZXI6MXB4IHNvbGlkIHJnYmEoMjU1LDI1NSwyNTUsLjI1KTsgYm9yZGVyLXJhZGl1czo2cHg7CiAgICAgICAgICAgY3Vyc29yOnBvaW50ZXI7IGZvbnQtc2l6ZToxM3B4OyB0ZXh0LWFsaWduOmxlZnQ7CiAgICAgICAgICAgZGlzcGxheTpmbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGp1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuOyBnYXA6NnB4OyB9Ci5tcy1idG46YWZ0ZXIgeyBjb250ZW50Oifilr4nOyBmb250LXNpemU6MTFweDsgb3BhY2l0eTouNzsgZmxleC1zaHJpbms6MDsgfQoubXMtYnRuLm9wZW46YWZ0ZXIgeyBjb250ZW50OifilrQnOyB9Ci5tcy1sYWJlbCB7IHdoaXRlLXNwYWNlOm5vd3JhcDsgb3ZlcmZsb3c6aGlkZGVuOyB0ZXh0LW92ZXJmbG93OmVsbGlwc2lzOyB9Ci5tcy1wYW5lbCB7IGRpc3BsYXk6bm9uZTsgcG9zaXRpb246YWJzb2x1dGU7IHRvcDpjYWxjKDEwMCUgKyA0cHgpOyBsZWZ0OjA7CiAgICAgICAgICAgIG1pbi13aWR0aDoxMDAlOyBtYXgtaGVpZ2h0OjIyMHB4OyBvdmVyZmxvdy15OmF1dG87CiAgICAgICAgICAgIGJhY2tncm91bmQ6I2ZmZjsgYm9yZGVyOjFweCBzb2xpZCAjZGRlOyBib3JkZXItcmFkaXVzOjhweDsKICAgICAgICAgICAgYm94LXNoYWRvdzowIDZweCAyMHB4IHJnYmEoMCwwLDAsLjE1KTsgei1pbmRleDo5OTk7IHBhZGRpbmc6NnB4IDA7IH0KLm1zLXBhbmVsLm9wZW4geyBkaXNwbGF5OmJsb2NrOyB9Ci5tcy1pdGVtIHsgZGlzcGxheTpmbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGdhcDo4cHg7IHBhZGRpbmc6N3B4IDEycHg7CiAgICAgICAgICAgZm9udC1zaXplOjEzcHg7IGNvbG9yOiMxRjM4NjQ7IGN1cnNvcjpwb2ludGVyOyB3aGl0ZS1zcGFjZTpub3dyYXA7IH0KLm1zLWl0ZW06aG92ZXIgeyBiYWNrZ3JvdW5kOiNmMGY0ZmY7IH0KLm1zLWl0ZW0gaW5wdXQgeyB3aWR0aDoxNHB4OyBoZWlnaHQ6MTRweDsgYWNjZW50LWNvbG9yOiMyRTc1QjY7IGN1cnNvcjpwb2ludGVyOyBmbGV4LXNocmluazowOyB9Ci5tcy1kaXZpZGVyIHsgaGVpZ2h0OjFweDsgYmFja2dyb3VuZDojZWVlOyBtYXJnaW46NHB4IDA7IH0KPC9zdHlsZT4KPC9oZWFkPgo8Ym9keT4KCjxkaXYgaWQ9ImxvYWRpbmciPjxkaXYgY2xhc3M9InNwaW5uZXIiPjwvZGl2PiBMb2FkaW5nIGRhc2hib2FyZOKApjwvZGl2PgoKPCEtLSBIZWFkZXIgLS0+CjxkaXYgY2xhc3M9ImhlYWRlciI+CiAgPGRpdiBjbGFzcz0iaGVhZGVyLWxlZnQiPgogICAgPGRpdj4KICAgICAgPGRpdiBzdHlsZT0iZm9udC1zaXplOjE3cHg7Zm9udC13ZWlnaHQ6NzAwOyI+QWdlZCBJbnZlbnRvcnkgRGFzaGJvYXJkPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImhlYWRlci1zdWIiIGlkPSJoZWFkZXJTdWIiPkFzIG9mIHRvZGF5PC9kaXY+CiAgICA8L2Rpdj4KICA8L2Rpdj4KICA8YSBjbGFzcz0idXBsb2FkLWxpbmsiIGhyZWY9Ii8iPuKGkSBVcGxvYWQgTmV3IERhdGE8L2E+CjwvZGl2PgoKPCEtLSBUYWIgbmF2aWdhdGlvbiAtLT4KPGRpdiBjbGFzcz0idGFiLW5hdiI+CiAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biBhY3RpdmUiIG9uY2xpY2s9InN3aXRjaFRhYignbWFpbicsIHRoaXMpIj5NYWluIERhc2hib2FyZDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InRhYi1idG4iIG9uY2xpY2s9InN3aXRjaFRhYignc2Ryb3AnLCB0aGlzKSIgaWQ9InNkcm9wVGFiQnRuIj4KICAgIFMtRHJvcCBSZXZpZXcgPHNwYW4gY2xhc3M9InRhYi1iYWRnZSIgaWQ9InNkcm9wQmFkZ2UiPuKAlDwvc3Bhbj4KICA8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJ0YWItYnRuIiBvbmNsaWNrPSJzd2l0Y2hUYWIoJ2Nsb3NlZCcsIHRoaXMpIiBpZD0iY2xvc2VkVGFiQnRuIj4KICAgIENsb3NlZCBPcmRlcnMgPHNwYW4gY2xhc3M9InRhYi1iYWRnZSIgc3R5bGU9ImJhY2tncm91bmQ6IzM3NTYyMyIgaWQ9ImNsb3NlZEJhZGdlIj7igJQ8L3NwYW4+CiAgPC9idXR0b24+CiAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgb25jbGljaz0ic3dpdGNoVGFiKCdvZmZzaXRlJywgdGhpcykiIGlkPSJvZmZzaXRlVGFiQnRuIj4KICAgIFN0b3JhZ2UgPHNwYW4gY2xhc3M9InRhYi1iYWRnZSIgc3R5bGU9ImJhY2tncm91bmQ6IzJFNzVCNiIgaWQ9Im9mZnNpdGVCYWRnZSI+4oCUPC9zcGFuPgogIDwvYnV0dG9uPgogIDxidXR0b24gY2xhc3M9InRhYi1idG4iIG9uY2xpY2s9InN3aXRjaFRhYignbmV3c3RvcmFnZScsIHRoaXMpIiBpZD0ibmV3c3RvcmFnZVRhYkJ0biI+CiAgICBOZXcgdG8gU3RvcmFnZSA8c3BhbiBjbGFzcz0idGFiLWJhZGdlIiBzdHlsZT0iYmFja2dyb3VuZDojOEIwMDAwIiBpZD0ibmV3c3RvcmFnZUJhZGdlIj7igJQ8L3NwYW4+CiAgPC9idXR0b24+CjwvZGl2PgoKPCEtLSBNYWluIGRhc2hib2FyZCBmaWx0ZXJzIC0tPgo8ZGl2IGNsYXNzPSJmaWx0ZXJzLWJhciIgaWQ9Im1haW5GaWx0ZXJzIj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBQcm9qZWN0PC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1wcm9qZWN0Ij48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtcHJvamVjdCcsIGV2ZW50KVwiPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIFByb2plY3RzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtcHJvamVjdC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQ29vcmRpbmF0b3I8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLWNvb3JkIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtY29vcmQnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBDb29yZGluYXRvcnM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1jb29yZC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQWdlIEJ1Y2tldDwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtYnVja2V0Ij48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtYnVja2V0JywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgQnVja2V0czwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLWJ1Y2tldC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGJ1dHRvbiBjbGFzcz0icmVzZXQtYnRuIiBvbmNsaWNrPSJyZXNldEZpbHRlcnMoKSI+JiMxMDAwNTsgUmVzZXQ8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIHN0eWxlPSJiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjIpO2NvbG9yOiNmZmY7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjQpOyIgb25jbGljaz0iZXhwb3J0TWFpbkNTVigpIj4mIzExMDE1OyBFeHBvcnQgQ1NWPC9idXR0b24+CjwvZGl2PgoKPCEtLSBTLURyb3AgZmlsdGVycyAtLT4KPGRpdiBjbGFzcz0iZmlsdGVycy1iYXIiIGlkPSJzZHJvcEZpbHRlcnMiIHN0eWxlPSJkaXNwbGF5Om5vbmUiPgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IERyb3AgTG9jYXRpb248L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLWRyb3Bsb2MiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1kcm9wbG9jJywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgTG9jYXRpb25zPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtZHJvcGxvYy1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQ29vcmRpbmF0b3I8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLWRyb3Bjb29yZCI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIiBvbmNsaWNrPVwibXNUb2dnbGUoJ21zLWRyb3Bjb29yZCcsIGV2ZW50KVwiPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIENvb3JkaW5hdG9yczwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLWRyb3Bjb29yZC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQWdlIEZsYWc8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLWRyb3BmbGFnIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtZHJvcGZsYWcnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBBZ2VzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtZHJvcGZsYWctcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgb25jbGljaz0icmVzZXRTZHJvcEZpbHRlcnMoKSI+JiMxMDAwNTsgUmVzZXQ8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIHN0eWxlPSJiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjIpO2NvbG9yOiNmZmY7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjQpOyIgb25jbGljaz0iZXhwb3J0U2Ryb3BDU1YoKSI+JiMxMTAxNTsgRXhwb3J0IENTVjwvYnV0dG9uPgo8L2Rpdj4KCjwhLS0gQ2xvc2VkIE9yZGVycyBmaWx0ZXJzIC0tPgo8ZGl2IGNsYXNzPSJmaWx0ZXJzLWJhciIgaWQ9ImNsb3NlZEZpbHRlcnMiIHN0eWxlPSJkaXNwbGF5Om5vbmUiPgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IENvb3JkaW5hdG9yPC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1jbGNvb3JkIj48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtY2xjb29yZCcsIGV2ZW50KVwiPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIENvb3JkaW5hdG9yczwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLWNsY29vcmQtcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IEFnZSBCdWNrZXQ8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLWNsYnVja2V0Ij48YnV0dG9uIGNsYXNzPSJtcy1idG4iIG9uY2xpY2s9XCJtc1RvZ2dsZSgnbXMtY2xidWNrZXQnLCBldmVudClcIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBCdWNrZXRzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtY2xidWNrZXQtcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxidXR0b24gY2xhc3M9InJlc2V0LWJ0biIgb25jbGljaz0icmVzZXRDbG9zZWRGaWx0ZXJzKCkiPiYjMTAwMDU7IFJlc2V0PC9idXR0b24+CiAgPGJ1dHRvbiBjbGFzcz0icmVzZXQtYnRuIiBzdHlsZT0iYmFja2dyb3VuZDpyZ2JhKDI1NSwyNTUsMjU1LC4yKTtjb2xvcjojZmZmO2JvcmRlci1jb2xvcjpyZ2JhKDI1NSwyNTUsMjU1LC40KTsiIG9uY2xpY2s9ImV4cG9ydENsb3NlZENTVigpIj4mIzExMDE1OyBFeHBvcnQgQ1NWPC9idXR0b24+CjwvZGl2PgoKPCEtLSBPZmZzaXRlIFN0b3JhZ2UgZmlsdGVycyAtLT4KPGRpdiBjbGFzcz0iZmlsdGVycy1iYXIiIGlkPSJvZmZzaXRlRmlsdGVycyIgc3R5bGU9ImRpc3BsYXk6bm9uZSI+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQnVpbGRpbmc8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLW9mYnVpbGQiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1vZmJ1aWxkJywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgQnVpbGRpbmdzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtb2ZidWlsZC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQ29vcmRpbmF0b3I8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLW9mY29vcmQiPjxidXR0b24gY2xhc3M9Im1zLWJ0biIgb25jbGljaz1cIm1zVG9nZ2xlKCdtcy1vZmNvb3JkJywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgQ29vcmRpbmF0b3JzPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtb2Zjb29yZC1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgT3JkZXIgU3RhdHVzPC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1vZnN0YXR1cyI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIiBvbmNsaWNrPVwibXNUb2dnbGUoJ21zLW9mc3RhdHVzJywgZXZlbnQpXCI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgU3RhdHVzZXM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1vZnN0YXR1cy1wYW5lbCI+PC9kaXY+PC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyb3VwIj4KICAgIDxsYWJlbD5GaWx0ZXIgYnkgQWdlIEZsYWc8L2xhYmVsPgogICAgPGRpdiBjbGFzcz0ibXMtd3JhcCIgaWQ9Im1zLW9mZmxhZyI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIiBvbmNsaWNrPVwibXNUb2dnbGUoJ21zLW9mZmxhZycsIGV2ZW50KVwiPjxzcGFuIGNsYXNzPSJtcy1sYWJlbCI+QWxsIEFnZXM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1vZmZsYWctcGFuZWwiPjwvZGl2PjwvZGl2PgogIDwvZGl2PgogIDxkaXYgY2xhc3M9ImZpbHRlci1ncm91cCI+CiAgICA8bGFiZWw+RmlsdGVyIGJ5IERpc3Bvc2l0aW9uPC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1vZmRpc3AiPjxidXR0b24gY2xhc3M9Im1zLWJ0biI+PHNwYW4gY2xhc3M9Im1zLWxhYmVsIj5BbGwgRGlzcG9zaXRpb25zPC9zcGFuPjwvYnV0dG9uPjxkaXYgY2xhc3M9Im1zLXBhbmVsIiBpZD0ibXMtb2ZkaXNwLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBCaWxsZWQgU3RhdHVzPC9sYWJlbD4KICAgIDxkaXYgY2xhc3M9Im1zLXdyYXAiIGlkPSJtcy1vZmJpbGxlZCI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbDwvc3Bhbj48L2J1dHRvbj48ZGl2IGNsYXNzPSJtcy1wYW5lbCIgaWQ9Im1zLW9mYmlsbGVkLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIG9uY2xpY2s9InJlc2V0T2Zmc2l0ZUZpbHRlcnMoKSI+JiMxMDAwNTsgUmVzZXQ8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIHN0eWxlPSJiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjIpO2NvbG9yOiNmZmY7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjQpOyIgb25jbGljaz0iZXhwb3J0T2Zmc2l0ZUNTVigpIj4mIzExMDE1OyBFeHBvcnQgQ1NWPC9idXR0b24+CjwvZGl2PgoKPCEtLSBOZXcgdG8gU3RvcmFnZSBmaWx0ZXJzIC0tPgo8ZGl2IGNsYXNzPSJmaWx0ZXJzLWJhciIgaWQ9Im5ld3N0b3JhZ2VGaWx0ZXJzIiBzdHlsZT0iZGlzcGxheTpub25lIj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXItZ3JvdXAiPgogICAgPGxhYmVsPkZpbHRlciBieSBDb29yZGluYXRvcjwvbGFiZWw+CiAgICA8ZGl2IGNsYXNzPSJtcy13cmFwIiBpZD0ibXMtbnNjb29yZCI+PGJ1dHRvbiBjbGFzcz0ibXMtYnRuIj48c3BhbiBjbGFzcz0ibXMtbGFiZWwiPkFsbCBDb29yZGluYXRvcnM8L3NwYW4+PC9idXR0b24+PGRpdiBjbGFzcz0ibXMtcGFuZWwiIGlkPSJtcy1uc2Nvb3JkLXBhbmVsIj48L2Rpdj48L2Rpdj4KICA8L2Rpdj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIG9uY2xpY2s9InJlc2V0TmV3U3RvcmFnZUZpbHRlcnMoKSI+JiMxMDAwNTsgUmVzZXQ8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJyZXNldC1idG4iIHN0eWxlPSJiYWNrZ3JvdW5kOnJnYmEoMjU1LDI1NSwyNTUsLjIpO2NvbG9yOiNmZmY7Ym9yZGVyLWNvbG9yOnJnYmEoMjU1LDI1NSwyNTUsLjQpOyIgb25jbGljaz0iZXhwb3J0TmV3U3RvcmFnZUNTVigpIj4mIzExMDE1OyBFeHBvcnQgQ1NWPC9idXR0b24+CjwvZGl2PgoKPGRpdiBjbGFzcz0ibWFpbiI+CgogIDwhLS0g4pWQ4pWQIE1BSU4gVEFCIOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkCAtLT4KICA8ZGl2IGNsYXNzPSJ0YWItcGFuZWwgYWN0aXZlIiBpZD0icGFuZWwtbWFpbiI+CgogICAgPGRpdiBjbGFzcz0ia3BpLWdyaWQiPgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBJbnZlbnRvcnkgVmFsdWU8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJrVG90YWwiPuKAlDwvZGl2PjxkaXYgY2xhc3M9ImtwaS1zdWIiIGlkPSJrQ29udGFpbmVycyI+4oCUPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1yZWQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+T3ZlciA5MCBEYXlzIFZhbHVlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ia092ZXI5MCI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiIgaWQ9ImtPdmVyOTBQY3QiPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+SW4gU3RvcmFnZSBWYWx1ZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9ImtTdG9yYWdlIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWFtYmVyIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPkZpbmFuY2UgSG9sZCBWYWx1ZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9ImtGaW5hbmNlIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWdyZWVuIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPkF2ZXJhZ2UgQWdlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ia0F2Z0FnZSI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiI+ZGF5czwvZGl2PjwvZGl2PgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0idHdvLWNvbCI+CiAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tY2FyZCI+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1oZWFkZXIiPgogICAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+SW52ZW50b3J5IGJ5IFByb2plY3QgQ29vcmRpbmF0b3I8L2Rpdj4KICAgICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tbWV0YSIgaWQ9ImNvb3JkTWV0YSI+PC9kaXY+CiAgICAgICAgPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0idGFibGUtd3JhcCI+CiAgICAgICAgICA8dGFibGUgaWQ9ImNvb3JkVGFibGUiPgogICAgICAgICAgICA8dGhlYWQ+PHRyPgogICAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nvb3JkVGFibGUnLDApIj5Db29yZGluYXRvcjwvdGg+CiAgICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnY29vcmRUYWJsZScsMSkiPkNvbnRhaW5lcnM8L3RoPgogICAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nvb3JkVGFibGUnLDIpIj5WYWx1ZTwvdGg+CiAgICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnY29vcmRUYWJsZScsMykiPk92ZXIgOTBkIFZhbHVlPC90aD4KICAgICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdjb29yZFRhYmxlJyw0KSI+JSBPdmVyIDkwZDwvdGg+CiAgICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnY29vcmRUYWJsZScsNSkiPkF2ZyBBZ2U8L3RoPgogICAgICAgICAgICA8L3RyPjwvdGhlYWQ+CiAgICAgICAgICAgIDx0Ym9keSBpZD0iY29vcmRCb2R5Ij48L3Rib2R5PgogICAgICAgICAgICA8dGZvb3QgaWQ9ImNvb3JkRm9vdCI+PC90Zm9vdD4KICAgICAgICAgIDwvdGFibGU+CiAgICAgICAgPC9kaXY+CiAgICAgIDwvZGl2PgoKICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5JbnZlbnRvcnkgYnkgQWdlIEJ1Y2tldDwvZGl2PgogICAgICAgIDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9InRhYmxlLXdyYXAiPgogICAgICAgICAgPHRhYmxlIGlkPSJhZ2VUYWJsZSI+CiAgICAgICAgICAgIDx0aGVhZD48dHI+CiAgICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnYWdlVGFibGUnLDApIj5BZ2UgQnVja2V0PC90aD4KICAgICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdhZ2VUYWJsZScsMSkiPkNvbnRhaW5lcnM8L3RoPgogICAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ2FnZVRhYmxlJywyKSI+VmFsdWU8L3RoPgogICAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2FnZVRhYmxlJywzKSI+JSBvZiBUb3RhbDwvdGg+CiAgICAgICAgICAgIDwvdHI+PC90aGVhZD4KICAgICAgICAgICAgPHRib2R5IGlkPSJhZ2VCb2R5Ij48L3Rib2R5PgogICAgICAgICAgICA8dGZvb3QgaWQ9ImFnZUZvb3QiPjwvdGZvb3Q+CiAgICAgICAgICA8L3RhYmxlPgogICAgICAgIDwvZGl2PgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgoKICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24taGVhZGVyIj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5PcmRlcnMgaW4gVmlldzwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tbWV0YSIgaWQ9Im9yZGVyc01ldGEiPjwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic2VhcmNoLWJveCI+CiAgICAgICAgPGlucHV0IHR5cGU9InRleHQiIGlkPSJvcmRlclNlYXJjaCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBieSBvcmRlciAjLCBwcm9qZWN0LCBvciBjb29yZGluYXRvcuKApiIgb25pbnB1dD0icmVuZGVyT3JkZXJzKCkiPgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0idGFibGUtd3JhcCI+CiAgICAgICAgPHRhYmxlIGlkPSJvcmRlcnNUYWJsZSI+CiAgICAgICAgICA8dGhlYWQ+PHRyPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvcmRlcnNUYWJsZScsMCkiPk9yZGVyICM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvcmRlcnNUYWJsZScsMSkiPlByb2plY3Q8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvcmRlcnNUYWJsZScsMikiPkNvb3JkaW5hdG9yPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnb3JkZXJzVGFibGUnLDMpIj5Db250YWluZXJzPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnb3JkZXJzVGFibGUnLDQpIj5WYWx1ZTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ29yZGVyc1RhYmxlJyw1KSI+QXZnIEFnZTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ29yZGVyc1RhYmxlJyw2KSI+TWF4IEFnZTwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29yZGVyc1RhYmxlJyw3KSI+QWdlIEJ1Y2tldDwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29yZGVyc1RhYmxlJyw4KSI+U2hpcCBEYXRlIFJhbmdlPC90aD4KICAgICAgICAgIDwvdHI+PC90aGVhZD4KICAgICAgICAgIDx0Ym9keSBpZD0ib3JkZXJzQm9keSI+PC90Ym9keT4KICAgICAgICAgIDx0Zm9vdCBpZD0ib3JkZXJzRm9vdCI+PC90Zm9vdD4KICAgICAgICA8L3RhYmxlPgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgoKICA8L2Rpdj48IS0tIC9wYW5lbC1tYWluIC0tPgoKICA8IS0tIOKVkOKVkCBTLURST1AgVEFCIOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkCAtLT4KICA8ZGl2IGNsYXNzPSJ0YWItcGFuZWwiIGlkPSJwYW5lbC1zZHJvcCI+CgogICAgPGRpdiBjbGFzcz0ia3BpLWdyaWQiIHN0eWxlPSJncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCg1LCAxZnIpOyI+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1yZWQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VG90YWwgSXRlbXMgaW4gRHJvcDwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9InNkS3BpSXRlbXMiPuKAlDwvZGl2PjxkaXYgY2xhc3M9ImtwaS1zdWIiPm92ZXIgMiBkYXlzPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1yZWQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VG90YWwgVmFsdWU8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJzZEtwaVZhbHVlIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VW5pcXVlIE9yZGVyczwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9InNkS3BpT3JkZXJzIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWFtYmVyIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPkF2ZyBBZ2U8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJzZEtwaUF2ZyI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiI+ZGF5czwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+TWF4IEFnZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9InNkS3BpTWF4Ij7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktc3ViIj5kYXlzPC9kaXY+PC9kaXY+CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+QnkgRHJvcCBMb2NhdGlvbjwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic2Ryb3AtbG9jLWdyaWQiIGlkPSJzZHJvcExvY0dyaWQiPjwvZGl2PgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIj4KICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1oZWFkZXIiPgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPkl0ZW0gRGV0YWlsPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1tZXRhIiBpZD0ic2Ryb3BNZXRhIj48L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNlYXJjaC1ib3giPgogICAgICAgIDxpbnB1dCB0eXBlPSJ0ZXh0IiBpZD0ic2Ryb3BTZWFyY2giIHBsYWNlaG9sZGVyPSJTZWFyY2ggYnkgb3JkZXIgIywgcHJvamVjdCwgcGFydCBub+KApiIgb25pbnB1dD0icmVuZGVyU2Ryb3AoKSI+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJ0YWJsZS13cmFwIj4KICAgICAgICA8dGFibGUgaWQ9InNkcm9wVGFibGUiPgogICAgICAgICAgPHRoZWFkPjx0cj4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsMCkiPkxvY2F0aW9uPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsMSkiPk9yZGVyICM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzZHJvcFRhYmxlJywyKSI+UHJvamVjdDwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDMpIj5Db29yZGluYXRvcjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDQpIj5QYXJ0IE5vLjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDUpIj5TZXJpYWwgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDYpIj5QYXJ0IEdyb3VwPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsNykiPlF0eTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDgpIj5BZ2UgKGRheXMpPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsOSkiPlZhbHVlPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnc2Ryb3BUYWJsZScsMTApIj5TaGlwIERhdGU8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdzZHJvcFRhYmxlJywxMSkiPk9yZGVyIFN0YXR1czwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ3Nkcm9wVGFibGUnLDEyKSI+QWdlIEZsYWc8L3RoPgogICAgICAgICAgPC90cj48L3RoZWFkPgogICAgICAgICAgPHRib2R5IGlkPSJzZHJvcEJvZHkiPjwvdGJvZHk+CiAgICAgICAgICA8dGZvb3QgaWQ9InNkcm9wRm9vdCI+PC90Zm9vdD4KICAgICAgICA8L3RhYmxlPgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgoKICA8L2Rpdj48IS0tIC9wYW5lbC1zZHJvcCAtLT4KCiAgPCEtLSDilZDilZAgQ0xPU0VEIE9SREVSUyBUQUIg4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQ4pWQIC0tPgogIDxkaXYgY2xhc3M9InRhYi1wYW5lbCIgaWQ9InBhbmVsLWNsb3NlZCI+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCIgc3R5bGU9ImdyaWQtdGVtcGxhdGUtY29sdW1uczpyZXBlYXQoNSwxZnIpOyI+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1ncmVlbiI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5DbG9zZWQgSXRlbXM8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJjbEtwaUl0ZW1zIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWdyZWVuIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlRvdGFsIFZhbHVlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0iY2xLcGlWYWx1ZSI+4oCUPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlVuaXF1ZSBPcmRlcnM8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJjbEtwaU9yZGVycyI+4oCUPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1hbWJlciI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5BdmcgQWdlPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0iY2xLcGlBdmciPuKAlDwvZGl2PjxkaXYgY2xhc3M9ImtwaS1zdWIiPmRheXM8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWFtYmVyIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPk1heCBBZ2U8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJjbEtwaU1heCI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiI+ZGF5czwvZGl2PjwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+Q2xvc2VkIE9yZGVyIEl0ZW1zPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1tZXRhIiBpZD0iY2xvc2VkTWV0YSI+PC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJzZWFyY2gtYm94Ij4KICAgICAgICA8aW5wdXQgdHlwZT0idGV4dCIgaWQ9ImNsb3NlZFNlYXJjaCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBieSBvcmRlciAjLCBwcm9qZWN0LCBwYXJ0IG5vLi4uIiBvbmlucHV0PSJyZW5kZXJDbG9zZWQoKSI+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJ0YWJsZS13cmFwIj4KICAgICAgICA8dGFibGUgaWQ9ImNsb3NlZFRhYmxlIj4KICAgICAgICAgIDx0aGVhZD48dHI+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJywwKSI+T3JkZXIgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJywxKSI+UHJvamVjdDwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJywyKSI+Q29vcmRpbmF0b3I8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdjbG9zZWRUYWJsZScsMykiPlBhcnQgTm8uPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDQpIj5TZXJpYWwgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJyw1KSI+UGFydCBHcm91cDwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJyw2KSI+TG9jYXRpb248L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdjbG9zZWRUYWJsZScsNykiPlNoaXAgRGF0ZTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ2Nsb3NlZFRhYmxlJyw4KSI+UXR5PC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDkpIj5BZ2UgKGRheXMpPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnY2xvc2VkVGFibGUnLDEwKSI+VmFsdWU8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdjbG9zZWRUYWJsZScsMTEpIj5BZ2UgQnVja2V0PC90aD4KICAgICAgICAgIDwvdHI+PC90aGVhZD4KICAgICAgICAgIDx0Ym9keSBpZD0iY2xvc2VkQm9keSI+PC90Ym9keT4KICAgICAgICAgIDx0Zm9vdCBpZD0iY2xvc2VkRm9vdCI+PC90Zm9vdD4KICAgICAgICA8L3RhYmxlPgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgogIDwvZGl2PjwhLS0gL3BhbmVsLWNsb3NlZCAtLT4KCiAgPCEtLSDilZDilZAgT0ZGU0lURSBTVE9SQUdFIFRBQiDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZAgLS0+CiAgPGRpdiBjbGFzcz0idGFiLXBhbmVsIiBpZD0icGFuZWwtb2Zmc2l0ZSI+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCIgc3R5bGU9ImdyaWQtdGVtcGxhdGUtY29sdW1uczpyZXBlYXQoNSwxZnIpOyI+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlRvdGFsIEl0ZW1zPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ib2ZLcGlJdGVtcyI+4oCUPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1yZWQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VG90YWwgVmFsdWU8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJvZktwaVZhbHVlIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VW5pcXVlIE9yZGVyczwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpT3JkZXJzIj7igJQ8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQgYWNjZW50LWFtYmVyIj48ZGl2IGNsYXNzPSJrcGktbGFiZWwiPkF2ZyBBZ2U8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJvZktwaUF2ZyI+4oCUPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXN1YiI+ZGF5czwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCBhY2NlbnQtYW1iZXIiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+TWF4IEFnZTwvZGl2PjxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpTWF4Ij7igJQ8L2Rpdj48ZGl2IGNsYXNzPSJrcGktc3ViIj5kYXlzPC9kaXY+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImtwaS1ncmlkIiBzdHlsZT0iZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOnJlcGVhdCgzLDFmcik7bWFyZ2luLXRvcDoxMnB4OyI+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIiBzdHlsZT0iYm9yZGVyLXRvcDo0cHggc29saWQgIzJFNzVCNiI+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5TT1AgQnVpbGQgQWhlYWQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJvZktwaVNPUCI+4oCUPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiIgaWQ9Im9mS3BpU09QVmFsIiBzdHlsZT0iZm9udC1zaXplOjEzcHg7Y29sb3I6IzJFNzVCNjtmb250LXdlaWdodDo2MDAiPuKAlDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1zdWIiPm9yZGVycyDCtyB2YWx1ZTwvZGl2PgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiIHN0eWxlPSJib3JkZXItdG9wOjRweCBzb2xpZCAjMzc1NjIzIj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlN0b3JhZ2UgQ2hhcmdlZDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpQ2hhcmdlZCI+4oCUPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiIgaWQ9Im9mS3BpQ2hhcmdlZFZhbCIgc3R5bGU9ImZvbnQtc2l6ZToxM3B4O2NvbG9yOiMzNzU2MjM7Zm9udC13ZWlnaHQ6NjAwIj7igJQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIj5vcmRlcnMgwrcgdmFsdWU8L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIiBzdHlsZT0iYm9yZGVyLXRvcDo0cHggc29saWQgI2RjMzU0NSI+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5VbmJpbGxlZCBTdG9yYWdlPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ib2ZLcGlVbmJpbGxlZCI+4oCUPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiIgaWQ9Im9mS3BpVW5iaWxsZWRWYWwiIHN0eWxlPSJmb250LXNpemU6MTNweDtjb2xvcjojZGMzNTQ1O2ZvbnQtd2VpZ2h0OjYwMCI+4oCUPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLXN1YiI+b3JkZXJzIG5vdCB5ZXQgY2hhcmdlZDwvZGl2PgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0ia3BpLWdyaWQiIHN0eWxlPSJncmlkLXRlbXBsYXRlLWNvbHVtbnM6cmVwZWF0KDIsMWZyKTttYXJnaW4tdG9wOjEycHg7Ij4KICAgICAgPGRpdiBjbGFzcz0ia3BpLWNhcmQiIHN0eWxlPSJib3JkZXItdG9wOjRweCBzb2xpZCAjOEI0NTEzIj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktbGFiZWwiPlByb2plY3RzIEJlaW5nIEJpbGxlZDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS12YWx1ZSIgaWQ9Im9mS3BpQmlsbGVkQ291bnQiPuKAlDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImtwaS1zdWIiPnN0b3JhZ2UgY2hhcmdlcyByYWlzZWQ8L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIiBzdHlsZT0iYm9yZGVyLXRvcDo0cHggc29saWQgIzhCNDUxMyI+CiAgICAgICAgPGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBCaWxsZWQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJvZktwaUJpbGxlZFRvdGFsIj7igJQ8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJrcGktc3ViIj5jaGFyZ2UgYW1vdW50PC9kaXY+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWNhcmQiPgogICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLWhlYWRlciI+PGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+QnkgTG9jYXRpb24gR3JvdXA8L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic2Ryb3AtbG9jLWdyaWQiIGlkPSJvZmZzaXRlTG9jR3JpZCI+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24taGVhZGVyIj4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5JdGVtIERldGFpbDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tbWV0YSIgaWQ9Im9mZnNpdGVNZXRhIj48L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNlYXJjaC1ib3giPgogICAgICAgIDxpbnB1dCB0eXBlPSJ0ZXh0IiBpZD0ib2Zmc2l0ZVNlYXJjaCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBieSBvcmRlciAjLCBwcm9qZWN0LCBwYXJ0IG5vLi4uIiBvbmlucHV0PSJyZW5kZXJPZmZzaXRlKCkiPgogICAgICA8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0idGFibGUtd3JhcCI+CiAgICAgICAgPHRhYmxlIGlkPSJvZmZzaXRlVGFibGUiPgogICAgICAgICAgPHRoZWFkPjx0cj4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywwKSI+QnVpbGRpbmc8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDEpIj5Mb2NhdGlvbiBHcm91cDwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMikiPk9yZGVyICM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDMpIj5Qcm9qZWN0PC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJyw0KSI+Q29vcmRpbmF0b3I8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDUpIj5Mb2NhdGlvbjwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsNikiPlBhcnQgTm8uPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJyw3KSI+U2VyaWFsICM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDgpIj5TaGlwIERhdGU8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDkpIj5Db250YWluZXJzPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywxMCkiPlZhbHVlPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywxMSkiPkF2ZyBBZ2U8L3RoPgogICAgICAgICAgICA8dGggY2xhc3M9Im51bSIgb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDEyKSI+TWF4IEFnZTwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTMpIj5PcmRlciBTdGF0dXM8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDE0KSI+QWdlIEZsYWc8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDE1KSI+Q2hhcmdlIEFtb3VudDwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTYpIj5DaGFyZ2UgVHlwZTwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTcpIj5CaWxsaW5nIE5vdGU8L3RoPgogICAgICAgICAgICA8dGg+RGlzcG9zaXRpb248L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCdvZmZzaXRlVGFibGUnLDE2KSI+QmlsbGVkPzwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ29mZnNpdGVUYWJsZScsMTcpIj5DaGFyZ2UgQW10PC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnb2Zmc2l0ZVRhYmxlJywxOCkiPkNoYXJnZSBUeXBlPC90aD4KICAgICAgICAgIDwvdHI+PC90aGVhZD4KICAgICAgICAgIDx0Ym9keSBpZD0ib2Zmc2l0ZUJvZHkiPjwvdGJvZHk+CiAgICAgICAgICA8dGZvb3QgaWQ9Im9mZnNpdGVGb290Ij48L3Rmb290PgogICAgICAgIDwvdGFibGU+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgPC9kaXY+PCEtLSAvcGFuZWwtb2Zmc2l0ZSAtLT4KCiAgPCEtLSDilZDilZAgTkVXIFRPIFNUT1JBR0UgVEFCIOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkOKVkCAtLT4KICA8ZGl2IGNsYXNzPSJ0YWItcGFuZWwiIGlkPSJwYW5lbC1uZXdzdG9yYWdlIj4KICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tY2FyZCIgc3R5bGU9Im1hcmdpbi1ib3R0b206MTZweDtiYWNrZ3JvdW5kOiNmZmY4Zjg7Ym9yZGVyOjFweCBzb2xpZCAjZjVjNmM2OyI+CiAgICAgIDxkaXYgc3R5bGU9InBhZGRpbmc6MTZweCAyMHB4O2ZvbnQtc2l6ZToxM3B4O2NvbG9yOiM3QjAwMDA7bGluZS1oZWlnaHQ6MS42Ij4KICAgICAgICA8c3Ryb25nPiYjMTI4NjgwOyBXaGF0IGlzIHRoaXMgdGFiPzwvc3Ryb25nPjxicj4KICAgICAgICBUaGlzIHRhYiBzaG93cyBvcmRlcnMgd2hvc2Ugc3RhdHVzIDxzdHJvbmc+Y2hhbmdlZCB0byBzdGFydCB3aXRoICJTdG9yYWdlIjwvc3Ryb25nPiBzaW5jZSB0aGUgcHJldmlvdXMgdXBsb2FkLgogICAgICAgIEl0IGlzIGEgY3VtdWxhdGl2ZSBsb2cg4oCUIGVudHJpZXMgcmVtYWluIGhlcmUgYWNyb3NzIGZ1dHVyZSB1cGxvYWRzIHNvIHlvdSBjYW4gdHJhY2sgd2hlbiBpdGVtcyBmaXJzdCBtb3ZlZCBpbnRvIHN0b3JhZ2UuCiAgICAgICAgVXBsb2FkIG5ldyBmaWxlcyBkYWlseSB0byBrZWVwIHRoaXMgbGlzdCBjdXJyZW50LgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0ia3BpLWdyaWQiIHN0eWxlPSJncmlkLXRlbXBsYXRlLWNvbHVtbnM6cmVwZWF0KDMsMWZyKTsiPgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCIgc3R5bGU9ImJvcmRlci10b3A6NHB4IHNvbGlkICM4QjAwMDAiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+TmV3IFRoaXMgVXBsb2FkPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ibnNLcGlOZXciPuKAlDwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJrcGktY2FyZCI+PGRpdiBjbGFzcz0ia3BpLWxhYmVsIj5Ub3RhbCBUcmFja2VkPC9kaXY+PGRpdiBjbGFzcz0ia3BpLXZhbHVlIiBpZD0ibnNLcGlUb3RhbCI+4oCUPC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImtwaS1jYXJkIGFjY2VudC1yZWQiPjxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+VG90YWwgVmFsdWU8L2Rpdj48ZGl2IGNsYXNzPSJrcGktdmFsdWUiIGlkPSJuc0twaVZhbHVlIj7igJQ8L2Rpdj48L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1jYXJkIj4KICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1oZWFkZXIiPgogICAgICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPk9yZGVycyBNb3ZlZCB0byBTdG9yYWdlPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VjdGlvbi1tZXRhIiBpZD0ibnNNZXRhIj48L2Rpdj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNlYXJjaC1ib3giPgogICAgICAgIDxpbnB1dCB0eXBlPSJ0ZXh0IiBpZD0ibnNTZWFyY2giIHBsYWNlaG9sZGVyPSJTZWFyY2ggYnkgb3JkZXIgIywgcHJvamVjdCwgY29vcmRpbmF0b3IuLi4iIG9uaW5wdXQ9InJlbmRlck5ld1N0b3JhZ2UoKSI+CiAgICAgIDwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJ0YWJsZS13cmFwIj4KICAgICAgICA8dGFibGUgaWQ9Im5zVGFibGUiPgogICAgICAgICAgPHRoZWFkPjx0cj4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsMCkiPkRldGVjdGVkIERhdGU8L3RoPgogICAgICAgICAgICA8dGggb25jbGljaz0ic29ydFRhYmxlKCduc1RhYmxlJywxKSI+T3JkZXIgIzwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ25zVGFibGUnLDIpIj5Qcm9qZWN0PC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsMykiPkNvb3JkaW5hdG9yPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsNCkiPlByZXZpb3VzIFN0YXR1czwvdGg+CiAgICAgICAgICAgIDx0aCBvbmNsaWNrPSJzb3J0VGFibGUoJ25zVGFibGUnLDUpIj5OZXcgU3RhdHVzPC90aD4KICAgICAgICAgICAgPHRoIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsNikiPlNoaXAgRGF0ZTwvdGg+CiAgICAgICAgICAgIDx0aCBjbGFzcz0ibnVtIiBvbmNsaWNrPSJzb3J0VGFibGUoJ25zVGFibGUnLDcpIj5Db250YWluZXJzPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsOCkiPlZhbHVlPC90aD4KICAgICAgICAgICAgPHRoIGNsYXNzPSJudW0iIG9uY2xpY2s9InNvcnRUYWJsZSgnbnNUYWJsZScsOSkiPkF2ZyBBZ2U8L3RoPgogICAgICAgICAgPC90cj48L3RoZWFkPgogICAgICAgICAgPHRib2R5IGlkPSJuc0JvZHkiPjwvdGJvZHk+CiAgICAgICAgICA8dGZvb3QgaWQ9Im5zRm9vdCI+PC90Zm9vdD4KICAgICAgICA8L3RhYmxlPgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgogIDwvZGl2PjwhLS0gL3BhbmVsLW5ld3N0b3JhZ2UgLS0+Cgo8L2Rpdj48IS0tIC9tYWluIC0tPgoKPHNjcmlwdD4KLy8g4pSA4pSAIEdsb2JhbHMg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmxldCBSQVcgPSBudWxsOwpsZXQgZmlsdGVyZWRPcmRlcnMgPSBbXTsKY29uc3Qgc29ydFN0YXRlID0ge307Cgpjb25zdCBBR0VfT1JERVIgID0gWycwLTdkJywnOC0xNGQnLCcxNS0zMGQnLCczMS02MGQnLCc2MS05MGQnLCc5MS0xODBkJywnMTgxLTM2NWQnLCczNjVkKyddOwpjb25zdCBBR0VfTEFCRUxTID0geycwLTdkJzonMOKAkzdkJywnOC0xNGQnOic44oCTMTRkJywnMTUtMzBkJzonMTXigJMzMGQnLCczMS02MGQnOiczMeKAkzYwZCcsCiAgJzYxLTkwZCc6JzYx4oCTOTBkJywnOTEtMTgwZCc6Jzkx4oCTMTgwZCcsJzE4MS0zNjVkJzonMTgx4oCTMzY1ZCcsJzM2NWQrJzonPjM2NWQnfTsKCi8vIOKUgOKUgCBGb3JtYXR0ZXJzIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApjb25zdCBmbXQkICA9IHYgPT4gdiA9PSBudWxsID8gJ+KAlCcgOiAnJCcgKyBNYXRoLnJvdW5kKHYpLnRvTG9jYWxlU3RyaW5nKCk7CmNvbnN0IGZtdE4gID0gdiA9PiB2ID09IG51bGwgPyAn4oCUJyA6IE1hdGgucm91bmQodikudG9Mb2NhbGVTdHJpbmcoKTsKY29uc3QgZm10UGN0PSB2ID0+IHYgPT0gbnVsbCA/ICfigJQnIDogKHYqMTAwKS50b0ZpeGVkKDEpKyclJzsKY29uc3QgZm10RiAgPSB2ID0+IHYgPT0gbnVsbCA/ICfigJQnIDogdi50b0ZpeGVkKDEpOwoKZnVuY3Rpb24gYWdlQnVja2V0VGFnKGNvZGUpIHsKICBpZiAoIWNvZGUpIHJldHVybiAnJzsKICBjb25zdCBjbHMgPSAoY29kZT09PSczNjVkKyd8fGNvZGU9PT0nMTgxLTM2NWQnKSA/ICd0YWctcmVkJwogICAgICAgICAgICA6IGNvZGU9PT0nOTEtMTgwZCcgPyAndGFnLWFtYmVyJyA6ICd0YWctYmx1ZSc7CiAgcmV0dXJuIGA8c3BhbiBjbGFzcz0idGFnICR7Y2xzfSI+JHtBR0VfTEFCRUxTW2NvZGVdfHxjb2RlfTwvc3Bhbj5gOwp9CmZ1bmN0aW9uIG1heEFnZVRhZyhhZ2UpIHsKICBjb25zdCBjbHMgPSBhZ2U+OTAgPyAndGFnLXJlZCcgOiBhZ2U+MzAgPyAndGFnLWFtYmVyJyA6ICd0YWctZ3JlZW4nOwogIHJldHVybiBgPHNwYW4gY2xhc3M9InRhZyAke2Nsc30iPiR7YWdlfWQ8L3NwYW4+YDsKfQpmdW5jdGlvbiBzZHJvcEFnZUZsYWcoYWdlKSB7CiAgaWYgKGFnZSA8PSAyKSAgcmV0dXJuICcyIERheXMgb3IgTGVzcyc7CiAgaWYgKGFnZSA8PSA1KSAgcmV0dXJuICcy4oCTNSBEYXlzJzsKICBpZiAoYWdlIDw9IDEwKSByZXR1cm4gJzbigJMxMCBEYXlzJzsKICByZXR1cm4gJzEwKyBEYXlzJzsKfQpmdW5jdGlvbiBzZHJvcEZsYWdUYWcoYWdlKSB7CiAgY29uc3QgZiA9IHNkcm9wQWdlRmxhZyhhZ2UpOwogIGNvbnN0IGNscyA9IGFnZSA+IDEwID8gJ3RhZy1yZWQnIDogYWdlID4gNSA/ICd0YWctYW1iZXInIDogYWdlID4gMiA/ICd0YWctZ3JlZW4nIDogJyc7CiAgcmV0dXJuIGA8c3BhbiBjbGFzcz0idGFnICR7Y2xzfSI+JHtmfTwvc3Bhbj5gOwp9CmZ1bmN0aW9uIGZsYWdUYWcoZmxhZykgewogIGNvbnN0IGNscyA9IGZsYWc9PT0nPjkwIERheXMnID8gJ3RhZy1yZWQnIDogZmxhZz09PSc+MzAgRGF5cycgPyAndGFnLWFtYmVyJyA6ICd0YWctZ3JlZW4nOwogIHJldHVybiBgPHNwYW4gY2xhc3M9InRhZyAke2Nsc30iPiR7ZmxhZ308L3NwYW4+YDsKfQoKLy8g4pSA4pSAIFRhYiBzd2l0Y2hpbmcg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHN3aXRjaFRhYihuYW1lLCBidG4pIHsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcudGFiLXBhbmVsJykuZm9yRWFjaChwID0+IHAuY2xhc3NMaXN0LnJlbW92ZSgnYWN0aXZlJykpOwogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy50YWItYnRuJykuZm9yRWFjaChiID0+IGIuY2xhc3NMaXN0LnJlbW92ZSgnYWN0aXZlJykpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdwYW5lbC0nICsgbmFtZSkuY2xhc3NMaXN0LmFkZCgnYWN0aXZlJyk7CiAgYnRuLmNsYXNzTGlzdC5hZGQoJ2FjdGl2ZScpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdtYWluRmlsdGVycycpLnN0eWxlLmRpc3BsYXkgICA9IG5hbWU9PT0nbWFpbicgICAgPyAnJyA6ICdub25lJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BGaWx0ZXJzJykuc3R5bGUuZGlzcGxheSAgPSBuYW1lPT09J3Nkcm9wJyAgID8gJycgOiAnbm9uZSc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nsb3NlZEZpbHRlcnMnKS5zdHlsZS5kaXNwbGF5ID0gbmFtZT09PSdjbG9zZWQnICA/ICcnIDogJ25vbmUnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZmZzaXRlRmlsdGVycycpLnN0eWxlLmRpc3BsYXkgICAgPSBuYW1lPT09J29mZnNpdGUnICAgICA/ICcnIDogJ25vbmUnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduZXdzdG9yYWdlRmlsdGVycycpLnN0eWxlLmRpc3BsYXkgPSBuYW1lPT09J25ld3N0b3JhZ2UnICA/ICcnIDogJ25vbmUnOwogIGlmIChuYW1lID09PSAnbmV3c3RvcmFnZScgJiYgIU5TX0xPQURFRCkgbG9hZE5ld1N0b3JhZ2UoKTsKfQoKLy8g4pSA4pSAIExvYWQgZGF0YSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKYXN5bmMgZnVuY3Rpb24gbG9hZERhdGEoKSB7CiAgYXdhaXQgbG9hZERpc3Bvc2l0aW9ucygpOwogIGF3YWl0IGxvYWRCaWxsaW5nKCk7CiAgdHJ5IHsKICAgIC8vIDEwIHNlY29uZCB0aW1lb3V0IOKAlCBpZiBzZXJ2ZXIgZG9lc24ndCByZXNwb25kLCBzaG93IGVycm9yCiAgICBjb25zdCBjb250cm9sbGVyID0gbmV3IEFib3J0Q29udHJvbGxlcigpOwogICAgY29uc3QgdGltZW91dCA9IHNldFRpbWVvdXQoKCkgPT4gY29udHJvbGxlci5hYm9ydCgpLCAxMDAwMCk7CiAgICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaCgnL2FwaS9kYXRhJywgeyBzaWduYWw6IGNvbnRyb2xsZXIuc2lnbmFsIH0pOwogICAgY2xlYXJUaW1lb3V0KHRpbWVvdXQpOwogICAgaWYgKCFyZXMub2spIHRocm93IG5ldyBFcnJvcignbm9fZGF0YScpOwogICAgUkFXID0gYXdhaXQgcmVzLmpzb24oKTsKICB9IGNhdGNoKGUpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdsb2FkaW5nJykuaW5uZXJIVE1MID0gYAogICAgICA8ZGl2IHN0eWxlPSJ0ZXh0LWFsaWduOmNlbnRlcjtwYWRkaW5nOjQwcHgiPgogICAgICAgIDxkaXYgc3R5bGU9ImZvbnQtc2l6ZTo0OHB4O21hcmdpbi1ib3R0b206MTZweCI+8J+TgjwvZGl2PgogICAgICAgIDxkaXYgc3R5bGU9ImZvbnQtc2l6ZToyMHB4O2ZvbnQtd2VpZ2h0OjcwMDttYXJnaW4tYm90dG9tOjhweCI+Tm8gZGF0YSBsb2FkZWQgeWV0PC9kaXY+CiAgICAgICAgPGRpdiBzdHlsZT0iZm9udC1zaXplOjE0cHg7b3BhY2l0eTouODttYXJnaW4tYm90dG9tOjI0cHgiPlVwbG9hZCB5b3VyIHdhcmVob3VzZSBhbmQgaW50ZXJuYWwgbm90ZXMgZmlsZXMgdG8gZ2V0IHN0YXJ0ZWQuPC9kaXY+CiAgICAgICAgPGEgaHJlZj0iLyIgc3R5bGU9ImJhY2tncm91bmQ6I2ZmZjtjb2xvcjojMUYzODY0O3BhZGRpbmc6MTJweCAyOHB4O2JvcmRlci1yYWRpdXM6OHB4O3RleHQtZGVjb3JhdGlvbjpub25lO2ZvbnQtd2VpZ2h0OjcwMDtmb250LXNpemU6MTRweCI+4oaRIFVwbG9hZCBGaWxlczwvYT4KICAgICAgPC9kaXY+YDsKICAgIHJldHVybjsKICB9CgogIGlmICghUkFXLnNkcm9wKSBSQVcuc2Ryb3AgPSB7IGtwaXM6e30sIGJ5X2xvY2F0aW9uOltdLCBpdGVtczpbXSB9OwoKICAvLyBTZXQgdGhlIHVwbG9hZGVkIGRhdGUgaW4gdGhlIGhlYWRlcgogIGlmIChSQVcudXBsb2FkZWQpIHsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdoZWFkZXJTdWInKS50ZXh0Q29udGVudCA9ICdBcyBvZiAnICsgUkFXLnVwbG9hZGVkOwogIH0KCiAgLy8gVXBkYXRlIHRhYiBiYWRnZXMKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BCYWRnZScpLnRleHRDb250ZW50ICAgPSBSQVcuc2Ryb3A/LmtwaXM/LnRvdGFsX2l0ZW1zICAgPz8gMDsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xvc2VkQmFkZ2UnKS50ZXh0Q29udGVudCAgPSBSQVcuY2xvc2VkPy5rcGlzPy50b3RhbF9pdGVtcyAgPz8gMDsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUJhZGdlJykudGV4dENvbnRlbnQgPSBSQVcub2Zmc2l0ZT8ua3Bpcz8udG90YWxfaXRlbXMgPz8gMDsKCiAgLy8gUG9wdWxhdGUgYWxsIG11bHRpLXNlbGVjdCBmaWx0ZXJzCiAgbXNQb3B1bGF0ZSgnbXMtcHJvamVjdCcsIFJBVy5wcm9qZWN0cyB8fCBbXSwgYXBwbHlGaWx0ZXJzKTsKICBjb25zdCBtYWluQ29vcmRzID0gWy4uLm5ldyBTZXQoKFJBVy5vcmRlcnNfdGFibGV8fFtdKS5tYXAociA9PiByLmNvb3JkaW5hdG9yKSldLmZpbHRlcihCb29sZWFuKS5zb3J0KCk7CiAgbXNQb3B1bGF0ZSgnbXMtY29vcmQnLCBtYWluQ29vcmRzLCBhcHBseUZpbHRlcnMpOwogIG1zUG9wdWxhdGUoJ21zLWJ1Y2tldCcsICBbJzAtN2QnLCc4LTE0ZCcsJzE1LTMwZCcsJzMxLTYwZCcsJzYxLTkwZCcsJzkxLTE4MGQnLCcxODEtMzY1ZCcsJzM2NWQrJ10sIGFwcGx5RmlsdGVycyk7CgogIGNvbnN0IHNkcm9wQ29vcmRzID0gWy4uLm5ldyBTZXQoKFJBVy5zZHJvcD8uaXRlbXN8fFtdKS5tYXAocj0+ci5jb29yZGluYXRvcikpXS5maWx0ZXIoQm9vbGVhbikuc29ydCgpOwogIG1zUG9wdWxhdGUoJ21zLWRyb3Bsb2MnLCAgIFJBVy5kcm9wX2xvY2F0aW9ucyB8fCBbXSwgcmVuZGVyU2Ryb3ApOwogIG1zUG9wdWxhdGUoJ21zLWRyb3Bjb29yZCcsIHNkcm9wQ29vcmRzLCByZW5kZXJTZHJvcCk7CiAgbXNQb3B1bGF0ZSgnbXMtZHJvcGZsYWcnLCAgWycyIERheXMgb3IgTGVzcycsJzLigJM1IERheXMnLCc24oCTMTAgRGF5cycsJzEwKyBEYXlzJ10sIHJlbmRlclNkcm9wKTsKCiAgY29uc3QgY2xDb29yZHMgPSBbLi4ubmV3IFNldCgoUkFXLmNsb3NlZD8uaXRlbXN8fFtdKS5tYXAociA9PiByLmNvb3JkaW5hdG9yKSldLmZpbHRlcihCb29sZWFuKS5zb3J0KCk7CiAgbXNQb3B1bGF0ZSgnbXMtY2xjb29yZCcsIGNsQ29vcmRzLCByZW5kZXJDbG9zZWQpOwogIG1zUG9wdWxhdGUoJ21zLWNsYnVja2V0JywgWycwLTdkJywnOC0xNGQnLCcxNS0zMGQnLCczMS02MGQnLCc2MS05MGQnLCc5MS0xODBkJywnMTgxLTM2NWQnLCczNjVkKyddLCByZW5kZXJDbG9zZWQpOwoKICBtc1BvcHVsYXRlKCdtcy1vZmJ1aWxkJywgIFsnRG93bnN0YWlycyBTdG9yYWdlJywnUG9ydGVyIFN0b3JhZ2UnLCdOb3J0aCBXYXJlaG91c2UnLCdUcmFpbGVycyddLCByZW5kZXJPZmZzaXRlKTsKICAvLyBEZXJpdmUgY29vcmRpbmF0b3JzIGZyb20gaXRlbXMgZGlyZWN0bHkgc28gVW5hc3NpZ25lZCBpcyBhbHdheXMgaW5jbHVkZWQKICBjb25zdCBvZkNvb3JkcyA9IFsuLi5uZXcgU2V0KChSQVcub2Zmc2l0ZT8uaXRlbXN8fFtdKS5tYXAociA9PiByLmNvb3JkaW5hdG9yKSldLmZpbHRlcihCb29sZWFuKS5zb3J0KCk7CiAgbXNQb3B1bGF0ZSgnbXMtb2Zjb29yZCcsIG9mQ29vcmRzLCByZW5kZXJPZmZzaXRlKTsKICBtc1BvcHVsYXRlKCdtcy1vZnN0YXR1cycsIFJBVy5vZmZzaXRlPy5zdGF0dXNlcyB8fCBbXSwgcmVuZGVyT2Zmc2l0ZSk7CiAgbXNQb3B1bGF0ZSgnbXMtb2ZmbGFnJywgICBbJz45MCBEYXlzJywnPjMwIERheXMnLCfiiaQzMCBEYXlzJ10sIHJlbmRlck9mZnNpdGUpOwogIG1zUG9wdWxhdGUoJ21zLW9mZGlzcCcsICAgWydTT1AgQnVpbGQgQWhlYWQnLCdTdG9yYWdlIENoYXJnZWQnLCcoTm9uZSknXSwgcmVuZGVyT2Zmc2l0ZSk7CiAgbXNQb3B1bGF0ZSgnbXMtb2ZiaWxsZWQnLCBbJ0JpbGxlZCcsJ05vdCBCaWxsZWQnXSwgcmVuZGVyT2Zmc2l0ZSk7CgogIC8vIFJlbmRlciBuZXcgdGFicwogIHJlbmRlckNsb3NlZEtwaXMoKTsKICByZW5kZXJPZmZzaXRlS3BpcygpOwogIHJlbmRlck9mZnNpdGVMb2NHcmlkKCk7CgogIGFwcGx5RmlsdGVycygpOwogIHJlbmRlclNkcm9wS3BpcygpOwogIHJlbmRlclNkcm9wTG9jR3JpZCgpOwogIHJlbmRlclNkcm9wKCk7CiAgcmVuZGVyQ2xvc2VkKCk7CiAgcmVuZGVyT2Zmc2l0ZSgpOwogIHVwZGF0ZURpc3Bvc2l0aW9uS1BJcygpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdsb2FkaW5nJykuc3R5bGUuZGlzcGxheSA9ICdub25lJzsKfQoKLy8g4pSA4pSAIE1haW4gZmlsdGVycyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gZ2V0RmlsdGVycygpIHsKICByZXR1cm4gewogICAgcHJvamVjdDogbXNHZXRTZWxlY3RlZCgnbXMtcHJvamVjdCcpLAogICAgY29vcmQ6ICAgbXNHZXRTZWxlY3RlZCgnbXMtY29vcmQnKSwKICAgIGJ1Y2tldDogIG1zR2V0U2VsZWN0ZWQoJ21zLWJ1Y2tldCcpLAogIH07Cn0KZnVuY3Rpb24gcmVzZXRGaWx0ZXJzKCkgewogIFsnZmlsdGVyUHJvamVjdCcsJ2ZpbHRlckNvb3JkJywnZmlsdGVyQnVja2V0J10uZm9yRWFjaChpZCA9PiBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChpZCkudmFsdWU9JycpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvcmRlclNlYXJjaCcpLnZhbHVlID0gJyc7CiAgYXBwbHlGaWx0ZXJzKCk7Cn0KCmZ1bmN0aW9uIGFwcGx5RmlsdGVycygpIHsKICBpZiAoIVJBVykgcmV0dXJuOwogIGNvbnN0IHtwcm9qZWN0LCBjb29yZCwgYnVja2V0fSA9IGdldEZpbHRlcnMoKTsKICBmaWx0ZXJlZE9yZGVycyA9IFJBVy5vcmRlcnNfdGFibGUuZmlsdGVyKG8gPT4KICAgICghcHJvamVjdC5sZW5ndGggfHwgcHJvamVjdC5pbmNsdWRlcyhvLnByb2plY3QpKSAmJgogICAgKCFjb29yZC5sZW5ndGggICB8fCBjb29yZC5pbmNsdWRlcyhvLmNvb3JkaW5hdG9yKSkgJiYKICAgICghYnVja2V0Lmxlbmd0aCAgfHwgYnVja2V0LmluY2x1ZGVzKG8uYWdlX2J1Y2tldCkpCiAgKTsKICByZW5kZXJLUElzKCk7IHJlbmRlckNvb3JkVGFibGUoKTsgcmVuZGVyQWdlVGFibGUoKTsgcmVuZGVyT3JkZXJzKCk7Cn0KCi8vIOKUgOKUgCBLUElzIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiByZW5kZXJLUElzKCkgewogIGNvbnN0IHtwcm9qZWN0LCBjb29yZCwgYnVja2V0fSA9IGdldEZpbHRlcnMoKTsKICBjb25zdCBoYXNGaWx0ZXIgPSBwcm9qZWN0Lmxlbmd0aCB8fCBjb29yZC5sZW5ndGggfHwgYnVja2V0Lmxlbmd0aDsKICBpZiAoIWhhc0ZpbHRlcikgewogICAgY29uc3QgayA9IFJBVy5rcGlzOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tUb3RhbCcpLnRleHRDb250ZW50ICAgICAgPSBmbXQkKGsudG90YWxfdmFsdWUpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tDb250YWluZXJzJykudGV4dENvbnRlbnQgPSBmbXROKGsudG90YWxfY29udGFpbmVycykgKyAnIGNvbnRhaW5lcnMnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tPdmVyOTAnKS50ZXh0Q29udGVudCAgICAgPSBmbXQkKGsub3ZlcjkwX3ZhbHVlKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrT3ZlcjkwUGN0JykudGV4dENvbnRlbnQgID0gZm10UGN0KGsub3ZlcjkwX3BjdCkgKyAnIG9mIHRvdGFsJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrU3RvcmFnZScpLnRleHRDb250ZW50ICAgID0gZm10JChrLmluX3N0b3JhZ2VfdmFsdWUpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tGaW5hbmNlJykudGV4dENvbnRlbnQgICAgPSBmbXQkKGsuZmluYW5jZV9ob2xkX3ZhbHVlKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrQXZnQWdlJykudGV4dENvbnRlbnQgICAgID0gZm10RihrLmF2Z19hZ2UpOwogIH0gZWxzZSB7CiAgICBsZXQgdHY9MCwgdGM9MCwgbzkwPTAsIHRhPTA7CiAgICBmaWx0ZXJlZE9yZGVycy5mb3JFYWNoKG8gPT4geyB0dis9by52YWx1ZTsgdGMrPW8uY29udGFpbmVyczsgaWYoby5tYXhfYWdlPjkwKSBvOTArPW8udmFsdWU7IHRhKz1vLmF2Z19hZ2Uqby5jb250YWluZXJzOyB9KTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrVG90YWwnKS50ZXh0Q29udGVudCAgICAgID0gZm10JCh0dik7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna0NvbnRhaW5lcnMnKS50ZXh0Q29udGVudCA9IGZtdE4odGMpICsgJyBjb250YWluZXJzJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrT3ZlcjkwJykudGV4dENvbnRlbnQgICAgID0gZm10JChvOTApOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tPdmVyOTBQY3QnKS50ZXh0Q29udGVudCAgPSB0diA/IGZtdFBjdChvOTAvdHYpKycgb2YgdG90YWwnIDogJ+KAlCc7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgna1N0b3JhZ2UnKS50ZXh0Q29udGVudCAgICA9ICfigJQnOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2tGaW5hbmNlJykudGV4dENvbnRlbnQgICAgPSAn4oCUJzsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdrQXZnQWdlJykudGV4dENvbnRlbnQgICAgID0gdGMgPyAodGEvdGMpLnRvRml4ZWQoMSkgOiAn4oCUJzsKICB9Cn0KCi8vIOKUgOKUgCBDb29yZGluYXRvciB0YWJsZSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyQ29vcmRUYWJsZSgpIHsKICBjb25zdCB7cHJvamVjdCwgY29vcmQsIGJ1Y2tldH0gPSBnZXRGaWx0ZXJzKCk7CiAgbGV0IHJvd3M7CiAgaWYgKCFwcm9qZWN0ICYmICFjb29yZCAmJiAhYnVja2V0KSB7CiAgICByb3dzID0gUkFXLmNvb3JkX3RhYmxlOwogIH0gZWxzZSB7CiAgICBjb25zdCBtYXAgPSB7fTsKICAgIGZpbHRlcmVkT3JkZXJzLmZvckVhY2gobyA9PiB7CiAgICAgIGlmICghbWFwW28uY29vcmRpbmF0b3JdKSBtYXBbby5jb29yZGluYXRvcl09e2Nvb3JkaW5hdG9yOm8uY29vcmRpbmF0b3IsY29udGFpbmVyczowLHZhbHVlOjAsb3ZlcjkwX3ZhbHVlOjAsYWdlczpbXX07CiAgICAgIG1hcFtvLmNvb3JkaW5hdG9yXS5jb250YWluZXJzKz1vLmNvbnRhaW5lcnM7IG1hcFtvLmNvb3JkaW5hdG9yXS52YWx1ZSs9by52YWx1ZTsKICAgICAgaWYoby5tYXhfYWdlPjkwKSBtYXBbby5jb29yZGluYXRvcl0ub3ZlcjkwX3ZhbHVlKz1vLnZhbHVlOwogICAgICBtYXBbby5jb29yZGluYXRvcl0uYWdlcy5wdXNoKC4uLkFycmF5KG8uY29udGFpbmVycykuZmlsbChvLmF2Z19hZ2UpKTsKICAgIH0pOwogICAgcm93cyA9IE9iamVjdC52YWx1ZXMobWFwKS5tYXAocj0+KHsuLi5yLAogICAgICBvdmVyOTBfcGN0OiByLnZhbHVlID8gci5vdmVyOTBfdmFsdWUvci52YWx1ZSA6IDAsCiAgICAgIGF2Z19hZ2U6IHIuYWdlcy5sZW5ndGggPyByLmFnZXMucmVkdWNlKChhLGIpPT5hK2IsMCkvci5hZ2VzLmxlbmd0aCA6IDAsCiAgICB9KSkuc29ydCgoYSxiKT0+YS5jb29yZGluYXRvci5sb2NhbGVDb21wYXJlKGIuY29vcmRpbmF0b3IpKTsKICB9CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nvb3JkTWV0YScpLnRleHRDb250ZW50ID0gcm93cy5sZW5ndGggKyAnIGNvb3JkaW5hdG9ycyc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nvb3JkQm9keScpLmlubmVySFRNTCA9IHJvd3MubWFwKHI9PmAKICAgIDx0cj4KICAgICAgPHRkPiR7ci5jb29yZGluYXRvcn08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXROKHIuY29udGFpbmVycyl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10JChyLnZhbHVlKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXQkKHIub3ZlcjkwX3ZhbHVlKX08L3RkPgogICAgICA8dGQ+PGRpdiBjbGFzcz0iYmFyLWNlbGwiPjxkaXYgY2xhc3M9ImJhciI+PGRpdiBjbGFzcz0iYmFyLWZpbGwgcmVkIiBzdHlsZT0id2lkdGg6JHtNYXRoLm1pbihyLm92ZXI5MF9wY3QqMTAwLDEwMCkudG9GaXhlZCgxKX0lIj48L2Rpdj48L2Rpdj48c3Bhbj4ke2ZtdFBjdChyLm92ZXI5MF9wY3QpfTwvc3Bhbj48L2Rpdj48L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXRGKHIuYXZnX2FnZSl9PC90ZD4KICAgIDwvdHI+YCkuam9pbignJyk7CiAgY29uc3QgdG90ID0gcm93cy5yZWR1Y2UoKGEscik9Pih7Y29udGFpbmVyczphLmNvbnRhaW5lcnMrci5jb250YWluZXJzLHZhbHVlOmEudmFsdWUrci52YWx1ZSxvdmVyOTBfdmFsdWU6YS5vdmVyOTBfdmFsdWUrci5vdmVyOTBfdmFsdWV9KSx7Y29udGFpbmVyczowLHZhbHVlOjAsb3ZlcjkwX3ZhbHVlOjB9KTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY29vcmRGb290JykuaW5uZXJIVE1MID0gYDx0cj48dGQ+VE9UQUw8L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4odG90LmNvbnRhaW5lcnMpfTwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10JCh0b3QudmFsdWUpfTwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10JCh0b3Qub3ZlcjkwX3ZhbHVlKX08L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke3RvdC52YWx1ZT9mbXRQY3QodG90Lm92ZXI5MF92YWx1ZS90b3QudmFsdWUpOifigJQnfTwvdGQ+PHRkPjwvdGQ+PC90cj5gOwp9CgovLyDilIDilIAgQWdlIGJ1Y2tldCB0YWJsZSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyQWdlVGFibGUoKSB7CiAgY29uc3Qge3Byb2plY3QsIGNvb3JkLCBidWNrZXR9ID0gZ2V0RmlsdGVycygpOwogIGxldCByb3dzOwogIGlmICghcHJvamVjdCAmJiAhY29vcmQgJiYgIWJ1Y2tldCkgewogICAgcm93cyA9IFJBVy5hZ2VfdGFibGU7CiAgfSBlbHNlIHsKICAgIGNvbnN0IG1hcCA9IHt9OyBBR0VfT1JERVIuZm9yRWFjaChrPT5tYXBba109e2J1Y2tldDprLGxhYmVsOmssY29udGFpbmVyczowLHZhbHVlOjB9KTsKICAgIGZpbHRlcmVkT3JkZXJzLmZvckVhY2gobz0+e2lmKG1hcFtvLmFnZV9idWNrZXRdKXttYXBbby5hZ2VfYnVja2V0XS5jb250YWluZXJzKz1vLmNvbnRhaW5lcnM7bWFwW28uYWdlX2J1Y2tldF0udmFsdWUrPW8udmFsdWU7fX0pOwogICAgY29uc3QgdHY9T2JqZWN0LnZhbHVlcyhtYXApLnJlZHVjZSgoYSxyKT0+YStyLnZhbHVlLDApOwogICAgcm93cz1BR0VfT1JERVIubWFwKGs9Pih7Li4ubWFwW2tdLHBjdDp0dj9tYXBba10udmFsdWUvdHY6MH0pKTsKICB9CiAgY29uc3QgdHY9cm93cy5yZWR1Y2UoKGEscik9PmErci52YWx1ZSwwKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYWdlQm9keScpLmlubmVySFRNTD1yb3dzLm1hcChyPT5gCiAgICA8dHI+CiAgICAgIDx0ZD4ke3IubGFiZWx8fEFHRV9MQUJFTFNbci5idWNrZXRdfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4oci5jb250YWluZXJzKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXQkKHIudmFsdWUpfTwvdGQ+CiAgICAgIDx0ZD48ZGl2IGNsYXNzPSJiYXItY2VsbCI+PGRpdiBjbGFzcz0iYmFyIj48ZGl2IGNsYXNzPSJiYXItZmlsbCIgc3R5bGU9IndpZHRoOiR7TWF0aC5taW4oKHIucGN0fHwwKSoxMDAsMTAwKS50b0ZpeGVkKDEpfSUiPjwvZGl2PjwvZGl2PjxzcGFuPiR7Zm10UGN0KHIucGN0KX08L3NwYW4+PC9kaXY+PC90ZD4KICAgIDwvdHI+YCkuam9pbignJyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FnZUZvb3QnKS5pbm5lckhUTUw9YDx0cj48dGQ+VE9UQUw8L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4ocm93cy5yZWR1Y2UoKGEscik9PmErci5jb250YWluZXJzLDApKX08L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQodHYpfTwvdGQ+PHRkPjwvdGQ+PC90cj5gOwp9CgovLyDilIDilIAgT3JkZXJzIHRhYmxlIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiByZW5kZXJPcmRlcnMoKSB7CiAgY29uc3QgcSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvcmRlclNlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBmaWx0ZXJlZE9yZGVyczsKICBpZiAocSkgcm93cz1yb3dzLmZpbHRlcihvPT5vLm9yZGVyLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSl8fG8ucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpfHxvLmNvb3JkaW5hdG9yLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvcmRlcnNNZXRhJykudGV4dENvbnRlbnQgPSByb3dzLmxlbmd0aCsnIG9yZGVycyc7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29yZGVyc0JvZHknKS5pbm5lckhUTUwgPSByb3dzLm1hcChvPT5gCiAgICA8dHI+CiAgICAgIDx0ZD48c3Ryb25nPiR7by5vcmRlcn08L3N0cm9uZz48L3RkPgogICAgICA8dGQgY2xhc3M9IndyYXAiPiR7by5wcm9qZWN0fTwvdGQ+CiAgICAgIDx0ZD4ke28uY29vcmRpbmF0b3J9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10TihvLmNvbnRhaW5lcnMpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQoby52YWx1ZSl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10RihvLmF2Z19hZ2UpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke21heEFnZVRhZyhvLm1heF9hZ2UpfTwvdGQ+CiAgICAgIDx0ZD4ke2FnZUJ1Y2tldFRhZyhvLmFnZV9idWNrZXQpfTwvdGQ+CiAgICAgIDx0ZD4ke28uc2hpcF9yYW5nZXx8J+KAlCd9PC90ZD4KICAgIDwvdHI+YCkuam9pbignJykgfHwgJzx0cj48dGQgY29sc3Bhbj0iOSIgc3R5bGU9InRleHQtYWxpZ246Y2VudGVyO3BhZGRpbmc6MzJweDtjb2xvcjp2YXIoLS1tdXRlZCkiPk5vIG9yZGVycyBtYXRjaCBmaWx0ZXJzLjwvdGQ+PC90cj4nOwogIGlmICghcm93cy5sZW5ndGgpIHsgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29yZGVyc0Zvb3QnKS5pbm5lckhUTUw9Jyc7IHJldHVybjsgfQogIGNvbnN0IHRvdD1yb3dzLnJlZHVjZSgoYSxvKT0+KHtjb250YWluZXJzOmEuY29udGFpbmVycytvLmNvbnRhaW5lcnMsdmFsdWU6YS52YWx1ZStvLnZhbHVlfSkse2NvbnRhaW5lcnM6MCx2YWx1ZTowfSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29yZGVyc0Zvb3QnKS5pbm5lckhUTUw9YDx0cj48dGQgY29sc3Bhbj0iMyI+VE9UQUwg4oCUICR7cm93cy5sZW5ndGh9IG9yZGVyczwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10Tih0b3QuY29udGFpbmVycyl9PC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXQkKHRvdC52YWx1ZSl9PC90ZD48dGQgY29sc3Bhbj0iNCI+PC90ZD48L3RyPmA7Cn0KCi8vIOKUgOKUgCBTLURyb3AgS1BJcyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyU2Ryb3BLcGlzKCkgewogIGNvbnN0IGsgPSBSQVcuc2Ryb3A/LmtwaXM7CiAgaWYgKCFrKSByZXR1cm47CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NkS3BpSXRlbXMnKS50ZXh0Q29udGVudCAgPSBmbXROKGsudG90YWxfaXRlbXMpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZEtwaVZhbHVlJykudGV4dENvbnRlbnQgID0gZm10JChrLnRvdGFsX3ZhbHVlKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2RLcGlPcmRlcnMnKS50ZXh0Q29udGVudCA9IGZtdE4oay51bmlxdWVfb3JkZXJzKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2RLcGlBdmcnKS50ZXh0Q29udGVudCAgICA9IGZtdEYoay5hdmdfYWdlKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2RLcGlNYXgnKS50ZXh0Q29udGVudCAgICA9IGZtdE4oay5tYXhfYWdlKTsKfQoKLy8g4pSA4pSAIFMtRHJvcCBsb2NhdGlvbiBjYXJkcyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyU2Ryb3BMb2NHcmlkKCkgewogIGNvbnN0IGxvY3MgPSBSQVcuc2Ryb3A/LmJ5X2xvY2F0aW9uIHx8IFtdOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZHJvcExvY0dyaWQnKS5pbm5lckhUTUwgPSBsb2NzLm1hcChsPT5gCiAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9InNkcm9wLWxvYy1uYW1lIj4ke2wubG9jYXRpb259PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNkcm9wLWxvYy1pdGVtcyI+JHtmbXROKGwuaXRlbXMpfSA8c3BhbiBzdHlsZT0iZm9udC1zaXplOjEzcHg7Zm9udC13ZWlnaHQ6NDAwO2NvbG9yOnZhcigtLW11dGVkKSI+aXRlbXM8L3NwYW4+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNkcm9wLWxvYy12YWwiPiR7Zm10JChsLnZhbHVlKX0gwrcgYXZnICR7Zm10RihsLmF2Z19hZ2UpfWQgwrcgbWF4ICR7Zm10TihsLm1heF9hZ2UpfWQ8L2Rpdj4KICAgIDwvZGl2PmApLmpvaW4oJycpOwp9CgovLyDilIDilIAgUy1Ecm9wIGRldGFpbCB0YWJsZSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVzZXRTZHJvcEZpbHRlcnMoKSB7CiAgbXNSZXNldCgnbXMtZHJvcGxvYycpOyBtc1Jlc2V0KCdtcy1kcm9wY29vcmQnKTsgbXNSZXNldCgnbXMtZHJvcGZsYWcnKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BTZWFyY2gnKS52YWx1ZSA9ICcnOwogIHJlbmRlclNkcm9wKCk7Cn0KCmZ1bmN0aW9uIHJlbmRlclNkcm9wKCkgewogIGNvbnN0IGl0ZW1zID0gUkFXLnNkcm9wPy5pdGVtcyB8fCBbXTsKICBjb25zdCBsb2MgICA9IG1zR2V0U2VsZWN0ZWQoJ21zLWRyb3Bsb2MnKTsKICBjb25zdCBjb29yZCA9IG1zR2V0U2VsZWN0ZWQoJ21zLWRyb3Bjb29yZCcpOwogIGNvbnN0IGZsYWcgID0gbXNHZXRTZWxlY3RlZCgnbXMtZHJvcGZsYWcnKTsKICBjb25zdCBxICAgICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZHJvcFNlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CgogIGxldCByb3dzID0gaXRlbXM7CiAgaWYgKGxvYy5sZW5ndGgpICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gbG9jLmluY2x1ZGVzKHIubG9jYXRpb24pKTsKICBpZiAoY29vcmQubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKGZsYWcubGVuZ3RoKSAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gZmxhZy5pbmNsdWRlcyhzZHJvcEFnZUZsYWcoci5hZ2UpKSk7CiAgaWYgKHEpICAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+CiAgICByLm9yZGVyLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwKICAgIHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICByLnBhcnRfbm8udG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fAogICAgci5jb29yZGluYXRvci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpCiAgKTsKCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Nkcm9wTWV0YScpLnRleHRDb250ZW50ID0gcm93cy5sZW5ndGggKyAnIGl0ZW1zJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BCb2R5JykuaW5uZXJIVE1MID0gcm93cy5tYXAocj0+YAogICAgPHRyPgogICAgICA8dGQ+JHtyLmxvY2F0aW9ufTwvdGQ+CiAgICAgIDx0ZD48c3Ryb25nPiR7ci5vcmRlcn08L3N0cm9uZz48L3RkPgogICAgICA8dGQgY2xhc3M9IndyYXAiPiR7ci5wcm9qZWN0fTwvdGQ+CiAgICAgIDx0ZD4ke3IuY29vcmRpbmF0b3J9PC90ZD4KICAgICAgPHRkPiR7ci5wYXJ0X25vfHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtyLnNlcmlhbHx8J+KAlCd9PC90ZD4KICAgICAgPHRkPiR7ci5wYXJ0X2dyb3VwfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4oci5xdHkpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke21heEFnZVRhZyhyLmFnZSl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10JChyLnZhbHVlKX08L3RkPgogICAgICA8dGQ+JHtyLnNoaXBfZGF0ZXx8J+KAlCd9PC90ZD4KICAgICAgPHRkPiR7ci5zdGF0dXN9PC90ZD4KICAgICAgPHRkPiR7c2Ryb3BGbGFnVGFnKHIuYWdlKX08L3RkPgogICAgPC90cj5gKS5qb2luKCcnKSB8fCAnPHRyPjx0ZCBjb2xzcGFuPSIxMSIgc3R5bGU9InRleHQtYWxpZ246Y2VudGVyO3BhZGRpbmc6MzJweDtjb2xvcjp2YXIoLS1tdXRlZCkiPk5vIGl0ZW1zIG1hdGNoIGZpbHRlcnMuPC90ZD48L3RyPic7CgogIGlmICghcm93cy5sZW5ndGgpIHsgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3Nkcm9wRm9vdCcpLmlubmVySFRNTD0nJzsgcmV0dXJuOyB9CiAgY29uc3QgdG90UXR5PXJvd3MucmVkdWNlKChhLHIpPT5hK3IucXR5LDApLCB0b3RWYWw9cm93cy5yZWR1Y2UoKGEscik9PmErci52YWx1ZSwwKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Ryb3BGb290JykuaW5uZXJIVE1MPWA8dHI+PHRkIGNvbHNwYW49IjYiPlRPVEFMIOKAlCAke3Jvd3MubGVuZ3RofSBpdGVtczwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10Tih0b3RRdHkpfTwvdGQ+PHRkPjwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10JCh0b3RWYWwpfTwvdGQ+PHRkIGNvbHNwYW49IjIiPjwvdGQ+PC90cj5gOwp9CgovLyDilIDilIAgU29ydGluZyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gc29ydFRhYmxlKHRhYmxlSWQsIGNvbElkeCkgewogIGNvbnN0IGtleSA9IHRhYmxlSWQrJ18nK2NvbElkeDsKICBjb25zdCBhc2MgPSBzb3J0U3RhdGVba2V5XSAhPT0gJ2FzYyc7CiAgc29ydFN0YXRlW2tleV0gPSBhc2MgPyAnYXNjJyA6ICdkZXNjJzsKICBjb25zdCB0YWJsZSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKHRhYmxlSWQpOwogIHRhYmxlLnF1ZXJ5U2VsZWN0b3JBbGwoJ3RoZWFkIHRoJykuZm9yRWFjaCgodGgsaSk9PnsKICAgIHRoLmNsYXNzTGlzdC5yZW1vdmUoJ3NvcnQtYXNjJywnc29ydC1kZXNjJyk7CiAgICBpZihpPT09Y29sSWR4KSB0aC5jbGFzc0xpc3QuYWRkKGFzYz8nc29ydC1hc2MnOidzb3J0LWRlc2MnKTsKICB9KTsKICBjb25zdCB0Ym9keSA9IHRhYmxlLnF1ZXJ5U2VsZWN0b3IoJ3Rib2R5Jyk7CiAgQXJyYXkuZnJvbSh0Ym9keS5xdWVyeVNlbGVjdG9yQWxsKCd0cicpKS5zb3J0KChhLGIpPT57CiAgICBsZXQgYXY9KGEuY2VsbHNbY29sSWR4XT8udGV4dENvbnRlbnR8fCcnKS5yZXBsYWNlKC9bJCwlXFxz4oaR4oaTXS9nLCcnKS50cmltKCk7CiAgICBsZXQgYnY9KGIuY2VsbHNbY29sSWR4XT8udGV4dENvbnRlbnR8fCcnKS5yZXBsYWNlKC9bJCwlXFxz4oaR4oaTXS9nLCcnKS50cmltKCk7CiAgICBjb25zdCBhbj1wYXJzZUZsb2F0KGF2KSxibj1wYXJzZUZsb2F0KGJ2KTsKICAgIGNvbnN0IGNtcD0oIWlzTmFOKGFuKSYmIWlzTmFOKGJuKSk/YW4tYm46YXYubG9jYWxlQ29tcGFyZShidik7CiAgICByZXR1cm4gYXNjP2NtcDotY21wOwogIH0pLmZvckVhY2gocj0+dGJvZHkuYXBwZW5kQ2hpbGQocikpOwp9CgovLyDilIDilIAgTXVsdGktc2VsZWN0IGRyb3Bkb3duIGVuZ2luZSDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gbXNUb2dnbGUoaWQpIHsKICBjb25zdCB3cmFwID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQpOwogIGNvbnN0IHBhbmVsPSB3cmFwLnF1ZXJ5U2VsZWN0b3IoJy5tcy1wYW5lbCcpOwogIGNvbnN0IGlzT3BlbiA9IHBhbmVsLmNsYXNzTGlzdC5jb250YWlucygnb3BlbicpOwogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5tcy1wYW5lbC5vcGVuJykuZm9yRWFjaChwID0+IHsKICAgIHAuY2xhc3NMaXN0LnJlbW92ZSgnb3BlbicpOwogICAgcC5jbG9zZXN0KCcubXMtd3JhcCcpLnF1ZXJ5U2VsZWN0b3IoJy5tcy1idG4nKS5jbGFzc0xpc3QucmVtb3ZlKCdvcGVuJyk7CiAgfSk7CiAgaWYgKCFpc09wZW4pIHsgcGFuZWwuY2xhc3NMaXN0LmFkZCgnb3BlbicpOyB3cmFwLnF1ZXJ5U2VsZWN0b3IoJy5tcy1idG4nKS5jbGFzc0xpc3QuYWRkKCdvcGVuJyk7IH0KfQovLyBVc2UgZXZlbnQgZGVsZWdhdGlvbiDigJQgY2F0Y2hlcyBkeW5hbWljYWxseSBhZGRlZCBidXR0b25zIHRvbwpkb2N1bWVudC5hZGRFdmVudExpc3RlbmVyKCdjbGljaycsIGUgPT4gewogIGNvbnN0IGJ0biA9IGUudGFyZ2V0LmNsb3Nlc3QoJy5tcy1idG4nKTsKICBpZiAoYnRuKSB7CiAgICBlLnN0b3BQcm9wYWdhdGlvbigpOwogICAgY29uc3Qgd3JhcCA9IGJ0bi5jbG9zZXN0KCcubXMtd3JhcCcpOwogICAgaWYgKHdyYXApIG1zVG9nZ2xlKHdyYXAuaWQpOwogICAgcmV0dXJuOwogIH0KICAvLyBDbGljayBvdXRzaWRlIOKAlCBjbG9zZSBhbGwKICBpZiAoIWUudGFyZ2V0LmNsb3Nlc3QoJy5tcy13cmFwJykpIHsKICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5tcy1wYW5lbC5vcGVuJykuZm9yRWFjaChwID0+IHsKICAgICAgcC5jbGFzc0xpc3QucmVtb3ZlKCdvcGVuJyk7CiAgICAgIHAuY2xvc2VzdCgnLm1zLXdyYXAnKS5xdWVyeVNlbGVjdG9yKCcubXMtYnRuJykuY2xhc3NMaXN0LnJlbW92ZSgnb3BlbicpOwogICAgfSk7CiAgfQp9KTsKCmZ1bmN0aW9uIG1zUG9wdWxhdGUoaWQsIG9wdGlvbnMsIG9uQ2hhbmdlKSB7CiAgY29uc3QgcGFuZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChpZCArICctcGFuZWwnKTsKICBwYW5lbC5pbm5lckhUTUwgPSAnJzsKCiAgLy8gSGVscGVyIHRvIGJ1aWxkIGEgY2hlY2tib3ggcm93IHVzaW5nIERPTSBtZXRob2RzIChhdm9pZHMgaW5uZXJIVE1MIGxhYmVsIGlzc3VlcykKICBmdW5jdGlvbiBtYWtlUm93KHZhbHVlLCBsYWJlbFRleHQsIGlzQm9sZCkgewogICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7CiAgICByb3cuY2xhc3NOYW1lID0gJ21zLWl0ZW0nOwogICAgY29uc3QgY2IgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCdpbnB1dCcpOwogICAgY2IudHlwZSA9ICdjaGVja2JveCc7IGNiLmNoZWNrZWQgPSB0cnVlOwogICAgaWYgKHZhbHVlICE9PSBudWxsKSBjYi52YWx1ZSA9IHZhbHVlOwogICAgZWxzZSBjYi5pZCA9IGlkICsgJy1hbGwnOwogICAgY29uc3QgbGJsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpOwogICAgbGJsLnRleHRDb250ZW50ID0gbGFiZWxUZXh0OwogICAgbGJsLnN0eWxlLmNzc1RleHQgPSAnY3Vyc29yOnBvaW50ZXI7Y29sb3I6IzFGMzg2NDtmb250LXNpemU6MTNweDt1c2VyLXNlbGVjdDpub25lOycgKyAoaXNCb2xkID8gJ2ZvbnQtd2VpZ2h0OjYwMDsnIDogJycpOwogICAgcm93LmFwcGVuZENoaWxkKGNiKTsgcm93LmFwcGVuZENoaWxkKGxibCk7CiAgICByZXR1cm4geyByb3csIGNiIH07CiAgfQoKICAvLyAiQWxsIiByb3cKICBjb25zdCB7IHJvdzogYWxsUm93LCBjYjogYWxsQ2IgfSA9IG1ha2VSb3cobnVsbCwgJ0FsbCcsIHRydWUpOwogIHBhbmVsLmFwcGVuZENoaWxkKGFsbFJvdyk7CiAgY29uc3QgZGl2ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnZGl2Jyk7IGRpdi5jbGFzc05hbWUgPSAnbXMtZGl2aWRlcic7IHBhbmVsLmFwcGVuZENoaWxkKGRpdik7CgogIC8vIE9wdGlvbiByb3dzCiAgb3B0aW9ucy5mb3JFYWNoKG9wdCA9PiB7CiAgICBjb25zdCB7IHJvdywgY2IgfSA9IG1ha2VSb3cob3B0LCBvcHQsIGZhbHNlKTsKICAgIHBhbmVsLmFwcGVuZENoaWxkKHJvdyk7CiAgfSk7CgogIC8vIEV2ZW50IGxpc3RlbmVycwogIHBhbmVsLnF1ZXJ5U2VsZWN0b3JBbGwoJ2lucHV0W3ZhbHVlXScpLmZvckVhY2goY2IgPT4gewogICAgY2IuYWRkRXZlbnRMaXN0ZW5lcignY2hhbmdlJywgKCkgPT4gewogICAgICBhbGxDYi5jaGVja2VkID0gWy4uLnBhbmVsLnF1ZXJ5U2VsZWN0b3JBbGwoJ2lucHV0W3ZhbHVlXScpXS5ldmVyeShjID0+IGMuY2hlY2tlZCk7CiAgICAgIG1zVXBkYXRlTGFiZWwoaWQsIG9wdGlvbnMubGVuZ3RoKTsKICAgICAgb25DaGFuZ2UoKTsKICAgIH0pOwogIH0pOwogIGFsbENiLmFkZEV2ZW50TGlzdGVuZXIoJ2NoYW5nZScsICgpID0+IHsKICAgIHBhbmVsLnF1ZXJ5U2VsZWN0b3JBbGwoJ2lucHV0W3ZhbHVlXScpLmZvckVhY2goY2IgPT4gY2IuY2hlY2tlZCA9IGFsbENiLmNoZWNrZWQpOwogICAgbXNVcGRhdGVMYWJlbChpZCwgb3B0aW9ucy5sZW5ndGgpOwogICAgb25DaGFuZ2UoKTsKICB9KTsKfQoKZnVuY3Rpb24gbXNHZXRTZWxlY3RlZChpZCkgewogIGNvbnN0IHBhbmVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQgKyAnLXBhbmVsJyk7CiAgaWYgKCFwYW5lbCkgcmV0dXJuIFtdOwogIGNvbnN0IGNoZWNrZWQgPSBbLi4ucGFuZWwucXVlcnlTZWxlY3RvckFsbCgnaW5wdXRbdmFsdWVdOmNoZWNrZWQnKV0ubWFwKGMgPT4gYy52YWx1ZSk7CiAgY29uc3QgYWxsICAgICA9IFsuLi5wYW5lbC5xdWVyeVNlbGVjdG9yQWxsKCdpbnB1dFt2YWx1ZV0nKV0ubWFwKGMgPT4gYy52YWx1ZSk7CiAgLy8gSWYgYWxsIGNoZWNrZWQsIHJldHVybiBlbXB0eSBhcnJheSAobWVhbnMgIm5vIGZpbHRlciIpCiAgcmV0dXJuIGNoZWNrZWQubGVuZ3RoID09PSBhbGwubGVuZ3RoID8gW10gOiBjaGVja2VkOwp9CgpmdW5jdGlvbiBtc1VwZGF0ZUxhYmVsKGlkLCB0b3RhbCkgewogIGNvbnN0IHBhbmVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoaWQgKyAnLXBhbmVsJyk7CiAgY29uc3QgYnRuICAgPSBkb2N1bWVudC5xdWVyeVNlbGVjdG9yKCcjJyArIGlkICsgJyAubXMtbGFiZWwnKTsKICBpZiAoIXBhbmVsIHx8ICFidG4pIHJldHVybjsKICBjb25zdCBjaGVja2VkID0gWy4uLnBhbmVsLnF1ZXJ5U2VsZWN0b3JBbGwoJ2lucHV0W3ZhbHVlXTpjaGVja2VkJyldOwogIGNvbnN0IGFsbCAgICAgPSBbLi4ucGFuZWwucXVlcnlTZWxlY3RvckFsbCgnaW5wdXRbdmFsdWVdJyldOwogIGlmIChjaGVja2VkLmxlbmd0aCA9PT0gMCkgICAgICAgICBidG4udGV4dENvbnRlbnQgPSAnTm9uZSBzZWxlY3RlZCc7CiAgZWxzZSBpZiAoY2hlY2tlZC5sZW5ndGggPT09IGFsbC5sZW5ndGgpIGJ0bi50ZXh0Q29udGVudCA9ICdBbGwnOwogIGVsc2UgaWYgKGNoZWNrZWQubGVuZ3RoID09PSAxKSAgICBidG4udGV4dENvbnRlbnQgPSBjaGVja2VkWzBdLnZhbHVlOwogIGVsc2UgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBidG4udGV4dENvbnRlbnQgPSBjaGVja2VkLmxlbmd0aCArICcgc2VsZWN0ZWQnOwp9CgpmdW5jdGlvbiBtc1Jlc2V0KGlkKSB7CiAgY29uc3QgcGFuZWwgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZChpZCArICctcGFuZWwnKTsKICBpZiAoIXBhbmVsKSByZXR1cm47CiAgcGFuZWwucXVlcnlTZWxlY3RvckFsbCgnaW5wdXQnKS5mb3JFYWNoKGNiID0+IGNiLmNoZWNrZWQgPSB0cnVlKTsKICBtc1VwZGF0ZUxhYmVsKGlkLCAwKTsKfQoKLy8g4pSA4pSAIENTViBFeHBvcnQg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmZ1bmN0aW9uIHRvQ1NWKGhlYWRlcnMsIHJvd3MpIHsKICBjb25zdCBlc2NhcGUgPSB2ID0+IHsKICAgIGlmICh2ID09IG51bGwpIHJldHVybiAnJzsKICAgIGNvbnN0IHMgPSBTdHJpbmcodik7CiAgICByZXR1cm4gcy5pbmNsdWRlcygnLCcpIHx8IHMuaW5jbHVkZXMoJyInKSB8fCBzLmluY2x1ZGVzKCdcbicpID8gJyInICsgcy5yZXBsYWNlKC8iL2csICciIicpICsgJyInIDogczsKICB9OwogIGNvbnN0IGxpbmVzID0gW2hlYWRlcnMubWFwKGVzY2FwZSkuam9pbignLCcpXTsKICByb3dzLmZvckVhY2gociA9PiBsaW5lcy5wdXNoKHIubWFwKGVzY2FwZSkuam9pbignLCcpKSk7CiAgcmV0dXJuIGxpbmVzLmpvaW4oJ1xuJyk7Cn0KCmZ1bmN0aW9uIGRvd25sb2FkQ1NWKGZpbGVuYW1lLCBjc3YpIHsKICBjb25zdCBibG9iID0gbmV3IEJsb2IoW2Nzdl0sIHt0eXBlOiAndGV4dC9jc3YnfSk7CiAgY29uc3QgdXJsICA9IFVSTC5jcmVhdGVPYmplY3RVUkwoYmxvYik7CiAgY29uc3QgYSAgICA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICBhLmhyZWYgPSB1cmw7IGEuZG93bmxvYWQgPSBmaWxlbmFtZTsgYS5jbGljaygpOwogIFVSTC5yZXZva2VPYmplY3RVUkwodXJsKTsKfQoKZnVuY3Rpb24gZXhwb3J0TWFpbkNTVigpIHsKICBjb25zdCBoZWFkZXJzID0gWydPcmRlciAjJywnUHJvamVjdCcsJ0Nvb3JkaW5hdG9yJywnQ29udGFpbmVycycsJ1ZhbHVlJywnQXZnIEFnZScsJ01heCBBZ2UnLCdBZ2UgQnVja2V0JywnU3RhdHVzJywnU2hpcCBEYXRlIFJhbmdlJ107CiAgY29uc3Qgcm93cyA9IGZpbHRlcmVkT3JkZXJzLm1hcChvID0+IFsKICAgIG8ub3JkZXIsIG8ucHJvamVjdCwgby5jb29yZGluYXRvciwgby5jb250YWluZXJzLAogICAgby52YWx1ZSwgby5hdmdfYWdlLCBvLm1heF9hZ2UsIG8uYWdlX2J1Y2tldCwgby5zdGF0dXMsIG8uc2hpcF9yYW5nZQogIF0pOwogIGNvbnN0IHByb2ogID0gbXNHZXRTZWxlY3RlZCgnbXMtcHJvamVjdCcpOwogIGNvbnN0IGNvb3JkID0gbXNHZXRTZWxlY3RlZCgnbXMtY29vcmQnKTsKICBjb25zdCBidWNrZXQ9IG1zR2V0U2VsZWN0ZWQoJ21zLWJ1Y2tldCcpOwogIGNvbnN0IGxhYmVsID0gWwogICAgcHJvai5sZW5ndGggICA/IHByb2ouam9pbignKycpICAgOiAnJywKICAgIGNvb3JkLmxlbmd0aCAgPyBjb29yZC5qb2luKCcrJykgIDogJycsCiAgICBidWNrZXQubGVuZ3RoID8gYnVja2V0LmpvaW4oJysnKSA6ICcnLAogIF0uZmlsdGVyKEJvb2xlYW4pLmpvaW4oJ18nKSB8fCAnQWxsJzsKICBkb3dubG9hZENTVignQWdlZF9JbnZlbnRvcnlfJyArIGxhYmVsICsgJy5jc3YnLCB0b0NTVihoZWFkZXJzLCByb3dzKSk7Cn0KCmZ1bmN0aW9uIGV4cG9ydFNkcm9wQ1NWKCkgewogIGNvbnN0IGl0ZW1zID0gUkFXLnNkcm9wPy5pdGVtcyB8fCBbXTsKICBjb25zdCBsb2MgICA9IG1zR2V0U2VsZWN0ZWQoJ21zLWRyb3Bsb2MnKTsKICBjb25zdCBjb29yZCA9IG1zR2V0U2VsZWN0ZWQoJ21zLWRyb3Bjb29yZCcpOwogIGNvbnN0IGZsYWcgID0gbXNHZXRTZWxlY3RlZCgnbXMtZHJvcGZsYWcnKTsKICBjb25zdCBxICAgICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzZHJvcFNlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBpdGVtczsKICBpZiAobG9jLmxlbmd0aCkgICByb3dzID0gcm93cy5maWx0ZXIociA9PiBsb2MuaW5jbHVkZXMoci5sb2NhdGlvbikpOwogIGlmIChjb29yZC5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGNvb3JkLmluY2x1ZGVzKHIuY29vcmRpbmF0b3IpKTsKICBpZiAoZmxhZy5sZW5ndGgpICByb3dzID0gcm93cy5maWx0ZXIociA9PiBmbGFnLmluY2x1ZGVzKHNkcm9wQWdlRmxhZyhyLmFnZSkpKTsKICBpZiAocSkgICAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+CiAgICByLm9yZGVyLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwgci5wcm9qZWN0LnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwKICAgIHIucGFydF9uby50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIuY29vcmRpbmF0b3IudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSk7CiAgY29uc3QgaGVhZGVycyA9IFsnTG9jYXRpb24nLCdPcmRlciAjJywnUHJvamVjdCcsJ0Nvb3JkaW5hdG9yJywnUGFydCBOby4nLCdTZXJpYWwgIycsJ1BhcnQgR3JvdXAnLCdRdHknLCdBZ2UgKGRheXMpJywnVmFsdWUnLCdTaGlwIERhdGUnLCdPcmRlciBTdGF0dXMnLCdBZ2UgRmxhZyddOwogIGNvbnN0IGNzdlJvd3MgPSByb3dzLm1hcChyID0+IFtyLmxvY2F0aW9uLCByLm9yZGVyLCByLnByb2plY3QsIHIuY29vcmRpbmF0b3IsIHIucGFydF9ubywgci5zZXJpYWx8fCcnLCByLnBhcnRfZ3JvdXAsIHIucXR5LCByLmFnZSwgci52YWx1ZSwgci5zaGlwX2RhdGV8fCcnLCByLnN0YXR1cywgci5mbGFnXSk7CiAgZG93bmxvYWRDU1YoJ1NEcm9wX1Jldmlldy5jc3YnLCB0b0NTVihoZWFkZXJzLCBjc3ZSb3dzKSk7Cn0KCi8vIOKUgOKUgCBDbG9zZWQgT3JkZXJzIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiByZW5kZXJDbG9zZWRLcGlzKCkgewogIGNvbnN0IGsgPSBSQVcuY2xvc2VkPy5rcGlzOwogIGlmICghaykgcmV0dXJuOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbEtwaUl0ZW1zJykudGV4dENvbnRlbnQgID0gZm10TihrLnRvdGFsX2l0ZW1zKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xLcGlWYWx1ZScpLnRleHRDb250ZW50ICA9IGZtdCQoay50b3RhbF92YWx1ZSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NsS3BpT3JkZXJzJykudGV4dENvbnRlbnQgPSBmbXROKGsudW5pcXVlX29yZGVycyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NsS3BpQXZnJykudGV4dENvbnRlbnQgICAgPSBmbXRGKGsuYXZnX2FnZSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2NsS3BpTWF4JykudGV4dENvbnRlbnQgICAgPSBmbXROKGsubWF4X2FnZSk7Cn0KCmZ1bmN0aW9uIHJlc2V0Q2xvc2VkRmlsdGVycygpIHsKICBtc1Jlc2V0KCdtcy1jbGNvb3JkJyk7IG1zUmVzZXQoJ21zLWNsYnVja2V0Jyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nsb3NlZFNlYXJjaCcpLnZhbHVlID0gJyc7CiAgcmVuZGVyQ2xvc2VkKCk7Cn0KCmZ1bmN0aW9uIHJlbmRlckNsb3NlZCgpIHsKICBjb25zdCBpdGVtcyA9IFJBVy5jbG9zZWQ/Lml0ZW1zIHx8IFtdOwogIGNvbnN0IGNvb3JkICA9IG1zR2V0U2VsZWN0ZWQoJ21zLWNsY29vcmQnKTsKICBjb25zdCBidWNrZXQgPSBtc0dldFNlbGVjdGVkKCdtcy1jbGJ1Y2tldCcpOwogIGNvbnN0IHEgICAgICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbG9zZWRTZWFyY2gnKS52YWx1ZS50b0xvd2VyQ2FzZSgpOwogIGxldCByb3dzID0gaXRlbXM7CiAgaWYgKGNvb3JkLmxlbmd0aCkgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGNvb3JkLmluY2x1ZGVzKHIuY29vcmRpbmF0b3IpKTsKICBpZiAoYnVja2V0Lmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gYnVja2V0LmluY2x1ZGVzKHIuYWdlX2J1Y2tldCkpOwogIGlmIChxKSAgICAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+CiAgICByLm9yZGVyLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwgci5wcm9qZWN0LnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwKICAgIHIucGFydF9uby50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xvc2VkTWV0YScpLnRleHRDb250ZW50ID0gcm93cy5sZW5ndGggKyAnIGl0ZW1zJzsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xvc2VkQm9keScpLmlubmVySFRNTCA9IHJvd3MubWFwKHIgPT4gYAogICAgPHRyPgogICAgICA8dGQ+PHN0cm9uZz4ke3Iub3JkZXJ9PC9zdHJvbmc+PC90ZD4KICAgICAgPHRkIGNsYXNzPSJ3cmFwIj4ke3IucHJvamVjdH08L3RkPgogICAgICA8dGQ+JHtyLmNvb3JkaW5hdG9yfTwvdGQ+CiAgICAgIDx0ZD4ke3IucGFydF9ub3x8J+KAlCd9PC90ZD4KICAgICAgPHRkPiR7ci5zZXJpYWx8fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3IucGFydF9ncm91cH08L3RkPgogICAgICA8dGQ+JHtyLmxvY2F0aW9ufTwvdGQ+CiAgICAgIDx0ZD4ke3Iuc2hpcF9kYXRlfHwn4oCUJ308L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXROKHIucXR5KX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHttYXhBZ2VUYWcoci5hZ2UpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdCQoci52YWx1ZSl9PC90ZD4KICAgICAgPHRkPiR7YWdlQnVja2V0VGFnKHIuYWdlX2J1Y2tldCl9PC90ZD4KICAgIDwvdHI+YCkuam9pbignJykgfHwgJzx0cj48dGQgY29sc3Bhbj0iMTAiIHN0eWxlPSJ0ZXh0LWFsaWduOmNlbnRlcjtwYWRkaW5nOjMycHg7Y29sb3I6dmFyKC0tbXV0ZWQpIj5ObyBpdGVtcyBtYXRjaCBmaWx0ZXJzLjwvdGQ+PC90cj4nOwogIGlmICghcm93cy5sZW5ndGgpIHsgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2Nsb3NlZEZvb3QnKS5pbm5lckhUTUwgPSAnJzsgcmV0dXJuOyB9CiAgY29uc3QgdG90ID0gcm93cy5yZWR1Y2UoKGEscikgPT4gKHtxdHk6YS5xdHkrci5xdHksIHZhbHVlOmEudmFsdWUrci52YWx1ZX0pLCB7cXR5OjAsdmFsdWU6MH0pOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdjbG9zZWRGb290JykuaW5uZXJIVE1MID0gYDx0cj48dGQgY29sc3Bhbj0iOSI+VE9UQUwg4oCUICR7cm93cy5sZW5ndGh9IGl0ZW1zPC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXROKHRvdC5xdHkpfTwvdGQ+PHRkPjwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10JCh0b3QudmFsdWUpfTwvdGQ+PHRkPjwvdGQ+PC90cj5gOwp9CgpmdW5jdGlvbiBleHBvcnRDbG9zZWRDU1YoKSB7CiAgY29uc3QgaXRlbXMgPSBSQVcuY2xvc2VkPy5pdGVtcyB8fCBbXTsKICBjb25zdCBjb29yZCAgPSBtc0dldFNlbGVjdGVkKCdtcy1jbGNvb3JkJyk7CiAgY29uc3QgYnVja2V0ID0gbXNHZXRTZWxlY3RlZCgnbXMtY2xidWNrZXQnKTsKICBjb25zdCBxICAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnY2xvc2VkU2VhcmNoJykudmFsdWUudG9Mb3dlckNhc2UoKTsKICBsZXQgcm93cyA9IGl0ZW1zOwogIGlmIChjb29yZC5sZW5ndGgpICByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKGJ1Y2tldC5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGJ1Y2tldC5pbmNsdWRlcyhyLmFnZV9idWNrZXQpKTsKICBpZiAocSkgICAgICByb3dzID0gcm93cy5maWx0ZXIociA9PiByLm9yZGVyLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwgci5wcm9qZWN0LnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkgfHwgci5wYXJ0X25vLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwogIGNvbnN0IGhlYWRlcnMgPSBbJ09yZGVyICMnLCdQcm9qZWN0JywnQ29vcmRpbmF0b3InLCdQYXJ0IE5vLicsJ1NlcmlhbCAjJywnUGFydCBHcm91cCcsJ0xvY2F0aW9uJywnU2hpcCBEYXRlJywnUXR5JywnQWdlIChkYXlzKScsJ1ZhbHVlJywnQWdlIEJ1Y2tldCcsJ09yZGVyIFN0YXR1cyddOwogIGNvbnN0IGNzdlJvd3MgPSByb3dzLm1hcChyID0+IFtyLm9yZGVyLHIucHJvamVjdCxyLmNvb3JkaW5hdG9yLHIucGFydF9ubyxyLnNlcmlhbHx8Jycsci5wYXJ0X2dyb3VwLHIubG9jYXRpb24sci5zaGlwX2RhdGV8fCcnLHIucXR5LHIuYWdlLHIudmFsdWUsci5hZ2VfYnVja2V0LHIuc3RhdHVzXSk7CiAgZG93bmxvYWRDU1YoJ0Nsb3NlZF9PcmRlcnMuY3N2JywgdG9DU1YoaGVhZGVycywgY3N2Um93cykpOwp9CgovLyDilIDilIAgT2Zmc2l0ZSBTdG9yYWdlIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiByZW5kZXJPZmZzaXRlS3BpcygpIHsKICBjb25zdCBrID0gUkFXLm9mZnNpdGU/LmtwaXM7CiAgaWYgKCFrKSByZXR1cm47CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29mS3BpSXRlbXMnKS50ZXh0Q29udGVudCAgPSBmbXROKGsudG90YWxfaXRlbXMpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZktwaVZhbHVlJykudGV4dENvbnRlbnQgID0gZm10JChrLnRvdGFsX3ZhbHVlKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2ZLcGlPcmRlcnMnKS50ZXh0Q29udGVudCA9IGZtdE4oay51bmlxdWVfb3JkZXJzKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2ZLcGlBdmcnKS50ZXh0Q29udGVudCAgICA9IGZtdEYoay5hdmdfYWdlKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2ZLcGlNYXgnKS50ZXh0Q29udGVudCAgICA9IGZtdE4oay5tYXhfYWdlKTsKfQoKZnVuY3Rpb24gcmVuZGVyT2Zmc2l0ZUxvY0dyaWQoKSB7CiAgY29uc3QgZ3JvdXBzID0gUkFXLm9mZnNpdGU/LmJ5X2dyb3VwIHx8IFtdOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZmZzaXRlTG9jR3JpZCcpLmlubmVySFRNTCA9IGdyb3Vwcy5tYXAoZyA9PiBgCiAgICA8ZGl2IGNsYXNzPSJzZHJvcC1sb2MtY2FyZCI+CiAgICAgIDxkaXYgY2xhc3M9InNkcm9wLWxvYy1uYW1lIj4ke2cuZ3JvdXB9PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNkcm9wLWxvYy1pdGVtcyI+JHtmbXROKGcub3JkZXJzKX0gPHNwYW4gc3R5bGU9ImZvbnQtc2l6ZToxM3B4O2ZvbnQtd2VpZ2h0OjQwMDtjb2xvcjp2YXIoLS1tdXRlZCkiPm9yZGVyczwvc3Bhbj4gJm5ic3A7wrcmbmJzcDsgJHtmbXROKGcuaXRlbXMpfSA8c3BhbiBzdHlsZT0iZm9udC1zaXplOjEzcHg7Zm9udC13ZWlnaHQ6NDAwO2NvbG9yOnZhcigtLW11dGVkKSI+Y29udGFpbmVyczwvc3Bhbj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic2Ryb3AtbG9jLXZhbCI+JHtmbXQkKGcudmFsdWUpfSDCtyBhdmcgJHtmbXRGKGcuYXZnX2FnZSl9ZCDCtyBtYXggJHtmbXROKGcubWF4X2FnZSl9ZDwvZGl2PgogICAgPC9kaXY+YCkuam9pbignJyk7Cn0KCmZ1bmN0aW9uIHJlc2V0T2Zmc2l0ZUZpbHRlcnMoKSB7CiAgbXNSZXNldCgnbXMtb2ZidWlsZCcpOyBtc1Jlc2V0KCdtcy1vZmNvb3JkJyk7IG1zUmVzZXQoJ21zLW9mc3RhdHVzJyk7IG1zUmVzZXQoJ21zLW9mZmxhZycpOyBtc1Jlc2V0KCdtcy1vZmRpc3AnKTsgbXNSZXNldCgnbXMtb2ZiaWxsZWQnKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZVNlYXJjaCcpLnZhbHVlID0gJyc7CiAgcmVuZGVyT2Zmc2l0ZSgpOwp9CgpmdW5jdGlvbiByZW5kZXJPZmZzaXRlKCkgewogIGNvbnN0IGl0ZW1zICA9IFJBVy5vZmZzaXRlPy5pdGVtcyB8fCBbXTsKICBjb25zdCBncm91cCAgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmJ1aWxkJyk7CiAgY29uc3QgY29vcmQgID0gbXNHZXRTZWxlY3RlZCgnbXMtb2Zjb29yZCcpOwogIGNvbnN0IHN0YXR1cyA9IG1zR2V0U2VsZWN0ZWQoJ21zLW9mc3RhdHVzJyk7CiAgY29uc3QgZmxhZyAgID0gbXNHZXRTZWxlY3RlZCgnbXMtb2ZmbGFnJyk7CiAgY29uc3QgZGlzcCAgID0gbXNHZXRTZWxlY3RlZCgnbXMtb2ZkaXNwJyk7CiAgY29uc3QgcSAgICAgID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ29mZnNpdGVTZWFyY2gnKS52YWx1ZS50b0xvd2VyQ2FzZSgpOwogIGxldCByb3dzID0gaXRlbXM7CiAgaWYgKGdyb3VwLmxlbmd0aCkgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGdyb3VwLmluY2x1ZGVzKHIuYnVpbGRpbmcpKTsKICBpZiAoY29vcmQubGVuZ3RoKSAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gY29vcmQuaW5jbHVkZXMoci5jb29yZGluYXRvcikpOwogIGlmIChzdGF0dXMubGVuZ3RoKSByb3dzID0gcm93cy5maWx0ZXIociA9PiBzdGF0dXMuaW5jbHVkZXMoci5zdGF0dXMpKTsKICBpZiAoZmxhZy5sZW5ndGgpICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gZmxhZy5pbmNsdWRlcyhyLmZsYWcpKTsKICBpZiAoZGlzcC5sZW5ndGgpICAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gewogICAgY29uc3QgZCA9IERJU1BPU0lUSU9OU1tyLm9yZGVyXSB8fCAnJzsKICAgIHJldHVybiBkaXNwLmluY2x1ZGVzKGQgfHwgJyhOb25lKScpOwogIH0pOwogIGNvbnN0IGJpbGxlZCA9IG1zR2V0U2VsZWN0ZWQoJ21zLW9mYmlsbGVkJyk7CiAgaWYgKGJpbGxlZC5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHsKICAgIGNvbnN0IGIgPSByLmJpbGxlZCA/ICdCaWxsZWQnIDogJ05vdCBCaWxsZWQnOwogICAgcmV0dXJuIGJpbGxlZC5pbmNsdWRlcyhiKTsKICB9KTsKICBpZiAocSkgICAgICByb3dzID0gcm93cy5maWx0ZXIociA9PgogICAgci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICByLnBhcnRfbm8udG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLmNvb3JkaW5hdG9yLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZmZzaXRlTWV0YScpLnRleHRDb250ZW50ID0gcm93cy5sZW5ndGgudG9Mb2NhbGVTdHJpbmcoKSArICcgaXRlbXMnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdvZmZzaXRlQm9keScpLmlubmVySFRNTCA9IHJvd3MubWFwKHIgPT4gewogICAgLy8gUGFyc2Ugc2hpcCBkYXRlIGZvciBjb25kaXRpb25hbCBmb3JtYXR0aW5nCiAgICBjb25zdCBzZCA9IHIuc2hpcF9kYXRlICYmIHIuc2hpcF9kYXRlICE9PSAn4oCUJyA/IG5ldyBEYXRlKHIuc2hpcF9kYXRlKSA6IG51bGw7CiAgICBjb25zdCB0b2RheSA9IG5ldyBEYXRlKCk7IHRvZGF5LnNldEhvdXJzKDAsMCwwLDApOwogICAgY29uc3QgZGF5c091dCA9IHNkID8gTWF0aC5yb3VuZCgoc2QgLSB0b2RheSkgLyA4NjQwMDAwMCkgOiBudWxsOwogICAgY29uc3Qgc3RhdHVzU3RvcmFnZSA9IHIuc3RhdHVzICYmIHIuc3RhdHVzLnRvTG93ZXJDYXNlKCkuc3RhcnRzV2l0aCgnc3RvcmFnZScpOwogICAgY29uc3QgaXNEb3duc3RhaXJzID0gci5idWlsZGluZyA9PT0gJ0Rvd25zdGFpcnMgU3RvcmFnZSc7CiAgICBjb25zdCBpc05vcnRoID0gci5idWlsZGluZyA9PT0gJ05vcnRoIFdhcmVob3VzZSc7CiAgICBjb25zdCBkYXRlRmxhZyA9IChpc0Rvd25zdGFpcnMgJiYgc3RhdHVzU3RvcmFnZSAmJiBkYXlzT3V0ICE9PSBudWxsICYmIGRheXNPdXQgPiA0NSkgfHwKICAgICAgICAgICAgICAgICAgICAgKGlzTm9ydGggICAgICYmIHN0YXR1c1N0b3JhZ2UgJiYgZGF5c091dCAhPT0gbnVsbCAmJiBkYXlzT3V0ID4gNjApOwogICAgY29uc3QgZGlzcCA9IERJU1BPU0lUSU9OU1tyLm9yZGVyfHwnJ10gfHwgJyc7CiAgICAvLyBSZWQgaWY6IHN0b3JhZ2Ugc3RhdHVzLCBub3QgU09QIEJ1aWxkIEFoZWFkLCBhbmQgbm90IGJpbGxlZAogICAgY29uc3QgYmlsbCA9IEJJTExJTkdbci5vcmRlcl0gfHwgbnVsbDsKICAgIGNvbnN0IG5lZWRzQmlsbGluZyA9IHN0YXR1c1N0b3JhZ2UgJiYgZGlzcCAhPT0gJ1NPUCBCdWlsZCBBaGVhZCcgJiYgIWJpbGw7CiAgICBjb25zdCByb3dTdHlsZSA9IG5lZWRzQmlsbGluZwogICAgICA/ICdiYWNrZ3JvdW5kOiNmZmYwZjA7Ym9yZGVyLWxlZnQ6M3B4IHNvbGlkICNkYzM1NDU7JwogICAgICA6IGRhdGVGbGFnID8gJ2JhY2tncm91bmQ6I2ZmZjNjZDtib3JkZXItbGVmdDozcHggc29saWQgI2ZmYzEwNzsnIDogJyc7CiAgICByZXR1cm4gYDx0ciBzdHlsZT0iJHtyb3dTdHlsZX0iPgogICAgICA8dGQ+JHtyLmJ1aWxkaW5nfHwn4oCUJ308L3RkPgogICAgICA8dGQ+JHtyLmxvY2F0aW9uX2dyb3VwfHwn4oCUJ308L3RkPgogICAgICA8dGQ+PHN0cm9uZz4ke3Iub3JkZXJ8fCfigJQnfTwvc3Ryb25nPjwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0id3JhcCI+JHtyLnByb2plY3R8fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3IuY29vcmRpbmF0b3J8fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3IubG9jYXRpb258fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3IucGFydF9ub3x8J+KAlCd9PC90ZD4KICAgICAgPHRkPiR7ci5zZXJpYWx8fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke3Iuc2hpcF9kYXRlfHwn4oCUJ308L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXROKHIuY29udGFpbmVycyl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7Zm10JChyLnZhbHVlKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXRGKHIuYXZnX2FnZSl9PC90ZD4KICAgICAgPHRkIGNsYXNzPSJudW0iPiR7bWF4QWdlVGFnKHIubWF4X2FnZSl9PC90ZD4KICAgICAgPHRkPiR7ci5zdGF0dXN8fCfigJQnfTwvdGQ+CiAgICAgIDx0ZD4ke2ZsYWdUYWcoci5mbGFnKX08L3RkPgogICAgICA8dGQ+JHtiaWxsID8gYmlsbC5jaGFyZ2VfdHlwZSA6ICfigJQnfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2JpbGwgPyBmbXQkKGJpbGwuY2hhcmdlX2Ftb3VudCkgOiBuZWVkc0JpbGxpbmcgPyAnPHNwYW4gc3R5bGU9ImNvbG9yOiNkYzM1NDU7Zm9udC13ZWlnaHQ6NjAwIj7imqAgTm90IEJpbGxlZDwvc3Bhbj4nIDogJ+KAlCd9PC90ZD4KICAgICAgPHRkPiR7YmlsbCA/IGJpbGwubm90ZSA6ICfigJQnfTwvdGQ+CiAgICAgIDx0ZD48c2VsZWN0IGNsYXNzPSJkaXNwLXNlbGVjdCIgZGF0YS1vcmRlcj0iJHtyLm9yZGVyfHwnJ30iIG9uY2hhbmdlPSJzYXZlRGlzcG9zaXRpb24odGhpcykiIHN0eWxlPSJmb250LXNpemU6MTFweDtwYWRkaW5nOjJweCA0cHg7Ym9yZGVyOjFweCBzb2xpZCAjY2NjO2JvcmRlci1yYWRpdXM6NHB4O3dpZHRoOjE0MHB4OyI+CiAgICAgICAgPG9wdGlvbiB2YWx1ZT0iIj7igJQgU2VsZWN0IOKAlDwvb3B0aW9uPgogICAgICAgIDxvcHRpb24gdmFsdWU9IlNPUCBCdWlsZCBBaGVhZCIgJHtkaXNwPT09J1NPUCBCdWlsZCBBaGVhZCc/J3NlbGVjdGVkJzonJ30+U09QIEJ1aWxkIEFoZWFkPC9vcHRpb24+CiAgICAgICAgPG9wdGlvbiB2YWx1ZT0iU3RvcmFnZSBDaGFyZ2VkIiAke2Rpc3A9PT0nU3RvcmFnZSBDaGFyZ2VkJz8nc2VsZWN0ZWQnOicnfT5TdG9yYWdlIENoYXJnZWQ8L29wdGlvbj4KICAgICAgPC9zZWxlY3Q+PC90ZD4KICAgIDwvdHI+YDsKICB9KS5qb2luKCcnKSB8fCAnPHRyPjx0ZCBjb2xzcGFuPSIxMCIgc3R5bGU9InRleHQtYWxpZ246Y2VudGVyO3BhZGRpbmc6MzJweDtjb2xvcjp2YXIoLS1tdXRlZCkiPk5vIGl0ZW1zIG1hdGNoIGZpbHRlcnMuPC90ZD48L3RyPic7CiAgaWYgKCFyb3dzLmxlbmd0aCkgeyBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUZvb3QnKS5pbm5lckhUTUwgPSAnJzsgcmV0dXJuOyB9CiAgY29uc3QgdG90ID0gcm93cy5yZWR1Y2UoKGEscikgPT4gKHtjb250YWluZXJzOmEuY29udGFpbmVycytyLmNvbnRhaW5lcnMsIHZhbHVlOmEudmFsdWUrci52YWx1ZX0pLCB7Y29udGFpbmVyczowLHZhbHVlOjB9KTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZUZvb3QnKS5pbm5lckhUTUwgPSBgPHRyPjx0ZCBjb2xzcGFuPSIxMiI+VE9UQUwg4oCUICR7cm93cy5sZW5ndGgudG9Mb2NhbGVTdHJpbmcoKX0gcm93czwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10Tih0b3QuY29udGFpbmVycyl9PC90ZD48dGQgY2xhc3M9Im51bSI+JHtmbXQkKHRvdC52YWx1ZSl9PC90ZD48dGQgY29sc3Bhbj0iNCI+PC90ZD48L3RyPmA7Cn0KCmZ1bmN0aW9uIGV4cG9ydE9mZnNpdGVDU1YoKSB7CiAgY29uc3QgaXRlbXMgID0gUkFXLm9mZnNpdGU/Lml0ZW1zIHx8IFtdOwogIGNvbnN0IGdyb3VwICA9IG1zR2V0U2VsZWN0ZWQoJ21zLW9mYnVpbGQnKTsKICBjb25zdCBjb29yZCAgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmNvb3JkJyk7CiAgY29uc3Qgc3RhdHVzID0gbXNHZXRTZWxlY3RlZCgnbXMtb2ZzdGF0dXMnKTsKICBjb25zdCBmbGFnICAgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmZsYWcnKTsKICBjb25zdCBkaXNwICAgPSBtc0dldFNlbGVjdGVkKCdtcy1vZmRpc3AnKTsKICBjb25zdCBxICAgICAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnb2Zmc2l0ZVNlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBpdGVtczsKICBpZiAoZ3JvdXAubGVuZ3RoKSAgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gZ3JvdXAuaW5jbHVkZXMoci5idWlsZGluZykpOwogIGlmIChjb29yZC5sZW5ndGgpICByb3dzID0gcm93cy5maWx0ZXIociA9PiBjb29yZC5pbmNsdWRlcyhyLmNvb3JkaW5hdG9yKSk7CiAgaWYgKHN0YXR1cy5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHN0YXR1cy5pbmNsdWRlcyhyLnN0YXR1cykpOwogIGlmIChmbGFnLmxlbmd0aCkgICByb3dzID0gcm93cy5maWx0ZXIociA9PiBmbGFnLmluY2x1ZGVzKHIuZmxhZykpOwogIGlmIChkaXNwLmxlbmd0aCkgICByb3dzID0gcm93cy5maWx0ZXIociA9PiB7CiAgICBjb25zdCBkID0gRElTUE9TSVRJT05TW3Iub3JkZXJdIHx8ICcnOwogICAgcmV0dXJuIGRpc3AuaW5jbHVkZXMoZCB8fCAnKE5vbmUpJyk7CiAgfSk7CiAgY29uc3QgYmlsbGVkID0gbXNHZXRTZWxlY3RlZCgnbXMtb2ZiaWxsZWQnKTsKICBpZiAoYmlsbGVkLmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gewogICAgY29uc3QgYiA9IHIuYmlsbGVkID8gJ0JpbGxlZCcgOiAnTm90IEJpbGxlZCc7CiAgICByZXR1cm4gYmlsbGVkLmluY2x1ZGVzKGIpOwogIH0pOwogIGlmIChxKSAgICAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHIub3JkZXIudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLnByb2plY3QudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLnBhcnRfbm8udG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSk7CiAgY29uc3QgaGVhZGVycyA9IFsnQnVpbGRpbmcnLCdMb2NhdGlvbiBHcm91cCcsJ09yZGVyICMnLCdQcm9qZWN0JywnQ29vcmRpbmF0b3InLCdMb2NhdGlvbicsJ1BhcnQgTm8uJywnU2VyaWFsICMnLCdTaGlwIERhdGUnLCdDb250YWluZXJzJywnVmFsdWUnLCdBdmcgQWdlJywnTWF4IEFnZScsJ09yZGVyIFN0YXR1cycsJ0FnZSBGbGFnJywnQ2hhcmdlIFR5cGUnLCdDaGFyZ2UgQW1vdW50JywnQmlsbGluZyBOb3RlJywnRGlzcG9zaXRpb24nLCdVbmJpbGxlZD8nXTsKICBjb25zdCBjc3ZSb3dzID0gcm93cy5tYXAociA9PiB7IGNvbnN0IGI9QklMTElOR1tyLm9yZGVyXXx8bnVsbDsgY29uc3QgZD1ESVNQT1NJVElPTlNbci5vcmRlcl18fCcnOyBjb25zdCBuYj1yLnN0YXR1cyYmci5zdGF0dXMudG9Mb3dlckNhc2UoKS5zdGFydHNXaXRoKCdzdG9yYWdlJykmJmQhPT0nU09QIEJ1aWxkIEFoZWFkJyYmIWI7IHJldHVybiBbci5idWlsZGluZ3x8Jycsci5sb2NhdGlvbl9ncm91cHx8Jycsci5vcmRlcnx8Jycsci5wcm9qZWN0fHwnJyxyLmNvb3JkaW5hdG9yfHwnJyxyLmxvY2F0aW9ufHwnJyxyLnBhcnRfbm98fCcnLHIuc2VyaWFsfHwnJyxyLnNoaXBfZGF0ZXx8Jycsci5jb250YWluZXJzLHIudmFsdWUsci5hdmdfYWdlLHIubWF4X2FnZSxyLnN0YXR1c3x8Jycsci5mbGFnLGI/Yi5jaGFyZ2VfdHlwZTonJyxiP2IuY2hhcmdlX2Ftb3VudDonJyxiP2Iubm90ZTonJyxkLG5iPydZZXMnOicnXTsgfSk7CiAgZG93bmxvYWRDU1YoJ09mZnNpdGVfU3RvcmFnZS5jc3YnLCB0b0NTVihoZWFkZXJzLCBjc3ZSb3dzKSk7Cn0KCi8vIOKUgOKUgCBCaWxsaW5nIGRhdGEg4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSACmxldCBCSUxMSU5HID0ge307ICAvLyBvcmRlciAtPiB7Y2hhcmdlX2Ftb3VudCwgY2hhcmdlX3R5cGUsIG5vdGUsIHBjfQoKYXN5bmMgZnVuY3Rpb24gbG9hZEJpbGxpbmcoKSB7CiAgdHJ5IHsKICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKCcvYXBpL2JpbGxpbmcnKTsKICAgIEJJTExJTkcgPSBhd2FpdCByZXMuanNvbigpOyAgLy8ge29yZGVyOiB7Y2hhcmdlX2Ftb3VudCwgY2hhcmdlX3R5cGUsIG5vdGV9fQogICAgLy8gS1BJcyBhcmUgY29tcHV0ZWQgYnkgdXBkYXRlRGlzcG9zaXRpb25LUElzKCkgd2hpY2ggcnVucyBhZnRlciB0aGlzCiAgfSBjYXRjaChlKSB7IEJJTExJTkcgPSB7fTsgfQp9CgovLyDilIDilIAgRGlzcG9zaXRpb25zIChwZXJzaXN0ZWQgcGVyLW9yZGVyIHNlbGVjdGlvbnMpIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApsZXQgRElTUE9TSVRJT05TID0ge307Cgphc3luYyBmdW5jdGlvbiBsb2FkRGlzcG9zaXRpb25zKCkgewogIHRyeSB7CiAgICBjb25zdCByZXMgPSBhd2FpdCBmZXRjaCgnL2FwaS9kaXNwb3NpdGlvbnMnKTsKICAgIERJU1BPU0lUSU9OUyA9IGF3YWl0IHJlcy5qc29uKCk7CiAgfSBjYXRjaChlKSB7IERJU1BPU0lUSU9OUyA9IHt9OyB9Cn0KCmZ1bmN0aW9uIHNhdmVEaXNwb3NpdGlvbihzZWxlY3QpIHsKICBjb25zdCBvcmRlciA9IHNlbGVjdC5kYXRhc2V0Lm9yZGVyOwogIGNvbnN0IGRpc3AgID0gc2VsZWN0LnZhbHVlOwogIGlmICghb3JkZXIpIHJldHVybjsKICBESVNQT1NJVElPTlNbb3JkZXJdID0gZGlzcDsKICAvLyBTeW5jIHNhbWUgb3JkZXIncyBkcm9wZG93biBvbiBvdGhlciB0YWIgaWYgcHJlc2VudAogIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoYC5kaXNwLXNlbGVjdFtkYXRhLW9yZGVyPSIke29yZGVyfSJdYCkuZm9yRWFjaChzID0+IHsgcy52YWx1ZSA9IGRpc3A7IH0pOwogIC8vIFVwZGF0ZSBkaXNwb3NpdGlvbiBLUEkgY291bnRzCiAgdXBkYXRlRGlzcG9zaXRpb25LUElzKCk7CiAgLy8gUE9TVCB0byBzZXJ2ZXIKICBmZXRjaCgnL2FwaS9zZXRfZGlzcG9zaXRpb24nLCB7CiAgICBtZXRob2Q6ICdQT1NUJywKICAgIGhlYWRlcnM6IHsnQ29udGVudC1UeXBlJzonYXBwbGljYXRpb24vanNvbid9LAogICAgYm9keTogSlNPTi5zdHJpbmdpZnkoe29yZGVyLCBkaXNwb3NpdGlvbjogZGlzcH0pCiAgfSkuY2F0Y2goZSA9PiBjb25zb2xlLmVycm9yKCdEaXNwb3NpdGlvbiBzYXZlIGZhaWxlZCcsIGUpKTsKfQoKZnVuY3Rpb24gaXNTdG9yYWdlU3RhdHVzKHMpIHsgcmV0dXJuIHMgJiYgcy50b0xvd2VyQ2FzZSgpLnN0YXJ0c1dpdGgoJ3N0b3JhZ2UnKTsgfQoKZnVuY3Rpb24gdXBkYXRlRGlzcG9zaXRpb25LUElzKCkgewogIGNvbnN0IGl0ZW1zID0gUkFXPy5vZmZzaXRlPy5pdGVtcyB8fCBbXTsKICBsZXQgc29wT3JkZXJzPTAsc29wVmFsPTAsY2hhcmdlZE9yZGVycz0wLGNoYXJnZWRWYWw9MCx1bmJpbGxlZD0wLGJpbGxlZENvdW50PTAsYmlsbGVkVG90YWw9MDsKICBpdGVtcy5mb3JFYWNoKHIgPT4gewogICAgY29uc3QgZCAgICA9IERJU1BPU0lUSU9OU1tyLm9yZGVyXSB8fCAnJzsKICAgIGNvbnN0IGJpbGwgPSBCSUxMSU5HW3Iub3JkZXJdIHx8IG51bGw7CiAgICBpZiAoZCA9PT0gJ1NPUCBCdWlsZCBBaGVhZCcpIHsgc29wT3JkZXJzKys7ICAgICBzb3BWYWwgICAgICs9IHIudmFsdWV8fDA7IH0KICAgIGlmIChkID09PSAnU3RvcmFnZSBDaGFyZ2VkJykgeyBjaGFyZ2VkT3JkZXJzKys7IGNoYXJnZWRWYWwgKz0gci52YWx1ZXx8MDsgfQogICAgaWYgKGlzU3RvcmFnZVN0YXR1cyhyLnN0YXR1cykgJiYgIWJpbGwgJiYgZCAhPT0gJ1NPUCBCdWlsZCBBaGVhZCcpIHVuYmlsbGVkKys7CiAgICBpZiAoYmlsbCkgeyBiaWxsZWRDb3VudCsrOyBiaWxsZWRUb3RhbCArPSBiaWxsLmNoYXJnZV9hbW91bnR8fDA7IH0KICB9KTsKICBjb25zdCBzZXQgPSAoaWQsdmFsKSA9PiB7IGNvbnN0IGVsPWRvY3VtZW50LmdldEVsZW1lbnRCeUlkKGlkKTsgaWYoZWwpIGVsLnRleHRDb250ZW50PXZhbDsgfTsKICBzZXQoJ29mS3BpU09QJywgICAgICAgICBzb3BPcmRlcnMudG9Mb2NhbGVTdHJpbmcoKSk7CiAgc2V0KCdvZktwaVNPUFZhbCcsICAgICAgZm10JChzb3BWYWwpKTsKICBzZXQoJ29mS3BpQ2hhcmdlZCcsICAgICBjaGFyZ2VkT3JkZXJzLnRvTG9jYWxlU3RyaW5nKCkpOwogIHNldCgnb2ZLcGlDaGFyZ2VkVmFsJywgIGZtdCQoY2hhcmdlZFZhbCkpOwogIHNldCgnb2ZLcGlVbmJpbGxlZCcsICAgIHVuYmlsbGVkLnRvTG9jYWxlU3RyaW5nKCkpOwogIHNldCgnb2ZLcGlCaWxsZWRDb3VudCcsIGJpbGxlZENvdW50LnRvTG9jYWxlU3RyaW5nKCkpOwogIHNldCgnb2ZLcGlCaWxsZWRUb3RhbCcsIGZtdCQoYmlsbGVkVG90YWwpKTsKfQoKLy8g4pSA4pSAIE5ldyB0byBTdG9yYWdlIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApsZXQgTlNfREFUQSA9IFtdOwpsZXQgTlNfTE9BREVEID0gZmFsc2U7Cgphc3luYyBmdW5jdGlvbiBsb2FkTmV3U3RvcmFnZSgpIHsKICBpZiAoTlNfTE9BREVEKSByZXR1cm47CiAgTlNfTE9BREVEID0gdHJ1ZTsKICB0cnkgewogICAgY29uc3QgcmVzID0gYXdhaXQgZmV0Y2goJy9hcGkvbmV3X3N0b3JhZ2UnKTsKICAgIE5TX0RBVEEgPSBhd2FpdCByZXMuanNvbigpOwogICAgY29uc3QgbnNDb29yZHMgPSBbLi4ubmV3IFNldChOU19EQVRBLm1hcChyID0+IHIuY29vcmRpbmF0b3IpKV0uZmlsdGVyKEJvb2xlYW4pLnNvcnQoKTsKICAgIG1zUG9wdWxhdGUoJ21zLW5zY29vcmQnLCBuc0Nvb3JkcywgcmVuZGVyTmV3U3RvcmFnZSk7CiAgICByZW5kZXJOZXdTdG9yYWdlKCk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbmV3c3RvcmFnZUJhZGdlJykudGV4dENvbnRlbnQgPSBOU19EQVRBLmxlbmd0aDsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc0twaVRvdGFsJykudGV4dENvbnRlbnQgPSBOU19EQVRBLmxlbmd0aC50b0xvY2FsZVN0cmluZygpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ25zS3BpVmFsdWUnKS50ZXh0Q29udGVudCA9IGZtdCQoTlNfREFUQS5yZWR1Y2UoKGEscikgPT4gYSArIHIudmFsdWUsIDApKTsKICAgIGNvbnN0IHRvZGF5ID0gbmV3IERhdGUoKS50b0xvY2FsZURhdGVTdHJpbmcoJ2VuLVVTJywge3llYXI6J251bWVyaWMnLG1vbnRoOidsb25nJyxkYXk6J251bWVyaWMnfSk7CiAgICBjb25zdCBuZXdUb2RheSA9IE5TX0RBVEEuZmlsdGVyKHIgPT4gci5kZXRlY3RlZCA9PT0gdG9kYXkpLmxlbmd0aDsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc0twaU5ldycpLnRleHRDb250ZW50ID0gbmV3VG9kYXkudG9Mb2NhbGVTdHJpbmcoKTsKICB9IGNhdGNoKGUpIHsKICAgIGNvbnNvbGUuZXJyb3IoJ05ldyBzdG9yYWdlIGxvYWQgZmFpbGVkJywgZSk7CiAgfQp9CgpmdW5jdGlvbiByZXNldE5ld1N0b3JhZ2VGaWx0ZXJzKCkgewogIG1zUmVzZXQoJ21zLW5zY29vcmQnKTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbnNTZWFyY2gnKS52YWx1ZSA9ICcnOwogIHJlbmRlck5ld1N0b3JhZ2UoKTsKfQoKZnVuY3Rpb24gcmVuZGVyTmV3U3RvcmFnZSgpIHsKICBjb25zdCBjb29yZCA9IG1zR2V0U2VsZWN0ZWQoJ21zLW5zY29vcmQnKTsKICBjb25zdCBxICAgICA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc1NlYXJjaCcpLnZhbHVlLnRvTG93ZXJDYXNlKCk7CiAgbGV0IHJvd3MgPSBOU19EQVRBOwogIGlmIChjb29yZC5sZW5ndGgpIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IGNvb3JkLmluY2x1ZGVzKHIuY29vcmRpbmF0b3IpKTsKICBpZiAocSkgICAgICAgICAgICByb3dzID0gcm93cy5maWx0ZXIociA9PgogICAgci5vcmRlci50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8IHIucHJvamVjdC50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHEpIHx8CiAgICByLmNvb3JkaW5hdG9yLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc01ldGEnKS50ZXh0Q29udGVudCA9IHJvd3MubGVuZ3RoICsgJyBvcmRlcnMnOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCduc0JvZHknKS5pbm5lckhUTUwgPSByb3dzLm1hcChyID0+IHsKICAgIGNvbnN0IHB1c2hlZCA9IHIuc2hpcF9kYXRlX3B1c2hlZDsKICAgIGNvbnN0IHJvd0JnICA9IHB1c2hlZCA/ICdiYWNrZ3JvdW5kOiNmZmYwZjA7Ym9yZGVyLWxlZnQ6M3B4IHNvbGlkICNkYzM1NDU7JyA6ICdiYWNrZ3JvdW5kOiNmZmY4Zjg7JzsKICAgIGNvbnN0IGRpc3BWYWwgPSBESVNQT1NJVElPTlNbci5vcmRlcl0gfHwgJyc7CiAgICByZXR1cm4gYDx0ciBzdHlsZT0iJHtyb3dCZ30iPgogICAgICA8dGQ+JHtyLmRldGVjdGVkfTwvdGQ+CiAgICAgIDx0ZD48c3Ryb25nPiR7ci5vcmRlcn08L3N0cm9uZz48L3RkPgogICAgICA8dGQgY2xhc3M9IndyYXAiPiR7ci5wcm9qZWN0fTwvdGQ+CiAgICAgIDx0ZD4ke3IuY29vcmRpbmF0b3J9PC90ZD4KICAgICAgPHRkIHN0eWxlPSJjb2xvcjojODg4Ij4ke3IucHJldl9zdGF0dXN9PC90ZD4KICAgICAgPHRkIHN0eWxlPSJjb2xvcjojOEIwMDAwO2ZvbnQtd2VpZ2h0OjYwMCI+JHtyLm5ld19zdGF0dXN9PC90ZD4KICAgICAgPHRkPiR7cHVzaGVkID8gJzxzcGFuIHN0eWxlPSJjb2xvcjojZGMzNTQ1O2ZvbnQtd2VpZ2h0OjYwMCIgdGl0bGU9IlNoaXAgZGF0ZSBwdXNoZWQgb3V0IGZyb20gJyArIChyLnByZXZfc2hpcF9kYXRlfHwnPycpICsgJyI+4pqgICcgKyByLnNoaXBfZGF0ZSArICc8L3NwYW4+JyA6IHIuc2hpcF9kYXRlfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4oci5jb250YWluZXJzKX08L3RkPgogICAgICA8dGQgY2xhc3M9Im51bSI+JHtmbXQkKHIudmFsdWUpfTwvdGQ+CiAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdEYoci5hdmdfYWdlKX08L3RkPgogICAgICA8dGQ+PHNlbGVjdCBjbGFzcz0iZGlzcC1zZWxlY3QiIGRhdGEtb3JkZXI9IiR7ci5vcmRlcn0iIG9uY2hhbmdlPSJzYXZlRGlzcG9zaXRpb24odGhpcykiIHN0eWxlPSJmb250LXNpemU6MTFweDtwYWRkaW5nOjJweCA0cHg7Ym9yZGVyOjFweCBzb2xpZCAjY2NjO2JvcmRlci1yYWRpdXM6NHB4O3dpZHRoOjE0MHB4OyI+CiAgICAgICAgPG9wdGlvbiB2YWx1ZT0iIj7igJQgU2VsZWN0IOKAlDwvb3B0aW9uPgogICAgICAgIDxvcHRpb24gdmFsdWU9IlNPUCBCdWlsZCBBaGVhZCIgJHtkaXNwVmFsPT09J1NPUCBCdWlsZCBBaGVhZCc/J3NlbGVjdGVkJzonJ30+U09QIEJ1aWxkIEFoZWFkPC9vcHRpb24+CiAgICAgICAgPG9wdGlvbiB2YWx1ZT0iU3RvcmFnZSBDaGFyZ2VkIiAke2Rpc3BWYWw9PT0nU3RvcmFnZSBDaGFyZ2VkJz8nc2VsZWN0ZWQnOicnfT5TdG9yYWdlIENoYXJnZWQ8L29wdGlvbj4KICAgICAgPC9zZWxlY3Q+PC90ZD4KICAgIDwvdHI+YDsKICB9KS5qb2luKCcnKSB8fCAnPHRyPjx0ZCBjb2xzcGFuPSIxMSIgc3R5bGU9InRleHQtYWxpZ246Y2VudGVyO3BhZGRpbmc6MzJweDtjb2xvcjp2YXIoLS1tdXRlZCkiPk5vIG9yZGVycyBkZXRlY3RlZCB5ZXQuIFVwbG9hZCBmaWxlcyBvbiBjb25zZWN1dGl2ZSBkYXlzIHRvIHN0YXJ0IHRyYWNraW5nIGNoYW5nZXMuPC90ZD48L3RyPic7CiAgaWYgKCFyb3dzLmxlbmd0aCkgeyBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbnNGb290JykuaW5uZXJIVE1MID0gJyc7IHJldHVybjsgfQogIGNvbnN0IHRvdCA9IHJvd3MucmVkdWNlKChhLHIpID0+ICh7Y29udGFpbmVyczogYS5jb250YWluZXJzK3IuY29udGFpbmVycywgdmFsdWU6IGEudmFsdWUrci52YWx1ZX0pLCB7Y29udGFpbmVyczowLCB2YWx1ZTowfSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ25zRm9vdCcpLmlubmVySFRNTCA9IGA8dHI+PHRkIGNvbHNwYW49IjgiPlRPVEFMIOKAlCAke3Jvd3MubGVuZ3RofSBvcmRlcnM8L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdE4odG90LmNvbnRhaW5lcnMpfTwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10JCh0b3QudmFsdWUpfTwvdGQ+PHRkPjwvdGQ+PC90cj5gOwp9CgpmdW5jdGlvbiBleHBvcnROZXdTdG9yYWdlQ1NWKCkgewogIGNvbnN0IGNvb3JkID0gbXNHZXRTZWxlY3RlZCgnbXMtbnNjb29yZCcpOwogIGNvbnN0IHEgICAgID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ25zU2VhcmNoJykudmFsdWUudG9Mb3dlckNhc2UoKTsKICBsZXQgcm93cyA9IE5TX0RBVEE7CiAgaWYgKGNvb3JkLmxlbmd0aCkgcm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gY29vcmQuaW5jbHVkZXMoci5jb29yZGluYXRvcikpOwogIGlmIChxKSAgICAgICAgICAgIHJvd3MgPSByb3dzLmZpbHRlcihyID0+IHIub3JkZXIudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSB8fCByLnByb2plY3QudG9Mb3dlckNhc2UoKS5pbmNsdWRlcyhxKSk7CiAgY29uc3QgaGVhZGVycyA9IFsnRGV0ZWN0ZWQgRGF0ZScsJ09yZGVyICMnLCdQcm9qZWN0JywnQ29vcmRpbmF0b3InLCdQcmV2aW91cyBTdGF0dXMnLCdOZXcgU3RhdHVzJywnU2hpcCBEYXRlJywnU2hpcCBEYXRlIFB1c2hlZD8nLCdDb250YWluZXJzJywnVmFsdWUnLCdBdmcgQWdlJywnRGlzcG9zaXRpb24nXTsKICBjb25zdCBjc3ZSb3dzID0gcm93cy5tYXAociA9PiBbci5kZXRlY3RlZCwgci5vcmRlciwgci5wcm9qZWN0LCByLmNvb3JkaW5hdG9yLCByLnByZXZfc3RhdHVzLCByLm5ld19zdGF0dXMsIHIuc2hpcF9kYXRlLCByLnNoaXBfZGF0ZV9wdXNoZWQgPyAnWWVzICh3YXMgJyArIChyLnByZXZfc2hpcF9kYXRlfHwnPycpICsgJyknIDogJ05vJywgci5jb250YWluZXJzLCByLnZhbHVlLCByLmF2Z19hZ2UsIERJU1BPU0lUSU9OU1tyLm9yZGVyXXx8JyddKTsKICBkb3dubG9hZENTVignTmV3X3RvX1N0b3JhZ2UuY3N2JywgdG9DU1YoaGVhZGVycywgY3N2Um93cykpOwp9CgovLyDilIDilIAgSW5pdCDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKbG9hZERhdGEoKTsKPC9zY3JpcHQ+CjwvYm9keT4KPC9odG1sPg==").decode('utf-8')

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

def _run_processing(wh_path, notes_path, warehouse_filename, notes_filename, billing_path=None, billing_filename=None):
    global _processing_state
    _processing_state = {'status': 'processing', 'error': None}
    try:
        data = process_files(wh_path, notes_path, billing_path)
        PROCESSED_PATH.write_text(json.dumps(data))
        META_PATH.write_text(json.dumps({
            'warehouse_filename': warehouse_filename,
            'notes_filename':     notes_filename,
            'billing_filename':   billing_filename or '',
            'rows':               data['kpis']['total_containers'],
            'sdrop_items':        data['sdrop']['kpis']['total_items'],
            'uploaded':           data['uploaded'],
        }))
        _processing_state = {'status': 'done', 'error': None}
    except Exception as e:
        import traceback
        _processing_state = {'status': 'error', 'error': str(e), 'detail': traceback.format_exc()}

@app.route('/upload', methods=['POST'])
def upload():
    import threading
    warehouse = request.files.get('warehouse')
    notes = request.files.get('notes')
    if not warehouse or not notes:
        return jsonify({'error': 'Both files are required'}), 400
    billing    = request.files.get('billing')
    wh_path    = DATA_DIR / 'warehouse_temp.xlsx'
    notes_path = DATA_DIR / 'notes_temp.csv'
    warehouse.save(wh_path)
    notes.save(notes_path)
    if billing:
        billing.save(BILLING_PATH)
    elif not BILLING_PATH.exists():
        # Create empty billing file placeholder
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Order Number','Charge Amount','Charge Type','PC','Note'])
        wb.save(BILLING_PATH)
    wh_name = warehouse.filename
    nt_name = notes.filename
    # Start processing in background thread — returns immediately to avoid timeout
    t = threading.Thread(target=_run_processing, args=(wh_path, notes_path, wh_name, nt_name), daemon=True)
    t.start()
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

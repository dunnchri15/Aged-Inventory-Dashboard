import os, json
from pathlib import Path
from flask import Flask, render_template, request, jsonify, redirect, url_for
import pandas as pd
import numpy as np

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

DATA_DIR = Path('data')
DATA_DIR.mkdir(exist_ok=True)
PROCESSED_PATH = DATA_DIR / 'processed.json'
META_PATH      = DATA_DIR / 'meta.json'

# Use calamine (fast Rust reader) with openpyxl as fallback
def read_excel_fast(path):
    try:
        return pd.read_excel(path, engine='calamine')
    except Exception:
        return pd.read_excel(path, engine='openpyxl')

# ── Data processing ───────────────────────────────────────────────────────────

AGE_BUCKET_ORDER = ['0-7d','8-14d','15-30d','31-60d','61-90d','91-180d','181-365d','365d+']
AGE_BUCKET_LABELS = {
    '0-7d': '0 – 7 Days', '8-14d': '8 – 14 Days', '15-30d': '15 – 30 Days',
    '31-60d': '31 – 60 Days', '61-90d': '61 – 90 Days', '91-180d': '91 – 180 Days',
    '181-365d': '181 – 365 Days', '365d+': 'Over 365 Days'
}

def assign_age_buckets(series):
    """Vectorized age bucket assignment — ~1000x faster than .apply()"""
    age = series
    result = pd.Series('365d+', index=series.index)
    result = result.where(age.isna(), result)
    result[age <= 365] = '181-365d'
    result[age <= 180] = '91-180d'
    result[age <= 90]  = '61-90d'
    result[age <= 60]  = '31-60d'
    result[age <= 30]  = '15-30d'
    result[age <= 14]  = '8-14d'
    result[age <= 7]   = '0-7d'
    result[age.isna()] = None
    return result

def process_files(warehouse_path, notes_path):
    df    = read_excel_fast(warehouse_path)
    notes = pd.read_csv(notes_path)

    df['PROJECT_NAME'] = df['SHIP_TO'].str.replace(r'^S-', '', regex=True).str.strip()
    df['AGE_BUCKET']   = assign_age_buckets(df['INV_AGE'])

    coord_map = (notes[['Order No', 'Project Coordinator']]
                 .drop_duplicates('Order No')
                 .dropna(subset=['Project Coordinator']))
    df = df.merge(coord_map, left_on='ORDERS', right_on='Order No', how='left')
    df.rename(columns={'Project Coordinator': 'COORDINATOR'}, inplace=True)

    df['SHIP_DATE'] = pd.to_datetime(df['SHIP_DATE'], errors='coerce')

    # ── KPIs ──────────────────────────────────────────────────────────────────
    total_value  = float(df['EXTENDED_COST'].sum())
    over90_mask  = df['INV_AGE'] > 90
    over90_value = float(df.loc[over90_mask, 'EXTENDED_COST'].sum())
    avg_age      = float(df['INV_AGE'].mean()) if len(df) else 0
    in_storage   = df[df['LOCATION_GROUP'].str.contains('Storage',  case=False, na=False)]
    finance_hold = df[df['ORDER_STATUS'].str.contains('Finance Hold', case=False, na=False)]

    kpis = {
        'total_value':        total_value,
        'over90_value':       over90_value,
        'over90_pct':         over90_value / total_value if total_value else 0,
        'total_containers':   len(df),
        'avg_age':            round(avg_age, 1),
        'in_storage_value':   float(in_storage['EXTENDED_COST'].sum()),
        'finance_hold_value': float(finance_hold['EXTENDED_COST'].sum()),
    }

    # ── Coordinator table — vectorized groupby ────────────────────────────────
    df_coord = df[df['COORDINATOR'].notna()].copy()
    coord_agg = df_coord.groupby('COORDINATOR').agg(
        containers=('ORDERS', 'count'),
        value=('EXTENDED_COST', 'sum'),
        over90_value=('EXTENDED_COST', lambda x: x[df_coord.loc[x.index, 'INV_AGE'] > 90].sum()),
        avg_age=('INV_AGE', 'mean'),
    ).reset_index()
    coord_agg['over90_pct'] = coord_agg['over90_value'] / coord_agg['value'].replace(0, np.nan)
    coord_agg['over90_pct'] = coord_agg['over90_pct'].fillna(0)
    coord_tbl = coord_agg.sort_values('COORDINATOR').to_dict('records')
    coord_tbl = [{
        'coordinator': r['COORDINATOR'],
        'containers':  int(r['containers']),
        'value':       round(float(r['value']), 2),
        'over90_value': round(float(r['over90_value']), 2),
        'over90_pct':  float(r['over90_pct']),
        'avg_age':     round(float(r['avg_age']), 1),
    } for r in coord_tbl]

    # ── Age bucket table — vectorized ─────────────────────────────────────────
    age_tbl = []
    for code in AGE_BUCKET_ORDER:
        mask = df['AGE_BUCKET'] == code
        val  = float(df.loc[mask, 'EXTENDED_COST'].sum())
        age_tbl.append({
            'bucket':     code,
            'label':      AGE_BUCKET_LABELS[code],
            'containers': int(mask.sum()),
            'value':      round(val, 2),
            'pct':        val / total_value if total_value else 0,
        })

    # ── Orders table — vectorized groupby ─────────────────────────────────────
    valid = df[df['ORDERS'].notna() & (df['ORDERS'].astype(str) != '.')].copy()
    valid['SHIP_TS'] = pd.to_datetime(valid['SHIP_DATE'], errors='coerce')

    # Groupby aggregation
    ord_agg = valid.groupby('ORDERS').agg(
        project=('PROJECT_NAME', 'first'),
        containers=('ORDERS', 'count'),
        value=('EXTENDED_COST', 'sum'),
        avg_age=('INV_AGE', 'mean'),
        max_age=('INV_AGE', 'max'),
        age_bucket=('AGE_BUCKET', lambda x: x.mode().iloc[0] if len(x) else ''),
        status=('ORDER_STATUS', lambda x: x.mode().iloc[0] if len(x) else ''),
        ship_min=('SHIP_TS', 'min'),
        ship_max=('SHIP_TS', 'max'),
    ).reset_index()

    # Coordinator per order (first non-null)
    coord_per_order = (valid[valid['COORDINATOR'].notna()]
                       .groupby('ORDERS')['COORDINATOR']
                       .first()
                       .reset_index()
                       .rename(columns={'COORDINATOR': 'coordinator'}))
    ord_agg = ord_agg.merge(coord_per_order, on='ORDERS', how='left')
    ord_agg['coordinator'] = ord_agg['coordinator'].fillna('Unassigned')

    # Ship date range string
    def ship_range(row):
        if pd.isna(row['ship_min']): return ''
        mn = row['ship_min'].strftime('%m/%d/%y')
        mx = row['ship_max'].strftime('%m/%d/%y')
        return mn if mn == mx else f'{mn} – {mx}'
    ord_agg['ship_range'] = ord_agg.apply(ship_range, axis=1)

    orders_tbl = [{
        'order':       str(r['ORDERS']),
        'project':     str(r['project'] or ''),
        'coordinator': r['coordinator'],
        'containers':  int(r['containers']),
        'value':       round(float(r['value']), 2),
        'avg_age':     round(float(r['avg_age']), 1),
        'max_age':     int(r['max_age']),
        'age_bucket':  str(r['age_bucket']),
        'status':      str(r['status']),
        'ship_range':  r['ship_range'],
    } for _, r in ord_agg.sort_values('project').iterrows()]

    # ── S-Drop table ──────────────────────────────────────────────────────────
    sdrop_raw = df[
        df['LOCATION'].str.contains('drop', case=False, na=False) &
        (df['INV_AGE'] > 2)
    ].copy().sort_values('INV_AGE', ascending=False)

    loc_sum = (sdrop_raw.groupby('LOCATION')
               .agg(items=('LOCATION','count'), value=('EXTENDED_COST','sum'),
                    avg_age=('INV_AGE','mean'), max_age=('INV_AGE','max'))
               .reset_index().sort_values('items', ascending=False))

    sdrop_by_location = [{
        'location': r['LOCATION'], 'items': int(r['items']),
        'value': round(float(r['value']), 2),
        'avg_age': round(float(r['avg_age']), 1), 'max_age': int(r['max_age']),
    } for _, r in loc_sum.iterrows()]

    sdrop_items = []
    for _, row in sdrop_raw.iterrows():
        age = int(row['INV_AGE'])
        flag = '>90 Days' if age > 90 else ('>30 Days' if age > 30 else '≤30 Days')
        sdrop_items.append({
            'location':    str(row['LOCATION']),
            'order':       str(row['ORDERS']) if pd.notna(row['ORDERS']) and str(row['ORDERS']) != '.' else '—',
            'project':     str(row['PROJECT_NAME']) if pd.notna(row['PROJECT_NAME']) and str(row['PROJECT_NAME']) != '.' else '—',
            'coordinator': str(row['COORDINATOR']) if pd.notna(row['COORDINATOR']) else 'Unassigned',
            'part_no':     str(row['PART_NO']) if pd.notna(row['PART_NO']) else '—',
            'part_group':  str(row['PART_GROUP']) if pd.notna(row['PART_GROUP']) and str(row['PART_GROUP']) != '.' else '—',
            'qty':         float(row['QUANTITY']) if pd.notna(row['QUANTITY']) else 0,
            'age':         age,
            'value':       round(float(row['EXTENDED_COST']), 2) if pd.notna(row['EXTENDED_COST']) else 0,
            'status':      str(row['ORDER_STATUS']) if pd.notna(row['ORDER_STATUS']) and str(row['ORDER_STATUS']) != '.' else '—',
            'flag':        flag,
        })

    uniq_orders = sdrop_raw[sdrop_raw['ORDERS'].notna() & (sdrop_raw['ORDERS'].astype(str) != '.')]['ORDERS'].nunique()
    sdrop = {
        'kpis': {
            'total_items': len(sdrop_raw), 'total_value': round(float(sdrop_raw['EXTENDED_COST'].sum()), 2),
            'unique_orders': int(uniq_orders),
            'avg_age': round(float(sdrop_raw['INV_AGE'].mean()), 1) if len(sdrop_raw) else 0,
            'max_age': int(sdrop_raw['INV_AGE'].max()) if len(sdrop_raw) else 0,
        },
        'by_location': sdrop_by_location,
        'items':       sdrop_items,
    }

    projects     = sorted(df['PROJECT_NAME'].dropna().unique().tolist())
    coordinators = sorted(df['COORDINATOR'].dropna().unique().tolist())
    drop_locs    = sorted(sdrop_raw['LOCATION'].unique().tolist())

    return {
        'kpis': kpis, 'coord_table': coord_tbl, 'age_table': age_tbl,
        'orders_table': orders_tbl, 'sdrop': sdrop,
        'projects': projects, 'coordinators': coordinators, 'drop_locations': drop_locs,
    }


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    has_data = PROCESSED_PATH.exists()
    meta = json.loads(META_PATH.read_text()) if META_PATH.exists() else {}
    return render_template('index.html', has_data=has_data, meta=meta)

@app.route('/upload', methods=['POST'])
def upload():
    warehouse = request.files.get('warehouse')
    notes     = request.files.get('notes')
    if not warehouse or not notes:
        return jsonify({'error': 'Both files are required'}), 400

    wh_path    = DATA_DIR / 'warehouse_temp.xlsx'
    notes_path = DATA_DIR / 'notes_temp.csv'
    warehouse.save(wh_path)
    notes.save(notes_path)

    try:
        data = process_files(wh_path, notes_path)
        PROCESSED_PATH.write_text(json.dumps(data))
        META_PATH.write_text(json.dumps({
            'warehouse_filename': warehouse.filename,
            'notes_filename':     notes.filename,
            'rows':               data['kpis']['total_containers'],
            'sdrop_items':        data['sdrop']['kpis']['total_items'],
            'uploaded':           pd.Timestamp.now().strftime('%B %d, %Y'),
        }))
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/dashboard')
def dashboard():
    if not PROCESSED_PATH.exists():
        return redirect(url_for('index'))
    meta = json.loads(META_PATH.read_text()) if META_PATH.exists() else {}
    return render_template('dashboard.html', meta=meta)

@app.route('/api/data')
def api_data():
    if not PROCESSED_PATH.exists():
        return jsonify({'error': 'No data loaded'}), 404
    return PROCESSED_PATH.read_text(), 200, {'Content-Type': 'application/json'}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

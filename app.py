import os, json, re
from pathlib import Path
from flask import Flask, render_template, request, jsonify, redirect, url_for
import pandas as pd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

DATA_DIR = Path('data')
DATA_DIR.mkdir(exist_ok=True)
PROCESSED_PATH = DATA_DIR / 'processed.json'
META_PATH = DATA_DIR / 'meta.json'


# ── Data processing ───────────────────────────────────────────────────────────

def age_bucket(age):
    if pd.isna(age): return None
    a = int(age)
    if a <= 7:    return '0-7d'
    elif a <= 14:  return '8-14d'
    elif a <= 30:  return '15-30d'
    elif a <= 60:  return '31-60d'
    elif a <= 90:  return '61-90d'
    elif a <= 180: return '91-180d'
    elif a <= 365: return '181-365d'
    else:          return '365d+'

AGE_BUCKET_ORDER = ['0-7d','8-14d','15-30d','31-60d','61-90d','91-180d','181-365d','365d+']
AGE_BUCKET_LABELS = {
    '0-7d': '0 – 7 Days', '8-14d': '8 – 14 Days', '15-30d': '15 – 30 Days',
    '31-60d': '31 – 60 Days', '61-90d': '61 – 90 Days', '91-180d': '91 – 180 Days',
    '181-365d': '181 – 365 Days', '365d+': 'Over 365 Days'
}

def process_files(warehouse_path, notes_path):
    df = pd.read_excel(warehouse_path)
    notes = pd.read_csv(notes_path)

    df['PROJECT_NAME'] = df['SHIP_TO'].str.replace(r'^S-', '', regex=True).str.strip()
    df['AGE_BUCKET'] = df['INV_AGE'].apply(age_bucket)

    coord_map = (notes[['Order No', 'Project Coordinator']]
                 .drop_duplicates('Order No')
                 .dropna(subset=['Project Coordinator']))
    df = df.merge(coord_map, left_on='ORDERS', right_on='Order No', how='left')
    df.rename(columns={'Project Coordinator': 'COORDINATOR'}, inplace=True)

    df['SHIP_DATE'] = pd.to_datetime(df['SHIP_DATE'], errors='coerce')
    df['SHIP_DATE_STR'] = df['SHIP_DATE'].dt.strftime('%m/%d/%Y').where(df['SHIP_DATE'].notna(), None)

    # ── KPIs ──────────────────────────────────────────────────────────────────
    total_value = float(df['EXTENDED_COST'].sum())
    over90 = df[df['INV_AGE'] > 90]
    over90_value = float(over90['EXTENDED_COST'].sum())
    avg_age = float(df['INV_AGE'].mean()) if len(df) else 0
    in_storage = df[df['LOCATION_GROUP'].str.contains('Storage', case=False, na=False)]
    finance_hold = df[df['ORDER_STATUS'].str.contains('Finance Hold', case=False, na=False)]

    kpis = {
        'total_value': total_value,
        'over90_value': over90_value,
        'over90_pct': over90_value / total_value if total_value else 0,
        'total_containers': len(df),
        'avg_age': round(avg_age, 1),
        'in_storage_value': float(in_storage['EXTENDED_COST'].sum()),
        'finance_hold_value': float(finance_hold['EXTENDED_COST'].sum()),
    }

    # ── Coordinator table ─────────────────────────────────────────────────────
    coord_tbl = []
    for coord, grp in df[df['COORDINATOR'].notna()].groupby('COORDINATOR'):
        o90 = grp[grp['INV_AGE'] > 90]
        coord_tbl.append({
            'coordinator': coord,
            'containers': len(grp),
            'value': round(float(grp['EXTENDED_COST'].sum()), 2),
            'over90_value': round(float(o90['EXTENDED_COST'].sum()), 2),
            'over90_pct': float(o90['EXTENDED_COST'].sum() / grp['EXTENDED_COST'].sum()) if grp['EXTENDED_COST'].sum() else 0,
            'avg_age': round(float(grp['INV_AGE'].mean()), 1),
        })
    coord_tbl.sort(key=lambda x: x['coordinator'])

    # ── Age bucket table ──────────────────────────────────────────────────────
    age_tbl = []
    for code in AGE_BUCKET_ORDER:
        grp = df[df['AGE_BUCKET'] == code]
        age_tbl.append({
            'bucket': code,
            'label': AGE_BUCKET_LABELS[code],
            'containers': len(grp),
            'value': round(float(grp['EXTENDED_COST'].sum()), 2),
            'pct': float(grp['EXTENDED_COST'].sum() / total_value) if total_value else 0,
        })

    # ── Orders table (per-order summary) ─────────────────────────────────────
    orders_tbl = []
    valid = df[df['ORDERS'].notna() & (df['ORDERS'] != '.')]
    for order, grp in valid.groupby('ORDERS'):
        o90 = grp[grp['INV_AGE'] > 90]
        ship_dates = grp['SHIP_DATE'].dropna()
        ship_range = ''
        if len(ship_dates):
            mn = ship_dates.min().strftime('%m/%d/%y')
            mx = ship_dates.max().strftime('%m/%d/%y')
            ship_range = mn if mn == mx else f'{mn} – {mx}'
        orders_tbl.append({
            'order': str(order),
            'project': str(grp['PROJECT_NAME'].iloc[0] or ''),
            'coordinator': str(grp['COORDINATOR'].dropna().iloc[0]) if grp['COORDINATOR'].notna().any() else 'Unassigned',
            'containers': len(grp),
            'value': round(float(grp['EXTENDED_COST'].sum()), 2),
            'avg_age': round(float(grp['INV_AGE'].mean()), 1),
            'max_age': int(grp['INV_AGE'].max()),
            'ship_range': ship_range,
            'age_bucket': str(grp['AGE_BUCKET'].mode().iloc[0]) if len(grp) else '',
            'status': str(grp['ORDER_STATUS'].mode().iloc[0]) if len(grp) else '',
        })
    orders_tbl.sort(key=lambda x: x['project'])

    # ── Filter lists ──────────────────────────────────────────────────────────
    projects = sorted(df['PROJECT_NAME'].dropna().unique().tolist())
    coordinators = sorted(df['COORDINATOR'].dropna().unique().tolist())

    return {
        'kpis': kpis,
        'coord_table': coord_tbl,
        'age_table': age_tbl,
        'orders_table': orders_tbl,
        'projects': projects,
        'coordinators': coordinators,
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
    notes = request.files.get('notes')

    if not warehouse or not notes:
        return jsonify({'error': 'Both files are required'}), 400

    wh_path = DATA_DIR / 'warehouse_temp.xlsx'
    notes_path = DATA_DIR / 'notes_temp.csv'
    warehouse.save(wh_path)
    notes.save(notes_path)

    try:
        data = process_files(wh_path, notes_path)
        PROCESSED_PATH.write_text(json.dumps(data))
        META_PATH.write_text(json.dumps({
            'warehouse_filename': warehouse.filename,
            'notes_filename': notes.filename,
            'rows': data['kpis']['total_containers'],
            'uploaded': pd.Timestamp.now().strftime('%B %d, %Y'),
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

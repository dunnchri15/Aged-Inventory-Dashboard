import os, json
from pathlib import Path
from flask import Flask, request, jsonify, redirect, Response
import pandas as pd
import numpy as np

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

DATA_DIR = Path('data')
DATA_DIR.mkdir(exist_ok=True)
PROCESSED_PATH = DATA_DIR / 'processed.json'
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
def process_files(warehouse_path, notes_path):
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

    loc_sum = (sdrop_df.groupby('LOCATION')
               .agg(items=('LOCATION','count'), value=('EXTENDED_COST','sum'),
                    avg_age=('INV_AGE','mean'), max_age=('INV_AGE','max'))
               .reset_index().sort_values('items', ascending=False))

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

    return {
        'kpis':         kpis,
        'coord_table':  coord_tbl,
        'age_table':    age_tbl,
        'orders_table': orders_tbl,
        'sdrop':        sdrop,
        'projects':     sorted(df['PROJECT_NAME'].dropna().unique().tolist()),
        'coordinators': sorted(df['COORDINATOR'].dropna().unique().tolist()),
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
['wz','nz'].forEach(function(id){
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
  [1000,4000,8000,11000].forEach(function(d,i){setTimeout(function(){setStep(i+2);},d);});
  try{
    var res=await fetch('/upload',{method:'POST',body:fd});
    var data=await res.json();
    if(data.success){setStep(5);document.getElementById('pf').style.width='100%';setTimeout(function(){window.location.href='/dashboard';},800);}
    else throw new Error(data.error||'Upload failed');
  }catch(ex){
    document.getElementById('prog').style.display='none';
    err.textContent='&#9888; '+ex.message;err.style.display='block';btn.disabled=false;
  }
});
</script></body></html>"""


DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Aged Inventory Dashboard</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    :root {
      --navy: #1F3864; --blue: #2E75B6; --light-blue: #D6E4F0; --section-bg: #EBF3FB;
      --green: #375623; --green-light: #E2EFDA; --red: #C00000; --red-light: #FFCCCC;
      --amber: #C55A11; --amber-light: #FFF2CC;
      --text: #1e293b; --muted: #64748b; --border: #e2e8f0; --white: #ffffff;
      --surface: #f8fafc;
    }

    body { font-family: 'Segoe UI', Arial, sans-serif; background: #f0f4f8;
           color: var(--text); font-size: 13px; }

    /* ── Header ── */
    .header { background: var(--navy); color: #fff; padding: 0 32px;
              display: flex; align-items: center; justify-content: space-between;
              height: 56px; position: sticky; top: 0; z-index: 100;
              box-shadow: 0 2px 8px rgba(0,0,0,.2); }
    .header-left { display: flex; align-items: center; gap: 12px; }
    .header h1 { font-size: 17px; font-weight: 700; }
    .header-sub { font-size: 11px; opacity: .7; margin-top: 1px; }
    .upload-link { background: rgba(255,255,255,.15); color: #fff; text-decoration: none;
                   padding: 6px 14px; border-radius: 6px; font-size: 12px; font-weight: 600;
                   transition: background .2s; border: 1px solid rgba(255,255,255,.25); }
    .upload-link:hover { background: rgba(255,255,255,.25); }

    /* ── Tab nav ── */
    .tab-nav { background: var(--navy); padding: 0 32px;
               display: flex; gap: 4px; border-top: 1px solid rgba(255,255,255,.1); }
    .tab-btn { padding: 10px 20px; font-size: 12px; font-weight: 600; color: rgba(255,255,255,.6);
               background: transparent; border: none; cursor: pointer; border-bottom: 3px solid transparent;
               transition: all .2s; white-space: nowrap; }
    .tab-btn:hover { color: rgba(255,255,255,.9); }
    .tab-btn.active { color: #fff; border-bottom-color: var(--amber-light); }
    .tab-btn .tab-badge { display: inline-block; background: var(--red); color: #fff;
                          border-radius: 10px; padding: 1px 7px; font-size: 10px;
                          margin-left: 6px; font-weight: 700; }

    /* ── Filters bar ── */
    .filters-bar { background: #162d54; padding: 10px 32px;
                   display: flex; gap: 12px; align-items: flex-end; flex-wrap: wrap;
                   border-bottom: 1px solid rgba(255,255,255,.08); }
    .filters-bar.sdrop-filters { display: none; }
    .filter-group { display: flex; flex-direction: column; gap: 4px; }
    .filter-group label { font-size: 10px; color: rgba(255,255,255,.65); font-weight: 600;
                          text-transform: uppercase; letter-spacing: .05em; }
    .filter-group select { background: rgba(255,255,255,.12); color: #fff; border: 1px solid rgba(255,255,255,.25);
                           border-radius: 6px; padding: 6px 28px 6px 10px; font-size: 12px; font-weight: 500;
                           cursor: pointer; appearance: none; min-width: 180px; max-width: 260px;
                           background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");
                           background-repeat: no-repeat; background-position: right 8px center; }
    .filter-group select option { background: var(--navy); }
    .filter-group select:focus { outline: none; border-color: var(--light-blue); }
    .reset-btn { background: rgba(255,255,255,.1); color: rgba(255,255,255,.8); border: 1px solid rgba(255,255,255,.2);
                 border-radius: 6px; padding: 6px 14px; font-size: 12px; cursor: pointer;
                 transition: all .2s; align-self: flex-end; }
    .reset-btn:hover { background: rgba(255,255,255,.2); color: #fff; }

    /* ── Main layout ── */
    .main { padding: 24px 32px; max-width: 1600px; margin: 0 auto; }
    .tab-panel { display: none; }
    .tab-panel.active { display: block; }

    /* ── KPI cards ── */
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
                gap: 16px; margin-bottom: 24px; }
    .kpi-card { background: #fff; border-radius: 10px; padding: 20px 22px;
                box-shadow: 0 1px 4px rgba(0,0,0,.08); border-top: 3px solid var(--blue); }
    .kpi-card.accent-red   { border-top-color: var(--red); }
    .kpi-card.accent-amber { border-top-color: var(--amber); }
    .kpi-card.accent-green { border-top-color: var(--green); }
    .kpi-label { font-size: 11px; color: var(--muted); font-weight: 600; text-transform: uppercase;
                 letter-spacing: .04em; margin-bottom: 8px; }
    .kpi-value { font-size: 26px; font-weight: 700; color: var(--navy); line-height: 1; }
    .kpi-sub   { font-size: 11px; color: var(--muted); margin-top: 5px; }

    /* ── Section cards ── */
    .section-card { background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.08);
                    margin-bottom: 24px; overflow: hidden; }
    .section-header { padding: 16px 20px; border-bottom: 1px solid var(--border);
                      display: flex; align-items: center; justify-content: space-between; }
    .section-title { font-size: 14px; font-weight: 700; color: var(--navy); }
    .section-meta  { font-size: 11px; color: var(--muted); }

    /* ── Search / filter input ── */
    .search-box { padding: 12px 20px; border-bottom: 1px solid var(--border); display: flex; gap: 10px; }
    .search-box input, .search-box select {
      padding: 7px 12px; border: 1px solid var(--border); border-radius: 6px;
      font-size: 12px; color: var(--text); background: var(--surface); }
    .search-box input { flex: 1; max-width: 320px; }
    .search-box input:focus, .search-box select:focus { outline: none; border-color: var(--blue); }

    /* ── Tables ── */
    .table-wrap { overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; font-size: 12px; }
    thead th { background: var(--navy); color: #fff; padding: 9px 12px; text-align: left;
               font-weight: 600; font-size: 11px; white-space: nowrap;
               cursor: pointer; user-select: none; }
    thead th:hover { background: var(--blue); }
    thead th.sort-asc::after  { content: ' ↑'; }
    thead th.sort-desc::after { content: ' ↓'; }
    tbody tr { border-bottom: 1px solid var(--border); transition: background .1s; }
    tbody tr:hover { background: var(--section-bg); }
    tbody tr:nth-child(even) { background: var(--surface); }
    tbody tr:nth-child(even):hover { background: var(--section-bg); }
    td { padding: 8px 12px; white-space: nowrap; }
    td.wrap { white-space: normal; max-width: 240px; }
    tfoot td { font-weight: 700; background: var(--light-blue); color: var(--navy);
               padding: 9px 12px; border-top: 2px solid var(--blue); }

    /* ── Cell helpers ── */
    .num { text-align: right; font-variant-numeric: tabular-nums; }
    .tag { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 10px; font-weight: 700; }
    .tag-red    { background: var(--red-light);   color: var(--red); }
    .tag-amber  { background: var(--amber-light); color: var(--amber); }
    .tag-green  { background: var(--green-light); color: var(--green); }
    .tag-blue   { background: var(--light-blue);  color: var(--navy); }
    .bar-cell   { display: flex; align-items: center; gap: 8px; min-width: 120px; }
    .bar        { flex: 1; height: 6px; background: var(--border); border-radius: 3px; overflow: hidden; }
    .bar-fill   { height: 100%; border-radius: 3px; background: var(--blue); transition: width .3s; }
    .bar-fill.red { background: var(--red); }

    /* ── Two-col layout ── */
    .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px; }
    @media (max-width: 900px) { .two-col { grid-template-columns: 1fr; } }

    /* ── S-Drop specific ── */
    .sdrop-loc-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px,1fr));
                      gap: 12px; padding: 16px 20px; border-bottom: 1px solid var(--border); }
    .sdrop-loc-card { background: var(--surface); border: 1px solid var(--border); border-radius: 8px;
                      padding: 14px 16px; }
    .sdrop-loc-name  { font-size: 11px; font-weight: 700; color: var(--navy);
                       text-transform: uppercase; margin-bottom: 6px; }
    .sdrop-loc-items { font-size: 20px; font-weight: 700; color: var(--navy); }
    .sdrop-loc-val   { font-size: 11px; color: var(--muted); margin-top: 2px; }

    /* ── Loading overlay ── */
    #loading { position: fixed; inset: 0; background: rgba(31,56,100,.85);
               display: flex; align-items: center; justify-content: center;
               z-index: 999; color: #fff; font-size: 15px; font-weight: 600; }
    .spinner { width: 36px; height: 36px; border: 3px solid rgba(255,255,255,.3);
               border-top-color: #fff; border-radius: 50%; animation: spin .8s linear infinite;
               margin-right: 16px; }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>
<body>

<div id="loading"><div class="spinner"></div> Loading dashboard…</div>

<!-- Header -->
<div class="header">
  <div class="header-left">
    <div>
      <div style="font-size:17px;font-weight:700;">Aged Inventory Dashboard</div>
      <div class="header-sub" id="headerSub">As of today</div>
    </div>
  </div>
  <a class="upload-link" href="/">↑ Upload New Data</a>
</div>

<!-- Tab navigation -->
<div class="tab-nav">
  <button class="tab-btn active" onclick="switchTab('main', this)">Main Dashboard</button>
  <button class="tab-btn" onclick="switchTab('sdrop', this)" id="sdropTabBtn">
    S-Drop Review <span class="tab-badge" id="sdropBadge">—</span>
  </button>
</div>

<!-- Main dashboard filters -->
<div class="filters-bar" id="mainFilters">
  <div class="filter-group">
    <label>Filter by Project</label>
    <select id="filterProject" onchange="applyFilters()"><option value="">All Projects</option></select>
  </div>
  <div class="filter-group">
    <label>Filter by Coordinator</label>
    <select id="filterCoord" onchange="applyFilters()"><option value="">All Coordinators</option></select>
  </div>
  <div class="filter-group">
    <label>Filter by Age Bucket</label>
    <select id="filterBucket" onchange="applyFilters()">
      <option value="">All Buckets</option>
      <option value="0-7d">0 – 7 Days</option>
      <option value="8-14d">8 – 14 Days</option>
      <option value="15-30d">15 – 30 Days</option>
      <option value="31-60d">31 – 60 Days</option>
      <option value="61-90d">61 – 90 Days</option>
      <option value="91-180d">91 – 180 Days</option>
      <option value="181-365d">181 – 365 Days</option>
      <option value="365d+">Over 365 Days</option>
    </select>
  </div>
  <button class="reset-btn" onclick="resetFilters()">&#10005; Reset</button>
  <button class="reset-btn" style="background:rgba(255,255,255,.2);color:#fff;border-color:rgba(255,255,255,.4);" onclick="exportMainCSV()">&#11015; Export CSV</button>
</div>

<!-- S-Drop filters -->
<div class="filters-bar sdrop-filters" id="sdropFilters">
  <div class="filter-group">
    <label>Filter by Drop Location</label>
    <select id="filterDropLoc" onchange="renderSdrop()"><option value="">All Drop Locations</option></select>
  </div>
  <div class="filter-group">
    <label>Filter by Coordinator</label>
    <select id="filterDropCoord" onchange="renderSdrop()"><option value="">All Coordinators</option></select>
  </div>
  <div class="filter-group">
    <label>Filter by Age Flag</label>
    <select id="filterDropFlag" onchange="renderSdrop()">
      <option value="">All Ages</option>
      <option value=">90 Days">Over 90 Days</option>
      <option value=">30 Days">Over 30 Days</option>
      <option value="&#8804;30 Days">30 Days or Less</option>
    </select>
  </div>
  <button class="reset-btn" onclick="resetSdropFilters()">&#10005; Reset</button>
  <button class="reset-btn" style="background:rgba(255,255,255,.2);color:#fff;border-color:rgba(255,255,255,.4);" onclick="exportSdropCSV()">&#11015; Export CSV</button>
</div>

<div class="main">

  <!-- ══ MAIN TAB ══════════════════════════════════════════════════════════ -->
  <div class="tab-panel active" id="panel-main">

    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-label">Total Inventory Value</div><div class="kpi-value" id="kTotal">—</div><div class="kpi-sub" id="kContainers">—</div></div>
      <div class="kpi-card accent-red"><div class="kpi-label">Over 90 Days Value</div><div class="kpi-value" id="kOver90">—</div><div class="kpi-sub" id="kOver90Pct">—</div></div>
      <div class="kpi-card accent-amber"><div class="kpi-label">In Storage Value</div><div class="kpi-value" id="kStorage">—</div></div>
      <div class="kpi-card accent-amber"><div class="kpi-label">Finance Hold Value</div><div class="kpi-value" id="kFinance">—</div></div>
      <div class="kpi-card accent-green"><div class="kpi-label">Average Age</div><div class="kpi-value" id="kAvgAge">—</div><div class="kpi-sub">days</div></div>
    </div>

    <div class="two-col">
      <div class="section-card">
        <div class="section-header">
          <div class="section-title">Inventory by Project Coordinator</div>
          <div class="section-meta" id="coordMeta"></div>
        </div>
        <div class="table-wrap">
          <table id="coordTable">
            <thead><tr>
              <th onclick="sortTable('coordTable',0)">Coordinator</th>
              <th class="num" onclick="sortTable('coordTable',1)">Containers</th>
              <th class="num" onclick="sortTable('coordTable',2)">Value</th>
              <th class="num" onclick="sortTable('coordTable',3)">Over 90d Value</th>
              <th class="num" onclick="sortTable('coordTable',4)">% Over 90d</th>
              <th class="num" onclick="sortTable('coordTable',5)">Avg Age</th>
            </tr></thead>
            <tbody id="coordBody"></tbody>
            <tfoot id="coordFoot"></tfoot>
          </table>
        </div>
      </div>

      <div class="section-card">
        <div class="section-header">
          <div class="section-title">Inventory by Age Bucket</div>
        </div>
        <div class="table-wrap">
          <table id="ageTable">
            <thead><tr>
              <th onclick="sortTable('ageTable',0)">Age Bucket</th>
              <th class="num" onclick="sortTable('ageTable',1)">Containers</th>
              <th class="num" onclick="sortTable('ageTable',2)">Value</th>
              <th onclick="sortTable('ageTable',3)">% of Total</th>
            </tr></thead>
            <tbody id="ageBody"></tbody>
            <tfoot id="ageFoot"></tfoot>
          </table>
        </div>
      </div>
    </div>

    <div class="section-card">
      <div class="section-header">
        <div class="section-title">Orders in View</div>
        <div class="section-meta" id="ordersMeta"></div>
      </div>
      <div class="search-box">
        <input type="text" id="orderSearch" placeholder="Search by order #, project, or coordinator…" oninput="renderOrders()">
      </div>
      <div class="table-wrap">
        <table id="ordersTable">
          <thead><tr>
            <th onclick="sortTable('ordersTable',0)">Order #</th>
            <th onclick="sortTable('ordersTable',1)">Project</th>
            <th onclick="sortTable('ordersTable',2)">Coordinator</th>
            <th class="num" onclick="sortTable('ordersTable',3)">Containers</th>
            <th class="num" onclick="sortTable('ordersTable',4)">Value</th>
            <th class="num" onclick="sortTable('ordersTable',5)">Avg Age</th>
            <th class="num" onclick="sortTable('ordersTable',6)">Max Age</th>
            <th onclick="sortTable('ordersTable',7)">Age Bucket</th>
            <th onclick="sortTable('ordersTable',8)">Ship Date Range</th>
          </tr></thead>
          <tbody id="ordersBody"></tbody>
          <tfoot id="ordersFoot"></tfoot>
        </table>
      </div>
    </div>

  </div><!-- /panel-main -->

  <!-- ══ S-DROP TAB ════════════════════════════════════════════════════════ -->
  <div class="tab-panel" id="panel-sdrop">

    <div class="kpi-grid" style="grid-template-columns: repeat(5, 1fr);">
      <div class="kpi-card accent-red"><div class="kpi-label">Total Items in Drop</div><div class="kpi-value" id="sdKpiItems">—</div><div class="kpi-sub">over 2 days</div></div>
      <div class="kpi-card accent-red"><div class="kpi-label">Total Value</div><div class="kpi-value" id="sdKpiValue">—</div></div>
      <div class="kpi-card"><div class="kpi-label">Unique Orders</div><div class="kpi-value" id="sdKpiOrders">—</div></div>
      <div class="kpi-card accent-amber"><div class="kpi-label">Avg Age</div><div class="kpi-value" id="sdKpiAvg">—</div><div class="kpi-sub">days</div></div>
      <div class="kpi-card accent-amber"><div class="kpi-label">Max Age</div><div class="kpi-value" id="sdKpiMax">—</div><div class="kpi-sub">days</div></div>
    </div>

    <div class="section-card">
      <div class="section-header">
        <div class="section-title">By Drop Location</div>
      </div>
      <div class="sdrop-loc-grid" id="sdropLocGrid"></div>
    </div>

    <div class="section-card">
      <div class="section-header">
        <div class="section-title">Item Detail</div>
        <div class="section-meta" id="sdropMeta"></div>
      </div>
      <div class="search-box">
        <input type="text" id="sdropSearch" placeholder="Search by order #, project, part no…" oninput="renderSdrop()">
      </div>
      <div class="table-wrap">
        <table id="sdropTable">
          <thead><tr>
            <th onclick="sortTable('sdropTable',0)">Location</th>
            <th onclick="sortTable('sdropTable',1)">Order #</th>
            <th onclick="sortTable('sdropTable',2)">Project</th>
            <th onclick="sortTable('sdropTable',3)">Coordinator</th>
            <th onclick="sortTable('sdropTable',4)">Part No.</th>
            <th onclick="sortTable('sdropTable',5)">Part Group</th>
            <th class="num" onclick="sortTable('sdropTable',6)">Qty</th>
            <th class="num" onclick="sortTable('sdropTable',7)">Age (days)</th>
            <th class="num" onclick="sortTable('sdropTable',8)">Value</th>
            <th onclick="sortTable('sdropTable',9)">Order Status</th>
            <th onclick="sortTable('sdropTable',10)">Age Flag</th>
          </tr></thead>
          <tbody id="sdropBody"></tbody>
          <tfoot id="sdropFoot"></tfoot>
        </table>
      </div>
    </div>

  </div><!-- /panel-sdrop -->

</div><!-- /main -->

<script>
// ── Globals ───────────────────────────────────────────────────────────────────
let RAW = null;
let filteredOrders = [];
const sortState = {};

const AGE_ORDER  = ['0-7d','8-14d','15-30d','31-60d','61-90d','91-180d','181-365d','365d+'];
const AGE_LABELS = {'0-7d':'0–7d','8-14d':'8–14d','15-30d':'15–30d','31-60d':'31–60d',
  '61-90d':'61–90d','91-180d':'91–180d','181-365d':'181–365d','365d+':'>365d'};

// ── Formatters ────────────────────────────────────────────────────────────────
const fmt$  = v => v == null ? '—' : '$' + Math.round(v).toLocaleString();
const fmtN  = v => v == null ? '—' : Math.round(v).toLocaleString();
const fmtPct= v => v == null ? '—' : (v*100).toFixed(1)+'%';
const fmtF  = v => v == null ? '—' : v.toFixed(1);

function ageBucketTag(code) {
  if (!code) return '';
  const cls = (code==='365d+'||code==='181-365d') ? 'tag-red'
            : code==='91-180d' ? 'tag-amber' : 'tag-blue';
  return `<span class="tag ${cls}">${AGE_LABELS[code]||code}</span>`;
}
function maxAgeTag(age) {
  const cls = age>90 ? 'tag-red' : age>30 ? 'tag-amber' : 'tag-green';
  return `<span class="tag ${cls}">${age}d</span>`;
}
function flagTag(flag) {
  const cls = flag==='>90 Days' ? 'tag-red' : flag==='>30 Days' ? 'tag-amber' : 'tag-green';
  return `<span class="tag ${cls}">${flag}</span>`;
}

// ── Tab switching ─────────────────────────────────────────────────────────────
function switchTab(name, btn) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-' + name).classList.add('active');
  btn.classList.add('active');
  document.getElementById('mainFilters').style.display  = name==='main'  ? '' : 'none';
  document.getElementById('sdropFilters').style.display = name==='sdrop' ? '' : 'none';

}

// ── Load data ─────────────────────────────────────────────────────────────────
async function loadData() {
  try {
    // 10 second timeout — if server doesn't respond, show error
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 10000);
    const res = await fetch('/api/data', { signal: controller.signal });
    clearTimeout(timeout);
    if (!res.ok) throw new Error('no_data');
    RAW = await res.json();
  } catch(e) {
    document.getElementById('loading').innerHTML = `
      <div style="text-align:center;padding:40px">
        <div style="font-size:48px;margin-bottom:16px">📂</div>
        <div style="font-size:20px;font-weight:700;margin-bottom:8px">No data loaded yet</div>
        <div style="font-size:14px;opacity:.8;margin-bottom:24px">Upload your warehouse and internal notes files to get started.</div>
        <a href="/" style="background:#fff;color:#1F3864;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">↑ Upload Files</a>
      </div>`;
    return;
  }

  if (!RAW.sdrop) RAW.sdrop = { kpis:{}, by_location:[], items:[] };

  // Set the uploaded date in the header
  if (RAW.uploaded) {
    document.getElementById('headerSub').textContent = 'As of ' + RAW.uploaded;
  }

  // Populate main filters
  RAW.projects.forEach(p => document.getElementById('filterProject').appendChild(new Option(p, p)));
  RAW.coordinators.forEach(c => document.getElementById('filterCoord').appendChild(new Option(c, c)));

  // Populate S-Drop filters
  (RAW.drop_locations || []).forEach(l => document.getElementById('filterDropLoc').appendChild(new Option(l, l)));
  const sdropCoords = [...new Set((RAW.sdrop.items || []).map(r => r.coordinator))].sort();
  sdropCoords.forEach(c => document.getElementById('filterDropCoord').appendChild(new Option(c, c)));

  // Update S-Drop badge
  document.getElementById('sdropBadge').textContent = RAW.sdrop?.kpis?.total_items ?? 0;

  applyFilters();
  renderSdropKpis();
  renderSdropLocGrid();
  renderSdrop();
  document.getElementById('loading').style.display = 'none';
}

// ── Main filters ──────────────────────────────────────────────────────────────
function getFilters() {
  return {
    project: document.getElementById('filterProject').value,
    coord:   document.getElementById('filterCoord').value,
    bucket:  document.getElementById('filterBucket').value,
  };
}
function resetFilters() {
  ['filterProject','filterCoord','filterBucket'].forEach(id => document.getElementById(id).value='');
  document.getElementById('orderSearch').value = '';
  applyFilters();
}

function applyFilters() {
  if (!RAW) return;
  const {project, coord, bucket} = getFilters();
  filteredOrders = RAW.orders_table.filter(o =>
    (!project || o.project === project) &&
    (!coord   || o.coordinator === coord) &&
    (!bucket  || o.age_bucket === bucket)
  );
  renderKPIs(); renderCoordTable(); renderAgeTable(); renderOrders();
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function renderKPIs() {
  const {project, coord, bucket} = getFilters();
  const hasFilter = project || coord || bucket;
  if (!hasFilter) {
    const k = RAW.kpis;
    document.getElementById('kTotal').textContent      = fmt$(k.total_value);
    document.getElementById('kContainers').textContent = fmtN(k.total_containers) + ' containers';
    document.getElementById('kOver90').textContent     = fmt$(k.over90_value);
    document.getElementById('kOver90Pct').textContent  = fmtPct(k.over90_pct) + ' of total';
    document.getElementById('kStorage').textContent    = fmt$(k.in_storage_value);
    document.getElementById('kFinance').textContent    = fmt$(k.finance_hold_value);
    document.getElementById('kAvgAge').textContent     = fmtF(k.avg_age);
  } else {
    let tv=0, tc=0, o90=0, ta=0;
    filteredOrders.forEach(o => { tv+=o.value; tc+=o.containers; if(o.max_age>90) o90+=o.value; ta+=o.avg_age*o.containers; });
    document.getElementById('kTotal').textContent      = fmt$(tv);
    document.getElementById('kContainers').textContent = fmtN(tc) + ' containers';
    document.getElementById('kOver90').textContent     = fmt$(o90);
    document.getElementById('kOver90Pct').textContent  = tv ? fmtPct(o90/tv)+' of total' : '—';
    document.getElementById('kStorage').textContent    = '—';
    document.getElementById('kFinance').textContent    = '—';
    document.getElementById('kAvgAge').textContent     = tc ? (ta/tc).toFixed(1) : '—';
  }
}

// ── Coordinator table ─────────────────────────────────────────────────────────
function renderCoordTable() {
  const {project, coord, bucket} = getFilters();
  let rows;
  if (!project && !coord && !bucket) {
    rows = RAW.coord_table;
  } else {
    const map = {};
    filteredOrders.forEach(o => {
      if (!map[o.coordinator]) map[o.coordinator]={coordinator:o.coordinator,containers:0,value:0,over90_value:0,ages:[]};
      map[o.coordinator].containers+=o.containers; map[o.coordinator].value+=o.value;
      if(o.max_age>90) map[o.coordinator].over90_value+=o.value;
      map[o.coordinator].ages.push(...Array(o.containers).fill(o.avg_age));
    });
    rows = Object.values(map).map(r=>({...r,
      over90_pct: r.value ? r.over90_value/r.value : 0,
      avg_age: r.ages.length ? r.ages.reduce((a,b)=>a+b,0)/r.ages.length : 0,
    })).sort((a,b)=>a.coordinator.localeCompare(b.coordinator));
  }
  document.getElementById('coordMeta').textContent = rows.length + ' coordinators';
  document.getElementById('coordBody').innerHTML = rows.map(r=>`
    <tr>
      <td>${r.coordinator}</td>
      <td class="num">${fmtN(r.containers)}</td>
      <td class="num">${fmt$(r.value)}</td>
      <td class="num">${fmt$(r.over90_value)}</td>
      <td><div class="bar-cell"><div class="bar"><div class="bar-fill red" style="width:${Math.min(r.over90_pct*100,100).toFixed(1)}%"></div></div><span>${fmtPct(r.over90_pct)}</span></div></td>
      <td class="num">${fmtF(r.avg_age)}</td>
    </tr>`).join('');
  const tot = rows.reduce((a,r)=>({containers:a.containers+r.containers,value:a.value+r.value,over90_value:a.over90_value+r.over90_value}),{containers:0,value:0,over90_value:0});
  document.getElementById('coordFoot').innerHTML = `<tr><td>TOTAL</td><td class="num">${fmtN(tot.containers)}</td><td class="num">${fmt$(tot.value)}</td><td class="num">${fmt$(tot.over90_value)}</td><td class="num">${tot.value?fmtPct(tot.over90_value/tot.value):'—'}</td><td></td></tr>`;
}

// ── Age bucket table ──────────────────────────────────────────────────────────
function renderAgeTable() {
  const {project, coord, bucket} = getFilters();
  let rows;
  if (!project && !coord && !bucket) {
    rows = RAW.age_table;
  } else {
    const map = {}; AGE_ORDER.forEach(k=>map[k]={bucket:k,label:k,containers:0,value:0});
    filteredOrders.forEach(o=>{if(map[o.age_bucket]){map[o.age_bucket].containers+=o.containers;map[o.age_bucket].value+=o.value;}});
    const tv=Object.values(map).reduce((a,r)=>a+r.value,0);
    rows=AGE_ORDER.map(k=>({...map[k],pct:tv?map[k].value/tv:0}));
  }
  const tv=rows.reduce((a,r)=>a+r.value,0);
  document.getElementById('ageBody').innerHTML=rows.map(r=>`
    <tr>
      <td>${r.label||AGE_LABELS[r.bucket]}</td>
      <td class="num">${fmtN(r.containers)}</td>
      <td class="num">${fmt$(r.value)}</td>
      <td><div class="bar-cell"><div class="bar"><div class="bar-fill" style="width:${Math.min((r.pct||0)*100,100).toFixed(1)}%"></div></div><span>${fmtPct(r.pct)}</span></div></td>
    </tr>`).join('');
  document.getElementById('ageFoot').innerHTML=`<tr><td>TOTAL</td><td class="num">${fmtN(rows.reduce((a,r)=>a+r.containers,0))}</td><td class="num">${fmt$(tv)}</td><td></td></tr>`;
}

// ── Orders table ──────────────────────────────────────────────────────────────
function renderOrders() {
  const q = document.getElementById('orderSearch').value.toLowerCase();
  let rows = filteredOrders;
  if (q) rows=rows.filter(o=>o.order.toLowerCase().includes(q)||o.project.toLowerCase().includes(q)||o.coordinator.toLowerCase().includes(q));
  document.getElementById('ordersMeta').textContent = rows.length+' orders';
  document.getElementById('ordersBody').innerHTML = rows.map(o=>`
    <tr>
      <td><strong>${o.order}</strong></td>
      <td class="wrap">${o.project}</td>
      <td>${o.coordinator}</td>
      <td class="num">${fmtN(o.containers)}</td>
      <td class="num">${fmt$(o.value)}</td>
      <td class="num">${fmtF(o.avg_age)}</td>
      <td class="num">${maxAgeTag(o.max_age)}</td>
      <td>${ageBucketTag(o.age_bucket)}</td>
      <td>${o.ship_range||'—'}</td>
    </tr>`).join('') || '<tr><td colspan="9" style="text-align:center;padding:32px;color:var(--muted)">No orders match filters.</td></tr>';
  if (!rows.length) { document.getElementById('ordersFoot').innerHTML=''; return; }
  const tot=rows.reduce((a,o)=>({containers:a.containers+o.containers,value:a.value+o.value}),{containers:0,value:0});
  document.getElementById('ordersFoot').innerHTML=`<tr><td colspan="3">TOTAL — ${rows.length} orders</td><td class="num">${fmtN(tot.containers)}</td><td class="num">${fmt$(tot.value)}</td><td colspan="4"></td></tr>`;
}

// ── S-Drop KPIs ───────────────────────────────────────────────────────────────
function renderSdropKpis() {
  const k = RAW.sdrop?.kpis;
  if (!k) return;
  document.getElementById('sdKpiItems').textContent  = fmtN(k.total_items);
  document.getElementById('sdKpiValue').textContent  = fmt$(k.total_value);
  document.getElementById('sdKpiOrders').textContent = fmtN(k.unique_orders);
  document.getElementById('sdKpiAvg').textContent    = fmtF(k.avg_age);
  document.getElementById('sdKpiMax').textContent    = fmtN(k.max_age);
}

// ── S-Drop location cards ─────────────────────────────────────────────────────
function renderSdropLocGrid() {
  const locs = RAW.sdrop?.by_location || [];
  document.getElementById('sdropLocGrid').innerHTML = locs.map(l=>`
    <div class="sdrop-loc-card">
      <div class="sdrop-loc-name">${l.location}</div>
      <div class="sdrop-loc-items">${fmtN(l.items)} <span style="font-size:13px;font-weight:400;color:var(--muted)">items</span></div>
      <div class="sdrop-loc-val">${fmt$(l.value)} · avg ${fmtF(l.avg_age)}d · max ${fmtN(l.max_age)}d</div>
    </div>`).join('');
}

// ── S-Drop detail table ───────────────────────────────────────────────────────
function resetSdropFilters() {
  document.getElementById('filterDropLoc').value   = '';
  document.getElementById('filterDropCoord').value = '';
  document.getElementById('filterDropFlag').value  = '';
  document.getElementById('sdropSearch').value     = '';
  renderSdrop();
}

function renderSdrop() {
  const items = RAW.sdrop?.items || [];
  const loc   = document.getElementById('filterDropLoc').value;
  const coord = document.getElementById('filterDropCoord').value;
  const flag  = document.getElementById('filterDropFlag').value;
  const q     = document.getElementById('sdropSearch').value.toLowerCase();

  let rows = items;
  if (loc)   rows = rows.filter(r => r.location === loc);
  if (coord) rows = rows.filter(r => r.coordinator === coord);
  if (flag)  rows = rows.filter(r => r.flag === flag);
  if (q)    rows = rows.filter(r =>
    r.order.toLowerCase().includes(q) ||
    r.project.toLowerCase().includes(q) ||
    r.part_no.toLowerCase().includes(q) ||
    r.coordinator.toLowerCase().includes(q)
  );

  document.getElementById('sdropMeta').textContent = rows.length + ' items';
  document.getElementById('sdropBody').innerHTML = rows.map(r=>`
    <tr>
      <td>${r.location}</td>
      <td><strong>${r.order}</strong></td>
      <td class="wrap">${r.project}</td>
      <td>${r.coordinator}</td>
      <td>${r.part_no}</td>
      <td>${r.part_group}</td>
      <td class="num">${fmtN(r.qty)}</td>
      <td class="num">${maxAgeTag(r.age)}</td>
      <td class="num">${fmt$(r.value)}</td>
      <td>${r.status}</td>
      <td>${flagTag(r.flag)}</td>
    </tr>`).join('') || '<tr><td colspan="11" style="text-align:center;padding:32px;color:var(--muted)">No items match filters.</td></tr>';

  if (!rows.length) { document.getElementById('sdropFoot').innerHTML=''; return; }
  const totQty=rows.reduce((a,r)=>a+r.qty,0), totVal=rows.reduce((a,r)=>a+r.value,0);
  document.getElementById('sdropFoot').innerHTML=`<tr><td colspan="6">TOTAL — ${rows.length} items</td><td class="num">${fmtN(totQty)}</td><td></td><td class="num">${fmt$(totVal)}</td><td colspan="2"></td></tr>`;
}

// ── Sorting ───────────────────────────────────────────────────────────────────
function sortTable(tableId, colIdx) {
  const key = tableId+'_'+colIdx;
  const asc = sortState[key] !== 'asc';
  sortState[key] = asc ? 'asc' : 'desc';
  const table = document.getElementById(tableId);
  table.querySelectorAll('thead th').forEach((th,i)=>{
    th.classList.remove('sort-asc','sort-desc');
    if(i===colIdx) th.classList.add(asc?'sort-asc':'sort-desc');
  });
  const tbody = table.querySelector('tbody');
  Array.from(tbody.querySelectorAll('tr')).sort((a,b)=>{
    let av=(a.cells[colIdx]?.textContent||'').replace(/[$,%\\s↑↓]/g,'').trim();
    let bv=(b.cells[colIdx]?.textContent||'').replace(/[$,%\\s↑↓]/g,'').trim();
    const an=parseFloat(av),bn=parseFloat(bv);
    const cmp=(!isNaN(an)&&!isNaN(bn))?an-bn:av.localeCompare(bv);
    return asc?cmp:-cmp;
  }).forEach(r=>tbody.appendChild(r));
}

// ── CSV Export ───────────────────────────────────────────────────────────────
function toCSV(headers, rows) {
  const escape = v => {
    if (v == null) return '';
    const s = String(v);
    return s.includes(',') || s.includes('"') || s.includes('\n') ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [headers.map(escape).join(',')];
  rows.forEach(r => lines.push(r.map(escape).join(',')));
  return lines.join('\n');
}

function downloadCSV(filename, csv) {
  const blob = new Blob([csv], {type: 'text/csv'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function exportMainCSV() {
  const headers = ['Order #','Project','Coordinator','Containers','Value','Avg Age','Max Age','Age Bucket','Status','Ship Date Range'];
  const rows = filteredOrders.map(o => [
    o.order, o.project, o.coordinator, o.containers,
    o.value, o.avg_age, o.max_age, o.age_bucket, o.status, o.ship_range
  ]);
  const proj  = document.getElementById('filterProject').value || 'All Projects';
  const coord = document.getElementById('filterCoord').value   || 'All Coordinators';
  const bucket= document.getElementById('filterBucket').value  || 'All Buckets';
  const label = [proj, coord, bucket].filter(x => x && !x.startsWith('All')).join('_') || 'All';
  downloadCSV('Aged_Inventory_' + label + '.csv', toCSV(headers, rows));
}

function exportSdropCSV() {
  const items = RAW.sdrop?.items || [];
  const loc   = document.getElementById('filterDropLoc').value;
  const coord = document.getElementById('filterDropCoord').value;
  const flag  = document.getElementById('filterDropFlag').value;
  const q     = document.getElementById('sdropSearch').value.toLowerCase();
  let rows = items;
  if (loc)   rows = rows.filter(r => r.location === loc);
  if (coord) rows = rows.filter(r => r.coordinator === coord);
  if (flag)  rows = rows.filter(r => r.flag === flag);
  if (q)     rows = rows.filter(r =>
    r.order.toLowerCase().includes(q) || r.project.toLowerCase().includes(q) ||
    r.part_no.toLowerCase().includes(q) || r.coordinator.toLowerCase().includes(q));
  const headers = ['Location','Order #','Project','Coordinator','Part No.','Part Group','Qty','Age (days)','Value','Order Status','Age Flag'];
  const csvRows = rows.map(r => [r.location, r.order, r.project, r.coordinator, r.part_no, r.part_group, r.qty, r.age, r.value, r.status, r.flag]);
  downloadCSV('SDrop_Review.csv', toCSV(headers, csvRows));
}

// ── Init ──────────────────────────────────────────────────────────────────────
loadData();
</script>
</body>
</html>
"""

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
            'sdrop_items': data['sdrop']['kpis']['total_items'],
            'uploaded': data['uploaded'],
        }))
        return jsonify({'success': True})
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500

@app.route('/dashboard')
def dashboard():
    if not PROCESSED_PATH.exists():
        return redirect('/')
    return Response(DASHBOARD_HTML, mimetype='text/html')

@app.route('/api/data')
def api_data():
    if not PROCESSED_PATH.exists():
        return jsonify({'error': 'No data — please upload files first'}), 404
    return PROCESSED_PATH.read_text(), 200, {'Content-Type': 'application/json'}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

import os, json
from pathlib import Path
from flask import Flask, request, jsonify, redirect, url_for, Response
import pandas as pd
import numpy as np

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

DATA_DIR = Path('data')
DATA_DIR.mkdir(exist_ok=True)
PROCESSED_PATH = DATA_DIR / 'processed.json'
META_PATH      = DATA_DIR / 'meta.json'

def read_excel_fast(path):
    try:
        return pd.read_excel(path, engine='calamine')
    except Exception:
        return pd.read_excel(path, engine='openpyxl')

AGE_BUCKET_ORDER = ['0-7d','8-14d','15-30d','31-60d','61-90d','91-180d','181-365d','365d+']
AGE_BUCKET_LABELS = {
    '0-7d':'0 – 7 Days','8-14d':'8 – 14 Days','15-30d':'15 – 30 Days',
    '31-60d':'31 – 60 Days','61-90d':'61 – 90 Days','91-180d':'91 – 180 Days',
    '181-365d':'181 – 365 Days','365d+':'Over 365 Days'
}

def assign_age_buckets(series):
    age = series
    result = pd.Series('365d+', index=series.index, dtype=str)
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

    coord_map = (notes[['Order No','Project Coordinator']]
                 .drop_duplicates('Order No')
                 .dropna(subset=['Project Coordinator']))
    df = df.merge(coord_map, left_on='ORDERS', right_on='Order No', how='left')
    df.rename(columns={'Project Coordinator':'COORDINATOR'}, inplace=True)
    df['SHIP_DATE'] = pd.to_datetime(df['SHIP_DATE'], errors='coerce')

    total_value  = float(df['EXTENDED_COST'].sum())
    over90_mask  = df['INV_AGE'] > 90
    over90_value = float(df.loc[over90_mask,'EXTENDED_COST'].sum())
    avg_age      = float(df['INV_AGE'].mean()) if len(df) else 0
    in_storage   = df[df['LOCATION_GROUP'].str.contains('Storage',case=False,na=False)]
    finance_hold = df[df['ORDER_STATUS'].str.contains('Finance Hold',case=False,na=False)]

    kpis = {
        'total_value': total_value, 'over90_value': over90_value,
        'over90_pct': over90_value/total_value if total_value else 0,
        'total_containers': len(df), 'avg_age': round(avg_age,1),
        'in_storage_value': float(in_storage['EXTENDED_COST'].sum()),
        'finance_hold_value': float(finance_hold['EXTENDED_COST'].sum()),
    }

    df_coord = df[df['COORDINATOR'].notna()].copy()
    coord_tbl = []
    for coord, grp in df_coord.groupby('COORDINATOR'):
        o90 = grp[grp['INV_AGE'] > 90]
        val = float(grp['EXTENDED_COST'].sum())
        coord_tbl.append({
            'coordinator': coord, 'containers': len(grp), 'value': round(val,2),
            'over90_value': round(float(o90['EXTENDED_COST'].sum()),2),
            'over90_pct': float(o90['EXTENDED_COST'].sum()/val) if val else 0,
            'avg_age': round(float(grp['INV_AGE'].mean()),1),
        })

    age_tbl = []
    for code in AGE_BUCKET_ORDER:
        mask = df['AGE_BUCKET'] == code
        val  = float(df.loc[mask,'EXTENDED_COST'].sum())
        age_tbl.append({'bucket':code,'label':AGE_BUCKET_LABELS[code],
                        'containers':int(mask.sum()),'value':round(val,2),
                        'pct':val/total_value if total_value else 0})

    valid = df[df['ORDERS'].notna() & (df['ORDERS'].astype(str) != '.')].copy()
    valid['SHIP_TS'] = pd.to_datetime(valid['SHIP_DATE'], errors='coerce')
    orders_tbl = []
    for order, grp in valid.groupby('ORDERS'):
        ship = grp['SHIP_TS'].dropna()
        sr = ''
        if len(ship):
            mn,mx = ship.min().strftime('%m/%d/%y'),ship.max().strftime('%m/%d/%y')
            sr = mn if mn==mx else f'{mn} – {mx}'
        coord = grp['COORDINATOR'].dropna().iloc[0] if grp['COORDINATOR'].notna().any() else 'Unassigned'
        orders_tbl.append({
            'order':str(order),'project':str(grp['PROJECT_NAME'].iloc[0] or ''),
            'coordinator':str(coord),'containers':len(grp),
            'value':round(float(grp['EXTENDED_COST'].sum()),2),
            'avg_age':round(float(grp['INV_AGE'].mean()),1),'max_age':int(grp['INV_AGE'].max()),
            'age_bucket':str(grp['AGE_BUCKET'].mode().iloc[0]) if len(grp) else '',
            'status':str(grp['ORDER_STATUS'].mode().iloc[0]) if len(grp) else '',
            'ship_range':sr,
        })
    orders_tbl.sort(key=lambda x: x['project'])

    sdrop_raw = df[df['LOCATION'].str.contains('drop',case=False,na=False) & (df['INV_AGE']>2)].copy()
    sdrop_raw = sdrop_raw.sort_values('INV_AGE',ascending=False)
    loc_sum = (sdrop_raw.groupby('LOCATION')
               .agg(items=('LOCATION','count'),value=('EXTENDED_COST','sum'),
                    avg_age=('INV_AGE','mean'),max_age=('INV_AGE','max'))
               .reset_index().sort_values('items',ascending=False))
    sdrop_by_location = [{'location':r['LOCATION'],'items':int(r['items']),
        'value':round(float(r['value']),2),'avg_age':round(float(r['avg_age']),1),
        'max_age':int(r['max_age'])} for _,r in loc_sum.iterrows()]
    sdrop_items = []
    for _,row in sdrop_raw.iterrows():
        age = int(row['INV_AGE'])
        flag = '>90 Days' if age>90 else ('>30 Days' if age>30 else '≤30 Days')
        sdrop_items.append({
            'location':str(row['LOCATION']),
            'order':str(row['ORDERS']) if pd.notna(row['ORDERS']) and str(row['ORDERS'])!='.' else '—',
            'project':str(row['PROJECT_NAME']) if pd.notna(row['PROJECT_NAME']) and str(row['PROJECT_NAME'])!='.' else '—',
            'coordinator':str(row['COORDINATOR']) if pd.notna(row['COORDINATOR']) else 'Unassigned',
            'part_no':str(row['PART_NO']) if pd.notna(row['PART_NO']) else '—',
            'part_group':str(row['PART_GROUP']) if pd.notna(row['PART_GROUP']) and str(row['PART_GROUP'])!='.' else '—',
            'qty':float(row['QUANTITY']) if pd.notna(row['QUANTITY']) else 0,
            'age':age,'value':round(float(row['EXTENDED_COST']),2) if pd.notna(row['EXTENDED_COST']) else 0,
            'status':str(row['ORDER_STATUS']) if pd.notna(row['ORDER_STATUS']) and str(row['ORDER_STATUS'])!='.' else '—',
            'flag':flag,
        })
    uniq = sdrop_raw[sdrop_raw['ORDERS'].notna()&(sdrop_raw['ORDERS'].astype(str)!='.')]['ORDERS'].nunique()
    sdrop = {'kpis':{'total_items':len(sdrop_raw),'total_value':round(float(sdrop_raw['EXTENDED_COST'].sum()),2),
        'unique_orders':int(uniq),'avg_age':round(float(sdrop_raw['INV_AGE'].mean()),1) if len(sdrop_raw) else 0,
        'max_age':int(sdrop_raw['INV_AGE'].max()) if len(sdrop_raw) else 0},
        'by_location':sdrop_by_location,'items':sdrop_items}

    return {
        'kpis':kpis,'coord_table':coord_tbl,'age_table':age_tbl,
        'orders_table':orders_tbl,'sdrop':sdrop,
        'projects':sorted(df['PROJECT_NAME'].dropna().unique().tolist()),
        'coordinators':sorted(df['COORDINATOR'].dropna().unique().tolist()),
        'drop_locations':sorted(sdrop_raw['LOCATION'].unique().tolist()),
        'uploaded': pd.Timestamp.now().strftime('%B %d, %Y'),
    }

# ── HTML pages (inline — no templates folder needed) ─────────────────────────

INDEX_HTML = r"""<!DOCTYPE html>
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
.drop-zone input[type=file]{display:none}
.file-chosen{font-size:12px;color:#2E75B6;margin-top:6px;font-weight:600}
.btn{width:100%;padding:13px;background:#1F3864;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;transition:background .2s;margin-top:8px}
.btn:hover{background:#2E75B6}
.btn:disabled{background:#94a3b8;cursor:not-allowed}
.view-link{display:block;text-align:center;margin-top:20px;font-size:13px;color:#2E75B6;text-decoration:none;font-weight:600}
.view-link:hover{text-decoration:underline}
.progress{display:none;margin-top:20px}
.progress-bar{height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden}
.progress-fill{height:100%;background:linear-gradient(90deg,#1F3864,#2E75B6);width:0;transition:width .5s ease;border-radius:4px}
.progress-steps{margin-top:14px;display:flex;flex-direction:column;gap:6px}
.step{display:flex;align-items:center;gap:10px;font-size:12px;color:#94a3b8;transition:color .3s}
.step.active{color:#1F3864;font-weight:600}
.step.done{color:#375623}
.step-icon{width:18px;height:18px;border-radius:50%;border:2px solid #cbd5e1;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .3s;font-size:10px}
.step.active .step-icon{border-color:#2E75B6;background:#2E75B6;color:#fff}
.step.done .step-icon{border-color:#375623;background:#375623;color:#fff}
.step.done .step-icon::after{content:'✓';font-size:10px}
.error-msg{background:#fee2e2;border:1px solid #fca5a5;border-radius:8px;padding:10px 14px;font-size:13px;color:#991b1b;margin-top:12px;display:none}
</style></head><body>
<div class="card">
  <div class="logo">
    <div class="logo-icon"><svg viewBox="0 0 24 24"><path d="M20 7H4a2 2 0 00-2 2v9a2 2 0 002 2h16a2 2 0 002-2V9a2 2 0 00-2-2z"/><path d="M16 3H8a2 2 0 00-2 2v2h12V5a2 2 0 00-2-2z"/></svg></div>
    <div><h1>Aged Inventory Dashboard</h1><div class="subtitle">Upload today's files to refresh the dashboard</div></div>
  </div>
  __STATUS__
  <form id="uploadForm">
    <div class="drop-zone" id="warehouseZone" onclick="document.getElementById('warehouseFile').click()">
      <input type="file" id="warehouseFile" accept=".xlsx" onchange="fileChosen(this,'warehouseZone','warehouseLabel')">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><polyline points="9 15 12 12 15 15"/></svg>
      <div class="label">Warehouse File (.xlsx)</div><div class="sublabel">Click to browse or drag and drop</div>
      <div class="file-chosen" id="warehouseLabel"></div>
    </div>
    <div class="drop-zone" id="notesZone" onclick="document.getElementById('notesFile').click()">
      <input type="file" id="notesFile" accept=".csv" onchange="fileChosen(this,'notesZone','notesLabel')">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>
      <div class="label">Internal Notes (.csv)</div><div class="sublabel">Click to browse or drag and drop</div>
      <div class="file-chosen" id="notesLabel"></div>
    </div>
    <div class="progress" id="progress">
      <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
      <div class="progress-steps">
        <div class="step" id="step1"><div class="step-icon"></div><span>Uploading files to server…</span></div>
        <div class="step" id="step2"><div class="step-icon"></div><span>Reading warehouse data (takes ~10 seconds)…</span></div>
        <div class="step" id="step3"><div class="step-icon"></div><span>Joining coordinator data…</span></div>
        <div class="step" id="step4"><div class="step-icon"></div><span>Building dashboard tables…</span></div>
        <div class="step" id="step5"><div class="step-icon"></div><span>Done! Loading dashboard…</span></div>
      </div>
    </div>
    <div class="error-msg" id="errorMsg"></div>
    <button class="btn" type="submit" id="uploadBtn" disabled>Upload &amp; Process</button>
  </form>
  __VIEWLINK__
</div>
<script>
function fileChosen(input,zoneId,labelId){
  document.getElementById(labelId).textContent=input.files[0]?'✓ '+input.files[0].name:'';
  checkReady();
}
function checkReady(){
  var wh=document.getElementById('warehouseFile').files[0];
  var nt=document.getElementById('notesFile').files[0];
  document.getElementById('uploadBtn').disabled=!(wh&&nt);
}
['warehouseZone','notesZone'].forEach(function(id){
  var zone=document.getElementById(id);
  zone.addEventListener('dragover',function(e){e.preventDefault();zone.classList.add('drag-over');});
  zone.addEventListener('dragleave',function(){zone.classList.remove('drag-over');});
  zone.addEventListener('drop',function(e){
    e.preventDefault();zone.classList.remove('drag-over');
    var fi=zone.querySelector('input[type=file]');
    var lb=zone.querySelector('.file-chosen');
    if(e.dataTransfer.files[0]){
      var dt=new DataTransfer();dt.items.add(e.dataTransfer.files[0]);
      fi.files=dt.files;lb.textContent='✓ '+e.dataTransfer.files[0].name;checkReady();
    }
  });
});
function setStep(n){
  for(var i=1;i<=5;i++){
    var el=document.getElementById('step'+i);
    el.classList.remove('active','done');
    if(i<n)el.classList.add('done');
    if(i===n)el.classList.add('active');
  }
  document.getElementById('progressFill').style.width=((n-1)/4*100)+'%';
}
document.getElementById('uploadForm').addEventListener('submit',async function(e){
  e.preventDefault();
  var btn=document.getElementById('uploadBtn');
  var errorMsg=document.getElementById('errorMsg');
  btn.disabled=true;errorMsg.style.display='none';
  document.getElementById('progress').style.display='block';
  setStep(1);
  var fd=new FormData();
  fd.append('warehouse',document.getElementById('warehouseFile').files[0]);
  fd.append('notes',document.getElementById('notesFile').files[0]);
  [1000,4000,8000,11000].forEach(function(delay,i){setTimeout(function(){setStep(i+2);},delay);});
  try{
    var res=await fetch('/upload',{method:'POST',body:fd});
    var data=await res.json();
    if(data.success){setStep(5);document.getElementById('progressFill').style.width='100%';setTimeout(function(){window.location.href='/dashboard';},800);}
    else{throw new Error(data.error||'Upload failed');}
  }catch(err){
    document.getElementById('progress').style.display='none';
    errorMsg.textContent='⚠ '+err.message;errorMsg.style.display='block';btn.disabled=false;
  }
});
</script></body></html>"""

DASHBOARD_HTML = open(Path(__file__).parent / 'templates' / 'dashboard.html').read() if (Path(__file__).parent / 'templates' / 'dashboard.html').exists() else None

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    has_data = PROCESSED_PATH.exists()
    meta = json.loads(META_PATH.read_text()) if META_PATH.exists() else {}
    html = INDEX_HTML
    if has_data:
        status = f'''<div class="status-banner">
          <strong>✓ Data loaded — {meta.get("uploaded","")}</strong>
          {meta.get("warehouse_filename","")} · {meta.get("rows",0):,} rows
          {(" · " + str(meta.get("sdrop_items","")) + " S-Drop items") if meta.get("sdrop_items") else ""}
        </div>'''
        link = '<a class="view-link" href="/dashboard">→ View current dashboard</a>'
    else:
        status = ''
        link = ''
    html = html.replace('__STATUS__', status).replace('__VIEWLINK__', link)
    return Response(html, mimetype='text/html')

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
            'notes_filename': notes.filename,
            'rows': data['kpis']['total_containers'],
            'sdrop_items': data['sdrop']['kpis']['total_items'],
            'uploaded': pd.Timestamp.now().strftime('%B %d, %Y'),
        }))
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/dashboard')
def dashboard():
    if not PROCESSED_PATH.exists():
        return redirect(url_for('index'))
    # Try template file first, fall back to inline
    tmpl = Path(__file__).parent / 'templates' / 'dashboard.html'
    if tmpl.exists():
        return Response(tmpl.read_text(), mimetype='text/html')
    return Response('<h1>Dashboard template missing</h1><p><a href="/">Go back</a></p>', mimetype='text/html')

@app.route('/api/data')
def api_data():
    if not PROCESSED_PATH.exists():
        return jsonify({'error': 'No data loaded'}), 404
    return PROCESSED_PATH.read_text(), 200, {'Content-Type': 'application/json'}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

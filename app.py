"""
eBarimt VAT Checker - Web App for Render.com
Browser calls the API directly (Mongolian IP required).
Render just serves the HTML page.
Run locally:  python app.py
"""

import os
import sys
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl")
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment

from io import BytesIO
from datetime import datetime
import json

PORT = int(os.environ.get("PORT", 5000))

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>eBarimt VAT Checker</title>
<style>
:root{
  --blue:#1a56db;--blue-dark:#1342b0;
  --green-bg:#f0fdf4;--green:#15803d;--green-bd:#bbf7d0;
  --red-bg:#fef2f2;--red:#b91c1c;--red-bd:#fecaca;
  --yellow-bg:#fefce8;--yellow:#a16207;--yellow-bd:#fde68a;
  --gray-50:#f9fafb;--gray-100:#f3f4f6;--gray-200:#e5e7eb;
  --gray-400:#9ca3af;--gray-600:#4b5563;--gray-800:#1f2937;
  --radius:8px;--radius-lg:12px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#f1f5f9;color:var(--gray-800);min-height:100vh}
.topbar{background:#fff;border-bottom:1px solid var(--gray-200);padding:0 32px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}
.topbar-brand{display:flex;align-items:center;gap:10px;font-size:15px;font-weight:600}
.topbar-sub{font-size:12px;color:var(--gray-400);font-weight:400}
.page{max-width:980px;margin:0 auto;padding:28px 24px}
.grid{display:grid;grid-template-columns:300px 1fr;gap:20px;align-items:start}
@media(max-width:720px){.grid{grid-template-columns:1fr}}
.card{background:#fff;border:1px solid var(--gray-200);border-radius:var(--radius-lg);padding:20px}
.card-title{font-size:11px;font-weight:700;color:var(--gray-400);text-transform:uppercase;letter-spacing:.08em;margin-bottom:14px}
.field-label{font-size:13px;color:var(--gray-600);margin-bottom:6px;display:block}
textarea{width:100%;height:200px;resize:vertical;font-family:monospace;font-size:12px;border:1px solid var(--gray-200);border-radius:var(--radius);padding:10px;color:var(--gray-800);outline:none;line-height:1.6}
textarea:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(26,86,219,.1)}
.or-div{text-align:center;font-size:12px;color:var(--gray-400);margin:10px 0}
.file-zone{border:1.5px dashed var(--gray-200);border-radius:var(--radius);padding:14px;text-align:center;cursor:pointer}
.file-zone:hover{border-color:var(--blue)}
.file-zone input{display:none}
.file-zone-text{font-size:13px;color:var(--gray-400)}
.file-zone-text b{color:var(--blue)}
#fileNameTag{font-size:12px;color:var(--green);margin-top:4px}
#countTag{font-size:12px;color:var(--gray-400);margin-top:8px}
.settings-row{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:14px}
.settings-row label{font-size:12px;color:var(--gray-600);display:flex;flex-direction:column;gap:4px}
select{padding:7px 8px;border:1px solid var(--gray-200);border-radius:var(--radius);font-size:13px;background:#fff;outline:none}
select:focus{border-color:var(--blue)}
.btn{width:100%;margin-top:12px;padding:10px;border-radius:var(--radius);border:none;font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;transition:background .15s}
.btn-start{background:var(--blue);color:#fff}
.btn-start:hover{background:var(--blue-dark)}
.btn-start:disabled{background:var(--gray-200);color:var(--gray-400);cursor:not-allowed}
.btn-stop{background:#fee2e2;color:var(--red);display:none}
.btn-stop:hover{background:#fecaca}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}
.stat{background:#fff;border:1px solid var(--gray-200);border-radius:var(--radius);padding:12px;text-align:center}
.stat-val{font-size:26px;font-weight:700}
.stat-lbl{font-size:11px;color:var(--gray-400);margin-top:2px}
.s-total .stat-val{color:var(--gray-800)}
.s-found .stat-val{color:var(--green)}
.s-notfound .stat-val{color:var(--red)}
.s-remain .stat-val{color:var(--blue)}
.progress-wrap{height:5px;background:var(--gray-100);border-radius:3px;overflow:hidden;margin-bottom:14px;display:none}
.progress-fill{height:100%;background:var(--blue);border-radius:3px;transition:width .2s}
.tabs{display:flex;gap:6px;margin-bottom:10px}
.tab{padding:5px 14px;border-radius:100px;border:1px solid var(--gray-200);font-size:12px;cursor:pointer;background:#fff;color:var(--gray-600);font-weight:500}
.tab.active{background:var(--blue);color:#fff;border-color:var(--blue)}
.table-wrap{max-height:400px;overflow-y:auto;border:1px solid var(--gray-200);border-radius:var(--radius-lg)}
table{width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed}
th{text-align:left;padding:9px 12px;font-size:11px;font-weight:600;color:var(--gray-400);text-transform:uppercase;letter-spacing:.04em;border-bottom:1px solid var(--gray-200);background:#fff;position:sticky;top:0}
td{padding:9px 12px;border-bottom:1px solid #f5f5f5;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
tr:last-child td{border-bottom:none}
tr:hover td{background:var(--gray-50)}
.mono{font-family:monospace;font-size:12px}
.badge{display:inline-flex;align-items:center;gap:3px;padding:2px 8px;border-radius:100px;font-size:11px;font-weight:600}
.b-found{background:var(--green-bg);color:var(--green);border:1px solid var(--green-bd)}
.b-nf{background:var(--red-bg);color:var(--red);border:1px solid var(--red-bd)}
.b-err{background:var(--yellow-bg);color:var(--yellow);border:1px solid var(--yellow-bd)}
.b-pend{background:var(--gray-100);color:var(--gray-400);border:1px solid var(--gray-200)}
.toolbar{display:flex;justify-content:space-between;align-items:center;margin-top:12px}
.toolbar-btns{display:flex;gap:8px}
.icon-btn{padding:7px 14px;border:1px solid var(--gray-200);border-radius:var(--radius);font-size:13px;font-weight:500;cursor:pointer;background:#fff;color:var(--gray-600);display:inline-flex;align-items:center;gap:6px}
.icon-btn:hover{background:var(--gray-50)}
.status-msg{font-size:12px;color:var(--gray-400)}
.empty{text-align:center;padding:48px 16px;color:var(--gray-400);font-size:13px}
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-brand">
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#1a56db" stroke-width="2"><path d="M9 14l2 2 4-4"/><rect x="3" y="4" width="18" height="16" rx="2"/></svg>
    eBarimt VAT Checker
    <span class="topbar-sub">APU Trading</span>
  </div>
  <div style="font-size:12px;color:var(--gray-400)" id="clock"></div>
</div>

<div class="page">
  <div class="grid">
    <div>
      <div class="card">
        <div class="card-title">Input</div>
        <span class="field-label">Paste registration numbers</span>
        <textarea id="regInput" placeholder="5113377&#10;УД95011056&#10;ОГ78011707&#10;..."></textarea>
        <div class="or-div">or upload a file</div>
        <label class="file-zone">
          <input type="file" id="fileInput" accept=".txt,.csv">
          <div class="file-zone-text">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display:block;margin:0 auto 4px"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            <b>Choose file</b> or drag &amp; drop<br>.txt or .csv
          </div>
          <div id="fileNameTag"></div>
        </label>
        <div id="countTag"></div>
        <div class="settings-row">
          <label>Parallel requests
            <select id="concurrency"><option value="1">1</option><option value="3">3</option><option value="5" selected>5</option><option value="10">10</option></select>
          </label>
          <label>Delay between calls
            <select id="delay"><option value="0">None</option><option value="100">100 ms</option><option value="200" selected>200 ms</option><option value="500">500 ms</option></select>
          </label>
        </div>
        <button class="btn btn-start" id="startBtn" onclick="startJob()">▶ Start verification</button>
        <button class="btn btn-stop" id="stopBtn" onclick="stopJob()">■ Stop</button>
      </div>
    </div>

    <div>
      <div class="stats">
        <div class="stat s-total"><div class="stat-val" id="sTot">0</div><div class="stat-lbl">Total</div></div>
        <div class="stat s-found"><div class="stat-val" id="sOk">0</div><div class="stat-lbl">Found</div></div>
        <div class="stat s-notfound"><div class="stat-val" id="sErr">0</div><div class="stat-lbl">Not found</div></div>
        <div class="stat s-remain"><div class="stat-val" id="sPend">0</div><div class="stat-lbl">Remaining</div></div>
      </div>
      <div class="progress-wrap" id="progressWrap"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
      <div class="tabs">
        <button class="tab active" onclick="setFilter('all',this)">All</button>
        <button class="tab" onclick="setFilter('Found',this)">Found</button>
        <button class="tab" onclick="setFilter('Not found',this)">Not found</button>
        <button class="tab" onclick="setFilter('Error',this)">Errors</button>
      </div>
      <div class="table-wrap">
        <table>
          <colgroup><col style="width:36px"><col style="width:130px"><col style="width:auto"><col style="width:100px"></colgroup>
          <thead><tr><th>#</th><th>Reg No</th><th>TIN</th><th>Status</th></tr></thead>
          <tbody id="tbody"><tr><td colspan="4"><div class="empty">Paste numbers and click Start</div></td></tr></tbody>
        </table>
      </div>
      <div class="toolbar">
        <span class="status-msg" id="statusMsg"></span>
        <div class="toolbar-btns">
          <button class="icon-btn" onclick="exportCSV()">⬇ Export CSV</button>
          <button class="icon-btn" onclick="clearAll()">🗑 Clear</button>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
const API = 'https://api.ebarimt.mn/api/info/check/getTinInfo?regNo=';
let allResults = [], abort = false, currentFilter = 'all';
let tot=0, ok=0, err=0, pend=0;

setInterval(()=>{
  const n=new Date();
  document.getElementById('clock').textContent=n.toLocaleDateString('mn-MN')+' '+n.toLocaleTimeString('mn-MN');
},1000);

document.getElementById('fileInput').addEventListener('change',e=>{
  const f=e.target.files[0];if(!f)return;
  document.getElementById('fileNameTag').textContent=f.name;
  const r=new FileReader();
  r.onload=ev=>{document.getElementById('regInput').value=ev.target.result;updateCount();};
  r.readAsText(f);
});
document.getElementById('regInput').addEventListener('input',updateCount);

function parseNums(){
  return[...new Set(
    document.getElementById('regInput').value
      .split(/[\n,;]+/).map(s=>s.trim()).filter(Boolean)
  )];
}
function updateCount(){
  const n=parseNums().length;
  document.getElementById('countTag').textContent=n?n+' numbers detected':'';
}

async function checkOne(regNo){
  const url=API+encodeURIComponent(regNo);
  const res=await fetch(url);
  if(!res.ok) throw new Error('HTTP '+res.status);
  const data=await res.json();
  if(data&&data.status===200&&data.data!==undefined&&data.data!==null&&data.data!==''){
    return{tin:String(data.data),status:'Found'};
  }
  return{tin:'',status:'Not found'};
}

async function startJob(){
  const nums=parseNums();
  if(!nums.length){alert('Enter at least one registration number.');return;}
  abort=false;
  tot=nums.length;ok=0;err=0;pend=nums.length;
  allResults=nums.map(r=>({reg:r,tin:'',status:'Pending'}));
  document.getElementById('startBtn').style.display='none';
  document.getElementById('stopBtn').style.display='';
  document.getElementById('progressWrap').style.display='block';
  document.getElementById('statusMsg').textContent='Running...';
  renderTable();updateStats();

  const conc=parseInt(document.getElementById('concurrency').value);
  const delay=parseInt(document.getElementById('delay').value);
  let idx=0;
  const lock={i:0};

  async function worker(){
    while(true){
      if(abort)break;
      const i=lock.i++;
      if(i>=nums.length)break;
      try{
        const r=await checkOne(nums[i]);
        allResults[i]={reg:nums[i],tin:r.tin,status:r.status};
        if(r.status==='Found')ok++;else err++;
      }catch(e){
        allResults[i]={reg:nums[i],tin:'',status:'Error'};
        err++;
      }
      pend--;
      updateStats();
      renderRow(i);
      if(delay)await sleep(delay);
    }
  }

  await Promise.all(Array.from({length:conc},()=>worker()));
  document.getElementById('startBtn').style.display='';
  document.getElementById('stopBtn').style.display='none';
  document.getElementById('statusMsg').textContent=
    'Done — '+ok+' found, '+err+' not found/error out of '+tot;
}

function stopJob(){
  abort=true;
  allResults.filter(r=>r.status==='Pending').forEach(r=>{r.status='Stopped';pend--;});
  renderTable();updateStats();
  document.getElementById('startBtn').style.display='';
  document.getElementById('stopBtn').style.display='none';
}

function sleep(ms){return new Promise(r=>setTimeout(r,ms));}

function updateStats(){
  document.getElementById('sTot').textContent=tot;
  document.getElementById('sOk').textContent=ok;
  document.getElementById('sErr').textContent=err;
  document.getElementById('sPend').textContent=pend;
  document.getElementById('progressFill').style.width=tot?Math.round((tot-pend)/tot*100)+'%':'0%';
}

function setFilter(f,el){
  currentFilter=f;
  document.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));
  el.classList.add('active');renderTable();
}

function badge(s){
  if(s==='Found')return'<span class="badge b-found">✓ Found</span>';
  if(s==='Not found')return'<span class="badge b-nf">✗ Not found</span>';
  if(s==='Error')return'<span class="badge b-err">⚠ Error</span>';
  if(s==='Stopped')return'<span class="badge b-pend">Stopped</span>';
  return'<span class="badge b-pend">Pending</span>';
}

function getFiltered(){
  if(currentFilter==='all')return allResults;
  return allResults.filter(r=>r.status===currentFilter);
}

function renderTable(){
  const rows=getFiltered();
  if(!rows.length){
    document.getElementById('tbody').innerHTML='<tr><td colspan="4"><div class="empty">No results match this filter</div></td></tr>';
    return;
  }
  document.getElementById('tbody').innerHTML=rows.map((r,i)=>`
    <tr>
      <td style="color:#9ca3af">${i+1}</td>
      <td class="mono">${r.reg}</td>
      <td class="mono">${r.tin||'—'}</td>
      <td>${badge(r.status)}</td>
    </tr>`).join('');
}

function renderRow(i){
  if(currentFilter!=='all'){renderTable();return;}
  const trs=document.getElementById('tbody').querySelectorAll('tr');
  if(trs[i])trs[i].innerHTML=`
    <td style="color:#9ca3af">${i+1}</td>
    <td class="mono">${allResults[i].reg}</td>
    <td class="mono">${allResults[i].tin||'—'}</td>
    <td>${badge(allResults[i].status)}</td>`;
  else renderTable();
}

function exportCSV(){
  if(!allResults.length){alert('No results to export.');return;}
  const csv=['Reg No,TIN,Status',...allResults.map(r=>[r.reg,r.tin,r.status].join(','))].join('\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));
  a.download='vat_results_'+new Date().toISOString().slice(0,10)+'.csv';
  a.click();
}

function clearAll(){
  allResults=[];tot=0;ok=0;err=0;pend=0;
  document.getElementById('regInput').value='';
  document.getElementById('countTag').textContent='';
  document.getElementById('fileNameTag').textContent='';
  document.getElementById('progressFill').style.width='0%';
  document.getElementById('progressWrap').style.display='none';
  document.getElementById('statusMsg').textContent='';
  ['sTot','sOk','sErr','sPend'].forEach(id=>document.getElementById(id).textContent='0');
  document.getElementById('tbody').innerHTML='<tr><td colspan="4"><div class="empty">Paste numbers and click Start</div></td></tr>';
}
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ('/', '/index.html'):
            body = HTML.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', len(body))
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        self.send_response(404)
        self.end_headers()


def main():
    server = HTTPServer(('0.0.0.0', PORT), Handler)
    print(f'\n  eBarimt VAT Checker running at http://localhost:{PORT}')
    print('  Press Ctrl+C to stop\n')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n  Stopped.')


if __name__ == '__main__':
    main()

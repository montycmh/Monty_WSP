function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

const arFiles = [null, null]; // [file0, file1]

function arDragOver(e, id){ e.preventDefault(); document.getElementById(id).style.borderColor='#1D9E75'; }
function arDragLeave(id){ document.getElementById(id).style.borderColor='#D3D1C7'; }
function arDrop(e, idx){
e.preventDefault();
  const id = idx===0?'ar-drop-0':'ar-drop-1';
document.getElementById(id).style.borderColor='#D3D1C7';
  const file = e.dataTransfer.files[0];
  if(file) arSetFile(idx, file);
}
function arFileSelected(idx, input){
  if(input.files[0]) arSetFile(idx, input.files[0]);
}
function arSetFile(idx, file){
  arFiles[idx] = file;
  const icon  = document.getElementById('ar-icon-'+idx);
  const label = document.getElementById('ar-label-'+idx);
  const drop  = document.getElementById(idx===0?'ar-drop-0':'ar-drop-1');
icon.className  = 'ti ti-file-check';
icon.style.color = '#1D9E75';
label.textContent = file.name;
label.style.color = '#0F6E56';
drop.style.borderColor = '#1D9E75';
drop.style.borderStyle = 'solid';
  // Enable button if both files loaded
  if(arFiles[0] && arFiles[1]){
    const btn = document.getElementById('ar-generate-btn');
btn.disabled = false;
btn.style.background = '#1E2761';
btn.style.color = '#fff';
btn.style.cursor = 'pointer';
  }
}

async function arGenerateReport(){
  const btn = document.getElementById('ar-generate-btn');
btn.textContent = 'Processing...';
btn.disabled = true;

  try {
    // Load SheetJS dynamically if not present
    if(typeof XLSX === 'undefined'){
      await new Promise((res,rej)=>{
        const s = document.createElement('script');
s.src = ' https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
s.onload = res; s.onerror = rej;
document.head.appendChild(s);
      });
    }

    const [wb0, wb1] = await Promise.all([arReadXLSX(arFiles[0]), arReadXLSX(arFiles[1])]);
    const d0 = arParseRows(wb0, arFiles[0].name);
    const d1 = arParseRows(wb1, arFiles[1].name);

    // Determine recent vs old by date in filename
    const recent = d0.fileDate >= d1.fileDate ? d0 : d1;
    const old    = d0.fileDate >= d1.fileDate ? d1 : d0;

    // Apply offset to all customer rows (not totals)
    arApplyOffset(recent.customers);
    arApplyOffset(old.customers);

    // Grand totals (raw, no offset)
    const gt = recent.grandTotal;
    const gtOld = old.grandTotal;

    // Top 15 by total desc
    const top15 = [...recent.customers].sort((a,b)=>b.total-a.total).slice(0,15);

    // Lookup old values for each top15 customer
    const oldMap = {};
old.customers.forEach(r => oldMap[r.customer] = r);

    // Detect flags (using raw pre-offset values)
    const flags = arDetectFlags(top15, recent.rawMap, old.rawMap);

    // Derive labels
    const rd = recent.fileDate;
    const blueDate  = new Date(rd.slice(0,4), parseInt(rd.slice(4,6))-2, 1);
    const greenDate = new Date(rd.slice(0,4), parseInt(rd.slice(4,6))-3, 1);
    const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const blueLabel  = MONTHS[blueDate.getMonth()]  + ' Close';
    const greenLabel = MONTHS[greenDate.getMonth()] + ' Close';
    const asOfLabel  = ${rd.slice(4,6)}/${rd.slice(6,8)}/${rd.slice(0,4)};
    const outName    = AR_Aging_Report_${MONTHS[blueDate.getMonth()]}${blueDate.getFullYear()}.html;
    const blueYear   = blueDate.getFullYear();

    const html = arBuildHTML({
      top15, oldMap, gt, gtOld, blueLabel, greenLabel, asOfLabel, flags, outName, blueYear,
      recentRawDate: recent.rawDate, oldRawDate: old.rawDate
    });
    const win = window.open('', '_blank');
win.document.write(html);
win.document.close();

  } catch(err){
    alert('Error generating report: ' + err.message);
console.error(err);
  }

  btn.innerHTML = '<i class="ti ti-report-analytics" style="font-size:16px;"></i>Generate Report';
btn.disabled = false;
btn.style.background = '#1E2761';
btn.style.color = '#fff';
}

async function arReadXLSX(file){
  return new Promise((res,rej)=>{
    const reader = new FileReader();
reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array'});
        res(wb);
      } catch(err){ rej(err); }
    };
reader.onerror = rej;
reader.readAsArrayBuffer(file);
  });
}

// ---------------------------------------------------------------------------
// The four functions below (arParseRows, arApplyOffset, arDetectFlags,
// arBuildHTML) were MISSING from the uploaded app.js — arGenerateReport()
// called them but they were never defined anywhere in the file, so every
// "Generate Report" click threw a ReferenceError as soon as it tried to run
// them (caught by the try/catch above, which is why nothing visibly happened
// beyond an alert box). They have been reconstructed below based on the
// business rules documented in the companion ar-aging-report skill reference
// (offset logic, Top 15 selection, flag detection, report layout). Please
// spot-check the first generated report against a known-good prior version
// before relying on it for close, since the exact original implementation
// wasn't available to compare against.
// ---------------------------------------------------------------------------

function arToNum(v){
  if(typeof v === 'number') return v;
  if(v===null || v===undefined || v==='') return 0;
  let s = String(v).trim();
  if(s==='') return 0;
  const neg = /^(.*)$/.test(s);
  s = s.replace(/[(),$\s]/g,'').replace(/,/g,'');
  const n = parseFloat(s);
  if(isNaN(n)) return 0;
  return neg ? -Math.abs(n) : n;
}

function arParseRows(wb, filename){
  const m = filename.match(/(\d{2})(\d{2})(\d{4}).xlsx?$/i);
  if(!m){
    throw new Error('Cannot parse date from filename "' + filename + '" — expected pattern JAZAR-ACO-MMDDYYYY.xlsx');
  }
  const mm = m[1], dd = m[2], yyyy = m[3];
  const fileDate = yyyy + mm + dd; // YYYYMMDD, sortable
  const rawDate = mm + dd + yyyy; // MMDDYYYY, as it appears in the filename

  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});

  // The known NetSuite export layout has headers on row 7 (index 6) and data
  // starting row 9 (index 8). We try to locate the header row by looking for
  // a "Customer" label first, and fall back to the fixed index if not found,
  // so small formatting differences between exports don't silently break this.
  let headerIdx = rows.findIndex(r => r && String(r[0]||'').toLowerCase().includes('customer'));
  let startIdx = headerIdx >= 0 ? headerIdx + 1 : 8;

  const customers = [];
  const rawMap = {};
  let grandTotal = null;

  for(let i=startIdx; i<rows.length; i++){
    const r = rows[i];
    if(!r || r.length===0) continue;
    const name = String(r[0]||'').trim();
    if(!name) continue;

    const cur   = arToNum(r[1]);
    const b30   = arToNum(r[2]);
    const b60   = arToNum(r[3]);
    const b90   = arToNum(r[4]);
    const b90p  = arToNum(r[5]);
    const total = arToNum(r[6]);

    if(name.toLowerCase() === 'total'){
      grandTotal = {cur, b30, b60, b90, b90p, total};
      break;
    }

    customers.push({customer:name, cur, b30, b60, b90, b90p, total});
    rawMap[name] = {cur, b30, b60, b90, b90p, total};
  }

  if(!grandTotal){
    throw new Error('Could not find a "Total" row in ' + filename + ' — check the file format/layout.');
  }
  if(customers.length === 0){
    throw new Error('No customer rows found in ' + filename + '.');
  }

  return {fileDate, rawDate, customers, grandTotal, rawMap};
}

function arApplyOffset(customers){
  // Reclassifies negative bucket values against positive buckets, oldest
  // bucket first. Total Balance (row.total) is left unchanged.
customers.forEach(row => {
    const vals = [row.b90p, row.b90, row.b60, row.b30, row.cur]; // oldest -> newest
    for(let i=0; i<vals.length; i++){
      if(vals[i] < 0){
        let credit = -vals[i];
        vals[i] = 0;
        for(let j=0; j<vals.length; j++){
          if(j!==i && vals[j] > 0 && credit > 0){
            const consumed = Math.min(vals[j], credit);
            vals[j] -= consumed;
            credit -= consumed;
          }
        }
        if(credit > 0){
          vals[4] -= credit; // remainder stays in "Not due yet" as negative
        }
      }
    }
row.b90p = vals[0]; row.b90 = vals[1]; row.b60 = vals[2]; row.b30 = vals[3]; row.cur = vals[4];
  });
}

function arDetectFlags(top15, recentRawMap, oldRawMap){
  const bucketsMap = [['b30','0 – 30'], ['b60','31 – 60'], ['b90','61 – 90'], ['b90p','90+']];
  const flags = {};

  top15.forEach(row => {
    const name = row.customer;
    const rRaw = recentRawMap[name] || {};
    const oRaw = oldRawMap[name] || {};
    const customerFlags = [];

    bucketsMap.forEach(([key, label]) => {
      const rVal = rRaw[key] || 0;
      const oVal = oRaw[key] || 0;
      if(rVal < 0 || oVal < 0){
        let note;
        if(oVal < 0 && rVal >= 0){
          note = 'Bucket was negative in prior close (' + Math.abs(oVal).toLocaleString('en-US') + '). Value changed sign in current close. Verify if credit was correctly applied or rebucketed.';
        } else if(rVal < 0 && oVal === 0){
          note = 'New negative appeared in current close (' + Math.abs(rVal).toLocaleString('en-US') + ') not present in prior close. Verify source of credit.';
        } else {
          note = 'Negative present in both closes. Current: ' + rVal.toLocaleString('en-US') + ' / Prior: ' + oVal.toLocaleString('en-US') + '. Verify if amounts are consistent.';
        }
customerFlags.push({bucket:label, recentVal:rVal, oldVal:oVal, note});
      }
    });

    if(customerFlags.length) flags[name] = customerFlags;
  });

  return flags;
}

function arBuildHTML(ctx){
  const RAW = ctx.top15.map(r => {
    const o = ctx.oldMap[r.customer] || {b90:0, b90p:0};
    const rowFlags = ctx.flags[r.customer] || [];
    return {
      name: r.customer,
      cur: r.cur, b30: r.b30, b60: r.b60, b90: r.b90, b90p: r.b90p, tot: r.total,
      ab90: o.b90 || 0, ab90p: o.b90p || 0,
      flag: rowFlags.length > 0,
      flag_detail: rowFlags.map(f => ({bucket:f.bucket, recent_val:f.recentVal, old_val:f.oldVal, note:f.note}))
    };
  });

  const RECENT_TOTAL = {cur:ctx.gt.cur, b30:ctx.gt.b30, b60:ctx.gt.b60, b90:ctx.gt.b90, b90p:ctx.gt.b90p, tot:ctx.gt.total};
  const OLD_TOTAL = {b90:ctx.gtOld.b90, b90p:ctx.gtOld.b90p, tot:ctx.gtOld.total};

  const escJs = s => JSON.stringify(s).replace(/</g, '\u003c');
  const recentMonthWord = ctx.blueLabel.replace(' Close','');
  const oldMonthWord = ctx.greenLabel.replace(' Close','');
  const n = RAW.length;
  const reportTitle = A/R Aging Report — ${ctx.blueLabel} ${ctx.blueYear};

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${esc(reportTitle)}</title>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
    font-size: 12px;
    background: #f5f6f8;
    color: #1a1a2e;
    padding: 24px;
    min-height: 100vh;
  }

  .page {
    max-width: 960px;
    margin: 0 auto;
    background: #fff;
    border-radius: 10px;
    box-shadow: 0 2px 16px rgba(0,0,0,0.08);
    padding: 16px 20px 24px;
  }

  /* Header */
  .report-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 20px;
    padding-bottom: 16px;
    border-bottom: 1.5px solid #e2e6ea;
  }
  .report-header h1 {
    font-size: 17px;
    font-weight: 600;
    color: #0c2340;
    letter-spacing: -0.01em;
  }
  .report-header .subtitle {
    font-size: 11px;
    color: #6b7a8d;
    margin-top: 3px;
  }
  .report-header .meta-pills {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    justify-content: flex-end;
  }
  .pill {
    font-size: 10.5px;
    font-weight: 500;
    padding: 3px 10px;
    border-radius: 20px;
    white-space: nowrap;
  }
  .pill-blue  { background: #e6f1fb; color: #0c447c; }
  .pill-green { background: #eaf3de; color: #3b6d11; }
  .pill-gray  { background: #f1efe8; color: #444441; }

  /* Table wrapper */
  .table-wrap { overflow-x: visible; width: 100%; }

  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 10px;
    min-width: 0;
    table-layout: fixed;
  }

  colgroup col:first-child { width: 120px; }

  /* Header rows */
  thead tr:first-child th { border-bottom: none; }
  thead tr:last-child th  { border-top: none; }

  th {
    padding: 4px 5px;
    text-align: right;
    border: 0.5px solid #d0d7df;
    white-space: normal;
    overflow: hidden;
    font-weight: 500;
    line-height: 1.2;
    word-break: break-word;
  }
th.left { text-align: left; }

  th.grp-may  { background: #0c447c; color: #b5d4f4; border-color: #185fa5; text-align: center; }
th.grp-apr  { background: #3b6d11; color: #c0dd97; border-color: #639922; text-align: center; }
th.grp-var  { background: #854f0b; color: #fac775; border-color: #ba7517; text-align: center; }
th.sub      { font-size: 10px; font-weight: 400; }

  /* Body cells */
  td {
    padding: 4px 5px;
    text-align: right;
    border: 0.5px solid #d0d7df;
    color: #1a1a2e;
    white-space: nowrap;
    overflow: visible;
    line-height: 1.3;
  }
td.name {
    text-align: left;
    font-size: 9.5px;
    color: #1a1a2e;
    white-space: normal;
    word-break: break-word;
    line-height: 1.3;
    overflow: visible;
  }

  tr.total-row td {
    background: #f0f4f8;
    font-weight: 600;
    border-top: 1.5px solid #b0bcc8;
  }
tr.pct-row td {
    color: #6b7a8d;
    font-size: 9px;
    background: #fafbfc;
  }
tr.cust:hover td { background: #f7f9fb; }
tr.spacer td { height: 8px; border: none; background: #f5f6f8; }

  .div-col { border-left: 2px solid #b0bcc8 !important; }

  /* Variance colors */
  .pos-var { color: #a32d2d; font-weight: 600; }
  .neg-var { color: #27670a; font-weight: 600; }
  .zero    { color: #b0bcc8; }

  /* Editable cells */
td.editable { background: #f0f7ff; cursor: text; }
td.editable:focus-within { background: #e0efff; outline: 1.5px solid #378add; border-radius: 2px; }
input.cell-edit {
    width: 100%;
    text-align: right;
    background: transparent;
    border: none;
    color: #1a1a2e;
    font-size: 10px;
    font-family: inherit;
    padding: 0;
    cursor: text;
    line-height: 1.3;
  }
input.cell-edit:focus { outline: none; }

  /* Flagged cells */
td.flagged, td.flagged input.cell-edit { background: #fff7e6 !important; color: #7a3e00; }
tr.cust:hover td.flagged { background: #fff0d0 !important; }
  .flag-icon { color: #ba7517; font-size: 11px; margin-left: 4px; vertical-align: middle; }

  /* Recalc note */
  .recalc-note {
    display: none;
    font-size: 10.5px;
    color: #378add;
    margin-top: 6px;
    font-style: italic;
  }

  /* Flags section */
  .flags-section {
    margin-top: 24px;
    border: 1px solid #f0c880;
    border-radius: 8px;
    overflow: hidden;
  }
  .flags-header {
    background: #854f0b;
    color: #fac775;
    font-size: 11.5px;
    font-weight: 600;
    padding: 8px 14px;
    display: flex;
    align-items: center;
    gap: 7px;
    letter-spacing: 0.01em;
  }
  .flags-header svg { flex-shrink: 0; }
  .flag-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 11px;
  }
  .flag-table th {
    background: #fdf3e0;
    color: #7a3e00;
    font-weight: 600;
    text-align: left;
    padding: 6px 12px;
    border-bottom: 1px solid #f0d8a0;
    border-right: 0.5px solid #f0d8a0;
  }
  .flag-table td {
    padding: 7px 12px;
    text-align: left;
    border-bottom: 0.5px solid #f0e8c8;
    border-right: 0.5px solid #f0e8c8;
    color: #3a2800;
    white-space: normal;
  }
  .flag-table tr:last-child td { border-bottom: none; }
  .flag-table .val-neg { color: #a32d2d; font-weight: 600; }
  .flag-table .val-pos { color: #1a1a2e; }

  /* Freeze button */
  .btn-freeze {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 11.5px;
    font-weight: 600;
    padding: 5px 14px;
    border-radius: 6px;
    border: 1.5px solid #0c447c;
    background: #fff;
    color: #0c447c;
    cursor: pointer;
    transition: background 0.15s, color 0.15s;
    white-space: nowrap;
  }
  .btn-freeze:hover { background: #e6f1fb; }
  .btn-freeze.locked {
    border-color: #3b6d11;
    color: #3b6d11;
  }
  .btn-freeze.locked:hover { background: #eaf3de; }

  /* Frozen state overrides */
body.frozen td.editable          { background: inherit !important; cursor: default; }
body.frozen td.editable:focus-within { outline: none; background: inherit !important; }
body.frozen input.cell-edit      { pointer-events: none; cursor: default; color: #1a1a2e; }
body.frozen td.flagged,
body.frozen td.flagged input.cell-edit { background: inherit !important; color: #1a1a2e !important; }
body.frozen tr.cust:hover td.flagged   { background: #f7f9fb !important; }
body.frozen .flag-icon           { display: none; }
body.frozen .flags-section       { display: none; }
body.frozen .recalc-note         { display: none !important; }
body.frozen td.name.flagged      { color: #1a1a2e; }

  /* Copy image button */
  .btn-copy {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 11.5px;
    font-weight: 600;
    padding: 5px 14px;
    border-radius: 6px;
    border: 1.5px solid #6b7a8d;
    background: #fff;
    color: #6b7a8d;
    cursor: pointer;
    transition: background 0.15s, color 0.15s, border-color 0.15s;
    white-space: nowrap;
  }
  .btn-copy:hover { background: #f0f4f8; border-color: #1a1a2e; color: #1a1a2e; }
  .btn-copy.success { border-color: #3b6d11; color: #3b6d11; background: #eaf3de; }
  .btn-copy.error   { border-color: #a32d2d; color: #a32d2d; background: #fdf0f0; }
  .btn-row { display: flex; gap: 8px; align-items: center; }

  /* Print */
  @media print {
    body { background: #fff; padding: 0; }
    .page { box-shadow: none; border-radius: 0; padding: 16px; }
input.cell-edit { -webkit-appearance: none; }
  }
</style>
<script src=" https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
</head>
<body>
<div class="page">

  <div class="report-header">
    <div>
      <h1>A/R Aging Report — Komodo Health</h1>
      <div class="subtitle">Generated from NetSuite export · Rules-based calculation</div>
    </div>
    <div style="display:flex;flex-direction:column;align-items:flex-end;gap:10px">
      <div class="meta-pills">
        <span class="pill pill-gray">Days overdue as of <strong>${esc(ctx.asOfLabel)}</strong></span>
        <span class="pill pill-blue">Blue → ${esc(ctx.blueLabel)} (${esc(ctx.recentRawDate)})</span>
        <span class="pill pill-green">Green → ${esc(ctx.greenLabel)} (${esc(ctx.oldRawDate)})</span>
      </div>
      <div class="btn-row">
        <button class="btn-copy" id="btn-copy" onclick="copyTableAsImage()">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
          <span id="copy-label">Copy as image</span>
        </button>
        <button class="btn-freeze" id="btn-freeze" onclick="toggleFreeze()">
          <svg id="freeze-icon" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
          <span id="freeze-label">Lock report</span>
        </button>
      </div>
    </div>
  </div>

  <div class="table-wrap">
    <table>
      <colgroup>
        <col style="width:108px">
        <col style="width:70px"><col style="width:62px"><col style="width:62px"><col style="width:62px"><col style="width:62px"><col style="width:70px"><col style="width:74px">
        <col style="width:62px"><col style="width:62px"><col style="width:70px">
        <col style="width:74px"><col style="width:74px">
      </colgroup>
      <thead>
        <tr>
          <th class="left grp-may" rowspan="2" style="font-size:9px">Days overdue<br><span style="font-weight:400;font-size:8.5px">as of ${esc(ctx.asOfLabel)}</span></th>
          <th class="grp-may" colspan="7" style="font-size:9.5px">A/R Aging Report (${esc(ctx.blueLabel)})</th>
          <th class="grp-apr div-col" colspan="3" style="font-size:9.5px">A/R Aging Report (${esc(ctx.greenLabel)})</th>
          <th class="grp-var div-col" colspan="2" style="font-size:9.5px">Variance</th>
        </tr>
        <tr>
          <th class="grp-may sub" style="font-size:8.5px">Not due yet</th>
          <th class="grp-may sub" style="font-size:8.5px">0 – 30</th>
          <th class="grp-may sub" style="font-size:8.5px">31 – 60</th>
          <th class="grp-may sub" style="font-size:8.5px">61 – 90</th>
          <th class="grp-may sub" style="font-size:8.5px">90+</th>
          <th class="grp-may sub" style="font-size:8.5px">60+ days total</th>
          <th class="grp-may sub" style="font-size:8.5px">Total balance</th>
          <th class="grp-apr sub div-col" style="font-size:8.5px">61 – 90</th>
          <th class="grp-apr sub" style="font-size:8.5px">90+</th>
          <th class="grp-apr sub" style="font-size:8.5px">60+ days total</th>
          <th class="grp-var sub div-col" style="font-size:8.5px">90+ vs prior close</th>
          <th class="grp-var sub" style="font-size:8.5px">60+ vs prior close</th>
        </tr>
      </thead>
      <tbody id="tb"></tbody>
    </table>
  </div>

  <div class="recalc-note" id="recalc-note">* Subtotals and variances recalculated after manual edit.</div>

  <div class="flags-section" id="flags-section">
    <div class="flags-header">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
      Negative bucket inconsistencies — Top ${n} (requires investigation)
    </div>
    <table class="flag-table">
      <thead>
        <tr>
          <th style="width:200px">Customer</th>
          <th style="width:90px">Bucket</th>
          <th style="width:130px">${esc(recentMonthWord)} value</th>
          <th style="width:130px">${esc(oldMonthWord)} value</th>
          <th>Note</th>
        </tr>
      </thead>
      <tbody id="flags-body"></tbody>
    </table>
  </div>

</div>

<script>
const fmt = n => {
  if (n === 0) return '-';
  const abs = Math.round(Math.abs(n)).toLocaleString('en-US');
  return n < 0 ? '(' + abs + ')' : abs;
};
const pct = (n, d) => d === 0 ? '-' : Math.round(n / d * 100) + '%';
const vcTd = (v, extraClass = '') => {
  if (v === 0) return '<td class="zero' + extraClass + '">-</td>';
  return '<td class="' + (v > 0 ? 'pos-var' : 'neg-var') + extraClass + '">' + fmt(v) + '</td>';
};

// === DATA (generated from the uploaded NetSuite exports) ===
// Grand totals: raw, no offset applied.
const RECENT_TOTAL = ${escJs(RECENT_TOTAL)};
RECENT_TOTAL['60p'] = RECENT_TOTAL.b90 + RECENT_TOTAL.b90p;

const OLD_TOTAL = ${escJs(OLD_TOTAL)};
OLD_TOTAL['60p'] = OLD_TOTAL.b90 + OLD_TOTAL.b90p;

// Top ${n} sorted by Total desc from recent file; prior-close values via lookup.
// Offset already applied per customer to the individual bucket values below.
const RAW = ${escJs(RAW)};

// Mutable state for editable cells
const state = {};
RAW.forEach(r => {
  state[r.name] = { cur: r.cur, b30: r.b30, b60: r.b60, b90: r.b90, b90p: r.b90p, tot: r.tot, ab90: r.ab90, ab90p: r.ab90p };
});

function calcTopN() {
  const t = { cur: 0, b30: 0, b60: 0, b90: 0, b90p: 0, tot: 0, ab90: 0, ab90p: 0 };
RAW.forEach(r => {
    const s = state[r.name];
t.cur += s.cur; t.b30 += s.b30; t.b60 += s.b60; t.b90 += s.b90; t.b90p += s.b90p;
t.tot += (s.cur + s.b30 + s.b60 + s.b90 + s.b90p); // recalculated from buckets so edits flow through
t.ab90 += s.ab90; t.ab90p += s.ab90p;
  });
  t['60p']  = t.b90 + t.b90p;
  t['a60p'] = t.ab90 + t.ab90p;
  return t;
}

function parseInput(val) {
  const isNeg = val.trim().startsWith('(') || val.trim().startsWith('-');
  const num = parseFloat(val.replace(/[(),\-\s,]/g, '').replace(/,/g, ''));
  if (isNaN(num)) return null;
  return isNeg ? -Math.abs(num) : num;
}

function editCell(name, field, val, extraClass = '') {
  const isFlagged = RAW.find(r => r.name === name)?.flag;
  const fc = isFlagged ? ' flagged' : '';
  return '<td class="editable' + extraClass + fc + '">' +
    '<input class="cell-edit" data-name="' + name + '" data-field="' + field + '"' +
    ' value="' + (val === 0 ? '' : fmt(val)) + '" placeholder="-" />' +
    '</td>';
}

function render() {
  const recent60p = RECENT_TOTAL['60p'];
  const old60p = OLD_TOTAL['60p'];
  let h = '';

  // --- Total Amount ($) row ---
  h += '<tr class="total-row">' +
    '<td class="name">Total amount ($)</td>' +
    '<td>' + fmt(RECENT_TOTAL.cur) + '</td>' +
    '<td>' + fmt(RECENT_TOTAL.b30) + '</td>' +
    '<td>' + fmt(RECENT_TOTAL.b60) + '</td>' +
    '<td>' + fmt(RECENT_TOTAL.b90) + '</td>' +
    '<td>' + fmt(RECENT_TOTAL.b90p) + '</td>' +
    '<td>' + fmt(recent60p) + '</td>' +
    '<td>' + fmt(RECENT_TOTAL.tot) + '</td>' +
    '<td class="div-col">' + fmt(OLD_TOTAL.b90) + '</td>' +
    '<td>' + fmt(OLD_TOTAL.b90p) + '</td>' +
    '<td>' + fmt(old60p) + '</td>' +
    vcTd(RECENT_TOTAL.b90p - OLD_TOTAL.b90p, ' div-col') +
    vcTd(recent60p - old60p) +
    '</tr>';

  // --- % of total (bucket / grand total) ---
  h += '<tr class="pct-row">' +
    '<td class="name">% of total</td>' +
    '<td>' + pct(RECENT_TOTAL.cur,  RECENT_TOTAL.tot) + '</td>' +
    '<td>' + pct(RECENT_TOTAL.b30,  RECENT_TOTAL.tot) + '</td>' +
    '<td>' + pct(RECENT_TOTAL.b60,  RECENT_TOTAL.tot) + '</td>' +
    '<td>' + pct(RECENT_TOTAL.b90,  RECENT_TOTAL.tot) + '</td>' +
    '<td>' + pct(RECENT_TOTAL.b90p, RECENT_TOTAL.tot) + '</td>' +
    '<td>' + pct(recent60p,         RECENT_TOTAL.tot) + '</td>' +
    '<td>100%</td>' +
    '<td class="div-col">' + pct(OLD_TOTAL.b90,  OLD_TOTAL.tot) + '</td>' +
    '<td>' + pct(OLD_TOTAL.b90p, OLD_TOTAL.tot) + '</td>' +
    '<td>' + pct(old60p,         OLD_TOTAL.tot) + '</td>' +
    '<td class="div-col zero">-</td>' +
    '<td class="zero">-</td>' +
    '</tr>';

  h += '<tr class="spacer"><td colspan="12"></td></tr>';

  // --- Top N subtotal ---
  const t = calcTopN();
  h += '<tr class="total-row">' +
    '<td class="name">Top ' + RAW.length + ' customers</td>' +
    '<td>' + fmt(t.cur) + '</td>' +
    '<td>' + fmt(t.b30) + '</td>' +
    '<td>' + fmt(t.b60) + '</td>' +
    '<td>' + fmt(t.b90) + '</td>' +
    '<td>' + fmt(t.b90p) + '</td>' +
    '<td>' + fmt(t['60p']) + '</td>' +
    '<td>' + fmt(t.tot) + '</td>' +
    '<td class="div-col">' + fmt(t.ab90) + '</td>' +
    '<td>' + fmt(t.ab90p) + '</td>' +
    '<td>' + fmt(t['a60p']) + '</td>' +
    vcTd(t.b90p - t.ab90p, ' div-col') +
    vcTd(t['60p'] - t['a60p']) +
    '</tr>';

  // --- % of total (Top N bucket / grand total same bucket) ---
  h += '<tr class="pct-row">' +
    '<td class="name">% of total</td>' +
    '<td>' + pct(t.cur,   RECENT_TOTAL.cur) + '</td>' +
    '<td>' + pct(t.b30,   RECENT_TOTAL.b30) + '</td>' +
    '<td>' + pct(t.b60,   RECENT_TOTAL.b60) + '</td>' +
    '<td>' + pct(t.b90,   RECENT_TOTAL.b90) + '</td>' +
    '<td>' + pct(t.b90p,  RECENT_TOTAL.b90p) + '</td>' +
    '<td>' + pct(t['60p'],recent60p) + '</td>' +
    '<td>' + pct(t.tot,   RECENT_TOTAL.tot) + '</td>' +
    '<td class="div-col">' + pct(t.ab90,  OLD_TOTAL.b90) + '</td>' +
    '<td>' + pct(t.ab90p, OLD_TOTAL.b90p) + '</td>' +
    '<td>' + pct(t['a60p'],old60p) + '</td>' +
    '<td class="div-col zero">-</td>' +
    '<td class="zero">-</td>' +
    '</tr>';

  // --- Individual customers ---
RAW.forEach(r => {
    const s = state[r.name];
    const tot   = s.cur + s.b30 + s.b60 + s.b90 + s.b90p; // recalculated from buckets so edits update the row total
    const c60p  = s.b90 + s.b90p;
    const ca60p = s.ab90 + s.ab90p;
    const fc    = r.flag ? ' flagged' : '';
    const flagIcon = r.flag ? ' <span class="flag-icon" title="Negative inconsistency detected">&#9888;</span>' : '';

    h += '<tr class="cust">' +
      '<td class="name' + fc + '">' + esc(r.name) + flagIcon + '</td>' +
      editCell(r.name, 'cur',  s.cur) +
      editCell(r.name, 'b30',  s.b30) +
      editCell(r.name, 'b60',  s.b60) +
      editCell(r.name, 'b90',  s.b90) +
      editCell(r.name, 'b90p', s.b90p) +
      '<td class="' + fc + '">' + fmt(c60p) + '</td>' +
      '<td class="' + fc + '">' + fmt(tot) + '</td>' +
      '<td class="div-col' + fc + '">' + fmt(s.ab90) + '</td>' +
      '<td class="' + fc + '">' + fmt(s.ab90p) + '</td>' +
      '<td class="' + fc + '">' + fmt(ca60p) + '</td>' +
      vcTd(s.b90p - s.ab90p, ' div-col') +
      vcTd(c60p - ca60p) +
      '</tr>';
  });

  document.getElementById('tb').innerHTML = h;

  // Bind input events
document.querySelectorAll('input.cell-edit').forEach(inp => {
inp.addEventListener('change', function () {
      const val = parseInput(this.value);
      if (val !== null) {
        state[this.dataset.name][this.dataset.field] = val;
document.getElementById('recalc-note').style.display = 'block';
        render();
      }
    });
  });

  // --- Flags section ---
  const flaggedRows = RAW.filter(r => r.flag);
  if (!flaggedRows.length) {
document.getElementById('flags-section').style.display = 'none';
    return;
  }
document.getElementById('flags-section').style.display = '';
  let fh = '';
flaggedRows.forEach(r => {
r.flag_detail.forEach((fd, i) => {
      const negClass = fd.recent_val < 0 ? 'val-neg' : 'val-pos';
      const oldClass = fd.old_val   < 0 ? 'val-neg' : 'val-pos';
      fh += '<tr>' +
        '<td>' + (i === 0 ? esc(r.name) : '') + '</td>' +
        '<td>' + esc(fd.bucket) + '</td>' +
        '<td class="' + negClass + '">' + fmt(fd.recent_val) + '</td>' +
        '<td class="' + oldClass + '">' + fmt(fd.old_val) + '</td>' +
        '<td>' + esc(fd.note) + '</td>' +
        '</tr>';
    });
  });
document.getElementById('flags-body').innerHTML = fh;
}

function esc(s){
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

render();

async function copyTableAsImage() {
  const btn   = document.getElementById('btn-copy');
  const label = document.getElementById('copy-label');

  label.textContent = 'Capturing...';
btn.disabled = true;

  try {
    const target = document.querySelector('table');
    const canvas = await html2canvas(target, {
      scale: 2,
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false,
      ignoreElements: el => el.classList.contains('btn-copy') || el.classList.contains('btn-freeze')
    });

    if (navigator.clipboard && window.ClipboardItem) {
canvas.toBlob(async blob => {
        try {
          await navigator.clipboard.write([
            new ClipboardItem({ 'image/png': blob })
          ]);
btn.classList.add('success');
label.textContent = '✓ Copied! Paste in PowerPoint';
          setTimeout(() => {
btn.classList.remove('success');
label.textContent = 'Copy as image';
btn.disabled = false;
          }, 3000);
        } catch (err) {
          downloadFallback(canvas, btn, label);
        }
      }, 'image/png');
    } else {
      downloadFallback(canvas, btn, label);
    }
  } catch (err) {
btn.classList.add('error');
label.textContent = 'Error — try again';
    setTimeout(() => {
btn.classList.remove('error');
label.textContent = 'Copy as image';
btn.disabled = false;
    }, 3000);
  }
}

function downloadFallback(canvas, btn, label) {
  const link = document.createElement('a');
link.download = ${escJs(ctx.outName.replace(/.html$/,'.png'))};
link.href = canvas.toDataURL('image/png');
link.click();
btn.classList.add('success');
label.textContent = '✓ Downloaded as PNG';
  setTimeout(() => {
btn.classList.remove('success');
label.textContent = 'Copy as image';
btn.disabled = false;
  }, 3000);
}

let frozen = false;
function toggleFreeze() {
  frozen = !frozen;
  const btn   = document.getElementById('btn-freeze');
  const label = document.getElementById('freeze-label');
  const icon  = document.getElementById('freeze-icon');

  if (frozen) {
document.body.classList.add('frozen');
btn.classList.add('locked');
label.textContent = 'Unlock report';
icon.innerHTML = '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/>';
  } else {
document.body.classList.remove('frozen');
btn.classList.remove('locked');
label.textContent = 'Lock report';
icon.innerHTML = '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>';
  }
}
</script>
</body>
</html>`;
}

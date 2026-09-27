function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const arFiles = [null, null];

function arDragOver(e, id) {
  e.preventDefault();
  const el = document.getElementById(id);
  if (el) el.style.borderColor = '#1D9E75';
}

function arDragLeave(id) {
  const el = document.getElementById(id);
  if (el) el.style.borderColor = '#D3D1C7';
}

function arDrop(e, idx) {
  e.preventDefault();
  const id = idx === 0 ? 'ar-drop-0' : 'ar-drop-1';
  const drop = document.getElementById(id);
  if (drop) drop.style.borderColor = '#D3D1C7';
  const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) arSetFile(idx, file);
}

function arFileSelected(idx, input) {
  if (input && input.files && input.files[0]) {
    arSetFile(idx, input.files[0]);
  }
}

function arSetFile(idx, file) {
  arFiles[idx] = file;

  const icon = document.getElementById('ar-icon-' + idx);
  const label = document.getElementById('ar-label-' + idx);
  const drop = document.getElementById(idx === 0 ? 'ar-drop-0' : 'ar-drop-1');

  if (icon) {
    icon.className = 'ti ti-file-check';
    icon.style.color = '#1D9E75';
  }

  if (label) {
    label.textContent = file.name;
    label.style.color = '#0F6E56';
  }

  if (drop) {
    drop.style.borderColor = '#1D9E75';
    drop.style.borderStyle = 'solid';
  }

  if (arFiles[0] && arFiles[1]) {
    const btn = document.getElementById('ar-generate-btn');
    if (btn) {
      btn.disabled = false;
      btn.style.background = '#1E2761';
      btn.style.color = '#fff';
      btn.style.cursor = 'pointer';
    }
  }
}

function arLoadScriptIfNeeded(globalName, src) {
  if (window[globalName]) return Promise.resolve();

  return new Promise((resolve, reject) => {
    const existing = document.querySelector('script[src="' + src + '"]');

    if (existing) {
      existing.addEventListener('load', resolve, { once: true });
      existing.addEventListener('error', reject, { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.onload = resolve;
    script.onerror = () => reject(new Error('Unable to load ' + src));
    document.head.appendChild(script);
  });
}

async function arGenerateReport() {
  const btn = document.getElementById('ar-generate-btn');

  if (!arFiles[0] || !arFiles[1]) {
    alert('Please upload both A/R aging files.');
    return;
  }

  if (btn) {
    btn.textContent = 'Processing...';
    btn.disabled = true;
  }

  try {
    await arLoadScriptIfNeeded(
      'XLSX',
      'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
    );

    const workbooks = await Promise.all([
      arReadXLSX(arFiles[0]),
      arReadXLSX(arFiles[1])
    ]);

    const parsed0 = arParseRows(workbooks[0], arFiles[0].name);
    const parsed1 = arParseRows(workbooks[1], arFiles[1].name);

    const recent = parsed0.fileDate >= parsed1.fileDate ? parsed0 : parsed1;
    const old = parsed0.fileDate >= parsed1.fileDate ? parsed1 : parsed0;

    arApplyOffset(recent.customers);
    arApplyOffset(old.customers);

    const top15 = recent.customers
      .slice()
      .sort((a, b) => b.total - a.total)
      .slice(0, 15);

    const oldMap = {};
    old.customers.forEach(row => {
      oldMap[row.customer] = row;
    });

    const flags = arDetectFlags(top15, recent.rawMap, old.rawMap);
    const rd = recent.fileDate;
    const year = Number(rd.slice(0, 4));
    const month = Number(rd.slice(4, 6));

    const blueDate = new Date(year, month - 2, 1);
    const greenDate = new Date(year, month - 3, 1);

    const months = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];

    const blueLabel = months[blueDate.getMonth()] + ' Close';
    const greenLabel = months[greenDate.getMonth()] + ' Close';
    const asOfLabel = rd.slice(4, 6) + '/' + rd.slice(6, 8) + '/' + rd.slice(0, 4);
    const outName = 'AR_Aging_Report_' + months[blueDate.getMonth()] + blueDate.getFullYear() + '.html';

    const html = arBuildHTML({
      top15,
      oldMap,
      gt: recent.grandTotal,
      gtOld: old.grandTotal,
      blueLabel,
      greenLabel,
      asOfLabel,
      flags,
      outName,
      blueYear: blueDate.getFullYear(),
      recentRawDate: recent.rawDate,
      oldRawDate: old.rawDate
    });

    const win = window.open('', '_blank');
    if (!win) throw new Error('Pop-up blocked. Please allow pop-ups and try again.');

    win.document.open();
    win.document.write(html);
    win.document.close();
  } catch (err) {
    console.error(err);
    alert('Error generating report: ' + err.message);
  } finally {
    if (btn) {
      btn.innerHTML = '<i class="ti ti-report-analytics" style="font-size:16px;"></i> Generate Report';
      btn.disabled = false;
    }
  }
}

function arReadXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = event => {
      try {
        resolve(XLSX.read(new Uint8Array(event.target.result), { type: 'array' }));
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () => reject(new Error('Unable to read ' + file.name));
    reader.readAsArrayBuffer(file);
  });
}

function arToNum(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === '') return 0;

  let text = String(value).trim();
  if (!text) return 0;

  const negative = /^\(.*\)$/.test(text) || /^-/.test(text);
  text = text.replace(/[(),$\s]/g, '');

  const number = parseFloat(text);
  if (Number.isNaN(number)) return 0;

  return negative ? -Math.abs(number) : number;
}

function arParseRows(workbook, filename) {
  const match =
    filename.match(/JAZAR-ACO-(\d{2})(\d{2})(\d{4})/i) ||
    filename.match(/(\d{2})(\d{2})(\d{4})/);

  if (!match) {
    throw new Error('Cannot parse date from filename "' + filename + '". Expected MMDDYYYY.');
  }

  const mm = match[1];
  const dd = match[2];
  const yyyy = match[3];
  const fileDate = yyyy + mm + dd;
  const rawDate = mm + dd + yyyy;

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // Exact match avoids mistaking "Customer View" title rows for the header.
  const headerIdx = rows.findIndex(row =>
    String(row && row[0] !== undefined ? row[0] : '').trim().toLowerCase() === 'customer'
  );

  if (headerIdx < 0) {
    throw new Error('Could not find the Customer header in ' + filename + '.');
  }

  const customers = [];
  const rawMap = {};
  let grandTotal = null;

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const first = String(row[0] === undefined ? '' : row[0]).trim();
    const second = String(row[1] === undefined ? '' : row[1]).trim();
    const rowText = row.map(v => String(v === undefined ? '' : v)).join(' ').toLowerCase();

    if (rowText.includes('top15')) break;

    if (i > headerIdx + 1 && first.toLowerCase() === 'customer') break;

    if (first.toLowerCase() === 'total') {
      grandTotal = {
        cur: arToNum(row[1]),
        b30: arToNum(row[2]),
        b60: arToNum(row[3]),
        b90: arToNum(row[4]),
        b90p: arToNum(row[5]),
        total: arToNum(row[6])
      };

      // Some exports put Total before customer rows; others put it last.
      if (customers.length > 0) break;
      continue;
    }

    let customer;
    let offset;

    // Ranked rows look like: 1 | Janssen | Current | 0-30 | ...
    if (/^\d+$/.test(first) && second) {
      customer = second;
      offset = 2;
    } else {
      customer = first;
      offset = 1;
    }

    if (!customer) continue;

    const parsed = {
      cur: arToNum(row[offset]),
      b30: arToNum(row[offset + 1]),
      b60: arToNum(row[offset + 2]),
      b90: arToNum(row[offset + 3]),
      b90p: arToNum(row[offset + 4]),
      total: arToNum(row[offset + 5])
    };

    if (Object.values(parsed).every(value => value === 0) && !row[offset]) continue;

    customers.push({ customer, ...parsed });
    rawMap[customer] = { ...parsed };
  }

  if (!grandTotal) throw new Error('Could not find a Total row in ' + filename + '.');
  if (!customers.length) throw new Error('No customer rows found in ' + filename + '.');

  return { fileDate, rawDate, customers, grandTotal, rawMap };
}

function arApplyOffset(customers) {
  customers.forEach(row => {
    const values = [row.b90p, row.b90, row.b60, row.b30, row.cur];

    for (let i = 0; i < values.length; i++) {
      if (values[i] < 0) {
        let credit = Math.abs(values[i]);
        values[i] = 0;

        for (let j = 0; j < values.length && credit > 0; j++) {
          if (j === i || values[j] <= 0) continue;
          const consumed = Math.min(values[j], credit);
          values[j] -= consumed;
          credit -= consumed;
        }

        if (credit > 0) values[4] -= credit;
      }
    }

    row.b90p = values[0];
    row.b90 = values[1];
    row.b60 = values[2];
    row.b30 = values[3];
    row.cur = values[4];
  });
}

function arDetectFlags(top15, recentRawMap, oldRawMap) {
  const buckets = [
    ['b30', '0 – 30'],
    ['b60', '31 – 60'],
    ['b90', '61 – 90'],
    ['b90p', '90+']
  ];

  const flags = {};

  top15.forEach(row => {
    const recent = recentRawMap[row.customer] || {};
    const old = oldRawMap[row.customer] || {};
    const customerFlags = [];

    buckets.forEach(([key, label]) => {
      const recentValue = Number(recent[key] || 0);
      const oldValue = Number(old[key] || 0);

      if (recentValue < 0 || oldValue < 0) {
        let note;

        if (oldValue < 0 && recentValue >= 0) {
          note = 'Bucket was negative in prior close (' + Math.abs(oldValue).toLocaleString('en-US') + '). Value changed sign in current close. Verify if credit was correctly applied or rebucketed.';
        } else if (recentValue < 0 && oldValue === 0) {
          note = 'New negative appeared in current close (' + Math.abs(recentValue).toLocaleString('en-US') + ') not present in prior close. Verify source of credit.';
        } else {
          note = 'Negative present in both closes. Current: ' + recentValue.toLocaleString('en-US') + ' / Prior: ' + oldValue.toLocaleString('en-US') + '. Verify if amounts are consistent.';
        }

        customerFlags.push({ bucket: label, recentVal: recentValue, oldVal: oldValue, note });
      }
    });

    if (customerFlags.length) flags[row.customer] = customerFlags;
  });

  return flags;
}

function arBuildHTML(ctx) {
  const raw = ctx.top15.map(row => {
    const old = ctx.oldMap[row.customer] || { b90: 0, b90p: 0 };
    const rowFlags = ctx.flags[row.customer] || [];

    return {
      name: row.customer,
      cur: row.cur,
      b30: row.b30,
      b60: row.b60,
      b90: row.b90,
      b90p: row.b90p,
      tot: row.total,
      ab90: old.b90 || 0,
      ab90p: old.b90p || 0,
      flag: rowFlags.length > 0,
      flag_detail: rowFlags.map(flag => ({
        bucket: flag.bucket,
        recent_val: flag.recentVal,
        old_val: flag.oldVal,
        note: flag.note
      }))
    };
  });

  const recentTotal = {
    cur: ctx.gt.cur,
    b30: ctx.gt.b30,
    b60: ctx.gt.b60,
    b90: ctx.gt.b90,
    b90p: ctx.gt.b90p,
    tot: ctx.gt.total
  };

  const oldTotal = {
    b90: ctx.gtOld.b90,
    b90p: ctx.gtOld.b90p,
    tot: ctx.gtOld.total
  };

  const escJs = value => JSON.stringify(value).replace(/</g, '\\u003c');
  const recentMonth = ctx.blueLabel.replace(' Close', '');
  const oldMonth = ctx.greenLabel.replace(' Close', '');
  const title = 'A/R Aging Report — ' + ctx.blueLabel + ' ' + ctx.blueYear;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${esc(title)}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"><\/script>
<style>
*{box-sizing:border-box}body{margin:0;padding:24px;background:#f5f6f8;color:#1a1a2e;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;font-size:12px}.page{max-width:1120px;margin:auto;padding:18px 22px 26px;background:#fff;border-radius:10px;box-shadow:0 2px 16px rgba(0,0,0,.08)}.report-header{display:flex;justify-content:space-between;gap:20px;align-items:flex-start;margin-bottom:20px;padding-bottom:16px;border-bottom:1.5px solid #e2e6ea}.report-header h1{margin:0;color:#0c2340;font-size:18px}.subtitle{margin-top:4px;color:#6b7a8d;font-size:11px}.meta-area{display:flex;flex-direction:column;align-items:flex-end;gap:10px}.meta-pills{display:flex;justify-content:flex-end;flex-wrap:wrap;gap:8px}.pill{padding:4px 10px;border-radius:20px;white-space:nowrap;font-size:10px}.pill-gray{background:#f1efe8;color:#444}.pill-blue{background:#e6f1fb;color:#0c447c}.pill-green{background:#eaf3de;color:#3b6d11}.btn-row{display:flex;gap:8px}.btn-copy,.btn-freeze{padding:6px 13px;border-radius:6px;background:#fff;cursor:pointer;font-size:11px;font-weight:600}.btn-copy{border:1.5px solid #6b7a8d;color:#6b7a8d}.btn-freeze{border:1.5px solid #0c447c;color:#0c447c}.btn-freeze.locked{border-color:#3b6d11;color:#3b6d11}.table-wrap{width:100%;overflow-x:auto}table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:10px}th,td{border:.5px solid #d0d7df;padding:5px}th{text-align:right;line-height:1.2;word-break:break-word}th.left,td.name{text-align:left}th.grp-blue{background:#0c447c;border-color:#185fa5;color:#b5d4f4;text-align:center}th.grp-green{background:#3b6d11;border-color:#639922;color:#c0dd97;text-align:center}th.grp-var{background:#854f0b;border-color:#ba7517;color:#fac775;text-align:center}th.sub{font-size:9px;font-weight:400}td{text-align:right;white-space:nowrap;line-height:1.3}td.name{white-space:normal;word-break:break-word}.total-row td{background:#f0f4f8;border-top:1.5px solid #b0bcc8;font-weight:600}.pct-row td{background:#fafbfc;color:#6b7a8d;font-size:9px}.spacer td{height:8px;border:none;background:#f5f6f8}.div-col{border-left:2px solid #b0bcc8!important}.pos-var{color:#a32d2d;font-weight:600}.neg-var{color:#27670a;font-weight:600}.zero{color:#b0bcc8}.editable{padding:0;background:#f0f7ff}.cell-edit{width:100%;padding:5px;border:0;outline:0;background:transparent;font:inherit;text-align:right}.cell-edit:focus{outline:1.5px solid #378add;background:#e0efff}.flagged,.flagged input{background:#fff7e6!important;color:#7a3e00!important}.flag-icon{color:#ba7517;margin-left:4px}.recalc-note{display:none;margin-top:7px;color:#378add;font-size:10px;font-style:italic}.flags-section{margin-top:24px;overflow:hidden;border:1px solid #f0c880;border-radius:8px}.flags-header{padding:9px 14px;background:#854f0b;color:#fac775;font-weight:600}.flag-table{width:100%;border-collapse:collapse;font-size:11px}.flag-table th{background:#fdf3e0;color:#7a3e00;text-align:left}.flag-table td{text-align:left;white-space:normal;color:#3a2800}.val-neg{color:#a32d2d!important;font-weight:600}body.frozen .editable{background:inherit!important}body.frozen .cell-edit{pointer-events:none}body.frozen .flags-section,body.frozen .recalc-note{display:none!important}body.frozen .flag-icon{display:none}@media print{body{padding:0;background:#fff}.page{box-shadow:none}.btn-row{display:none}}
</style>
</head>
<body>
<div class="page">
<div class="report-header">
<div><h1>A/R Aging Report — Komodo Health</h1><div class="subtitle">Generated from NetSuite export · Rules-based calculation</div></div>
<div class="meta-area">
<div class="meta-pills"><span class="pill pill-gray">Days overdue as of <strong>${esc(ctx.asOfLabel)}</strong></span><span class="pill pill-blue">Blue → ${esc(ctx.blueLabel)} (${esc(ctx.recentRawDate)})</span><span class="pill pill-green">Green → ${esc(ctx.greenLabel)} (${esc(ctx.oldRawDate)})</span></div>
<div class="btn-row"><button class="btn-copy" id="btn-copy" onclick="copyTableAsImage()"><span id="copy-label">Copy as image</span></button><button class="btn-freeze" id="btn-freeze" onclick="toggleFreeze()"><span id="freeze-label">Lock report</span></button></div>
</div>
</div>
<div class="table-wrap"><table><colgroup><col style="width:125px"><col span="6" style="width:70px"><col style="width:80px"><col span="2" style="width:70px"><col style="width:80px"><col span="2" style="width:85px"></colgroup>
<thead><tr><th class="left grp-blue" rowspan="2">Days overdue<br><span style="font-size:8px;font-weight:400">as of ${esc(ctx.asOfLabel)}</span></th><th class="grp-blue" colspan="7">A/R Aging Report (${esc(ctx.blueLabel)})</th><th class="grp-green div-col" colspan="3">A/R Aging Report (${esc(ctx.greenLabel)})</th><th class="grp-var div-col" colspan="2">Variance</th></tr>
<tr><th class="grp-blue sub">Not due yet</th><th class="grp-blue sub">0 – 30</th><th class="grp-blue sub">31 – 60</th><th class="grp-blue sub">61 – 90</th><th class="grp-blue sub">90+</th><th class="grp-blue sub">60+ days total</th><th class="grp-blue sub">Total balance</th><th class="grp-green sub div-col">61 – 90</th><th class="grp-green sub">90+</th><th class="grp-green sub">60+ days total</th><th class="grp-var sub div-col">90+ vs prior close</th><th class="grp-var sub">60+ vs prior close</th></tr></thead>
<tbody id="tb"></tbody></table></div>
<div class="recalc-note" id="recalc-note">* Subtotals and variances recalculated after manual edit.</div>
<div class="flags-section" id="flags-section"><div class="flags-header">⚠ Negative bucket inconsistencies — Top ${raw.length}</div><table class="flag-table"><thead><tr><th>Customer</th><th>Bucket</th><th>${esc(recentMonth)} value</th><th>${esc(oldMonth)} value</th><th>Note</th></tr></thead><tbody id="flags-body"></tbody></table></div>
</div>
<script>
const esc = value => String(value ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
const fmt = n => { n=Number(n||0); if(n===0)return '-'; const a=Math.round(Math.abs(n)).toLocaleString('en-US'); return n<0?'('+a+')':a; };
const pct = (n,d) => d===0 ? '-' : Math.round(n/d*100)+'%';
const vcTd = (v,c='') => v===0 ? '<td class="zero'+c+'">-</td>' : '<td class="'+(v>0?'pos-var':'neg-var')+c+'">'+fmt(v)+'</td>';
const RECENT_TOTAL = ${escJs(recentTotal)}; RECENT_TOTAL['60p']=RECENT_TOTAL.b90+RECENT_TOTAL.b90p;
const OLD_TOTAL = ${escJs(oldTotal)}; OLD_TOTAL['60p']=OLD_TOTAL.b90+OLD_TOTAL.b90p;
const RAW = ${escJs(raw)};
const state = {};
RAW.forEach(r=>state[r.name]={cur:r.cur,b30:r.b30,b60:r.b60,b90:r.b90,b90p:r.b90p,tot:r.tot,ab90:r.ab90,ab90p:r.ab90p});
function calcTopN(){const t={cur:0,b30:0,b60:0,b90:0,b90p:0,tot:0,ab90:0,ab90p:0};RAW.forEach(r=>{const s=state[r.name];t.cur+=s.cur;t.b30+=s.b30;t.b60+=s.b60;t.b90+=s.b90;t.b90p+=s.b90p;t.tot+=s.cur+s.b30+s.b60+s.b90+s.b90p;t.ab90+=s.ab90;t.ab90p+=s.ab90p;});t['60p']=t.b90+t.b90p;t['a60p']=t.ab90+t.ab90p;return t;}
function parseInput(v){v=String(v||'').trim();if(!v)return null;const neg=v.startsWith('(')||v.startsWith('-');const n=parseFloat(v.replace(/[(),\\-\\s,]/g,''));return Number.isNaN(n)?null:(neg?-Math.abs(n):n);}
function editCell(name,field,value,extra=''){const r=RAW.find(x=>x.name===name);const f=r&&r.flag?' flagged':'';return '<td class="editable'+extra+f+'"><input class="cell-edit" data-name="'+esc(name)+'" data-field="'+field+'" value="'+(value===0?'':fmt(value))+'" placeholder="-" /></td>';}
function render(){const r60=RECENT_TOTAL['60p'];const o60=OLD_TOTAL['60p'];let h='';h+='<tr class="total-row"><td class="name">Total amount ($)</td><td>'+fmt(RECENT_TOTAL.cur)+'</td><td>'+fmt(RECENT_TOTAL.b30)+'</td><td>'+fmt(RECENT_TOTAL.b60)+'</td><td>'+fmt(RECENT_TOTAL.b90)+'</td><td>'+fmt(RECENT_TOTAL.b90p)+'</td><td>'+fmt(r60)+'</td><td>'+fmt(RECENT_TOTAL.tot)+'</td><td class="div-col">'+fmt(OLD_TOTAL.b90)+'</td><td>'+fmt(OLD_TOTAL.b90p)+'</td><td>'+fmt(o60)+'</td>'+vcTd(RECENT_TOTAL.b90p-OLD_TOTAL.b90p,' div-col')+vcTd(r60-o60)+'</tr>';
h+='<tr class="pct-row"><td class="name">% of total</td><td>'+pct(RECENT_TOTAL.cur,RECENT_TOTAL.tot)+'</td><td>'+pct(RECENT_TOTAL.b30,RECENT_TOTAL.tot)+'</td><td>'+pct(RECENT_TOTAL.b60,RECENT_TOTAL.tot)+'</td><td>'+pct(RECENT_TOTAL.b90,RECENT_TOTAL.tot)+'</td><td>'+pct(RECENT_TOTAL.b90p,RECENT_TOTAL.tot)+'</td><td>'+pct(r60,RECENT_TOTAL.tot)+'</td><td>100%</td><td class="div-col">'+pct(OLD_TOTAL.b90,OLD_TOTAL.tot)+'</td><td>'+pct(OLD_TOTAL.b90p,OLD_TOTAL.tot)+'</td><td>'+pct(o60,OLD_TOTAL.tot)+'</td><td>-</td><td>-</td></tr><tr class="spacer"><td colspan="13"></td></tr>';
const t=calcTopN();h+='<tr class="total-row"><td class="name">Top '+RAW.length+' customers</td><td>'+fmt(t.cur)+'</td><td>'+fmt(t.b30)+'</td><td>'+fmt(t.b60)+'</td><td>'+fmt(t.b90)+'</td><td>'+fmt(t.b90p)+'</td><td>'+fmt(t['60p'])+'</td><td>'+fmt(t.tot)+'</td><td class="div-col">'+fmt(t.ab90)+'</td><td>'+fmt(t.ab90p)+'</td><td>'+fmt(t['a60p'])+'</td>'+vcTd(t.b90p-t.ab90p,' div-col')+vcTd(t['60p']-t['a60p'])+'</tr>';
h+='<tr class="pct-row"><td class="name">% of total</td><td>'+pct(t.cur,RECENT_TOTAL.cur)+'</td><td>'+pct(t.b30,RECENT_TOTAL.b30)+'</td><td>'+pct(t.b60,RECENT_TOTAL.b60)+'</td><td>'+pct(t.b90,RECENT_TOTAL.b90)+'</td><td>'+pct(t.b90p,RECENT_TOTAL.b90p)+'</td><td>'+pct(t['60p'],r60)+'</td><td>'+pct(t.tot,RECENT_TOTAL.tot)+'</td><td class="div-col">'+pct(t.ab90,OLD_TOTAL.b90)+'</td><td>'+pct(t.ab90p,OLD_TOTAL.b90p)+'</td><td>'+pct(t['a60p'],o60)+'</td><td>-</td><td>-</td></tr>';
RAW.forEach(r=>{const s=state[r.name];const total=s.cur+s.b30+s.b60+s.b90+s.b90p;const c60=s.b90+s.b90p;const o60r=s.ab90+s.ab90p;const f=r.flag?' flagged':'';const icon=r.flag?' <span class="flag-icon">⚠</span>':'';h+='<tr class="cust"><td class="name'+f+'">'+esc(r.name)+icon+'</td>'+editCell(r.name,'cur',s.cur)+editCell(r.name,'b30',s.b30)+editCell(r.name,'b60',s.b60)+editCell(r.name,'b90',s.b90)+editCell(r.name,'b90p',s.b90p)+'<td class="'+f+'">'+fmt(c60)+'</td><td class="'+f+'">'+fmt(total)+'</td><td class="div-col'+f+'">'+fmt(s.ab90)+'</td><td class="'+f+'">'+fmt(s.ab90p)+'</td><td class="'+f+'">'+fmt(o60r)+'</td>'+vcTd(s.b90p-s.ab90p,' div-col')+vcTd(c60-o60r)+'</tr>';});
document.getElementById('tb').innerHTML=h;document.querySelectorAll('.cell-edit').forEach(input=>input.addEventListener('change',function(){const v=parseInput(this.value);if(v===null)return;state[this.dataset.name][this.dataset.field]=v;document.getElementById('recalc-note').style.display='block';render();}));
const flagged=RAW.filter(r=>r.flag);const section=document.getElementById('flags-section');if(!flagged.length){section.style.display='none';return;}section.style.display='';let fh='';flagged.forEach(r=>r.flag_detail.forEach((d,i)=>{fh+='<tr><td>'+(i===0?esc(r.name):'')+'</td><td>'+esc(d.bucket)+'</td><td class="'+(d.recent_val<0?'val-neg':'')+'">'+fmt(d.recent_val)+'</td><td class="'+(d.old_val<0?'val-neg':'')+'">'+fmt(d.old_val)+'</td><td>'+esc(d.note)+'</td></tr>';}));document.getElementById('flags-body').innerHTML=fh;}
render();
async function copyTableAsImage(){const b=document.getElementById('btn-copy');const l=document.getElementById('copy-label');l.textContent='Capturing...';b.disabled=true;try{if(!window.html2canvas)await new Promise((res,rej)=>{const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';s.onload=res;s.onerror=rej;document.head.appendChild(s);});const canvas=await html2canvas(document.querySelector('.table-wrap table'),{scale:2,backgroundColor:'#fff',useCORS:true,logging:false});if(navigator.clipboard&&window.ClipboardItem){canvas.toBlob(async blob=>{try{await navigator.clipboard.write([new ClipboardItem({'image/png':blob})]);l.textContent='✓ Copied! Paste in PowerPoint';}catch(e){downloadFallback(canvas,b,l);}},'image/png');}else downloadFallback(canvas,b,l);}catch(e){l.textContent='Error — try again';setTimeout(()=>{l.textContent='Copy as image';b.disabled=false;},3000);}}
function downloadFallback(canvas,b,l){const a=document.createElement('a');a.download=${escJs(ctx.outName.replace(/\.html$/i,'.png'))};a.href=canvas.toDataURL('image/png');a.click();l.textContent='✓ Downloaded as PNG';setTimeout(()=>{l.textContent='Copy as image';b.disabled=false;},3000);}
let frozen=false;function toggleFreeze(){frozen=!frozen;document.body.classList.toggle('frozen',frozen);const b=document.getElementById('btn-freeze');const l=document.getElementById('freeze-label');b.classList.toggle('locked',frozen);l.textContent=frozen?'Unlock report':'Lock report';}
<\/script>
</body>
</html>`;
}

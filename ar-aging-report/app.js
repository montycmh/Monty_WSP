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

  const file = e.dataTransfer?.files?.[0];
  if (file) arSetFile(idx, file);
}

function arFileSelected(idx, input) {
  if (input?.files?.[0]) {
    arSetFile(idx, input.files[0]);
  }
}

function arSetFile(idx, file) {
  arFiles[idx] = file;

  const icon = document.getElementById(`ar-icon-${idx}`);
  const label = document.getElementById(`ar-label-${idx}`);
  const drop = document.getElementById(
    idx === 0 ? 'ar-drop-0' : 'ar-drop-1'
  );

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

async function arGenerateReport() {
  const btn = document.getElementById('ar-generate-btn');

  if (!arFiles[0] || !arFiles[1]) {
    alert('Please upload both A/R aging files.');
    return;
  }

  if (btn) {
    btn.innerHTML = 'Processing...';
    btn.disabled = true;
  }

  try {
    await arLoadScriptIfNeeded(
      'XLSX',
      'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
    );

    const [wb0, wb1] = await Promise.all([
      arReadXLSX(arFiles[0]),
      arReadXLSX(arFiles[1])
    ]);

    const d0 = arParseRows(wb0, arFiles[0].name);
    const d1 = arParseRows(wb1, arFiles[1].name);

    const recent = d0.fileDate >= d1.fileDate ? d0 : d1;
    const old = d0.fileDate >= d1.fileDate ? d1 : d0;

    arApplyOffset(recent.customers);
    arApplyOffset(old.customers);

    const gt = recent.grandTotal;
    const gtOld = old.grandTotal;

    const top15 = [...recent.customers]
      .sort((a, b) => b.total - a.total)
      .slice(0, 15);

    const oldMap = {};
    old.customers.forEach(row => {
      oldMap[row.customer] = row;
    });

    const flags = arDetectFlags(
      top15,
      recent.rawMap,
      old.rawMap
    );

    const rd = recent.fileDate;

    const blueDate = new Date(
      Number(rd.slice(0, 4)),
      Number(rd.slice(4, 6)) - 2,
      1
    );

    const greenDate = new Date(
      Number(rd.slice(0, 4)),
      Number(rd.slice(4, 6)) - 3,
      1
    );

    const MONTHS = [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December'
    ];

    const blueLabel = `${MONTHS[blueDate.getMonth()]} Close`;
    const greenLabel = `${MONTHS[greenDate.getMonth()]} Close`;

    const asOfLabel =
      `${rd.slice(4, 6)}/${rd.slice(6, 8)}/${rd.slice(0, 4)}`;

    const outName =
      `AR_Aging_Report_${MONTHS[blueDate.getMonth()]}${blueDate.getFullYear()}.html`;

    const html = arBuildHTML({
      top15,
      oldMap,
      gt,
      gtOld,
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

    if (!win) {
      throw new Error(
        'The report window was blocked. Please allow pop-ups and try again.'
      );
    }

    win.document.open();
    win.document.write(html);
    win.document.close();
  } catch (err) {
    console.error(err);
    alert(`Error generating report: ${err.message}`);
  } finally {
    if (btn) {
      btn.innerHTML =
        '<i class="ti ti-report-analytics" style="font-size:16px;"></i> Generate Report';
      btn.disabled = false;
      btn.style.background = '#1E2761';
      btn.style.color = '#fff';
    }
  }
}

function arLoadScriptIfNeeded(globalName, src) {
  if (window[globalName]) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[src="${src}"]`);

    if (existing) {
      existing.addEventListener('load', resolve, { once: true });
      existing.addEventListener('error', reject, { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = src;
    script.async = true;

    script.onload = resolve;
    script.onerror = () => {
      reject(new Error(`Unable to load library: ${src}`));
    };

    document.head.appendChild(script);
  });
}

async function arReadXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = event => {
      try {
        const workbook = XLSX.read(
          new Uint8Array(event.target.result),
          { type: 'array' }
        );

        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () => {
      reject(new Error(`Unable to read file: ${file.name}`));
    };

    reader.readAsArrayBuffer(file);
  });
}

function arToNum(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  if (value === null || value === undefined || value === '') {
    return 0;
  }

  let text = String(value).trim();

  if (!text) {
    return 0;
  }

  const isNegative =
    /^\(.*\)$/.test(text) || /^-/.test(text);

  text = text.replace(/[(),$\s]/g, '');

  const number = parseFloat(text);

  if (Number.isNaN(number)) {
    return 0;
  }

  return isNegative ? -Math.abs(number) : number;
}

function arParseRows(workbook, filename) {
  const match = filename.match(
    /(\d{2})(\d{2})(\d{4})\.(xlsx?|xls)$/i
  );

  if (!match) {
    throw new Error(
      `Cannot parse date from filename "${filename}". Expected MMDDYYYY.xlsx.`
    );
  }

  const mm = match[1];
  const dd = match[2];
  const yyyy = match[3];

  const fileDate = `${yyyy}${mm}${dd}`;
  const rawDate = `${mm}${dd}${yyyy}`;

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ''
  });

  let headerIdx = rows.findIndex(row => {
    const firstCell = String(row?.[0] ?? '').toLowerCase();
    return firstCell.includes('customer');
  });

  const startIdx = headerIdx >= 0 ? headerIdx + 1 : 8;

  const customers = [];
  const rawMap = {};
  let grandTotal = null;

  for (let i = startIdx; i < rows.length; i++) {
    const row = rows[i];

    if (!row || row.length === 0) {
      continue;
    }

    const customer = String(row[0] ?? '').trim();

    if (!customer) {
      continue;
    }

    const parsed = {
      cur: arToNum(row[1]),
      b30: arToNum(row[2]),
      b60: arToNum(row[3]),
      b90: arToNum(row[4]),
      b90p: arToNum(row[5]),
      total: arToNum(row[6])
    };

    if (customer.toLowerCase() === 'total') {
      grandTotal = parsed;
      break;
    }

    const customerRow = {
      customer,
      ...parsed
    };

    customers.push(customerRow);
    rawMap[customer] = { ...parsed };
  }

  if (!grandTotal) {
    throw new Error(
      `Could not find a "Total" row in ${filename}.`
    );
  }

  if (customers.length === 0) {
    throw new Error(
      `No customer rows found in ${filename}.`
    );
  }

  return {
    fileDate,
    rawDate,
    customers,
    grandTotal,
    rawMap
  };
}

function arApplyOffset(customers) {
  customers.forEach(row => {
    const values = [
      row.b90p,
      row.b90,
      row.b60,
      row.b30,
      row.cur
    ];

    for (let i = 0; i < values.length; i++) {
      if (values[i] < 0) {
        let credit = Math.abs(values[i]);
        values[i] = 0;

        for (let j = 0; j < values.length; j++) {
          if (j === i || values[j] <= 0 || credit <= 0) {
            continue;
          }

          const consumed = Math.min(values[j], credit);

          values[j] -= consumed;
          credit -= consumed;
        }

        if (credit > 0) {
          values[4] -= credit;
        }
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
    const customer = row.customer;
    const recent = recentRawMap[customer] || {};
    const old = oldRawMap[customer] || {};
    const customerFlags = [];

    buckets.forEach(([key, label]) => {
      const recentValue = Number(recent[key] || 0);
      const oldValue = Number(old[key] || 0);

      if (recentValue < 0 || oldValue < 0) {
        let note;

        if (oldValue < 0 && recentValue >= 0) {
          note =
            `Bucket was negative in prior close (${Math.abs(oldValue).toLocaleString('en-US')}). ` +
            'Value changed sign in current close. Verify if credit was correctly applied or rebucketed.';
        } else if (recentValue < 0 && oldValue === 0) {
          note =
            `New negative appeared in current close (${Math.abs(recentValue).toLocaleString('en-US')}) ` +
            'not present in prior close. Verify source of credit.';
        } else {
          note =
            `Negative present in both closes. Current: ${recentValue.toLocaleString('en-US')} / ` +
            `Prior: ${oldValue.toLocaleString('en-US')}. Verify if amounts are consistent.`;
        }

        customerFlags.push({
          bucket: label,
          recentVal: recentValue,
          oldVal: oldValue,
          note
        });
      }
    });

    if (customerFlags.length > 0) {
      flags[customer] = customerFlags;
    }
  });

  return flags;
}

function arBuildHTML(ctx) {
  const raw = ctx.top15.map(row => {
    const old = ctx.oldMap[row.customer] || {
      b90: 0,
      b90p: 0
    };

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

  const escJs = value =>
    JSON.stringify(value).replace(/</g, '\\u003c');

  const recentMonthWord =
    ctx.blueLabel.replace(' Close', '');

  const oldMonthWord =
    ctx.greenLabel.replace(' Close', '');

  const reportTitle =
    `A/R Aging Report — ${ctx.blueLabel} ${ctx.blueYear}`;

  const reportHtml = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${esc(reportTitle)}</title>

<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"><\/script>

<style>
* {
  box-sizing: border-box;
}

body {
  margin: 0;
  padding: 24px;
  background: #f5f6f8;
  color: #1a1a2e;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
  font-size: 12px;
}

.page {
  max-width: 1120px;
  margin: 0 auto;
  padding: 18px 22px 26px;
  background: #ffffff;
  border-radius: 10px;
  box-shadow: 0 2px 16px rgba(0, 0, 0, 0.08);
}

.report-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 20px;
  margin-bottom: 20px;
  padding-bottom: 16px;
  border-bottom: 1.5px solid #e2e6ea;
}

.report-header h1 {
  margin: 0;
  color: #0c2340;
  font-size: 18px;
  font-weight: 600;
}

.subtitle {
  margin-top: 4px;
  color: #6b7a8d;
  font-size: 11px;
}

.meta-area {
  display: flex;
  flex-direction: column;
  align-items: flex-end;
  gap: 10px;
}

.meta-pills {
  display: flex;
  justify-content: flex-end;
  flex-wrap: wrap;
  gap: 8px;
}

.pill {
  padding: 4px 10px;
  border-radius: 20px;
  white-space: nowrap;
  font-size: 10px;
  font-weight: 500;
}

.pill-gray {
  background: #f1efe8;
  color: #444441;
}

.pill-blue {
  background: #e6f1fb;
  color: #0c447c;
}

.pill-green {
  background: #eaf3de;
  color: #3b6d11;
}

.btn-row {
  display: flex;
  gap: 8px;
}

.btn-copy,
.btn-freeze {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 13px;
  border-radius: 6px;
  background: #ffffff;
  cursor: pointer;
  font-size: 11px;
  font-weight: 600;
}

.btn-copy {
  border: 1.5px solid #6b7a8d;
  color: #6b7a8d;
}

.btn-freeze {
  border: 1.5px solid #0c447c;
  color: #0c447c;
}

.btn-copy:hover {
  background: #f0f4f8;
}

.btn-freeze:hover {
  background: #e6f1fb;
}

.btn-freeze.locked {
  border-color: #3b6d11;
  color: #3b6d11;
}

.table-wrap {
  width: 100%;
  overflow-x: auto;
}

table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
  font-size: 10px;
}

th,
td {
  border: 0.5px solid #d0d7df;
  padding: 5px 5px;
}

th {
  text-align: right;
  line-height: 1.2;
  word-break: break-word;
}

th.left,
td.name {
  text-align: left;
}

th.grp-blue {
  background: #0c447c;
  border-color: #185fa5;
  color: #b5d4f4;
  text-align: center;
}

th.grp-green {
  background: #3b6d11;
  border-color: #639922;
  color: #c0dd97;
  text-align: center;
}

th.grp-var {
  background: #854f0b;
  border-color: #ba7517;
  color: #fac775;
  text-align: center;
}

th.sub {
  font-size: 9px;
  font-weight: 400;
}

td {
  color: #1a1a2e;
  text-align: right;
  white-space: nowrap;
  line-height: 1.3;
}

td.name {
  white-space: normal;
  word-break: break-word;
  line-height: 1.3;
}

.total-row td {
  background: #f0f4f8;
  border-top: 1.5px solid #b0bcc8;
  font-weight: 600;
}

.pct-row td {
  background: #fafbfc;
  color: #6b7a8d;
  font-size: 9px;
}

.cust:hover td {
  background: #f7f9fb;
}

.spacer td {
  height: 8px;
  border: none;
  background: #f5f6f8;
}

.div-col {
  border-left: 2px solid #b0bcc8 !important;
}

.pos-var {
  color: #a32d2d;
  font-weight: 600;
}

.neg-var {
  color: #27670a;
  font-weight: 600;
}

.zero {
  color: #b0bcc8;
}

.editable {
  padding: 0;
  background: #f0f7ff;
}

.cell-edit {
  width: 100%;
  padding: 5px;
  border: none;
  outline: none;
  background: transparent;
  color: #1a1a2e;
  font: inherit;
  text-align: right;
}

.cell-edit:focus {
  outline: 1.5px solid #378add;
  background: #e0efff;
}

.flagged,
.flagged input {
  background: #fff7e6 !important;
  color: #7a3e00 !important;
}

.flag-icon {
  color: #ba7517;
  font-size: 11px;
  margin-left: 4px;
}

.recalc-note {
  display: none;
  margin-top: 7px;
  color: #378add;
  font-size: 10px;
  font-style: italic;
}

.flags-section {
  margin-top: 24px;
  overflow: hidden;
  border: 1px solid #f0c880;
  border-radius: 8px;
}

.flags-header {
  display: flex;
  align-items: center;
  gap: 7px;
  padding: 9px 14px;
  background: #854f0b;
  color: #fac775;
  font-size: 11px;
  font-weight: 600;
}

.flag-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 11px;
}

.flag-table th {
  padding: 7px 12px;
  background: #fdf3e0;
  border-color: #f0d8a0;
  color: #7a3e00;
  text-align: left;
}

.flag-table td {
  padding: 8px 12px;
  border-color: #f0e8c8;
  color: #3a2800;
  text-align: left;
  white-space: normal;
}

.val-neg {
  color: #a32d2d !important;
  font-weight: 600;
}

.val-pos {
  color: #1a1a2e !important;
}

body.frozen .editable {
  background: inherit !important;
  cursor: default;
}

body.frozen .cell-edit {
  pointer-events: none;
  cursor: default;
}

body.frozen .flags-section,
body.frozen .recalc-note {
  display: none !important;
}

body.frozen .flag-icon {
  display: none;
}

@media print {
  body {
    padding: 0;
    background: #ffffff;
  }

  .page {
    box-shadow: none;
    border-radius: 0;
  }

  .btn-row {
    display: none;
  }
}
</style>
</head>

<body>
<div class="page">

  <div class="report-header">
    <div>
      <h1>A/R Aging Report — Komodo Health</h1>
      <div class="subtitle">
        Generated from NetSuite export · Rules-based calculation
      </div>
    </div>

    <div class="meta-area">
      <div class="meta-pills">
        <span class="pill pill-gray">
          Days overdue as of <strong>${esc(ctx.asOfLabel)}</strong>
        </span>

        <span class="pill pill-blue">
          Blue → ${esc(ctx.blueLabel)} (${esc(ctx.recentRawDate)})
        </span>

        <span class="pill pill-green">
          Green → ${esc(ctx.greenLabel)} (${esc(ctx.oldRawDate)})
        </span>
      </div>

      <div class="btn-row">
        <button class="btn-copy" id="btn-copy" onclick="copyTableAsImage()">
          <span id="copy-label">Copy as image</span>
        </button>

        <button class="btn-freeze" id="btn-freeze" onclick="toggleFreeze()">
          <span id="freeze-label">Lock report</span>
        </button>
      </div>
    </div>
  </div>

  <div class="table-wrap">
    <table>
      <colgroup>
        <col style="width: 125px">
        <col style="width: 75px">
        <col style="width: 67px">
        <col style="width: 67px">
        <col style="width: 67px">
        <col style="width: 67px">
        <col style="width: 78px">
        <col style="width: 80px">
        <col style="width: 67px">
        <col style="width: 67px">
        <col style="width: 78px">
        <col style="width: 82px">
        <col style="width: 82px">
      </colgroup>

      <thead>
        <tr>
          <th class="left grp-blue" rowspan="2">
            Days overdue<br>
            <span style="font-size: 8px; font-weight: 400">
              as of ${esc(ctx.asOfLabel)}
            </span>
          </th>

          <th class="grp-blue" colspan="7">
            A/R Aging Report (${esc(ctx.blueLabel)})
          </th>

          <th class="grp-green div-col" colspan="3">
            A/R Aging Report (${esc(ctx.greenLabel)})
          </th>

          <th class="grp-var div-col" colspan="2">
            Variance
          </th>
        </tr>

        <tr>
          <th class="grp-blue sub">Not due yet</th>
          <th class="grp-blue sub">0 – 30</th>
          <th class="grp-blue sub">31 – 60</th>
          <th class="grp-blue sub">61 – 90</th>
          <th class="grp-blue sub">90+</th>
          <th class="grp-blue sub">60+ days total</th>
          <th class="grp-blue sub">Total balance</th>

          <th class="grp-green sub div-col">61 – 90</th>
          <th class="grp-green sub">90+</th>
          <th class="grp-green sub">60+ days total</th>

          <th class="grp-var sub div-col">90+ vs prior close</th>
          <th class="grp-var sub">60+ vs prior close</th>
        </tr>
      </thead>

      <tbody id="tb"></tbody>
    </table>
  </div>

  <div class="recalc-note" id="recalc-note">
    * Subtotals and variances recalculated after manual edit.
  </div>

  <div class="flags-section" id="flags-section">
    <div class="flags-header">
      ⚠ Negative bucket inconsistencies — Top ${raw.length}
    </div>

    <table class="flag-table">
      <thead>
        <tr>
          <th style="width: 190px">Customer</th>
          <th style="width: 90px">Bucket</th>
          <th style="width: 125px">${esc(recentMonthWord)} value</th>
          <th style="width: 125px">${esc(oldMonthWord)} value</th>
   

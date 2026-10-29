/* =====================================================================
   OPEX Review Board — app logic (browser only; files never leave the browser)
   Review rules: new entries vs missing actuals, vendor-ID / alias matching,
   No Vendor rows reviewed by GL account, Department/GL recommendation by
   frequency (last-month tie-break), reclass export (SpreadsheetML .xls).
   ===================================================================== */
const MAX_BOARDS = 5;

// showBackMessage() lives in ../assets/launchpad-ui.js

const EXCLUDE_TERMS = [
  'A/P', 'A/R', 'Cash and Cash Equivalents', 'Accounts Payable',
  'Taxes Payable', 'Other Liabilities', 'VAT', 'TDS',
  'SVB Checking', 'Expensify', 'Zip Clearing'
];

const RELATED_ENTITY_ALIASES = {
  'dun & bradstreet': ['d&b qtc', 'd and b qtc', 'dnb qtc'],
  'google': ['google cloud', 'google workspace'],
  'amazon': ['amazon web services', 'aws'],
  'microsoft': ['microsoft azure', 'msft']
};

const TAG_META = {
  'Prior': { filter: 'reclass', pillClass: 'tag-prior', rowFilter: 'reclass' },
  'New GL': { filter: 'reclass', pillClass: 'tag-new', rowFilter: 'reclass' },
  'Related': { filter: 'reclass', pillClass: 'tag-related', rowFilter: 'reclass' },
  'No match': { filter: 'review', pillClass: 'tag-nomatch', rowFilter: 'review' },
  'No Vendor': { filter: 'novendor', pillClass: 'tag-novendor', rowFilter: 'novendor' },
  'Missing': { filter: 'missing', pillClass: 'tag-missing', rowFilter: 'missing' }
};

const FILTER_LABELS = {
  all: 'Showing all entries',
  reclass: 'Filtered: New Entries / Reclass',
  review: 'Filtered: Review',
  novendor: 'Filtered: No Vendor',
  missing: 'Filtered: Missing Actuals'
};

const state = {
  files: { gl: null },
  opexSlots: [],
  glRows: null,
  manualOptions: null,
  boards: [],
  activeBoardIndex: 0,
  slotSeq: 0
};

function activeBoard() {
  return state.boards[state.activeBoardIndex] || null;
}

function showView(view) {
  const hasBoards = state.boards.length > 0;
  const loadView = view === 'load' || !hasBoards;
  $('load-panel').classList.toggle('is-hidden', !loadView);
  $('file-panel').classList.toggle('is-hidden', loadView);
  $('file-chip').classList.toggle('is-hidden', loadView);
  if (loadView) $('board-alerts').classList.add('is-hidden');
  else if ($('board-alerts').innerHTML) $('board-alerts').classList.remove('is-hidden');
  document.querySelectorAll('.review-grid').forEach(el => el.classList.toggle('is-hidden', loadView));
  document.querySelectorAll('[data-glean-id="missing-panel"]').forEach(el => el.classList.toggle('is-hidden', loadView));
  $('filter-strip').classList.toggle('is-hidden', loadView);
  $('manage-files').classList.toggle('is-hidden', loadView || !hasBoards);
  $('view-boards').classList.toggle('is-hidden', !hasBoards);
  $('board-tabs').classList.toggle('is-hidden', loadView);
  if (loadView) $('active-filter-banner').classList.remove('is-visible');
  syncTopbarHeight();
}

function syncTopbarHeight() {
  const bar = document.querySelector('.topbar');
  if (bar) document.documentElement.style.setProperty('--topbar-height', `${bar.offsetHeight}px`);
}
window.addEventListener('load', syncTopbarHeight);
window.addEventListener('resize', syncTopbarHeight);

function $(id) { return document.getElementById(id); }
function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
function fileSizeText(size) {
  if (!size && size !== 0) return '';
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
}

function updateLoadUi() {
  $('gl-file-meta').textContent = state.files.gl ? `${state.files.gl.name} · ${fileSizeText(state.files.gl.size)}` : 'No file selected';
  const filledSlots = state.opexSlots.filter(slot => slot.file).length;
  const ready = state.files.gl && filledSlots > 0;
  $('build-board').disabled = !ready;
  $('add-opex').disabled = state.opexSlots.length >= MAX_BOARDS;
  $('load-status').textContent = ready
    ? `Ready to build ${filledSlots} board${filledSlots === 1 ? '' : 's'}.`
    : 'Waiting for a GL file and at least one OPEX file.';
}

function renderOpexSlots() {
  const container = $('opex-slots');
  container.innerHTML = state.opexSlots.map((slot, index) => {
    const meta = slot.file ? `${escapeHtml(slot.file.name)} · ${fileSizeText(slot.file.size)}` : 'No file selected';
    const removable = state.opexSlots.length > 1;
    return `
      <div class="upload-card opex-slot" data-slot-id="${slot.id}" data-glean-id="opex-slot-${index + 1}">
        <div class="slot-head">
          <h4>OPEX board ${index + 1}</h4>
          ${removable ? `<button class="clear-button" type="button" data-action="remove-slot" aria-label="Remove OPEX board">×</button>` : ''}
        </div>
        <p>Supported patterns: <span class="mono">Tbl_*_OPEX Input *.xlsx</span>, <span class="mono">[Tbl]_*_OPEX Input *.xlsx</span> and <span class="mono">Tbl_OPEX_Input_*.xlsx</span>.</p>
        <input type="text" class="board-name-input" data-action="board-name" value="${escapeHtml(slot.name || '')}" placeholder="Board name (optional)" autocomplete="off">
        <div class="picker">
          <button class="secondary-button" type="button" data-action="pick-opex">Choose OPEX file</button>
          <input class="hidden-input opex-input" type="file" accept=".xlsx,.xlsm,.xls,.xlsb">
        </div>
        <div class="file-meta">${meta}</div>
      </div>
    `;
  }).join('');
  updateLoadUi();
  syncTopbarHeight();
}

function addOpexSlot() {
  if (state.opexSlots.length >= MAX_BOARDS) return;
  state.slotSeq += 1;
  state.opexSlots.push({ id: `slot-${state.slotSeq}`, file: null, name: '' });
  renderOpexSlots();
}

function removeOpexSlot(slotId) {
  if (state.opexSlots.length <= 1) return;
  state.opexSlots = state.opexSlots.filter(slot => slot.id !== slotId);
  renderOpexSlots();
}

$('pick-gl').addEventListener('click', () => $('gl-file').click());
$('gl-file').addEventListener('change', (event) => {
  state.files.gl = event.target.files[0] || null;
  updateLoadUi();
});
$('add-opex').addEventListener('click', addOpexSlot);

$('opex-slots').addEventListener('click', (event) => {
  const slotEl = event.target.closest('.opex-slot');
  if (!slotEl) return;
  const slotId = slotEl.dataset.slotId;
  const pick = event.target.closest('[data-action="pick-opex"]');
  if (pick) {
    slotEl.querySelector('.opex-input').click();
    return;
  }
  const remove = event.target.closest('[data-action="remove-slot"]');
  if (remove) {
    removeOpexSlot(slotId);
  }
});

$('opex-slots').addEventListener('change', (event) => {
  const input = event.target.closest('.opex-input');
  if (!input) return;
  const slotEl = event.target.closest('.opex-slot');
  const slot = state.opexSlots.find(item => item.id === slotEl.dataset.slotId);
  if (!slot) return;
  slot.file = input.files[0] || null;
  renderOpexSlots();
});

$('opex-slots').addEventListener('input', (event) => {
  const nameInput = event.target.closest('[data-action="board-name"]');
  if (!nameInput) return;
  const slotEl = event.target.closest('.opex-slot');
  const slot = state.opexSlots.find(item => item.id === slotEl.dataset.slotId);
  if (slot) slot.name = nameInput.value;
});

function validateFilename(name, kind) {
  if (!name) return { ok: false, message: 'Missing file name.' };
  const opexOk = /^\[?Tbl\]?_.+_OPEX[ _]Input.*\.xlsx$/i.test(name) || /^Tbl_OPEX_Input_.+\.xlsx$/i.test(name);
  const glOk = /^GLTransactionSummary.*\.(xlsx|xls)$/i.test(name);
  return kind === 'opex'
    ? { ok: opexOk, message: opexOk ? '' : 'OPEX filename does not match the expected patterns.' }
    : { ok: glOk, message: glOk ? '' : 'GL filename does not match the expected GLTransactionSummary pattern.' };
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellDates: true, raw: true });
}

function directCellValue(sheet, row, col) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })];
  return cell ? cell.v : null;
}

function buildMergeReader(sheet) {
  const mergedValues = new Map();
  const merges = sheet['!merges'] || [];
  for (const merge of merges) {
    const topValue = directCellValue(sheet, merge.s.r + 1, merge.s.c + 1);
    for (let r = merge.s.r + 1; r <= merge.e.r + 1; r += 1) {
      for (let c = merge.s.c + 1; c <= merge.e.c + 1; c += 1) {
        mergedValues.set(`${r}:${c}`, topValue);
      }
    }
  }
  return function getCell(row, col) {
    const value = directCellValue(sheet, row, col);
    if (value !== null && value !== undefined && value !== '') return value;
    return mergedValues.has(`${row}:${col}`) ? mergedValues.get(`${row}:${col}`) : null;
  };
}

function hv(value) {
  if (value === null || value === undefined || value === '') return false;
  const num = Number(value);
  if (!Number.isNaN(num)) return num !== 0;
  return true;
}

function getVid(vendor) {
  const match = String(vendor ?? '').trim().match(/^(\d+)/);
  return match ? match[1] : null;
}

function getVname(vendor) {
  return String(vendor ?? '').replace(/^\d+\s*-\s*/, '').trim();
}

function normalizeName(value) {
  return String(value ?? '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
}

function aliasMatch(a, b) {
  const na = normalizeName(a);
  const nb = normalizeName(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  for (const [canon, aliases] of Object.entries(RELATED_ENTITY_ALIASES)) {
    const group = new Set([normalizeName(canon), ...aliases.map(normalizeName)]);
    if (group.has(na) && group.has(nb)) return true;
  }
  return false;
}

function isExpense(name) {
  const text = String(name ?? '').toLowerCase();
  return !EXCLUDE_TERMS.some(term => text.includes(term.toLowerCase()));
}

function realDept(dept) {
  return !!dept && dept !== '- No Department -' && !String(dept).includes('99999');
}

function realEntity(value) {
  if (value === null || value === undefined) return '';
  const text = String(value).trim();
  return text && text !== '- No Entity -' ? text : '';
}

function displayEntity(row) {
  return realEntity(row['Entity: Name']) || realEntity(row['Entity (Line)']);
}

function glLeaf(glText) {
  const text = String(glText ?? '');
  const parts = text.split(':');
  return parts[parts.length - 1].trim();
}

function fmtAmount(value) {
  if (value === null || value === undefined || value === '') return '';
  const num = Number(value);
  return Number.isNaN(num) ? value : Math.round(num * 100) / 100;
}

function toExcelDateString(value) {
  const date = toJsDate(value);
  if (!date) return value == null ? '' : String(value);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function toJsDate(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const attempt = new Date(value);
  return Number.isNaN(attempt.getTime()) ? null : attempt;
}

function inferMonthYear(label) {
  const match = String(label ?? '').match(/([A-Za-z]{3})\s+(\d{2})/);
  if (!match) return { month: null, year: null };
  const months = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  return { month: months[match[1].toLowerCase()] || null, year: 2000 + Number(match[2]) };
}

function currency(value) {
  if (value === null || value === undefined || value === '') return '—';
  const num = Number(value);
  return Number.isNaN(num) ? String(value) : num.toLocaleString(undefined, { maximumFractionDigits: 0 });
}

function currency2(value) {
  const num = Number(value || 0);
  return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function extractClassFromFilename(name) {
  const match = String(name ?? '').match(/\[Tbl\]_(.+?)_OPEX Input/i);
  return match ? match[1].trim() : '';
}

function iconForCategory(l2) {
  const text = String(l2 ?? '').toLowerCase();
  if (text.includes('professional') || text.includes('consult')) return '<i class="ti ti-briefcase"></i>';
  if (text.includes('recruit')) return '<i class="ti ti-users"></i>';
  if (text.includes('facilit')) return '<i class="ti ti-building"></i>';
  if (text.includes('equipment') || text.includes('software')) return '<i class="ti ti-device-laptop"></i>';
  if (text.includes('tax') || text.includes('income')) return '<i class="ti ti-receipt-tax"></i>';
  if (text.includes('admin') || text.includes('insurance')) return '<i class="ti ti-file-text"></i>';
  return '<i class="ti ti-tag"></i>';
}

function nearestPriorActualValue(row, priorMonths) {
  for (let i = priorMonths.length - 1; i >= 0; i -= 1) {
    const value = row[priorMonths[i]];
    if (hv(value)) return fmtAmount(value);
  }
  return '';
}

function parseOpex(workbook, fileName) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const getCell = buildMergeReader(sheet);
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const maxCol = range.e.c + 1;
  const maxRow = range.e.r + 1;
  const monthNames = [];
  const typeLabels = [];
  for (let c = 1; c <= maxCol; c += 1) {
    monthNames.push(getCell(1, c));
    typeLabels.push(getCell(2, c));
  }
  const actualCols = [];
  const forecastCols = [];
  typeLabels.forEach((label, index) => {
    if (label === 'Actual') actualCols.push(index + 1);
    if (label === 'Working') forecastCols.push(index + 1);
  });
  if (!actualCols.length) throw new Error('No Actual columns found in the OPEX workbook.');
  const lastActualCol = actualCols[actualCols.length - 1];
  const priorActualCols = actualCols.slice(0, -1);
  const lastMon = String(monthNames[lastActualCol - 1] ?? '');
  const priorMon = priorActualCols.map(col => String(monthNames[col - 1] ?? ''));
  const fcastMon = forecastCols.map(col => String(monthNames[col - 1] ?? ''));

  let lastL1 = null;
  let lastL2 = null;
  let lastGl = null;
  const rows = [];
  for (let r = 4; r <= maxRow; r += 1) {
    const l1 = getCell(r, 1);
    const l2 = getCell(r, 2);
    const gl = getCell(r, 3);
    const vendor = getCell(r, 4);
    if (l1) lastL1 = l1;
    if (l2) lastL2 = l2;
    if (gl) lastGl = gl;
    if (!vendor) continue;
    const row = {
      __row_number: r,
      L1: lastL1,
      L2: lastL2,
      GL: lastGl,
      Vendor: String(vendor)
    };
    [...actualCols, ...forecastCols].forEach(col => {
      row[String(monthNames[col - 1] ?? '')] = getCell(r, col);
    });
    rows.push(row);
  }

  const filtered = rows.filter(row => !String(row.L2 ?? '').includes('Comp and Benefits')).map((row, index) => {
    const hasPrior = priorMon.some(mon => hv(row[mon]));
    const hasLast = hv(row[lastMon]);
    const hasFuture = fcastMon.some(mon => hv(row[mon]));
    return {
      ...row,
      __id: `opex-${index + 1}`,
      has_prior: hasPrior,
      has_last: hasLast,
      has_future: hasFuture,
      __nearest_prior_amount: nearestPriorActualValue(row, priorMon)
    };
  });

  const newEntries = filtered.filter(row => row.has_last && !row.has_prior);
  const missingActuals = filtered.filter(row => !row.has_last && row.has_prior && row.has_future);
  return {
    rows: filtered,
    newEntries,
    missingActuals,
    actualCols,
    forecastCols,
    lastActualCol,
    priorActualCols,
    lastMon,
    priorMon,
    fcastMon,
    className: extractClassFromFilename(fileName)
  };
}

function buildNewRows(allRows, newEntries, priorMonths, lastMon) {
  const finalRows = [];
  const addedPrior = new Set();
  newEntries.forEach(currentRow => {
    const vid = getVid(currentRow.Vendor);
    const currentGl = currentRow.GL;
    if (!vid) {
      finalRows.push(makeBoardRow(currentRow, 'No Vendor', 'Review GL lines', priorMonths, lastMon));
      return;
    }
    let sameVidPrior = [];
    if (vid) {
      sameVidPrior = allRows.filter(row => getVid(row.Vendor) === vid && row.GL !== currentGl && row.has_prior);
    }
    if (sameVidPrior.length) {
      const best = sameVidPrior[0];
      const priorGlNum = String(best.GL ?? '').split(' - ')[0].trim();
      if (!addedPrior.has(best.__id)) {
        finalRows.push(makeBoardRow(best, 'Prior', 'No Action', priorMonths, lastMon));
        addedPrior.add(best.__id);
      }
      finalRows.push(makeBoardRow(currentRow, 'New GL', `Request Reclass to GL ${priorGlNum}`, priorMonths, lastMon));
      return;
    }
    const vname = getVname(currentRow.Vendor);
    const related = allRows.filter(row => row.__id !== currentRow.__id && row.GL !== currentGl && row.has_prior && aliasMatch(getVname(row.Vendor), vname));
    if (related.length) {
      const best = related[0];
      if (!addedPrior.has(best.__id)) {
        finalRows.push(makeBoardRow(best, 'Related', 'No Action', priorMonths, lastMon));
        addedPrior.add(best.__id);
      }
      finalRows.push(makeBoardRow(currentRow, 'Related', 'No Action', priorMonths, lastMon));
    } else {
      finalRows.push(makeBoardRow(currentRow, 'No match', 'To be Review / Reclass', priorMonths, lastMon));
    }
  });
  return finalRows;
}

function buildMissingRows(missingActuals, priorMonths, lastMon) {
  return missingActuals.map(row => makeBoardRow(row, 'Missing', 'Missing — check with Acc / Reclass', priorMonths, lastMon));
}

function parseGl(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const getCell = buildMergeReader(sheet);
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const maxCol = range.e.c + 1;
  const maxRow = range.e.r + 1;
  const headers = [];
  for (let c = 1; c <= maxCol; c += 1) headers.push(getCell(7, c));
  const rows = [];
  for (let r = 8; r <= maxRow; r += 1) {
    const row = {};
    let hasAny = false;
    for (let c = 1; c <= maxCol; c += 1) {
      const header = headers[c - 1];
      if (!header) continue;
      const value = getCell(r, c);
      if (value !== null && value !== undefined && value !== '') hasAny = true;
      row[String(header)] = value;
    }
    if (hasAny) rows.push(row);
  }
  return rows;
}

function exportRowPayload(row) {
  return {
    'Name (GL-style)': row['Name (GL-style)'] == null ? '' : String(row['Name (GL-style)']),
    'Department: Name': row['Department: Name'] == null ? '' : String(row['Department: Name']),
    'Description': row['Description'] == null ? '' : String(row['Description']),
    'Memo': row['Memo'] == null ? '' : String(row['Memo']),
    'Entity': displayEntity(row),
    'Transaction Type': row['Transaction Type'] == null ? '' : String(row['Transaction Type']),
    'Document Number': row['Document Number'] == null ? '' : String(row['Document Number']),
    'Date': toExcelDateString(row['Date']),
    'Amount': fmtAmount(row['Amount']),
    'Date Created': toExcelDateString(row['Date Created'])
  };
}

function pickRecommendationWithLastMonthTie(entries, accessor, isLastAccessor) {
  const counts = new Map();
  entries.forEach(entry => {
    const key = accessor(entry);
    if (!key) return;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  if (!counts.size) return null;
  const topCount = Math.max(...counts.values());
  const topKeys = [...counts.entries()].filter(([, count]) => count === topCount).map(([key]) => key);
  const lastMonthMatches = entries.filter(entry => topKeys.includes(accessor(entry)) && isLastAccessor(entry)).map(accessor);
  if (lastMonthMatches.length) {
    const lastCounts = new Map();
    lastMonthMatches.forEach(key => lastCounts.set(key, (lastCounts.get(key) || 0) + 1));
    const bestCount = Math.max(...lastCounts.values());
    return [...lastCounts.entries()].filter(([, count]) => count === bestCount).map(([key]) => key).sort()[0] || null;
  }
  return topKeys.sort()[0] || null;
}

function distribution(entries, accessor) {
  const counts = new Map();
  entries.forEach(entry => {
    const key = accessor(entry);
    if (!key) return;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  const total = [...counts.values()].reduce((sum, value) => sum + value, 0);
  const ranked = [...counts.entries()]
    .map(([key, count]) => ({ key, count }))
    .sort((a, b) => b.count - a.count || String(a.key).localeCompare(String(b.key)));
  return { ranked, total };
}

function freqDetail(dist, total, picked) {
  if (!dist || !dist.length) return '';
  const rows = dist.slice(0, 5).map(item => {
    const pct = total ? Math.round((item.count / total) * 100) : 0;
    const isPick = picked && item.key === picked;
    return `<div class="freq-row ${isPick ? 'is-pick' : ''}">
      <span class="freq-key">${escapeHtml(item.key)}</span>
      <span class="freq-bar"><span style="width:${pct}%"></span></span>
      <span class="freq-count">${item.count} / ${total} · ${pct}%</span>
    </div>`;
  }).join('');
  const extra = dist.length > 5 ? `<div class="freq-more">+${dist.length - 5} more value(s)</div>` : '';
  return `<div class="freq-list">${rows}${extra}</div>`;
}

function lookupVendor(glRows, vendorName, lastMonthNum, lastYear) {
  const vendorLower = String(vendorName ?? '').toLowerCase();
  let matches = glRows.filter(row => String(row['Entity: Name'] ?? '').toLowerCase().includes(vendorLower));
  let foundBy = matches.length ? 'Entity: Name' : null;
  if (!matches.length) {
    matches = glRows.filter(row => String(row['Entity (Line)'] ?? '').toLowerCase().includes(vendorLower));
    foundBy = matches.length ? 'Entity (Line)' : 'Not found';
  }
  if (!matches.length) {
    return { found_by: 'Not found', transactions: [], dept: null, gl: null, count: 0, amt_min: 0, amt_max: 0, amt_avg: 0, export_rows: [] };
  }

  let expenseRows = matches.filter(row => isExpense(row['Name']));
  if (!expenseRows.length) expenseRows = matches;

  const txns = expenseRows.map(row => {
    const date = toJsDate(row['Date']);
    const isLast = !!date && date.getMonth() + 1 === lastMonthNum && date.getFullYear() === lastYear;
    const amount = fmtAmount(row['Amount']);
    const glFull = row['Name (GL-style)'] == null ? '' : String(row['Name (GL-style)']).trim();
    const dept = row['Department: Name'] ? String(row['Department: Name']) : '—';
    return {
      date: toExcelDateString(row['Date']),
      dept,
      gl_short: glLeaf(glFull),
      gl_full: glFull,
      amount: typeof amount === 'number' ? amount : 0,
      is_last: isLast,
      export_row: exportRowPayload(row),
      month_num: date ? date.getMonth() + 1 : null,
      year: date ? date.getFullYear() : null
    };
  });

  const validDeptTxns = txns.filter(txn => realDept(txn.dept));
  const glTxns = txns.filter(txn => txn.gl_full && txn.gl_full !== '—');
  const deptRecommendation = pickRecommendationWithLastMonthTie(validDeptTxns, txn => txn.dept, txn => txn.is_last);
  const glRecommendation = pickRecommendationWithLastMonthTie(glTxns, txn => txn.gl_full, txn => txn.is_last);
  const deptDist = distribution(validDeptTxns, txn => txn.dept);
  const glDist = distribution(glTxns, txn => txn.gl_full);
  const amounts = txns.map(txn => txn.amount).filter(value => typeof value === 'number');

  return {
    found_by: foundBy,
    transactions: txns.slice(0, 80),
    dept: deptRecommendation,
    gl: glRecommendation,
    dept_dist: deptDist.ranked,
    dept_total: deptDist.total,
    gl_dist: glDist.ranked,
    gl_total: glDist.total,
    count: txns.length,
    amt_min: amounts.length ? Math.min(...amounts) : 0,
    amt_max: amounts.length ? Math.max(...amounts) : 0,
    amt_avg: amounts.length ? Math.round((amounts.reduce((sum, value) => sum + value, 0) / amounts.length) * 100) / 100 : 0,
    export_rows: txns.map(txn => ({
      ...txn.export_row,
      _dept: txn.dept,
      _gl_full: txn.gl_full,
      _gl_short: txn.gl_short,
      _amount_num: txn.amount,
      _month_num: txn.month_num,
      _year: txn.year
    }))
  };
}

function makeBoardRow(sourceRow, tag, comment, priorMonths, lastMon) {
  const { month, year } = inferMonthYear(lastMon);
  return {
    id: `${sourceRow.__id}-${tag.replace(/\s+/g, '-').toLowerCase()}`,
    source: sourceRow,
    tag,
    comment,
    filter: TAG_META[tag].rowFilter,
    priorValues: priorMonths.map(mon => sourceRow[mon]),
    lastValue: sourceRow[lastMon],
    lastMonthNum: month,
    lastYear: year,
    sourceGlLeaf: glLeaf(sourceRow.GL),
    fullGl: String(sourceRow.GL ?? ''),
    vendorFull: String(sourceRow.Vendor ?? ''),
    vid: getVid(sourceRow.Vendor),
    vname: getVname(sourceRow.Vendor),
    recommendation: { dept: null, gl: null, source: null },
    lineSel: null,
    nearestPriorAmount: sourceRow.__nearest_prior_amount,
    selected: false
  };
}

function buildManualOptions(glRows) {
  const departments = [...new Set(glRows.map(row => String(row['Department: Name'] ?? '').trim()).filter(value => realDept(value)))].sort((a, b) => a.localeCompare(b));
  const glAccounts = [...new Set(glRows.map(row => String(row['Name (GL-style)'] ?? '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  return { departments, glAccounts };
}

/* ---------- account lookup (No Vendor rows and vendors not found by name) ----------
   The OPEX GL "650100 - Travel - Mileage" matches the GL file "Number" column; the
   board's department code (e.g. 9300) matches "Department: Name". Only the month in review. */
function accountNumber(text) {
  const match = String(text ?? '').match(/^\s*(\d{4,})/);
  return match ? match[1] : '';
}
function glRowAccount(row) {
  const direct = String(row['Number'] ?? '').trim();
  if (direct) return direct;
  const leaf = glLeaf(row['Name (GL-style)']);
  return accountNumber(leaf);
}
// excludeNames: vendors that have their own OPEX row under the same GL (their lines are not part of No Vendor).
function lookupAccountLines(glRows, acct, deptCode, monthNum, year, excludeNames) {
  if (!acct) return { acct, lines: [], total: 0 };
  const skip = (excludeNames || []).map(normalizeName).filter(name => name.length >= 3);
  const isOtherVendor = row => {
    const names = [row['Entity: Name'], row['Entity (Line)']].map(normalizeName).filter(Boolean);
    return skip.some(vendor => names.some(name => name.includes(vendor)));
  };
  const lines = glRows.filter(row => {
    if (glRowAccount(row) !== acct) return false;
    const date = toJsDate(row['Date']);
    if (!date || date.getMonth() + 1 !== monthNum || date.getFullYear() !== year) return false;
    if (deptCode && !String(row['Department: Name'] ?? '').trim().startsWith(deptCode)) return false;
    return isExpense(row['Name']) && !isOtherVendor(row);
  }).map(row => {
    const amount = fmtAmount(row['Amount']);
    return {
      date: toExcelDateString(row['Date']),
      type: String(row['Transaction Type'] ?? ''),
      doc: String(row['Document Number'] ?? ''),
      entity: displayEntity(row) || '—',
      memo: String(row['Memo'] ?? row['Description'] ?? ''),
      dept: row['Department: Name'] ? String(row['Department: Name']) : '—',
      gl_full: String(row['Name (GL-style)'] ?? '').trim(),
      amount: typeof amount === 'number' ? amount : 0,
      export_row: exportRowPayload(row)
    };
  });
  const total = Math.round(lines.reduce((sum, line) => sum + line.amount, 0) * 100) / 100;
  return { acct, lines, total };
}

function buildBoardState(opexData, glRows, fileNames) {
  const newRows = buildNewRows(opexData.rows, opexData.newEntries, opexData.priorMon, opexData.lastMon);
  const missingRows = buildMissingRows(opexData.missingActuals, opexData.priorMon, opexData.lastMon);
  const inferred = inferMonthYear(opexData.lastMon);
  const glData = {};
  opexData.rows.forEach(row => {
    const vid = getVid(row.Vendor);
    const vname = getVname(row.Vendor);
    if (vid && !glData[vid] && !String(row.Vendor).includes('No Vendor')) {
      glData[vid] = { vname, ...lookupVendor(glRows, vname, inferred.month || 1, inferred.year || 1900) };
    }
  });
  const boardLabel = String((opexData.rows[0] && opexData.rows[0].L1) || '').trim();
  const deptCode = accountNumber(boardLabel);
  // Rows reviewed by GL account: No Vendor rows, and vendors not found by name in the GL.
  const acctData = {};
  newRows.forEach(row => {
    const vendorMissing = row.vid && (!glData[row.vid] || glData[row.vid].found_by === 'Not found');
    if (row.tag === 'No Vendor' || vendorMissing) {
      const otherVendors = opexData.rows
        .filter(item => item.GL === row.source.GL && getVid(item.Vendor) && getVid(item.Vendor) !== row.vid)
        .map(item => getVname(item.Vendor));
      acctData[row.id] = lookupAccountLines(glRows, accountNumber(row.fullGl), deptCode, inferred.month || 1, inferred.year || 1900, otherVendors);
    }
  });
  const vendorsWithoutId = [...newRows, ...missingRows].filter(row => !row.vid).length;
  const lookupNotFound = Object.values(glData).filter(item => item.found_by === 'Not found').length;
  const counts = {
    actionable: newRows.length + missingRows.length,
    newRows: newRows.length,
    reviewRows: newRows.filter(row => row.tag === 'No match').length,
    noVendorRows: newRows.filter(row => row.tag === 'No Vendor').length,
    missingRows: missingRows.length,
    reclassRows: newRows.filter(row => row.filter === 'reclass').length,
    vendorsWithoutId,
    lookupNotFound
  };
  return {
    lastMon: opexData.lastMon,
    priorMon: opexData.priorMon,
    fcastMon: opexData.fcastMon,
    className: opexData.className || boardLabel,
    deptCode,
    rows: { newRows, missingRows },
    glData,
    acctData,
    counts,
    fileNames,
    warnings: collectWarnings(fileNames, vendorsWithoutId, lookupNotFound),
    activeFilter: 'all',
    selected: new Set(),
    focusId: null,
    modalQueue: [],
    modalIndex: 0
  };
}

function collectWarnings(fileNames, vendorsWithoutId, lookupNotFound) {
  const warnings = [];
  const opexValidation = validateFilename(fileNames.opexName, 'opex');
  const glValidation = validateFilename(fileNames.glName, 'gl');
  if (!opexValidation.ok) warnings.push({ level: 'warning', text: `${opexValidation.message} Continuing anyway.` });
  if (!glValidation.ok) warnings.push({ level: 'warning', text: `${glValidation.message} Continuing anyway.` });
  return warnings;
}

function setBuildStatus(message) {
  $('load-status').textContent = message;
}

async function buildBoards() {
  if (!state.files.gl) return;
  const slotsWithFiles = state.opexSlots.filter(slot => slot.file);
  if (!slotsWithFiles.length) return;
  if (typeof XLSX === 'undefined') {
    setBuildStatus('The SheetJS library did not load. Check internet access or vendor the library locally.');
    return;
  }
  $('build-board').disabled = true;
  setBuildStatus('Reading GL Transaction Summary... (large files can take a few seconds)');
  try {
    const glWb = await readWorkbook(state.files.gl);
    const glRows = parseGl(glWb);
    state.glRows = glRows;
    state.manualOptions = buildManualOptions(glRows);
    const boards = [];
    for (let i = 0; i < slotsWithFiles.length; i += 1) {
      const slot = slotsWithFiles[i];
      setBuildStatus(`Building board ${i + 1} of ${slotsWithFiles.length}...`);
      const opexWb = await readWorkbook(slot.file);
      const opexData = parseOpex(opexWb, slot.file.name);
      const board = buildBoardState(opexData, glRows, {
        opexName: slot.file.name,
        glName: state.files.gl.name
      });
      board.index = i;
      board.id = `board-${i}`;
      const fallback = opexData.className || slot.file.name.replace(/\.[^.]+$/, '');
      board.name = (slot.name && slot.name.trim()) || fallback;
      boards.push(board);
    }
    state.boards = boards;
    state.activeBoardIndex = 0;
    renderApp();
    showView('boards');
    setBuildStatus(`Built ${boards.length} board${boards.length === 1 ? '' : 's'}.`);
  } catch (error) {
    console.error(error);
    setBuildStatus(`Build failed: ${error.message}`);
  } finally {
    updateLoadUi();
  }
}

$('build-board').addEventListener('click', buildBoards);
$('print-board').addEventListener('click', () => window.print());
$('clear-filter').addEventListener('click', () => applyFilter('all'));

function renderApp() {
  renderManualOptionLists(state.manualOptions);
  renderBoardTabs();
  renderActiveBoard();
}

function renderBoardTabs() {
  const tabs = $('board-tabs');
  if (state.boards.length <= 1) {
    tabs.classList.remove('is-visible');
    tabs.innerHTML = '';
  } else {
    tabs.classList.add('is-visible');
    tabs.innerHTML = state.boards.map((board, index) => `
      <button class="board-tab ${index === state.activeBoardIndex ? 'is-active' : ''}" type="button" data-board-index="${index}" data-glean-id="board-tab-${index + 1}">
        <span>${escapeHtml(board.name)}</span>
        <span class="tab-count">${board.counts.actionable}</span>
      </button>
    `).join('');
  }
  syncTopbarHeight();
}

function switchBoard(index) {
  if (index < 0 || index >= state.boards.length) return;
  state.activeBoardIndex = index;
  renderBoardTabs();
  renderActiveBoard();
}

function renderActiveBoard() {
  const board = activeBoard();
  if (!board) return;
  const boardLabel = state.boards.length > 1 ? `${board.name} · ` : '';
  $('brand-subtitle').textContent = `${boardLabel}${board.lastMon} · Excl. Comp and Benefits · by MONTY`;
  $('last-actual-loaded').textContent = board.lastMon;
  $('count-all').textContent = board.counts.actionable;
  $('count-reclass').textContent = board.counts.reclassRows;
  $('count-review').textContent = board.counts.reviewRows;
  $('count-novendor').textContent = board.counts.noVendorRows;
  $('count-missing').textContent = board.counts.missingRows;
  renderWarnings(board.warnings);
  renderTables();
  $('new-empty').hidden = board.rows.newRows.length > 0;
  $('new-table-wrap').hidden = board.rows.newRows.length === 0;
  $('missing-empty').hidden = board.rows.missingRows.length > 0;
  $('missing-table-wrap').hidden = board.rows.missingRows.length === 0;
  renderHistory();
  updateSelectionUi();
  applyFilter(board.activeFilter);
  scrollMonthsToEnd();
  syncTopbarHeight();
}

function renderManualOptionLists(manualOptions) {
  $('department-options').innerHTML = (manualOptions?.departments || []).map(value => `<option value="${escapeHtml(value)}"></option>`).join('');
  $('gl-options').innerHTML = (manualOptions?.glAccounts || []).map(value => `<option value="${escapeHtml(value)}"></option>`).join('');
}

// Only file-name checks stay as a banner; the rest is shown where it is used (history / file).
function renderWarnings(warnings) {
  const box = $('board-alerts');
  if (!warnings.length) { box.classList.add('is-hidden'); box.innerHTML = ''; return; }
  box.classList.remove('is-hidden');
  box.innerHTML = warnings.map(item => `<div class="alert-line"><i class="ti ti-alert-circle"></i>${escapeHtml(item.text)}</div>`).join('');
}

/* ---------- tables: frozen left (Category, Vendor, GL), scrolling months, frozen right ---------- */
const COLS = { vendor: 190, gl: 250, month: 96, tag: 104, comment: 170, reclass: 190, check: 60 };
const VISIBLE_MONTHS = 3;

function renderHeaders(priorMonths, lastMonth, targetId, prefix) {
  const isNew = prefix === 'new';
  const months = [...priorMonths, lastMonth];
  const right = isNew ? ['tag', 'comment', 'reclass', 'check'] : ['tag', 'comment', 'check'];
  $(`${prefix}-colgroup`).innerHTML =
    `<col style="width:${COLS.vendor}px"><col style="width:${COLS.gl}px">` +
    months.map(() => `<col style="width:${COLS.month}px">`).join('') +
    right.map(key => `<col style="width:${COLS[key]}px">`).join('');
  const table = $(targetId).closest('table');
  const fixed = COLS.vendor + COLS.gl + right.reduce((sum, key) => sum + COLS[key], 0);
  table.style.width = `${fixed + months.length * COLS.month}px`;
  table.style.minWidth = isNew ? '' : '100%';
  // Reclass Review only takes the width of 3 months; the rest of the row goes to Transaction History.
  if (isNew) document.querySelector('.review-grid').style.setProperty('--review-max', `${fixed + Math.min(months.length, VISIBLE_MONTHS) * COLS.month + 2}px`);
  const rightOffsets = stickyRightOffsets(right);
  $(targetId).innerHTML = `
    <th class="fz-l" style="left:0">Vendor</th>
    <th class="fz-l fz-edge-l" style="left:${COLS.vendor}px">GL Account · Category</th>
    ${priorMonths.map(mon => `<th class="month">${escapeHtml(mon)}</th>`).join('')}
    <th class="month is-review">${escapeHtml(lastMonth)}</th>
    <th class="fz-r fz-edge-r" style="right:${rightOffsets.tag}px">Tag</th>
    <th class="fz-r" style="right:${rightOffsets.comment}px">Comment</th>
    ${isNew ? `<th class="fz-r" style="right:${rightOffsets.reclass}px">Reclass to</th>` : ''}
    <th class="fz-r check-col" style="right:0">Select</th>
  `;
}

function stickyRightOffsets(right) {
  const offsets = {};
  let acc = 0;
  for (let i = right.length - 1; i >= 0; i -= 1) { offsets[right[i]] = acc; acc += COLS[right[i]]; }
  return offsets;
}

function tagPill(tag) {
  const cls = TAG_META[tag].pillClass;
  return `<span class="tag-pill ${cls}">${escapeHtml(tag)}</span>`;
}

function commentHtml(comment) {
  if (comment === 'No Action') return '<span class="comment-note comment-na">No Action</span>';
  if (comment.startsWith('Request Reclass')) return `<span class="comment-reclass">${escapeHtml(comment)}</span>`;
  if (comment.startsWith('To be Review')) return `<span class="comment-review">${escapeHtml(comment)}</span>`;
  if (comment.startsWith('Review GL lines')) return `<span class="comment-novendor">${escapeHtml(comment)}</span>`;
  return `<span class="comment-missing">${escapeHtml(comment)}</span>`;
}

function vendorBlock(vendor) {
  const vid = getVid(vendor);
  const vname = getVname(vendor);
  const sub = vid ? `Vendor ID: ${escapeHtml(vid)}` : 'No Vendor ID';
  return `<div class="vendor-name">${escapeHtml(vname || 'No Vendor')}</div><div class="vendor-sub">${sub}</div>`;
}

function reclassToHtml(row, board) {
  const status = rowFileStatus(row, board);
  const rec = row.recommendation;
  if (rec.dept && rec.gl) {
    const src = rec.source === 'manual' ? 'Manual' : 'Recommended';
    const lines = row.lineSel ? ` · ${row.lineSel.size} line${row.lineSel.size === 1 ? '' : 's'}` : '';
    return `<div class="rt ${status.ok ? 'rt-ok' : 'rt-warn'}"><b>${escapeHtml(rec.dept)}</b><span class="mono">${escapeHtml(glLeaf(rec.gl) || rec.gl)}</span><small>${src}${lines}${status.ok ? '' : ' · ' + escapeHtml(status.short)}</small></div>`;
  }
  return `<span class="rt-pending">${board.selected.has(row.id) ? 'Pending · choose in history' : '—'}</span>`;
}

function renderRows(targetId, rows, priorMonths, lastMonth, board) {
  const isNew = targetId === 'new-tbody';
  const right = isNew ? ['tag', 'comment', 'reclass', 'check'] : ['tag', 'comment', 'check'];
  const ro = stickyRightOffsets(right);
  const tbody = $(targetId);
  tbody.innerHTML = rows.map(row => {
    const priorCells = priorMonths.map(mon => `<td class="numeric month">${escapeHtml(currency(row.source[mon]))}</td>`).join('');
    const isSelected = board.selected.has(row.id);
    const isFocus = board.focusId === row.id;
    const rowClass = `board-row${isSelected ? ' is-selected' : ''}${isFocus ? ' is-focus' : ''}`;
    return `
      <tr
        class="${rowClass}"
        data-row-id="${escapeHtml(row.id)}"
        data-glean-id="row-${escapeHtml(board.id)}-${escapeHtml(row.id)}"
        data-section="${isNew ? 'new' : 'missing'}"
        data-filter="${escapeHtml(row.filter)}"
        data-vid="${escapeHtml(row.vid || '')}"
        data-vname="${escapeHtml(row.vname)}"
        data-last-amount="${escapeHtml(row.lastValue === null || row.lastValue === undefined ? '' : String(fmtAmount(row.lastValue)))}"
        data-last-month-num="${escapeHtml(row.lastMonthNum || '')}"
        data-last-year="${escapeHtml(row.lastYear || '')}"
        data-source-gl="${escapeHtml(row.sourceGlLeaf)}"
        data-full-gl="${escapeHtml(row.fullGl)}"
        data-vendor-full="${escapeHtml(row.vendorFull)}"
        data-prev-amount="${escapeHtml(row.nearestPriorAmount === null || row.nearestPriorAmount === undefined ? '' : String(row.nearestPriorAmount))}"
      >
        <td class="fz-l" style="left:0">${vendorBlock(row.source.Vendor)}</td>
        <td class="fz-l fz-edge-l gl-cell" style="left:${COLS.vendor}px"><div class="gl-main mono">${escapeHtml(String(row.source.GL ?? '—'))}</div><div class="gl-cat"><span class="category-icon">${iconForCategory(row.source.L2)}</span>${escapeHtml(String(row.source.L2 ?? '—').replace(/^Total\s+/i, ''))}</div></td>
        ${priorCells}
        <td class="numeric month is-review">${escapeHtml(currency(row.lastValue))}</td>
        <td class="fz-r fz-edge-r" style="right:${ro.tag}px">${tagPill(row.tag)}</td>
        <td class="fz-r" style="right:${ro.comment}px">${commentHtml(row.comment)}</td>
        ${isNew ? `<td class="fz-r reclass-cell" style="right:${ro.reclass}px">${reclassToHtml(row, board)}</td>` : ''}
        <td class="fz-r check-col numeric" style="right:0"><span class="select-box">✓</span></td>
      </tr>
    `;
  }).join('');
}

function renderTables() {
  const board = activeBoard();
  if (!board) return;
  const keep = ['new-table-wrap', 'missing-table-wrap'].map(id => $(id).scrollLeft);
  renderHeaders(board.priorMon, board.lastMon, 'new-header-row', 'new');
  renderHeaders(board.priorMon, board.lastMon, 'missing-header-row', 'missing');
  renderRows('new-tbody', board.rows.newRows, board.priorMon, board.lastMon, board);
  renderRows('missing-tbody', board.rows.missingRows, board.priorMon, board.lastMon, board);
  ['new-table-wrap', 'missing-table-wrap'].forEach((id, i) => { $(id).scrollLeft = keep[i]; });
  applyFilter(board.activeFilter);
}

// Start with the last 3 months in view; older months are to the left of the scroll bar.
function scrollMonthsToEnd() {
  requestAnimationFrame(() => {
    ['new-table-wrap', 'missing-table-wrap'].forEach(id => { const el = $(id); if (el) el.scrollLeft = el.scrollWidth; });
  });
}

function applyFilter(filter) {
  const board = activeBoard();
  if (board) board.activeFilter = filter;
  document.querySelectorAll('.filter-card').forEach(card => card.classList.toggle('is-active', card.dataset.filter === filter));
  const banner = $('active-filter-banner');
  if (filter === 'all') {
    banner.classList.remove('is-visible');
  } else {
    banner.classList.add('is-visible');
    $('active-filter-label').textContent = FILTER_LABELS[filter];
  }
  document.querySelectorAll('.board-row').forEach(row => {
    row.classList.remove('is-dimmed', 'is-highlighted');
    if (filter === 'all') return;
    if (row.dataset.filter === filter) row.classList.add('is-highlighted');
    else row.classList.add('is-dimmed');
  });
}

function rowById(rowId) {
  const board = activeBoard();
  if (!board) return null;
  return [...board.rows.newRows, ...board.rows.missingRows].find(row => row.id === rowId) || null;
}

function isMissingRow(board, row) {
  return board.rows.missingRows.includes(row);
}

function toggleRowSelection(rowElement) {
  const board = activeBoard();
  if (!board || !rowElement) return;
  if (board.activeFilter !== 'all' && rowElement.classList.contains('is-dimmed')) return;
  const rowId = rowElement.dataset.rowId;
  if (!rowId) return;
  if (board.selected.has(rowId)) board.selected.delete(rowId);
  else board.selected.add(rowId);
  const rowData = rowById(rowId);
  if (rowData) rowData.selected = board.selected.has(rowId);
  board.focusId = rowId;
  renderTables();
  renderHistory();
  updateSelectionUi();
}

function boardSelectedNewCount(board) {
  return board ? board.rows.newRows.filter(row => board.selected.has(row.id)).length : 0;
}

/* ---------- file status: what goes into the reclass file, and why a row is still pending ---------- */
function rowFileStatus(row, board) {
  const rec = row.recommendation;
  if (!rec.dept || !rec.gl) return { ok: false, short: 'No department / GL', reason: 'Choose Department and GL in Transaction History.' };
  const acct = board.acctData[row.id];
  if (acct) {
    if (!acct.lines.length) return { ok: false, short: 'No GL lines', reason: `No GL lines for account ${acct.acct} in ${board.lastMon} (department ${board.deptCode || 'any'}).` };
    if (!row.lineSel || !row.lineSel.size) return { ok: false, short: 'No lines selected', reason: 'Select at least one GL line in Transaction History.' };
    return { ok: true };
  }
  const lookup = row.vid ? board.glData[row.vid] : null;
  if (!lookup) return { ok: false, short: 'No GL data', reason: 'Vendor has no GL transactions.' };
  if (!resolveExportSource(lookup.export_rows || [], row)) return { ok: false, short: 'Txn not found', reason: `Original ${board.lastMon} transaction (amount ${currency2(row.lastValue)}) not found in the GL.` };
  return { ok: true };
}

function pendingRows(board) {
  return board.rows.newRows
    .filter(row => board.selected.has(row.id))
    .map(row => ({ row, status: rowFileStatus(row, board) }))
    .filter(item => !item.status.ok);
}

function reclassReadyCount(board) {
  return board ? buildReclassExportRows(board, false).length : 0;
}

function updateSelectionUi() {
  const board = activeBoard();
  $('selected-counter').textContent = String(boardSelectedNewCount(board));
  let consolidated = 0;
  state.boards.forEach(b => {
    consolidated += reclassReadyCount(b);
    consolidated += b.rows.missingRows.filter(row => b.selected.has(row.id)).length;
  });
  $('consolidated-counter').textContent = String(consolidated);
  renderFilePreview();
}

function renderFilePreview() {
  const board = activeBoard();
  if (!board) return;
  const items = buildReclassExportItems(board, false);
  const reclass = items.map(item => item.payload);
  const missing = buildMissingExportRows(board, false);
  const pending = pendingRows(board);
  let consolidated = 0;
  state.boards.forEach(b => { consolidated += buildReclassExportRows(b, false).length + buildMissingExportRows(b, false).length; });
  $('fs-reclass').textContent = reclass.length;
  $('fs-missing').textContent = missing.length;
  $('fs-pending').textContent = pending.length;
  $('fs-consolidated').textContent = consolidated;
  $('chip-reclass').textContent = reclass.length;
  $('chip-missing').textContent = missing.length;
  $('chip-pending').textContent = pending.length;
  $('chip-pending-wrap').classList.toggle('is-hidden', !pending.length);
  $('file-subtitle').textContent = `${board.name} · ${board.lastMon}. Live preview — it updates as you select rows and assign Department and GL.`;

  $('fp-reclass-empty').hidden = reclass.length > 0;
  const SRC_COLS = ['Name (GL-style)', 'Department: Name', 'Description', 'Memo', 'Entity', 'Transaction Type', 'Document Number', 'Date', 'Amount', 'Date Created'];
  const cell = (key, value) => {
    const text = value === null || value === undefined ? '' : String(value);
    if (key === 'Amount') return `<td class="num mono">${escapeHtml(text === '' ? '' : currency2(value))}</td>`;
    const cls = ['Name (GL-style)', 'Document Number', 'Date', 'Date Created'].includes(key) ? 'mono' : '';
    const wide = ['Name (GL-style)', 'Description', 'Memo'].includes(key) ? ' wide' : '';
    return `<td class="${cls}${wide}" title="${escapeHtml(text)}">${escapeHtml(text) || '<span class="blank">—</span>'}</td>`;
  };
  $('fp-reclass').innerHTML = reclass.length ? `
    <thead>
      <tr class="grp"><th rowspan="2" class="match-col">Match <small>preview only</small></th><th colspan="${SRC_COLS.length}" class="g-src">From GL Transaction Summary</th><th class="g-file">File</th><th colspan="2" class="g-to">Reclass to</th></tr>
      <tr>${SRC_COLS.map(key => `<th class="${key === 'Amount' ? 'num' : ''}">${escapeHtml(key)}</th>`).join('')}<th>Data Source</th><th class="g-to-col">Reclass to Business Units</th><th class="g-to-col">Reclass to GL account</th></tr>
    </thead>
    <tbody>${items.map(item => {
      const r = item.payload;
      return `<tr>
        <td class="match-col"><span class="match match-${item.match.level}">${escapeHtml(item.match.text)}</span><small>${escapeHtml(item.row.vname || 'No Vendor')} · ${escapeHtml(board.lastMon)} ${escapeHtml(currency2(item.row.lastValue))}</small></td>
        ${SRC_COLS.map(key => cell(key, r[key])).join('')}
        ${cell('Data Source', r['Data Source'])}
        <td class="to-col"><b>${escapeHtml(r['Reclass to Business Units'])}</b></td>
        <td class="to-col mono" title="${escapeHtml(r['Reclass to GL account'])}"><b>${escapeHtml(r['Reclass to GL account'])}</b></td>
      </tr>`;
    }).join('')}</tbody>` : '';
  $('fp-missing-empty').hidden = missing.length > 0;
  $('fp-missing').innerHTML = missing.length ? `
    <thead><tr><th>Vendor</th><th class="num">Amount TCV</th><th>GL Account</th><th>Class</th></tr></thead>
    <tbody>${missing.map(r => `<tr>
      <td>${escapeHtml(r['Vendor'])}</td>
      <td class="num mono">${escapeHtml(r['Amount TCV'] === '' ? '—' : currency2(r['Amount TCV']))}</td>
      <td class="mono">${escapeHtml(r['GL Account'])}</td>
      <td>${escapeHtml(r['Class'])}</td>
    </tr>`).join('')}</tbody>` : '';
  $('fp-pending-block').classList.toggle('is-hidden', !pending.length);
  $('fp-pending').innerHTML = pending.map(item => `
    <li data-focus-row="${escapeHtml(item.row.id)}"><b>${escapeHtml(item.row.vname || 'No Vendor')}</b><span class="mono">${escapeHtml(item.row.sourceGlLeaf)}</span><em>${escapeHtml(item.status.reason)}</em><i class="ti ti-arrow-right"></i></li>`).join('');
}

/* ---------- Transaction History (right panel) ---------- */
function historyQueue(board) {
  return board.rows.newRows.filter(row => board.selected.has(row.id));
}

function focusRow(rowId, scroll) {
  const board = activeBoard();
  if (!board) return;
  board.focusId = rowId;
  renderTables();
  renderHistory();
  if (scroll) $('history-dock').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function openModalForSelected() {
  const board = activeBoard();
  if (!board) return;
  const queue = historyQueue(board);
  if (!queue.length) return;
  focusRow(queue[0].id, true);
}

function closeModal() {}

function statCards(items) {
  return `<div class="hist-stats">${items.map(([label, value]) => `<div><span>${escapeHtml(label)}</span><b>${escapeHtml(value)}</b></div>`).join('')}</div>`;
}

function assignmentHtml(row, rec, missing) {
  if (missing) return '<div class="hist-note"><i class="ti ti-info-circle"></i>Missing accruals go to the file with vendor, amount, GL and class — no reclass target needed.</div>';
  const current = row.recommendation;
  const dept = current.dept || (rec && rec.dept) || '';
  const gl = current.gl || (rec && rec.gl) || '';
  return `
    <div class="assign-card">
      <div class="assign-head"><b>Reclass to</b>${current.dept && current.gl ? `<span class="assign-src">${current.source === 'manual' ? 'Manual choice' : 'Recommendation applied'}</span>` : '<span class="assign-src is-pending">Not assigned</span>'}</div>
      <label><span>Department</span><input type="text" id="hist-dept" class="manual-input" list="department-options" value="${escapeHtml(dept)}" placeholder="Choose a department" autocomplete="off"></label>
      <label><span>GL account</span><input type="text" id="hist-gl" class="manual-input gl-manual" list="gl-options" value="${escapeHtml(gl)}" placeholder="Choose a GL account" autocomplete="off"></label>
      <div class="assign-actions">
        <button class="secondary-button" id="hist-clear" type="button">Clear</button>
        ${rec && rec.dept && rec.gl ? '<button class="secondary-button" id="modal-apply" type="button">Apply recommendation</button>' : ''}
        <button class="button" id="hist-save" type="button"><i class="ti ti-check"></i> Save</button>
      </div>
    </div>`;
}

function renderHistory() {
  const board = activeBoard();
  const body = $('modal-body');
  const queue = board ? historyQueue(board) : [];
  const row = board && board.focusId ? rowById(board.focusId) : null;
  const idx = row ? queue.findIndex(item => item.id === row.id) : -1;
  $('modal-nav').classList.toggle('is-visible', idx >= 0 && queue.length > 1);
  $('modal-nav-label').textContent = `${idx + 1} of ${queue.length}`;
  if (!row) {
    $('hist-meta').textContent = 'GL lines behind the selected row.';
    body.innerHTML = '<div class="hist-empty"><i class="ti ti-pointer"></i><b>Select a row</b><span>Its GL transactions, the recommendation and your reclass choice show up here.</span></div>';
    return;
  }
  const missing = isMissingRow(board, row);
  const head = `<div class="hist-title"><div><b>${escapeHtml(row.vname || 'No Vendor')}</b><span class="mono">${escapeHtml(row.fullGl)}</span></div>${tagPill(row.tag)}</div>`;
  const acct = board.acctData[row.id];
  if (acct) {
    $('hist-meta').textContent = `By GL account · ${acct.acct} · dept ${board.deptCode || 'all'} · ${board.lastMon}`;
    if (!row.lineSel) row.lineSel = new Set(acct.lines.map((_, i) => i));
    const opexAmount = Number(fmtAmount(row.lastValue)) || 0;
    const checked = acct.lines.reduce((sum, line, i) => sum + (row.lineSel.has(i) ? line.amount : 0), 0);
    const diff = Math.round((checked - opexAmount) * 100) / 100;
    const tie = Math.abs(diff) < 0.01
      ? `<span class="tie ok"><i class="ti ti-circle-check"></i>Matches the OPEX amount</span>`
      : `<span class="tie warn"><i class="ti ti-alert-triangle"></i>Difference ${escapeHtml(currency2(diff))} vs OPEX</span>`;
    const why = row.tag === 'No Vendor' ? 'No vendor ID — showing the GL lines of this account.' : 'Vendor not found by name in the GL — showing the GL lines of this account.';
    body.innerHTML = head + `<div class="hist-note"><i class="ti ti-info-circle"></i>${escapeHtml(why)}</div>` +
      statCards([['OPEX ' + board.lastMon, currency2(opexAmount)], ['Selected lines', currency2(checked)], ['Lines', String(acct.lines.length)]]) +
      (acct.lines.length ? `<div class="tie-row">${tie}<button class="link-button" id="lines-toggle" type="button">${row.lineSel.size === acct.lines.length ? 'Unselect all' : 'Select all'}</button></div>
      <div class="lines-list">${acct.lines.map((line, i) => `
        <label class="line-item${row.lineSel.has(i) ? ' is-on' : ''}"><input type="checkbox" class="line-check" data-line="${i}" ${row.lineSel.has(i) ? 'checked' : ''}>
          <span class="li-main"><b>${escapeHtml(line.entity)}</b><small>${escapeHtml(line.type)}${line.doc ? ' · ' + escapeHtml(line.doc) : ''} · ${escapeHtml(line.date)}</small><small class="li-memo">${escapeHtml(line.memo)}</small></span>
          <span class="li-amt mono">${escapeHtml(currency2(line.amount))}</span>
        </label>`).join('')}</div>` : `<div class="not-found">No GL lines for account ${escapeHtml(acct.acct)} in ${escapeHtml(board.lastMon)}.</div>`) +
      assignmentHtml(row, null, missing);
    return;
  }
  const data = row.vid ? board.glData[row.vid] : null;
  $('hist-meta').textContent = `${board.name} · Vendor ID ${row.vid || '—'} · ${board.fileNames.glName}`;
  if (!data || data.found_by === 'Not found') {
    body.innerHTML = head + '<div class="not-found">Vendor not found in GL Transaction Summary.</div>' + assignmentHtml(row, null, missing);
    return;
  }
  const deptPickCount = (data.dept_dist || []).find(item => item.key === data.dept)?.count || 0;
  const glPickCount = (data.gl_dist || []).find(item => item.key === data.gl)?.count || 0;
  const foundBadgeClass = data.found_by === 'Entity (Line)' ? 'found-badge line' : 'found-badge';
  const txRows = (data.transactions || []).map(txn => `
    <tr class="${txn.is_last ? 'last-month' : ''}">
      <td class="mono">${escapeHtml(txn.date || '—')}</td>
      <td>${escapeHtml(txn.dept || '—')}</td>
      <td class="mono">${escapeHtml(txn.gl_short || txn.gl_full || '—')}</td>
      <td class="mono" style="text-align:right">${escapeHtml(currency2(txn.amount || 0))}</td>
    </tr>`).join('');
  body.innerHTML = head + `<div class="${foundBadgeClass}">Found by ${escapeHtml(data.found_by)}</div>` +
    statCards([['Transactions', String(data.count || 0)], ['Average', currency2(data.amt_avg || 0)], ['Maximum', currency2(data.amt_max || 0)]]) +
    (missing ? '' : `<div class="recommendation-grid">
      <div class="recommendation-card">
        <div class="label">Department recommendation</div>
        <div class="value">${escapeHtml(data.dept || 'No recommendation')}</div>
        ${data.dept ? `<div class="freq-caption">Appears in ${deptPickCount} of ${data.dept_total} department-tagged transaction(s).</div>` : ''}
        ${freqDetail(data.dept_dist, data.dept_total, data.dept)}
      </div>
      <div class="recommendation-card">
        <div class="label">GL recommendation</div>
        <div class="value mono-value">${escapeHtml(data.gl ? glLeaf(data.gl) : 'No recommendation')}</div>
        ${data.gl ? `<div class="freq-caption">Appears in ${glPickCount} of ${data.gl_total} transaction(s).</div>` : ''}
        ${freqDetail((data.gl_dist || []).map(item => ({ key: glLeaf(item.key), count: item.count })), data.gl_total, data.gl ? glLeaf(data.gl) : null)}
      </div>
    </div>`) +
    assignmentHtml(row, data, missing) +
    `<div class="hist-sub">Transactions <small>${board.lastMon} highlighted</small></div>
    <div class="table-wrap txn-wrap">
      <table class="transactions-table">
        <thead><tr><th>Date</th><th>Department</th><th>GL</th><th style="text-align:right">Amount</th></tr></thead>
        <tbody>${txRows}</tbody>
      </table>
    </div>`;
}

function saveAssignment(source) {
  const board = activeBoard();
  const row = board && board.focusId ? rowById(board.focusId) : null;
  if (!row || isMissingRow(board, row)) return;
  let dept, gl;
  if (source === 'rec') {
    const data = row.vid ? board.glData[row.vid] : null;
    if (!data || !data.dept || !data.gl) return;
    dept = data.dept; gl = data.gl;
  } else {
    dept = ($('hist-dept')?.value || '').trim();
    gl = ($('hist-gl')?.value || '').trim();
    if (!dept || !gl) { ($('hist-dept')?.value ? $('hist-gl') : $('hist-dept'))?.focus(); return; }
    const data = row.vid ? board.glData[row.vid] : null;
    if (data && data.dept === dept && data.gl === gl) source = 'rec';
  }
  row.recommendation = { dept, gl, source: source === 'rec' ? 'rec' : 'manual' };
  if (!board.selected.has(row.id)) { board.selected.add(row.id); row.selected = true; }
  renderTables();
  updateSelectionUi();
  // Move on to the next selected row, like the old modal did.
  const queue = historyQueue(board);
  const idx = queue.findIndex(item => item.id === row.id);
  if (idx >= 0 && idx < queue.length - 1) focusRow(queue[idx + 1].id, false);
  else renderHistory();
}

function clearRecommendation() {
  const board = activeBoard();
  const row = board && board.focusId ? rowById(board.focusId) : null;
  if (!row) return;
  row.recommendation = { dept: null, gl: null, source: null };
  renderTables();
  renderHistory();
  updateSelectionUi();
}

function applyCurrentRecommendation() { saveAssignment('rec'); }

function amountEq(a, b) {
  const av = Number(a);
  const bv = Number(b);
  if (Number.isNaN(av) || Number.isNaN(bv)) return false;
  return Math.abs(av - bv) < 0.005;
}

function resolveExportSource(exportRows, boardRow) {
  const boardAmount = Number(fmtAmount(boardRow.lastValue));
  const sourceGlFull = String(boardRow.fullGl || '').trim();
  const sourceGlLeaf = glLeaf(sourceGlFull);
  const month = Number(boardRow.lastMonthNum || 0);
  const year = Number(boardRow.lastYear || 0);
  let matched = exportRows.filter(row =>
    Number(row._month_num || 0) === month &&
    Number(row._year || 0) === year &&
    amountEq(row._amount_num, boardAmount) &&
    String(row['Name (GL-style)'] || '').trim() === sourceGlFull
  );
  if (!matched.length) {
    matched = exportRows.filter(row =>
      Number(row._month_num || 0) === month &&
      Number(row._year || 0) === year &&
      amountEq(row._amount_num, boardAmount) &&
      glLeaf(row['Name (GL-style)'] || '') === sourceGlLeaf
    );
  }
  if (!matched.length) {
    matched = exportRows.filter(row =>
      Number(row._month_num || 0) === month &&
      Number(row._year || 0) === year &&
      amountEq(row._amount_num, boardAmount)
    );
  }
  if (!matched.length) {
    matched = exportRows.filter(row =>
      Number(row._month_num || 0) === month &&
      Number(row._year || 0) === year &&
      glLeaf(row['Name (GL-style)'] || '') === sourceGlLeaf
    );
  }
  if (!matched.length) {
    matched = exportRows.filter(row =>
      amountEq(row._amount_num, boardAmount) && glLeaf(row['Name (GL-style)'] || '') === sourceGlLeaf
    );
  }
  return matched[0] || null;
}

function buildReclassExportRows(board, includeBoard) {
  return buildReclassExportItems(board, includeBoard).map(item => item.payload);
}

// How the exported GL line was found (preview only; the file is not affected).
function matchInfo(source, boardRow) {
  const month = Number(boardRow.lastMonthNum || 0);
  const year = Number(boardRow.lastYear || 0);
  const sameMonth = Number(source._month_num || 0) === month && Number(source._year || 0) === year;
  const sameAmount = amountEq(source._amount_num, Number(fmtAmount(boardRow.lastValue)));
  const fullGl = String(boardRow.fullGl || '').trim();
  const gl = String(source['Name (GL-style)'] || '').trim();
  const sameGl = gl === fullGl;
  const sameLeaf = glLeaf(gl) === glLeaf(fullGl);
  if (sameMonth && sameAmount && (sameGl || sameLeaf)) return { level: 'ok', text: 'Month · amount · GL' };
  if (sameMonth && sameAmount) return { level: 'warn', text: 'Month · amount — GL differs' };
  if (sameMonth && sameLeaf) return { level: 'warn', text: 'Month · GL — amount differs' };
  return { level: 'warn', text: 'Amount · GL — other month' };
}

function buildReclassExportItems(board, includeBoard) {
  return board.rows.newRows
    .filter(row => board.selected.has(row.id) && row.recommendation.dept && row.recommendation.gl)
    .flatMap(row => {
      const acct = board.acctData && board.acctData[row.id];
      if (acct) {
        return acct.lines
          .filter((line, i) => row.lineSel && row.lineSel.has(i))
          .map(line => {
            const payload = {};
            if (includeBoard) payload['Board'] = board.name;
            Object.assign(payload, line.export_row, {
              'Data Source': 'Netsuite',
              'Reclass to Business Units': row.recommendation.dept,
              'Reclass to GL account': row.recommendation.gl
            });
            return { payload, row, match: { level: 'line', text: `GL line · account ${acct.acct}` } };
          });
      }
      const lookup = row.vid ? board.glData[row.vid] : null;
      if (!lookup) return [];
      const source = resolveExportSource(lookup.export_rows || [], row);
      if (!source) return [];
      const payload = {};
      if (includeBoard) payload['Board'] = board.name;
      Object.assign(payload, {
        'Name (GL-style)': source['Name (GL-style)'] || '',
        'Department: Name': source['Department: Name'] || '',
        'Description': source['Description'] || '',
        'Memo': source['Memo'] || '',
        'Entity': source['Entity'] || '',
        'Transaction Type': source['Transaction Type'] || '',
        'Document Number': source['Document Number'] || '',
        'Date': source['Date'] || '',
        'Amount': source['Amount'] == null ? '' : source['Amount'],
        'Date Created': source['Date Created'] || '',
        'Data Source': 'Netsuite',
        'Reclass to Business Units': row.recommendation.dept,
        'Reclass to GL account': row.recommendation.gl
      });
      return [{ payload, row, match: matchInfo(source, row) }];
    });
}

function buildMissingExportRows(board, includeBoard) {
  return board.rows.missingRows
    .filter(row => board.selected.has(row.id))
    .map(row => {
      const payload = {};
      if (includeBoard) payload['Board'] = board.name;
      Object.assign(payload, {
        'Vendor': row.vendorFull || '',
        'Amount TCV': row.nearestPriorAmount === '' ? '' : Number(row.nearestPriorAmount),
        'GL Account': row.fullGl || '',
        'Class': board.className || ''
      });
      return payload;
    })
    .filter(row => row['Vendor'] || row['Amount TCV'] || row['GL Account'] || row['Class']);
}

function xmlEsc(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function ssCell(value, type) {
  const cellType = type || 'String';
  return `<Cell><Data ss:Type="${cellType}">${xmlEsc(value)}</Data></Cell>`;
}

function rowFromValues(values, types) {
  const cells = values.map((value, index) => ssCell(value, types && types[index] ? types[index] : 'String'));
  return `<Row>${cells.join('')}</Row>`;
}

function blankRow(count) {
  return `<Row>${Array.from({ length: count }, () => ssCell('', 'String')).join('')}</Row>`;
}

function numericHeaders(headers) {
  return new Set(['Amount', 'Amount TCV'].filter(h => headers.includes(h)));
}

function valuesFromHeaders(row, headers) {
  return headers.map(header => (row[header] === null || row[header] === undefined) ? '' : row[header]);
}

function typesFromHeaders(row, headers, numericSet) {
  return headers.map(header => (numericSet.has(header) && typeof row[header] === 'number') ? 'Number' : 'String');
}

function downloadExcelXml(reclassHeaders, reclassRows, missingHeaders, missingRows, filename) {
  const maxCols = Math.max(reclassHeaders.length, missingHeaders.length);
  const reclassNumeric = numericHeaders(reclassHeaders);
  const missingNumeric = numericHeaders(missingHeaders);
  const xml = [];
  xml.push('<?xml version="1.0"?>');
  xml.push('<?mso-application progid="Excel.Sheet"?>');
  xml.push('<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">');
  xml.push('<Worksheet ss:Name="Reclass Export"><Table>');
  xml.push(rowFromValues(['RECLASSES'], ['String']));
  xml.push(blankRow(maxCols));
  xml.push(rowFromValues(reclassHeaders, reclassHeaders.map(() => 'String')));
  reclassRows.forEach(row => {
    xml.push(rowFromValues(valuesFromHeaders(row, reclassHeaders), typesFromHeaders(row, reclassHeaders, reclassNumeric)));
  });
  xml.push(blankRow(maxCols));
  xml.push(blankRow(maxCols));
  xml.push(rowFromValues(['MISSING ACCRUALS'], ['String']));
  xml.push(blankRow(maxCols));
  xml.push(rowFromValues(missingHeaders.concat(Array(Math.max(0, maxCols - missingHeaders.length)).fill('')), Array(maxCols).fill('String')));
  missingRows.forEach(row => {
    xml.push(rowFromValues(valuesFromHeaders(row, missingHeaders), typesFromHeaders(row, missingHeaders, missingNumeric)));
  });
  xml.push('</Table></Worksheet></Workbook>');
  const blob = new Blob([xml.join('')], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 600);
}

function safeFileToken(value) {
  return String(value || 'board').replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '').toLowerCase() || 'board';
}

function generateReclassExport() {
  const board = activeBoard();
  if (!board) return;
  const reclassHeaders = ['Name (GL-style)', 'Department: Name', 'Description', 'Memo', 'Entity', 'Transaction Type', 'Document Number', 'Date', 'Amount', 'Date Created', 'Data Source', 'Reclass to Business Units', 'Reclass to GL account'];
  const missingHeaders = ['Vendor', 'Amount TCV', 'GL Account', 'Class'];
  const reclassRows = buildReclassExportRows(board, false);
  const missingRows = buildMissingExportRows(board, false);
  if (!reclassRows.length && !missingRows.length) {
    window.alert('No exportable rows were found from this board’s selection. For reclass rows, apply Department and GL recommendations first.');
    return;
  }
  downloadExcelXml(reclassHeaders, reclassRows, missingHeaders, missingRows, `opex_reclass_${safeFileToken(board.name)}.xls`);
}

function generateConsolidatedExport() {
  if (!state.boards.length) return;
  const reclassHeaders = ['Board', 'Name (GL-style)', 'Department: Name', 'Description', 'Memo', 'Entity', 'Transaction Type', 'Document Number', 'Date', 'Amount', 'Date Created', 'Data Source', 'Reclass to Business Units', 'Reclass to GL account'];
  const missingHeaders = ['Board', 'Vendor', 'Amount TCV', 'GL Account', 'Class'];
  const reclassRows = [];
  const missingRows = [];
  state.boards.forEach(board => {
    reclassRows.push(...buildReclassExportRows(board, true));
    missingRows.push(...buildMissingExportRows(board, true));
  });
  if (!reclassRows.length && !missingRows.length) {
    window.alert('No exportable rows were found across any board. Select rows and apply recommendations first.');
    return;
  }
  downloadExcelXml(reclassHeaders, reclassRows, missingHeaders, missingRows, 'opex_reclass_consolidated.xls');
}

$('board-tabs').addEventListener('click', event => {
  const tab = event.target.closest('.board-tab');
  if (!tab) return;
  switchBoard(Number(tab.dataset.boardIndex));
});

document.addEventListener('click', event => {
  const filterCard = event.target.closest('.filter-card');
  if (filterCard) {
    applyFilter(filterCard.dataset.filter);
    return;
  }
  const btnId = event.target.closest('button')?.id;
  if (btnId === 'reclassify-selected') { openModalForSelected(); return; }
  if (btnId === 'generate-reclass') { generateReclassExport(); return; }
  if (btnId === 'generate-consolidated') { generateConsolidatedExport(); return; }
  if (btnId === 'file-chip') { $('file-panel').scrollIntoView({ behavior: 'smooth', block: 'start' }); return; }
  if (btnId === 'modal-apply') { applyCurrentRecommendation(); return; }
  if (btnId === 'hist-save') { saveAssignment('manual'); return; }
  if (btnId === 'hist-clear') { clearRecommendation(); return; }
  if (btnId === 'lines-toggle') {
    const board = activeBoard(); const row = board && rowById(board.focusId); const acct = row && board.acctData[row.id];
    if (acct) { row.lineSel = row.lineSel.size === acct.lines.length ? new Set() : new Set(acct.lines.map((_, i) => i)); renderHistory(); renderTables(); updateSelectionUi(); }
    return;
  }
  if (btnId === 'modal-prev' || btnId === 'modal-next') {
    const board = activeBoard(); if (!board) return;
    const queue = historyQueue(board);
    const idx = queue.findIndex(item => item.id === board.focusId);
    const next = queue[Math.min(queue.length - 1, Math.max(0, idx + (btnId === 'modal-next' ? 1 : -1)))];
    if (next) focusRow(next.id, false);
    return;
  }
  const pendingItem = event.target.closest('[data-focus-row]');
  if (pendingItem) { focusRow(pendingItem.dataset.focusRow, true); return; }
  const rowElement = event.target.closest('.board-row');
  if (rowElement) {
    if (event.target.closest('button') || event.target.closest('input')) return;
    toggleRowSelection(rowElement);
  }
});

document.addEventListener('change', event => {
  const check = event.target.closest('.line-check');
  if (!check) return;
  const board = activeBoard(); const row = board && rowById(board.focusId);
  if (!row || !row.lineSel) return;
  const i = Number(check.dataset.line);
  if (check.checked) row.lineSel.add(i); else row.lineSel.delete(i);
  renderHistory(); renderTables(); updateSelectionUi();
});

document.addEventListener('keydown', event => {
  if (event.key === 'Enter' && (event.target.id === 'hist-dept' || event.target.id === 'hist-gl')) saveAssignment('manual');
});


$('manage-files').addEventListener('click', () => showView('load'));
$('view-boards').addEventListener('click', () => showView('boards'));

addOpexSlot();
updateLoadUi();
showView('load');

if (window.GleanBridge && typeof window.GleanBridge.postMessage === 'function') {
  window.GleanBridge.postMessage({
    actionId: 'export-pdf',
    type: 'glean-add-menu',
    metadata: { label: 'Export as PDF', icon: 'export' }
  });
  window.GleanBridge.onMessage('action', function(data) {
    if (data.actionId === 'export-pdf') window.print();
  });
}

function prepareChartsForPrint() {}
function cleanUpAfterPrint() {}
window.addEventListener('beforeprint', prepareChartsForPrint);
window.addEventListener('afterprint', cleanUpAfterPrint);

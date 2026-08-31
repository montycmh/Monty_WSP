(function(){
Chart.register(ChartDataLabels);
const state = {
files: { quarter:null, year:null, hc:null, te:null, opex:null },
model: null,
edits: {},
boardCount: 1,
boards: [],          // per-board file sets: [{quarter,year,hc,te,opex}, ...]
generated: []        // built standalone boards: [{title, subtitle, fileBase, html}, ...]
};
const chartRefs = {};
let REVIEW_MONTH_IDX = -1, REVIEW_Q_IDX = -1;
const monthOrder = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const fiscalOrder = ['Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan'];
const dom = {
budgetUtilNav: document.getElementById('budget-util-nav'),
buildBtn: document.getElementById('build-btn'),
resetBtn: document.getElementById('reset-btn'),
backNav: document.getElementById('back-nav'),
homeNav: document.getElementById('home-nav'),
errors: document.getElementById('build-errors'),
root: document.getElementById('dashboard-root'),
sidebarFooter: document.getElementById('sidebar-footer'),
boardCountSeg: document.getElementById('board-count-seg'),
boardGroups: document.getElementById('board-groups'),
launcher: document.getElementById('launcher')
};
const FEED_DEFS = [
{ key:'quarter', title:'QUARTER', sub:'Quarterly OPEX feed' },
{ key:'year',    title:'YEAR',    sub:'Full-year OPEX feed' },
{ key:'hc',      title:'HC',      sub:'Headcount feed' },
{ key:'te',      title:'T&E',     sub:'Travel & expense feed' },
{ key:'opex',    title:'OPEX',    sub:'OPEX Feed (optional)' }
];
// ---------- Dynamic multi-board upload groups ----------
function emptyFiles(){ return { quarter:null, year:null, hc:null, te:null, opex:null }; }
function ensureBoards(count){
while(state.boards.length < count) state.boards.push(emptyFiles());
state.boards.length = count;
}
function boardGroupHtml(boardIdx, single){
const cards = FEED_DEFS.map(def => {
const id = `b${boardIdx}-${def.key}-file`;
const nameId = `b${boardIdx}-${def.key}-name`;
return `<label class="upload-card" for="${id}">
<div class="upload-icon"><i class="ti ti-file-spreadsheet"></i></div>
<div class="upload-title">${def.title}</div>
<div class="upload-sub">${def.sub}</div>
<div class="upload-name" id="${nameId}">Choose file</div>
<input type="file" id="${id}" accept=".xlsx,.xls" hidden data-board="${boardIdx}" data-feed="${def.key}" />
</label>`;
}).join('');
const legend = single
? ''
: `<div class="board-legend"><span class="bg-num">Board ${boardIdx+1}</span><span class="bg-tag" id="bg-tag-${boardIdx}">Awaiting Quarter &amp; Year</span></div>`;
return `<div class="board-group${single?' single':''}" data-board="${boardIdx}">${legend}<div class="upload-grid">${cards}</div></div>`;
}
function renderBoardGroups(){
const count = state.boardCount;
ensureBoards(count);
const single = count === 1;
dom.boardGroups.innerHTML = state.boards.map((_, i) => boardGroupHtml(i, single)).join('');
dom.boardGroups.querySelectorAll('input[type="file"]').forEach(input => {
input.addEventListener('change', onFeedChange);
});
// Re-hydrate labels from any retained state.
state.boards.forEach((files, i) => {
FEED_DEFS.forEach(def => {
const f = files[def.key];
const label = document.getElementById(`b${i}-${def.key}-name`);
if(label) label.textContent = f ? f.name : 'Choose file';
const input = document.getElementById(`b${i}-${def.key}-file`);
if(input) input.closest('.upload-card').classList.toggle('loaded', !!f);
});
updateBoardTag(i);
});
updateBuildButton();
}
function onFeedChange(e){
const input = e.target;
const boardIdx = Number(input.dataset.board);
const feed = input.dataset.feed;
const file = input.files[0] || null;
if(!state.boards[boardIdx]) state.boards[boardIdx] = emptyFiles();
state.boards[boardIdx][feed] = file;
const label = document.getElementById(`b${boardIdx}-${feed}-name`);
if(label) label.textContent = file ? file.name : 'Choose file';
input.closest('.upload-card').classList.toggle('loaded', !!file);
updateBoardTag(boardIdx);
updateBuildButton();
}
function updateBoardTag(boardIdx){
const tag = document.getElementById(`bg-tag-${boardIdx}`);
if(!tag) return;
const files = state.boards[boardIdx] || {};
if(files.quarter && files.year){
let meta = null;
try{ meta = parseMeta(files.quarter.name); }catch(e){}
tag.textContent = meta ? `${meta.dashboardCode} · ${meta.monthToken} ${meta.fyToken}` : 'Ready';
tag.classList.add('ready');
} else {
tag.textContent = 'Awaiting Quarter & Year';
tag.classList.remove('ready');
}
}
function setBoardCount(count){
state.boardCount = count;
if(dom.boardCountSeg){
dom.boardCountSeg.querySelectorAll('.bcs-btn').forEach(b => b.classList.toggle('active', Number(b.dataset.count) === count));
}
if(dom.launcher){ dom.launcher.classList.add('hidden'); dom.launcher.innerHTML = ''; }
state.generated = [];
renderBoardGroups();
}
if(dom.boardCountSeg){
dom.boardCountSeg.querySelectorAll('.bcs-btn').forEach(btn => {
btn.addEventListener('click', () => setBoardCount(Number(btn.dataset.count)));
});
}
renderBoardGroups();
document.getElementById('save-nav').addEventListener('click', saveState);
document.getElementById('download-nav').addEventListener('click', downloadHtml);
const pdfNav = document.getElementById('download-pdf-nav');
if(pdfNav) pdfNav.addEventListener('click', downloadPdf);
if(dom.backNav) dom.backNav.addEventListener('click', backToWorkspace);
dom.resetBtn.addEventListener('click', resetAll);
dom.buildBtn.addEventListener('click', onGenerate);
document.querySelectorAll('aside nav ul li[data-nav]').forEach(li => {
li.addEventListener('click', () => navTo(li.dataset.nav, li));
});
if(dom.budgetUtilNav) dom.budgetUtilNav.addEventListener('click', openBudgetUtilization);
function updateBuildButton(){
ensureBoards(state.boardCount);
const ready = state.boards.length > 0 && state.boards.every(f => f && f.quarter && f.year);
dom.buildBtn.disabled = !ready;
if(dom.buildBtn){
const lbl = state.boardCount > 1 ? `Generate ${state.boardCount} Boards` : 'Generate Board';
dom.buildBtn.innerHTML = `<i class="ti ti-wand"></i> ${lbl}`;
}
}
function showError(msg){ dom.errors.textContent = msg || ''; }
function toast(msg){
let t = document.getElementById('toast');
if(!t){ t = document.createElement('div'); t.id='toast'; t.className='toast'; document.body.appendChild(t); }
t.textContent = msg; t.classList.add('show');
setTimeout(() => t.classList.remove('show'), 1800);
}
function resetAll(){
state.model = null;
state.edits = {};
state.files = { quarter:null, year:null, hc:null, te:null, opex:null };
state.boards = [];
state.generated = [];
dom.root.innerHTML = '';
dom.root.classList.add('hidden');
if(dom.backNav) dom.backNav.classList.add('hidden');
if(dom.homeNav) dom.homeNav.classList.remove('hidden');
if(dom.budgetUtilNav) dom.budgetUtilNav.classList.add('hidden');
if(dom.launcher){ dom.launcher.classList.add('hidden'); dom.launcher.innerHTML = ''; }
showError('');
document.getElementById('upload-shell').classList.remove('hidden');
dom.sidebarFooter.textContent = 'Upload Quarter and Year feeds to generate a board (HC and T&E optional).';
document.querySelectorAll('aside nav ul li').forEach((li, idx) => li.classList.toggle('active', idx===0));
setBoardCount(1);
updateBuildButton();
}
function backToWorkspace(){
document.getElementById('upload-shell').classList.remove('hidden');
dom.root.classList.add('hidden');
if(dom.backNav) dom.backNav.classList.add('hidden');
if(dom.homeNav) dom.homeNav.classList.remove('hidden');
if(dom.budgetUtilNav) dom.budgetUtilNav.classList.add('hidden');
showError('');
document.querySelectorAll('aside nav ul li').forEach((li, idx) => li.classList.toggle('active', idx===0));
window.scrollTo({ top: 0, behavior: 'smooth' });
}
async function buildDashboard(files){
try{
showError('');
dom.buildBtn.disabled = true;
dom.buildBtn.innerHTML = '<i class="ti ti-loader-2"></i> Building...';
const parsed = await parseFiles(files || state.boards[0] || state.files);
const model = deriveModel(parsed);
state.model = model;
renderDashboard(model);
restoreSavedState();
document.getElementById('upload-shell').classList.add('hidden');
dom.root.classList.remove('hidden');
requestAnimationFrame(function(){ document.querySelectorAll('#dashboard-root textarea').forEach(autoResize); });
if(dom.backNav) dom.backNav.classList.remove('hidden');
if(dom.homeNav) dom.homeNav.classList.add('hidden');
if(dom.budgetUtilNav) dom.budgetUtilNav.classList.toggle('hidden', !model.opex);
dom.sidebarFooter.textContent = model.meta.badgeLabel;
}catch(err){
console.error(err);
showError(err.message || 'Could not build the dashboard from the uploaded feeds.');
}finally{
dom.buildBtn.disabled = false;
dom.buildBtn.innerHTML = '<i class="ti ti-wand"></i> Generate Board';
}
}
async function parseFiles(files){
const [quarterWb, yearWb, hcWb, teWb, opexWb] = await Promise.all([
readWorkbook(files.quarter),
readWorkbook(files.year),
files.hc ? readWorkbook(files.hc) : Promise.resolve(null),
files.te ? readWorkbook(files.te) : Promise.resolve(null),
files.opex ? readWorkbook(files.opex) : Promise.resolve(null)
]);
const meta = parseMeta(files.quarter.name);
const opexMeta = files.opex ? parseOpexMeta(files.opex.name) : null;
return {
meta,
quarter: parseQuarterFeed(quarterWb),
year: parseYearFeed(yearWb, meta),
hc: hcWb ? parseHcFeed(hcWb) : null,
te: teWb ? parseTeFeed(teWb) : null,
opex: opexWb ? parseOpexFeed(opexWb, opexMeta) : null
};
}
function readWorkbook(file){
return new Promise((resolve, reject) => {
const reader = new FileReader();
reader.onload = e => {
try{
const wb = XLSX.read(new Uint8Array(e.target.result), { type:'array' });
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
resolve({ workbook: wb, sheet: ws, rows });
}catch(err){ reject(err); }
};
reader.onerror = reject;
reader.readAsArrayBuffer(file);
});
}
function parseMeta(filename){
const cleaned = filename.replace(/\.xlsx?$/i,'');
const fyMatch = cleaned.match(/(FY\d{2})/i);
const monthMatch = cleaned.match(/(?:FY\d{2}[_ -]?)([A-Za-z]{3,9})/i);
const beforeFeed = cleaned.match(/(?:QUARTER|YEAR|HC|TE)[_ -]+(.+?)[_ -]*Feed/i);
let dashboardCode = beforeFeed ? beforeFeed[1].replace(/[_]+/g,' ').replace(/\s+/g,' ').trim() : 'BvA';
if(/^\d+\s*-\s*.+$/.test(dashboardCode)){
dashboardCode = dashboardCode.replace(/^\d+\s*-\s*/, '').trim();
}
dashboardCode = dashboardCode.replace(/\s*Feed$/i,'').trim();
return {
dashboardCode,
fyToken: fyMatch ? fyMatch[1].toUpperCase() : 'FY',
monthToken: monthMatch ? normalizeMonth(monthMatch[1]) : 'Month',
preparedBy: 'Monty'
};
}
function normalizeMonth(str){
const short = String(str||'').slice(0,3).toLowerCase();
const match = monthOrder.find(m => m.toLowerCase() === short);
return match || str;
}
function v(x){
if(x === null || x === undefined || x === '') return 0;
if(typeof x === 'number') return x;
const s = String(x).replace(/[$,]/g,'').trim();
if(!s) return 0;
if(/^\((.*)\)$/.test(s)) return -Number(RegExp.$1);
const n = Number(s);
return Number.isFinite(n) ? n : 0;
}
// Sum three consecutive month cells (a quarter block stores 3 months per metric).
function sum3(row, start){ return v(row[start]) + v(row[start+1]) + v(row[start+2]); }
function fmtK(n){
if(!Number.isFinite(n) || n === 0) return '—';
const rounded = Math.round(n / 1000);
const abs = Math.abs(rounded).toLocaleString('en-US');
if(rounded > 0) return '+$' + abs + 'K';
return '-$' + abs + 'K';
}
function fmtKplain(n){
if(!Number.isFinite(n) || n === 0) return '—';
const rounded = Math.round(n / 1000);
const abs = Math.abs(rounded).toLocaleString('en-US');
return rounded < 0 ? '($' + abs + 'K)' : '$' + abs + 'K';
}
function varClass(n){
if(!n) return 'var-neu';
return n < 0 ? 'var-fav' : 'var-unfav';
}
function varianceBadge(n){
if(!n) return '<span class="badge b-plan">PLAN</span>';
return n < 0 ? '<span class="badge b-fav">↓ FAVORABLE</span>' : '<span class="badge b-unfav">↑ UNFAVORABLE</span>';
}
function classifyRow(label){
const txt = String(label||'').trim();
if(!txt) return 'blank';
if(/^Expense$/i.test(txt)) return 'expense';
if(/^Total\s+/i.test(txt)) return 'l2';
if(/^[67]\d{5}\s*-/.test(txt)) return 'gl';
if(/^\d+\s*-/.test(txt)) return 'vendor';
if(/^No Vendor\s*-/i.test(txt)) return 'novendor';
return 'other';
}
function parseQuarterFeed(wb){
const rows = wb.rows;
const quarterHeader = String(rows[0][5] || rows[0][9] || rows[0][13] || '').trim();
const monthLabels = [String(rows[1][5]||'').trim(), String(rows[1][9]||'').trim(), String(rows[1][13]||'').trim()];
const fcstRaw = [rows[2][3], rows[2][7], rows[2][11], rows[2][15]].map(x => String(x||'')).find(x => /forecast|fcst/i.test(x)) || 'Forecast';
const forecastLabel = fcstRaw.replace(/Forecast/i,'FCST').replace(/\s+/g,' ').trim();
const dataRows = [];
for(let i=3;i<rows.length;i++){
const label = String(rows[i][0] || '').trim();
if(!label) continue;
dataRows.push({
index: i,
label,
rowType: classifyRow(label),
__raw: rows[i],
total: { w:v(rows[i][1]), p:v(rows[i][2]), f:v(rows[i][3]) },
months: [
{ label: monthLabels[0], w:v(rows[i][5]), p:v(rows[i][6]), f:v(rows[i][7]) },
{ label: monthLabels[1], w:v(rows[i][9]), p:v(rows[i][10]), f:v(rows[i][11]) },
{ label: monthLabels[2], w:v(rows[i][13]), p:v(rows[i][14]), f:v(rows[i][15]) }
]
});
}
return { quarterHeader, monthLabels, forecastLabel, dataRows };
}
// FIXED: The Year feed lays out each quarter as a 12-column block grouped by
// metric (3 months of Working, 3 of Plan, 3 of Forecast, 3 of Working-vs-FCST).
// Column map (0-indexed array from sheet_to_json):
//   FY Total: Working=1, Plan=2, Forecast=3, (WvF=4)
//   Q1 -> Working 5,6,7 | Plan 8,9,10  | Fcst 11,12,13
//   Q2 -> Working 17,18,19 | Plan 20,21,22 | Fcst 23,24,25
//   Q3 -> Working 29,30,31 | Plan 32,33,34 | Fcst 35,36,37
//   Q4 -> Working 41,42,43 | Plan 44,45,46 | Fcst 47,48,49
// Each quarter total = sum of its three monthly cells per metric.
function parseYearFeed(wb, meta){
const rows = wb.rows;
const headerRows = rows.slice(0,4);
const qBases = [5, 17, 29, 41]; // start index of each quarter's Working months
const qLabelRow = rows[1] || [];
const dataRows = [];
for(let i=4;i<rows.length;i++){
const label = String(rows[i][0] || '').trim();
if(!label) continue;
const raw = rows[i];
dataRows.push({
index: i,
label,
rowType: classifyRow(label),
__raw: raw,
total: { w:v(raw[1]), p:v(raw[2]), f:v(raw[3]) },
quarters: qBases.map((b, qi) => ({
label: String(qLabelRow[b] || ('Q' + (qi+1))).trim(),
w: sum3(raw, b),      // Working  = 3 month cells
p: sum3(raw, b + 3),  // Plan     = next 3 month cells
f: sum3(raw, b + 6)   // Forecast = next 3 month cells
}))
});
}
const monthCols = detectYearMonthlyColumns(headerRows);
return { headerRows, dataRows, monthCols };
}
function detectYearMonthlyColumns(headerRows){
const monthRow = headerRows[3] || [];
const metricRow = headerRows[2] || [];
const out = [];
let currentMetric = '';
for(let c=0;c<monthRow.length;c++){
const monthLabel = String(monthRow[c] || '').trim();
const rawMetric = String(metricRow[c] || '').trim();
if(rawMetric) currentMetric = rawMetric;
const metric = currentMetric;
if(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i.test(monthLabel) && /working|plan/i.test(metric)){
out.push({ col:c, monthLabel, metric });
}
}
return out;
}
function parseHcFeed(wb){
const rows = wb.rows;
const salaryIdx = rows.findIndex(r => String(r[0]||r[1]||'').toLowerCase().includes('salary accrued'));
if(salaryIdx < 0) throw new Error('Could not find Salary Accrued row in HC feed.');
const salaryRow = rows[salaryIdx];
const employeeRows = rows.slice(salaryIdx+1).filter(r => String(r[0]||'').trim());
return {
salary: {
planTotal: v(salaryRow[1]), q: [v(salaryRow[2]), v(salaryRow[3]), v(salaryRow[4]), v(salaryRow[5])],
workTotal: v(salaryRow[6]), qWork: [v(salaryRow[7]), v(salaryRow[8]), v(salaryRow[9]), v(salaryRow[10])]
},
employeeRows: employeeRows.map(r => ({
name: String(r[0]||'').trim(),
planTotal: v(r[1]), workTotal: v(r[6]), variance: v(r[6]) - v(r[1]),
isTbh: /^TBH\b/i.test(String(r[0]||'').trim())
}))
};
}
function parseTeFeed(wb){
const rows = wb.rows;
const headerIdx = rows.findIndex(r => String(r[0]||'').trim().toLowerCase() === 'employee' && String(r[1]||'').trim().toLowerCase() === 'vendor');
if(headerIdx < 0) throw new Error('Could not find Employee/Vendor header row in T&E feed.');
const header = rows[headerIdx].map(x => String(x||'').trim());
const monthHeaders = header.slice(2).filter(Boolean);
const body = rows.slice(headerIdx+1).filter(r => String(r[0]||'').trim());
const totalIdx = body.findIndex(r => String(r[0]||'').trim().toLowerCase() === 'total');
const detailRows = (totalIdx >= 0 ? body.slice(0,totalIdx) : body).map(r => ({ employee:String(r[0]||'').trim(), vendor:String(r[1]||'').trim(), values: monthHeaders.map((_,i) => v(r[i+2])) }));
const totalRow = totalIdx >= 0 ? { employee:'Total', vendor:'', values: monthHeaders.map((_,i) => v(body[totalIdx][i+2])) } : { employee:'Total', vendor:'', values: monthHeaders.map((_,i) => detailRows.reduce((s,row)=>s+row.values[i],0)) };
return { monthHeaders, detailRows, totalRow };
}
// Nomenclatura: <BU>_OPEXPLAN_<Mmm><YY>  ej. "IT_OPEXPLAN_Jun26"
//   businessUnit = texto antes del primer "_"
//   month token  = texto despues del ultimo "_"  (Jun26 -> Jun + 26)
function parseOpexMeta(filename){
const cleaned = String(filename||'').replace(/\.xlsx?$/i,'');
const parts = cleaned.split('_');
const businessUnit = parts.length ? parts[0].trim() : '';
const monthRaw = parts.length > 1 ? parts[parts.length-1].trim() : '';
const mMatch = monthRaw.match(/^([A-Za-z]{3,9})[ '\-]?(\d{2,4})?$/);
const monthToken = mMatch ? normalizeMonth(mMatch[1]) : normalizeMonth(monthRaw);
const yearToken = mMatch && mMatch[2] ? mMatch[2] : '';
return { businessUnit, monthToken, yearToken, monthRaw, label: 'OPEX Feed' };
}
// Conserva filas y libro crudos para la vista Budget Utilization.
function parseOpexFeed(wb, meta){
return { meta: meta || {}, rows: wb.rows, workbook: wb.workbook, sheet: wb.sheet };
}
function deriveModel(parsed){
const quarterExpense = parsed.quarter.dataRows.find(r => r.rowType === 'expense');
const yearExpense = parsed.year.dataRows.find(r => r.rowType === 'expense');
if(!quarterExpense || !yearExpense) throw new Error('Could not find Expense row in Quarter or Year feed.');
const monthIdx = parsed.quarter.monthLabels.findIndex(m => normalizeMonth(m) === parsed.meta.monthToken);
const effectiveMonthIdx = monthIdx >= 0 ? monthIdx : 0;
const monthLabel = parsed.quarter.monthLabels[effectiveMonthIdx] || parsed.meta.monthToken;
const badgeLabel = `${parsed.meta.dashboardCode} · ${parsed.meta.monthToken} ${parsed.meta.fyToken} · ${parsed.quarter.quarterHeader}`;
const l2QuarterRows = parsed.quarter.dataRows.filter(r => r.rowType === 'l2');
const l2YearRows = parsed.year.dataRows.filter(r => r.rowType === 'l2');
const quarterPlanVar = l2QuarterRows.map(r => ({ ...r, variance: r.total.w - r.total.p }));
const quarterFcstVar = l2QuarterRows.map(r => ({ ...r, variance: r.total.w - r.total.f }));
const yearPlanVar = l2YearRows.map(r => ({ ...r, variance: r.total.w - r.total.p }));
const yearFcstVar = l2YearRows.map(r => ({ ...r, variance: r.total.w - r.total.f }));
const quarterTopPlan = rankTop(quarterPlanVar, 3);
const quarterTopFcst = rankTop(quarterFcstVar, 3);
const yearTopPlan = rankTop(yearPlanVar, 3);
const yearTopFcst = rankTop(yearFcstVar, 3);
const quarterDriverBlocksPlan = deriveQuarterDriverBlocks(parsed.quarter, quarterPlanVar, 'p', 35000);
const quarterDriverBlocksFcst = deriveQuarterDriverBlocks(parsed.quarter, quarterFcstVar, 'f', 35000);
const yearDriverBlocksPlan = deriveYearDriverBlocks(parsed.year, yearPlanVar, 'p', 75000);
const yearDriverBlocksFcst = deriveYearDriverBlocks(parsed.year, yearFcstVar, 'f', 75000);
const hc = parsed.hc ? deriveHc(parsed.hc) : null;
const trend = deriveTrend(parsed.year, parsed.meta.monthToken, yearExpense.label);
const te = parsed.te ? deriveTe(parsed.te) : null;
const actions = deriveActions(parsed, quarterTopPlan, yearTopPlan, quarterDriverBlocksPlan, hc, te);
return {
meta: {
dashboardCode: parsed.meta.dashboardCode,
fyToken: parsed.meta.fyToken,
monthToken: parsed.meta.monthToken,
currentQuarterLabel: parsed.quarter.quarterHeader,
forecastLabel: parsed.quarter.forecastLabel,
badgeLabel,
preparedBy: parsed.meta.preparedBy
},
quarter: {
monthLabels: parsed.quarter.monthLabels,
allRows: parsed.quarter.dataRows,
expense: quarterExpense,
l2Rows: l2QuarterRows,
kpis: {
kpi1: quarterExpense.months[effectiveMonthIdx].w - quarterExpense.months[effectiveMonthIdx].p,
kpi2: quarterExpense.months[effectiveMonthIdx].w - quarterExpense.months[effectiveMonthIdx].f,
kpi3: quarterExpense.total.w - quarterExpense.total.p,
kpi4: quarterExpense.total.w - quarterExpense.total.f
},
topPlan: quarterTopPlan,
topFcst: quarterTopFcst,
driverBlocksPlan: quarterDriverBlocksPlan,
driverBlocksFcst: quarterDriverBlocksFcst,
currentMonthLabel: monthLabel
},
year: {
elapsedMonths: trend.labels,
allRows: parsed.year.dataRows,
expense: yearExpense,
l2Rows: l2YearRows,
kpis: { kpi5: yearExpense.total.w - yearExpense.total.p, kpi6: yearExpense.total.w - yearExpense.total.f },
trend,
topPlan: yearTopPlan,
topFcst: yearTopFcst,
driverBlocksPlan: yearDriverBlocksPlan,
driverBlocksFcst: yearDriverBlocksFcst
},
hc,
te,
opex: parsed.opex || null,
actions
};
}
function rankTop(rows, n){
const sorted = rows.slice().sort((a,b) => Math.abs(b.variance) - Math.abs(a.variance));
const top = sorted.slice(0,n);
const totalVar = rows.reduce((s,r)=>s+r.variance,0);
const others = totalVar - top.reduce((s,r)=>s+r.variance,0);
return { rows: top, others, totalVar };
}
function deriveQuarterDriverBlocks(q, rows, benchmarkKey, threshold){
const qualified = rows.filter(r => Math.abs(r.variance) > threshold);
return qualified.map(r => ({
id: slug(r.label + benchmarkKey),
label: r.label,
variance: r.variance,
benchmarkKey,
vendors: extractQuarterVendors(q.dataRows, r, benchmarkKey),
comments: []
}));
}
function deriveYearDriverBlocks(y, rows, benchmarkKey, threshold){
return rows.filter(r => Math.abs(r.variance) > threshold).map(r => ({
id: slug(r.label + benchmarkKey),
label: r.label,
variance: r.variance,
benchmarkKey,
vendors: [{ name:'No Vendor -', variance:r.variance, comment: summarizeVariance(r.label, r.variance, 'full-year') }],
comments: []
}));
}
function extractQuarterVendors(dataRows, targetRow, benchmarkKey){
const idx = dataRows.findIndex(r => r.index === targetRow.index);
const vendors = [];
for(let i = idx - 1; i >= 0; i--){
const row = dataRows[i];
if(row.rowType === 'l2' || row.rowType === 'expense') break;
if(row.rowType === 'vendor' || row.rowType === 'novendor'){
const bench = row.total[benchmarkKey] || 0;
const variance = row.total.w - bench;
vendors.push({
name: row.label,
variance,
comment: summarizeVariance(row.label, variance, 'quarter')
});
}
}
const picked = vendors.sort((a,b)=>Math.abs(b.variance)-Math.abs(a.variance)).slice(0,4);
if(!picked.length){
return [{ name:'No Vendor -', variance:targetRow.variance, comment:'Variance appears pooled inside unattributed detail rows. Review timing and spread within this L2 block.' }];
}
return picked;
}
function summarizeVariance(name, variance, scope){
const direction = variance < 0 ? 'below benchmark / favorable' : 'above benchmark / unfavorable';
const cleaned = String(name||'').replace(/^\d+\s*-\s*/,'').replace(/^Total\s+/,'');
return `${cleaned} is ${direction} at ${fmtK(variance)} on a ${scope} basis. Validate timing, scope, and whether the run-rate should persist into the remaining periods.`;
}
function deriveTrend(yearFeed, monthToken, expenseLabel){
const expenseRow = yearFeed.dataRows.find(r => r.label === expenseLabel) || yearFeed.dataRows.find(r => r.rowType === 'expense');
const cutoff = fiscalOrder.indexOf(monthToken);
const allowed = cutoff >= 0 ? fiscalOrder.slice(0, cutoff + 1) : fiscalOrder.slice(0, 1);
const workingCols = yearFeed.monthCols.filter(c => /working/i.test(c.metric) && !/vs/i.test(c.metric));
const planCols = yearFeed.monthCols.filter(c => /plan/i.test(c.metric));
const labels = [];
const working = [];
const plan = [];
allowed.forEach(mon => {
const wCol = workingCols.find(c => normalizeMonth(c.monthLabel) === mon);
const pCol = planCols.find(c => normalizeMonth(c.monthLabel) === mon);
if(wCol && pCol && expenseRow && expenseRow.__raw){
labels.push(mon);
working.push(v(expenseRow.__raw[wCol.col]));
plan.push(v(expenseRow.__raw[pCol.col]));
}
});
if(!labels.length){
return { labels: allowed, working: allowed.map(()=>0), plan: allowed.map(()=>0) };
}
return { labels, working, plan };
}
function deriveHc(hc){
const rows = hc.employeeRows;
const activePlan = rows.filter(r => !r.isTbh && r.planTotal > 0).length;
const activeWork = rows.filter(r => !r.isTbh && r.planTotal > 0 && r.workTotal > 0).length;
const tbhPlan = rows.filter(r => r.isTbh && r.planTotal > 0).length;
const tbhWork = rows.filter(r => r.isTbh && r.workTotal > 0).length;
const newHires = rows.filter(r => r.planTotal === 0 && r.workTotal > 0).length;
const endingPlan = rows.filter(r => r.planTotal > 0).length;
const endingWork = rows.filter(r => r.workTotal > 0).length;
const topVarianceEmployees = rows.slice().sort((a,b)=>Math.abs(b.variance)-Math.abs(a.variance)).slice(0,4).map(r => ({
...r,
comment: hcComment(r)
}));
return {
salaryAccrued: hc.salary,
employeeRows: rows,
movementCounts: { activePlan, activeWork, tbhPlan, tbhWork, newHires, endingPlan, endingWork },
topVarianceEmployees
};
}
function hcComment(r){
if(r.isTbh && r.workTotal === 0) return 'Open req not filled yet; plan remains in place while working has not started.';
if(r.isTbh && r.workTotal > 0) return 'TBH line is now flowing through working, consistent with an in-year hire against an open req.';
if(r.planTotal === 0 && r.workTotal > 0) return 'Working cost is showing without plan budget, consistent with a new hire or transfer not included in plan.';
if(r.variance < 0) return 'Working is below plan, likely driven by delayed fill, lower run rate, or partial-year employment.';
if(r.variance > 0) return 'Working is above plan, likely driven by earlier start timing, higher comp, or accrual timing.';
return 'No material variance versus plan.';
}
function deriveTe(te){
const rows = te.detailRows.map(r => ({ ...r, grandTotal: r.values.reduce((s,n)=>s+n,0) }));
const totalRow = { ...te.totalRow, grandTotal: te.totalRow.values.reduce((s,n)=>s+n,0) };
return { headers: ['Employee','Vendor',...te.monthHeaders,'Grand Total'], rows, totalRow };
}
function deriveActions(parsed, quarterTopPlan, yearTopPlan, qBlocks, hc, te){
const actions = [];
quarterTopPlan.rows.slice(0,2).forEach(r => actions.push(`Review ${r.label} at ${fmtK(r.variance)} versus plan and confirm whether the quarter run-rate should persist.`));
yearTopPlan.rows.slice(0,2).forEach(r => actions.push(`Confirm the full-year outlook for ${r.label}; current variance is ${fmtK(r.variance)} versus plan.`));
if(qBlocks[0] && qBlocks[0].vendors[0]) actions.push(`Validate the vendor driver for ${qBlocks[0].label}, especially ${qBlocks[0].vendors[0].name}.`);
if(hc && hc.topVarianceEmployees[0]) actions.push(`Confirm HC variance driver for ${hc.topVarianceEmployees[0].name} and whether it reflects start timing, transfer, or comp variance.`);
while(actions.length < 6) actions.push('Review any remaining variance outside the top drivers and confirm whether follow-up is needed before close.');
return actions.slice(0,6);
}
function renderDashboard(model){
document.getElementById('upload-shell').classList.add('hidden');
REVIEW_MONTH_IDX = model.quarter.monthLabels.findIndex(m => normalizeMonth(m) === model.meta.monthToken);
REVIEW_Q_IDX = quarterNum(model.meta.currentQuarterLabel);
const html = `
<style>
#dashboard-root th.tot-col{ background:#e0e7ff !important; color:#312e81 !important; }
#dashboard-root td.tot-col{ background:#eef2ff !important; }
#dashboard-root th.tot-end, #dashboard-root td.tot-end{ border-right:2px solid #a5b4fc !important; }
#dashboard-root th.tot-col:first-of-type, #dashboard-root td.tot-col:first-of-type{ border-left:2px solid #a5b4fc !important; }
#dashboard-root th.rev-col{ background:#e0e7ff !important; color:#312e81 !important; }
#dashboard-root td.rev-col{ background:#eef2ff !important; }
#dashboard-root th.rev-start, #dashboard-root td.rev-start{ border-left:2px solid #a5b4fc !important; }
#dashboard-root th.rev-end, #dashboard-root td.rev-end{ border-right:2px solid #a5b4fc !important; }
#dashboard-root textarea.vedit,
#dashboard-root textarea.hc-edit,
#dashboard-root textarea.act-text,
#dashboard-root .add-comment-input{ font-size:13.5px !important; line-height:1.5 !important; color:#0f172a !important; }
#dashboard-root .vrow p,
#dashboard-root .hcrow p{ font-size:13px !important; }
</style>
${renderOverview(model)}
${renderVarianceSection('sec-qplan', `${model.meta.currentQuarterLabel} vs ${model.meta.fyToken} Plan`, model.quarter.topPlan, model.quarter.driverBlocksPlan, model.quarter.l2Rows, 'p', model.quarter.monthLabels, model.quarter.expense.total.w, model.quarter.expense.total.p, `${model.meta.fyToken} Plan`, model.quarter.expense)}
${renderVarianceSection('sec-qfcst', `${model.meta.currentQuarterLabel} vs ${model.meta.forecastLabel}`, model.quarter.topFcst, model.quarter.driverBlocksFcst, model.quarter.l2Rows, 'f', model.quarter.monthLabels, model.quarter.expense.total.w, model.quarter.expense.total.f, model.meta.forecastLabel, model.quarter.expense)}
${renderYearSection('sec-fyplan', `Full Year vs ${model.meta.fyToken} Plan`, model.year.topPlan, model.year.driverBlocksPlan, model.year.l2Rows, 'p', model.year.expense.total.w, model.year.expense.total.p, `${model.meta.fyToken} Plan`, model.year.expense)}
${renderYearSection('sec-fyfcst', `Full Year vs ${model.meta.forecastLabel}`, model.year.topFcst, model.year.driverBlocksFcst, model.year.l2Rows, 'f', model.year.expense.total.w, model.year.expense.total.f, model.meta.forecastLabel, model.year.expense)}
${renderTE(model)}
${renderHC(model)}
${renderActions(model)}
`;
dom.root.innerHTML = html;
bindInteractive(model);
renderCharts(model);
}
function renderOverview(model){
const k = model.quarter.kpis;
const y = model.year.kpis;
return `<section class="sec" id="sec-overview">
<div class="sec-hdr"><div class="sec-hdr-left"><span class="sec-ic"><i class="ti ti-layout-dashboard"></i></span><span class="sec-title">Performance Overview</span></div><span class="sec-tag">Prepared by ${escapeHtml(model.meta.preparedBy)} · ${escapeHtml(model.meta.badgeLabel)}</span></div>
<div class="kpi-grid">
${kpiCard(model.quarter.currentMonthLabel, k.kpi1, 'Current Month vs Plan')}
${kpiCard(model.quarter.currentMonthLabel, k.kpi2, `Current Month vs ${model.meta.forecastLabel}`)}
${kpiCard(model.meta.currentQuarterLabel, k.kpi3, 'Quarter vs Plan')}
${kpiCard(model.meta.currentQuarterLabel, k.kpi4, `Quarter vs ${model.meta.forecastLabel}`)}
${kpiCard(model.meta.fyToken, y.kpi5, 'Full Year vs Plan')}
</div>
<div class="sec-hdr"><span>Trend</span><span class="sec-tag">Working vs Plan · elapsed fiscal months</span></div>
<div class="chart-legend"><div class="legend-key"><span class="legend-swatch work"></span> Working</div><div class="legend-key"><span class="legend-swatch plan dashed"></span> Plan</div></div>
<div class="trend-wrap"><canvas id="trendC"></canvas></div>
</section>`;
}
function kpiCard(label, val, sub){
const cls = val < 0 ? 'kpi-fav' : val > 0 ? 'kpi-unfav' : 'kpi-neu';
const chip = val < 0 ? 'fav' : val > 0 ? 'unfav' : 'plan';
let icon = 'ti-chart-bar';
if(/month/i.test(sub)) icon = 'ti-calendar-dollar';
else if(/full year/i.test(sub)) icon = 'ti-calendar-stats';
else if(/quarter/i.test(sub)) icon = 'ti-chart-bar';
return `<div class="kpi-card"><div class="kpi-ic ${chip}"><i class="ti ${icon}"></i></div><div class="kpi-label">${escapeHtml(label)}</div><div class="kpi-val ${cls}">${fmtK(val)}</div><div class="kpi-sub">${escapeHtml(sub)}</div>${varianceBadge(val)}</div>`;
}
function renderVarianceSection(id, title, topBundle, blocks, rows, benchmarkKey, monthLabels, workingEnd, baseStart, benchmarkLabel, totalRow){
const secIcon = id.indexOf('fcst') !== -1 ? 'ti-chart-dots-3' : 'ti-trending-down';
const labels = [benchmarkLabel, ...topBundle.rows.map(r => cleanLabel(r.label)), 'Others', 'Working'];
const vars = [...topBundle.rows.map(r=>r.variance), topBundle.others];
return `<section class="sec" id="${id}">
<div class="sec-hdr"><div class="sec-hdr-left"><span class="sec-ic"><i class="ti ${secIcon}"></i></span><span class="sec-title">${escapeHtml(title)}</span></div><span class="sec-tag">${escapeHtml(benchmarkLabel)}</span></div>
<div class="sublbl">Waterfall</div>
<div class="wf-wrap"><canvas id="${id}-wf"></canvas></div>
<div class="sublbl">L2 detail</div>
<div class="tbl-wrap">${renderQuarterTable(rows, benchmarkKey, monthLabels, totalRow, id)}</div>
<div class="sublbl">Drivers</div>
${blocks.map(b => renderDriverBlock(b)).join('') || '<div class="comment-block">No driver block crossed the materiality threshold for this section.</div>'}
<div class="sublbl">Additional Comments</div>
<div class="comment-block" data-comment-block="${id}-comments">${renderAdditionalComments(`${id}-comments`)}</div>
<script type="application/json" id="${id}-wf-data">${JSON.stringify({labels, baseStart, vars, workingEnd})}</script>
</section>`;
}
function renderYearSection(id, title, topBundle, blocks, rows, benchmarkKey, workingEnd, baseStart, benchmarkLabel, totalRow){
const secIcon = id.indexOf('fcst') !== -1 ? 'ti-chart-line' : 'ti-calendar-stats';
const labels = [benchmarkLabel, ...topBundle.rows.map(r => cleanLabel(r.label)), 'Others', 'Working'];
const vars = [...topBundle.rows.map(r=>r.variance), topBundle.others];
return `<section class="sec" id="${id}">
<div class="sec-hdr"><div class="sec-hdr-left"><span class="sec-ic"><i class="ti ${secIcon}"></i></span><span class="sec-title">${escapeHtml(title)}</span></div><span class="sec-tag">${escapeHtml(benchmarkLabel)}</span></div>
<div class="sublbl">Waterfall</div>
<div class="wf-wrap"><canvas id="${id}-wf"></canvas></div>
<div class="sublbl">L2 detail</div>
<div class="tbl-outer"><div class="tbl-inner">${renderYearTable(rows, benchmarkKey, totalRow, id)}</div></div>
<div class="sublbl">Drivers</div>
${blocks.map(b => renderDriverBlock(b)).join('') || '<div class="comment-block">No driver block crossed the materiality threshold for this section.</div>'}
<div class="sublbl">Additional Comments</div>
<div class="comment-block" data-comment-block="${id}-comments">${renderAdditionalComments(`${id}-comments`)}</div>
<script type="application/json" id="${id}-wf-data">${JSON.stringify({labels, baseStart, vars, workingEnd})}</script>
</section>`;
}
function renderQuarterTable(rows, benchmarkKey, monthLabels, totalRow, sectionId){
const drill = sectionId === 'sec-qplan' || sectionId === 'sec-qfcst';
let h = '<table><thead><tr><th>Category</th><th class="yw tot-col">Q Working</th><th class="tot-col">' + (benchmarkKey==='p'?'Q Plan':'Q FCST') + '</th><th class="yv tot-col tot-end">Q Var</th>';
monthLabels.forEach((m, idx) => { const rv = idx===REVIEW_MONTH_IDX; h += `<th class="yw${rv?' rev-col rev-start':''}">${escapeHtml(m)} W</th><th class="${rv?'rev-col':''}">${escapeHtml(m)} ${benchmarkKey==='p'?'P':'F'}</th><th class="mv${rv?' rev-col rev-end':''}">Var</th>`; });
h += '</tr></thead><tbody>';
const totalRows = totalRow ? [{ ...totalRow, label: totalRow.label || 'Expense', rowType: 'expense' }] : [];
const orderedRows = rows.filter(r => r.rowType !== 'expense').concat(totalRows);
orderedRows.forEach(r => {
const qVar = r.total.w - r.total[benchmarkKey];
const trCls = r.rowType === 'expense' ? 'tr-exp' : '';
const dc = drill ? ' drill-cell' : '';
const da = (period, periodLabel) => drill ? ` data-drill="1" data-scope="q" data-row-index="${r.index}" data-period="${period}" data-benchmark="${benchmarkKey}" data-label="${escapeHtml(cleanLabel(r.label))}" data-period-label="${escapeHtml(periodLabel)}"` : '';
h += `<tr class="${trCls}"><td>${escapeHtml(cleanLabel(r.label))}</td>`+
`<td class="yw tot-col${dc}"${da('q','Quarter total')}>${fmtK(r.total.w)}</td>`+
`<td class="tot-col${dc}"${da('q','Quarter total')}>${fmtK(r.total[benchmarkKey])}</td>`+
`<td class="yv tot-col tot-end ${varClass(qVar)}${dc}"${da('q','Quarter total')}>${fmtK(qVar)}</td>`;
r.months.forEach((m, idx) => {
const mv = m.w - m[benchmarkKey];
const rv = idx===REVIEW_MONTH_IDX;
const pl = m.label || monthLabels[idx] || ('Month ' + (idx+1));
h += `<td class="yw${rv?' rev-col rev-start':''}${dc}"${da(idx,pl)}>${fmtK(m.w)}</td>`+
`<td class="${rv?'rev-col':''}${dc}"${da(idx,pl)}>${fmtK(m[benchmarkKey])}</td>`+
`<td class="mv${rv?' rev-col rev-end':''} ${varClass(mv)}${dc}"${da(idx,pl)}>${fmtK(mv)}</td>`;
});
h += '</tr>';
});
h += '</tbody></table>';
return h;
}
function renderYearTable(rows, benchmarkKey, totalRow, sectionId){
const drill = sectionId === 'sec-fyplan' || sectionId === 'sec-fyfcst';
let h = '<table><thead><tr><th>Category</th><th class="yw tot-col">FY Working</th><th class="tot-col">' + (benchmarkKey==='p'?'FY Plan':'FY FCST') + '</th><th class="yv tot-col tot-end">FY Var</th>';
['Q1','Q2','Q3','Q4'].forEach((q, idx) => { const rv = idx===REVIEW_Q_IDX; h += `<th class="yw${rv?' rev-col rev-start':''}">${q} W</th><th class="${rv?'rev-col':''}">${q} ${benchmarkKey==='p'?'P':'F'}</th><th class="mv${rv?' rev-col rev-end':''}">Var</th>`; });
h += '</tr></thead><tbody>';
const totalRows = totalRow ? [{ ...totalRow, label: totalRow.label || 'Expense', rowType: 'expense' }] : [];
const orderedRows = rows.filter(r => r.rowType !== 'expense').concat(totalRows);
orderedRows.forEach(r => {
const fyVar = r.total.w - r.total[benchmarkKey];
const trCls = r.rowType === 'expense' ? 'tr-exp' : '';
const dc = drill ? ' drill-cell' : '';
const da = (period, periodLabel) => drill ? ` data-drill="1" data-scope="y" data-row-index="${r.index}" data-period="${period}" data-benchmark="${benchmarkKey}" data-label="${escapeHtml(cleanLabel(r.label))}" data-period-label="${escapeHtml(periodLabel)}"` : '';
h += `<tr class="${trCls}"><td>${escapeHtml(cleanLabel(r.label))}</td>`+
`<td class="yw tot-col${dc}"${da('fy','Full year')}>${fmtK(r.total.w)}</td>`+
`<td class="tot-col${dc}"${da('fy','Full year')}>${fmtK(r.total[benchmarkKey])}</td>`+
`<td class="yv tot-col tot-end ${varClass(fyVar)}${dc}"${da('fy','Full year')}>${fmtK(fyVar)}</td>`;
r.quarters.forEach((q, idx) => {
const qVar = q.w - q[benchmarkKey];
const rv = idx===REVIEW_Q_IDX;
const pl = q.label || ('Q' + (idx+1));
h += `<td class="yw${rv?' rev-col rev-start':''}${dc}"${da(idx,pl)}>${fmtK(q.w)}</td>`+
`<td class="${rv?'rev-col':''}${dc}"${da(idx,pl)}>${fmtK(q[benchmarkKey])}</td>`+
`<td class="mv${rv?' rev-col rev-end':''} ${varClass(qVar)}${dc}"${da(idx,pl)}>${fmtK(qVar)}</td>`;
});
h += '</tr>';
});
h += '</tbody></table>';
return h;
}
function renderDriverBlock(block){
return `<div class="drv-block ${block.variance < 0 ? 'fav-block':'unfav-block'}" id="blk-${block.id}-wrap">
<div class="drv-block-inner">
<div class="drv-left">
<button class="del-block" data-del-block="blk-${block.id}"><i class="ti ti-trash"></i></button>
<div class="drv-label">${escapeHtml(cleanLabel(block.label))}</div>
<div class="drv-amt ${block.variance<0?'drv-fav':'drv-unfav'}">${fmtK(block.variance)}</div>
<div class="drv-benchmark">Working vs ${block.benchmarkKey==='p'?'Plan':'FCST'}</div>
<div class="drv-badge-wrap">${varianceBadge(block.variance)}</div>
</div>
<div class="drv-right">
<div class="drv-vendors-hdr">Vendor Drivers</div>
${block.vendors.map((v,i) => renderVendorRow(block.id, i, v)).join('')}
<div class="inline-add">
<input class="add-comment-input" data-add-input="${block.id}" placeholder="Add comment row..." />
<button class="mini-btn" data-add-comment="${block.id}"><i class="ti ti-plus"></i></button>
</div>
<div class="drv-footer">Review the underlying service, timing, and whether the current run-rate should carry through the rest of the period.</div>
</div>
</div>
</div>`;
}
function renderVendorRow(blockId, idx, vendor){
const rowId = `${blockId}-${idx}`;
return `<div class="vrow" id="${rowId}-row">
<span style="width:8px;height:8px;border-radius:50%;background:${vendor.variance<0?'#16a34a':'#e11d48'};flex-shrink:0;margin-top:3px"></span>
<div style="flex:1;min-width:0">
<p style="font-size:11.5px;font-weight:600;color:#0f172a;margin-bottom:3px">${escapeHtml(vendor.name)}</p>
<textarea class="vedit" rows="2" data-persist="comment:${rowId}">${escapeHtml(vendor.comment || '')}</textarea>
</div>
<span style="font-size:12px;font-weight:600;color:${vendor.variance<0?'#16a34a':'#e11d48'};flex-shrink:0;margin-top:2px;min-width:48px;text-align:right">${fmtK(vendor.variance)}</span>
<button class="del-btn" data-del-row="${rowId}"><i class="ti ti-trash"></i></button>
</div>`;
}
function renderAdditionalComments(id){
return [].map(i => `<div class="comment-row" id="${id}-${i}-row"><textarea class="vedit" rows="2" data-persist="comment:${id}-${i}" placeholder="Additional comment..."></textarea><button class="del-btn" data-del-row="${id}-${i}"><i class="ti ti-trash"></i></button></div>`).join('') +
`<div class="inline-add"><input class="add-comment-input" data-add-input="${id}" placeholder="Add comment row..." /><button class="mini-btn" data-add-comment="${id}"><i class="ti ti-plus"></i></button></div>`;
}
function renderNoActivity(id, icon, title, tag){
return `<section class="sec" id="${id}"><div class="sec-hdr"><div class="sec-hdr-left"><span class="sec-ic"><i class="ti ${icon}"></i></span><span class="sec-title">${escapeHtml(title)}</span></div><span class="sec-tag">${escapeHtml(tag)}</span></div><div class="comment-block" style="text-align:center;padding:30px 16px;color:#64748b;font-size:14px;font-weight:600"><i class="ti ti-info-circle" style="font-size:20px;display:block;margin-bottom:8px;color:#94a3b8"></i>No activity registered for this period.</div></section>`;
}
function renderTE(model){
if(!model.te) return renderNoActivity('sec-te','ti-plane','Travel & Expense','No activity');
const headers = model.te.headers;
const firstMonthIdx = 2;
const lastMonthIdx = headers.length - 2;
let h = `<section class="sec" id="sec-te"><div class="sec-hdr"><div class="sec-hdr-left"><span class="sec-ic"><i class="ti ti-plane"></i></span><span class="sec-title">Travel & Expense</span></div><span class="sec-tag">Employee and vendor detail</span></div><div class="tbl-outer"><table><thead><tr>`;
headers.forEach((hdr, idx) => h += `<th class="${idx===firstMonthIdx?'te-q1':''} ${idx===lastMonthIdx?'te-q2':''} ${idx>=1?'te-center':''}">${escapeHtml(hdr)}</th>`);
h += '</tr></thead><tbody>';
model.te.rows.filter(row => Number.isFinite(row.grandTotal) && row.grandTotal !== 0).forEach(row => {
h += `<tr><td>${escapeHtml(row.employee)}</td><td class="te-center">${escapeHtml(row.vendor)}</td>`;
row.values.forEach((n, idx) => h += `<td class="te-center ${idx===0?'te-q1':''} ${idx===row.values.length-1?'te-q2':''}">${fmtKplain(n)}</td>`);
h += `<td class="te-center">${fmtKplain(row.grandTotal)}</td></tr>`;
});
h += `<tr class="tr-exp"><td>${escapeHtml(model.te.totalRow.employee)}</td><td></td>`;
model.te.totalRow.values.forEach((n, idx) => h += `<td class="te-center ${idx===0?'te-qt1':''} ${idx===model.te.totalRow.values.length-1?'te-qt2':''}">${fmtKplain(n)}</td>`);
h += `<td class="te-center te-grand">${fmtKplain(model.te.totalRow.grandTotal)}</td></tr></tbody></table></div></section>`;
return h;
}
function renderHC(model){
if(!model.hc) return renderNoActivity('sec-hc','ti-users','Headcount Cost','No activity');
const m = model.hc.movementCounts;
const variance = model.hc.salaryAccrued.workTotal - model.hc.salaryAccrued.planTotal;
return `<section class="sec" id="sec-hc">
<div class="sec-hdr"><div class="sec-hdr-left"><span class="sec-ic"><i class="ti ti-users"></i></span><span class="sec-title">Headcount Cost</span></div><span class="sec-tag">Salary Accrued + movement</span></div>
<div class="hc-grid">
<div class="hc-card hc-card-wide hc-work">
<div class="hc-summary-grid">
<div class="hc-stat"><div class="kpi-label">Plan Total</div><div class="kpi-val kpi-neu">${fmtK(model.hc.salaryAccrued.planTotal)}</div></div>
<div class="hc-stat"><div class="kpi-label">Working Total</div><div class="kpi-val ${variance<0?'kpi-fav':'kpi-unfav'}">${fmtK(model.hc.salaryAccrued.workTotal)}</div></div>
<div class="hc-stat"><div class="kpi-label">Variance</div><div class="kpi-val ${variance<0?'kpi-fav':variance>0?'kpi-unfav':'kpi-neu'}">${fmtK(variance)}</div><div class="hc-badge-wrap">${varianceBadge(variance)}</div></div>
</div>
<div class="mini-kpi hc-quarter-grid">${model.hc.salaryAccrued.q.map((n,i)=>`<div><div class="k">Q${i+1}</div><div class="v">Plan ${fmtK(n)}</div><div class="v2">Working ${fmtK(model.hc.salaryAccrued.qWork[i])}</div></div>`).join('')}</div>
</div>
</div>
<div class="sublbl">HC movement</div>
<div class="tbl-wrap hc-move-wrap"><table class="hc-move-table"><thead><tr><th>Metric</th><th class="yw">Plan</th><th class="te-q2">Working</th><th class="yv">Var</th></tr></thead><tbody>
${hcRow('Active Employees', m.activePlan, m.activeWork)}
${hcRow('TBH Roles', m.tbhPlan, m.tbhWork)}
${hcRow('New Hires (not in plan)', 0, m.newHires)}
${hcRow('Ending HC', m.endingPlan, m.endingWork, true)}
</tbody></table></div>
<div class="sublbl">Employee commentary</div>
<div>${model.hc.topVarianceEmployees.map((r,i)=>`<div class="hcrow" id="hc-${i}-row"><span style="width:8px;height:8px;border-radius:50%;background:${r.variance<0?'#16a34a':'#e11d48'};flex-shrink:0;margin-top:3px"></span><div style="flex:1"><p style="font-size:11.5px;font-weight:600;color:#0f172a;margin-bottom:3px">${escapeHtml(r.name)}</p><textarea class="hc-edit" rows="2" data-persist="comment:hc-${i}">${escapeHtml(r.comment)}</textarea></div><span style="font-size:12px;font-weight:600;color:${r.variance<0?'#16a34a':'#e11d48'};min-width:48px;text-align:right">${fmtK(r.variance)}</span><button class="del-btn" data-del-row="hc-${i}"><i class="ti ti-trash"></i></button></div>`).join('')}</div>
</section>`;
}
function hcRow(label, plan, work, highlight){
const variance = work - plan;
return `<tr class="${highlight?'tr-exp':''}"><td>${escapeHtml(label)}</td><td class="yw">${plan}</td><td class="te-q2">${work}</td><td class="yv ${varClass(variance)}">${variance}</td></tr>`;
}
function renderActions(model){
return `<section class="sec" id="sec-actions"><div class="sec-hdr"><div class="sec-hdr-left"><span class="sec-ic"><i class="ti ti-checklist"></i></span><span class="sec-title">Follow-up Actions</span></div><span id="actCtr" class="sec-tag">0 of 6 completed</span></div><div class="act-box" id="actList">${model.actions.map((a,i)=>`<div class="act-item" id="act-${i}"><input type="checkbox" class="act-cb" data-act-cb="act-${i}"><textarea class="act-text" rows="1" data-persist="action:act-${i}">${escapeHtml(a)}</textarea><button class="del-btn" data-del-row="act-${i}"><i class="ti ti-trash"></i></button></div>`).join('')}</div><button class="add-act" id="add-act"><i class="ti ti-plus"></i> Add action</button></section>`;
}
function bindInteractive(model){
document.querySelectorAll('[data-del-row]').forEach(btn => btn.addEventListener('click', () => { const id = btn.dataset.delRow; const row = document.getElementById(id + '-row') || document.getElementById(id); if(row) row.classList.add('hidden'); saveState(); }));
document.querySelectorAll('[data-del-block]').forEach(btn => btn.addEventListener('click', () => { const id = btn.dataset.delBlock; const row = document.getElementById(id + '-wrap') || document.getElementById(id + '-wrap'.replace('blk-','')); if(row) row.classList.add('hidden'); saveState(); }));
document.querySelectorAll('[data-add-comment]').forEach(btn => btn.addEventListener('click', () => addCommentRow(btn.dataset.addComment)));
document.getElementById('add-act').addEventListener('click', addAction);
document.querySelectorAll('[data-act-cb]').forEach(cb => cb.addEventListener('change', saveState));
document.querySelectorAll('[data-persist]').forEach(el => el.addEventListener('input', () => { autoResize(el); saveState(); }));
if(!dom.root.dataset.drillBound){
dom.root.addEventListener('click', e => {
const cell = e.target.closest('[data-drill="1"]');
if(cell) openDrill(cell);
});
dom.root.dataset.drillBound = '1';
}
updateActionCounter();
document.querySelectorAll('textarea').forEach(autoResize);
window.addEventListener('scroll', scrollSpy, { passive:true });
}
function addCommentRow(blockId){
const input = document.querySelector(`[data-add-input="${blockId}"]`);
const text = input ? input.value.trim() : '';
const rowId = slug(blockId + '-' + Date.now());
const row = document.createElement('div');
row.className = 'comment-row';
row.id = rowId + '-row';
row.innerHTML = `<textarea class="vedit" rows="2" data-persist="comment:${rowId}">${escapeHtml(text)}</textarea><button class="del-btn" data-del-row="${rowId}"><i class="ti ti-trash"></i></button>`;
const addWrap = input.closest('.inline-add');
addWrap.parentNode.insertBefore(row, addWrap);
row.querySelector('[data-del-row]').addEventListener('click', () => { row.classList.add('hidden'); saveState(); });
row.querySelector('[data-persist]').addEventListener('input', e => { autoResize(e.target); saveState(); });
autoResize(row.querySelector('textarea'));
if(input) input.value = '';
saveState();
}
function addAction(){
const list = document.getElementById('actList');
const id = 'act-' + Date.now();
const div = document.createElement('div');
div.className = 'act-item';
div.id = id;
div.innerHTML = `<input type="checkbox" class="act-cb" data-act-cb="${id}"><textarea class="act-text" rows="1" data-persist="action:${id}" placeholder="New action item..."></textarea><button class="del-btn" data-del-row="${id}"><i class="ti ti-trash"></i></button>`;
list.appendChild(div);
div.querySelector('[data-act-cb]').addEventListener('change', saveState);
div.querySelector('[data-del-row]').addEventListener('click', () => { div.classList.add('hidden'); saveState(); updateActionCounter(); });
div.querySelector('[data-persist]').addEventListener('input', e => { autoResize(e.target); saveState(); });
autoResize(div.querySelector('textarea'));
updateActionCounter();
saveState();
}
function updateActionCounter(){
const items = Array.from(document.querySelectorAll('.act-item')).filter(el => !el.classList.contains('hidden'));
const done = items.filter(el => el.querySelector('.act-cb').checked).length;
const ctr = document.getElementById('actCtr');
if(ctr) ctr.textContent = `${done} of ${items.length} completed`;
}
function autoResize(el){ if(!el || el.tagName !== 'TEXTAREA') return; el.style.height='auto'; el.style.height = el.scrollHeight + 'px'; }
function saveState(){
if(!state.model) return;
const key = storageKey();
const payload = {
comments: {},
actions: {},
hidden: []
};
document.querySelectorAll('[data-persist]').forEach(el => payload.comments[el.dataset.persist] = el.value);
document.querySelectorAll('[data-act-cb]').forEach(cb => payload.actions[cb.dataset.actCb] = cb.checked);
document.querySelectorAll('.hidden[id]').forEach(el => payload.hidden.push(el.id));
localStorage.setItem(key, JSON.stringify(payload));
updateActionCounter();
toast('Board saved');
}
function restoreSavedState(){
const raw = localStorage.getItem(storageKey());
if(!raw) return;
try{
const payload = JSON.parse(raw);
Object.entries(payload.comments || {}).forEach(([k,v]) => {
const el = document.querySelector(`[data-persist="${cssEscape(k)}"]`);
if(el){ el.value = v; autoResize(el); }
});
Object.entries(payload.actions || {}).forEach(([k,v]) => {
const cb = document.querySelector(`[data-act-cb="${cssEscape(k)}"]`);
if(cb) cb.checked = !!v;
});
(payload.hidden || []).forEach(id => {
const el = document.getElementById(id);
if(el) el.classList.add('hidden');
});
updateActionCounter();
}catch(err){ console.warn(err); }
}
function storageKey(){
return 'bva:' + slug(state.model.meta.badgeLabel);
}
function exportFileBase(){
return `bva_${slug(state.model.meta.dashboardCode)}_${slug(state.model.meta.fyToken)}_${slug(state.model.meta.monthToken)}`;
}
async function fetchInlineCss(){
try{
const link = document.querySelector('link[href$="styles.css"]');
const href = link ? link.href : './styles.css';
const res = await fetch(href);
if(res.ok) return await res.text();
}catch(e){ console.warn('Could not fetch styles.css directly, falling back to loaded stylesheet rules.', e); }
try{
const sheet = Array.from(document.styleSheets).find(s => (s.href||'').indexOf('styles.css') !== -1);
if(sheet) return Array.from(sheet.cssRules).map(r => r.cssText).join('\n');
}catch(e){ console.warn('Could not read stylesheet rules for inline export.', e); }
return '';
}
function inlineCssIntoClone(clone, cssText){
if(!cssText) return;
const styleTag = document.createElement('style');
styleTag.textContent = cssText;
const link = clone.querySelector('link[href$="styles.css"]');
if(link) link.replaceWith(styleTag);
else { const head = clone.querySelector('head'); if(head) head.appendChild(styleTag); }
}
function freezeCanvasesAsImages(clone){
document.querySelectorAll('canvas[id]').forEach(liveCanvas => {
let dataUrl;
try{
dataUrl = liveCanvas.toDataURL('image/png');
}catch(e){
console.warn('Could not snapshot canvas ' + liveCanvas.id, e);
return;
}
const cloneCanvas = clone.querySelector('#' + liveCanvas.id);
if(!cloneCanvas) return;
const img = document.createElement('img');
img.src = dataUrl;
img.alt = liveCanvas.id;
img.style.display = 'block';
img.style.width = '100%';
img.style.height = '100%';
img.style.maxWidth = '100%';
img.style.maxHeight = '100%';
img.style.objectFit = 'contain';
img.style.objectPosition = 'center';
cloneCanvas.replaceWith(img);
});
}
function stripLiveScripts(clone, keepPatterns){
const keep = keepPatterns || [];
clone.querySelectorAll('script[src]').forEach(s => {
const src = s.getAttribute('src') || '';
const shouldKeep = keep.some(p => src.indexOf(p) !== -1);
if(!shouldKeep) s.remove();
});
}
function buildExportScript(){
return `(function(){
function autoResize(el){ if(!el) return; el.style.height='auto'; el.style.height = el.scrollHeight + 'px'; }
document.querySelectorAll('textarea').forEach(autoResize);
document.querySelectorAll('textarea').forEach(function(el){ el.addEventListener('input', function(){ autoResize(el); }); });
function updateActionCounter(){
var items = Array.prototype.slice.call(document.querySelectorAll('.act-item')).filter(function(el){ return !el.classList.contains('hidden'); });
var done = items.filter(function(el){ var cb = el.querySelector('.act-cb'); return cb && cb.checked; }).length;
var ctr = document.getElementById('actCtr');
if(ctr) ctr.textContent = done + ' of ' + items.length + ' completed';
}
document.querySelectorAll('[data-act-cb]').forEach(function(cb){ cb.addEventListener('change', updateActionCounter); });
updateActionCounter();
function navTo(sectionId, el){
var target = document.getElementById(sectionId);
if(!target) return;
var top = target.getBoundingClientRect().top + window.scrollY - 20;
window.scrollTo({ top: top, behavior: 'smooth' });
document.querySelectorAll('aside nav ul li[data-nav]').forEach(function(li){ li.classList.remove('active'); });
if(el) el.classList.add('active');
}
document.querySelectorAll('aside nav ul li[data-nav]').forEach(function(li){
li.addEventListener('click', function(){ navTo(li.dataset.nav, li); });
});
function scrollSpy(){
var sections = ['sec-overview','sec-qplan','sec-qfcst','sec-fyplan','sec-fyfcst','sec-te','sec-hc','sec-actions'];
var navItems = document.querySelectorAll('aside nav ul li[data-nav]');
var scrollY = window.scrollY + 80;
var current = 0;
sections.forEach(function(id, i){ var el = document.getElementById(id); if(el && el.offsetTop <= scrollY) current = i; });
navItems.forEach(function(li){ li.classList.remove('active'); });
if(navItems[current]) navItems[current].classList.add('active');
}
window.addEventListener('scroll', scrollSpy, { passive: true });
})();`;
}
function appendExportScript(clone){
const script = document.createElement('script');
script.textContent = buildExportScript();
const body = clone.querySelector('body');
if(body) body.appendChild(script);
}
// Build a fully self-contained HTML string for the currently rendered board.
// opts.interactive === true keeps Save / Download / add-line / trash controls
// live inside the exported window (used for multi-board windows).
async function buildStandaloneHtml(opts){
const interactive = !!(opts && opts.interactive);
const cssText = await fetchInlineCss();
// Persist current field values into the DOM so the clone captures typed text.
document.querySelectorAll('#dashboard-root textarea').forEach(t => { t.textContent = t.value; });
document.querySelectorAll('#dashboard-root input').forEach(i => {
if(i.type === 'checkbox'){ if(i.checked) i.setAttribute('checked','checked'); else i.removeAttribute('checked'); }
else { i.setAttribute('value', i.value); }
});
const clone = document.documentElement.cloneNode(true);
if(interactive){
// Keep interactive controls; only remove things that make no sense in a
// standalone window (workspace navigation, uploader, batch UI, open drill).
clone.querySelectorAll('#back-nav, #home-nav, #budget-util-nav, .topbar-actions, #upload-shell, #drill-overlay').forEach(el => el.remove());
freezeCanvasesAsImages(clone);
stripLiveScripts(clone, ['html2canvas', 'jspdf']); // keep PDF libs for in-window export
inlineCssIntoClone(clone, cssText);
appendInteractiveScript(clone);
appendDrillExportScript(clone);
} else {
clone.querySelectorAll('.hidden, .del-btn, .del-block, .add-act, .mini-btn, .ghost-btn, .topbar-actions, #save-nav, #download-nav, #download-pdf-nav, #back-nav, #home-nav, #drill-overlay, #upload-shell').forEach(el => el.remove());
freezeCanvasesAsImages(clone);
stripLiveScripts(clone);
inlineCssIntoClone(clone, cssText);
appendExportScript(clone);
appendDrillExportScript(clone);
}
return '<!DOCTYPE html>\n' + clone.outerHTML;
}
// Inject the self-contained interactive bootstrap (Save / Download HTML /
// Download PDF / add comment / add action / delete row+block / nav+scrollspy).
function appendInteractiveScript(clone){
const script = document.createElement('script');
script.textContent = buildInteractiveBootstrap(storageKey(), exportFileBase());
const body = clone.querySelector('body');
if(body) body.appendChild(script);
}
// Returns a vanilla, dependency-free IIFE string that re-wires every editable
// control inside an exported window. STORAGE_KEY / FILE_BASE are baked in so
// each board window persists to its own localStorage slot.
function buildInteractiveBootstrap(storageKeyStr, fileBaseStr){
const KEY = JSON.stringify(storageKeyStr);
const FILE = JSON.stringify(fileBaseStr);
return '(function(){\n'
+ 'var STORAGE_KEY=' + KEY + ',FILE_BASE=' + FILE + ';\n'
+ 'function slug(s){return String(s||"").toLowerCase().replace(/[^a-z0-9]+/g,"-").replace(/(^-|-$)/g,"");}\n'
+ 'function escapeHtml(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}\n'
+ 'function cssEscape(s){return String(s).replace(/"/g,\'\\"\');}\n'
+ 'function autoResize(el){if(!el||el.tagName!=="TEXTAREA")return;el.style.height="auto";el.style.height=el.scrollHeight+"px";}\n'
+ 'function toast(m){var t=document.getElementById("toast");if(!t){t=document.createElement("div");t.id="toast";t.className="toast";document.body.appendChild(t);}t.textContent=m;t.classList.add("show");setTimeout(function(){t.classList.remove("show");},1800);}\n'
+ 'function updateActionCounter(){var items=Array.prototype.slice.call(document.querySelectorAll(".act-item")).filter(function(el){return !el.classList.contains("hidden");});var done=items.filter(function(el){var cb=el.querySelector(".act-cb");return cb&&cb.checked;}).length;var ctr=document.getElementById("actCtr");if(ctr)ctr.textContent=done+" of "+items.length+" completed";}\n'
+ 'function saveState(silent){var p={comments:{},actions:{},hidden:[]};document.querySelectorAll("[data-persist]").forEach(function(el){p.comments[el.dataset.persist]=el.value;});document.querySelectorAll("[data-act-cb]").forEach(function(cb){p.actions[cb.dataset.actCb]=cb.checked;});document.querySelectorAll(".hidden[id]").forEach(function(el){p.hidden.push(el.id);});try{localStorage.setItem(STORAGE_KEY,JSON.stringify(p));}catch(e){}updateActionCounter();if(!silent)toast("Board saved");}\n'
+ 'function restoreSavedState(){var raw;try{raw=localStorage.getItem(STORAGE_KEY);}catch(e){}if(!raw)return;try{var p=JSON.parse(raw);Object.keys(p.comments||{}).forEach(function(k){var el=document.querySelector("[data-persist=\\""+cssEscape(k)+"\\"]");if(el){el.value=p.comments[k];autoResize(el);}});Object.keys(p.actions||{}).forEach(function(k){var cb=document.querySelector("[data-act-cb=\\""+cssEscape(k)+"\\"]");if(cb)cb.checked=!!p.actions[k];});(p.hidden||[]).forEach(function(id){var el=document.getElementById(id);if(el)el.classList.add("hidden");});updateActionCounter();}catch(e){}}\n'
+ 'function addCommentRow(blockId){var input=document.querySelector("[data-add-input=\\""+blockId+"\\"]");var text=input?input.value.trim():"";var rowId=slug(blockId+"-"+Date.now());var row=document.createElement("div");row.className="comment-row";row.id=rowId+"-row";row.innerHTML="<textarea class=\\"vedit\\" rows=\\"2\\" data-persist=\\"comment:"+rowId+"\\">"+escapeHtml(text)+"</textarea><button class=\\"del-btn\\" data-del-row=\\""+rowId+"\\"><i class=\\"ti ti-trash\\"></i></button>";var addWrap=input.closest(".inline-add");addWrap.parentNode.insertBefore(row,addWrap);var ta=row.querySelector("textarea");ta.addEventListener("input",function(){autoResize(ta);saveState(true);});autoResize(ta);if(input)input.value="";saveState(true);}\n'
+ 'function addAction(){var list=document.getElementById("actList");if(!list)return;var id="act-"+Date.now();var div=document.createElement("div");div.className="act-item";div.id=id;div.innerHTML="<input type=\\"checkbox\\" class=\\"act-cb\\" data-act-cb=\\""+id+"\\"><textarea class=\\"act-text\\" rows=\\"1\\" data-persist=\\"action:"+id+"\\" placeholder=\\"New action item...\\"></textarea><button class=\\"del-btn\\" data-del-row=\\""+id+"\\"><i class=\\"ti ti-trash\\"></i></button>";list.appendChild(div);div.querySelector("[data-act-cb]").addEventListener("change",function(){saveState(true);});var ta=div.querySelector("textarea");ta.addEventListener("input",function(){autoResize(ta);saveState(true);});autoResize(ta);updateActionCounter();saveState(true);}\n'
+ 'function downloadHtml(){try{document.querySelectorAll("textarea").forEach(function(t){t.textContent=t.value;});document.querySelectorAll("input").forEach(function(i){if(i.type==="checkbox"){if(i.checked)i.setAttribute("checked","checked");else i.removeAttribute("checked");}else{i.setAttribute("value",i.value);}});var clone=document.documentElement.cloneNode(true);clone.querySelectorAll("#toast").forEach(function(el){el.remove();});var html="<!DOCTYPE html>\\n"+clone.outerHTML;var blob=new Blob([html],{type:"text/html"});var a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=FILE_BASE+".html";a.click();URL.revokeObjectURL(a.href);toast("HTML downloaded");}catch(e){console.error(e);toast("Could not export HTML");}}\n'
+ 'function downloadPdf(){if(typeof html2canvas==="undefined"||!window.jspdf){toast("PDF libraries did not load");return;}saveState(true);var target=document.querySelector(".main-workspace");if(!target){toast("Could not find dashboard content");return;}var hiddenEls=Array.prototype.slice.call(target.querySelectorAll(".hidden, .del-btn, .del-block, .add-act, .mini-btn, .ghost-btn, .topbar-actions, #back-nav, #home-nav"));var restore=hiddenEls.map(function(el){return [el,el.style.display];});var aside=document.querySelector("aside");var asideDisplay=aside?aside.style.display:null;var prevML=target.style.marginLeft,prevW=target.style.width;toast("Building PDF...");hiddenEls.forEach(function(el){el.style.display="none";});if(aside)aside.style.display="none";target.style.marginLeft="0";target.style.width="100%";setTimeout(function(){html2canvas(target,{scale:2,useCORS:true,backgroundColor:"#f8fafc"}).then(function(canvas){var jsPDF=window.jspdf.jsPDF;var pdf=new jsPDF("p","pt","a4");var pageWidth=pdf.internal.pageSize.getWidth();var pageHeight=pdf.internal.pageSize.getHeight();var ratio=canvas.width/pageWidth;var pageHeightPx=Math.max(1,Math.floor(pageHeight*ratio));var rendered=0,first=true;while(rendered<canvas.height){var sh=Math.min(pageHeightPx,canvas.height-rendered);var sc=document.createElement("canvas");sc.width=canvas.width;sc.height=sh;sc.getContext("2d").drawImage(canvas,0,rendered,canvas.width,sh,0,0,canvas.width,sh);var img=sc.toDataURL("image/jpeg",0.92);if(!first)pdf.addPage();pdf.addImage(img,"JPEG",0,0,pageWidth,sh/ratio);rendered+=sh;first=false;}pdf.save(FILE_BASE+".pdf");toast("PDF downloaded");}).catch(function(e){console.error(e);toast("Could not export PDF");}).then(function(){restore.forEach(function(pr){pr[0].style.display=pr[1];});if(aside)aside.style.display=asideDisplay;target.style.marginLeft=prevML;target.style.width=prevW;});},60);}\n'
+ 'function navTo(sectionId,el){var t=document.getElementById(sectionId);if(!t)return;var top=t.getBoundingClientRect().top+window.scrollY-20;window.scrollTo({top:top,behavior:"smooth"});document.querySelectorAll("aside nav ul li[data-nav]").forEach(function(li){li.classList.remove("active");});if(el)el.classList.add("active");}\n'
+ 'function scrollSpy(){var sections=["sec-overview","sec-qplan","sec-qfcst","sec-fyplan","sec-fyfcst","sec-te","sec-hc","sec-actions"];var navItems=document.querySelectorAll("aside nav ul li[data-nav]");var sy=window.scrollY+80;var current=0;sections.forEach(function(id,i){var el=document.getElementById(id);if(el&&el.offsetTop<=sy)current=i;});navItems.forEach(function(li){li.classList.remove("active");});if(navItems[current])navItems[current].classList.add("active");}\n'
+ 'document.body.addEventListener("click",function(e){var t=e.target;if(!t.closest)return;var dr=t.closest("[data-del-row]");if(dr){var id=dr.dataset.delRow;var row=document.getElementById(id+"-row")||document.getElementById(id);if(row)row.classList.add("hidden");saveState(true);updateActionCounter();return;}var db=t.closest("[data-del-block]");if(db){var w=document.getElementById(db.dataset.delBlock+"-wrap");if(w)w.classList.add("hidden");saveState(true);return;}var ac=t.closest("[data-add-comment]");if(ac){addCommentRow(ac.dataset.addComment);return;}if(t.closest("#add-act")){addAction();return;}if(t.closest("#save-nav")){saveState();return;}if(t.closest("#download-nav")){downloadHtml();return;}if(t.closest("#download-pdf-nav")){downloadPdf();return;}var nav=t.closest("aside nav ul li[data-nav]");if(nav){navTo(nav.dataset.nav,nav);return;}});\n'
+ 'document.querySelectorAll("[data-persist]").forEach(function(el){el.addEventListener("input",function(){autoResize(el);saveState(true);});});\n'
+ 'document.querySelectorAll("[data-act-cb]").forEach(function(cb){cb.addEventListener("change",function(){saveState(true);});});\n'
+ 'window.addEventListener("scroll",scrollSpy,{passive:true});\n'
+ 'restoreSavedState();updateActionCounter();document.querySelectorAll("textarea").forEach(autoResize);\n'
+ '})();';
}
async function downloadHtml(){
if(!state.model){ toast('Build a dashboard first'); return; }
saveState();
const nav = document.getElementById('download-nav');
try{
if(nav) nav.classList.add('disabled-nav');
const html = await buildStandaloneHtml();
const blob = new Blob([html], { type:'text/html' });
const a = document.createElement('a');
a.href = URL.createObjectURL(blob);
a.download = `${exportFileBase()}.html`;
a.click();
URL.revokeObjectURL(a.href);
toast('HTML downloaded');
}catch(err){
console.error(err);
toast('Could not export HTML');
}finally{
if(nav) nav.classList.remove('disabled-nav');
}
}
// ---------- Multi-board generation ----------
async function onGenerate(){
showError('');
if(dom.launcher){ dom.launcher.classList.add('hidden'); dom.launcher.innerHTML = ''; }
if(state.boardCount === 1){
await buildDashboard(state.boards[0]);
return;
}
await buildAllBoards();
}
// Validate BU+month uniqueness across boards. Returns an error string or null.
function validateBoardKeys(){
const seen = new Map();
for(let i = 0; i < state.boards.length; i++){
const files = state.boards[i];
let meta = null;
try{ meta = parseMeta(files.quarter.name); }catch(e){}
if(!meta){ return `Board ${i+1}: could not read the business unit/month from the Quarter filename.`; }
const key = slug(meta.dashboardCode) + '|' + slug(meta.monthToken) + '|' + slug(meta.fyToken);
if(seen.has(key)){
const other = seen.get(key) + 1;
return `Boards ${other} and ${i+1} are the same board (${meta.dashboardCode} · ${meta.monthToken} ${meta.fyToken}). Use a different business unit, or the same BU with a different month.`;
}
seen.set(key, i);
}
return null;
}
async function buildAllBoards(){
const err = validateBoardKeys();
if(err){ showError(err); toast('Duplicate board detected'); return; }
dom.buildBtn.disabled = true;
const originalLabel = dom.buildBtn.innerHTML;
state.generated = [];
try{
for(let i = 0; i < state.boards.length; i++){
dom.buildBtn.innerHTML = `<i class="ti ti-loader-2"></i> Building ${i+1} / ${state.boards.length}...`;
const parsed = await parseFiles(state.boards[i]);
const model = deriveModel(parsed);
state.model = model;
renderDashboard(model);
restoreSavedState();
document.getElementById('upload-shell').classList.add('hidden');
dom.root.classList.remove('hidden');
// Let charts paint before snapshotting.
await new Promise(r => setTimeout(r, 550));
const html = await buildStandaloneHtml({ interactive:true });
state.generated.push({
title: model.meta.dashboardCode,
subtitle: `${model.meta.monthToken} ${model.meta.fyToken} · ${model.meta.currentQuarterLabel}`,
fileBase: exportFileBase(),
html
});
}
// Return to the upload shell and show launcher buttons.
dom.root.classList.add('hidden');
dom.root.innerHTML = '';
state.model = null;
document.getElementById('upload-shell').classList.remove('hidden');
if(dom.backNav) dom.backNav.classList.add('hidden');
if(dom.homeNav) dom.homeNav.classList.remove('hidden');
renderLauncher();
toast(`${state.generated.length} boards ready`);
}catch(e){
console.error(e);
showError(e.message || 'Could not build one of the boards. Check the feeds and try again.');
}finally{
dom.buildBtn.disabled = false;
updateBuildButton();
}
}
function renderLauncher(){
if(!dom.launcher) return;
const cards = state.generated.map((g, i) => `
<button type="button" class="launch-btn" data-launch="${i}">
<span class="lb-ic"><i class="ti ti-external-link"></i></span>
<span class="lb-body">
<span class="lb-title">Open Board ${i+1}</span>
<span class="lb-sub">${escapeHtml(g.title)} · ${escapeHtml(g.subtitle)}</span>
</span>
</button>
`).join('');
dom.launcher.innerHTML = `
<div class="launcher-hdr">
<i class="ti ti-circle-check"></i>
${state.generated.length} boards generated
</div>
<div class="launcher-actions">
<button type="button" class="launcher-action primary" data-open-package>
<i class="ti ti-layout-tabs"></i>
Open Board Package
</button>
<button type="button" class="launcher-action" data-download-package>
<i class="ti ti-file-zip"></i>
Download ZIP
</button>
</div>
<div class="launch-grid">${cards}</div>
<div class="launcher-note">
Open Board Package lets you edit each board in its own tab.
Download ZIP captures the current comments, actions and changes.
</div>
`;
dom.launcher.classList.remove('hidden');
dom.launcher.querySelectorAll('[data-launch]').forEach(btn => {
btn.addEventListener('click', () => {
openGeneratedBoard(Number(btn.dataset.launch), btn);
});
});
const openPackageBtn = dom.launcher.querySelector('[data-open-package]');
if(openPackageBtn){
openPackageBtn.addEventListener('click', openGeneratedPackage);
}
const downloadPackageBtn = dom.launcher.querySelector('[data-download-package]');
if(downloadPackageBtn){
downloadPackageBtn.addEventListener('click', downloadGeneratedPackage);
}
dom.launcher.scrollIntoView({ behavior:'smooth', block:'nearest' });
}
function openGeneratedBoard(index, btn){
const g = state.generated[index];
if(!g){ toast('Board not found'); return; }
const w = window.open('', '_blank');
if(!w){ toast('Pop-up blocked — allow pop-ups and click again'); return; }
try{
w.document.open();
w.document.write(g.html);
w.document.close();
w.document.title = `${g.title} — ${g.subtitle}`;
if(btn){ btn.classList.add('opened'); }
}catch(e){
console.error(e);
toast('Could not open the board window');
}
}
function packageRecords(){
const seen = new Set();
return state.generated.reduce((out, g, i) => {
const fileName = `${g.fileBase || `bva_board_${i+1}`}.html`;
if(seen.has(fileName)) return out;
seen.add(fileName);
out.push({
title: g.title,
subtitle: g.subtitle,
fileName,
html: g.html
});
return out;
}, []);
}
function openGeneratedPackage(){
const boards = packageRecords();
const w = window.open('', '_blank');
if(!w){
toast('Pop-up blocked — allow pop-ups and try again');
return;
}
w.document.open();
w.document.write(buildPackageHubHtml(boards));
w.document.close();
}
async function downloadGeneratedPackage(){
if(typeof JSZip === 'undefined'){
toast('JSZip did not load');
return;
}
const boards = packageRecords();
const zip = new JSZip();
boards.forEach(board => {
zip.file(board.fileName, board.html);
});
zip.file('index.html', buildPackageHubHtml(boards));
const blob = await zip.generateAsync({ type:'blob' });
const url = URL.createObjectURL(blob);
const link = document.createElement('a');
link.href = url;
link.download = 'bva_boards.zip';
link.click();
setTimeout(() => URL.revokeObjectURL(url), 1000);
toast('ZIP downloaded');
}
function encodeBase64Utf8(value){
const bytes = new TextEncoder().encode(String(value == null ? '' : value));
let binary = '';
for(let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
return btoa(binary);
}

function buildPackageHubHtml(boards){
const boardFixCss = `
<style id="bva-package-board-fix">
html,
body{
width:100%!important;
min-width:0!important;
margin:0!important;
padding:0!important;
overflow-x:hidden!important;
}
body{
display:flex!important;
min-height:100vh!important;
background:#f7f9fd!important;
}
aside{
display:flex!important;
width:244px!important;
position:fixed!important;
left:0!important;
top:0!important;
height:100vh!important;
z-index:200!important;
}
.main-workspace{
display:block!important;
width:calc(100% - 244px)!important;
max-width:none!important;
margin-left:244px!important;
padding:28px 30px 56px!important;
overflow-x:hidden!important;
}
.trend-wrap,
.wf-wrap{
position:relative!important;
width:100%!important;
max-width:100%!important;
overflow:hidden!important;
}
.trend-wrap img,
.wf-wrap img,
.trend-wrap canvas,
.wf-wrap canvas{
display:block!important;
width:100%!important;
max-width:100%!important;
height:100%!important;
max-height:100%!important;
object-fit:contain!important;
object-position:center!important;
}
</style>`;

const packagePayload = {
  boards: boards.map(board => ({
    title: board.title,
    subtitle: board.subtitle,
    fileName: board.fileName,
    htmlBase64: encodeBase64Utf8(board.html)
  }))
};
const packageData = encodeBase64Utf8(JSON.stringify(packagePayload));

return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BvA Board Package</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<style>
*{box-sizing:border-box}
html,
body{
width:100%;
height:100%;
margin:0;
padding:0;
overflow:hidden;
}
body{
font-family:Inter,Arial,sans-serif;
background:#eef2fb;
color:#0f172a;
}
.hub{
width:100%;
height:100vh;
max-width:none;
margin:0;
background:#fff;
border:0;
border-radius:0;
box-shadow:none;
overflow:hidden;
display:flex;
flex-direction:column;
}
.hub-header{
display:flex;
justify-content:space-between;
align-items:center;
gap:16px;
padding:18px 24px;
border-bottom:1px solid #e9eef7;
flex-shrink:0;
}
.hub-title{font-size:22px;font-weight:800}
.hub-sub{margin-top:5px;color:#64748b;font-size:12px}
.hub-actions{display:flex;gap:9px}
.hub-btn{border:1px solid #cbd5e1;border-radius:9px;padding:10px 14px;background:#fff;color:#334155;font-weight:800;cursor:pointer}
.hub-btn.primary{background:#4f46e5;color:#fff;border-color:#4f46e5}
.tabs{display:flex;gap:6px;overflow-x:auto;padding:12px 20px;border-bottom:1px solid #e9eef7;background:#f8fafc;flex-shrink:0}
.tab{border:1px solid #cbd5e1;border-radius:9px;padding:9px 14px;background:#fff;color:#475569;font-weight:800;cursor:pointer;white-space:nowrap}
.tab.active{background:#4f46e5;border-color:#4f46e5;color:#fff}
.status{padding:7px 22px;color:#64748b;font-size:11px;border-bottom:1px solid #e9eef7;flex-shrink:0}
.frames{flex:1;min-height:0;height:auto;overflow:hidden}
.board-frame{display:none;width:100%;height:100%;min-height:0;border:0;background:#f7f9fd}
.board-frame.active{display:block}
</style>
</head>
<body>
<div class="hub">
  <div class="hub-header">
    <div>
      <div class="hub-title">BvA Board Package</div>
      <div class="hub-sub">Each tab contains an independent editable board.</div>
    </div>
    <div class="hub-actions">
      <button class="hub-btn primary" id="download-zip">Download ZIP</button>
    </div>
  </div>
  <div class="tabs" id="tabs"></div>
  <div class="status" id="status"></div>
  <div class="frames" id="frames"></div>
</div>
<script id="bva-package-data">window.__BVA_PACKAGE_DATA__ = ${JSON.stringify(packageData)};</script>
<script>
(function(){
  function decodeBase64Utf8(value){
    var binary = atob(value || '');
    var bytes = new Uint8Array(binary.length);
    for(var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    if(window.TextDecoder) return new TextDecoder('utf-8').decode(bytes);
    var escaped = '';
    for(var j = 0; j < bytes.length; j++) escaped += '%' + ('00' + bytes[j].toString(16)).slice(-2);
    return decodeURIComponent(escaped);
  }

  function readBoards(){
    try{
      var payload = JSON.parse(decodeBase64Utf8(window.__BVA_PACKAGE_DATA__));
      return (payload.boards || []).map(function(board){
        return {
          title: board.title || '',
          subtitle: board.subtitle || '',
          fileName: board.fileName || 'board.html',
          html: decodeBase64Utf8(board.htmlBase64 || '')
        };
      });
    }catch(error){
      console.error('Could not decode BvA package data.', error);
      return [];
    }
  }

  var boards = readBoards();
  var tabs = document.getElementById('tabs');
  var frames = document.getElementById('frames');
  var status = document.getElementById('status');

  boards.forEach(function(board, index){
    var tab = document.createElement('button');
    tab.className = 'tab';
    tab.textContent = board.title || ('Board ' + (index + 1));
    tabs.appendChild(tab);

    var frame = document.createElement('iframe');
    frame.className = 'board-frame';
    frame.dataset.index = index;
    frame.srcdoc = board.html;
    frames.appendChild(frame);

    tab.addEventListener('click', function(){ activate(index); });
  });

  function activate(index){
    tabs.querySelectorAll('.tab').forEach(function(tab, i){
      tab.classList.toggle('active', i === index);
    });
    frames.querySelectorAll('.board-frame').forEach(function(frame, i){
      frame.classList.toggle('active', i === index);
    });
    var board = boards[index];
    status.textContent = board ? board.title + ' · ' + board.subtitle : '';
  }

  function waitForFrame(frame){
    if(frame.contentDocument && frame.contentDocument.readyState === 'complete'){
      return Promise.resolve();
    }
    return new Promise(function(resolve){
      frame.addEventListener('load', resolve, { once:true });
    });
  }

  function snapshotFrame(frame){
    var sourceDoc = frame.contentDocument;
    if(!sourceDoc) throw new Error('Could not read board frame');

    var clone = sourceDoc.documentElement.cloneNode(true);
    var sourceFields = sourceDoc.querySelectorAll('textarea,input,select');
    var clonedFields = clone.querySelectorAll('textarea,input,select');

    sourceFields.forEach(function(source, index){
      var target = clonedFields[index];
      if(!target) return;

      if(source.tagName === 'TEXTAREA'){
        target.textContent = source.value;
      }else if(source.type === 'checkbox' || source.type === 'radio'){
        if(source.checked) target.setAttribute('checked', 'checked');
        else target.removeAttribute('checked');
      }else if(source.tagName === 'SELECT'){
        Array.from(target.options).forEach(function(option, optionIndex){
          var selected = source.options[optionIndex] && source.options[optionIndex].selected;
          option.selected = selected;
          if(selected) option.setAttribute('selected', 'selected');
          else option.removeAttribute('selected');
        });
      }else{
        target.setAttribute('value', source.value);
      }
    });

    clone.querySelectorAll('#toast').forEach(function(element){ element.remove(); });
    return '<!DOCTYPE html>\\n' + clone.outerHTML;
  }

  function encodeBase64Utf8(value){
    var bytes = new TextEncoder().encode(String(value == null ? '' : value));
    var binary = '';
    for(var i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
  }

  async function downloadZip(){
    if(typeof JSZip === 'undefined'){
      alert('JSZip could not be loaded.');
      return;
    }

    var frameList = Array.from(frames.querySelectorAll('.board-frame'));
    await Promise.all(frameList.map(waitForFrame));

    var currentBoards = frameList.map(function(frame, index){
      return {
        title: boards[index].title,
        subtitle: boards[index].subtitle,
        fileName: boards[index].fileName,
        html: snapshotFrame(frame)
      };
    });

    var zip = new JSZip();
    currentBoards.forEach(function(board){ zip.file(board.fileName, board.html); });
    zip.file('index.html', buildUpdatedIndex(currentBoards));

    var blob = await zip.generateAsync({ type:'blob' });
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = 'bva_boards_updated.zip';
    link.click();
    setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
  }

  function buildUpdatedIndex(updatedBoards){
    var payload = {
      boards: updatedBoards.map(function(board){
        return {
          title: board.title,
          subtitle: board.subtitle,
          fileName: board.fileName,
          htmlBase64: encodeBase64Utf8(board.html)
        };
      })
    };
    var encoded = encodeBase64Utf8(JSON.stringify(payload));
    var clone = document.documentElement.cloneNode(true);
    var dataScript = clone.querySelector('#bva-package-data');
    if(dataScript){
      dataScript.textContent = 'window.__BVA_PACKAGE_DATA__ = ' + JSON.stringify(encoded) + ';';
    }
    var clonedTabs = clone.querySelector('#tabs');
    var clonedFrames = clone.querySelector('#frames');
    var clonedStatus = clone.querySelector('#status');
    if(clonedTabs) clonedTabs.innerHTML = '';
    if(clonedFrames) clonedFrames.innerHTML = '';
    if(clonedStatus) clonedStatus.textContent = '';
    return '<!DOCTYPE html>\\n' + clone.outerHTML;
  }

  document.getElementById('download-zip').addEventListener('click', downloadZip);
  if(boards.length) activate(0);
})();
</script>
</body>
</html>`;
}

async function downloadPdf(){
if(!state.model){ toast('Build a dashboard first'); return; }
if(typeof html2canvas === 'undefined' || !window.jspdf){
toast('PDF libraries did not load — check your connection and try again');
return;
}
saveState();
const nav = document.getElementById('download-pdf-nav');
const target = document.querySelector('.main-workspace');
if(!target){ toast('Could not find dashboard content'); return; }
const hiddenEls = Array.from(target.querySelectorAll('.hidden, .del-btn, .del-block, .add-act, .mini-btn, .ghost-btn, .topbar-actions, #back-nav, #home-nav'));
const restoreDisplay = hiddenEls.map(el => [el, el.style.display]);
const aside = document.querySelector('aside');
const asideDisplay = aside ? aside.style.display : null;
const prevMarginLeft = target.style.marginLeft;
const prevWidth = target.style.width;
try{
if(nav) nav.classList.add('disabled-nav');
toast('Building PDF…');
hiddenEls.forEach(el => { el.style.display = 'none'; });
if(aside) aside.style.display = 'none';
target.style.marginLeft = '0';
target.style.width = '100%';
await new Promise(r => setTimeout(r, 60));
const canvas = await html2canvas(target, { scale: 2, useCORS: true, backgroundColor: '#f8fafc' });
const { jsPDF } = window.jspdf;
const pdf = new jsPDF('p', 'pt', 'a4');
const pageWidth = pdf.internal.pageSize.getWidth();
const pageHeight = pdf.internal.pageSize.getHeight();
const ratio = canvas.width / pageWidth;
const pageHeightPx = Math.max(1, Math.floor(pageHeight * ratio));
let renderedHeight = 0;
let first = true;
while(renderedHeight < canvas.height){
const sliceHeight = Math.min(pageHeightPx, canvas.height - renderedHeight);
const sliceCanvas = document.createElement('canvas');
sliceCanvas.width = canvas.width;
sliceCanvas.height = sliceHeight;
const ctx = sliceCanvas.getContext('2d');
ctx.drawImage(canvas, 0, renderedHeight, canvas.width, sliceHeight, 0, 0, canvas.width, sliceHeight);
const imgData = sliceCanvas.toDataURL('image/jpeg', 0.92);
const sliceHeightPt = sliceHeight / ratio;
if(!first) pdf.addPage();
pdf.addImage(imgData, 'JPEG', 0, 0, pageWidth, sliceHeightPt);
renderedHeight += sliceHeight;
first = false;
}
pdf.save(`${exportFileBase()}.pdf`);
toast('PDF downloaded');
}catch(err){
console.error(err);
toast('Could not export PDF');
}finally{
restoreDisplay.forEach(([el, disp]) => { el.style.display = disp; });
if(aside) aside.style.display = asideDisplay;
target.style.marginLeft = prevMarginLeft;
target.style.width = prevWidth;
if(nav) nav.classList.remove('disabled-nav');
}
}
function openBudgetUtilization(){
if(!state.model || !state.model.opex){ toast('Upload the OPEX Feed to view Budget Utilization'); return; }
const payload = {
meta: state.model.meta,
opex: state.model.opex,
savedAt: Date.now()
};
try{ localStorage.setItem('bva:budget-utilization', JSON.stringify(payload)); }
catch(e){ console.warn(e); }
window.open('./budget-utilization/index.html', 'BudgetUtilization', 'width=1200,height=800,scrollbars=yes,resizable=yes');
}
function navTo(sectionId, el){
const target = document.getElementById(sectionId);
if(!target) return;
const top = target.getBoundingClientRect().top + window.scrollY - 20;
window.scrollTo({ top, behavior:'smooth' });
document.querySelectorAll('aside nav ul li[data-nav]').forEach(li => li.classList.remove('active'));
if(el) el.classList.add('active');
}
function scrollSpy(){
const sections = ['sec-overview','sec-qplan','sec-qfcst','sec-fyplan','sec-fyfcst','sec-te','sec-hc','sec-actions'];
const navItems = document.querySelectorAll('aside nav ul li[data-nav]');
const scrollY = window.scrollY + 80;
let current = 0;
sections.forEach((id, i) => { const el = document.getElementById(id); if(el && el.offsetTop <= scrollY) current = i; });
navItems.forEach(li => li.classList.remove('active'));
if(navItems[current]) navItems[current].classList.add('active');
}
function decodeHtmlEntities(str){
const ta = document.createElement('textarea');
ta.innerHTML = str;
return ta.value;
}
function parseEmbeddedJson(node){
const raw = node ? String(node.textContent || '').trim() : '';
if(!raw) return null;
try{ return JSON.parse(raw); }catch(e){}
try{ return JSON.parse(decodeHtmlEntities(raw)); }catch(e){}
return null;
}
function destroyChart(id){
if(chartRefs[id]){
try{ chartRefs[id].destroy(); }catch(e){}
delete chartRefs[id];
}
}
function renderCharts(model){
renderTrendChart(model.year.trend);
['sec-qplan','sec-qfcst','sec-fyplan','sec-fyfcst'].forEach(id => {
const node = document.getElementById(id + '-wf-data');
const data = parseEmbeddedJson(node);
if(!data) return;
const padded = waterfallBounds(data.baseStart, data.vars, data.workingEnd);
buildWF(id + '-wf', data.labels, data.baseStart, data.vars, data.workingEnd, padded.min, padded.max);
});
}
function waterfallBounds(base, vars, end){
let run = base;
let min = Math.min(base, end), max = Math.max(base, end);
vars.forEach(v => { run += v; min = Math.min(min, run); max = Math.max(max, run); });
const pad = Math.max(50000, Math.ceil((max - min) * 0.1 / 50000) * 50000 || 50000);
return { min: Math.floor((min - pad) / 50000) * 50000, max: Math.ceil((max + pad) / 50000) * 50000 };
}
function buildWF(canvasId, labels, baseVal, varsArr, endVal, yMin, yMax) {
var bases=[], bars=[], bgs=[], run=baseVal;
bases.push(0); bars.push(baseVal); bgs.push('#4f46e5');
varsArr.forEach(function(v) {
bases.push(v < 0 ? run + v : run);
bars.push(Math.abs(v));
bgs.push(v < 0 ? '#16a34a' : '#e11d48');
run += v;
});
bases.push(0); bars.push(endVal); bgs.push('#4f46e5');
var canvas = document.getElementById(canvasId); if (!canvas) return;
destroyChart(canvasId);
chartRefs[canvasId] = new Chart(canvas, {
type: 'bar',
data: {
labels: labels,
datasets: [
{ data: bases, backgroundColor: 'rgba(0,0,0,0)', borderWidth: 0, datalabels: { display: false } },
{ data: bars, backgroundColor: bgs, borderWidth: 0, borderRadius: 6,
datalabels: {
display: true, anchor: 'end', align: 'end', offset: 2,
color: '#334155', font: { size: 13, weight: '700' },
formatter: function(v, ctx) {
var i = ctx.dataIndex;
if (i === 0 || i === bars.length - 1) return '$' + Math.round((bases[i] + v) / 1000) + 'K';
var o = varsArr[i - 1];
return (o > 0 ? '+' : '-') + '$' + Math.round(Math.abs(o) / 1000) + 'K';
}
}
}
]
},
options: {
responsive: true, maintainAspectRatio: false, layout: { padding: { top: 44 } },
scales: {
x: { stacked: true, grid: { display: false }, ticks: { font: { size: 9.5 }, color: '#64748b', autoSkip: false, maxRotation: 20 } },
y: { stacked: true, min: yMin, max: yMax, grid: { color: '#f1f5f9' }, ticks: { font: { size: 9 }, color: '#64748b', callback: function(v) { return '$' + (v/1000).toFixed(0) + 'K'; } } }
},
plugins: { legend: { display: false }, tooltip: { enabled: false } }
}
});
}
function renderTrendChart(trend){
const canvas = document.getElementById('trendC');
if(!canvas) return;
const allVals = [...trend.working, ...trend.plan].filter(n => Number.isFinite(n));
const min = allVals.length ? Math.floor((Math.min(...allVals) - 50000)/50000)*50000 : 0;
const max = allVals.length ? Math.ceil((Math.max(...allVals) + 50000)/50000)*50000 : 100000;
destroyChart('trendC');
chartRefs['trendC'] = new Chart(canvas, {
type: 'line',
data: {
labels: trend.labels,
datasets: [
{ label: 'Working', data: trend.working, borderColor: '#16a34a', backgroundColor: 'rgba(22,163,74,0.14)', fill: true, borderWidth: 3, pointRadius: 4, pointBackgroundColor: '#16a34a', tension: 0.35, datalabels: { display: false } },
{ label: 'Plan', data: trend.plan, borderColor: '#2563eb', backgroundColor: 'rgba(37,99,235,0.10)', fill: true, borderWidth: 2.5, pointRadius: 4, pointBackgroundColor: '#2563eb', borderDash: [6,4], tension: 0.35, datalabels: { display: false } }
]
},
options: {
responsive: true, maintainAspectRatio: false,
scales: {
x: { grid: { display: false }, ticks: { font: { size: 10 }, color: '#64748b' } },
y: { min, max, grid: { color: '#f1f5f9' }, ticks: { font: { size: 9 }, color: '#64748b', callback: function(v) { return '$' + (v/1000).toFixed(0) + 'K'; } } }
},
plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c) { return c.dataset.label + ': $' + Math.round(c.raw/1000) + 'K'; } } } }
}
});
}
// ---------- Vendor drill-down (piloto Q vs Plan) ----------
function buildGlMap(rows){
const map = {}; let pending = [];
for(let i=0;i<rows.length;i++){
const r = rows[i];
if(r.rowType === 'vendor' || r.rowType === 'novendor'){ pending.push(r.index); continue; }
if(r.rowType === 'gl'){ pending.forEach(idx => { map[idx] = r.label; }); pending = []; continue; }
if(r.rowType === 'l2' || r.rowType === 'expense'){ pending = []; }
}
return map;
}
function collectVendorRows(allRows, l2RowIndex){
const idx = allRows.findIndex(r => r.index === l2RowIndex);
const out = [];
for(let i = idx - 1; i >= 0; i--){
const row = allRows[i];
if(row.rowType === 'l2' || row.rowType === 'expense') break;
if(row.rowType === 'vendor' || row.rowType === 'novendor') out.push(row);
}
return out;
}
function vendorPeriodValues(row, period, benchmarkKey){
if(period === 'q') return { working: row.total.w, benchmark: row.total[benchmarkKey] };
const m = row.months[period];
return { working: m.w, benchmark: m[benchmarkKey] };
}
function computeQuarterVendorBreakdown(l2RowIndex, period, benchmarkKey){
const allRows = (state.model && state.model.quarter && state.model.quarter.allRows) || [];
const glMap = buildGlMap(allRows);
const target = allRows.find(r => r.index === l2RowIndex);
let vendorRows;
if(target && target.rowType === 'expense'){
vendorRows = allRows.filter(r => r.rowType === 'vendor' || r.rowType === 'novendor');
} else {
vendorRows = collectVendorRows(allRows, l2RowIndex);
}
return vendorRows.map(row => {
const vals = vendorPeriodValues(row, period, benchmarkKey);
return { name: cleanLabel(row.label), gl: glMap[row.index] || '', working: vals.working, benchmark: vals.benchmark, variance: vals.working - vals.benchmark };
}).sort((a,b) => Math.abs(b.variance) - Math.abs(a.variance));
}
function collectVendorRowsFlex(allRows, l2RowIndex){
const idx = allRows.findIndex(r => r.index === l2RowIndex);
const vend = [], gl = [];
for(let i = idx - 1; i >= 0; i--){
const row = allRows[i];
if(row.rowType === 'l2' || row.rowType === 'expense') break;
if(row.rowType === 'vendor' || row.rowType === 'novendor') vend.push(row);
else if(row.rowType === 'gl') gl.push(row);
}
return vend.length ? vend : gl;
}
function yearVendorPeriodValues(row, period, benchmarkKey){
if(period === 'fy') return { working: row.total.w, benchmark: row.total[benchmarkKey] };
const q = row.quarters[period];
return { working: q.w, benchmark: q[benchmarkKey] };
}
function computeYearVendorBreakdown(l2RowIndex, period, benchmarkKey){
const allRows = (state.model && state.model.year && state.model.year.allRows) || [];
const glMap = buildGlMap(allRows);
const target = allRows.find(r => r.index === l2RowIndex);
let vendorRows;
if(target && target.rowType === 'expense'){
vendorRows = allRows.filter(r => r.rowType === 'vendor' || r.rowType === 'novendor');
if(!vendorRows.length) vendorRows = allRows.filter(r => r.rowType === 'gl');
} else {
vendorRows = collectVendorRowsFlex(allRows, l2RowIndex);
}
return vendorRows.map(row => {
const vals = yearVendorPeriodValues(row, period, benchmarkKey);
return { name: cleanLabel(row.label), gl: glMap[row.index] || '', working: vals.working, benchmark: vals.benchmark, variance: vals.working - vals.benchmark };
}).sort((a,b) => Math.abs(b.variance) - Math.abs(a.variance));
}
// Self-contained drill-down used by the exported HTML (app.js is stripped on export,
// so this function carries its own helpers + reads vendor data from an embedded object).
function initDrilldown(root, data){
data = data || { quarter: [], year: [] };
function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function fmtK(n){ if(!isFinite(n)||n===0) return '—'; var r=Math.round(n/1000); var a=Math.abs(r).toLocaleString('en-US'); return r>0?('+$'+a+'K'):('-$'+a+'K'); }
function fmtKplain(n){ if(!isFinite(n)||n===0) return '—'; var r=Math.round(n/1000); var a=Math.abs(r).toLocaleString('en-US'); return r<0?('($'+a+'K)'):('$'+a+'K'); }
function cleanLabel(s){ return String(s==null?'':s).replace(/^Total\s+/,''); }
function collect(rows, l2Index, flex){
var idx=rows.findIndex(function(r){return r.index===l2Index;});
var vend=[], gl=[];
for(var i=idx-1;i>=0;i--){ var row=rows[i]; if(row.rowType==='l2'||row.rowType==='expense') break; if(row.rowType==='vendor'||row.rowType==='novendor') vend.push(row); else if(row.rowType==='gl') gl.push(row); }
return flex ? (vend.length?vend:gl) : vend;
}
function glmap(rows){ var m={},pending=[]; for(var i=0;i<rows.length;i++){ var r=rows[i]; if(r.rowType==='vendor'||r.rowType==='novendor'){pending.push(r.index);continue;} if(r.rowType==='gl'){for(var k=0;k<pending.length;k++) m[pending[k]]=r.label; pending=[];continue;} if(r.rowType==='l2'||r.rowType==='expense'){pending=[];} } return m; }
function qVals(row, period, bk){ if(period==='q') return {working:row.total.w, benchmark:row.total[bk]}; var m=row.months[period]; return {working:m.w, benchmark:m[bk]}; }
function yVals(row, period, bk){ if(period==='fy') return {working:row.total.w, benchmark:row.total[bk]}; var q=row.quarters[period]; return {working:q.w, benchmark:q[bk]}; }
function computeQ(rowIndex, period, bk){
var rows=data.quarter||[]; var gm=glmap(rows); var t=rows.find(function(r){return r.index===rowIndex;}); var vr;
if(t&&t.rowType==='expense') vr=rows.filter(function(r){return r.rowType==='vendor'||r.rowType==='novendor';}); else vr=collect(rows,rowIndex,false);
return vr.map(function(row){ var val=qVals(row,period,bk); return {name:cleanLabel(row.label), gl:gm[row.index]||'', working:val.working, benchmark:val.benchmark, variance:val.working-val.benchmark}; }).sort(function(a,b){return Math.abs(b.variance)-Math.abs(a.variance);});
}
function computeY(rowIndex, period, bk){
var rows=data.year||[]; var gm=glmap(rows); var t=rows.find(function(r){return r.index===rowIndex;}); var vr;
if(t&&t.rowType==='expense'){ vr=rows.filter(function(r){return r.rowType==='vendor'||r.rowType==='novendor';}); if(!vr.length) vr=rows.filter(function(r){return r.rowType==='gl';}); } else vr=collect(rows,rowIndex,true);
return vr.map(function(row){ var val=yVals(row,period,bk); return {name:cleanLabel(row.label), gl:gm[row.index]||'', working:val.working, benchmark:val.benchmark, variance:val.working-val.benchmark}; }).sort(function(a,b){return Math.abs(b.variance)-Math.abs(a.variance);});
}
var ctx=null;
function ensure(){
var o=document.getElementById('drill-overlay'); if(o) return o;
o=document.createElement('div'); o.id='drill-overlay'; o.className='drill-overlay';
o.innerHTML='<div class="drill-panel" id="drill-panel" role="dialog" aria-modal="true"></div>';
document.body.appendChild(o);
o.addEventListener('click', function(e){ if(e.target===o) close(); });
document.addEventListener('keydown', function(e){ if(e.key==='Escape') close(); });
return o;
}
function close(){ var o=document.getElementById('drill-overlay'); var p=document.getElementById('drill-panel'); if(p)p.classList.remove('show'); if(o)o.classList.remove('show'); }
function open(cell){
var rowIndex=Number(cell.dataset.rowIndex), scope=cell.dataset.scope||'q', bk=cell.dataset.benchmark, bl=bk==='p'?'Plan':'FCST', vendors, threshold;
if(scope==='y'){ var py=cell.dataset.period==='fy'?'fy':Number(cell.dataset.period); vendors=computeY(rowIndex,py,bk); threshold=75; }
else { var pq=cell.dataset.period==='q'?'q':Number(cell.dataset.period); vendors=computeQ(rowIndex,pq,bk); threshold=25; }
render({title:cell.dataset.label, periodLabel:cell.dataset.periodLabel, benchmarkLabel:bl, vendors:vendors, threshold:threshold});
}
function render(c){
ctx=Object.assign({mode:'materiality', threshold:25, search:''}, c);
var overlay=ensure(), panel=document.getElementById('drill-panel');
panel.innerHTML=''
+'<div class="drill-hdr"><div><h3>'+esc(c.title)+'</h3><div class="drill-sub">'+esc(c.periodLabel)+' \u00b7 Working vs '+esc(c.benchmarkLabel)+'</div></div><button class="drill-close" id="drill-close">&times;</button></div>'
+'<div class="drill-controls"><div class="drill-seg"><button class="drill-seg-btn" data-mode="materiality">By materiality</button><button class="drill-seg-btn" data-mode="activity" title="With Activity (excluding zero)">With activity</button><button class="drill-seg-btn" data-mode="all">Show all</button></div>'
+'<div class="drill-thr" id="drill-thr-wrap"><span>\u00b1\u00a0$</span><input type="number" id="drill-thr" min="0" step="5" value="'+c.threshold+'" /><span>K</span></div></div>'
+'<div class="drill-search" id="drill-search-wrap"><i class="ti ti-search"></i><input type="text" id="drill-search" placeholder="Search vendor by id or name..." /></div>'
+'<div class="drill-body" id="drill-body"></div>';
panel.querySelector('#drill-close').addEventListener('click', close);
panel.querySelectorAll('.drill-seg-btn').forEach(function(btn){ btn.addEventListener('click', function(){ ctx.mode=btn.dataset.mode; body(); }); });
var thr=panel.querySelector('#drill-thr'); if(thr) thr.addEventListener('input', function(){ var val=parseFloat(thr.value); ctx.threshold=isFinite(val)?val:0; if(ctx.mode==='materiality') body(); });
var s=panel.querySelector('#drill-search'); if(s) s.addEventListener('input', function(){ ctx.search=s.value; if(ctx.mode==='all') body(); });
var be=panel.querySelector('#drill-body'); if(be) be.addEventListener('click', function(e){ var b=e.target.closest('.drill-gl-btn'); if(!b) return; var row=b.closest('.drill-bar-row'); if(!row) return; var open=row.classList.toggle('gl-open'); var ic=b.querySelector('i'); if(ic) ic.className=open?'ti ti-eye-off':'ti ti-eye'; });
body(); overlay.classList.add('show'); requestAnimationFrame(function(){ panel.classList.add('show'); });
}
function body(){
var panel=document.getElementById('drill-panel'); if(!panel||!ctx) return;
var bd=panel.querySelector('#drill-body');
panel.querySelectorAll('.drill-seg-btn').forEach(function(b){ b.classList.toggle('active', b.dataset.mode===ctx.mode); });
var tw=panel.querySelector('#drill-thr-wrap'); if(tw) tw.style.display=ctx.mode==='materiality'?'flex':'none';
var sw=panel.querySelector('#drill-search-wrap'); if(sw) sw.style.display=ctx.mode==='all'?'flex':'none';
var all=ctx.vendors;
var totW=all.reduce(function(s,v){return s+v.working;},0), totB=all.reduce(function(s,v){return s+v.benchmark;},0), totVar=totW-totB, util=totB?Math.round(totW/totB*100):null;
var thrAbs=(ctx.threshold||0)*1000, shown;
if(ctx.mode==='materiality') shown=all.filter(function(v){return Math.abs(v.variance)>=thrAbs;});
else if(ctx.mode==='activity') shown=all.filter(function(v){return v.working!==0||v.benchmark!==0;});
else { shown=all; var q=(ctx.search||'').trim().toLowerCase(); if(q) shown=shown.filter(function(v){return v.name.toLowerCase().indexOf(q)!==-1;}); }
var maxVal=Math.max.apply(null,[1].concat(shown.map(function(v){return Math.max(Math.abs(v.working),Math.abs(v.benchmark));})));
function bar(v){
var wPct=Math.min(100,Math.abs(v.working)/maxVal*100), bPct=Math.min(100,Math.abs(v.benchmark)/maxVal*100), over=v.variance>0;
var vu=v.benchmark?((v.variance>0?'+':'')+Math.round(v.variance/v.benchmark*100)+'%'):(v.working?'not in plan':'\u2014');
return '<div class="drill-bar-row"><div class="drill-bar-top"><span class="drill-bar-left"><span class="drill-bar-name">'+esc(v.name)+'</span>'+(v.gl?'<button class="drill-gl-btn" type="button" title="Show GL account"><i class="ti ti-eye"></i></button>':'')+'</span><span class="drill-bar-fig">'+fmtKplain(v.working)+' / '+fmtKplain(v.benchmark)+'</span></div><div class="drill-track"><div class="drill-fill '+(over?'unfav':'fav')+'" style="width:'+wPct+'%"></div>'+(v.benchmark?'<div class="drill-plan-marker" style="left:'+bPct+'%"></div>':'')+'</div><div class="drill-util '+(over?'var-unfav':'var-fav')+'">'+esc(vu)+' vs '+esc(ctx.benchmarkLabel.toLowerCase())+' \u00b7 '+fmtK(v.variance)+'</div>'+(v.gl?'<div class="drill-gl-line"><i class="ti ti-receipt-2"></i>GL account \u00b7 '+esc(v.gl)+'</div>':'')+'</div>';
}
var unfav=shown.filter(function(v){return v.variance>0;}).sort(function(a,b){return b.variance-a.variance;});
var fav=shown.filter(function(v){return v.variance<0;}).sort(function(a,b){return a.variance-b.variance;});
var neu=shown.filter(function(v){return v.variance===0;});
function gh(cls,main,note,count){ return '<div class="drill-group-hdr '+cls+'"><span class="ghl"><span class="gmain">'+main+'</span><span class="gnote">('+note+')</span></span><span class="gcount">'+count+'</span></div>'; }
var bars='';
if(fav.length) bars+=gh('fav','Favorable','Savings',fav.length)+fav.map(bar).join('');
if(unfav.length) bars+=gh('unfav','Unfavorable','Overspend',unfav.length)+unfav.map(bar).join('');
if(neu.length) bars+=gh('neu','No Variance','In line with Plan',neu.length)+neu.map(bar).join('');
var ct;
if(ctx.mode==='materiality') ct='Showing '+shown.length+' of '+all.length+' vendors \u00b7 materiality \u00b1 $'+ctx.threshold+'K';
else if(ctx.mode==='activity') ct='Showing '+shown.length+' of '+all.length+' vendors \u00b7 with activity (excluding zero)';
else { var q2=(ctx.search||'').trim(); ct=q2?('Showing '+shown.length+' of '+all.length+' vendors \u00b7 search "'+q2+'"'):('Showing all '+all.length+' vendors'); }
bd.innerHTML='<div class="drill-summary"><div class="ds"><div class="k">Working</div><div class="val">'+fmtKplain(totW)+'</div></div><div class="ds"><div class="k">'+esc(ctx.benchmarkLabel)+'</div><div class="val">'+fmtKplain(totB)+'</div></div><div class="ds"><div class="k">Variance</div><div class="val '+(totVar<0?'kpi-fav':totVar>0?'kpi-unfav':'kpi-neu')+'">'+fmtK(totVar)+'</div></div><div class="ds"><div class="k">Utilization</div><div class="val">'+(util===null?'\u2014':util+'%')+'</div></div></div><div class="drill-count">'+esc(ct)+'</div>'+(shown.length?bars:('<div class="drill-empty">'+(ctx.mode==='materiality'?'No vendors within this materiality range.':'No vendors to display.')+'</div>'));
}
var container=root||document;
container.addEventListener('click', function(e){ var t=e.target; if(!t||!t.closest) return; var cell=t.closest('[data-drill="1"]'); if(cell) open(cell); });
}
// Build the minimal vendor dataset embedded into the exported HTML.
function buildDrillData(){
var m = state.model || {};
function slim(rows, kind){
return (rows||[]).map(function(r){
var o = { index:r.index, rowType:r.rowType, label:r.label, total:{ w:r.total.w, p:r.total.p, f:r.total.f } };
if(kind==='q') o.months = (r.months||[]).map(function(x){ return { w:x.w, p:x.p, f:x.f }; });
else o.quarters = (r.quarters||[]).map(function(x){ return { w:x.w, p:x.p, f:x.f }; });
return o;
});
}
return { quarter: slim(m.quarter && m.quarter.allRows, 'q'), year: slim(m.year && m.year.allRows, 'y') };
}
function appendDrillExportScript(clone){
try{
var json = JSON.stringify(buildDrillData()).replace(/</g, '\\u003c');
var script = document.createElement('script');
script.textContent = '(' + initDrilldown.toString() + ')(document, ' + json + ');';
var body = clone.querySelector('body');
if(body) body.appendChild(script);
}catch(e){ console.warn('Could not embed drill-down into export.', e); }
}
function ensureDrillPanel(){
let overlay = document.getElementById('drill-overlay');
if(overlay) return overlay;
overlay = document.createElement('div');
overlay.id = 'drill-overlay';
overlay.className = 'drill-overlay';
overlay.innerHTML = '<div class="drill-panel" id="drill-panel" role="dialog" aria-modal="true"></div>';
document.body.appendChild(overlay);
overlay.addEventListener('click', e => { if(e.target === overlay) closeDrill(); });
document.addEventListener('keydown', e => { if(e.key === 'Escape') closeDrill(); });
return overlay;
}
function closeDrill(){
const overlay = document.getElementById('drill-overlay');
const panel = document.getElementById('drill-panel');
if(panel) panel.classList.remove('show');
if(overlay) overlay.classList.remove('show');
}
function openDrill(cell){
const rowIndex = Number(cell.dataset.rowIndex);
const scope = cell.dataset.scope || 'q';
const benchmarkKey = cell.dataset.benchmark;
const benchmarkLabel = benchmarkKey === 'p' ? 'Plan' : 'FCST';
let vendors, threshold;
if(scope === 'y'){
const period = cell.dataset.period === 'fy' ? 'fy' : Number(cell.dataset.period);
vendors = computeYearVendorBreakdown(rowIndex, period, benchmarkKey);
threshold = 75;
} else {
const period = cell.dataset.period === 'q' ? 'q' : Number(cell.dataset.period);
vendors = computeQuarterVendorBreakdown(rowIndex, period, benchmarkKey);
threshold = 25;
}
renderDrillPanel({ title: cell.dataset.label, periodLabel: cell.dataset.periodLabel, benchmarkLabel, vendors, threshold });
}
let drillCtx = null;
function renderDrillPanel(ctx){
drillCtx = Object.assign({ mode:'materiality', threshold:25, search:'' }, ctx);
const overlay = ensureDrillPanel();
const panel = document.getElementById('drill-panel');
panel.innerHTML = `
<div class="drill-hdr">
<div>
<h3>${escapeHtml(ctx.title)}</h3>
<div class="drill-sub">${escapeHtml(ctx.periodLabel)} · Working vs ${escapeHtml(ctx.benchmarkLabel)}</div>
</div>
<button class="drill-close" id="drill-close">&times;</button>
</div>
<div class="drill-controls">
<div class="drill-seg">
<button class="drill-seg-btn" data-mode="materiality">By materiality</button>
<button class="drill-seg-btn" data-mode="activity" title="With Activity (excluding zero)">With activity</button>
<button class="drill-seg-btn" data-mode="all">Show all</button>
</div>
<div class="drill-thr" id="drill-thr-wrap">
<span>±&nbsp;$</span><input type="number" id="drill-thr" min="0" step="5" value="${ctx.threshold}" /><span>K</span>
</div>
</div>
<div class="drill-search" id="drill-search-wrap">
<i class="ti ti-search"></i><input type="text" id="drill-search" placeholder="Search vendor by id or name..." value="${escapeHtml(ctx.search||'')}" />
</div>
<div class="drill-body" id="drill-body"></div>`;
panel.querySelector('#drill-close').addEventListener('click', closeDrill);
panel.querySelectorAll('.drill-seg-btn').forEach(btn => {
btn.addEventListener('click', () => { drillCtx.mode = btn.dataset.mode; renderDrillBody(); });
});
const thr = panel.querySelector('#drill-thr');
if(thr) thr.addEventListener('input', () => {
const val = parseFloat(thr.value);
drillCtx.threshold = Number.isFinite(val) ? val : 0;
if(drillCtx.mode === 'materiality') renderDrillBody();
});
const search = panel.querySelector('#drill-search');
if(search) search.addEventListener('input', () => { drillCtx.search = search.value; if(drillCtx.mode === 'all') renderDrillBody(); });
const bodyEl = panel.querySelector('#drill-body');
if(bodyEl) bodyEl.addEventListener('click', e => {
const btn = e.target.closest('.drill-gl-btn');
if(!btn) return;
const row = btn.closest('.drill-bar-row');
if(!row) return;
const open = row.classList.toggle('gl-open');
const ic = btn.querySelector('i');
if(ic) ic.className = open ? 'ti ti-eye-off' : 'ti ti-eye';
});
renderDrillBody();
overlay.classList.add('show');
requestAnimationFrame(() => panel.classList.add('show'));
}
function renderDrillBody(){
const panel = document.getElementById('drill-panel');
if(!panel || !drillCtx) return;
const body = panel.querySelector('#drill-body');
panel.querySelectorAll('.drill-seg-btn').forEach(b => b.classList.toggle('active', b.dataset.mode === drillCtx.mode));
const thrWrap = panel.querySelector('#drill-thr-wrap');
if(thrWrap) thrWrap.style.display = drillCtx.mode === 'materiality' ? 'flex' : 'none';
const searchWrap = panel.querySelector('#drill-search-wrap');
if(searchWrap) searchWrap.style.display = drillCtx.mode === 'all' ? 'flex' : 'none';
const all = drillCtx.vendors;
const totW = all.reduce((s,v)=>s+v.working,0);
const totB = all.reduce((s,v)=>s+v.benchmark,0);
const totVar = totW - totB;
const util = totB ? Math.round(totW/totB*100) : null;
const thrAbs = (drillCtx.threshold || 0) * 1000;
let shown;
if(drillCtx.mode === 'materiality'){
shown = all.filter(v => Math.abs(v.variance) >= thrAbs);
} else if(drillCtx.mode === 'activity'){
shown = all.filter(v => v.working !== 0 || v.benchmark !== 0);
} else {
shown = all;
const q = (drillCtx.search || '').trim().toLowerCase();
if(q) shown = shown.filter(v => v.name.toLowerCase().includes(q));
}
const maxVal = Math.max(1, ...shown.map(v => Math.max(Math.abs(v.working), Math.abs(v.benchmark))));
const renderBar = v => {
const wPct = Math.min(100, Math.abs(v.working)/maxVal*100);
const bPct = Math.min(100, Math.abs(v.benchmark)/maxVal*100);
const over = v.variance > 0;
const vpct = v.benchmark ? ((v.variance>0?'+':'') + Math.round(v.variance/v.benchmark*100) + '%') : (v.working ? 'not in plan' : '—');
return `<div class="drill-bar-row">
<div class="drill-bar-top">
<span class="drill-bar-left"><span class="drill-bar-name">${escapeHtml(v.name)}</span>${v.gl ? '<button class="drill-gl-btn" type="button" title="Show GL account"><i class="ti ti-eye"></i></button>' : ''}</span>
<span class="drill-bar-fig">${fmtKplain(v.working)} / ${fmtKplain(v.benchmark)}</span>
</div>
<div class="drill-track">
<div class="drill-fill ${over?'unfav':'fav'}" style="width:${wPct}%"></div>
${v.benchmark ? `<div class="drill-plan-marker" style="left:${bPct}%"></div>` : ''}
</div>
<div class="drill-util ${over?'var-unfav':'var-fav'}">${escapeHtml(vpct)} vs ${escapeHtml(drillCtx.benchmarkLabel.toLowerCase())} · ${fmtK(v.variance)}</div>
${v.gl ? `<div class="drill-gl-line"><i class="ti ti-receipt-2"></i>GL account · ${escapeHtml(v.gl)}</div>` : ''}
</div>`;
};
const unfav = shown.filter(v => v.variance > 0).sort((a,b)=>b.variance - a.variance);
const fav = shown.filter(v => v.variance < 0).sort((a,b)=>a.variance - b.variance);
const neutral = shown.filter(v => v.variance === 0);
const groupHdr = (cls, main, note, count) => `<div class="drill-group-hdr ${cls}"><span class="ghl"><span class="gmain">${main}</span><span class="gnote">(${note})</span></span><span class="gcount">${count}</span></div>`;
let bars = '';
if(fav.length) bars += groupHdr('fav', 'Favorable', 'Savings', fav.length) + fav.map(renderBar).join('');
if(unfav.length) bars += groupHdr('unfav', 'Unfavorable', 'Overspend', unfav.length) + unfav.map(renderBar).join('');
if(neutral.length) bars += groupHdr('neu', 'No Variance', 'In line with Plan', neutral.length) + neutral.map(renderBar).join('');
let countTxt;
if(drillCtx.mode === 'materiality'){
countTxt = `Showing ${shown.length} of ${all.length} vendors · materiality ± $${drillCtx.threshold}K`;
} else if(drillCtx.mode === 'activity'){
countTxt = `Showing ${shown.length} of ${all.length} vendors · with activity (excluding zero)`;
} else {
const q = (drillCtx.search || '').trim();
countTxt = q ? `Showing ${shown.length} of ${all.length} vendors · search \"${q}\"` : `Showing all ${all.length} vendors`;
}
body.innerHTML = `
<div class="drill-summary">
<div class="ds"><div class="k">Working</div><div class="val">${fmtKplain(totW)}</div></div>
<div class="ds"><div class="k">${escapeHtml(drillCtx.benchmarkLabel)}</div><div class="val">${fmtKplain(totB)}</div></div>
<div class="ds"><div class="k">Variance</div><div class="val ${totVar<0?'kpi-fav':totVar>0?'kpi-unfav':'kpi-neu'}">${fmtK(totVar)}</div></div>
<div class="ds"><div class="k">Utilization</div><div class="val">${util===null?'—':util+'%'}</div></div>
</div>
<div class="drill-count">${escapeHtml(countTxt)}</div>
${shown.length ? bars : `<div class="drill-empty">${drillCtx.mode === 'materiality' ? 'No vendors within this materiality range.' : 'No vendors to display.'}</div>`}`;
}
function quarterNum(label){ const m = String(label||'').match(/Q\s*([1-4])/i); return m ? parseInt(m[1],10)-1 : -1; }
function cleanLabel(label){ return String(label||'').replace(/^Total\s+/,''); }
function slug(s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/(^-|-$)/g,''); }
function escapeHtml(str){ return String(str||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function cssEscape(str){ return String(str).replace(/"/g,'\\"'); }
})();

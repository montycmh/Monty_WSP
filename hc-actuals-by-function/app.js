let hcaFile = null;

function hcaDragOver(e,id){ e.preventDefault(); document.getElementById(id).style.borderColor='#1e3a5f'; }
function hcaDragLeave(id){ document.getElementById(id).style.borderColor='#D3D1C7'; }
function hcaDrop(e,idx){ e.preventDefault(); document.getElementById('hca-drop-0').style.borderColor='#D3D1C7'; const f=e.dataTransfer.files[0]; if(f) hcaSetFile(f); }
function hcaFileSelected(idx,input){ if(input.files[0]) hcaSetFile(input.files[0]); }

function hcaSetFile(file){
  hcaFile = file;
  document.getElementById('hca-icon-0').className = 'ti ti-file-check';
  document.getElementById('hca-icon-0').style.color = '#1e3a5f';
  document.getElementById('hca-label-0').textContent = file.name;
  document.getElementById('hca-label-0').style.color = '#1e3a5f';
  const drop = document.getElementById('hca-drop-0');
  drop.classList.add('filled');
  drop.style.borderColor = '#1e3a5f';
  drop.style.borderStyle = 'solid';
  const btn = document.getElementById('hca-generate-btn');
  btn.disabled = false;
  btn.classList.add('enabled');
}

async function hcaGenerateReport(){
  const btn = document.getElementById('hca-generate-btn');
  btn.innerHTML = '<i class="ti ti-loader-2"></i>Processing...';
  btn.disabled = true;
  btn.classList.remove('enabled');

  try {
    if(typeof XLSX === 'undefined'){
      await new Promise((res,rej)=>{ const s=document.createElement('script'); s.src='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'; s.onload=res; s.onerror=rej; document.head.appendChild(s); });
    }

    const fname = hcaFile.name;
    const m = fname.match(/HC_Actuals_by[ _]function_?([A-Za-z]+)(\d{2})\.xlsx$/i);
    if(!m) throw new Error(`Filename "${fname}" doesn't match HC_Actuals_by_function_MmmYY.xlsx`);

    const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const monStr = m[1];
    const yr2 = parseInt(m[2], 10);
    const yr = 2000 + yr2;
    const monIdx = MONTH_NAMES.findIndex(x=>x.toLowerCase()===monStr.slice(0,3).toLowerCase());
    if(monIdx < 0) throw new Error(`Cannot parse month "${monStr}"`);

    function priorMonth(baseIdx, baseYr, offset){
      let idx = baseIdx - offset;
      let y = baseYr;
      while(idx < 0){ idx += 12; y--; }
      return { name: MONTH_NAMES[idx], yr: y, label: `${MONTH_NAMES[idx]} ${String(y).slice(2)}` };
    }

    const M4 = { name: MONTH_NAMES[monIdx], yr, label: `${MONTH_NAMES[monIdx]} ${yr2}` };
    const M3 = priorMonth(monIdx, yr, 1);
    const M2 = priorMonth(monIdx, yr, 2);
    const M1 = priorMonth(monIdx, yr, 3);
    const months = [M1, M2, M3, M4];

    const wb = await new Promise((res,rej)=>{
      const reader = new FileReader();
      reader.onload = e=>{ try{ res(XLSX.read(new Uint8Array(e.target.result), {type:'array'})); }catch(err){ rej(err); } };
      reader.onerror = rej;
      reader.readAsArrayBuffer(hcaFile);
    });

    const mappingSheet = wb.Sheets['Mapping'] || wb.Sheets[wb.SheetNames[1]];
    if(!mappingSheet) throw new Error('Mapping sheet not found');
    const mappingRaw = XLSX.utils.sheet_to_json(mappingSheet, {header:1, defval:''});

    const normClass = v => String(v == null ? '' : v).replace(/\s+/g,' ').replace(/\s*-\s*/g,' - ').trim();
    const classToBU = {};
    mappingRaw.slice(1).forEach(row => {
      const bu = String(row[0] == null ? '' : row[0]).trim();
      const cls = normClass(row[1]);
      if(bu && cls) classToBU[cls] = bu;
    });
    if(!Object.keys(classToBU).length) throw new Error('Could not build Mapping lookup from Mapping sheet.');

    const sheet1 = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet1, {header:1, defval:null});

    let headerRowIdx = -1;
    let colMap = {};
    for(let i=0; i<Math.min(10, raw.length); i++){
      const row = raw[i] || [];
      let found = 0;
      const temp = {};
      months.forEach(mo => {
        const target = mo.label.toLowerCase();
        const ci = row.findIndex(c => c && String(c).trim().toLowerCase() === target);
        if(ci >= 0){ temp[mo.label] = ci; found++; }
      });
      if(found === 4){ headerRowIdx = i; colMap = temp; break; }
    }
    if(headerRowIdx < 0) throw new Error('Could not find month headers in sheet1.');

    const ALL_SECTION_HEADERS = new Set([
      'FTE',
      'New Hires',
      'Transfer In',
      'Transfer Out',
      'CMH - Actual Termination',
      'Headcount (Baseline)',
      'Headcount (w/Attrition)',
      'Net Adds (w/Attrition)'
    ]);
    const MOVEMENT_SECTIONS = new Set(['New Hires','Transfer In','Transfer Out','CMH - Actual Termination']);

    const buData = {};
    function ensureBU(bu){
      if(!buData[bu]) buData[bu] = {
        hc:{[M1.label]:0,[M2.label]:0,[M3.label]:0,[M4.label]:0},
        new:0,
        att:0,
        trans:0,
        preM4:0
      };
    }
    const toInt = v => (v===null || v===undefined || v==='') ? 0 : Math.round(parseFloat(String(v).replace(/,/g,'')) || 0);

    let currentSection = null;
    for(let i=headerRowIdx + 1; i<raw.length; i++){
      const row = raw[i] || [];
      const cell0 = row[0] !== null && row[0] !== undefined ? String(row[0]).trim() : '';
      const cls = normClass(row[1]);

      if(cell0 && ALL_SECTION_HEADERS.has(cell0)){
        currentSection = cell0;
      } else if(cell0) {
        currentSection = null;
      }

      if(!currentSection || !cls) continue;
      const bu = classToBU[cls];
      if(!bu) continue;
      ensureBU(bu);

      if(currentSection === 'FTE'){
        months.forEach(mo => {
          const ci = colMap[mo.label];
          if(ci !== undefined) buData[bu].hc[mo.label] += toInt(row[ci]);
        });
      } else if(MOVEMENT_SECTIONS.has(currentSection)){
        const ci = colMap[M4.label];
        const val = ci !== undefined ? toInt(row[ci]) : 0;
        if(currentSection === 'New Hires'){
          buData[bu].new += val;
        } else if(currentSection === 'CMH - Actual Termination'){
          buData[bu].att -= Math.abs(val);
        } else if(currentSection === 'Transfer In'){
          buData[bu].trans += Math.abs(val);
        } else if(currentSection === 'Transfer Out'){
          buData[bu].trans += val;
        }
      }
    }

    // Preserve the raw latest FTE and calculate the latest month from the bridge.
    Object.values(buData).forEach(d => {
      d.preM4 = d.hc[M4.label] || 0;
      d.hc[M4.label] =
        (d.hc[M3.label] || 0) +
        (d.new || 0) +
        (d.att || 0) +
        (d.trans || 0);
    });

    const html = hcaBuildHTML({buData, months, M4, fname});
    const win = window.open('', '_blank');
    win.document.write(html);
    win.document.close();

  } catch(err){
    alert('Error generating HC Actuals: ' + err.message);
    console.error(err);
  }

  btn.innerHTML = '<i class="ti ti-report-analytics"></i>Generate Report';
  btn.disabled = false;
  btn.classList.add('enabled');
}

function hcaBuildHTML({buData, months, M4, fname}){
  const [M1,M2,M3] = months;
  const ML = months.map(m=>m.label);

  function sumGroup(members){
    const out = {hc:{[ML[0]]:0,[ML[1]]:0,[ML[2]]:0,[ML[3]]:0},new:0,att:0,trans:0,preM4:0};
    members.forEach(bu=>{
      const d = buData[bu];
      if(!d) return;
      ML.forEach(m=>out.hc[m]+=d.hc[m]||0);
      out.new+=d.new||0;
      out.att+=d.att||0;
      out.trans+=d.trans||0;
      out.preM4+=d.preM4||0;
    });
    return out;
  }

  function emptyBU(){
    return {hc:{[ML[0]]:0,[ML[1]]:0,[ML[2]]:0,[ML[3]]:0},new:0,att:0,trans:0,preM4:0};
  }

  function leafBU(bu){
    if(!buData[bu]) return emptyBU();
    const d = buData[bu];
    return {hc:{...d.hc},new:d.new||0,att:d.att||0,trans:d.trans||0,preM4:d.preM4||0};
  }

  const T = {};
  T['Engineering'] = leafBU('Engineering');
  T['Tech Solutions'] = leafBU('Tech Solutions');
  T['Lab'] = leafBU('Lab');
  T['BAP Researchers'] = leafBU('BAP Researchers');
  T['Data Products'] = leafBU('Data Products');
  T['Data Strategy'] = leafBU('Data Strategy');
  T['Product'] = leafBU('Product');
  T['LT'] = leafBU('LT');
  T['Comm Strat & Rev Ops'] = leafBU('Comm Strat & Rev Ops');
  T['Sales Engineering'] = leafBU('Sales Engineering');
  T['Customer Support'] = leafBU('Customer Support');
  T['Analytics Services'] = leafBU('Analytics Services');
  T['FAST'] = leafBU('FAST');
  T['Sales Development'] = leafBU('Sales Development');
  T['Retention'] = leafBU('Retention');
  T['Sales'] = leafBU('Sales');
  T['Finance'] = leafBU('Finance');
  T['People'] = leafBU('People');
  T['Legal'] = leafBU('Legal');
  T['Executive Operations'] = leafBU('Executive Operations');
  T['IT'] = leafBU('IT');
  T['Total Mavens'] = leafBU('Total Mavens');
  T['Total Marketing'] = leafBU('Total Marketing');

  T['Total Dev Org'] = sumGroup(['Engineering','Tech Solutions','Lab']);
  T['Total Data Product'] = sumGroup(['BAP Researchers','Data Products','Data Strategy','Product']);
  T['Total Ops'] = sumGroup(['LT','Comm Strat & Rev Ops','Sales Engineering','Customer Support','Analytics Services','FAST','Sales Development']);
  T['Total Sales + Retention'] = sumGroup(['Retention','Sales']);
  T['Total Commercial'] = sumGroup(['LT','Comm Strat & Rev Ops','Sales Engineering','Customer Support','Analytics Services','FAST','Sales Development','Retention','Sales']);
  T['Total G&A'] = sumGroup(['Finance','People','Legal','Executive Operations','IT']);

  const net = {hc:{[ML[0]]:0,[ML[1]]:0,[ML[2]]:0,[ML[3]]:0},new:0,att:0,trans:0,preM4:0};
  ['Total Dev Org','Total Data Product','Total Commercial','Total Mavens','Total Marketing','Total G&A'].forEach(k=>{
    const d=T[k];
    ML.forEach(m=>net.hc[m]+=d.hc[m]||0);
    net.new+=d.new||0;
    net.att+=d.att||0;
    net.trans+=d.trans||0;
    net.preM4+=d.preM4||0;
  });

  let rowIdx = 0;
  const rowList = [];
  const hideBtn = i => `<span class="rhide" onclick="hideRow(${i})" title="Hide this row">&times;</span>`;
  function mov(v, cls){ if(v===0) return '<span class="zero">0</span>'; return `<span class="${cls}">${v}</span>`; }
  function dRow(label, d){
    const i = rowIdx++; rowList.push({i,label,g:'detail'});
    return `<tr class="detail" data-row="${i}"><td>${hideBtn(i)}${label}</td><td>${d.hc[ML[0]]||0}</td><td>${d.hc[ML[1]]||0}</td><td>${d.hc[ML[2]]||0}</td><td>${mov(d.new||0,'col-new-val')}</td><td>${mov(d.att||0,'col-att-val')}</td><td>${mov(d.trans||0,'col-trans-val')}</td><td>${d.hc[ML[3]]||0}</td><td class="internal-only">${d.preM4||0}</td></tr>`;
  }
  function sRow(label, d, light){
    const i = rowIdx++; rowList.push({i,label,g:'subtotal'});
    const cls = light ? 'subtotal-light' : 'subtotal';
    return `<tr class="${cls}" data-row="${i}"><td>${hideBtn(i)}${label}</td><td>${d.hc[ML[0]]||0}</td><td>${d.hc[ML[1]]||0}</td><td>${d.hc[ML[2]]||0}</td><td>${mov(d.new||0,'col-new-val')}</td><td>${mov(d.att||0,'col-att-val')}</td><td>${mov(d.trans||0,'col-trans-val')}</td><td>${d.hc[ML[3]]||0}</td><td class="internal-only">${d.preM4||0}</td></tr>`;
  }
  const sp = `<tr class="spacer"><td colspan="9"></td></tr>`;

  let tbody = '';
  ['Engineering','Tech Solutions','Lab'].forEach(b=>tbody+=dRow(b,T[b]));
  tbody+=sRow('Total Dev Org',T['Total Dev Org']); tbody+=sp;
  ['BAP Researchers','Data Products','Data Strategy','Product'].forEach(b=>tbody+=dRow(b,T[b]));
  tbody+=sRow('Total Data Product',T['Total Data Product']); tbody+=sp;
  ['LT','Comm Strat & Rev Ops','Sales Engineering','Customer Support','Analytics Services','FAST','Sales Development'].forEach(b=>tbody+=dRow(b,T[b]));
  tbody+=sRow('Total Ops',T['Total Ops'],true); tbody+=sp;
  ['Retention','Sales'].forEach(b=>tbody+=dRow(b,T[b]));
  tbody+=sRow('Total Sales + Retention',T['Total Sales + Retention'],true); tbody+=sp;
  tbody+=sRow('Total Commercial',T['Total Commercial']); tbody+=sp;
  tbody+=sRow('Total Mavens',T['Total Mavens']); tbody+=sp;
  tbody+=sRow('Total Marketing',T['Total Marketing']); tbody+=sp;
  ['Finance','People','Legal','Executive Operations','IT'].forEach(b=>tbody+=dRow(b,T[b]));
  tbody+=sRow('Total G&A',T['Total G&A']); tbody+=sp;

  const netI = rowIdx++; rowList.push({i:netI,label:'Total Net FTE Headcount',g:'total'});
  tbody+=`<tr class="total-net" data-row="${netI}"><td>${hideBtn(netI)}Total Net FTE Headcount</td><td>${net.hc[ML[0]]}</td><td>${net.hc[ML[1]]}</td><td>${net.hc[ML[2]]}</td><td>${mov(net.new,'col-new-val')}</td><td>${mov(net.att,'col-att-val')}</td><td>${mov(net.trans,'col-trans-val')}</td><td>${net.hc[ML[3]]}</td><td class="internal-only">${net.preM4}</td></tr>`;
  tbody+=sp;

  const crI = rowIdx++; rowList.push({i:crI,label:'Total Headcount + Costa Rica',g:'total'});
  tbody+=`<tr class="total-cr" data-row="${crI}"><td>${hideBtn(crI)}Total Headcount + Costa Rica</td><td id="tot-m1">${net.hc[ML[0]]}</td><td id="tot-m2">${net.hc[ML[1]]}</td><td id="tot-m3">${net.hc[ML[2]]}</td><td id="tot-new">${mov(net.new,'col-new-val')}</td><td id="tot-att">${mov(net.att,'col-att-val')}</td><td id="tot-trans">${mov(net.trans,'col-trans-val')}</td><td id="tot-m4-calc">${net.hc[ML[3]]}</td><td class="internal-only" id="tot-m4-pre">${net.preM4}</td></tr>`;
  tbody+=sp;

  const rowChecklist = rowList.map(r=>`<label class="rowchk"><input type="checkbox" data-i="${r.i}" checked onchange="toggleRow(${r.i},this.checked)"> ${r.label}</label>`).join('');
  const netJSON = JSON.stringify(net);
  const outName = `HC_Actuals_${M4.name}${String(M4.yr).slice(2)}.html`;

  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>HC Actuals by Function - ${M4.name} ${String(M4.yr).slice(2)}</title><style>*{box-sizing:border-box;margin:0;padding:0;}body{font-family:'Calibri',Arial,sans-serif;background:#f0f2f5;padding:24px 32px;display:flex;flex-direction:column;align-items:center;gap:14px;}.toolbar{width:860px;display:flex;justify-content:flex-end;gap:8px;}.btn-copy{display:flex;align-items:center;gap:6px;background:#1e3a5f;color:#ffffff;border:none;border-radius:5px;padding:7px 16px;font-size:11px;font-weight:600;font-family:inherit;cursor:pointer;transition:background .15s;}.btn-copy:hover{background:#274d7a;}.btn-copy:active{background:#152b47;}.btn-copy.copied{background:#15803d;}.btn-dl{display:flex;align-items:center;gap:6px;background:#fff;color:#1e3a5f;border:1.5px solid #1e3a5f;border-radius:5px;padding:6px 14px;font-size:11px;font-weight:600;font-family:inherit;cursor:pointer;transition:background .15s;}.btn-dl:hover{background:#e6f1fb;}.row-panel{width:860px;background:#fff;border:1px solid #d1d5db;border-radius:6px;padding:12px 14px;display:none;grid-template-columns:repeat(3,1fr);gap:5px 16px;font-size:11px;}.row-panel.open{display:grid;}.rowchk{display:flex;align-items:center;gap:6px;color:#111827;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}.rowchk input{cursor:pointer;}.panel-actions{grid-column:1/-1;display:flex;justify-content:space-between;align-items:center;border-top:1px solid #eee;padding-top:8px;margin-top:4px;}.panel-title{font-weight:700;color:#1e3a5f;}.btn-reset{background:#1e3a5f;color:#fff;border:none;border-radius:4px;padding:5px 12px;font-size:11px;cursor:pointer;font-family:inherit;font-weight:600;}.btn-reset:hover{background:#274d7a;}.table-wrapper{width:860px;background:#ffffff;}table{width:860px;border-collapse:collapse;font-size:10px;}thead th{background:#1e3a5f;color:#ffffff;font-weight:700;font-size:9.5px;padding:5px 8px;text-align:center;white-space:nowrap;border:1px solid #374151;}thead th:first-child{text-align:left;width:168px;}thead th.col-new,thead th.col-att,thead th.col-trans{background:#d1d5db;color:#111827;}th.internal-only,td.internal-only{background:#fff7ed!important;color:#92400e!important;}tr.detail td{padding:2px 8px;text-align:center;border:1px solid #e5e7eb;color:#111827;background:#ffffff;font-size:10px;}tr.detail td:first-child{text-align:left;padding-left:14px;}tr.detail:hover td{background:#f9fafb;}tr.subtotal td{padding:3px 8px;text-align:center;background:#c8c8c8;font-weight:700;font-size:10px;color:#111827;border:1px solid #9ca3af;}tr.subtotal td:first-child{text-align:left;padding-left:6px;}tr.subtotal-light td{padding:3px 8px;text-align:center;background:#e9eaeb;font-weight:700;font-size:10px;color:#111827;border:1px solid #c8c8c8;}tr.subtotal-light td:first-child{text-align:left;padding-left:6px;}tr.total-net td,tr.total-cr td{padding:4px 8px;text-align:center;background:#ffffff;color:#111827;font-weight:700;font-size:10px;border-top:2px solid #1e3a5f;border-bottom:2px solid #1e3a5f;border-left:none;border-right:none;}tr.total-net td:first-child,tr.total-cr td:first-child{text-align:left;padding-left:6px;border-left:2px solid #1e3a5f;}tr.total-net td:last-child,tr.total-cr td:last-child{border-right:2px solid #1e3a5f;}tr.cr-edit td{padding:2px 8px;text-align:center;border:1px solid #e5e7eb;background:#f8fafc;font-size:10px;}tr.cr-edit td:first-child{text-align:left;padding-left:14px;color:#6b7280;font-style:italic;}tr.spacer td{padding:2px 0;background:#f3f4f6;border:none;}.col-new-val{color:#15803d;font-weight:600;}.col-att-val{color:#b91c1c;font-weight:600;}.col-trans-val{color:#1d4ed8;font-weight:600;}.zero{color:#d1d5db;}.rhide{display:inline-block;width:11px;margin-right:5px;color:#b91c1c;cursor:pointer;opacity:0;font-weight:700;transition:opacity .1s;}tr:hover .rhide{opacity:1;}input.cr-in{width:100%;text-align:center;background:transparent;border:none;border-bottom:1px dashed #9ca3af;font-size:10px;font-weight:600;font-family:inherit;outline:none;padding:1px 0;color:#111827;}input.cr-in:focus{border-bottom-color:#374151;}input.cr-in.inp-new{color:#15803d;}input.cr-in.inp-att{color:#b91c1c;}input.cr-in.inp-trans{color:#1d4ed8;}</style></head><body><div class="toolbar"><button class="btn-dl" onclick="toggleRowPanel()"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/></svg>Show / Hide Rows</button><button class="btn-dl" onclick="downloadReport()"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>Download HTML</button><button class="btn-copy" id="btn-copy" onclick="copyAsImage()"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>Copy as Image</button></div><div class="row-panel" id="row-panel">${rowChecklist}<div class="panel-actions"><span class="panel-title">Toggle rows to display</span><button class="btn-reset" onclick="resetRows()">Show all</button></div></div><div class="table-wrapper"><table id="capture"><thead><tr><th>Function</th><th>${ML[0]}</th><th>${ML[1]}</th><th>${ML[2]}</th><th class="col-new">(+) New</th><th class="col-att">(-) Attrition</th><th class="col-trans">(+/-) Transfers</th><th data-calculated-header>${ML[3]} (Calculated)</th><th class="internal-only">${ML[3]} (Pre-Adjustment)</th></tr></thead><tbody id="tbody">${tbody}</tbody></table></div><div class="table-wrapper"><table style="width:860px;border-collapse:collapse;font-size:10px;"><tbody id="cr-tbody"></tbody></table></div><script>const NET=${netJSON};const ML=${JSON.stringify(ML)};document.getElementById('cr-tbody').innerHTML='<tr class="cr-edit"><td>Costa Rica</td><td><input class="cr-in" id="cr-m1" type="number" value="0"></td><td><input class="cr-in" id="cr-m2" type="number" value="0"></td><td><input class="cr-in" id="cr-m3" type="number" value="0"></td><td><input class="cr-in inp-new" id="cr-new" type="number" value="0"></td><td><input class="cr-in inp-att" id="cr-att" type="number" value="0"></td><td><input class="cr-in inp-trans" id="cr-trans" type="number" value="0"></td><td><input class="cr-in" id="cr-m4" type="number" value="0"></td><td class="internal-only"></td></tr>';['cr-m1','cr-m2','cr-m3','cr-m4','cr-new','cr-att','cr-trans'].forEach(id=>{ document.getElementById(id).addEventListener('input',update); });function mov(v,cls){ if(v===0) return '<span class="zero">0</span>'; return '<span class="'+cls+'">'+v+'</span>'; }function update(){const g=id=>parseInt(document.getElementById(id).value)||0;document.getElementById('tot-m1').textContent=NET.hc[ML[0]]+g('cr-m1');document.getElementById('tot-m2').textContent=NET.hc[ML[1]]+g('cr-m2');document.getElementById('tot-m3').textContent=NET.hc[ML[2]]+g('cr-m3');document.getElementById('tot-m4-calc').textContent=NET.hc[ML[3]]+g('cr-m4');document.getElementById('tot-new').innerHTML=mov(NET.new+g('cr-new'),'col-new-val');document.getElementById('tot-att').innerHTML=mov(NET.att+g('cr-att'),'col-att-val');document.getElementById('tot-trans').innerHTML=mov(NET.trans+g('cr-trans'),'col-trans-val');}function toggleRowPanel(){document.getElementById('row-panel').classList.toggle('open');}function toggleRow(i,show){const tr=document.querySelector('tr[data-row="'+i+'"]');if(tr)tr.style.display=show?'':'none';}function hideRow(i){toggleRow(i,false);const cb=document.querySelector('#row-panel input[data-i="'+i+'"]');if(cb)cb.checked=false;}function resetRows(){document.querySelectorAll('tr[data-row]').forEach(tr=>{tr.style.display='';});document.querySelectorAll('#row-panel input[type=checkbox]').forEach(cb=>{cb.checked=true;});}function downloadReport(){ const h='<!DOCTYPE html>'+document.documentElement.outerHTML; const a=document.createElement('a'); a.href='data:text/html;charset=utf-8,'+encodeURIComponent(h); a.download='${outName}'; a.click(); }async function copyAsImage(){ const btn=document.getElementById('btn-copy'); btn.textContent='Capturing...'; btn.disabled=true; let holder=null; try{ const clone=document.getElementById('capture').cloneNode(true); clone.id='capture-copy'; clone.querySelectorAll('.internal-only').forEach(el=>el.remove()); const header=clone.querySelector('[data-calculated-header]'); if(header) header.textContent=ML[3]; holder=document.createElement('div'); holder.style.position='fixed'; holder.style.left='-100000px'; holder.style.top='0'; holder.style.background='#fff'; holder.appendChild(clone); document.body.appendChild(holder); const canvas=await html2canvas(clone,{scale:2,backgroundColor:'#ffffff',useCORS:true,logging:false}); holder.remove(); holder=null; canvas.toBlob(async blob=>{ try{ await navigator.clipboard.write([new ClipboardItem({'image/png':blob})]); btn.innerHTML='Copied!'; btn.classList.add('copied'); setTimeout(()=>{ btn.innerHTML='Copy as Image'; btn.classList.remove('copied'); btn.disabled=false; },2000);}catch(e){ alert('Could not copy. Try right-clicking the table.'); btn.textContent='Copy as Image'; btn.disabled=false; } },'image/png'); }catch(e){ if(holder)holder.remove(); alert('Capture failed: '+e.message); btn.textContent='Copy as Image'; btn.disabled=false; }}<\/script><script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"><\/script></body></html>`;
}

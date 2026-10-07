/* ============================================================
   FP&A Analyst Bot — rule-based Q&A for the BvA board.
   - No AI model, no network calls: intent rules + the board's own data.
   - Questions in English or Spanish; answers in English.
   - Fully self-contained: app.js serializes this function with
     toString() into every exported board, so it must not reference
     anything outside its own body (DOM APIs only).
   ============================================================ */
function initAnalystBot(DATA){
  var d = document;
  ['fpa-bot-root','fpa-bot-style'].forEach(function(id){ var e = d.getElementById(id); if(e && e.parentNode) e.parentNode.removeChild(e); });
  if(!DATA || !DATA.rows || !DATA.rows.length) return null;

  /* ---------------- meta & periods ---------------- */
  var META = DATA.meta || {};
  var MONTHS = (DATA.monthLabels && DATA.monthLabels.length === 12) ? DATA.monthLabels : ['Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan'];
  var QL = (DATA.quarterLabels && DATA.quarterLabels.length === 4) ? DATA.quarterLabels : ['Q1','Q2','Q3','Q4'];
  var RM = (typeof META.monthIdx === 'number' && META.monthIdx >= 0 && META.monthIdx < 12) ? META.monthIdx : 0;
  var RQ = Math.floor(RM / 3);
  var PLAN = META.planLabel || 'Plan';
  var FCST = META.fcstLabel || 'Forecast';
  var BU = META.code || 'BU';
  var FY = META.fy || 'FY';
  var CAL = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  var startCal = CAL.indexOf(String(MONTHS[0]).slice(0,3).toLowerCase());
  if(startCal < 0) startCal = 1;
  function fiscalOf(cal){ return (cal - startCal + 12) % 12; }
  var THR = { m:10000, q:25000, ytd:50000, fy:75000 };

  /* ---------------- formatting ---------------- */
  function esc(s){ return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function norm(s){
    return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'')
      .replace(/&/g,' and ').replace(/[^a-z0-9+#\s]/g,' ').replace(/\s+/g,' ').trim();
  }
  function fm(n){
    n = Number(n) || 0; var a = Math.abs(n), s;
    if(a >= 1e6) s = '$' + (a/1e6).toFixed(2) + 'M';
    else if(a >= 1e3) s = '$' + (a/1e3).toFixed(1) + 'K';
    else s = '$' + Math.round(a).toLocaleString('en-US');
    return (n < 0 && Math.round(a) !== 0 ? '-' : '') + s;
  }
  function fv(n){ n = Number(n) || 0; if(Math.abs(n) < 0.5) return '$0'; return (n > 0 ? '+' : '-') + fm(Math.abs(n)); }
  function pct(v, b){ if(!b || Math.abs(b) < 1) return null; var p = v / Math.abs(b) * 100; return (p > 0 ? '+' : '') + p.toFixed(1) + '%'; }
  function pctPlain(a, b){ if(!b || Math.abs(b) < 1) return '—'; return (a / b * 100).toFixed(0) + '%'; }
  function cls(v, thr){ if(Math.abs(v) < (thr || 1)) return 'fb-neu'; return v > 0 ? 'fb-bad' : 'fb-good'; }
  function verdict(v, thr){ if(Math.abs(v) < (thr || 1)) return 'in line'; return v > 0 ? 'unfavorable' : 'favorable'; }
  function tag(v, thr){ return '<span class="fb-tag ' + cls(v, thr) + '">' + verdict(v, thr) + '</span>'; }
  function vspan(v, thr){ return '<span class="' + cls(v, thr) + '">' + fv(v) + '</span>'; }
  function bName(bk){ return bk === 'f' ? FCST : PLAN; }
  function pLabel(P){
    if(!P) return '';
    if(P.t === 'm') return MONTHS[P.i];
    if(P.t === 'q') return QL[P.i];
    if(P.t === 'ytd') return 'YTD (' + MONTHS[0] + ' – ' + MONTHS[RM] + ')';
    return FY + ' full year';
  }
  function samePeriod(a, b){ return a && b && a.t === b.t && (a.i || 0) === (b.i || 0); }
  function thrOf(P, scale){ return (THR[P.t] || 25000) * (scale || 1); }

  /* ---------------- values ---------------- */
  function rowVal(r, P){
    var a;
    if(P.t === 'm') a = r.M && r.M[P.i];
    else if(P.t === 'q') a = r.Q && r.Q[P.i];
    else if(P.t === 'fy') a = r.T;
    else {
      a = [0,0,0];
      for(var k = 0; k <= RM; k++){ var mm = r.M && r.M[k]; if(mm) a = [a[0]+mm[0], a[1]+mm[1], a[2]+mm[2]]; }
    }
    a = a || [0,0,0];
    return { w:+a[0] || 0, p:+a[1] || 0, f:+a[2] || 0 };
  }
  function sumVals(rows, P){
    var o = { w:0, p:0, f:0 };
    rows.forEach(function(r){ var x = rowVal(r, P); o.w += x.w; o.p += x.p; o.f += x.f; });
    return o;
  }
  function withB(o, bk){ o.b = bk === 'f' ? o.f : o.p; o.v = o.w - o.b; o.vp = o.w - o.p; o.vf = o.w - o.f; return o; }
  function ev(E, P, bk){ return withB(sumVals(E.rows, P), bk); }

  /* ---------------- index: categories, accounts, vendors ---------------- */
  function splitCode(label){
    var m = String(label).match(/^\s*(\d{3,})\s*-\s*(.*)$/);
    return m ? { code:m[1], name:m[2].trim() } : { code:'', name:String(label).replace(/\s*-\s*$/,'').trim() };
  }
  var CATS = [], GLS = [], VEND = {}, VLIST = [], VROWS = [], TOTAL = null, pendV = [], pendG = [];
  DATA.rows.forEach(function(r){
    if(r.t === 'vendor' || r.t === 'novendor'){ pendV.push(r); return; }
    if(r.t === 'gl'){
      var sc = splitCode(r.l);
      var g = { type:'gl', key:'gl:' + sc.code, label:r.l, name:sc.name, code:sc.code, rows:[r], vrows:pendV.slice() };
      GLS.push(g); pendG.push(g); pendV = []; return;
    }
    if(r.t === 'l2'){
      var c = { type:'cat', key:'cat:' + norm(r.l), label:r.l, name:String(r.l).replace(/^Total\s+/i,'').replace(/\s+/g,' ').trim(), code:'', rows:[r], gls:pendG.slice() };
      pendG.forEach(function(gg){ gg.cat = c; });
      CATS.push(c); pendG = []; pendV = []; return;
    }
    if(r.t === 'expense'){ TOTAL = { type:'total', key:'total', label:r.l, name:BU + ' total expense', code:'', rows:[r] }; }
  });
  if(!TOTAL) TOTAL = { type:'total', key:'total', label:'Expense', name:BU + ' total expense', code:'', rows:CATS.map(function(c){ return c.rows[0]; }) };
  GLS.forEach(function(g){
    g.vrows.forEach(function(r){
      var sc = splitCode(r.l), nov = r.t === 'novendor' || /^no vendor/i.test(r.l);
      var name = nov ? 'No Vendor' : sc.name;
      var key = 'v:' + (nov ? 'novendor' : (sc.code || norm(name)));
      var V = VEND[key];
      if(!V){ V = VEND[key] = { type:'vendor', key:key, label:r.l, name:name, code:nov ? '' : sc.code, rows:[], pairs:[], nov:nov }; VLIST.push(V); }
      V.rows.push(r); V.pairs.push({ row:r, gl:g });
      VROWS.push({ row:r, gl:g, cat:g.cat || null, vendor:V });
    });
  });

  /* ---------------- index: people (HC + T&E) ---------------- */
  var HC = DATA.hc || null, TE = DATA.te || null, EMPS = {}, ELIST = [];
  function getEmp(raw){
    raw = String(raw || '').trim();
    if(!raw || /^no employee/i.test(raw) || /^total$/i.test(raw)) return null;
    var tbh = /^TBH\b/i.test(raw);
    var name = tbh ? raw : raw.replace(/\s+-\s+\S+$/,'').trim();
    var key = 'e:' + norm(name), E = EMPS[key];
    if(!E){ E = EMPS[key] = { type:'employee', key:key, label:raw, name:name, code:'', tbh:tbh, hc:null, te:[], ids:[] }; ELIST.push(E); }
    var idm = raw.match(/\s-\s(\S+)\s*$/); if(idm) E.ids.push(norm(idm[1]).replace(/\s+/g,''));
    return E;
  }
  if(HC && HC.employees) HC.employees.forEach(function(e){ var E = getEmp(e.n); if(E) E.hc = e; });
  if(TE && TE.rows) TE.rows.forEach(function(t){ var E = getEmp(t.e); if(E) E.te.push(t); });

  /* ---------------- entity matching ---------------- */
  var NAME_SKIP = { tbh:1, inc:1, llc:1, ltd:1, limited:1, corp:1, co:1, the:1, and:1, of:1, for:1, total:1, private:1, a:1 };
  var NAME_LOW = { services:1, service:1, expensed:1, expense:1, exp:1, other:1, fees:1, online:1, non:1, client:1 };
  var QSTOP = ('what whats which who whom where when why how much many is are was were be been being the a an of in on at for to from by '
    + 'vs versus against compared compare with and or our we us i me my you your it its this that these those there here do does did done doing going '
    + 'can could would should will please show give tell list about explain describe see view display get find look up down so any all some '
    + 'spend spent spending cost costs expense expenses variance variances var plan planned budget budgeted forecast fcst fc working actual actuals '
    + 'top biggest largest main major key drivers driver vendor vendors supplier suppliers account accounts gl gls category categories line lines '
    + 'month months monthly quarter quarters quarterly year years yearly annual full fy ytd qtd mtd date current last previous next '
    + 'favorable unfavorable fav unfav over under above below summary overview trend utilization remaining left status '
    + 'que cual cuales cuanto cuanta cuantos cuantas como de del la el los las lo en por para con contra y o u un una unos unas al es son fue esta este estos estas esa ese '
    + 'gasto gastos gastamos gastado presupuesto pronostico varianza variacion mes meses trimestre trimestres ano anos anual proveedor proveedores cuenta cuentas categoria categorias '
    + 'mayor mayores principales principal resumen tendencia dame muestra muestrame dime explica sobre porque paso pasa pasando vamos va estamos hay tiene tenemos me nos se su sus '
    + 'buffer buffers parked park reallocated real projected projection proyectado closed excluding underlying savings ahorro ahorros where donde '
    + 'driving drove drive cause causing reason happened happening going doing tracking running look trending breakdown detail details '
    + 'mas employee employees empleado empleados person people persona personas who quien same now ahora mismo k m usd dollars much number numbers amount amounts value values total totals overall entire whole').split(' ');
  var QSTOPSET = {}; QSTOP.forEach(function(w){ QSTOPSET[w] = 1; });
  var MONTHWORDS = { jan:0, january:0, ene:0, enero:0, feb:1, february:1, febrero:1, mar:2, march:2, marzo:2, apr:3, april:3, abr:3, abril:3,
    may:4, mayo:4, jun:5, june:5, junio:5, jul:6, july:6, julio:6, aug:7, august:7, ago:7, agosto:7, sep:8, sept:8, september:8, septiembre:8, setiembre:8,
    oct:9, october:9, octubre:9, nov:10, november:10, noviembre:10, dec:11, december:11, dic:11, diciembre:11 };

  function prep(E){
    var toks = [];
    norm(E.name).split(' ').forEach(function(t){
      if(t.length < 2 || NAME_SKIP[t]) return;
      if(toks.some(function(x){ return x.t === t; })) return;
      toks.push({ t:t, w:NAME_LOW[t] ? 0.35 : 1 });
    });
    E._toks = toks;
    E._full = norm(E.name);
    E._codes = [];
    if(E.code) E._codes.push(E.code);
    (E.ids || []).forEach(function(x){ if(x) E._codes.push(x); });
  }
  function lev(a, b){
    if(a === b) return 0;
    var m = a.length, n = b.length, prev = [], cur, i, j;
    for(j = 0; j <= n; j++) prev[j] = j;
    for(i = 1; i <= m; i++){
      cur = [i];
      for(j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j-1] + 1, prev[j-1] + (a[i-1] === b[j-1] ? 0 : 1));
      prev = cur;
    }
    return prev[n];
  }
  function tokScore(qt, et){
    if(qt === et) return 1;
    if(qt.length >= 3 && et.length > qt.length && et.indexOf(qt) === 0) return qt.length >= 4 ? 0.85 : 0.7;
    if(qt.length >= 5 && et.length >= 5 && Math.abs(qt.length - et.length) <= 2 && lev(qt, et) <= (et.length >= 8 ? 2 : 1)) return 0.75;
    return 0;
  }
  function scoreEntity(E, qts, qn){
    if(!E._toks) prep(E);
    var tot = 0, got = 0;
    E._toks.forEach(function(t){
      tot += t.w;
      var best = 0;
      qts.forEach(function(qt){ var s = tokScore(qt, t.t); if(s > best) best = s; });
      got += best * t.w;
    });
    var codeHit = E._codes.some(function(c){ return qts.indexOf(c) >= 0; });
    if(!got && !codeHit) return 0;
    var s = got + (tot ? got / tot : 0) * 0.5 + (codeHit ? 10 : 0);
    if(E._full.length >= 4 && (' ' + qn + ' ').indexOf(' ' + E._full + ' ') >= 0) s += 3;
    return s;
  }
  /* ---------------- index: No Vendor line items (optional NO VENDOR feed) ----------------
     DATA.nv.gls[] = { code, label, cat, board:{w,p,f}[12] (No Vendor rows of the Year feed) | null,
                       items:[{ d, v, nv, w[12], p[12], f[12], dt }] }
     Working line items exist only after the review month (closed months are actuals without
     line detail); Forecast line items start at DATA.nv.fStart.                                 */
  var NV = (DATA.nv && DATA.nv.gls && DATA.nv.gls.length) ? DATA.nv : null;
  var NVFS = NV && NV.fStart !== null && NV.fStart !== undefined ? NV.fStart : 99;
  var NVL = [];
  function nvG(code){ if(!NV) return null; for(var i = 0; i < NV.gls.length; i++){ if(NV.gls[i].code === code) return NV.gls[i]; } return null; }
  function nvSum(a, ms){ var s = 0; ms.forEach(function(m){ s += (a && a[m]) || 0; }); return s; }
  if(NV) NV.gls.forEach(function(g){
    g.items.forEach(function(it, i){ NVL.push({ type:'nvline', key:'nl:' + g.code + ':' + i, label:it.d, name:it.d, code:'', it:it, g:g }); });
  });
  // Same split as the board: projected months get line-level Working vs benchmark,
  // closed months only have the actual No Vendor total next to the planned lines.
  function nvSplit(g, months, bk){
    var proj = months.filter(function(m){ return m > RM; }), closed = months.filter(function(m){ return m <= RM; });
    function det(ms){ return bk === 'f' ? ms.filter(function(m){ return m >= NVFS; }) : ms; }
    var bd = g.board || { w:[], p:[], f:[] };
    var o = { proj:proj, closed:closed, rows:[], plan:[], detW:0, detB:0,
      boardW:nvSum(bd.w, proj), boardB:nvSum(bd[bk], proj), cW:nvSum(bd.w, closed), cB:nvSum(bd[bk], closed) };
    g.items.forEach(function(it){
      if(!it.nv) return;
      var w = nvSum(it.w, proj), b = nvSum(it[bk], det(proj));
      o.detW += w; o.detB += b;
      if(Math.abs(w - b) >= 1) o.rows.push({ it:it, w:w, b:b, v:w - b });   // only lines that cause variance
      var cb = nvSum(it[bk], det(closed));
      if(Math.abs(cb) >= 0.5) o.plan.push({ it:it, b:cb });
    });
    o.rows.sort(function(a, b){ return Math.abs(b.v) - Math.abs(a.v); });
    o.plan.sort(function(a, b){ return Math.abs(b.b) - Math.abs(a.b); });
    o.unW = o.boardW - o.detW; o.unB = o.boardB - o.detB; o.unV = o.unW - o.unB;
    o.has = o.rows.length > 0 || (!!g.board && Math.abs(o.unV) >= 1) || (!!g.board && o.closed.length > 0 && Math.abs(o.cW - o.cB) >= 1);
    return o;
  }
  // Names of the line items that move a GL away from its benchmark in one month (sign > 0: above).
  function nvNames(code, m, bk, sign){
    var g = nvG(code); if(!g || m <= RM) return [];
    return g.items.filter(function(it){ return it.nv; })
      .map(function(it){ return { d:it.d, x:((it.w && it.w[m]) || 0) - (m >= NVFS || bk !== 'f' ? ((it[bk] && it[bk][m]) || 0) : 0) }; })
      .filter(function(r){ return sign > 0 ? r.x >= 100 : r.x <= -100; })
      .sort(function(a, b){ return Math.abs(b.x) - Math.abs(a.x); })
      .map(function(r){ return r.d; });
  }
  function nvNamesList(list, bk, sign){
    var seen = {}, out = [];
    list.forEach(function(e){ nvNames(e.gl.code, e.m, bk, sign).forEach(function(n){ if(!seen[n]){ seen[n] = 1; out.push(n); } }); });
    return out;
  }
  var ALL = [].concat(CATS, GLS, VLIST.filter(function(v){ return !v.nov; }), ELIST, NVL);
  var TYPE_BONUS = { cat:0.3, gl:0.2, vendor:0.1, employee:0.05, nvline:0 };
  function queryTokens(qn){
    return qn.split(' ').filter(function(t){
      if(!t || QSTOPSET[t] || MONTHWORDS.hasOwnProperty(t)) return false;
      if(/^q[1-4]$/.test(t) || /^fy\d*$/.test(t) || /^\d{1,2}$/.test(t) || /^20\d\d$/.test(t)) return false;
      return t.length >= 2;
    });
  }
  function findEntities(qn, hint, types){
    var qts = queryTokens(qn), res = [];
    if(/\bno vendor\b|\bsin proveedor\b/.test(qn) && VEND['v:novendor']) res.push({ E:VEND['v:novendor'], s:99 });
    if(!qts.length) return res;
    ALL.forEach(function(E){
      if(types && types.indexOf(E.type) < 0) return;
      var s = scoreEntity(E, qts, qn);
      if(s < 0.95) return;
      if(hint && E.type === hint) s += 0.6;
      s += TYPE_BONUS[E.type] || 0;
      res.push({ E:E, s:s });
    });
    res.sort(function(a, b){ return b.s - a.s; });
    return res;
  }
  function typeHint(qn){
    if(/\b(vendor|vendors|supplier|suppliers|proveedor|proveedores)\b/.test(qn)) return 'vendor';
    if(/\b(account|accounts|gl|gls|cuenta|cuentas)\b/.test(qn)) return 'gl';
    if(/\b(category|categories|categoria|categorias|l2|bucket)\b/.test(qn)) return 'cat';
    if(/\b(employee|empleado|who|quien|person|persona)\b/.test(qn)) return 'employee';
    return null;
  }
  function typeLabel(E){ if(E.type === 'nvline') return E.it.nv ? 'No Vendor line item' : 'Line item · ' + E.it.v; return ({ cat:'Category (L2)', gl:'GL account', vendor:'Vendor', employee:'Employee', total:'Total' })[E.type] || ''; }

  /* ---------------- parsing: period, benchmark, direction ---------------- */
  function parsePeriod(qn){
    if(/\b(ytd|year to date|a la fecha|acumulad[oa]|en lo que va|so far this year)\b/.test(qn)) return { t:'ytd' };
    var m = qn.match(/\bq\s?([1-4])\b/);
    if(m) return { t:'q', i:+m[1] - 1 };
    var ORD = { first:0, primer:0, primero:0, '1st':0, second:1, segundo:1, '2nd':1, third:2, tercer:2, tercero:2, '3rd':2, fourth:3, cuarto:3, '4th':3 };
    m = qn.match(/\b(first|second|third|fourth|1st|2nd|3rd|4th|primer|primero|segundo|tercer|tercero|cuarto)\s+(quarter|trimestre)\b/);
    if(m) return { t:'q', i:ORD[m[1]] };
    if(/\b(last|previous|prior)\s+quarter\b|\btrimestre (pasado|anterior)\b/.test(qn)) return { t:'q', i:Math.max(0, RQ - 1) };
    if(/\bnext quarter\b|\bproximo trimestre\b|\bsiguiente trimestre\b/.test(qn)) return { t:'q', i:Math.min(3, RQ + 1) };
    if(/\b(last|previous|prior)\s+month\b|\bmes (pasado|anterior)\b/.test(qn)) return { t:'m', i:Math.max(0, RM - 1) };
    var toks = qn.split(' ');
    for(var i = 0; i < toks.length; i++){
      if(MONTHWORDS.hasOwnProperty(toks[i])){
        if(toks[i] === 'may' && i + 1 < toks.length && /^(be|have|need|want|i|we)$/.test(toks[i+1])) continue;
        return { t:'m', i:fiscalOf(MONTHWORDS[toks[i]]) };
      }
    }
    if(/\b(this|current|el|este)\s+(month|mes)\b|\bmtd\b|\bmes actual\b|\bmonth\b|\bmes\b/.test(qn)) return { t:'m', i:RM };
    if(/\b(this|current|este)\s+(quarter|trimestre)\b|\bqtd\b|\bquarter\b|\btrimestre\b|\bquarterly\b|\btrimestral\b/.test(qn)) return { t:'q', i:RQ };
    if(/\b(full year|fy\s?\d*|annual|year|anual|ano|whole year|todo el ano|ano completo|fiscal year)\b/.test(qn)) return { t:'fy' };
    return null;
  }
  function parseBk(qn){
    var f = /\b(forecast|fcst|fc|6\+6|pronostico|proyeccion|reforecast|rf|outlook)\b/.test(qn);
    var p = /\b(plan|budget|presupuesto|aop|planned|presupuestado|budgeted)\b/.test(qn);
    return f && p ? 'both' : f ? 'f' : p ? 'p' : null;
  }
  function parseDir(qn){
    if(/unfav|desfavorab|over ?spend|over budget|over plan|overrun|above (plan|budget|forecast)|sobregir|exceso|excedid|por encima|worst|peor|peores|overspent|over-spent|\bover\b/.test(qn)) return 'unfav';
    if(/(^|\s)fav\b|(^|\s)favorab|saving|ahorro|under ?spend|under budget|under plan|below (plan|budget|forecast)|por debajo|\bbest\b|mejor|mejores|underspent|\bunder\b/.test(qn)) return 'fav';
    return null;
  }
  function parseN(qn){
    var m = qn.match(/\btop\s*(\d{1,2})\b/) || qn.match(/\b(\d{1,2})\s+(biggest|largest|top|main|principales|mayores|vendors|proveedores|accounts|cuentas|categories|categorias)\b/);
    var n = m ? +m[1] : 5;
    return Math.max(1, Math.min(n, 15));
  }
  function parseDim(qn){
    if(/\b(vendor|vendors|supplier|suppliers|proveedor|proveedores)\b/.test(qn)) return 'vendor';
    if(/\b(account|accounts|gl|gls|cuenta|cuentas)\b/.test(qn)) return 'gl';
    if(/\b(category|categories|categoria|categorias|l2|bucket|buckets)\b/.test(qn)) return 'cat';
    return null;
  }

  /* ---------------- board notes (comments typed on the board) ---------------- */
  var SEC = {
    'sec-qplan': QL[RQ] + ' vs Plan', 'sec-qfcst': QL[RQ] + ' vs Forecast',
    'sec-fyplan': 'Full year vs Plan', 'sec-fyfcst': 'Full year vs Forecast', 'sec-hc': 'Headcount'
  };
  var AUTO = /Validate timing, scope, and whether the run-rate|Variance appears pooled inside unattributed detail rows|^Open req not filled yet|^TBH line is now flowing|^Working cost is showing without plan budget|^Working is (below|above) plan, likely driven|^No material variance versus plan/;
  function isHidden(el){ return !!(el && el.closest && el.closest('.hidden')); }
  function boardNotes(){
    var out = [];
    d.querySelectorAll('.drv-block').forEach(function(blk){
      if(isHidden(blk)) return;
      var sec = blk.closest('section'), sid = sec ? sec.id : '';
      var lab = blk.querySelector('.drv-label'); lab = lab ? lab.textContent.trim() : '';
      blk.querySelectorAll('textarea').forEach(function(t){
        var row = t.closest('.vrow, .comment-row');
        if(row && isHidden(row)) return;
        var txt = String(t.value || t.textContent || '').trim();
        if(!txt || AUTO.test(txt)) return;
        var p = row ? row.querySelector('p') : null;
        out.push({ sid:sid, label:lab, who:p ? p.textContent.trim() : '', text:txt });
      });
    });
    d.querySelectorAll('[data-comment-block]').forEach(function(box){
      if(isHidden(box)) return;
      var sec = box.closest('section'), sid = sec ? sec.id : '';
      box.querySelectorAll('textarea').forEach(function(t){
        var row = t.closest('.comment-row'); if(row && isHidden(row)) return;
        var txt = String(t.value || t.textContent || '').trim();
        if(txt) out.push({ sid:sid, label:'', who:'', text:txt });
      });
    });
    d.querySelectorAll('.hcrow').forEach(function(row){
      if(isHidden(row)) return;
      var t = row.querySelector('textarea'), p = row.querySelector('p');
      var txt = t ? String(t.value || t.textContent || '').trim() : '';
      if(txt && !AUTO.test(txt)) out.push({ sid:'sec-hc', label:'', who:p ? p.textContent.trim() : '', text:txt });
    });
    return out;
  }
  function notesFor(E){
    var nm = norm(E.name);
    return boardNotes().filter(function(n){
      var lab = norm(n.label), txt = norm(n.text), who = norm(n.who);
      if(E.type === 'cat') return lab === nm;
      if(E.type === 'total') return !n.label && !n.who && n.sid !== 'sec-hc';
      if(E.type === 'vendor' || E.type === 'gl') return (who && who.indexOf(nm) >= 0) || (nm.length >= 4 && (' ' + txt + ' ').indexOf(' ' + nm + ' ') >= 0);
      if(E.type === 'employee'){
        if(who && who.indexOf(nm) >= 0) return true;
        var parts = nm.split(' ').filter(function(p){ return p.length >= 3; });
        return parts.length >= 2 && parts.every(function(p){ return (' ' + txt + ' ').indexOf(' ' + p + ' ') >= 0; });
      }
      return false;
    });
  }
  function notesHtml(list){
    if(!list.length) return '';
    return '<div class="fb-sub">Analyst notes on the board</div>' + list.slice(0, 4).map(function(n){
      return '<div class="fb-note"><span>' + esc(SEC[n.sid] || 'Board') + (n.label ? ' · ' + esc(n.label) : '') + (n.who && !/^no vendor/i.test(n.who) ? ' · ' + esc(n.who) : '') + '</span>' + esc(n.text) + '</div>';
    }).join('');
  }

  /* ---------------- building blocks ---------------- */
  function scaleOf(E){ return (E.type === 'vendor' || E.type === 'gl') ? 0.4 : 1; }
  function glanceTable(E, hiP){
    var Ps = [{ t:'m', i:RM }, { t:'q', i:RQ }, { t:'ytd' }, { t:'fy' }];
    if(hiP && !Ps.some(function(P){ return samePeriod(P, hiP); })) Ps.unshift(hiP);
    var sc = scaleOf(E);
    var h = '<table class="fb-t fb-tw"><thead><tr><th></th><th>Working</th><th>Plan</th><th>vs Plan</th><th>vs FCST</th></tr></thead><tbody>';
    Ps.forEach(function(P){
      var x = ev(E, P, 'p'), t = thrOf(P, sc);
      h += '<tr' + (samePeriod(P, hiP) ? ' class="fb-hl"' : '') + '><td>' + esc(pLabel(P)) + '</td><td>' + fm(x.w) + '</td><td>' + fm(x.p) + '</td><td>' + vspan(x.vp, t) + '</td><td>' + vspan(x.vf, t) + '</td></tr>';
    });
    return h + '</tbody></table>';
  }
  function headline(E, P, bk){
    var x = ev(E, P, bk), t = thrOf(P, scaleOf(E)), nm = '<b>' + esc(E.name) + '</b>', pl = esc(pLabel(P)), bn = esc(bName(bk));
    var tg = Math.abs(x.v) >= t ? ' ' + tag(x.v, t) : '.';
    if(Math.abs(x.b) < 1 && Math.abs(x.w) >= 1) return nm + ' has <b>' + fm(x.w) + '</b> of Working in ' + pl + ' with <b>no ' + bn + '</b>: the full amount is unplanned' + tg;
    if(Math.abs(x.w) < 1 && Math.abs(x.b) >= 1) return nm + ' shows <b>no Working spend</b> in ' + pl + ' against ' + fm(x.b) + ' of ' + bn + ', so that budget was not used' + tg;
    if(Math.abs(x.w) < 1 && Math.abs(x.b) < 1) return nm + ' has no Working or ' + bn + ' in ' + pl + '.';
    var dir = Math.abs(x.v) < t ? 'tracking close to' : (x.v > 0 ? 'running over' : 'running under');
    return nm + ' is ' + dir + ' ' + bn + ' in ' + pl + ': Working <b>' + fm(x.w) + '</b> vs ' + fm(x.b) + ' → <b>' + vspan(x.v, t) + '</b>' + (pct(x.v, x.b) ? ' (' + pct(x.v, x.b) + ')' : '') + ' ' + tag(x.v, t);
  }
  function kids(E, P, bk){
    var out = [];
    function push(name, sub, rows, ref){ var x = withB(sumVals(rows, P), bk); out.push({ name:name, sub:sub, v:x, E:ref || null }); }
    if(E.type === 'total') CATS.forEach(function(c){ push(c.name, '', c.rows, c); });
    else if(E.type === 'cat') E.gls.forEach(function(g){ push(g.name, g.code, g.rows, g); });
    else if(E.type === 'gl') E.vrows.forEach(function(r){ var sc = splitCode(r.l); push(r.t === 'novendor' ? 'No Vendor' : sc.name, sc.code, [r]); });
    else if(E.type === 'vendor') E.pairs.forEach(function(p){ push(p.gl.name, p.gl.code + (p.gl.cat ? ' · ' + p.gl.cat.name : ''), [p.row], p.gl); });
    return out.filter(function(k){ return Math.abs(k.v.w) >= 1 || Math.abs(k.v.b) >= 1; })
      .sort(function(a, b){ return Math.abs(b.v.v) - Math.abs(a.v.v); });
  }
  function vendorsIn(E, P, bk){
    var map = {}, order = [];
    VROWS.forEach(function(x){
      if(E && E.type === 'cat' && x.cat !== E) return;
      if(E && E.type === 'gl' && x.gl !== E) return;
      var k = x.vendor.key;
      if(!map[k]){ map[k] = { E:x.vendor, rows:[] }; order.push(k); }
      map[k].rows.push(x.row);
    });
    return order.map(function(k){ var m = map[k]; return { E:m.E, name:m.E.name, nov:m.E.nov, v:withB(sumVals(m.rows, P), bk) }; })
      .filter(function(k){ return Math.abs(k.v.w) >= 1 || Math.abs(k.v.b) >= 1; });
  }
  function kidTable(list, n, P, bk, scale){
    if(!list.length) return '';
    var t = thrOf(P, scale || 0.4);
    var h = '<table class="fb-t"><thead><tr><th></th><th>Working</th><th>' + (bk === 'f' ? 'FCST' : 'Plan') + '</th><th>Var</th></tr></thead><tbody>';
    list.slice(0, n).forEach(function(k){
      h += '<tr><td>' + esc(k.name) + (k.sub ? '<small>' + esc(k.sub) + '</small>' : '') + '</td><td>' + fm(k.v.w) + '</td><td>' + fm(k.v.b) + '</td><td>' + vspan(k.v.v, t) + '</td></tr>';
    });
    if(list.length > n) h += '<tr class="fb-more"><td colspan="4">+ ' + (list.length - n) + ' more</td></tr>';
    return h + '</tbody></table>';
  }
  function driverSentence(list, t){
    var up = list.filter(function(k){ return k.v.v >= t; }).sort(function(a, b){ return b.v.v - a.v.v; }).slice(0, 3);
    var dn = list.filter(function(k){ return k.v.v <= -t; }).sort(function(a, b){ return a.v.v - b.v.v; }).slice(0, 3);
    function nm(k){ return esc(k.name) + ' (' + vspan(k.v.v, t) + ')'; }
    if(!up.length && !dn.length){
      var mu = list.filter(function(k){ return k.v.v >= t * 0.2; }).sort(function(a, b){ return b.v.v - a.v.v; }).slice(0, 2);
      var md = list.filter(function(k){ return k.v.v <= -t * 0.2; }).sort(function(a, b){ return a.v.v - b.v.v; }).slice(0, 2);
      if(!mu.length && !md.length) return 'No line crosses the ±' + fm(t) + ' materiality threshold; the variance is spread across many small items.';
      var parts = [];
      if(mu.length) parts.push(mu.map(nm).join(', '));
      if(md.length) parts.push((mu.length ? 'offset by ' : '') + md.map(nm).join(', '));
      return 'Nothing crosses the ±' + fm(t) + ' materiality threshold. Largest moves: ' + parts.join(' ') + '.';
    }
    var s = [];
    if(up.length) s.push('Overspend in ' + up.map(nm).join(', '));
    if(dn.length) s.push((up.length ? 'offset by savings in ' : 'Savings in ') + dn.map(nm).join(', '));
    return s.join(', ') + '.';
  }
  function chipsFor(list){ return list.filter(Boolean).slice(0, 4); }

  /* ---------------- real vs projected & buffer ----------------
     Working = actuals through the review month + projection after it.
     Buffer  = "No Vendor" lines outside Comp & Benefits, in months after the review month:
               above Plan  -> savings reallocated (parked) into future months;
               below Plan  -> buffer released to fund overspends.                        */
  function isCB(c){ return !!(c && /comp(ensation)?\s*(and|&)\s*ben/i.test(c.name)); }
  function monthsOf(P){
    var a = [], k;
    if(P.t === 'm') return [P.i];
    if(P.t === 'q') return [P.i * 3, P.i * 3 + 1, P.i * 3 + 2];
    if(P.t === 'ytd'){ for(k = 0; k <= RM; k++) a.push(k); return a; }
    for(k = 0; k < 12; k++) a.push(k); return a;
  }
  function hasFuture(P){ return monthsOf(P).some(function(m){ return m > RM; }); }
  function closedLabel(P){
    var c = monthsOf(P).filter(function(m){ return m <= RM; });
    if(!c.length) return null;
    return c.length === 1 ? MONTHS[c[0]] : MONTHS[c[0]] + ' – ' + MONTHS[c[c.length - 1]];
  }
  function projLabel(P){
    var f = monthsOf(P).filter(function(m){ return m > RM; });
    if(!f.length) return null;
    return f.length === 1 ? MONTHS[f[0]] : MONTHS[f[0]] + ' – ' + MONTHS[f[f.length - 1]];
  }
  function decompose(scope, P, bk){
    var o = { closed:0, parked:0, funding:0, cb:0, other:0, total:0, parkedL:[], fundL:[], overL:[], otherL:[] };
    VROWS.forEach(function(x){
      if(scope && scope.type === 'cat' && x.cat !== scope) return;
      if(scope && scope.type === 'gl' && x.gl !== scope) return;
      if(scope && scope.type === 'vendor' && x.vendor !== scope) return;
      monthsOf(P).forEach(function(m){
        var a = x.row.M && x.row.M[m]; if(!a) return;
        var dd = a[0] - (bk === 'f' ? a[2] : a[1]);
        if(Math.abs(dd) < 0.5) return;
        o.total += dd;
        if(m <= RM){ o.closed += dd; return; }
        if(x.vendor.nov && !isCB(x.cat)){
          if(dd > 0){ o.parked += dd; o.parkedL.push({ m:m, gl:x.gl, d:dd }); }
          else { o.funding += dd; o.fundL.push({ m:m, gl:x.gl, d:dd }); }
        } else if(isCB(x.cat)) o.cb += dd;
        else { o.other += dd; o.otherL.push({ m:m, x:x, d:dd }); if(dd > 0) o.overL.push({ m:m, x:x, d:dd }); }
      });
    });
    o.buffer = o.parked + o.funding; o.excl = o.total - o.buffer;
    return o;
  }
  function nspan(v){ return '<span class="fb-neu">' + fv(v) + '</span>'; }
  function monthList(list){
    var seen = {}, ms = [];
    list.forEach(function(e){ if(!seen[e.m]){ seen[e.m] = 1; ms.push(e.m); } });
    return ms.sort(function(a, b){ return a - b; }).map(function(m){ return MONTHS[m]; }).join(', ');
  }
  function topNames(list, n){
    var agg = {}, order = [];
    list.forEach(function(e){ var k = e.x.vendor.nov ? e.x.gl.name : e.x.vendor.name; if(!(k in agg)){ agg[k] = 0; order.push(k); } agg[k] += e.d; });
    return order.map(function(k){ return { k:k, d:agg[k] }; }).filter(function(r){ return Math.abs(r.d) >= 0.5; })
      .sort(function(a, b){ return Math.abs(b.d) - Math.abs(a.d); }).slice(0, n)
      .map(function(r){ return r.k + ' ' + fv(r.d); }).join(', ');
  }
  // Bridge: closed-month actuals -> buffer moves -> projected changes = reported variance.
  function decompTable(o, P, bk, scope){
    var t = thrOf(P, scope && scope.type !== 'total' ? 0.4 : 1), cl = closedLabel(P), pj = projLabel(P), bn = bk === 'f' ? 'Forecast' : 'Plan';
    var rows = [];
    if(cl) rows.push([cl + ' actual ' + (o.closed <= 0 ? 'savings' : 'overspend'), 'Real spend vs ' + bn + ' in closed months', o.closed]);
    if(pj){
      if(o.parked >= 0.5){ var pn = nvNamesList(o.parkedL, bk, 1); rows.push(['Savings moved to buffer (' + monthList(o.parkedL) + ')', 'Re-budgeted into future months · not spend' + (pn.length ? ' · line items: ' + pn.slice(0, 2).join(', ') + (pn.length > 2 ? ', …' : '') : ''), o.parked, 1]); }
      if(o.funding <= -0.5){
        var cov = topNames(o.overL.filter(function(x){ return o.fundL.some(function(f){ return f.m === x.m; }); }), 2);
        rows.push(['Buffer used to cover overspends (' + monthList(o.fundL) + ')', cov ? 'Covers ' + cov + ' (same month)' : 'Released to fund overspends', o.funding, 1]);
      }
      if(Math.abs(o.cb) >= 0.5) rows.push([(o.cb > 0 ? 'Higher' : 'Lower') + ' payroll projected (' + pj + ')', 'Comp & Benefits', o.cb]);
      if(Math.abs(o.other) >= 0.5){ var tn = topNames(o.otherL, 2); rows.push(['Other projected spend changes (' + pj + ')', tn ? 'Mostly ' + tn : '', o.other]); }
    }
    var h = '<table class="fb-t fb-bridge"><thead><tr><th></th><th>vs ' + bn + '</th></tr></thead><tbody>';
    rows.forEach(function(r){ h += '<tr' + (r[3] ? ' class="fb-bufrow"' : '') + '><td>' + esc(r[0]) + (r[1] ? '<small>' + esc(r[1]) + '</small>' : '') + '</td><td>' + (r[3] ? nspan(r[2]) : vspan(r[2], t)) + '</td></tr>'; });
    h += '<tr class="fb-tot"><td>Reported variance vs ' + bn + '</td><td>' + vspan(o.total, t) + '</td></tr>';
    if(pj && Math.abs(o.buffer) >= 0.5) h += '<tr class="fb-hl"><td>Underlying result without buffer moves</td><td>' + vspan(o.excl, t) + '</td></tr>';
    return h + '</tbody></table>';
  }
  // Conclusion first, then how it builds up.
  function bufferSentence(o, P, bk, who){
    if(!hasFuture(P) || Math.abs(o.buffer) < 0.5) return '';
    var t = thrOf(P), bn = esc(bName(bk)), pl = esc(pLabel(P));
    var look = Math.abs(o.total) < t ? 'looks in line with ' + bn + ' (' + vspan(o.total, t) + ')' : 'shows ' + vspan(o.total, t) + ' ' + (o.total > 0 ? 'over' : 'under') + ' ' + bn;
    var link = verdict(o.total, t) !== verdict(o.excl, t) ? ', but without buffer moves ' : '; without buffer moves ';
    var s = '<b>' + pl + ' ' + look + link + esc(who) + ' is ' + vspan(o.excl, t) + ' ' + verdict(o.excl, t) + '.</b> ';
    var parts = [], cl = closedLabel(P), proj = o.cb + o.other;
    if(cl) parts.push(esc(cl) + ' actuals came in ' + vspan(o.closed, t) + ' vs ' + bn);
    if(o.parked >= 0.5) parts.push(fm(o.parked) + ' of savings were moved into the buffer (' + esc(monthList(o.parkedL)) + '), which is not spend');
    if(o.funding <= -0.5) parts.push(fm(-o.funding) + ' of buffer was used to cover overspends');
    if(Math.abs(proj) >= 0.5) parts.push('projected spend changes ' + (proj > 0 ? 'add ' + fm(proj) : 'save ' + fm(-proj)));
    if(parts.length){ var j = parts.join('; '); s += j.charAt(0).toUpperCase() + j.slice(1) + '.'; }
    return s;
  }

  /* ---------------- answers ---------------- */
  function aHelp(){
    return {
      html: '<p>I answer from this board\'s feeds only (no AI model, nothing leaves your browser). I understand English and Spanish; I answer in English.</p>'
        + '<div class="fb-sub">Things you can ask</div><ul class="fb-ul">'
        + '<li><b>Overall:</b> "Executive summary", "How are we doing vs plan this quarter?"</li>'
        + '<li><b>Any category, account or vendor:</b> "Zoom vs plan", "What happened with Consulting in Q2?", "670800 YTD"</li>'
        + '<li><b>Compare:</b> "Zoom vs Microsoft", "Compare CDW and Anthropic full year"</li>'
        + '<li><b>Rankings:</b> "Top 5 unfavorable vendors", "Biggest savings by account full year vs forecast"</li>'
        + '<li><b>Exceptions:</b> "What is not in plan?", "Budget with no spend YTD"</li>'
        + '<li><b>Pacing:</b> "Budget utilization", "Monthly trend for Software"</li>'
        + '<li><b>Buffer:</b> "Real vs projected", "Where are the savings parked?", "Software excluding buffer"</li>'
        + (NV ? '<li><b>No Vendor detail:</b> "No Vendor line items", "No Vendor in Consulting", or the name of a line item</li>' : '')
        + '<li><b>People & T&E:</b> "Headcount", "Open TBH roles", "T&E by employee", "Jorge Herrera"</li>'
        + '<li><b>Board items:</b> "Open actions"</li></ul>'
        + '<p class="fb-dim">Defaults: current quarter (' + esc(QL[RQ]) + ') and ' + esc(PLAN) + ' unless you name a month, quarter, YTD, full year or Forecast.</p>',
      chips: ['Executive summary', 'Top 5 unfavorable vendors this quarter', 'What is not in plan YTD?', 'Budget utilization']
    };
  }

  function aSummary(){
    var Ps = [{ t:'m', i:RM }, { t:'q', i:RQ }, { t:'ytd' }, { t:'fy' }];
    var tbl = '<table class="fb-t"><thead><tr><th></th><th>Working</th><th>vs Plan</th><th>vs FCST</th></tr></thead><tbody>';
    Ps.forEach(function(P){ var x = ev(TOTAL, P, 'p'), t = thrOf(P); tbl += '<tr><td>' + esc(pLabel(P)) + '</td><td>' + fm(x.w) + '</td><td>' + vspan(x.vp, t) + '</td><td>' + vspan(x.vf, t) + '</td></tr>'; });
    tbl += '</tbody></table>';
    var Pq = { t:'q', i:RQ }, Pf = { t:'fy' };
    var q = ev(TOTAL, Pq, 'p'), fy = ev(TOTAL, Pf, 'p'), fyf = ev(TOTAL, Pf, 'f');
    function vs(x, t, lbl){
      if(Math.abs(x.v) < t) return 'in line with ' + esc(lbl) + ' (' + vspan(x.v, t) + ')';
      return vspan(x.v, t) + (pct(x.v, x.b) ? ' (' + pct(x.v, x.b) + ')' : '') + ' ' + (x.v > 0 ? 'over' : 'under') + ' ' + esc(lbl) + ' ' + tag(x.v, t);
    }
    var h = '<p><b>' + esc(BU) + ' · ' + esc(META.month || '') + ' ' + esc(FY) + '.</b> ' + esc(QL[RQ]) + ' Working is <b>' + fm(q.w) + '</b>, ' + vs(q, THR.q, PLAN)
      + '. Full year is <b>' + fm(fy.w) + '</b>, ' + vs(fy, THR.fy, PLAN) + ', and ' + vs(fyf, THR.fy, FCST) + '.</p>';
    var ofy = decompose(null, Pf, 'p');
    if(Math.abs(ofy.buffer) >= 0.5){
      h += '<p>' + bufferSentence(ofy, Pf, 'p', BU) + '</p>';
    }
    h += tbl;
    h += '<div class="fb-sub">What drives ' + esc(QL[RQ]) + ' vs Plan</div><p>' + driverSentence(kids(TOTAL, Pq, 'p'), THR.q) + '</p>';
    h += '<div class="fb-sub">What drives the full year vs Plan</div><p>' + driverSentence(kids(TOTAL, Pf, 'p'), THR.fy) + '</p>';
    var vq = vendorsIn(null, Pq, 'p').filter(function(k){ return !k.nov; }).sort(function(a, b){ return b.v.v - a.v.v; });
    var watch = [];
    if(vq[0] && vq[0].v.v >= THR.q * 0.4) watch.push('Largest vendor overspend this quarter: <b>' + esc(vq[0].name) + '</b> ' + vspan(vq[0].v.v, 1) + '.');
    var unpl = VROWS.map(function(x){ return withB(rowVal(x.row, { t:'ytd' }), 'p'); }).filter(function(x){ return Math.abs(x.p) < 1 && Math.abs(x.w) >= 1; });
    if(unpl.length) watch.push('Unplanned spend YTD: <b>' + fm(unpl.reduce(function(s, x){ return s + x.w; }, 0)) + '</b> across ' + unpl.length + ' vendor/account lines with no Plan.');
    if(HC && HC.salary){ var hv = HC.salary.workTotal - HC.salary.planTotal; watch.push('Salary Accrued (HC feed) full year: ' + fm(HC.salary.workTotal) + ' vs ' + fm(HC.salary.planTotal) + ' Plan → ' + vspan(hv, THR.q) + '.'); }
    if(watch.length) h += '<div class="fb-sub">Watch items</div><ul class="fb-ul">' + watch.map(function(w){ return '<li>' + w + '</li>'; }).join('') + '</ul>';
    var notes = boardNotes().filter(function(n){ return n.label && n.sid !== 'sec-hc'; }).concat(notesFor(TOTAL));
    h += notesHtml(notes);
    return { html:h, chips:['Real vs projected full year', 'Top 5 unfavorable vendors this quarter', 'Full year vs forecast', 'What is not in plan YTD?'] };
  }

  function aBuffer(scope, P, bk){
    bk = bk === 'both' ? 'p' : (bk || 'p');
    var who = scope ? scope.name : BU;
    var h = '<div class="fb-h">Real vs projected · ' + esc(who) + ' <small>' + esc(pLabel(P)) + ' vs ' + esc(bName(bk)) + '</small></div>';
    if(!hasFuture(P)){
      var oc = decompose(scope, P, bk);
      h += '<p>All months in ' + esc(pLabel(P)) + ' are closed, so the whole variance (' + vspan(oc.total, thrOf(P)) + ') is real. Buffer only exists in months after ' + esc(MONTHS[RM]) + '.</p>';
      return { html:h, chips:['Real vs projected full year', 'Real vs projected ' + QL[RQ]] };
    }
    var o = decompose(scope, P, bk);
    var line = bufferSentence(o, P, bk, who);
    h += '<p>' + (line || 'There are no buffer moves in ' + esc(pLabel(P)) + ', so the reported variance (' + vspan(o.total, thrOf(P)) + ') is the underlying result.') + '</p>';
    h += '<div class="fb-sub">How the variance builds up</div>' + decompTable(o, P, bk, scope);
    h += '<p class="fb-dim">How to read it: Working = actuals through ' + esc(MONTHS[RM]) + ' + projection after. Buffer moves only shift savings between months ("No Vendor" lines outside Comp &amp; Benefits); they are not spend.</p>';
    if(!scope){
      var rows = CATS.map(function(c){ var x = decompose(c, P, bk); return { n:c.name, o:x }; }).filter(function(r){ return Math.abs(r.o.total) >= 0.5; })
        .sort(function(a, b){ return Math.abs(b.o.total) - Math.abs(a.o.total); });
      var t = thrOf(P, 0.4);
      h += '<div class="fb-sub">By category</div><table class="fb-t fb-tw"><thead><tr><th></th><th>Actual</th><th>Buffer moves</th><th>Projection</th><th>Without buffer</th></tr></thead><tbody>';
      rows.forEach(function(r){ h += '<tr><td>' + esc(r.n) + '</td><td>' + vspan(r.o.closed, t) + '</td><td>' + nspan(r.o.buffer) + '</td><td>' + vspan(r.o.cb + r.o.other, t) + '</td><td>' + vspan(r.o.excl, t) + '</td></tr>'; });
      h += '</tbody></table>';
    }
    function grp(list){
      var m = {}, order = [];
      list.forEach(function(e){ var k = e.gl.key + '|' + e.m; if(!m[k]){ m[k] = { gl:e.gl, m:e.m, d:0 }; order.push(k); } m[k].d += e.d; });
      return order.map(function(k){ return m[k]; });
    }
    var pk = grp(o.parkedL).sort(function(a, b){ return b.d - a.d; });
    if(pk.length){
      h += '<div class="fb-sub">Savings moved to buffer</div><table class="fb-t"><tbody>';
      pk.slice(0, 6).forEach(function(e){ var nn = nvNames(e.gl.code, e.m, bk, 1); h += '<tr><td>' + esc(e.gl.code + ' ' + e.gl.name) + '<small>' + esc(MONTHS[e.m]) + (nn.length ? ' · ' + esc(nn.slice(0, 2).join(', ')) : '') + '</small></td><td>' + nspan(e.d) + '</td></tr>'; });
      if(pk.length > 6) h += '<tr class="fb-more"><td colspan="2">+ ' + (pk.length - 6) + ' more</td></tr>';
      h += '</tbody></table>';
    }
    var fd = grp(o.fundL).sort(function(a, b){ return a.d - b.d; });
    if(fd.length){
      h += '<div class="fb-sub">Buffer used to cover overspends</div><table class="fb-t"><tbody>';
      fd.slice(0, 5).forEach(function(e){
        var same = o.overL.filter(function(x){ return x.m === e.m; }).sort(function(a, b){ return b.d - a.d; }).slice(0, 2);
        var fn = nvNames(e.gl.code, e.m, bk, -1);
        h += '<tr><td>' + esc(e.gl.code + ' ' + e.gl.name) + '<small>' + esc(MONTHS[e.m]) + (fn.length ? ' · ' + esc(fn.slice(0, 2).join(', ')) : '') + (same.length ? ' · covers ' + same.map(function(x){ return esc(x.x.vendor.name) + ' ' + fv(x.d); }).join(', ') : '') + '</small></td><td>' + nspan(e.d) + '</td></tr>';
      });
      h += '</tbody></table>';
    }
    if(NV) h += '<p class="fb-dim">Line item names come from the No Vendor feed. Ask "No Vendor line items" for the full list.</p>';
    return { html:h, chips:chipsFor([
      P.t === 'fy' ? 'Real vs projected ' + QL[RQ] : 'Real vs projected full year',
      'Real vs projected vs ' + (bk === 'p' ? 'forecast' : 'plan'),
      'Budget utilization',
      scope ? 'Real vs projected full year' : 'Executive summary'
    ]) };
  }

  function aTotal(P, bk){
    var bks = bk === 'both' ? ['p','f'] : [bk || 'p'];
    var bk0 = bks[0], t = thrOf(P);
    var od = hasFuture(P) ? decompose(null, P, bk0) : null, bs = od ? bufferSentence(od, P, bk0, BU) : '';
    // With buffer moves, the conclusion replaces the generic headline for the main benchmark.
    var h = bks.map(function(b, i){ return '<p>' + (i === 0 && bs ? bs : headline(TOTAL, P, b)) + '</p>'; }).join('');
    if(od){
      h += '<div class="fb-sub">How the variance builds up</div>' + decompTable(od, P, bk0, null);
    }
    var cats = kids(TOTAL, P, bk0);
    h += '<div class="fb-sub">By category · ' + esc(pLabel(P)) + ' vs ' + esc(bName(bk0)) + '</div><p>' + driverSentence(cats, t) + '</p>' + kidTable(cats, 8, P, bk0, 1);
    var vs = vendorsIn(null, P, bk0).filter(function(k){ return !k.nov; }).sort(function(a, b){ return Math.abs(b.v.v) - Math.abs(a.v.v); });
    if(vs.length) h += '<div class="fb-sub">Largest vendor variances</div>' + kidTable(vs, 5, P, bk0);
    return { html:h, chips:chipsFor([bk0 === 'p' ? 'Same vs forecast' : 'Same vs plan', P.t !== 'fy' ? 'Full year vs plan' : 'This quarter vs plan', 'Top 5 unfavorable vendors ' + pLabel(P), hasFuture(P) ? 'Where are the savings parked?' : 'Monthly trend']) };
  }

  function aEntity(E, P, bk, why){
    var bks = bk === 'both' ? ['p','f'] : [bk || 'p'], bk0 = bks[0];
    var h = '<div class="fb-h">' + esc(E.name) + ' <small>' + esc(typeLabel(E)) + (E.code ? ' · ' + esc(E.code) : '') + '</small></div>';
    h += bks.map(function(b){ return '<p>' + headline(E, P, b) + '</p>'; }).join('');
    var x = ev(E, P, bk0), t = thrOf(P, scaleOf(E));
    if(E.type === 'vendor'){
      var where = E.pairs.map(function(p){ return esc(p.gl.code + ' ' + p.gl.name) + (p.gl.cat ? ' <span class="fb-dim">(' + esc(p.gl.cat.name) + ')</span>' : ''); });
      h += '<p class="fb-dim">Booked to: ' + where.join('; ') + '</p>';
    } else if(E.type === 'gl' && E.cat){
      h += '<p class="fb-dim">Part of category ' + esc(E.cat.name) + '.</p>';
    } else if(E.type === 'cat'){
      var tv = ev(TOTAL, P, bk0).v;
      if(Math.abs(tv) >= 1) h += '<p class="fb-dim">Net ' + esc(pLabel(P)) + ' variance for ' + esc(BU) + ' is ' + vspan(tv, thrOf(P)) + '; this category contributes ' + vspan(x.v, t) + '.</p>';
    }
    if((E.type === 'cat' || E.type === 'gl' || (E.type === 'vendor' && E.nov)) && hasFuture(P)){
      var oe = decompose(E, P, bk0);
      var be = bufferSentence(oe, P, bk0, E.name);
      if(be) h += '<p>' + be + '</p>';
    }
    h += glanceTable(E, P);
    if(E.type === 'cat'){
      var g = kids(E, P, bk0);
      if(g.length) h += '<div class="fb-sub">' + (why ? 'What explains it · ' : '') + 'By account · ' + esc(pLabel(P)) + '</div>' + kidTable(g, 6, P, bk0);
      var vs = vendorsIn(E, P, bk0).filter(function(k){ return !k.nov; }).sort(function(a, b){ return Math.abs(b.v.v) - Math.abs(a.v.v); });
      if(vs.length) h += '<div class="fb-sub">Top vendors</div>' + kidTable(vs, 5, P, bk0);
      var nv = vendorsIn(E, P, bk0).filter(function(k){ return k.nov; })[0];
      if(nv && Math.abs(nv.v.v) >= 1) h += '<p class="fb-dim">"No Vendor" lines in this category: ' + fm(nv.v.w) + ' Working vs ' + fm(nv.v.b) + ' → ' + vspan(nv.v.v, t) + '.</p>';
    } else if(E.type === 'gl'){
      var gv = kids(E, P, bk0);
      if(gv.length) h += '<div class="fb-sub">By vendor · ' + esc(pLabel(P)) + '</div>' + kidTable(gv, 6, P, bk0);
      var ng = nvG(E.code);
      if(ng) h += nvBlock(ng, P, bk0, false);
    } else if(E.type === 'vendor' && E.pairs.length > 1){
      h += '<div class="fb-sub">By account · ' + esc(pLabel(P)) + '</div>' + kidTable(kids(E, P, bk0), 6, P, bk0);
    }
    var notes = notesFor(E);
    h += notesHtml(notes);
    if(why && !notes.length) h += '<p class="fb-dim">No analyst comment on the board mentions ' + esc(E.name) + ' yet. The numbers above are the quantitative driver; add a comment in the driver block to capture the business reason.</p>';
    var chips = [
      E.name + ' vs ' + (bk0 === 'p' ? 'forecast' : 'plan'),
      P.t === 'fy' ? E.name + ' ' + QL[RQ] : E.name + ' full year',
      'Monthly trend for ' + E.name,
      E.type === 'cat' ? 'Top vendors in ' + E.name : (E.type === 'vendor' && E.pairs[0] && E.pairs[0].gl.cat ? E.pairs[0].gl.cat.name + ' vs plan' : null)
    ];
    return { html:h, chips:chipsFor(chips) };
  }

  function aTop(P, bk, dim, dir, n, scope){
    bk = bk === 'both' ? 'p' : (bk || 'p');
    dim = dim || (scope ? 'vendor' : 'cat');
    var list, dimLabel;
    if(dim === 'vendor'){ list = vendorsIn(scope, P, bk).filter(function(k){ return !k.nov; }); dimLabel = 'vendors'; }
    else if(dim === 'gl'){ list = GLS.filter(function(g){ return !scope || (scope.type === 'cat' && g.cat === scope) || g === scope; }).map(function(g){ return { name:g.name, sub:g.code + (g.cat ? ' · ' + g.cat.name : ''), v:ev(g, P, bk), E:g }; }); dimLabel = 'accounts'; }
    else { list = CATS.map(function(c){ return { name:c.name, v:ev(c, P, bk), E:c }; }); dimLabel = 'categories'; }
    list = list.filter(function(k){ return Math.abs(k.v.w) >= 1 || Math.abs(k.v.b) >= 1; });
    if(!list.length && dim === 'vendor' && scope && scope.type === 'cat') return aTop(P, bk, 'gl', dir, n, scope);
    if(dir === 'unfav') list = list.filter(function(k){ return k.v.v > 0.5; }).sort(function(a, b){ return b.v.v - a.v.v; });
    else if(dir === 'fav') list = list.filter(function(k){ return k.v.v < -0.5; }).sort(function(a, b){ return a.v.v - b.v.v; });
    else list = list.filter(function(k){ return Math.abs(k.v.v) >= 0.5; }).sort(function(a, b){ return Math.abs(b.v.v) - Math.abs(a.v.v); });
    var dirLabel = dir === 'unfav' ? 'unfavorable ' : dir === 'fav' ? 'favorable ' : '';
    var title = 'Top ' + Math.min(n, list.length) + ' ' + dirLabel + dimLabel + (scope ? ' in ' + scope.name : '') + ' · ' + pLabel(P) + ' vs ' + bName(bk);
    if(!list.length) return { html:'<p>No ' + dirLabel + dimLabel + (scope ? ' in ' + esc(scope.name) : '') + ' with a variance in ' + esc(pLabel(P)) + ' vs ' + esc(bName(bk)) + '.</p>', chips:['Executive summary'] };
    var shown = list.slice(0, n), sum = shown.reduce(function(s, k){ return s + k.v.v; }, 0);
    var net = scope ? ev(scope, P, bk).v : ev(TOTAL, P, bk).v;
    var h = '<div class="fb-h">' + esc(title) + '</div>' + kidTable(list, n, P, bk, dim === 'cat' ? 1 : 0.4);
    h += '<p>Together these are ' + vspan(sum, 1) + '; the net ' + esc(pLabel(P)) + ' variance' + (scope ? ' for ' + esc(scope.name) : ' for ' + esc(BU)) + ' is ' + vspan(net, 1) + '.</p>';
    if(dim === 'vendor'){
      var nv = vendorsIn(scope, P, bk).filter(function(k){ return k.nov; })[0];
      if(nv && Math.abs(nv.v.v) >= 1) h += '<p class="fb-dim">Excludes "No Vendor" lines (spend booked without a vendor, e.g. payroll): ' + vspan(nv.v.v, 1) + '.</p>';
    }
    var flip = dir === 'unfav' ? 'favorable' : 'unfavorable';
    return { html:h, chips:chipsFor([
      'Top ' + n + ' ' + flip + ' ' + dimLabel + (scope ? ' in ' + scope.name : '') + ' ' + pLabel(P),
      'Top ' + n + ' ' + dirLabel + dimLabel + (scope ? ' in ' + scope.name : '') + ' ' + (P.t === 'fy' ? QL[RQ] : 'full year'),
      dim !== 'vendor' ? 'Top ' + n + ' ' + dirLabel + 'vendors ' + pLabel(P) : 'Top ' + n + ' ' + dirLabel + 'accounts ' + pLabel(P),
      shown[0] && shown[0].E ? shown[0].E.name + ' ' + pLabel(P) : null
    ]) };
  }

  function aUnplanned(P, bk, scope){
    bk = bk === 'both' ? 'p' : (bk || 'p');
    var rows = VROWS.filter(function(x){ return !scope || (scope.type === 'cat' && x.cat === scope) || (scope.type === 'gl' && x.gl === scope); })
      .map(function(x){ return { name:x.vendor.name, sub:x.gl.code + ' ' + x.gl.name, v:withB(rowVal(x.row, P), bk) }; })
      .filter(function(k){ return Math.abs(k.v.b) < 1 && Math.abs(k.v.w) >= 1; })
      .sort(function(a, b){ return b.v.w - a.v.w; });
    if(!rows.length) return { html:'<p>Every vendor/account line with Working spend in ' + esc(pLabel(P)) + (scope ? ' for ' + esc(scope.name) : '') + ' has ' + esc(bName(bk)) + ' behind it. Nothing is unplanned.</p>', chips:['Budget with no spend YTD', 'Executive summary'] };
    var tot = rows.reduce(function(s, k){ return s + k.v.w; }, 0);
    var h = '<p><b>' + rows.length + '</b> vendor/account lines carry Working spend with <b>no ' + esc(bName(bk)) + '</b> in ' + esc(pLabel(P)) + (scope ? ' for ' + esc(scope.name) : '') + ', totaling <b>' + fm(tot) + '</b>.</p>';
    h += kidTable(rows, 10, P, bk);
    h += '<p class="fb-dim">Each line is one vendor under one GL. Negative amounts are credits or reversals.</p>';
    return { html:h, chips:chipsFor(['Budget with no spend ' + pLabel(P), 'What is not in plan full year', 'Top 5 unfavorable vendors ' + pLabel(P)]) };
  }

  function aUnused(P, bk, scope){
    bk = bk === 'both' ? 'p' : (bk || 'p');
    var rows = VROWS.filter(function(x){ return !scope || (scope.type === 'cat' && x.cat === scope) || (scope.type === 'gl' && x.gl === scope); })
      .map(function(x){ return { name:x.vendor.name, sub:x.gl.code + ' ' + x.gl.name, v:withB(rowVal(x.row, P), bk) }; })
      .filter(function(k){ return Math.abs(k.v.w) < 1 && k.v.b >= 1; })
      .sort(function(a, b){ return b.v.b - a.v.b; });
    if(!rows.length) return { html:'<p>No line has ' + esc(bName(bk)) + ' without Working spend in ' + esc(pLabel(P)) + '.</p>', chips:['What is not in plan YTD?'] };
    var tot = rows.reduce(function(s, k){ return s + k.v.b; }, 0);
    var h = '<p><b>' + rows.length + '</b> lines have <b>' + esc(bName(bk)) + ' but no Working spend</b> in ' + esc(pLabel(P)) + (scope ? ' for ' + esc(scope.name) : '') + ': <b>' + fm(tot) + '</b> of budget not used so far.</p>';
    h += kidTable(rows, 10, P, bk);
    h += '<p class="fb-dim">Check whether these are timing (invoice not yet received / accrual missing) or true savings that can be released.</p>';
    return { html:h, chips:chipsFor(['What is not in plan ' + pLabel(P), 'Budget utilization', 'Executive summary']) };
  }

  function aUtil(E, bk){
    E = E || TOTAL;
    bk = bk === 'both' ? 'p' : (bk || 'p');
    var ytd = ev(E, { t:'ytd' }, bk), fy = ev(E, { t:'fy' }, bk);
    var elapsed = (RM + 1) / 12, used = fy.b ? ytd.w / fy.b : null;
    var left = 11 - RM, remaining = fy.b - ytd.w, runRate = ytd.w / (RM + 1), need = left > 0 ? remaining / left : 0;
    var h = '<div class="fb-h">Budget utilization · ' + esc(E.name) + '</div>';
    if(used === null){
      h += '<p>' + esc(E.name) + ' has no full-year ' + esc(bName(bk)) + '. YTD Working is ' + fm(ytd.w) + '.</p>';
      return { html:h, chips:['Budget utilization'] };
    }
    var pace = used * 100 - elapsed * 100;
    h += '<p>Through <b>' + esc(MONTHS[RM]) + '</b> (' + (RM + 1) + ' of 12 months, ' + (elapsed * 100).toFixed(0) + '% of the year), ' + esc(E.name) + ' has used <b>' + (used * 100).toFixed(1) + '%</b> of its full-year ' + esc(bName(bk)) + ' (' + fm(ytd.w) + ' of ' + fm(fy.b) + '). '
      + (Math.abs(pace) < 3 ? 'That is roughly on a straight-line pace.' : pace > 0 ? 'That is <span class="fb-bad">ahead of a straight-line pace</span> by ' + pace.toFixed(1) + ' pts.' : 'That is <span class="fb-good">behind a straight-line pace</span> by ' + Math.abs(pace).toFixed(1) + ' pts.') + '</p>';
    h += '<p>Against the time-phased ' + esc(bName(bk)) + ', YTD is ' + vspan(ytd.v, THR.ytd * scaleOf(E)) + ' (' + pctPlain(ytd.w, ytd.b) + ' of YTD ' + esc(bName(bk)) + ').</p>';
    if(left > 0) h += '<p>Remaining budget: <b>' + fm(remaining) + '</b> for ' + left + ' months → <b>' + fm(need) + '/month</b> to land on ' + esc(bName(bk)) + ', vs a YTD average run-rate of ' + fm(runRate) + '/month.</p>';
    h += '<p>Full-year Working is ' + fm(fy.w) + ', ' + vspan(fy.v, THR.fy * scaleOf(E)) + ' vs ' + esc(bName(bk)) + ' ' + tag(fy.v, THR.fy * scaleOf(E)) + '.</p>';
    if(E.type === 'total'){
      var rows = CATS.map(function(c){ var a = ev(c, { t:'ytd' }, bk), b = ev(c, { t:'fy' }, bk); return { name:c.name, ytd:a.w, plan:b.b, used:b.b ? a.w / b.b : null }; })
        .filter(function(r){ return Math.abs(r.plan) >= 1 || Math.abs(r.ytd) >= 1; })
        .sort(function(a, b){ return (b.used === null ? 9 : b.used) - (a.used === null ? 9 : a.used); });
      h += '<div class="fb-sub">By category (YTD Working ÷ FY ' + esc(bk === 'f' ? 'FCST' : 'Plan') + ')</div><table class="fb-t fb-tw"><thead><tr><th></th><th>FY budget</th><th>YTD</th><th>Used</th><th>Left</th></tr></thead><tbody>';
      rows.forEach(function(r){
        var u = r.used === null ? 'no plan' : (r.used * 100).toFixed(0) + '%';
        var c = r.used === null ? 'fb-bad' : (r.used - elapsed > 0.05 ? 'fb-bad' : r.used - elapsed < -0.05 ? 'fb-good' : 'fb-neu');
        h += '<tr><td>' + esc(r.name) + '</td><td>' + fm(r.plan) + '</td><td>' + fm(r.ytd) + '</td><td class="' + c + '">' + u + '</td><td>' + fm(r.plan - r.ytd) + '</td></tr>';
      });
      h += '</tbody></table><p class="fb-dim">Red = consuming faster than ' + (elapsed * 100).toFixed(0) + '% of the year elapsed; green = slower.</p>';
    }
    if(DATA.hasOpex) h += '<p class="fb-dim">The OPEX Feed has its own Budget Utilization view in the sidebar.</p>';
    return { html:h, chips:chipsFor(['Monthly trend' + (E.type !== 'total' ? ' for ' + E.name : ''), 'Budget with no spend YTD', E.type === 'total' ? 'Executive summary' : E.name + ' vs plan']) };
  }

  function aTrend(E, bk){
    E = E || TOTAL;
    bk = bk === 'both' ? 'p' : (bk || 'p');
    var rows = [], cum = 0;
    for(var i = 0; i <= RM; i++){ var x = ev(E, { t:'m', i:i }, bk); cum += x.v; rows.push({ m:MONTHS[i], x:x, cum:cum }); }
    var t = THR.m * scaleOf(E);
    var h = '<div class="fb-h">Monthly trend · ' + esc(E.name) + ' vs ' + esc(bName(bk)) + '</div>';
    var avg = rows.reduce(function(s, r){ return s + r.x.w; }, 0) / rows.length;
    var last3 = rows.slice(-3), prev3 = rows.slice(-6, -3);
    var a3 = last3.reduce(function(s, r){ return s + r.x.w; }, 0) / last3.length;
    var peak = rows.slice().sort(function(a, b){ return b.x.w - a.x.w; })[0];
    var s = '<p>Average monthly Working ' + MONTHS[0] + '–' + MONTHS[RM] + ' is <b>' + fm(avg) + '</b>. ';
    if(prev3.length === 3){
      var p3 = prev3.reduce(function(s2, r){ return s2 + r.x.w; }, 0) / 3, ch = p3 ? (a3 - p3) / Math.abs(p3) * 100 : 0;
      s += 'The last 3 months average ' + fm(a3) + ' vs ' + fm(p3) + ' in the 3 months before (' + (ch > 0 ? '+' : '') + ch.toFixed(0) + '%), so spend is ' + (Math.abs(ch) < 5 ? 'flat' : ch > 0 ? 'trending up' : 'trending down') + '. ';
    }
    s += 'Peak month: ' + esc(peak.m) + ' at ' + fm(peak.x.w) + '. Cumulative variance through ' + esc(MONTHS[RM]) + ': ' + vspan(cum, THR.ytd * scaleOf(E)) + '.</p>';
    h += s + '<table class="fb-t fb-tw"><thead><tr><th></th><th>Working</th><th>' + (bk === 'f' ? 'FCST' : 'Plan') + '</th><th>Var</th><th>Cum. var</th></tr></thead><tbody>';
    rows.forEach(function(r){ h += '<tr' + (r.m === MONTHS[RM] ? ' class="fb-hl"' : '') + '><td>' + esc(r.m) + '</td><td>' + fm(r.x.w) + '</td><td>' + fm(r.x.b) + '</td><td>' + vspan(r.x.v, t) + '</td><td>' + vspan(r.cum, t) + '</td></tr>'; });
    h += '</tbody></table>';
    var fy = ev(E, { t:'fy' }, bk);
    if(Math.abs(fy.b) >= 1) h += '<p class="fb-dim">Simple run-rate check: YTD average × 12 = ' + fm(avg * 12) + ' vs full-year ' + esc(bName(bk)) + ' ' + fm(fy.b) + ' (' + vspan(avg * 12 - fy.b, THR.fy * scaleOf(E)) + '). This is a naive extrapolation, not a forecast.</p>';
    return { html:h, chips:chipsFor([E.type === 'total' ? 'Budget utilization' : 'Budget utilization for ' + E.name, E.type === 'total' ? 'Executive summary' : E.name + ' vs plan', 'Monthly trend' + (E.type !== 'total' ? ' for ' + E.name : '') + ' vs ' + (bk === 'p' ? 'forecast' : 'plan')]) };
  }

  function hcStatus(e){
    if(e.tbh){ if(e.p > 0 && Math.abs(e.w) < 1) return 'Open, not started'; if(e.p > 0 && e.w > 0) return 'Filling / started'; if(Math.abs(e.p) < 1 && e.w > 0) return 'Added, not in plan'; return 'No cost'; }
    if(Math.abs(e.p) < 1 && e.w > 0) return 'Not in plan (new hire / transfer)';
    if(e.p > 0 && Math.abs(e.w) < 1) return 'Planned, no cost';
    return e.w - e.p > 0 ? 'Above plan' : e.w - e.p < 0 ? 'Below plan' : 'On plan';
  }
  function aHC(sub){
    if(!HC || !HC.salary) return { html:'<p>This board was built without an HC feed, so headcount detail is not available. The Comp and Benefits category in the BvA feeds still covers payroll cost.</p>', chips:['Comp and Benefits vs plan'] };
    var s = HC.salary, v = s.workTotal - s.planTotal;
    var emps = (HC.employees || []).map(function(e){ return { n:e.n, p:+e.p || 0, w:+e.w || 0, tbh:!!e.tbh }; });
    var h = '';
    if(sub === 'tbh' || sub === 'new'){
      var list = sub === 'tbh' ? emps.filter(function(e){ return e.tbh; }) : emps.filter(function(e){ return !e.tbh && Math.abs(e.p) < 1 && e.w > 0; });
      h += '<div class="fb-h">' + (sub === 'tbh' ? 'TBH roles' : 'Hires not in plan') + ' · ' + list.length + '</div>';
      if(!list.length) return { html:h + '<p>None in the HC feed.</p>', chips:['Headcount'] };
      h += '<table class="fb-t"><thead><tr><th></th><th>Plan</th><th>Working</th><th>Var</th></tr></thead><tbody>';
      list.sort(function(a, b){ return Math.abs(b.w - b.p) - Math.abs(a.w - a.p); }).forEach(function(e){ h += '<tr><td>' + esc(e.n) + '<small>' + esc(hcStatus(e)) + '</small></td><td>' + fm(e.p) + '</td><td>' + fm(e.w) + '</td><td>' + vspan(e.w - e.p, 5000) + '</td></tr>'; });
      h += '</tbody></table>';
      if(sub === 'tbh'){
        var open = list.filter(function(e){ return e.p > 0 && Math.abs(e.w) < 1; });
        if(open.length) h += '<p>' + open.length + ' open req' + (open.length > 1 ? 's' : '') + ' with ' + fm(open.reduce(function(a, e){ return a + e.p; }, 0)) + ' of Plan and no Working cost yet: a source of HC savings if the hire slips.</p>';
      }
      return { html:h, chips:['Headcount', sub === 'tbh' ? 'Hires not in plan' : 'Open TBH roles'] };
    }
    h += '<div class="fb-h">Headcount cost · Salary Accrued</div>';
    h += '<p>Full year Salary Accrued is <b>' + fm(s.workTotal) + '</b> Working vs ' + fm(s.planTotal) + ' Plan → <b>' + vspan(v, THR.q) + '</b>' + (pct(v, s.planTotal) ? ' (' + pct(v, s.planTotal) + ')' : '') + ' ' + tag(v, THR.q) + '.';
    if(s.q && s.qWork && s.q[RQ] != null) h += ' ' + esc(QL[RQ]) + ': ' + fm(s.qWork[RQ]) + ' vs ' + fm(s.q[RQ]) + ' (' + vspan(s.qWork[RQ] - s.q[RQ], THR.m) + ').';
    h += '</p>';
    if(s.q && s.qWork){
      h += '<table class="fb-t"><thead><tr><th></th><th>Plan</th><th>Working</th><th>Var</th></tr></thead><tbody>';
      for(var i = 0; i < 4; i++) h += '<tr' + (i === RQ ? ' class="fb-hl"' : '') + '><td>' + esc(QL[i]) + '</td><td>' + fm(s.q[i]) + '</td><td>' + fm(s.qWork[i]) + '</td><td>' + vspan((s.qWork[i] || 0) - (s.q[i] || 0), THR.m) + '</td></tr>';
      h += '</tbody></table>';
    }
    var m = HC.moves || {};
    var openT = emps.filter(function(e){ return e.tbh && e.p > 0 && Math.abs(e.w) < 1; }).length;
    var newH = emps.filter(function(e){ return !e.tbh && Math.abs(e.p) < 1 && e.w > 0; });
    h += '<p>Ending HC: <b>' + (m.endingWork != null ? m.endingWork : '—') + '</b> Working vs ' + (m.endingPlan != null ? m.endingPlan : '—') + ' Plan. TBH roles in plan: ' + (m.tbhPlan != null ? m.tbhPlan : '—') + ' (' + openT + ' still open). Hires not in plan: ' + newH.length + (newH.length ? ' (' + newH.map(function(e){ return esc(e.n.replace(/\s+-\s+\S+$/, '')); }).join(', ') + ')' : '') + '.</p>';
    var top = emps.slice().sort(function(a, b){ return Math.abs(b.w - b.p) - Math.abs(a.w - a.p); }).slice(0, 5);
    h += '<div class="fb-sub">Largest employee variances</div><table class="fb-t"><thead><tr><th></th><th>Plan</th><th>Working</th><th>Var</th></tr></thead><tbody>';
    top.forEach(function(e){ h += '<tr><td>' + esc(e.n) + '<small>' + esc(hcStatus(e)) + '</small></td><td>' + fm(e.p) + '</td><td>' + fm(e.w) + '</td><td>' + vspan(e.w - e.p, 5000) + '</td></tr>'; });
    h += '</tbody></table>';
    var hn = boardNotes().filter(function(n){ return n.sid === 'sec-hc'; });
    h += notesHtml(hn);
    return { html:h, chips:['Open TBH roles', 'Hires not in plan', 'Comp and Benefits vs plan'] };
  }

  function aEmployee(E){
    var h = '<div class="fb-h">' + esc(E.name) + ' <small>' + (E.tbh ? 'TBH role' : 'Employee') + '</small></div>';
    if(E.hc){
      var e = { n:E.hc.n, p:+E.hc.p || 0, w:+E.hc.w || 0, tbh:!!E.hc.tbh }, v = e.w - e.p;
      h += '<p>Salary Accrued full year: Working <b>' + fm(e.w) + '</b> vs Plan ' + fm(e.p) + ' → <b>' + vspan(v, 5000) + '</b>. Status: ' + esc(hcStatus(e)) + '.</p>';
    } else {
      h += '<p class="fb-dim">Not in the HC feed.</p>';
    }
    if(E.te.length && TE){
      var tot = 0;
      h += '<div class="fb-sub">T&amp;E</div><table class="fb-t fb-tw"><thead><tr><th></th>' + TE.headers.map(function(x){ return '<th>' + esc(x) + '</th>'; }).join('') + '<th>Total</th></tr></thead><tbody>';
      E.te.forEach(function(t){ tot += +t.g || 0; h += '<tr><td>' + esc(String(t.v || '').replace(/\s*-\s*$/, '') || '—') + '</td>' + t.vals.map(function(n){ return '<td>' + (Math.abs(n) >= 0.5 ? fm(n) : '—') + '</td>'; }).join('') + '<td>' + fm(t.g) + '</td></tr>'; });
      h += '</tbody></table><p>Total T&amp;E: <b>' + fm(tot) + '</b>.</p>';
    }
    h += notesHtml(notesFor(E));
    return { html:h, chips:['Headcount', 'T&E by employee'] };
  }

  function aTE(){
    var cat = CATS.filter(function(c){ return /travel/i.test(c.name); })[0];
    var h = '<div class="fb-h">Travel &amp; Expense</div>';
    if(cat){
      var y = ev(cat, { t:'ytd' }, 'p'), f = ev(cat, { t:'fy' }, 'p');
      h += '<p>In the BvA feeds, <b>' + esc(cat.name) + '</b> is ' + fm(y.w) + ' YTD vs ' + fm(y.p) + ' Plan (' + vspan(y.v, 2500) + ') and ' + fm(f.w) + ' full year vs ' + fm(f.p) + ' (' + vspan(f.v, 2500) + ').</p>';
    }
    if(!TE || !TE.rows || !TE.rows.length){ h += '<p class="fb-dim">No T&amp;E feed was loaded for this board.</p>'; return { html:h, chips:cat ? [cat.name + ' vs plan'] : [] }; }
    var people = {};
    TE.rows.forEach(function(t){ var k = String(t.e || '').replace(/\s+-\s+\S+$/, '').trim() || 'Unassigned'; if(/^no employee/i.test(k)) k = 'No employee ID'; people[k] = (people[k] || 0) + (+t.g || 0); });
    var ranked = Object.keys(people).map(function(k){ return { n:k, g:people[k] }; }).filter(function(r){ return Math.abs(r.g) >= 0.5; }).sort(function(a, b){ return b.g - a.g; });
    var totals = TE.total && TE.total.vals ? TE.total.vals : [];
    h += '<div class="fb-sub">T&amp;E feed by month</div><table class="fb-t fb-tw"><thead><tr>' + TE.headers.map(function(x){ return '<th>' + esc(x) + '</th>'; }).join('') + '<th>Total</th></tr></thead><tbody><tr>'
      + totals.map(function(n){ return '<td>' + fm(n) + '</td>'; }).join('') + '<td><b>' + fm(TE.total.g) + '</b></td></tr></tbody></table>';
    h += '<div class="fb-sub">By employee</div><table class="fb-t"><thead><tr><th></th><th>Total</th><th>Share</th></tr></thead><tbody>';
    ranked.slice(0, 8).forEach(function(r){ h += '<tr><td>' + esc(r.n) + '</td><td>' + fm(r.g) + '</td><td>' + pctPlain(r.g, TE.total.g) + '</td></tr>'; });
    h += '</tbody></table>';
    if(ranked[0]) h += '<p>' + esc(ranked[0].n) + ' accounts for ' + pctPlain(ranked[0].g, TE.total.g) + ' of the T&amp;E in the feed.</p>';
    return { html:h, chips:chipsFor([ranked[0] ? ranked[0].n : null, cat ? cat.name + ' vs plan' : null, 'Headcount']) };
  }

  function aActions(){
    var items = [];
    d.querySelectorAll('#actList .act-item').forEach(function(it){
      if(isHidden(it)) return;
      var t = it.querySelector('textarea'), cb = it.querySelector('input[type="checkbox"]');
      var txt = t ? String(t.value || t.textContent || '').trim() : '';
      if(txt) items.push({ t:txt, done:!!(cb && cb.checked) });
    });
    if(!items.length) return { html:'<p>There are no follow-up actions on the board.</p>', chips:['Executive summary'] };
    var open = items.filter(function(i){ return !i.done; });
    var h = '<p><b>' + open.length + '</b> open of ' + items.length + ' follow-up actions.</p><ul class="fb-ul">';
    items.forEach(function(i){ h += '<li' + (i.done ? ' class="fb-done"' : '') + '>' + (i.done ? '✓ ' : '') + esc(i.t) + '</li>'; });
    return { html:h + '</ul>', chips:['Executive summary'] };
  }

  function aFallback(qn){
    var weak = [];
    var qts = queryTokens(qn);
    if(qts.length) ALL.forEach(function(E){ var s = scoreEntity(E, qts, qn); if(s >= 0.6) weak.push({ E:E, s:s }); });
    weak.sort(function(a, b){ return b.s - a.s; });
    var miss = qts.filter(function(t){ return !/^(and|y|what|about)$/.test(t); });
    var h = (miss.length ? '<p>I could not find "<b>' + esc(miss.join(' ')) + '</b>" in this board\'s categories, GL accounts, vendors or people.</p>' : '')
      + '<p>I work with fixed rules over the board data, so try naming a <b>category, account, vendor or person</b>, a <b>period</b> (Aug, Q2, YTD, full year) and <b>Plan or Forecast</b>.</p>';
    var chips = weak.slice(0, 3).map(function(w){ return w.E.name + ' vs plan'; });
    if(chips.length) h += '<p class="fb-dim">Did you mean one of these?</p>';
    return { html:h, chips:chips.concat(['Top 5 unfavorable vendors this quarter', 'Help']).slice(0, 4) };
  }

  /* ---------------- No Vendor line items ---------------- */
  function rangeLbl(ms){ if(!ms.length) return ''; return ms.length === 1 ? MONTHS[ms[0]] : MONTHS[ms[0]] + ' – ' + MONTHS[ms[ms.length - 1]]; }
  function nvBlock(g, P, bk, full){
    if(!g.items.some(function(it){ return it.nv; })) return '';
    var o = nvSplit(g, monthsOf(P), bk), bn = bk === 'f' ? 'FCST' : 'Plan', t = 2500;
    if(!o.has) return '';
    var h = '<div class="fb-sub">' + (full ? esc(g.label) : 'No Vendor line items causing the variance · ' + esc(pLabel(P))) + '</div>';
    if(o.proj.length && (o.rows.length || (g.board && Math.abs(o.unV) >= 1))){
      h += '<table class="fb-t"><thead><tr><th>' + esc(rangeLbl(o.proj)) + '</th><th>Working</th><th>' + bn + '</th><th>Var</th></tr></thead><tbody>';
      o.rows.slice(0, 6).forEach(function(r){
        var tg = Math.abs(r.b) < 0.5 ? 'not in ' + bn : Math.abs(r.w) < 0.5 ? 'no Working' : '';
        h += '<tr><td>' + esc(r.it.d) + (tg ? '<small>' + tg + '</small>' : '') + '</td><td>' + fm(r.w) + '</td><td>' + fm(r.b) + '</td><td>' + vspan(r.v, t) + '</td></tr>';
      });
      if(o.rows.length > 6) h += '<tr class="fb-more"><td colspan="4">+ ' + (o.rows.length - 6) + ' more</td></tr>';
      if(g.board && Math.abs(o.unV) >= 1) h += '<tr class="fb-bufrow"><td>Not explained by line items</td><td>' + fm(o.unW) + '</td><td>' + fm(o.unB) + '</td><td>' + vspan(o.unV, t) + '</td></tr>';
      h += '</tbody></table>';
    }
    if(o.closed.length && g.board && Math.abs(o.cW - o.cB) >= 1){
      h += '<p class="fb-dim">' + esc(rangeLbl(o.closed)) + ' are closed: actual No Vendor ' + fm(o.cW) + ' vs ' + fm(o.cB) + ' ' + bn + ' (' + vspan(o.cW - o.cB, t) + '). Actuals carry no line detail'
        + (o.plan.length ? '; planned lines: ' + o.plan.slice(0, 3).map(function(r){ return esc(r.it.d) + ' ' + fm(r.b); }).join(', ') + (o.plan.length > 3 ? ', …' : '') : '') + '.</p>';
    }
    return h;
  }
  function aNV(scope, P, bk){
    bk = bk === 'both' ? 'p' : (bk || 'p');
    if(!NV) return { html:'<p>This board was built without the No Vendor feed, so I can\'t break down the "No Vendor" amounts. Upload the optional NO VENDOR feed and regenerate the board.</p>', chips:['Real vs projected full year', 'Executive summary'] };
    var gl = NV.gls.filter(function(g){ return !scope || (scope.type === 'gl' && g.code === scope.code) || (scope.type === 'cat' && norm(g.cat) === norm(scope.name)); });
    var withItems = gl.filter(function(g){ return g.items.some(function(it){ return it.nv; }); });
    var h = '<div class="fb-h">No Vendor line items' + (scope ? ' · ' + esc(scope.name) : '') + ' <small>' + esc(pLabel(P)) + ' vs ' + esc(bName(bk)) + '</small></div>';
    var ms = monthsOf(P), proj = ms.filter(function(m){ return m > RM; }), tw = 0, tb = 0, uw = 0;
    withItems.forEach(function(g){ var o = nvSplit(g, ms, bk); tw += o.detW; tb += o.boardW; if(g.board) uw += o.unW; });
    if(proj.length && withItems.length) h += '<p>In the projected months (' + esc(rangeLbl(proj)) + '), line items explain <b>' + fm(tw) + '</b> of ' + fm(tb) + ' No Vendor Working' + (Math.abs(uw) >= 0.5 ? '; <b>' + fm(uw) + '</b> is not explained by line items.' : ', so it fully reconciles.') + '</p>';
    else if(!proj.length) h += '<p>All months in ' + esc(pLabel(P)) + ' are closed. No Vendor actuals carry no line detail, so I show the planned lines next to the actual total.</p>';
    NV.warnings.concat(NV.notes).forEach(function(n){ h += '<p class="fb-dim">' + esc(n) + '</p>'; });
    var shown = 0, quiet = [];
    withItems.forEach(function(g){ var bh = nvBlock(g, P, bk, true); if(bh){ h += bh; shown++; } else quiet.push(g.label); });
    if(!withItems.length) h += '<p>There are no No Vendor line items' + (scope ? ' for ' + esc(scope.name) : '') + ' in the feed.</p>';
    else if(!shown) h += '<p>No line item causes a variance vs ' + esc(bName(bk)) + ' in ' + esc(pLabel(P)) + '.</p>';
    if(shown && quiet.length) h += '<p class="fb-dim">No variance in: ' + quiet.map(esc).join('; ') + '.</p>';
    var nod = gl.filter(function(g){ return !g.items.length && g.board; });
    if(nod.length) h += '<p class="fb-dim">No Vendor amounts without line detail: ' + nod.slice(0, 4).map(function(g){ return esc(g.label); }).join('; ') + (nod.length > 4 ? '; +' + (nod.length - 4) + ' more' : '') + '.</p>';
    var first = withItems[0] && withItems[0].items.filter(function(it){ return it.nv; })[0];
    return { html:h, chips:chipsFor([
      'No Vendor line items' + (scope ? ' ' + scope.name : '') + ' vs ' + (bk === 'p' ? 'forecast' : 'plan'),
      'No Vendor line items' + (scope ? ' ' + scope.name : '') + ' ' + (P.t === 'fy' ? QL[RQ] : 'full year'),
      'Where are the savings parked?',
      first ? first.d : null
    ]) };
  }
  function aNVLine(E){
    var it = E.it, g = E.g, all = [0,1,2,3,4,5,6,7,8,9,10,11];
    var proj = all.filter(function(m){ return m > RM; }), closed = all.filter(function(m){ return m <= RM; });
    var fproj = proj.filter(function(m){ return m >= NVFS; });
    var h = '<div class="fb-h">' + esc(it.d) + ' <small>' + esc(typeLabel(E)) + '</small></div>';
    h += '<p class="fb-dim">' + esc(g.label) + (g.cat ? ' · ' + esc(g.cat) : '') + (it.dt ? ' · contract ' + esc(it.dt) : '') + '</p>';
    var w = nvSum(it.w, proj), p = nvSum(it.p, proj), f = nvSum(it.f, fproj), pFY = nvSum(it.p, all), pc = nvSum(it.p, closed);
    if(proj.length){
      h += '<p>Projected ' + esc(rangeLbl(proj)) + ': Working <b>' + fm(w) + '</b> vs ' + fm(p) + ' ' + esc(PLAN) + ' (' + vspan(w - p, 2500) + ') and ' + fm(f) + ' ' + esc(FCST) + ' (' + vspan(w - f, 2500) + ').'
        + (Math.abs(p) < 0.5 && Math.abs(w) >= 0.5 ? ' This line is <b>not in the Plan</b>.' : Math.abs(w) < 0.5 && Math.abs(p) >= 0.5 ? ' Working no longer carries this line.' : '') + '</p>';
    }
    h += '<p>Full-year ' + esc(PLAN) + ' for this line: ' + fm(pFY) + (closed.length && Math.abs(pc) >= 0.5 ? ', of which ' + fm(pc) + ' falls in the closed months (' + esc(rangeLbl(closed)) + '), where actuals have no line detail.' : '.') + '</p>';
    if(proj.length){
      h += '<table class="fb-t fb-tw"><thead><tr><th></th><th>Working</th><th>Plan</th><th>Forecast</th></tr></thead><tbody>';
      proj.forEach(function(m){ h += '<tr><td>' + esc(MONTHS[m]) + '</td><td>' + fm(it.w[m] || 0) + '</td><td>' + fm(it.p[m] || 0) + '</td><td>' + (m >= NVFS ? fm(it.f[m] || 0) : '—') + '</td></tr>'; });
      h += '</tbody></table>';
    }
    if(!it.nv) h += '<p class="fb-dim">This line is booked with vendor ' + esc(it.v) + ', so it shows in the vendor rows, not in No Vendor.</p>';
    return { html:h, chips:chipsFor(['No Vendor line items' + (g.cat ? ' ' + g.cat : ''), 'Where are the savings parked?', 'Real vs projected full year']) };
  }

  /* ---------------- router ---------------- */
  var CTX = null;
  var FOLLOWABLE = { entity:1, top:1, total:1, unplanned:1, unused:1, util:1, trend:1, buffer:1, compare:1, nv:1 };
  /* ---------------- comparisons ---------------- */
  function multiEntities(qn){
    var chunks = qn.split(/\b(?:vs|versus|against|contra|compared to|compare|comparar|compara|comparado con|frente a|and|y|with|con)\b|,/);
    var out = [], seen = {};
    chunks.forEach(function(c){
      c = String(c || '').trim(); if(!c) return;
      var f = findEntities(c, typeHint(c), ['cat','gl','vendor']);
      if(f[0] && !seen[f[0].E.key]){ seen[f[0].E.key] = 1; out.push(f[0].E); }
    });
    return out.slice(0, 4);
  }
  function short(n){ n = String(n); return n.length > 18 ? n.slice(0, 17) + '…' : n; }
  function aCompare(list, P, bk){
    bk = bk === 'both' ? 'p' : (bk || 'p');
    var sc = list.every(function(E){ return E.type === 'cat'; }) ? 1 : 0.4, t = thrOf(P, sc);
    var h = '<div class="fb-h">' + list.map(function(E){ return esc(E.name); }).join(' vs ') + ' <small>' + esc(pLabel(P)) + ' vs ' + esc(bName(bk)) + '</small></div>';
    h += '<table class="fb-t fb-tw"><thead><tr><th></th><th>Working</th><th>' + (bk === 'f' ? 'FCST' : 'Plan') + '</th><th>Var</th><th>%</th></tr></thead><tbody>';
    var vals = list.map(function(E){ return { E:E, x:ev(E, P, bk) }; });
    vals.forEach(function(r){ h += '<tr><td>' + esc(r.E.name) + '<small>' + esc(typeLabel(r.E)) + '</small></td><td>' + fm(r.x.w) + '</td><td>' + fm(r.x.b) + '</td><td>' + vspan(r.x.v, t) + '</td><td>' + (pct(r.x.v, r.x.b) || '—') + '</td></tr>'; });
    h += '</tbody></table>';
    var parts = vals.map(function(r){
      if(Math.abs(r.x.b) < 1 && Math.abs(r.x.w) >= 1) return esc(r.E.name) + ' has ' + fm(r.x.w) + ' with no ' + esc(bName(bk));
      if(Math.abs(r.x.v) < t) return esc(r.E.name) + ' is roughly on ' + esc(bName(bk)) + ' (' + vspan(r.x.v, t) + ')';
      return esc(r.E.name) + ' is ' + (r.x.v > 0 ? 'over' : 'under') + ' by ' + vspan(r.x.v, t);
    });
    h += '<p>In ' + esc(pLabel(P)) + ', ' + parts.join('; ') + '.';
    var big = vals.slice().sort(function(a, b){ return b.x.w - a.x.w; });
    if(big.length > 1 && big[1].x.w > 0) h += ' ' + esc(big[0].E.name) + ' spends ' + (big[0].x.w / big[1].x.w).toFixed(1) + '× ' + esc(big[1].E.name) + '.';
    h += '</p>';
    var Ps = [{ t:'m', i:RM }, { t:'q', i:RQ }, { t:'ytd' }, { t:'fy' }];
    h += '<div class="fb-sub">Variance vs ' + esc(bName(bk)) + ' by period</div><table class="fb-t fb-tw"><thead><tr><th></th>' + list.map(function(E){ return '<th>' + esc(short(E.name)) + '</th>'; }).join('') + '</tr></thead><tbody>';
    Ps.forEach(function(Q){ h += '<tr' + (samePeriod(Q, P) ? ' class="fb-hl"' : '') + '><td>' + esc(pLabel(Q)) + '</td>' + list.map(function(E){ return '<td>' + vspan(ev(E, Q, bk).v, thrOf(Q, sc)) + '</td>'; }).join('') + '</tr>'; });
    h += '</tbody></table>';
    var names = list.map(function(E){ return E.name; }).join(' vs ');
    return { html:h, chips:chipsFor([
      P.t === 'fy' ? names + ' ' + QL[RQ] : names + ' full year',
      names + ' vs ' + (bk === 'p' ? 'forecast' : 'plan'),
      list[0].name + ' ' + pLabel(P),
      list[1].name + ' ' + pLabel(P)
    ]) };
  }

  /* ---------------- premise check ("why is X over plan?") ---------------- */
  function assertDir(qn){
    if(/\b(over|above|overspen\w*|exceed\w*|sobre|encima|excedid\w*|sobregir\w*|higher than|mas alto)\b/.test(qn) && !/\bover time\b/.test(qn)) return 'unfav';
    if(/\b(under|below|underspen\w*|debajo|ahorr\w*|savings?|lower than|mas bajo)\b/.test(qn)) return 'fav';
    return null;
  }
  function premise(E, P, bk, dir){
    if(!dir) return null;
    var t = thrOf(P, scaleOf(E));
    function ok(x){ return dir === 'unfav' ? x.v >= 0.5 : x.v <= -0.5; }
    var cur = ev(E, P, bk);
    if(ok(cur)) return null;
    var cands = [{ t:'fy' }, { t:'ytd' }, { t:'q', i:RQ }, { t:'m', i:RM }, { t:'q', i:0 }, { t:'q', i:1 }, { t:'q', i:2 }, { t:'q', i:3 }];
    for(var m = 0; m <= RM; m++) cands.push({ t:'m', i:m });
    var best = null;
    cands.forEach(function(Q){ var x = ev(E, Q, bk); if(ok(x) && (!best || Math.abs(x.v) > Math.abs(best.x.v))) best = { P:Q, x:x }; });
    var word = dir === 'unfav' ? 'over' : 'under', nm = esc(E.name), bn = esc(bName(bk));
    var now = 'In ' + esc(pLabel(P)) + ' ' + nm + ' is ' + (Math.abs(cur.v) < 0.5 ? 'exactly on ' + bn : (cur.v > 0 ? 'over' : 'under') + ' ' + bn + ' (' + vspan(cur.v, t) + ')');
    if(best) return { P:best.P, note:'<p class="fb-note-q"><b>Quick check:</b> ' + now + '. It is ' + word + ' ' + bn + ' in <b>' + esc(pLabel(best.P)) + '</b> (' + vspan(best.x.v, thrOf(best.P, scaleOf(E))) + '), so here is that view.</p>' };
    return { P:P, note:'<p class="fb-note-q"><b>Quick check:</b> ' + now + ', and it is not ' + word + ' ' + bn + ' in any period on this board.</p>' };
  }

  /* ---------------- out of scope: recommendations, scenarios, opinions ---------------- */
  function aOutScope(kind, E, P, bk, qn){
    bk = bk === 'both' ? 'p' : (bk || 'p');
    var PP = P || { t:'fy' }, h = '', R, chips;
    if(kind === 'rec'){
      h = '<p><b>I don\'t make recommendations.</b> What to cut or keep is a business call. What I can do is show where the money is going, so that call is easier:</p>';
      if(E){ R = aEntity(E, PP, bk, true); }
      else {
        R = aTop(PP, bk, 'vendor', 'unfav', 5, null);
        var unus = VROWS.map(function(x){ return withB(rowVal(x.row, { t:'ytd' }), 'p'); }).filter(function(x){ return Math.abs(x.w) < 1 && x.p >= 1; });
        if(unus.length) R.html += '<p class="fb-dim">Also: ' + unus.length + ' lines have Plan but no spend YTD (' + fm(unus.reduce(function(a, x){ return a + x.p; }, 0)) + ').</p>';
      }
      chips = ['Budget with no spend YTD', 'Top 5 unfavorable accounts full year', 'Real vs projected full year'];
    } else if(kind === 'scenario'){
      h = '<p><b>I can\'t run what-if scenarios.</b> I only report what is in this board\'s feeds (Working, Plan and Forecast).</p>';
      var pm = qn.match(/(\d+(?:\.\d+)?)\s?(?:%|percent|por ?ciento)/);
      if(pm && E){
        var k = +pm[1] / 100, fy = ev(E, { t:'fy' }, bk), rem = 0;
        for(var m = RM + 1; m < 12; m++){ var a = sumVals(E.rows, { t:'m', i:m }); rem += bk === 'f' ? a.f : a.p; }
        h += '<p>For reference, plain arithmetic, not a forecast: ' + pm[1] + '% of ' + esc(E.name) + '\'s full-year Working (' + fm(fy.w) + ') is <b>' + fm(fy.w * k) + '</b>'
          + (rem > 0 ? ', and ' + pm[1] + '% of its remaining ' + esc(bName(bk)) + ' after ' + esc(MONTHS[RM]) + ' (' + fm(rem) + ') is <b>' + fm(rem * k) + '</b>' : '') + '.</p>';
      }
      h += '<p class="fb-dim">Here is the baseline you would start from:</p>';
      R = E ? aEntity(E, PP, bk, false) : aTotal(PP, bk);
      chips = ['Budget utilization', 'Real vs projected full year', 'Executive summary'];
    } else {
      h = '<p><b>I can\'t judge whether something is "normal".</b> Here is the context an analyst would use to decide: the monthly pattern, the run-rate and the variance vs Plan.</p>';
      R = aTrend(E || null, bk);
      R.html += '<p class="fb-dim">For reference, I flag variances as material above ±$10K for a month, ±$25K for a quarter and ±$75K for the full year (lower for single vendors or accounts).</p>';
      chips = [E ? E.name + ' vs plan' : 'Executive summary', 'Budget utilization', 'Top 5 unfavorable vendors this quarter'];
    }
    return { html:h + R.html, chips:chipsFor(chips) };
  }

  /* ---------------- random questions: a light joke, then back to the board ---------------- */
  function hashStr(q){ var n = 0; for(var i = 0; i < q.length; i++) n = (n * 31 + q.charCodeAt(i)) >>> 0; return n; }
  var RANDOM_TOPIC = /\b(weather|clima|tiempo hace|rain|lluvia|sunny|joke|chiste|funny|lunch|almuerzo|dinner|cena|coffee|cafe|hungry|hambre|pizza|food|comida|love|amor|date|novia|novio|football|futbol|soccer|sports|deporte|game|partido|how are you|como estas|que tal|who are you|quien eres|your name|tu nombre|meaning of life|sentido de la vida|bitcoin|crypto|stock market|bolsa|movie|pelicula|music|musica|song|cancion|vacation|vacaciones|weekend|fin de semana)\b/;
  function teaser(n){
    var Pq = { t:'q', i:RQ }, Pf = { t:'fy' }, opts = [];
    var vq = vendorsIn(null, Pq, 'p').filter(function(k){ return !k.nov; }).sort(function(a, b){ return b.v.v - a.v.v; });
    if(vq[0] && vq[0].v.v > 0) opts.push(esc(vq[0].name) + ' is ' + fv(vq[0].v.v) + ' over Plan in ' + esc(QL[RQ]) + '.');
    var o = decompose(null, Pf, 'p');
    if(Math.abs(o.buffer) >= 0.5) opts.push('without buffer moves, ' + esc(BU) + ' is ' + fv(o.excl) + ' vs Plan for the full year.');
    var ytd = ev(TOTAL, { t:'ytd' }, 'p'); opts.push(esc(BU) + ' is ' + fv(ytd.v) + ' vs Plan year to date.');
    var vs = vendorsIn(null, { t:'ytd' }, 'p').filter(function(k){ return !k.nov; }).sort(function(a, b){ return a.v.v - b.v.v; });
    if(vs[0] && vs[0].v.v < 0) opts.push('the biggest vendor saving YTD is ' + esc(vs[0].name) + ' at ' + fv(vs[0].v.v) + '.');
    return opts[n % opts.length];
  }
  function aRandom(qn){
    var n = hashStr(qn), Pf = { t:'fy' }, line;
    var fyf = ev(TOTAL, Pf, 'f'), fy = ev(TOTAL, Pf, 'p'), ytd = ev(TOTAL, { t:'ytd' }, 'p');
    if(/weather|clima|tiempo hace|rain|lluvia|sunny/.test(qn)) line = 'The only forecast I follow is the ' + esc(FCST) + ', and today it says ' + esc(BU) + ' is ' + fv(fyf.v) + ' vs it for the full year. Bring an umbrella.';
    else if(/joke|chiste|funny/.test(qn)) line = 'Why did the budget break up with the forecast? Too many unexplained variances.';
    else if(/lunch|almuerzo|dinner|cena|coffee|cafe|hungry|hambre|pizza|food|comida/.test(qn)) line = 'I don\'t eat, I only expense. ' + (TE && TE.total ? 'T&amp;E in this board\'s feed is ' + fm(TE.total.g) + ' so far.' : 'And this board has no T&amp;E feed, so I can\'t even do that.');
    else if(/love|amor|date|novia|novio/.test(qn)) line = 'My only relationship is with the Plan, and it\'s complicated: ' + fv(fy.v) + ' for the full year.';
    else if(/football|futbol|soccer|sports|deporte|game|partido/.test(qn)) line = 'The only score I keep is the variance: ' + esc(BU) + ' is ' + fv(ytd.v) + ' vs Plan year to date. Still in the game.';
    else if(/how are you|como estas|que tal/.test(qn)) line = 'Running favorable, thanks for asking: ' + esc(BU) + ' is ' + fv(ytd.v) + ' vs Plan year to date.';
    else if(/who are you|quien eres|your name|tu nombre/.test(qn)) line = 'I\'m Felipe, a rule-based FP&amp;A analyst. I live inside this board and I only speak Working, Plan and Forecast.';
    else line = [
      'That\'s outside my cost center. I\'m budgeted only for BvA questions (and I\'m tracking favorable).',
      'I searched every GL account for that. Closest match: nothing. Want something I can actually reconcile?',
      'Interesting question, but if it doesn\'t have a Working and a Plan column, I\'m lost.',
      'I\'d need a bigger budget to answer that. Meanwhile, here\'s something I do know.'
    ][n % 4];
    var h = '<p>' + line + '</p><p class="fb-dim">Fun fact from this board: ' + teaser(n) + '</p><p>Ask me about a category, account, vendor or period, or pick one of these:</p>';
    return { html:h, chips:['Executive summary', 'Top 5 unfavorable vendors this quarter', 'Real vs projected full year', 'Help'] };
  }

  var RX = {
    menu: /^(menu|main menu|menu principal|sections?|secciones|seccion|home|back|volver|regresar|inicio)\b/,
    help: /^(help|ayuda|hi|hello|hola|hey|start)\b|what can (you|i) (do|ask)|que (puedo|puedes)|como funciona|how does this work|what do you know/,
    actions: /\b(actions?|follow ?ups?|to ?dos?|pendientes?|acciones|tareas|action items?)\b/,
    hcStrong: /\b(headcount|hc|tbh|tbhs|open reqs?|open roles?|reqs?|vacantes?|plantilla|hires?|hiring|new hires?|contrataciones?|employees|empleados|salary accrued|salario devengado)\b/,
    hcWeak: /\b(salary|salaries|salario|salarios|payroll|nomina|people cost)\b/,
    tbh: /\b(tbh|tbhs|open reqs?|open roles?|vacantes?|reqs?)\b/,
    newh: /\b(new hires?|hires? not in plan|not in plan hires?|contrataciones? no planead|nuevos? ingresos?)\b/,
    te: /\b(t and e|t e|tne|travel|viajes?|viaticos?|expense reports?)\b/,
    unplanned: /not in (the )?(plan|budget|forecast)|unplanned|unbudgeted|no presupuestad|fuera del (plan|presupuesto)|sin (plan|presupuesto)|no planead|without (a )?(plan|budget)|no (plan|budget) behind|new vendors?|nuevos proveedores/,
    unused: /unused|no spend|not spent|sin gasto|sin usar|no usad|untouched|zero spend|no activity|sin actividad|budget (with|but) no|plan (with|but) no (spend|working)|not used/,
    buffer: /\b(buffers?|parked|park|reallocat\w*|re allocat\w*|reasign\w*|realocad\w*|excluding buffer|ex buffer|sin (el )?buffer|without (the )?buffer|underlying|real vs (projected|proyectado|projection)|closed vs projected|actuals? vs projected|actual vs projection|savings (parked|allocated|moved|reallocated)|where are the savings|donde estan los ahorros|ahorros? (reasignados?|guardados?|movidos?))\b/,
    util: /utiliz|consum|burn|remaining|restante|disponible|queda|quedan|headroom|how much (budget )?(is )?left|percent of (plan|budget)|% of (plan|budget)|pacing|pace/,
    trend: /\b(trend|trends|trending|tendencia|month over month|mom|monthly|por mes|mensual|run ?rate|evolution|evolucion|over time|by month|cada mes|month by month)\b/,
    summary: /\b(summary|summarize|summarise|resumen|resume|overview|highlights?|executive|exec|big picture|panorama|key takeaways?|recap|status|estado general)\b|how are we doing|como vamos|como estamos|how is it going|how are we tracking/,
    top: /\b(top|biggest|largest|main|major|ranking|rank|worst|best|mayor(es)?|principal(es)?|peores|mejores|most|highest|lowest|key drivers|drivers|mas)\b|which (vendors|accounts|categories)|cuales (son )?(los|las) (proveedores|cuentas|categorias)/,
    why: /\bwhy\b|por que|porque|explain|explica|reason|razon|what happened|que paso|que pasa|driving|drove|driver|cause/,
    money: /\b(total|overall|how much|cuanto|spend|spent|spending|gasto|gastos|gastamos|variance|varianza|vs|versus|against|contra|compared|working|actuals?|expense|expenses|cost|costs|budget|plan|forecast|fcst)\b/,
    opinion: /\b(is (this|that|it|[a-z0-9 ]{1,40}) (normal|ok|okay|good|bad|healthy|worrying|concerning|a problem|reasonable|expected)|es normal|esta bien|es (bueno|malo|preocupante|razonable)|should i (worry|be worried)|debo preocupar\w*|me debo preocupar)\b/,
    scenario: /\bwhat if\b|\bque pasa\w* si\b|\bsi (recort|cort|reduc|aument|baj|sub)\w*|\bscenarios?\b|\bescenarios?\b|\bsimula\w*|\bimpact of (cutting|reducing|adding|removing)\b|\d+(\.\d+)?\s?%|\b\d+ (percent|por ?ciento)\b/,
    rec: /\b(should (we|i)|recommend\w*|recomiend\w*|recomendac\w*|deberia\w*|debemos|what (can|could) we cut|que (podemos|deberiamos) (recortar|cortar)|advice|aconsej\w*|consejo\w*|suggest\w*|sugier\w*|sugerenc\w*)\b/,
    compare: /\b(compare|comparar|compara|comparison|comparacion|versus|vs|against|contra|side by side|frente a)\b/,
    nv: /\bno vendor\b|\bsin proveedor\b|\bline items?\b|\bpartidas?\b|\bdetalle de no vendor\b/,
    follow: /^(and|y|what about|how about|same|lo mismo|and in|and for|y en|y para|y vs|and vs|and the|y el|y la|now|ahora)\b/
  };
  function exec(S){
    switch(S.intent){
      case 'help': return aHelp();
      case 'actions': return aActions();
      case 'summary': return aSummary();
      case 'hc': return aHC(S.sub);
      case 'te': return aTE();
      case 'employee': return aEmployee(S.E);
      case 'unplanned': return aUnplanned(S.P || { t:'ytd' }, S.bk, S.scope);
      case 'unused': return aUnused(S.P || { t:'ytd' }, S.bk, S.scope);
      case 'util': return aUtil(S.E, S.bk);
      case 'trend': return aTrend(S.E, S.bk);
      case 'buffer': return aBuffer(S.scope, S.P || { t:'fy' }, S.bk);
      case 'top': return aTop(S.P || { t:'q', i:RQ }, S.bk, S.dim, S.dir, S.n || 5, S.scope);
      case 'entity':
        var pe = premise(S.E, S.P || { t:'q', i:RQ }, S.bk || 'p', S.assert);
        var re = aEntity(S.E, pe ? pe.P : (S.P || { t:'q', i:RQ }), S.bk || 'p', S.why);
        if(pe) re.html = pe.note + re.html;
        return re;
      case 'total':
        var pt = premise(TOTAL, S.P || { t:'q', i:RQ }, S.bk || 'p', S.assert);
        var rt = aTotal(pt ? pt.P : (S.P || { t:'q', i:RQ }), S.bk || 'p');
        if(pt) rt.html = pt.note + rt.html;
        return rt;
      case 'compare': return aCompare(S.list, S.P || { t:'q', i:RQ }, S.bk || 'p');
      case 'outscope': return aOutScope(S.kind, S.E, S.P, S.bk, S.raw || S.qn);
      case 'random': return aRandom(S.qn);
      case 'nv': return aNV(S.scope, S.P || { t:'fy' }, S.bk);
      case 'nvline': return aNVLine(S.E);
    }
    return aFallback(S.qn || '');
  }
  function route(raw){
    var qn = norm(raw);
    if(!qn) return { intent:'help' };
    var P = parsePeriod(qn), bk = parseBk(qn), P0 = P, bk0 = bk;
    // Inside a section, unspecified period / benchmark default to that section's view.
    if(SECTION){ if(!P && SECTION.P) P = SECTION.P; if(!bk && SECTION.bk) bk = SECTION.bk; }
    var I = {}; Object.keys(RX).forEach(function(k){ I[k] = RX[k].test(qn); });
    var isTop = I.top && !I.summary;
    var found = findEntities(qn, isTop ? null : typeHint(qn));
    var best = found[0] ? found[0].E : null;
    var S = { qn:qn, P:P, bk:bk, raw:String(raw || '').toLowerCase() };
    if(/\d+(\.\d+)?\s?%/.test(S.raw)) I.scenario = true;
    if(I.menu) return (S.intent = 'menu', S);
    if(I.help) return (S.intent = 'help', S);
    if(I.actions && !best) return (S.intent = 'actions', S);
    var nvScope = found.filter(function(x){ return x.E.type === 'cat' || x.E.type === 'gl'; })[0];
    if(NV && I.nv && !I.buffer) return (S.intent = 'nv', S.scope = nvScope ? nvScope.E : null, S);
    if(NV && SECTION && SECTION.id === 'nv' && best && (best.type === 'cat' || best.type === 'gl') && !isTop && !I.compare) return (S.intent = 'nv', S.scope = best, S);
    var bestNE = best && best.type !== 'employee' ? best : null;
    if(I.opinion) return (S.intent = 'outscope', S.kind = 'opinion', S.E = bestNE || (CTX && CTX.E) || null, S);
    if(I.scenario) return (S.intent = 'outscope', S.kind = 'scenario', S.E = bestNE, S);
    if(I.rec) return (S.intent = 'outscope', S.kind = 'rec', S.E = bestNE, S);
    if(I.compare || /\b(and|y)\b/.test(qn)){
      var multi = multiEntities(qn);
      if(multi.length >= 2) return (S.intent = 'compare', S.list = multi, S);
    }
    if(best && best.type === 'nvline' && !isTop) return (S.intent = 'nvline', S.E = best, S);
    if(best && best.type === 'employee' && !isTop) return (S.intent = 'employee', S.E = best, S);
    if(I.hcStrong || (I.hcWeak && !best)){
      if(!(best && (best.type === 'gl' || best.type === 'cat') && !I.hcStrong)){
        S.intent = 'hc'; S.sub = I.newh ? 'new' : I.tbh ? 'tbh' : null; return S;
      }
    }
    if(I.te && (!best || (best.type === 'cat' && /travel/i.test(best.name)))) return (S.intent = 'te', S);
    var scope = best && (best.type === 'cat' || best.type === 'gl') ? best : null;
    if(I.unplanned) return (S.intent = 'unplanned', S.scope = scope, S);
    if(I.unused) return (S.intent = 'unused', S.scope = scope, S);
    if(I.buffer) return (S.intent = 'buffer', S.scope = scope, S);
    if(I.util) return (S.intent = 'util', S.E = best && best.type !== 'employee' ? best : null, S);
    if(I.trend) return (S.intent = 'trend', S.E = best && best.type !== 'employee' ? best : null, S);
    if(I.summary && !best) return P ? (S.intent = 'total', S) : (S.intent = 'summary', S);
    if(isTop && scope && /\bdrivers?\b/.test(qn) && !parseDim(qn)) return (S.intent = 'entity', S.E = scope, S.why = true, S);
    if(isTop && (!best || scope || parseDim(qn))){
      S.intent = 'top'; S.dim = parseDim(qn); S.dir = parseDir(qn); S.n = parseN(qn); S.scope = scope;
      if(!S.dim && !scope && !/\bdrivers?\b/.test(qn) && /\b(vendor|proveedor)/.test(qn)) S.dim = 'vendor';
      return S;
    }
    if(best){ S.intent = 'entity'; S.E = best; S.why = I.why; S.assert = P0 ? null : assertDir(qn); S.alts = found.slice(1).filter(function(x){ return x.s >= found[0].s - 0.3 && x.E.name !== best.name; }).slice(0, 2).map(function(x){ return x.E; }); return S; }
    // follow-ups: "and Q2?", "y vs forecast?", "same for YTD"
    var unk = queryTokens(qn);
    var shortQ = qn.split(' ').length <= 4;
    if(CTX && FOLLOWABLE[CTX.intent] && (I.follow || ((P0 || bk0) && shortQ)) && !I.summary && !unk.length){
      var F = {}; Object.keys(CTX).forEach(function(k){ F[k] = CTX[k]; });
      F.qn = qn; if(P0) F.P = P0; if(bk0) F.bk = bk0;
      var dir = parseDir(qn); if(dir && F.intent === 'top') F.dir = dir;
      return F;
    }
    if(I.summary) return (S.intent = 'summary', S);
    if(isTop){ S.intent = 'top'; S.dim = parseDim(qn); S.dir = parseDir(qn); S.n = parseN(qn); return S; }
    var nTok = qn.split(' ').length;
    if(RANDOM_TOPIC.test(qn) || (!I.money && !P0 && !bk0 && !I.why && !I.follow && nTok >= 3)){ S.intent = 'random'; return S; }
    if(unk.length && (I.follow || !I.money)){ S.intent = 'fallback'; return S; }
    if(I.money || P || bk || I.why){ S.intent = 'total'; S.unk = unk; S.assert = P0 ? null : assertDir(qn); return S; }
    S.intent = 'fallback';
    return S;
  }
  function answer(raw){
    var S;
    try{
      S = route(raw);
      if(S.intent === 'menu') return { menu:true };
      var R = exec(S);
      if(S.intent === 'total' && S.unk && S.unk.length){
        R.html = '<p class="fb-dim">I could not find "' + esc(S.unk.join(' ')) + '" among this board\'s categories, accounts, vendors or people, so this is the ' + esc(BU) + ' total.</p>' + R.html;
      }
      if(S.intent === 'entity' && S.alts && S.alts.length){
        R.html += '<p class="fb-dim">Also matched: ' + S.alts.map(function(a){ return esc(a.name) + ' (' + esc(typeLabel(a)) + ')'; }).join(', ') + '.</p>';
        R.chips = S.alts.map(function(a){ return a.name + ' (' + typeLabel(a).toLowerCase() + ')'; }).concat(R.chips || []).slice(0, 4);
      }
      if(S.intent !== 'fallback' && S.intent !== 'help') CTX = S;
      return R;
    }catch(err){
      if(window.console) console.warn('FP&A bot error', err);
      return { html:'<p>Something went wrong while computing that answer. Try rephrasing, or ask for "help".</p>', chips:['Help'] };
    }
  }

  /* ---------------- sections (guided menu) ---------------- */
  var SECTION = null;
  function bestCat(P, bk){
    var l = CATS.map(function(c){ return { c:c, v:ev(c, P, bk).v }; }).sort(function(a, b){ return Math.abs(b.v) - Math.abs(a.v); });
    return l[0] && Math.abs(l[0].v) >= 1 ? l[0].c : null;
  }
  function periodChips(P, bk){
    var c = bestCat(P, bk);
    return ['Top 5 unfavorable vendors', 'Top 5 favorable vendors', 'Top 5 unfavorable accounts', c ? 'Why is ' + c.name + ' off ' + (bk === 'f' ? 'forecast' : 'plan') + '?' : null].filter(Boolean);
  }
  function aExceptions(){
    var P = { t:'ytd' };
    var lines = VROWS.map(function(x){ return withB(rowVal(x.row, P), 'p'); });
    var unpl = lines.filter(function(x){ return Math.abs(x.p) < 1 && Math.abs(x.w) >= 1; });
    var unus = lines.filter(function(x){ return Math.abs(x.w) < 1 && x.p >= 1; });
    var ytd = ev(TOTAL, P, 'p'), fy = ev(TOTAL, { t:'fy' }, 'p');
    var used = fy.p ? ytd.w / fy.p * 100 : null, elapsed = (RM + 1) / 12 * 100;
    var h = '<p>Three checks on <b>' + esc(pLabel(P)) + '</b> vs ' + esc(PLAN) + ':</p><ul class="fb-ul">';
    h += '<li><b>Unplanned spend:</b> ' + unpl.length + ' vendor/account lines with Working but no Plan, totaling <b>' + fm(unpl.reduce(function(a, x){ return a + x.w; }, 0)) + '</b>.</li>';
    h += '<li><b>Unused budget:</b> ' + unus.length + ' lines with Plan but no Working, <b>' + fm(unus.reduce(function(a, x){ return a + x.p; }, 0)) + '</b> not used so far.</li>';
    if(used !== null) h += '<li><b>Pacing:</b> ' + used.toFixed(1) + '% of the full-year Plan used vs ' + elapsed.toFixed(0) + '% of the year elapsed (' + (Math.abs(used - elapsed) < 3 ? 'on pace' : used > elapsed ? '<span class="fb-bad">ahead of pace</span>' : '<span class="fb-good">behind pace</span>') + ').</li>';
    h += '</ul>';
    return { html:h, chips:['What is not in plan?', 'Budget with no spend', 'Budget utilization', 'Monthly trend'] };
  }
  function secNotes(sid){ return notesHtml(boardNotes().filter(function(n){ return n.sid === sid; })); }
  var PQ = { t:'q', i:RQ }, PF = { t:'fy' };
  var SECTIONS = [
    { id:'overview', icon:'ti-layout-dashboard', title:'Overview', sub:'KPIs, drivers and watch items', sid:'sec-overview', P:null, bk:null,
      scope:'the whole board', hint:'Ask about the overall picture...', intro:aSummary },
    { id:'qplan', icon:'ti-trending-down', title:QL[RQ] + ' vs Plan', sub:'Quarter variance by category, account and vendor', sid:'sec-qplan', P:PQ, bk:'p',
      scope:QL[RQ] + ' vs ' + PLAN, hint:'Ask about ' + QL[RQ] + ' vs Plan...', intro:function(){ var R = aTotal(PQ, 'p'); R.html += secNotes('sec-qplan'); R.chips = periodChips(PQ, 'p'); return R; } },
    { id:'qfcst', icon:'ti-chart-dots-3', title:QL[RQ] + ' vs Forecast', sub:'Quarter variance vs ' + FCST, sid:'sec-qfcst', P:PQ, bk:'f',
      scope:QL[RQ] + ' vs ' + FCST, hint:'Ask about ' + QL[RQ] + ' vs Forecast...', intro:function(){ var R = aTotal(PQ, 'f'); R.html += secNotes('sec-qfcst'); R.chips = periodChips(PQ, 'f'); return R; } },
    { id:'fyplan', icon:'ti-calendar-stats', title:'Full year vs Plan', sub:FY + ' outlook vs ' + PLAN, sid:'sec-fyplan', P:PF, bk:'p',
      scope:'full year vs ' + PLAN, hint:'Ask about the full year vs Plan...', intro:function(){ var R = aTotal(PF, 'p'); R.html += secNotes('sec-fyplan'); R.chips = periodChips(PF, 'p'); return R; } },
    { id:'fyfcst', icon:'ti-chart-line', title:'Full year vs Forecast', sub:FY + ' outlook vs ' + FCST, sid:'sec-fyfcst', P:PF, bk:'f',
      scope:'full year vs ' + FCST, hint:'Ask about the full year vs Forecast...', intro:function(){ var R = aTotal(PF, 'f'); R.html += secNotes('sec-fyfcst'); R.chips = periodChips(PF, 'f'); return R; } },
    { id:'buffer', icon:'ti-arrows-split-2', title:'Real vs projected', sub:'Actuals, buffer (reallocated savings) and projection', sid:null, P:PF, bk:'p',
      scope:'full year vs ' + PLAN, hint:'Ask about buffer, actuals or projection...', intro:function(){ return aBuffer(null, PF, 'p'); } },
    { id:'exceptions', icon:'ti-alert-triangle', title:'Exceptions & pacing', sub:'Unplanned spend, unused budget, utilization', sid:null, P:{ t:'ytd' }, bk:'p',
      scope:'YTD vs ' + PLAN, hint:'Ask about unplanned spend, unused budget...', intro:aExceptions },
    { id:'hc', icon:'ti-users', title:'Headcount', sub:'Salary Accrued, TBH roles, hires', sid:'sec-hc', P:null, bk:null,
      scope:null, hint:'Ask about headcount, a TBH or an employee...', intro:function(){ return aHC(null); } },
    { id:'te', icon:'ti-plane', title:'Travel & Expense', sub:'T&E by month and employee', sid:'sec-te', P:null, bk:null,
      scope:null, hint:'Ask about T&E or an employee...', intro:aTE },
    { id:'actions', icon:'ti-checklist', title:'Follow-up actions', sub:'Open items on the board', sid:'sec-actions', P:null, bk:null,
      scope:null, hint:'Ask anything...', intro:aActions },
    { id:'free', icon:'ti-message-question', title:'Ask anything', sub:'Free questions across the whole board', sid:null, P:null, bk:null,
      scope:null, hint:'Ask about a vendor, account, category, quarter...', intro:function(){ var R = aHelp(); return R; } }
  ];

  if(NV) SECTIONS.splice(6, 0, { id:'nv', icon:'ti-list-details', title:'No Vendor detail', sub:'Line items causing the No Vendor variance', sid:null, P:PF, bk:'p',
    scope:'full year vs ' + PLAN, hint:'Ask about a No Vendor line item or account...', intro:function(){ return aNV(null, PF, 'p'); } });

  /* ---------------- UI ---------------- */
  var CSS = ''
    + '#fpa-bot-root{font-family:inherit}'
    + '.fb-fab{position:fixed;right:22px;bottom:22px;z-index:690;display:inline-flex;align-items:center;gap:8px;border:none;border-radius:999px;padding:12px 18px;background:linear-gradient(90deg,#4f46e5,#6366f1);color:#fff;font-size:13px;font-weight:800;cursor:pointer;box-shadow:0 8px 24px rgba(79,70,229,.38);transition:transform .15s}'
    + '.fb-fab:hover{transform:translateY(-2px)}.fb-fab i{font-size:18px}'
    + '.fb-panel{position:fixed;right:22px;bottom:80px;z-index:700;width:min(440px,calc(100vw - 32px));height:min(660px,calc(100vh - 104px));background:#fff;border:1px solid #e9eef7;border-radius:18px;box-shadow:0 18px 48px rgba(15,23,42,.22);display:none;flex-direction:column;overflow:hidden}'
    + '.fb-panel.open{display:flex}'
    + '.fb-head{padding:14px 16px;background:linear-gradient(120deg,#1e2450,#312e81);color:#fff;display:flex;align-items:center;gap:10px}'
    + '.fb-head .fb-av{width:34px;height:34px;border-radius:10px;background:rgba(255,255,255,.14);display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0}'
    + '.fb-face{width:28px;height:28px;display:block;overflow:visible;animation:fb-bob 3.2s ease-in-out infinite}'
    + '.fb-face .fb-eyes{transform-box:fill-box;transform-origin:center;animation:fb-blink 4.2s infinite}'
    + '.fb-face .fb-smile{transform-box:fill-box;transform-origin:center top;animation:fb-grin 6.4s ease-in-out infinite}'
    + '.fb-face .fb-ant{animation:fb-glow 2.4s ease-in-out infinite}'
    + '@keyframes fb-blink{0%,90%,100%{transform:scaleY(1)}93%{transform:scaleY(.12)}96%{transform:scaleY(1)}}'
    + '@keyframes fb-grin{0%,55%,100%{transform:scale(1,1)}65%,85%{transform:scale(1.18,1.45)}}'
    + '@keyframes fb-glow{0%,100%{opacity:.55}50%{opacity:1}}'
    + '@keyframes fb-bob{0%,100%{transform:translateY(0)}50%{transform:translateY(-1.5px)}}'
    + '.fb-face .fb-eyes-c,.fb-face .fb-hand{opacity:0;transition:opacity .15s}'
    + '.fb-face .fb-hand{transform-box:fill-box;transform-origin:70% 100%}'
    + '.fb-face .fb-ear-r{transition:opacity .15s}.fb-thinking .fb-face .fb-ear-r{opacity:0}'
    + '.fb-thinking .fb-face{animation:fb-nod .84s ease-in-out infinite}'
    + '.fb-thinking .fb-face .fb-eyes{opacity:0;animation:none}'
    + '.fb-thinking .fb-face .fb-eyes-c{opacity:1}'
    + '.fb-thinking .fb-face .fb-smile{animation:none;transform:scale(1.18,1.45)}'
    + '.fb-thinking .fb-face .fb-hand{opacity:1;animation:fb-tap .42s ease-in-out infinite}'
    + '@keyframes fb-tap{0%,100%{transform:translate(0,0) rotate(0)}50%{transform:translate(-.6px,.4px) rotate(-11deg)}}'
    + '@keyframes fb-nod{0%,100%{transform:rotate(0)}50%{transform:rotate(-4deg)}}'
    + '.fb-typing{display:flex;align-items:center;gap:4px;color:#64748b;font-size:11.5px;font-style:italic}'
    + '.fb-typing i{width:6px;height:6px;border-radius:50%;background:#a5b4fc;animation:fb-dot 1s infinite}.fb-typing i:nth-child(2){animation-delay:.15s}.fb-typing i:nth-child(3){animation-delay:.3s}.fb-typing span{margin-left:4px}'
    + '@keyframes fb-dot{0%,80%,100%{opacity:.35;transform:translateY(0)}40%{opacity:1;transform:translateY(-3px)}}'
    + '@media (prefers-reduced-motion:reduce){.fb-face,.fb-face *,.fb-typing i{animation:none!important}}'
    + '.fb-head .fb-tt{font-size:14px;font-weight:800;line-height:1.2}.fb-head .fb-st{font-size:11px;color:#c7d2fe;margin-top:2px}'
    + '.fb-hbtns{margin-left:auto;display:flex;align-items:center;gap:2px}.fb-clr{font-size:17px!important}'
    + '.fb-head button{background:transparent;border:none;color:#c7d2fe;font-size:20px;cursor:pointer;padding:2px 6px;border-radius:8px}.fb-head button:hover{background:rgba(255,255,255,.1);color:#fff}'
    + '.fb-msgs{flex:1;overflow-y:auto;padding:14px;background:#f7f9fd;display:flex;flex-direction:column;gap:10px}'
    + '.fb-msg{max-width:94%;font-size:12.5px;line-height:1.5;color:#1e293b;border-radius:14px;padding:10px 12px}'
    + '.fb-msg.bot{background:#fff;border:1px solid #e9eef7;align-self:flex-start;box-shadow:0 2px 8px rgba(15,23,42,.04)}'
    + '.fb-msg.me{background:linear-gradient(90deg,#4f46e5,#6366f1);color:#fff;align-self:flex-end;border-bottom-right-radius:4px}'
    + '.fb-msg p{margin:0 0 8px}.fb-msg p:last-child{margin-bottom:0}'
    + '.fb-h{font-size:13px;font-weight:800;color:#0f172a;margin-bottom:6px}.fb-h small{font-size:10.5px;font-weight:700;color:#94a3b8;margin-left:4px}'
    + '.fb-sub{font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#94a3b8;margin:12px 0 6px}'
    + '.fb-t{width:100%;border-collapse:collapse;font-size:11.5px;margin:6px 0 8px}'
    + '.fb-t th{font-size:10px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.03em;background:#f1f5fb;padding:6px 7px;text-align:right;border-bottom:1px solid #e9eef7;white-space:nowrap}'
    + '.fb-t th:first-child{text-align:left}'
    + '.fb-t td{padding:6px 7px;border-bottom:1px solid #f1f5f9;text-align:right;white-space:nowrap;color:#334155}'
    + '.fb-t td:first-child{text-align:left;white-space:normal;font-weight:600;color:#0f172a}'
    + '.fb-t td small{display:block;font-size:10px;font-weight:600;color:#94a3b8}'
    + '.fb-t.fb-tw{font-size:10.5px;table-layout:auto}.fb-t.fb-tw th{font-size:8.5px;letter-spacing:.02em;padding:5px 4px;white-space:normal;line-height:1.2;vertical-align:bottom}.fb-t.fb-tw td{padding:5px 4px}.fb-t.fb-tw td:first-child{font-size:10.5px;padding-left:5px}'
    + '.fb-msg.bot{min-width:0;overflow-x:auto}.fb-msg,.fb-chips{flex-shrink:0}'
    + '.fb-t tr.fb-hl td{background:#eef2ff}.fb-t tr.fb-more td{color:#94a3b8;font-style:italic;text-align:left}'
    + '.fb-good{color:#16a34a;font-weight:700}.fb-bad{color:#e11d48;font-weight:700}.fb-neu{color:#475569;font-weight:700}'
    + '.fb-tag{display:inline-block;font-size:10px;font-weight:800;border-radius:999px;padding:1px 8px;margin-left:2px;vertical-align:1px}'
    + '.fb-tag.fb-good{background:#e8f8ee}.fb-tag.fb-bad{background:#fdeaef}.fb-tag.fb-neu{background:#eaf1ff;color:#2563eb}'
    + '.fb-dim{color:#64748b;font-size:11.5px}'
    + '.fb-ul{margin:4px 0 8px 18px;padding:0}.fb-ul li{margin:3px 0}.fb-ul li.fb-done{color:#94a3b8;text-decoration:line-through}'
    + '.fb-note{background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:7px 9px;margin:6px 0;font-size:11.5px;color:#422006}'
    + '.fb-note span{display:block;font-size:9.5px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:#a16207;margin-bottom:2px}'
    + '.fb-chips{display:flex;flex-wrap:wrap;gap:6px;align-self:flex-start;max-width:100%}'
    + '.fb-chip{border:1px solid #c7d2fe;background:#fff;color:#4338ca;border-radius:999px;padding:5px 10px;font-size:11px;font-weight:700;cursor:pointer;text-align:left}'
    + '.fb-chip:hover{background:#eef2ff}'
    + '.fb-in{display:flex;gap:8px;padding:10px;border-top:1px solid #e9eef7;background:#fff}'
    + '.fb-in input{flex:1;border:1px solid #cbd5e1;border-radius:10px;padding:10px 12px;font-size:12.5px;font-family:inherit;color:#0f172a;min-width:0}'
    + '.fb-in input:focus{outline:2px solid #c7d2fe;border-color:transparent}'
    + '.fb-in button{border:none;border-radius:10px;padding:0 14px;background:linear-gradient(90deg,#4f46e5,#6366f1);color:#fff;font-size:16px;cursor:pointer}'
    + '.fb-ctx{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:7px 12px;border-top:1px solid #e9eef7;background:#f8faff;font-size:11.5px;color:#475569}'
    + '.fb-ctx-l{display:flex;align-items:center;gap:6px;min-width:0}.fb-ctx-l i{color:#6366f1;font-size:14px}.fb-ctx-t{font-weight:800;color:#312e81;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
    + '.fb-ctx-btn{display:inline-flex;align-items:center;gap:5px;flex-shrink:0;border:1px solid #c7d2fe;background:#fff;color:#4338ca;border-radius:8px;padding:4px 9px;font-size:11px;font-weight:800;cursor:pointer;font-family:inherit}.fb-ctx-btn:hover{background:#eef2ff}'
    + '.fb-menu{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:8px 0}'
    + '.fb-card{display:flex;gap:9px;align-items:flex-start;text-align:left;border:1px solid #e2e8f0;background:#fafbff;border-radius:12px;padding:10px;cursor:pointer;font-family:inherit;transition:.12s}'
    + '.fb-card:hover{border-color:#818cf8;background:#eef2ff}.fb-card.on{border-color:#4f46e5;background:#eef2ff;box-shadow:inset 0 0 0 1px #4f46e5}.fb-menu .fb-card:last-child:nth-child(odd){grid-column:1 / -1}'
    + '.fb-t.fb-bridge td:first-child{font-weight:600}.fb-t tr.fb-bufrow td{background:#f8fafc}.fb-t tr.fb-bufrow td:first-child{font-style:italic}'
    + '.fb-note-q{background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:7px 9px;font-size:11.5px;color:#312e81}'
    + '.fb-t tr.fb-tot td{font-weight:800;border-top:1.5px solid #e2e8f0}'
    + '.fb-card i{font-size:18px;color:#4f46e5;margin-top:1px;flex-shrink:0}.fb-card b{display:block;font-size:12px;color:#0f172a;line-height:1.25}.fb-card small{display:block;font-size:10.5px;color:#64748b;margin-top:3px;line-height:1.35}'
    + '.fb-sec-tag{display:inline-flex;align-items:center;gap:5px;font-size:10px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;color:#4338ca;background:#eef2ff;border-radius:999px;padding:3px 9px;margin-bottom:8px}'
    + '.fb-jump{display:inline-flex;align-items:center;gap:5px;border:none;background:none;color:#4f46e5;font-weight:800;font-size:11.5px;cursor:pointer;padding:0;margin:2px 0 8px;font-family:inherit}.fb-jump:hover{text-decoration:underline}'
    + '.fb-modal{position:absolute;inset:0;z-index:5;background:rgba(15,23,42,.38);display:flex;align-items:center;justify-content:center;padding:20px}.fb-modal[hidden]{display:none}'
    + '.fb-mbox{background:#fff;border-radius:16px;padding:20px 18px 16px;max-width:320px;width:100%;box-shadow:0 18px 40px rgba(15,23,42,.25);text-align:center}'
    + '.fb-mface{width:64px;height:64px;border-radius:18px;background:linear-gradient(120deg,#1e2450,#312e81);display:flex;align-items:center;justify-content:center;margin:0 auto 12px;box-shadow:0 6px 18px rgba(49,46,129,.28)}.fb-mface .fb-face{width:46px;height:46px}'
    + '.fb-face .fb-wave{opacity:0;transition:opacity .15s;transform-box:fill-box;transform-origin:40% 100%}'
    + '.fb-bye .fb-face .fb-wave{opacity:1;animation:fb-wave .38s ease-in-out infinite alternate}'
    + '.fb-bye .fb-face .fb-eyes{opacity:0;animation:none}.fb-bye .fb-face .fb-eyes-c{opacity:1}.fb-bye .fb-face .fb-ear-r{opacity:0}'
    + '.fb-bye .fb-face .fb-smile{animation:none;transform:scale(1.18,1.45)}'
    + '.fb-bye .fb-mact{display:none}.fb-bye .fb-mt{font-size:16px}'
    + '@keyframes fb-wave{from{transform:rotate(-14deg)}to{transform:rotate(16deg)}}'
    + '.fb-mi{width:42px;height:42px;border-radius:12px;background:#fff7ed;color:#ea580c;display:flex;align-items:center;justify-content:center;font-size:22px;margin:0 auto 10px}'
    + '.fb-mt{font-size:14px;font-weight:800;color:#0f172a;margin-bottom:6px}.fb-mp{font-size:12px;line-height:1.5;color:#475569;margin:0 0 16px}'
    + '.fb-mact{display:flex;gap:8px;justify-content:center;align-items:center}'
    + '.fb-mbtn{border-radius:10px;padding:9px 14px!important;font-size:12px!important;line-height:1.2!important;height:auto!important;white-space:nowrap;font-weight:800;cursor:pointer;font-family:inherit}'
    + '.fb-mbtn.fb-mb-sec{border:1px solid #cbd5e1;background:#fff;color:#334155}.fb-mbtn.fb-mb-sec:hover{background:#f1f5f9}'
    + '.fb-mbtn.fb-mb-pri{border:1px solid transparent;background:linear-gradient(90deg,#4f46e5,#6366f1);color:#fff}'
    + '.toast{right:auto!important;left:50%;transform:translate(-50%,10px)!important}.toast.show{transform:translate(-50%,0)!important}'
    + '@media print{#fpa-bot-root{display:none!important}}';
  var st = d.createElement('style'); st.id = 'fpa-bot-style'; st.textContent = CSS; (d.head || d.documentElement).appendChild(st);

  var FACE = '<svg class="fb-face" viewBox="0 0 32 32" aria-hidden="true"><line x1="16" y1="3.5" x2="16" y2="7.5" stroke="#c7d2fe" stroke-width="1.6" stroke-linecap="round"/><circle class="fb-ant" cx="16" cy="3" r="2" fill="#a5f3fc"/><rect x="5" y="7.5" width="22" height="19" rx="7" fill="#eef2ff"/><rect x="2.6" y="14" width="2.6" height="6" rx="1.3" fill="#c7d2fe"/><rect class="fb-ear-r" x="26.8" y="14" width="2.6" height="6" rx="1.3" fill="#c7d2fe"/><g class="fb-eyes"><ellipse cx="12" cy="15.5" rx="2.2" ry="2.6" fill="#312e81"/><ellipse cx="20" cy="15.5" rx="2.2" ry="2.6" fill="#312e81"/><circle cx="12.8" cy="14.6" r=".7" fill="#fff"/><circle cx="20.8" cy="14.6" r=".7" fill="#fff"/></g><circle cx="9.4" cy="20.2" r="1.5" fill="#fda4af" opacity=".75"/><circle cx="22.6" cy="20.2" r="1.5" fill="#fda4af" opacity=".75"/><path class="fb-smile" d="M12 20.6 Q16 24.6 20 20.6" fill="none" stroke="#312e81" stroke-width="1.8" stroke-linecap="round"/><g class="fb-eyes-c" fill="none" stroke="#312e81" stroke-width="1.7" stroke-linecap="round"><path d="M9.8 16.4 Q12 13.9 14.2 16.4"/><path d="M17.8 16.4 Q20 13.9 22.2 16.4"/></g><g class="fb-hand"><rect x="29.4" y="15.2" width="3.4" height="9" rx="1.7" fill="#c7d2fe" stroke="#312e81" stroke-width=".7"/><rect x="24.6" y="10.6" width="6" height="2.6" rx="1.3" fill="#fef3c7" stroke="#312e81" stroke-width=".7" transform="rotate(28 30.6 11.9)"/><circle cx="31.1" cy="14.4" r="3.4" fill="#fef3c7" stroke="#312e81" stroke-width=".7"/><path d="M29.6 15.6 Q31.1 16.6 32.6 15.6" fill="none" stroke="#312e81" stroke-width=".55" stroke-linecap="round"/></g><g class="fb-wave"><rect x="28.9" y="9" width="3.4" height="12.5" rx="1.7" fill="#c7d2fe" stroke="#312e81" stroke-width=".7" transform="rotate(14 30.6 21.5)"/><path d="M29.2 4.6 L28.7 1.4 M31.3 3.9 L31.3 0.4 M33.4 4.6 L34 1.4 M28.2 7.6 L26.3 6.4" stroke="#312e81" stroke-width="1.5" stroke-linecap="round"/><path d="M29.2 4.6 L28.7 1.4 M31.3 3.9 L31.3 0.4 M33.4 4.6 L34 1.4 M28.2 7.6 L26.3 6.4" stroke="#fef3c7" stroke-width=".7" stroke-linecap="round"/><circle cx="31.3" cy="7.2" r="3.7" fill="#fef3c7" stroke="#312e81" stroke-width=".7"/></g></svg>';
  var root = d.createElement('div'); root.id = 'fpa-bot-root';
  root.innerHTML = '<button type="button" class="fb-fab" aria-label="Ask Felipe, the FP&amp;A analyst"><i class="ti ti-message-chatbot"></i> Ask FP&amp;A</button>'
    + '<div class="fb-panel" role="dialog" aria-label="Felipe, FP&amp;A analyst">'
    + '<div class="fb-head"><div class="fb-av">' + FACE + '</div><div><div class="fb-tt">Felipe Analyst</div><div class="fb-st">' + esc(BU) + ' · ' + esc(META.month || '') + ' ' + esc(FY) + ' · ' + esc(QL[RQ]) + '</div></div><div class="fb-hbtns"><button type="button" class="fb-clr" title="Clear chat" aria-label="Clear chat"><i class="ti ti-eraser"></i></button><button type="button" class="fb-x" title="Close" aria-label="Close">&times;</button></div></div>'
    + '<div class="fb-msgs"></div>'
    + '<div class="fb-ctx"><span class="fb-ctx-l"><i class="ti ti-target-arrow"></i><span>Section:</span><span class="fb-ctx-t">Whole board</span></span><button type="button" class="fb-ctx-btn"><i class="ti ti-layout-grid"></i> Sections</button></div>'
    + '<div class="fb-modal" hidden><div class="fb-mbox"><div class="fb-mface">' + FACE + '</div><div class="fb-mt"></div><p class="fb-mp"></p><div class="fb-mact"><button type="button" class="fb-mbtn fb-mb-sec"></button><button type="button" class="fb-mbtn fb-mb-pri"></button></div></div></div>'
    + '<form class="fb-in"><input type="text" placeholder="Ask about a vendor, account, category, quarter..." autocomplete="off" /><button type="submit" aria-label="Send"><i class="ti ti-send"></i></button></form>'
    + '</div>';
  d.body.appendChild(root);
  // A board saved from an open window carries a stale copy of the bot after this script; drop it once parsing ends.
  function dedupe(){
    d.querySelectorAll('#fpa-bot-root').forEach(function(el){ if(el !== root && el.parentNode) el.parentNode.removeChild(el); });
    d.querySelectorAll('#fpa-bot-style').forEach(function(el){ if(el !== st && el.parentNode) el.parentNode.removeChild(el); });
  }
  if(d.readyState === 'loading') d.addEventListener('DOMContentLoaded', dedupe); else dedupe();
  var fab = root.querySelector('.fb-fab'), panel = root.querySelector('.fb-panel'), msgs = root.querySelector('.fb-msgs');
  var form = root.querySelector('.fb-in'), input = form.querySelector('input'), started = false;

  function addMsg(who, html, chips){
    var m = d.createElement('div'); m.className = 'fb-msg ' + who;
    if(who === 'me') m.textContent = html; else m.innerHTML = html;
    msgs.appendChild(m);
    var old = msgs.querySelectorAll('.fb-chips'); old.forEach(function(c){ c.parentNode.removeChild(c); });
    if(chips && chips.length){
      var c = d.createElement('div'); c.className = 'fb-chips';
      chips.forEach(function(t){ var b = d.createElement('button'); b.type = 'button'; b.className = 'fb-chip'; b.textContent = t; c.appendChild(b); });
      msgs.appendChild(c);
    }
    if(who === 'me') msgs.scrollTop = msgs.scrollHeight; else m.scrollIntoView({ block:'start' });
  }
  var ctxT = root.querySelector('.fb-ctx-t'), DEF_HINT = input.placeholder;
  function updateCtx(){
    ctxT.textContent = SECTION ? SECTION.title : 'Whole board';
    input.placeholder = SECTION ? SECTION.hint : DEF_HINT;
  }
  function menuHtml(){
    return '<div class="fb-h">What do you want to review?</div><div class="fb-menu">'
      + SECTIONS.map(function(sx){
          var on = !!(SECTION && SECTION.id === sx.id);
          return '<button type="button" class="fb-card' + (on ? ' on' : '') + '" data-sec="' + sx.id + '"><i class="ti ' + sx.icon + '"></i><span><b>' + esc(sx.title) + '</b><small>' + esc(sx.sub) + '</small></span></button>';
        }).join('')
      + '</div><p class="fb-dim">Come back here anytime with <b>Sections</b> below or by typing <b>menu</b>.</p>';
  }
  function showMenu(){ addMsg('bot', menuHtml(), null); }
  // "Thinking" beat: Felipe closes his eyes, smiles and taps his head before answering.
  var THINK_MS = 750, thinkTok = 0, busy = false;
  function think(fn){
    var tok = ++thinkTok; busy = true; root.classList.add('fb-thinking');
    var ty = d.createElement('div'); ty.className = 'fb-msg bot fb-typing';
    ty.innerHTML = '<i></i><i></i><i></i><span>Felipe is thinking...</span>';
    msgs.appendChild(ty); msgs.scrollTop = msgs.scrollHeight;
    setTimeout(function(){
      if(ty.parentNode) ty.parentNode.removeChild(ty);
      if(tok !== thinkTok) return;
      root.classList.remove('fb-thinking'); busy = false; fn();
    }, THINK_MS);
  }
  function enterSection(id){
    var sx = SECTIONS.filter(function(x){ return x.id === id; })[0]; if(!sx) return;
    SECTION = sx.id === 'free' ? null : sx; CTX = null; updateCtx();
    addMsg('me', '→ ' + sx.title);
    think(function(){ showSection(sx); });
  }
  function showSection(sx){
    var R;
    try{ R = sx.intro(); }catch(err){ if(window.console) console.warn('FP&A bot error', err); R = { html:'<p>Could not build this section.</p>', chips:[] }; }
    var html = '<div class="fb-sec-tag"><i class="ti ' + sx.icon + '"></i> ' + esc(sx.title) + '</div>' + R.html;
    if(sx.sid && d.getElementById(sx.sid)) html += '<button type="button" class="fb-jump" data-jump="' + sx.sid + '"><i class="ti ti-arrow-up-right"></i> View this section on the board</button>';
    if(sx.scope) html += '<p class="fb-dim">Questions you type now default to <b>' + esc(sx.scope) + '</b> unless you name another period or benchmark.</p>';
    addMsg('bot', html, R.chips);
    setTimeout(function(){ input.focus(); }, 30);
  }
  function ask(q){
    q = String(q || '').trim(); if(!q || busy) return;
    addMsg('me', q);
    think(function(){
      var R = answer(q);
      if(R.menu){ showMenu(); return; }
      addMsg('bot', R.html, R.chips);
    });
  }
  function open(){
    panel.classList.add('open'); fab.style.display = 'none';
    if(!started){
      started = true;
      addMsg('bot', '<p>Hi, I\'m <b>Felipe</b>, the FP&amp;A analyst for the <b>' + esc(BU) + ' ' + esc(META.month || '') + ' ' + esc(FY) + '</b> board. I compute every answer from this board\'s feeds and comments. I don\'t use an AI model and nothing leaves your browser.</p><p class="fb-dim">Pick a section to focus on, or just type a question about any category, GL account, vendor, month, quarter, YTD or full year.</p>', null);
      showMenu();
    }
    setTimeout(function(){ input.focus(); }, 30);
  }
  function hide(){ panel.classList.remove('open'); fab.style.display = ''; }
  function reset(){
    thinkTok++; busy = false; root.classList.remove('fb-thinking');
    msgs.innerHTML = ''; started = false; SECTION = null; CTX = null; input.value = ''; updateCtx();
  }
  function hasHistory(){ return !!msgs.querySelector('.fb-msg.me'); }
  var modal = root.querySelector('.fb-modal'), modalOk = null;
  var byeTimer = null;
  function confirmBox(title, text, okLabel, onOk, cancelLabel){
    modal.classList.remove('fb-bye');
    modal.querySelector('.fb-mt').textContent = title;
    modal.querySelector('.fb-mp').textContent = text;
    modal.querySelector('.fb-mbtn.fb-mb-pri').textContent = okLabel;
    modal.querySelector('.fb-mbtn.fb-mb-sec').textContent = cancelLabel || 'Cancel';
    modalOk = onOk; modal.hidden = false;
    modal.querySelector('.fb-mbtn.fb-mb-sec').focus();
  }
  function closeModal(){ if(byeTimer) return; modal.hidden = true; modalOk = null; modal.classList.remove('fb-bye'); }
  // Felipe waves goodbye, then the session closes and is cleared.
  function goodbye(){
    modal.classList.add('fb-bye');
    modal.querySelector('.fb-mt').textContent = 'Goodbye!';
    modal.querySelector('.fb-mp').textContent = 'Thanks for stopping by. See you at the next close.';
    modal.hidden = false;
    byeTimer = setTimeout(function(){
      byeTimer = null; reset(); hide();
      modal.hidden = true; modalOk = null; modal.classList.remove('fb-bye');
    }, 1600);
    return true;
  }
  modal.querySelector('.fb-mbtn.fb-mb-sec').addEventListener('click', closeModal);
  modal.querySelector('.fb-mbtn.fb-mb-pri').addEventListener('click', function(){ var f = modalOk; modalOk = null; var keep = f ? f() : false; if(!keep) closeModal(); });
  modal.addEventListener('click', function(e){ if(e.target === modal && !byeTimer) closeModal(); });
  // Closing ends the chat session: history is cleared (with a warning if there is any).
  function close(){
    if(byeTimer) return;
    if(!hasHistory()){ reset(); hide(); return; }
    confirmBox('Close our session?', 'Heads up: if we close it, I\'ll clear our question history and we\'ll start fresh next time.', 'Close & clear', goodbye, 'Keep chatting');
  }
  function clearChat(){
    if(!hasHistory()){ reset(); open(); return; }
    confirmBox('Start over?', 'I\'ll wipe our conversation and take you back to the sections menu.', 'Clear chat', function(){ reset(); open(); }, 'Cancel');
  }
  fab.addEventListener('click', open);
  root.querySelector('.fb-x').addEventListener('click', close);
  root.querySelector('.fb-clr').addEventListener('click', clearChat);
  form.addEventListener('submit', function(e){ e.preventDefault(); if(busy) return; var q = input.value; input.value = ''; ask(q); });
  msgs.addEventListener('click', function(e){
    if(!e.target.closest) return;
    var b = e.target.closest('.fb-chip'); if(b){ ask(b.textContent); return; }
    if(busy) return;
    var c = e.target.closest('[data-sec]'); if(c){ enterSection(c.getAttribute('data-sec')); return; }
    var j = e.target.closest('[data-jump]'); if(j){ var el = d.getElementById(j.getAttribute('data-jump')); if(el) window.scrollTo({ top:el.getBoundingClientRect().top + window.pageYOffset - 20, behavior:'smooth' }); }
  });
  root.querySelector('.fb-ctx-btn').addEventListener('click', showMenu);
  panel.addEventListener('keydown', function(e){ if(e.key === 'Escape'){ if(!modal.hidden) closeModal(); else close(); } });

  var api = { ask:answer, open:open, close:close, clear:clearChat, section:enterSection, menu:showMenu };
  try{ window.__fpaBot = api; }catch(e){}
  return api;
}

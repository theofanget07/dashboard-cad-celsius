/* ============================================================
   Dashboard CAD — Groupe E Celsius
   script.js  v4.4  — 18/02/2026
   Gantt: 3 datasets groupés, barThickness 18px fixe, centrage parfait
   ============================================================ */

// ---- Globals ----
let prevRows = [], realRows = [];
let chart, chartBar, chartPieBudgetRev, chartPieAT;
let chartHonoMetier, chartGantt;
let manualPlanningDates = {};

// ---- DOM refs ----
const prevInput    = document.getElementById('prevFile');
const realInput    = document.getElementById('realFile');
const kSlider      = document.getElementById('coefK');
const kNumber      = document.getElementById('coefK_num');
const statusEl     = document.getElementById('status');
const viewGlobalEl = document.getElementById('viewGlobal');
const viewSOnlyEl  = document.getElementById('viewSOnly');
const kpiSection   = document.getElementById('kpiSection');
const kpiHono      = document.getElementById('kpiSectionHonoraires');
const tablePrev    = document.getElementById('tablePreview');
const tableCmp     = document.getElementById('tableCompare');
const tableRess    = document.getElementById('tableRessources');
const tablePlan    = document.getElementById('tablePlanning');

const PROJECT_NAMES = {
  'CL150096_11_05_01': 'Général - Direction',
  'CL150096_11_05_02': 'Centrale',
  'CL150096_11_05_03': 'Bâtiment',
  'CL150096_11_05_04': 'Réseau CAD',
  'CL150096_11_05_05': 'Sous-stations',
  'CL150096_11_05_06': 'Participations clients'
};
const PROJECT_ORDER = Object.keys(PROJECT_NAMES);
function getPN(id) { return PROJECT_NAMES[id] || id; }

// ---- Palette Groupe E ----
const GE_BLUE  = '#163a5f';
const GE_BLUE2 = '#1e4d8c';
const GE_BLUE3 = '#2e6db4';
const GE_BLUE4 = '#5b9bd5';
const GE_ORANGE = '#e8873a';
const GE_RED    = '#c0392b';
const GE_GREEN  = '#27ae60';

// Camembert — palette classique sobre
const PIE_COLORS = ['#2e6db4','#e8873a','#27ae60','#8e44ad','#c0392b','#16a085'];

// ============================================================
// UTILS
// ============================================================
function readExcel(file, cb) {
  const r = new FileReader();
  r.onload = e => {
    const wb    = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    cb(XLSX.utils.sheet_to_json(sheet, { raw: true }));
  };
  r.readAsArrayBuffer(file);
}

function toNum(v) {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string') return 0;
  return parseFloat(v.replace(/[\u00A0\u202F']/g, '').replace(/[ ]/g, '').replace(',', '.')) || 0;
}
function toCHF(n) { return (n || 0).toLocaleString('fr-CH', { minimumFractionDigits: 0 }); }
function toDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  if (typeof v === 'string') {
    const cleaned = v.trim().replace(' 00:00:00','').replace('T00:00:00','');
    let d = new Date(cleaned);
    if (!isNaN(d)) return d;
    const m = cleaned.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})$/);
    if (m) return new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2]-1, +m[1]);
  }
  return null;
}
function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function yyyymm(d)       { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; }
function buildAxis(minD, maxD) {
  const axis = [], c = firstOfMonth(minD), e = firstOfMonth(maxD);
  while (c <= e) { axis.push(yyyymm(c)); c.setMonth(c.getMonth()+1); }
  return axis;
}

function sCurve(total, axis, k, sIdx, eIdx) {
  const N = axis.length, cum = Array(N).fill(null), L = eIdx - sIdx + 1;
  if (L <= 0 || total === 0 || sIdx < 0 || eIdx < 0) return cum;
  const s0 = 1/(1+Math.exp(-k*(0-0.5))), s1 = 1/(1+Math.exp(-k*(1-0.5)));
  for (let j = 0; j < L; j++) {
    const i = sIdx + j; if (i >= N) break;
    const x = j/(L-1||1), sx = 1/(1+Math.exp(-k*(x-0.5)));
    cum[i] = ((sx-s0)/(s1-s0))*total;
  }
  return cum;
}

function modelRange(rows, modelName, axis) {
  const filtered = rows.filter(r => r['Modèle de prévision'] === modelName);
  const dates = filtered.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
  if (!dates.length) return [-1, -1];
  const minD = new Date(Math.min(...dates)), maxD = new Date(Math.max(...dates));
  const mmMin = yyyymm(firstOfMonth(minD)), mmMax = yyyymm(firstOfMonth(maxD));
  let sIdx = axis.indexOf(mmMin), eIdx = axis.indexOf(mmMax);
  if (sIdx < 0) sIdx = axis.findIndex(m => m >= mmMin);
  if (eIdx < 0) { for (let i = axis.length-1; i >= 0; i--) { if (axis[i] <= mmMax) { eIdx = i; break; } } }
  if (sIdx < 0) sIdx = 0; if (eIdx < 0) eIdx = axis.length-1;
  return [sIdx, eIdx];
}

// Plugin ligne "Aujourd'hui"
const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(ch, _args, opts) {
    if (!opts || !opts.label) return;
    const { ctx, chartArea: { top, bottom }, scales: { x } } = ch;
    const labels = ch.data.labels;
    const toT = s => { const [y,m] = s.split('-'); return new Date(+y,+m-1,1).getTime(); };
    const t = toT(opts.label);
    let bi = 0, bd = Infinity;
    labels.forEach((l,i) => { const d = Math.abs(toT(l)-t); if (d < bd) { bd=d; bi=i; } });
    const xp = x.getPixelForValue(bi);
    ctx.save();
    ctx.strokeStyle = GE_RED; ctx.lineWidth = 2; ctx.setLineDash([5,5]);
    ctx.beginPath(); ctx.moveTo(xp, top); ctx.lineTo(xp, bottom); ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = GE_RED; ctx.font = 'bold 11px Arial'; ctx.textAlign = 'center';
    ctx.fillText("Aujourd'hui", xp, top-5);
    ctx.restore();
  }
};

// ============================================================
// EVENTS
// ============================================================
prevInput.addEventListener('change', e => { if (e.target.files[0]) readExcel(e.target.files[0], rows => { prevRows = rows; maybeProcess(); }); });
realInput.addEventListener('change', e => { if (e.target.files[0]) readExcel(e.target.files[0], rows => { realRows = rows; maybeProcess(); }); });
[kSlider, kNumber].forEach(el => {
  el.addEventListener('input', () => { kSlider.value = kNumber.value = el.value; if (prevRows.length && realRows.length) processAll(); });
});
[viewGlobalEl, viewSOnlyEl].forEach(el => {
  el.addEventListener('change', () => { if (prevRows.length && realRows.length) processAll(); });
});
document.getElementById('btnApplyPlanning').addEventListener('click', () => {
  if (!prevRows.length || !realRows.length) return;
  PROJECT_ORDER.forEach(id => {
    const s = document.getElementById(`rs_${id}`), e = document.getElementById(`re_${id}`);
    if (!manualPlanningDates[id]) manualPlanningDates[id] = {};
    manualPlanningDates[id].start = s && s.value ? new Date(s.value) : null;
    manualPlanningDates[id].end   = e && e.value ? new Date(e.value) : null;
  });
  processPlanning();
});
document.getElementById('btnPDF').addEventListener('click', generatePDF);

function maybeProcess() {
  if (prevRows.length && realRows.length) { statusEl.textContent = '✅ Fichiers chargés — calcul en cours...'; statusEl.className = 'status'; processAll(); }
}
function processAll() { processBudgets(); processImputations(); processPlanning(); }

// ============================================================
// PARTIE 1 — BUDGETS
// ============================================================
function processBudgets() {
  const k = parseFloat(kNumber.value) || 12;
  const MODELS = ['Budget', 'Budget_Rev', 'AT'];
  const isSOnly = viewSOnlyEl.checked;

  const model = {};
  MODELS.forEach(mn => {
    const rows  = prevRows.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a,r) => a+toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = { total, start: dates.length ? new Date(Math.min(...dates)) : null, end: dates.length ? new Date(Math.max(...dates)) : null };
  });

  const realDates = realRows.map(r => toDate(r['Date du projet'])).filter(Boolean);
  const allDates  = [];
  MODELS.forEach(mn => { if(model[mn].start) allDates.push(model[mn].start); if(model[mn].end) allDates.push(model[mn].end); });
  realDates.forEach(d => allDates.push(d));
  if (!allDates.length) { statusEl.textContent = '⚠️ Aucune date trouvée.'; return; }

  const globalMin = new Date(Math.min(...allDates)), globalMax = new Date(Math.max(...allDates));
  const modelMin  = new Date(Math.min(...MODELS.filter(m=>model[m].start).map(m=>model[m].start)));
  const modelMax  = new Date(Math.max(...MODELS.filter(m=>model[m].end).map(m=>model[m].end)));
  const axisMin   = isSOnly ? modelMin : globalMin;
  const axisMax   = isSOnly ? modelMax : globalMax;
  const axis      = buildAxis(axisMin, axisMax);

  const series = {};
  MODELS.forEach(mn => { const [si,ei] = modelRange(prevRows, mn, axis); series[mn] = sCurve(model[mn].total, axis, k, si, ei); });

  const today   = new Date(), mmToday = yyyymm(today);
  const realMap = new Map(axis.map(m => [m, 0]));
  let   initCum = 0;
  realRows.forEach(r => {
    const d = toDate(r['Date du projet']); if (!d) return;
    const mm = yyyymm(firstOfMonth(d)); if (mm > mmToday) return;
    if (realMap.has(mm)) realMap.set(mm, realMap.get(mm)+toNum(r['Montant total du coût']));
    else if (isSOnly && mm < axis[0]) initCum += toNum(r['Montant total du coût']);
  });

  const realMonth = axis.map(m => realMap.get(m)||0);
  let acc = initCum; const realCum = realMonth.map(v => { acc+=v; return acc; });
  const realCut = realCum.map((v,i) => axis[i] > mmToday ? null : v);

  const totalBR = model['Budget_Rev'].total, totalAT = model['AT'].total, totalR = realCum[realCum.length-1]||0;
  const ecart = totalR-totalAT, pct = totalAT > 0 ? (totalR/totalAT*100).toFixed(1) : 0;
  kpiSection.style.display = 'grid';
  kpiSection.innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Budget Révisé Total</div><div class="kpi-value">${toCHF(totalBR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">AT Total</div><div class="kpi-value">${toCHF(totalAT)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Coûts Réels</div><div class="kpi-value">${toCHF(totalR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Écart Réel / AT</div><div class="kpi-value ${ecart<=0?'positive':'negative'}">${toCHF(ecart)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">% Avanc. / AT</div><div class="kpi-value ${pct>100?'negative':'positive'}">${pct}%</div><div class="kpi-unit"></div></div>
  `;

  drawCourbesS(axis, series, realCut, today); buildTablePreview(axis, series['Budget_Rev'], realMonth, initCum);
  drawBarChart(); drawPieCharts(); buildTableCompare();
  statusEl.textContent = '✅ Calcul terminé avec succès'; statusEl.className = 'status success';
}

function drawCourbesS(axis, series, realCut, today) {
  const ds = [
    { label:'Budget (S cumul)',        data:series['Budget'],     borderColor:GE_BLUE4,   backgroundColor:'rgba(91,155,213,0.07)', pointRadius:0, tension:0.3, spanGaps:true },
    { label:'Budget Révisé (S cumul)', data:series['Budget_Rev'], borderColor:GE_ORANGE,  backgroundColor:'rgba(232,135,58,0.07)', pointRadius:0, tension:0.3, spanGaps:true, borderWidth:3 },
    { label:'AT (S cumul)',            data:series['AT'],         borderColor:GE_BLUE2,   backgroundColor:'rgba(30,77,140,0.07)',  pointRadius:0, tension:0.3, spanGaps:true },
    { label:'Réels cumulés',           data:realCut,              borderColor:GE_RED,     backgroundColor:'rgba(192,57,43,0.07)', pointRadius:2, pointHoverRadius:5, tension:0.3, spanGaps:true, borderWidth:3 }
  ];
  const opts = {
    responsive:true, maintainAspectRatio:false,
    plugins:{
      legend:{ position:'top', labels:{ padding:15, font:{size:13,weight:'600'}, usePointStyle:true } },
      tooltip:{ mode:'index', intersect:false, backgroundColor:'rgba(22,58,95,0.95)', callbacks:{ label:ctx=>ctx.parsed.y!==null?`${ctx.dataset.label}: ${toCHF(ctx.parsed.y)} CHF`:'' } },
      todayLine:{ label:yyyymm(today) }
    },
    scales:{
      x:{ title:{display:true,text:'Mois'}, grid:{color:'#eef2f7'}, ticks:{font:{size:11}, callback:(_,i)=>{ const l=axis[i]; return l&&(l.endsWith('-01')||l.endsWith('-04')||l.endsWith('-07')||l.endsWith('-10'))?l:''; }} },
      y:{ title:{display:true,text:'CHF (cumulé)'}, beginAtZero:true, grid:{color:'#eef2f7'}, ticks:{callback:v=>toCHF(v),font:{size:11}} }
    },
    interaction:{mode:'index',intersect:false}
  };
  if (chart) chart.destroy();
  chart = new Chart(document.getElementById('chart').getContext('2d'), { type:'line', data:{labels:axis,datasets:ds}, options:opts, plugins:[todayLinePlugin] });
}

function buildTablePreview(axis, brCum, realMonth, initCum) {
  tablePrev.innerHTML = `<thead><tr><th>Mois</th><th>Montant BR (S)</th><th>Cumul BR (S)</th><th>Réel / mois</th><th>Réel cumulé</th><th>Écart</th></tr></thead><tbody></tbody>`;
  const tbody = tablePrev.querySelector('tbody'); let cumBR = 0, cumR = initCum;
  axis.forEach((m,i) => {
    const brVal = brCum[i], rVal = realMonth[i]||0;
    if (brVal !== null) cumBR = brVal;
    const monthBR = (brVal !== null && i === 0) ? brVal : (brVal !== null && brCum[i-1] !== null) ? brVal-brCum[i-1] : 0;
    cumR += rVal;
    if (brVal === null && rVal === 0 && cumR <= initCum) return;
    const ecart = cumR-cumBR;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${m}</td><td>${toCHF(monthBR)}</td><td>${toCHF(brVal)}</td><td>${toCHF(rVal)}</td><td>${toCHF(cumR)}</td><td class="${ecart<=0?'positive':'negative'}">${toCHF(ecart)}</td>`;
    tbody.appendChild(tr);
  });
}

function buildTableCompare() {
  tableCmp.innerHTML = `<thead><tr><th>Projet</th><th>Budget</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart / Rev</th><th>Écart / AT</th><th>% / AT</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb = tableCmp.querySelector('tbody'), tf = tableCmp.querySelector('tfoot');
  let sB=0,sBR=0,sAT=0,sR=0;
  PROJECT_ORDER.forEach(id => {
    const b  = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='Budget').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const br = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const at = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const rr = realRows.filter(r=>r['ID Projet']===id).reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    sB+=b;sBR+=br;sAT+=at;sR+=rr;
    const eR=rr-br, eA=rr-at, pct=at>0?(rr/at*100).toFixed(1)+'%':'—';
    tb.insertAdjacentHTML('beforeend',`<tr><td><strong>${getPN(id)}</strong></td><td>${toCHF(b)}</td><td>${toCHF(br)}</td><td>${toCHF(at)}</td><td>${toCHF(rr)}</td><td class="${eR<=0?'positive':'negative'}">${toCHF(eR)}</td><td class="${eA<=0?'positive':'negative'}">${toCHF(eA)}</td><td>${pct}</td></tr>`);
  });
  const pT = sAT>0?(sR/sAT*100).toFixed(1)+'%':'—';
  tf.innerHTML = `<tr><td><strong>TOTAL</strong></td><td>${toCHF(sB)}</td><td>${toCHF(sBR)}</td><td>${toCHF(sAT)}</td><td>${toCHF(sR)}</td><td class="${sR-sBR<=0?'positive':'negative'}">${toCHF(sR-sBR)}</td><td class="${sR-sAT<=0?'positive':'negative'}">${toCHF(sR-sAT)}</td><td><strong>${pT}</strong></td></tr>`;
}

function drawPieCharts() {
  const labelsBR=[],dataBR=[],labelsAT=[],dataAT=[];
  PROJECT_ORDER.forEach(id => {
    const br = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const at = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    if(br>0){labelsBR.push(getPN(id));dataBR.push(br);}
    if(at>0){labelsAT.push(getPN(id));dataAT.push(at);}
  });
  const total = arr => arr.reduce((a,b)=>a+b,0);
  const pieCfg = (labels, data) => ({
    type: 'pie',
    data: { labels, datasets: [{ data, backgroundColor: PIE_COLORS, borderWidth: 2, borderColor: '#ffffff' }] },
    options: { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'right', labels: { font:{size:12}, padding:14, usePointStyle:true, boxWidth:12 } },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${toCHF(ctx.parsed)} CHF (${((ctx.parsed/total(data))*100).toFixed(1)}%)` } } } }
  });
  if(chartPieBudgetRev) chartPieBudgetRev.destroy(); if(chartPieAT) chartPieAT.destroy();
  chartPieBudgetRev = new Chart(document.getElementById('chartPieBudgetRev').getContext('2d'), pieCfg(labelsBR, dataBR));
  chartPieAT        = new Chart(document.getElementById('chartPieAT').getContext('2d'),        pieCfg(labelsAT, dataAT));
}

function drawBarChart() {
  const labels=[],dBR=[],dAT=[],dR=[];
  PROJECT_ORDER.forEach(id => {
    labels.push(getPN(id));
    dBR.push(prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
    dAT.push(prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
    dR.push(realRows.filter(r=>r['ID Projet']===id).reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
  });
  if(chartBar) chartBar.destroy();
  chartBar = new Chart(document.getElementById('chartBar').getContext('2d'), {
    type:'bar', data:{ labels, datasets:[
      {label:'Budget Révisé', data:dBR, backgroundColor:GE_ORANGE},
      {label:'AT',             data:dAT, backgroundColor:GE_BLUE2},
      {label:'Réel',           data:dR,  backgroundColor:GE_RED} ]},
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'top'}, tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF`}} },
      scales:{ x:{ticks:{callback:v=>toCHF(v)}} } }
  });
}

// ============================================================
// PARTIE 2 — IMPUTATIONS
// ============================================================
function extractTypeHono(nomTache) {
  if (!nomTache || typeof nomTache !== 'string') return nomTache || 'Autre';
  return nomTache.replace(/_(?:Réseau CAD|Sous-stations|Général|Centrale|Bâtiment|Participations|SIA).*$/, '').trim() || nomTache;
}
function isExternal(r) {
  const tache = (r['Nom de la tâche'] || '').toLowerCase(), fourn = r['Nom du Fournisseur'] || '';
  return tache.includes('externe') || fourn.trim() !== '';
}

function processImputations() {
  const isHono = r => r['Nom de la tâche'] && r['Nom de la tâche'].toLowerCase().includes('honoraires');
  const honoBR = prevRows.filter(r=>isHono(r)&&r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
  const honoAT = prevRows.filter(r=>isHono(r)&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
  const honoR  = realRows.filter(isHono).reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
  const reste  = honoAT - honoR;
  kpiHono.style.display = 'grid';
  kpiHono.innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Honoraires BR</div><div class="kpi-value">${toCHF(honoBR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Honoraires AT</div><div class="kpi-value">${toCHF(honoAT)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Honoraires Réels</div><div class="kpi-value">${toCHF(honoR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Reste à imputer (AT)</div><div class="kpi-value ${reste>=0?'positive':'negative'}">${toCHF(reste)}</div><div class="kpi-unit">CHF</div></div>
  `;
  const typesBR={}, typesAT={}, typesR={};
  prevRows.forEach(r => {
    if(!isHono(r)) return;
    const t=extractTypeHono(r['Nom de la tâche']), v=toNum(r['Montant total du coût']);
    if(r['Modèle de prévision']==='Budget_Rev') typesBR[t]=(typesBR[t]||0)+v;
    if(r['Modèle de prévision']==='AT')         typesAT[t]=(typesAT[t]||0)+v;
  });
  realRows.forEach(r => { if(!isHono(r)) return; const t=extractTypeHono(r['Nom de la tâche']); typesR[t]=(typesR[t]||0)+toNum(r['Montant total du coût']); });
  const allTypes = [...new Set([...Object.keys(typesBR),...Object.keys(typesAT),...Object.keys(typesR)])];
  if(chartHonoMetier) chartHonoMetier.destroy();
  chartHonoMetier = new Chart(document.getElementById('chartHonoMetier').getContext('2d'), {
    type:'bar', data:{ labels:allTypes, datasets:[
      {label:'Budget Révisé', data:allTypes.map(t=>typesBR[t]||0), backgroundColor:GE_ORANGE},
      {label:'AT',             data:allTypes.map(t=>typesAT[t]||0), backgroundColor:GE_BLUE2},
      {label:'Réel',           data:allTypes.map(t=>typesR[t]||0),  backgroundColor:GE_RED} ]},
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'top'}, tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF`}}},
      scales:{x:{ticks:{callback:v=>toCHF(v)}}} }
  });
  const buildMap = rows => {
    const m = {};
    rows.forEach(r => {
      if(!isHono(r)) return;
      const ress = r['Ressource.Nom de la ressource'] || (r['Nom du Fournisseur']?`Fourn.: ${r['Nom du Fournisseur']}`:'Non spécifié');
      const type = extractTypeHono(r['Nom de la tâche']||''), proj = getPN(r['ID Projet']||''), ext = isExternal(r);
      const key  = `${ext?'EXT':'INT'}||${ress}||${type}||${proj}`;
      if(!m[key]) m[key]={ress,type,proj,ext,total:0};
      m[key].total += toNum(r['Montant total du coût']);
    });
    return Object.values(m).filter(r=>r.total!==0).sort((a,b)=>b.total-a.total);
  };
  const allRows = buildMap(realRows), extRows = allRows.filter(r=>r.ext), intRows = allRows.filter(r=>!r.ext);
  const totExt = extRows.reduce((a,r)=>a+r.total,0), totInt = intRows.reduce((a,r)=>a+r.total,0), totAll = totExt+totInt;
  const makeRow = (r,cls='') => `<tr class="${cls}"><td>${r.ress}</td><td>${r.type}</td><td>${r.proj}</td><td class="${r.total<0?'negative':''}">${toCHF(r.total)}</td></tr>`;
  const sepRow  = (label,val,cls) => `<tr class="${cls}"><td colspan="3"><strong>${label}</strong></td><td><strong>${toCHF(val)}</strong></td></tr>`;
  tableRess.innerHTML = `<thead><tr><th>Ressource</th><th>Type d'Honoraires</th><th>Projet</th><th>Total Réel (CHF)</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb=tableRess.querySelector('tbody'), tf=tableRess.querySelector('tfoot');
  tb.insertAdjacentHTML('beforeend',`<tr class="section-header-int"><td colspan="4">🔵 Honoraires Internes</td></tr>`);
  intRows.forEach(r=>tb.insertAdjacentHTML('beforeend',makeRow(r,'row-int')));
  tb.insertAdjacentHTML('beforeend',sepRow('Sous-total Internes',totInt,'subtotal-int'));
  tb.insertAdjacentHTML('beforeend',`<tr class="section-header-ext"><td colspan="4">🟠 Honoraires Externes</td></tr>`);
  extRows.forEach(r=>tb.insertAdjacentHTML('beforeend',makeRow(r,'row-ext')));
  tb.insertAdjacentHTML('beforeend',sepRow('Sous-total Externes',totExt,'subtotal-ext'));
  tf.innerHTML = `<tr><td colspan="3"><strong>TOTAL GÉNÉRAL</strong></td><td><strong>${toCHF(totAll)}</strong></td></tr>`;
}

// ============================================================
// PARTIE 3 — PLANNING
// ============================================================
function processPlanning() {
  const plan = {};
  PROJECT_ORDER.forEach(id => {
    const getR = mode => {
      const rows  = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']===mode);
      const dates = rows.flatMap(r=>[toDate(r['Date de début']),toDate(r['Date de fin'])]).filter(Boolean);
      return dates.length ? {s:new Date(Math.min(...dates)),e:new Date(Math.max(...dates))} : {s:null,e:null};
    };
    const br=getR('Budget_Rev'), at=getR('AT'), man=manualPlanningDates[id]||{};
    let rS=man.start||null, rE=man.end||null;
    if(!rS||!rE){
      const rd=realRows.filter(r=>r['ID Projet']===id).map(r=>toDate(r['Date du projet'])).filter(Boolean);
      if(rd.length){if(!rS)rS=new Date(Math.min(...rd));if(!rE)rE=new Date(Math.max(...rd));}
    }
    plan[id]={name:getPN(id),brS:br.s,brE:br.e,atS:at.s,atE:at.e,rS,rE};
  });
  buildTablePlanning(plan); drawGantt(plan);
}

function buildTablePlanning(plan) {
  tablePlan.innerHTML=`<thead><tr><th>Projet</th><th>BR Début</th><th>BR Fin</th><th>AT Début</th><th>AT Fin</th><th>Réel Début</th><th>Réel Fin</th><th>Écart (j)</th><th>Statut</th></tr></thead><tbody></tbody>`;
  const fmt=d=>d&&!isNaN(d)?d.toISOString().split('T')[0]:'';
  const disp=d=>d&&!isNaN(d)?d.toLocaleDateString('fr-CH'):'—';
  const tb=tablePlan.querySelector('tbody');
  PROJECT_ORDER.forEach(id=>{
    const p=plan[id], ecart=p.atE&&p.rE?Math.round((p.rE-p.atE)/86400000):null;
    let stat='prevu',sLab='Prévu';
    if(p.rS&&!p.rE){stat='en-cours';sLab='En cours';} else if(p.rE){stat=ecart>0?'retard':'termine';sLab=ecart>0?'Retard':'Terminé';}
    tb.insertAdjacentHTML('beforeend',`<tr><td><strong>${p.name}</strong></td><td>${disp(p.brS)}</td><td>${disp(p.brE)}</td><td>${disp(p.atS)}</td><td>${disp(p.atE)}</td><td><input type="date" id="rs_${id}" class="date-input" value="${fmt(p.rS)}"/></td><td><input type="date" id="re_${id}" class="date-input" value="${fmt(p.rE)}"/></td><td>${ecart!==null?ecart:'—'}</td><td><span class="badge ${stat}">${sLab}</span></td></tr>`);
  });
}

// ============================================================
// GANTT  v4.4
// Approche : 3 datasets groupés (BR / AT / Réel), 1 point par projet.
// barThickness en pixels = épaisseur absolue garantie + centrage natif Chart.js.
// ============================================================
function drawGantt(plan) {
  const allD = [];
  PROJECT_ORDER.forEach(id => {
    const p = plan[id];
    [p.brS, p.brE, p.atS, p.atE, p.rS, p.rE].forEach(d => { if (d && !isNaN(d)) allD.push(d); });
  });
  if (!allD.length) return;

  const minD = new Date(Math.min(...allD));
  const maxD = new Date(Math.max(...allD));
  const span = Math.round((maxD - minD) / 86400000) + 60;
  const toDay = d => (d && !isNaN(d)) ? Math.round((d - minD) / 86400000) : null;

  // Labels Y : un label par projet (les 3 barres s'empilent en groupement)
  const yLabels = PROJECT_ORDER.map(id => getPN(id));

  // Construction des 3 datasets
  // Chaque dataset a un point par projet : { x: [start, end], y: nomProjet }
  // Si une date manque, on met null pour que Chart.js saute ce point
  const mkPoint = (s, e, label) => {
    const s0 = toDay(s), e0 = toDay(e);
    if (s0 === null || e0 === null) return { x: [null, null], y: label };
    return { x: [s0, e0], y: label };
  };

  const dsBR = {
    label: 'Budget Révisé',
    data: PROJECT_ORDER.map(id => mkPoint(plan[id].brS, plan[id].brE, getPN(id))),
    backgroundColor: GE_ORANGE + 'cc',
    borderColor: GE_ORANGE,
    borderWidth: 2,
    borderSkipped: false,
    borderRadius: 3,
    barThickness: 18          // épaisseur fixe en pixels
  };
  const dsAT = {
    label: 'AT',
    data: PROJECT_ORDER.map(id => mkPoint(plan[id].atS, plan[id].atE, getPN(id))),
    backgroundColor: GE_BLUE2 + 'cc',
    borderColor: GE_BLUE2,
    borderWidth: 2,
    borderSkipped: false,
    borderRadius: 3,
    barThickness: 18
  };
  const dsReal = {
    label: 'Réel',
    data: PROJECT_ORDER.map(id => mkPoint(plan[id].rS, plan[id].rE, getPN(id))),
    backgroundColor: GE_RED + 'cc',
    borderColor: GE_RED,
    borderWidth: 2,
    borderSkipped: false,
    borderRadius: 3,
    barThickness: 18
  };

  const tickStep = Math.max(30, Math.round(span / 12));
  if (chartGantt) chartGantt.destroy();

  chartGantt = new Chart(document.getElementById('chartGantt').getContext('2d'), {
    type: 'bar',
    data: { labels: yLabels, datasets: [dsBR, dsAT, dsReal] },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          position: 'top',
          labels: { usePointStyle: true, padding: 16, font: { size: 12, weight: '600' } }
        },
        tooltip: {
          callbacks: {
            title: ctx => `${ctx[0].label} — ${ctx[0].dataset.label}`,
            label: ctx => {
              const [d0, d1] = ctx.parsed.x;
              if (d0 === null || d1 === null) return 'Pas de données';
              const sD = new Date(minD.getTime() + d0 * 86400000);
              const eD = new Date(minD.getTime() + d1 * 86400000);
              const dur = Math.round((eD - sD) / 86400000);
              return [
                `Début : ${sD.toLocaleDateString('fr-CH')}`,
                `Fin   : ${eD.toLocaleDateString('fr-CH')}`,
                `Durée : ${dur} jours`
              ];
            }
          }
        }
      },
      scales: {
        x: {
          type: 'linear',
          min: 0,
          max: span,
          title: { display: true, text: 'Timeline' },
          grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: {
            stepSize: tickStep,
            callback: value => {
              const d = new Date(minD.getTime() + value * 86400000);
              return d.toLocaleDateString('fr-CH', { month: 'short', year: '2-digit' });
            }
          }
        },
        y: {
          type: 'category',
          labels: yLabels,
          offset: true,
          grid: {
            color: 'rgba(0,0,0,0.08)',
            lineWidth: 1
          },
          ticks: {
            font: { weight: 'bold', size: 11 },
            color: '#163a5f'
          }
        }
      }
    }
  });
}

// ============================================================
// EXPORTS
// ============================================================
function exportChartPNG(canvasId, filename) {
  const canvas = document.getElementById(canvasId); if(!canvas) return;
  const a = document.createElement('a'); a.download = filename+'.png'; a.href = canvas.toDataURL('image/png',1); a.click();
}
window.exportChartPNG = exportChartPNG;

async function exportTablePNG(cardId, filename) {
  const el = document.getElementById(cardId); if(!el) return;
  const canvas = await html2canvas(el,{scale:2,backgroundColor:'#ffffff',logging:false});
  canvas.toBlob(blob=>{ const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.download=filename+'.png'; a.href=url; a.click(); URL.revokeObjectURL(url); });
}
window.exportTablePNG = exportTablePNG;

function exportTableXLSX(tableId, sheetName) {
  const el = document.getElementById(tableId); if(!el) return;
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(el), sheetName.substring(0,31)); XLSX.writeFile(wb, sheetName+'.xlsx');
}
window.exportTableXLSX = exportTableXLSX;

// ============================================================
// PDF  v4.4
// ============================================================
async function generatePDF() {
  const { jsPDF } = window.jspdf;
  const today = new Date(), dd = String(today.getDate()).padStart(2,'0'), mm = String(today.getMonth()+1).padStart(2,'0'), yyyy = today.getFullYear();
  const fname = `${yyyy}${mm}${dd}_Etat_de_Projet.pdf`, dateStr = today.toLocaleDateString('fr-CH',{day:'2-digit',month:'long',year:'numeric'});
  const pdf = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
  const W = 297, H = 210, PAD = 10, HDR = 18, FTR = 8;
  const CONTENT_H = H - HDR - FTR - PAD, CONTENT_W = W - 2*PAD;

  function drawHeader(title, pageLabel) {
    pdf.setFillColor(22,58,95); pdf.rect(0,0,W,HDR,'F'); pdf.setFillColor(232,135,58); pdf.rect(0,HDR-1.5,W,1.5,'F');
    pdf.setTextColor(255,255,255); pdf.setFontSize(10); pdf.setFont('helvetica','bold'); pdf.text(title, PAD, 11.5);
    pdf.setFontSize(8); pdf.setFont('helvetica','normal'); pdf.text(`${fname.replace('.pdf','')}   |   ${pageLabel}`, W-PAD, 11.5, {align:'right'});
  }
  function drawFooter(page, total) {
    pdf.setFillColor(240,244,249); pdf.rect(0, H-FTR, W, FTR, 'F'); pdf.setTextColor(120,140,160); pdf.setFontSize(7.5); pdf.setFont('helvetica','normal');
    pdf.text('Groupe E Celsius — Dashboard Projets CAD — Confidentiel', PAD, H-2); pdf.text(`Page ${page} / ${total}`, W-PAD, H-2, {align:'right'});
  }

  pdf.setFillColor(22,58,95); pdf.rect(0,0,W,H,'F'); pdf.setFillColor(232,135,58); pdf.rect(0,0,8,H,'F');
  pdf.setFillColor(255,255,255); pdf.roundedRect(30, 40, W-60, 130, 6, 6, 'F');
  pdf.setTextColor(22,58,95); pdf.setFontSize(26); pdf.setFont('helvetica','bold'); pdf.text('Dashboard Projets CAD', W/2, 82, {align:'center'});
  pdf.setFontSize(15); pdf.setFont('helvetica','normal'); pdf.text('Groupe E Celsius', W/2, 97, {align:'center'});
  pdf.setDrawColor(232,135,58); pdf.setLineWidth(0.8); pdf.line(W/2-40, 103, W/2+40, 103);
  pdf.setFontSize(12); pdf.setTextColor(22,58,95); pdf.text('État de Projet — SIA 5 — Réalisation', W/2, 112, {align:'center'});
  pdf.setFontSize(10); pdf.setTextColor(90,122,154); pdf.text(`Généré le ${dateStr}`, W/2, 124, {align:'center'});
  pdf.setFontSize(9); pdf.setTextColor(180,190,200); pdf.text(fname, W/2, 160, {align:'center'});

  const sections = [
    { id:'kpiSection', title:'Indicateurs Clés de Performance', type:'kpi' },
    { id:'cardCourbesS', title:'Courbes S — Coûts Cumulés', type:'full' },
    { id:'tablePrevCard', title:'Budget Révisé — Détails Mensuels', type:'full' },
    { ids:['cardPieBR','cardPieAT'], titles:['Répartition Budget Révisé','Répartition AT'], type:'double' },
    { id:'cardBarProjet', title:'Avancement par Projet (BR/AT/Réel)', type:'full' },
    { id:'tableCmpCard', title:'Comparatif par Projet', type:'full' },
    { id:'kpiSectionHonoraires', title:'KPI Honoraires', type:'kpi' },
    { id:'cardHonoMetier', title:'Honoraires par Type de Métier', type:'full' },
    { id:'tableRessCard', title:'Coûts Réels par Ressource — Honoraires', type:'full' },
    { id:'cardGantt', title:'Planning Gantt', type:'full' },
    { id:'tablePlanCard', title:'Tableau Consolidé Planning', type:'full' }
  ];

  async function captureEl(id) {
    const el = document.getElementById(id); if(!el || el.style.display==='none') return null;
    try { return await html2canvas(el, { scale:2, backgroundColor:'#ffffff', logging:false, useCORS:true }); }
    catch(e) { console.warn('capture error', id, e); return null; }
  }
  const captured = {};
  for(const sec of sections) {
    if(sec.type==='double') { for(const id of sec.ids) captured[id] = await captureEl(id); }
    else { captured[sec.id] = await captureEl(sec.id); }
  }
  let totalPages = 1;
  for(const sec of sections) {
    if(sec.type==='double') { totalPages++; continue; }
    const c = captured[sec.id]; if(!c) continue; const imgH = CONTENT_W * (c.height/c.width); totalPages += Math.ceil(imgH/CONTENT_H);
  }
  let pageNum = 1; drawFooter(pageNum, totalPages);

  for(const sec of sections) {
    if(sec.type === 'double') {
      const canvases = sec.ids.map(id=>captured[id]).filter(Boolean); if(!canvases.length) continue; pageNum++; pdf.addPage();
      drawHeader(sec.titles.join('  /  '), `${pageNum}/${totalPages}`); drawFooter(pageNum, totalPages);
      const colW = (CONTENT_W-8)/2, yTop = HDR+PAD;
      canvases.forEach((c,ci) => {
        const xOff = PAD + ci*(colW+8), ratio = c.height/c.width, iH = Math.min(colW*ratio, CONTENT_H-6);
        pdf.setFillColor(255,255,255); pdf.roundedRect(xOff-2, yTop-2, colW+4, iH+4+10, 3, 3, 'F');
        pdf.setDrawColor(220,228,240); pdf.setLineWidth(0.4); pdf.roundedRect(xOff-2, yTop-2, colW+4, iH+4+10, 3, 3, 'S');
        pdf.setFontSize(9); pdf.setFont('helvetica','bold'); pdf.setTextColor(22,58,95); pdf.text(sec.titles[ci]||'', xOff+colW/2, yTop+7, {align:'center'});
        const imgData = c.toDataURL('image/jpeg',0.92); pdf.addImage(imgData,'JPEG', xOff, yTop+10, colW, iH);
      });
      continue;
    }
    if(sec.type === 'kpi') {
      const c = captured[sec.id]; if(!c) continue; const imgH_mm = CONTENT_W * (c.height/c.width); pageNum++; pdf.addPage();
      drawHeader(sec.title, `${pageNum}/${totalPages}`); drawFooter(pageNum, totalPages); const iH = Math.min(imgH_mm, CONTENT_H);
      const imgData = c.toDataURL('image/jpeg',0.93); pdf.addImage(imgData,'JPEG', PAD, HDR+4, CONTENT_W, iH); continue;
    }
    const c = captured[sec.id]; if(!c) continue; const srcW = c.width, srcH = c.height;
    const imgH_mm = CONTENT_W * (srcH/srcW), pagesNeeded = Math.ceil(imgH_mm / CONTENT_H);
    for(let pg=0; pg<pagesNeeded; pg++) {
      pageNum++; pdf.addPage(); const titleSuffix = pagesNeeded>1 ? ` (${pg+1}/${pagesNeeded})` : '';
      drawHeader(sec.title+titleSuffix, `${pageNum}/${totalPages}`); drawFooter(pageNum, totalPages);
      const slicePxH = Math.ceil(srcH/pagesNeeded), srcYStart = pg*slicePxH, srcYEnd = Math.min(srcYStart+slicePxH, srcH);
      const sl = document.createElement('canvas'); sl.width = srcW; sl.height = srcYEnd-srcYStart;
      sl.getContext('2d').drawImage(c, 0, -srcYStart); const slData = sl.toDataURL('image/jpeg',0.92);
      const slH_mm = CONTENT_W*(sl.height/sl.width); pdf.addImage(slData,'JPEG', PAD, HDR+4, CONTENT_W, Math.min(slH_mm, CONTENT_H));
    }
  }
  pdf.save(fname);
}
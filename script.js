/* ============================================================
   Dashboard CAD — Groupe E Celsius
   script.js  v3.0  — 18/02/2026
   Fix: courbe S range indépendant par type, exports individuels,
        PDF rapport, Gantt légendé, honoraires filtrés
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

// ---- Palette Groupe E bleu ----
const GE_BLUE       = '#163a5f';
const GE_BLUE2      = '#1e4d8c';
const GE_BLUE3      = '#2e6db4';
const GE_BLUE4      = '#5b9bd5';
const GE_ORANGE     = '#f59e0b';
const GE_RED        = '#ef4444';
const GE_GREEN      = '#10b981';
const GE_PURPLE     = '#6366f1';
const PIE_COLORS    = [GE_BLUE, GE_BLUE2, GE_BLUE3, GE_BLUE4, GE_ORANGE, GE_PURPLE];

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

function toCHF(n) {
  return (n || 0).toLocaleString('fr-CH', { minimumFractionDigits: 0 });
}

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
  const axis = [];
  const c = firstOfMonth(minD), e = firstOfMonth(maxD);
  while (c <= e) { axis.push(yyyymm(c)); c.setMonth(c.getMonth()+1); }
  return axis;
}

/**
 * Courbe S cumulative.
 * range = [sIdx, eIdx] — indices dans axis correspondant aux dates propres du type.
 * En dehors du range les valeurs restent null (spanGaps=true les relie graphiquement).
 */
function sCurve(total, axis, k, sIdx, eIdx) {
  const N   = axis.length;
  const cum = Array(N).fill(null);
  const L   = eIdx - sIdx + 1;
  if (L <= 0 || total === 0 || sIdx < 0 || eIdx < 0) return cum;
  const s0 = 1 / (1 + Math.exp(-k * (0 - 0.5)));
  const s1 = 1 / (1 + Math.exp(-k * (1 - 0.5)));
  for (let j = 0; j < L; j++) {
    const i = sIdx + j;
    if (i >= N) break;
    const x  = j / (L - 1 || 1);
    const sx = 1 / (1 + Math.exp(-k * (x - 0.5)));
    cum[i]   = ((sx - s0) / (s1 - s0)) * total;
  }
  return cum;
}

/**
 * Calcule le range (sIdx, eIdx) d'un modèle sur l'axe fourni.
 * Utilise les colonnes "Date de début" et "Date de fin" du fichier Prévisions.
 */
function modelRange(rows, modelName, axis) {
  const filtered = rows.filter(r => r['Modèle de prévision'] === modelName);
  const dates = filtered.flatMap(r => [
    toDate(r['Date de début']), toDate(r['Date de fin'])
  ]).filter(Boolean);
  if (!dates.length) return [-1, -1];
  const minD = new Date(Math.min(...dates));
  const maxD = new Date(Math.max(...dates));
  const mmMin = yyyymm(firstOfMonth(minD));
  const mmMax = yyyymm(firstOfMonth(maxD));
  let sIdx = axis.indexOf(mmMin);
  let eIdx = axis.indexOf(mmMax);
  // si non trouvé sur l'axe, chercher le plus proche
  if (sIdx < 0) sIdx = axis.findIndex(m => m >= mmMin);
  if (eIdx < 0) {
    for (let i = axis.length - 1; i >= 0; i--) { if (axis[i] <= mmMax) { eIdx = i; break; } }
  }
  if (sIdx < 0) sIdx = 0;
  if (eIdx < 0) eIdx = axis.length - 1;
  return [sIdx, eIdx];
}

// Plugin ligne "Aujourd'hui"
const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(ch, _args, opts) {
    if (!opts || !opts.label) return;
    const { ctx, chartArea: { top, bottom }, scales: { x } } = ch;
    const labels = ch.data.labels;
    const toT    = s => { const [y,m] = s.split('-'); return new Date(+y,+m-1,1).getTime(); };
    const t      = toT(opts.label);
    let bi = 0, bd = Infinity;
    labels.forEach((l, i) => { const d = Math.abs(toT(l) - t); if (d < bd) { bd = d; bi = i; } });
    const xp = x.getPixelForValue(bi);
    ctx.save();
    ctx.strokeStyle = opts.color || GE_RED;
    ctx.lineWidth   = opts.width || 2;
    ctx.setLineDash([5, 5]);
    ctx.beginPath(); ctx.moveTo(xp, top); ctx.lineTo(xp, bottom); ctx.stroke();
    ctx.setLineDash([]);
    if (opts.text) {
      ctx.fillStyle = opts.color || GE_RED;
      ctx.font      = 'bold 11px Arial';
      ctx.textAlign = 'center';
      ctx.fillText(opts.text, xp, top - 5);
    }
    ctx.restore();
  }
};

// ============================================================
// EVENTS
// ============================================================
prevInput.addEventListener('change', e => {
  if (e.target.files[0]) readExcel(e.target.files[0], rows => { prevRows = rows; maybeProcess(); });
});
realInput.addEventListener('change', e => {
  if (e.target.files[0]) readExcel(e.target.files[0], rows => { realRows = rows; maybeProcess(); });
});
[kSlider, kNumber].forEach(el => {
  el.addEventListener('input', () => {
    kSlider.value = kNumber.value = el.value;
    if (prevRows.length && realRows.length) processAll();
  });
});
[viewGlobalEl, viewSOnlyEl].forEach(el => {
  el.addEventListener('change', () => { if (prevRows.length && realRows.length) processAll(); });
});
document.getElementById('btnApplyPlanning').addEventListener('click', () => {
  if (!prevRows.length || !realRows.length) return;
  PROJECT_ORDER.forEach(id => {
    const s = document.getElementById(`rs_${id}`);
    const e = document.getElementById(`re_${id}`);
    if (!manualPlanningDates[id]) manualPlanningDates[id] = {};
    manualPlanningDates[id].start = s && s.value ? new Date(s.value) : null;
    manualPlanningDates[id].end   = e && e.value ? new Date(e.value) : null;
  });
  processPlanning();
});
document.getElementById('btnPDF').addEventListener('click', generatePDF);

function maybeProcess() {
  if (prevRows.length && realRows.length) {
    statusEl.textContent = '\u2705 Fichiers chargés — calcul en cours...';
    statusEl.className   = 'status';
    processAll();
  }
}

function processAll() {
  processBudgets();
  processImputations();
  processPlanning();
}

// ============================================================
// PARTIE 1 — BUDGETS
// ============================================================
function processBudgets() {
  const k        = parseFloat(kNumber.value) || 12;
  const MODELS   = ['Budget', 'Budget_Rev', 'AT'];
  const isSOnly  = viewSOnlyEl.checked;

  // Totaux + dates min/max par modèle
  const model = {};
  MODELS.forEach(mn => {
    const rows  = prevRows.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = {
      total,
      start: dates.length ? new Date(Math.min(...dates)) : null,
      end:   dates.length ? new Date(Math.max(...dates)) : null
    };
  });

  // Réels
  const realDates = realRows.map(r => toDate(r['Date du projet'])).filter(Boolean);

  // Axe global
  const allDates = [];
  MODELS.forEach(mn => { if (model[mn].start) allDates.push(model[mn].start); if (model[mn].end) allDates.push(model[mn].end); });
  realDates.forEach(d => allDates.push(d));
  if (!allDates.length) { statusEl.textContent = '\u26a0\ufe0f Aucune date trouvée.'; return; }

  const globalMin = new Date(Math.min(...allDates));
  const globalMax = new Date(Math.max(...allDates));

  // En vue S-only: axe limité à la plage des modèles
  const modelMin = new Date(Math.min(...MODELS.filter(m => model[m].start).map(m => model[m].start)));
  const modelMax = new Date(Math.max(...MODELS.filter(m => model[m].end).map(m => model[m].end)));
  const axisMin  = isSOnly ? modelMin : globalMin;
  const axisMax  = isSOnly ? modelMax : globalMax;
  const axis     = buildAxis(axisMin, axisMax);

  // ---- CORRECTION BUG COURBE S ----
  // Chaque modèle a son propre range calculé depuis SES dates dans prevRows
  const series = {};
  MODELS.forEach(mn => {
    const [si, ei] = modelRange(prevRows, mn, axis);
    series[mn] = sCurve(model[mn].total, axis, k, si, ei);
  });

  // Réels par mois
  const today    = new Date();
  const mmToday  = yyyymm(today);
  const realMap  = new Map(axis.map(m => [m, 0]));
  let   initCum  = 0; // cumul avant le début de l'axe (vue S-only)

  realRows.forEach(r => {
    const d = toDate(r['Date du projet']);
    if (!d) return;
    const mm = yyyymm(firstOfMonth(d));
    if (mm > mmToday) return;
    if (realMap.has(mm)) realMap.set(mm, realMap.get(mm) + toNum(r['Montant total du coût']));
    else if (isSOnly && mm < axis[0]) initCum += toNum(r['Montant total du coût']);
  });

  const realMonth = axis.map(m => realMap.get(m) || 0);
  let acc = initCum;
  const realCum = realMonth.map(v => { acc += v; return acc; });
  const realCut = realCum.map((v, i) => axis[i] > mmToday ? null : v);

  // KPI
  const totalBR  = model['Budget_Rev'].total;
  const totalAT  = model['AT'].total;
  const totalR   = realCum[realCum.length - 1] || 0;
  const ecart    = totalR - totalAT;
  const pct      = totalAT > 0 ? (totalR / totalAT * 100).toFixed(1) : 0;
  kpiSection.style.display = 'grid';
  kpiSection.innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Budget Révisé Total</div><div class="kpi-value">${toCHF(totalBR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">AT Total</div><div class="kpi-value">${toCHF(totalAT)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Coûts Réels</div><div class="kpi-value">${toCHF(totalR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Écart Réel / AT</div><div class="kpi-value ${ecart <= 0 ? 'positive' : 'negative'}">${toCHF(ecart)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">% Avanc. / AT</div><div class="kpi-value ${pct > 100 ? 'negative' : 'positive'}">${pct}%</div><div class="kpi-unit"></div></div>
  `;

  drawCourbesS(axis, series, realCut, today);
  buildTablePreview(axis, series['Budget_Rev'], realMonth, initCum);
  drawBarChart();
  drawPieCharts();
  buildTableCompare();

  statusEl.textContent = '\u2705 Calcul terminé avec succès';
  statusEl.className   = 'status success';
}

function drawCourbesS(axis, series, realCut, today) {
  const ds = [
    { label: 'Budget (S cumul)',         data: series['Budget'],     borderColor: GE_BLUE4,  backgroundColor: 'rgba(91,155,213,0.08)', pointRadius: 0, tension: 0.3, spanGaps: true },
    { label: 'Budget Révisé (S cumul)',  data: series['Budget_Rev'], borderColor: GE_ORANGE, backgroundColor: 'rgba(245,158,11,0.08)', pointRadius: 0, tension: 0.3, spanGaps: true, borderWidth: 3 },
    { label: 'AT (S cumul)',             data: series['AT'],         borderColor: GE_BLUE2,  backgroundColor: 'rgba(30,77,140,0.08)',  pointRadius: 0, tension: 0.3, spanGaps: true },
    { label: 'Réels cumulés',            data: realCut,              borderColor: GE_RED,    backgroundColor: 'rgba(239,68,68,0.08)', pointRadius: 2, pointHoverRadius: 5, tension: 0.3, spanGaps: true, borderWidth: 3 }
  ];
  const opts = {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top', labels: { padding: 15, font: { size: 13, weight: '600' }, usePointStyle: true } },
      tooltip: { mode: 'index', intersect: false, backgroundColor: 'rgba(22,58,95,0.95)', callbacks: { label: ctx => ctx.parsed.y !== null ? `${ctx.dataset.label}: ${toCHF(ctx.parsed.y)} CHF` : '' } },
      todayLine: { label: yyyymm(today), text: "Aujourd'hui", color: GE_RED, width: 2 }
    },
    scales: {
      x: { title: { display: true, text: 'Mois' }, grid: { color: '#eef2f7' }, ticks: { font: { size: 11 }, callback: (_, i) => { const l = axis[i]; return l && (l.endsWith('-01')||l.endsWith('-04')||l.endsWith('-07')||l.endsWith('-10')) ? l : ''; } } },
      y: { title: { display: true, text: 'CHF (cumulé)' }, beginAtZero: true, grid: { color: '#eef2f7' }, ticks: { callback: v => toCHF(v), font: { size: 11 } } }
    },
    interaction: { mode: 'index', intersect: false }
  };
  if (chart) chart.destroy();
  chart = new Chart(document.getElementById('chart').getContext('2d'), { type: 'line', data: { labels: axis, datasets: ds }, options: opts, plugins: [todayLinePlugin] });
}

function buildTablePreview(axis, brCum, realMonth, initCum) {
  tablePrev.innerHTML = `<thead><tr><th>Mois</th><th>Montant BR (S)</th><th>Cumul BR (S)</th><th>Réel / mois</th><th>Réel cumulé</th><th>Écart</th></tr></thead><tbody></tbody>`;
  const tbody = tablePrev.querySelector('tbody');
  let cumBR = 0, cumR = initCum;
  let prevBR = null;
  axis.forEach((m, i) => {
    const brVal = brCum[i];
    const rVal  = realMonth[i] || 0;
    if (brVal !== null) cumBR = brVal;
    const monthBR = (brVal !== null && i === 0) ? brVal
                  : (brVal !== null && brCum[i-1] !== null) ? brVal - brCum[i-1]
                  : 0;
    cumR += rVal;
    if (brVal === null && rVal === 0 && cumR <= initCum) return;
    const ecart = cumR - cumBR;
    const tr    = document.createElement('tr');
    tr.innerHTML = `<td>${m}</td><td>${toCHF(monthBR)}</td><td>${toCHF(brVal)}</td><td>${toCHF(rVal)}</td><td>${toCHF(cumR)}</td><td class="${ecart<=0?'positive':'negative'}">${toCHF(ecart)}</td>`;
    tbody.appendChild(tr);
  });
}

function buildTableCompare() {
  tableCmp.innerHTML = `<thead><tr><th>Projet</th><th>Budget</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart / Rev</th><th>Écart / AT</th><th>% / AT</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb = tableCmp.querySelector('tbody');
  const tf = tableCmp.querySelector('tfoot');
  let sB = 0, sBR = 0, sAT = 0, sR = 0;
  PROJECT_ORDER.forEach(id => {
    const b   = prevRows.filter(r => r['ID projet']===id && r['Modèle de prévision']==='Budget').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const br  = prevRows.filter(r => r['ID projet']===id && r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const at  = prevRows.filter(r => r['ID projet']===id && r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const rr  = realRows.filter(r => r['ID Projet']===id).reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    sB+=b; sBR+=br; sAT+=at; sR+=rr;
    const eR  = rr - br, eA = rr - at;
    const pct = at > 0 ? (rr/at*100).toFixed(1)+'%' : '—';
    tb.insertAdjacentHTML('beforeend', `<tr><td><strong>${getPN(id)}</strong></td><td>${toCHF(b)}</td><td>${toCHF(br)}</td><td>${toCHF(at)}</td><td>${toCHF(rr)}</td><td class="${eR<=0?'positive':'negative'}">${toCHF(eR)}</td><td class="${eA<=0?'positive':'negative'}">${toCHF(eA)}</td><td>${pct}</td></tr>`);
  });
  const pT = sAT>0 ? (sR/sAT*100).toFixed(1)+'%' : '—';
  tf.innerHTML = `<tr><td><strong>TOTAL</strong></td><td>${toCHF(sB)}</td><td>${toCHF(sBR)}</td><td>${toCHF(sAT)}</td><td>${toCHF(sR)}</td><td class="${sR-sBR<=0?'positive':'negative'}">${toCHF(sR-sBR)}</td><td class="${sR-sAT<=0?'positive':'negative'}">${toCHF(sR-sAT)}</td><td><strong>${pT}</strong></td></tr>`;
}

function drawPieCharts() {
  const labelsBR=[], dataBR=[], labelsAT=[], dataAT=[];
  PROJECT_ORDER.forEach((id,i) => {
    const br = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    const at = prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
    if (br>0){labelsBR.push(getPN(id));dataBR.push(br);}
    if (at>0){labelsAT.push(getPN(id));dataAT.push(at);}
  });
  const pieCfg = (labels, data) => ({ type:'pie', data:{ labels, datasets:[{data,backgroundColor:PIE_COLORS}] }, options:{ responsive:true,maintainAspectRatio:false,plugins:{ legend:{position:'right'}, tooltip:{callbacks:{label:ctx=>`${ctx.label}: ${toCHF(ctx.parsed)} CHF`}} } } });
  if (chartPieBudgetRev) chartPieBudgetRev.destroy();
  if (chartPieAT)        chartPieAT.destroy();
  chartPieBudgetRev = new Chart(document.getElementById('chartPieBudgetRev').getContext('2d'), pieCfg(labelsBR,dataBR));
  chartPieAT        = new Chart(document.getElementById('chartPieAT').getContext('2d'),        pieCfg(labelsAT,dataAT));
}

function drawBarChart() {
  const labels=[], dBR=[], dAT=[], dR=[];
  PROJECT_ORDER.forEach(id => {
    labels.push(getPN(id));
    dBR.push(prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
    dAT.push(prevRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
    dR.push(realRows.filter(r=>r['ID Projet']===id).reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
  });
  if (chartBar) chartBar.destroy();
  chartBar = new Chart(document.getElementById('chartBar').getContext('2d'), {
    type:'bar',
    data:{ labels, datasets:[
      { label:'Budget Révisé', data:dBR, backgroundColor:GE_ORANGE },
      { label:'AT',            data:dAT, backgroundColor:GE_BLUE2 },
      { label:'Réel',          data:dR,  backgroundColor:GE_RED }
    ]},
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'top'}, tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF`}} },
      scales:{ x:{ ticks:{callback:v=>toCHF(v)} } }
    }
  });
}

// ============================================================
// PARTIE 2 — IMPUTATIONS
// ============================================================
function extractTypeHono(nomTache) {
  if (!nomTache || typeof nomTache !== 'string') return nomTache || 'Autre';
  return nomTache.replace(/_(?:Réseau CAD|Sous-stations|Général|Centrale|Bâtiment|Participations|SIA).*$/, '').trim() || nomTache;
}

function processImputations() {
  const isHono = r => r['Nom de la tâche'] && r['Nom de la tâche'].toLowerCase().includes('honoraires');

  // KPI
  const honoBR   = prevRows.filter(r => isHono(r) && r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
  const honoAT   = prevRows.filter(r => isHono(r) && r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
  const honoR    = realRows.filter(isHono).reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
  const reste    = honoAT - honoR;
  kpiHono.style.display = 'grid';
  kpiHono.innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Honoraires BR</div><div class="kpi-value">${toCHF(honoBR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Honoraires AT</div><div class="kpi-value">${toCHF(honoAT)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Honoraires Réels</div><div class="kpi-value">${toCHF(honoR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Reste à imputer (AT)</div><div class="kpi-value ${reste>=0?'positive':'negative'}">${toCHF(reste)}</div><div class="kpi-unit">CHF</div></div>
  `;

  // Graphe honoraires par type
  const typesBR={}, typesAT={}, typesR={};
  prevRows.forEach(r => {
    if (!isHono(r)) return;
    const t = extractTypeHono(r['Nom de la tâche']);
    const v = toNum(r['Montant total du coût']);
    if (r['Modèle de prévision']==='Budget_Rev') typesBR[t]=(typesBR[t]||0)+v;
    if (r['Modèle de prévision']==='AT')         typesAT[t]=(typesAT[t]||0)+v;
  });
  realRows.forEach(r => {
    if (!isHono(r)) return;
    const t = extractTypeHono(r['Nom de la tâche']);
    typesR[t] = (typesR[t]||0) + toNum(r['Montant total du coût']);
  });
  const allTypes = [...new Set([...Object.keys(typesBR),...Object.keys(typesAT),...Object.keys(typesR)])];
  if (chartHonoMetier) chartHonoMetier.destroy();
  chartHonoMetier = new Chart(document.getElementById('chartHonoMetier').getContext('2d'), {
    type:'bar',
    data:{ labels:allTypes, datasets:[
      { label:'Budget Révisé', data:allTypes.map(t=>typesBR[t]||0), backgroundColor:GE_ORANGE },
      { label:'AT',            data:allTypes.map(t=>typesAT[t]||0), backgroundColor:GE_BLUE2 },
      { label:'Réel',          data:allTypes.map(t=>typesR[t]||0),  backgroundColor:GE_RED }
    ]},
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'top'}, tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF`}} },
      scales:{ x:{ticks:{callback:v=>toCHF(v)}} }
    }
  });

  // Tableau ressources — HONORAIRES UNIQUEMENT — totaux par ressource
  const ressMap = {};
  realRows.forEach(r => {
    if (!isHono(r)) return;   // filtre honoraires
    const nomT = r['Nom de la tâche'] || '';
    const ress = r['Ressource.Nom de la ressource'] || (r['Nom du Fournisseur'] ? `Fourn.: ${r['Nom du Fournisseur']}` : 'Non spécifié');
    const type = extractTypeHono(nomT);
    const proj = getPN(r['ID Projet'] || '');
    const key  = `${ress}||${type}||${proj}`;
    if (!ressMap[key]) ressMap[key] = { ress, type, proj, total: 0 };
    ressMap[key].total += toNum(r['Montant total du coût']);
  });

  const rows = Object.values(ressMap).filter(r=>r.total!==0).sort((a,b)=>b.total-a.total);
  tableRess.innerHTML = `<thead><tr><th>Ressource</th><th>Type d'Honoraires</th><th>Projet</th><th>Total Réel (CHF)</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb = tableRess.querySelector('tbody');
  const tf = tableRess.querySelector('tfoot');
  let tot = 0;
  rows.forEach(r => {
    tot += r.total;
    tb.insertAdjacentHTML('beforeend', `<tr><td>${r.ress}</td><td>${r.type}</td><td>${r.proj}</td><td class="${r.total<0?'negative':''}">${toCHF(r.total)}</td></tr>`);
  });
  tf.innerHTML = `<tr><td colspan="3"><strong>TOTAL</strong></td><td><strong>${toCHF(tot)}</strong></td></tr>`;
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
      return dates.length ? { s: new Date(Math.min(...dates)), e: new Date(Math.max(...dates)) } : { s:null, e:null };
    };
    const br   = getR('Budget_Rev');
    const at   = getR('AT');
    const man  = manualPlanningDates[id] || {};
    let rS = man.start || null, rE = man.end || null;
    if (!rS || !rE) {
      const rd = realRows.filter(r=>r['ID Projet']===id).map(r=>toDate(r['Date du projet'])).filter(Boolean);
      if (rd.length) { if (!rS) rS=new Date(Math.min(...rd)); if (!rE) rE=new Date(Math.max(...rd)); }
    }
    plan[id] = { name:getPN(id), brS:br.s, brE:br.e, atS:at.s, atE:at.e, rS, rE };
  });
  buildTablePlanning(plan);
  drawGantt(plan);
}

function buildTablePlanning(plan) {
  tablePlan.innerHTML = `<thead><tr><th>Projet</th><th>BR Début</th><th>BR Fin</th><th>AT Début</th><th>AT Fin</th><th>Réel Début</th><th>Réel Fin</th><th>Écart (j)</th><th>Statut</th></tr></thead><tbody></tbody>`;
  const fmt  = d => d&&!isNaN(d) ? d.toISOString().split('T')[0] : '';
  const disp = d => d&&!isNaN(d) ? d.toLocaleDateString('fr-CH') : '—';
  const tb   = tablePlan.querySelector('tbody');
  PROJECT_ORDER.forEach(id => {
    const p = plan[id];
    const ecart = p.atE && p.rE ? Math.round((p.rE - p.atE)/86400000) : null;
    let stat='prevu', sLab='Prévu';
    if (p.rS && !p.rE)         { stat='en-cours'; sLab='En cours'; }
    else if (p.rE) { stat = ecart > 0 ? 'retard' : 'termine'; sLab = ecart > 0 ? 'Retard' : 'Terminé'; }
    tb.insertAdjacentHTML('beforeend', `
      <tr>
        <td><strong>${p.name}</strong></td>
        <td>${disp(p.brS)}</td><td>${disp(p.brE)}</td>
        <td>${disp(p.atS)}</td><td>${disp(p.atE)}</td>
        <td><input type="date" id="rs_${id}" class="date-input" value="${fmt(p.rS)}"/></td>
        <td><input type="date" id="re_${id}" class="date-input" value="${fmt(p.rE)}"/></td>
        <td>${ecart!==null?ecart:'—'}</td>
        <td><span class="badge ${stat}">${sLab}</span></td>
      </tr>`);
  });
}

function drawGantt(plan) {
  const allD = [];
  PROJECT_ORDER.forEach(id => {
    const p = plan[id];
    [p.brS,p.brE,p.atS,p.atE,p.rS,p.rE].forEach(d=>{ if(d&&!isNaN(d)) allD.push(d); });
  });
  if (!allD.length) return;
  const minD   = new Date(Math.min(...allD));
  const toDay  = d => d&&!isNaN(d) ? Math.round((d-minD)/86400000) : null;
  const datasets = [];

  PROJECT_ORDER.forEach((id, idx) => {
    const p = plan[id];
    const y = idx * 3; // 3 lignes par projet: BR, AT, Réel
    if (p.brS && p.brE) datasets.push({ label:`${p.name} — Budget Révisé`, data:[{x:[toDay(p.brS),toDay(p.brE)],y}],    backgroundColor:GE_ORANGE, borderColor:GE_ORANGE, borderWidth:1, borderSkipped:false });
    if (p.atS && p.atE) datasets.push({ label:`${p.name} — AT`,            data:[{x:[toDay(p.atS),toDay(p.atE)],y:y+1}],  backgroundColor:GE_BLUE2,  borderColor:GE_BLUE2,  borderWidth:1, borderSkipped:false });
    if (p.rS  && p.rE)  datasets.push({ label:`${p.name} — Réel`,          data:[{x:[toDay(p.rS), toDay(p.rE)], y:y+2}], backgroundColor:GE_RED,    borderColor:GE_RED,    borderWidth:1, borderSkipped:false });
  });

  const numRows = PROJECT_ORDER.length * 3;
  const yLabels = [];
  PROJECT_ORDER.forEach(id => {
    yLabels.push(getPN(id) + ' (BR)', getPN(id) + ' (AT)', getPN(id) + ' (Réel)');
  });

  if (chartGantt) chartGantt.destroy();
  chartGantt = new Chart(document.getElementById('chartGantt').getContext('2d'), {
    type: 'bar',
    data: { datasets },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }, // légende via HTML
        tooltip: {
          callbacks: {
            title: ctx => ctx[0].dataset.label,
            label: ctx => {
              const [d0, d1] = ctx.parsed.x;
              const sD = new Date(minD.getTime() + d0 * 86400000);
              const eD = new Date(minD.getTime() + d1 * 86400000);
              return `${sD.toLocaleDateString('fr-CH')} – ${eD.toLocaleDateString('fr-CH')}`;
            }
          }
        }
      },
      scales: {
        x: {
          type: 'linear',
          title: { display: true, text: 'Timeline' },
          ticks: {
            callback: value => {
              const d = new Date(minD.getTime() + value * 86400000);
              return d.toLocaleDateString('fr-CH', { month: 'short', year: '2-digit' });
            },
            maxTicksLimit: 14
          },
          grid: { color: '#eef2f7' }
        },
        y: {
          type: 'linear',
          min: -0.5,
          max: numRows - 0.5,
          ticks: {
            stepSize: 1,
            callback: value => yLabels[value] || ''
          },
          grid: { color: '#eef2f7' }
        }
      }
    }
  });
}

// ============================================================
// EXPORTS INDIVIDUELS
// ============================================================
function exportChartPNG(canvasId, filename) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const a = document.createElement('a');
  a.download = filename + '.png';
  a.href = canvas.toDataURL('image/png', 1);
  a.click();
}

window.exportChartPNG = exportChartPNG;

async function exportTablePNG(cardId, filename) {
  const el = document.getElementById(cardId);
  if (!el) return;
  const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff', logging: false });
  canvas.toBlob(blob => {
    const url = URL.createObjectURL(blob);
    const a   = document.createElement('a');
    a.download = filename + '.png';
    a.href     = url;
    a.click();
    URL.revokeObjectURL(url);
  });
}

window.exportTablePNG = exportTablePNG;

function exportTableXLSX(tableId, sheetName) {
  const el = document.getElementById(tableId);
  if (!el) return;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(el), sheetName.substring(0, 31));
  XLSX.writeFile(wb, sheetName + '.xlsx');
}

window.exportTableXLSX = exportTableXLSX;

// ============================================================
// PDF RAPPORT COMPLET
// ============================================================
async function generatePDF() {
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const W   = pdf.internal.pageSize.getWidth();
  const H   = pdf.internal.pageSize.getHeight();

  // Nom du fichier: AAAAMMJJ_Etat_de_Projet
  const today = new Date();
  const fname = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}_Etat_de_Projet.pdf`;

  // Page de garde
  pdf.setFillColor(22, 58, 95);
  pdf.rect(0, 0, W, H, 'F');
  pdf.setTextColor(255, 255, 255);
  pdf.setFontSize(26);
  pdf.setFont('helvetica', 'bold');
  pdf.text('Dashboard Projets CAD', W/2, H/2 - 20, { align: 'center' });
  pdf.setFontSize(16);
  pdf.setFont('helvetica', 'normal');
  pdf.text('Groupe E Celsius — État de Projet SIA 5', W/2, H/2, { align: 'center' });
  pdf.setFontSize(12);
  pdf.text(`Généré le ${today.toLocaleDateString('fr-CH', { day:'2-digit', month:'long', year:'numeric' })}`, W/2, H/2 + 16, { align: 'center' });

  // Capture des graphiques et tableaux
  const sections = [
    { id: 'cardCourbesS',   title: 'Courbes S — Coûts Cumulés' },
    { id: 'tablePrevCard',  title: 'Budget Révisé — Détails Mensuels' },
    { id: 'tableCmpCard',   title: 'Comparatif par Projet' },
    { id: 'tableRessCard',  title: 'Coûts Réels par Ressource (Honoraires)' },
    { id: 'cardGantt',      title: 'Planning Gantt' },
    { id: 'tablePlanCard',  title: 'Tableau Planning' }
  ];

  for (const sec of sections) {
    const el = document.getElementById(sec.id);
    if (!el) continue;
    const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff', logging: false });
    const imgData = canvas.toDataURL('image/jpeg', 0.92);
    const ratio   = canvas.width / canvas.height;
    const imgW    = W - 20;
    const imgH    = imgW / ratio;

    pdf.addPage();
    // Bandeau titre
    pdf.setFillColor(22, 58, 95);
    pdf.rect(0, 0, W, 14, 'F');
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(11);
    pdf.setFont('helvetica', 'bold');
    pdf.text(sec.title, 10, 9);
    pdf.setFontSize(9);
    pdf.setFont('helvetica', 'normal');
    pdf.text(fname.replace('.pdf',''), W - 10, 9, { align: 'right' });

    // Image centrée
    const yImg = 18;
    const hMax = H - yImg - 6;
    const hFin = imgH > hMax ? hMax : imgH;
    pdf.addImage(imgData, 'JPEG', 10, yImg, imgW, hFin);
  }

  pdf.save(fname);
}

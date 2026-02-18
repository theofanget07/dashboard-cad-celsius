/* ============================================================
   Dashboard CAD — Groupe E Celsius
   script.js v5.1 — 18/02/2026
   Corrections : Aujourd'hui partout, vue Global/S globale,
   tableau mensuel corrigé toutes phases, honoraires SIA3 fix,
   exports Excel/PNG réintégrés
   ============================================================ */

// ============================================================
// GLOBALS
// ============================================================
let prevRows = [], realRows = [];
const charts = {}; // Stockage centralisé Chart.js
let manualPlanningDates = {}; // Dates éditées manuellement planning SIA5

// ---- Palette Groupe E ----
const GE_BLUE   = '#163a5f';
const GE_BLUE2  = '#1e4d8c';
const GE_BLUE3  = '#2e6db4';
const GE_BLUE4  = '#5b9bd5';
const GE_ORANGE = '#e8873a';
const GE_RED    = '#c0392b';
const GE_GREEN  = '#27ae60';

// Camembert : 12 couleurs distinctes
const PIE_COLORS = [
  '#2e6db4','#e8873a','#27ae60','#8e44ad',
  '#c0392b','#16a085','#d4ac0d','#2980b9',
  '#e74c3c','#1abc9c','#f39c12','#7f8c8d'
];

// SIA 5 — 6 sous-projets
const PROJECT_NAMES = {
  'CL150096_11_05_01': 'Général - Direction',
  'CL150096_11_05_02': 'Centrale',
  'CL150096_11_05_03': 'Bâtiment',
  'CL150096_11_05_04': 'Réseau CAD',
  'CL150096_11_05_05': 'Sous-stations',
  'CL150096_11_05_06': 'Participations clients'
};
const PROJECT_ORDER = Object.keys(PROJECT_NAMES);
const getPN = id => PROJECT_NAMES[id] || id;

// ============================================================
// UTILS
// ============================================================

function readExcel(file, cb) {
  const r = new FileReader();
  r.onload = e => {
    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    cb(XLSX.utils.sheet_to_json(sheet, { raw: true }));
  };
  r.readAsArrayBuffer(file);
}

function toNum(v) {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string') return 0;
  return parseFloat(
    v.replace(/[\u00A0\u202F']/g, '')
     .replace(/[ ]/g, '')
     .replace(',', '.')
  ) || 0;
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
    const m = cleaned.match(/^(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})$/);
    if (m) return new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2]-1, +m[1]);
  }
  return null;
}

function toDateInput(d) {
  if (!d) return '';
  const dd = new Date(d);
  if (isNaN(dd)) return '';
  return `${dd.getFullYear()}-${String(dd.getMonth()+1).padStart(2,'0')}-${String(dd.getDate()).padStart(2,'0')}`;
}

function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function yyyymm(d) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; }
function fmtMonth(mm) {
  const [y, m] = mm.split('-');
  const months = ['Jan','Fév','Mar','Avr','Mai','Juin','Juil','Août','Sep','Oct','Nov','Déc'];
  return `${months[+m-1]} ${y}`;
}

function buildAxis(minD, maxD) {
  const axis = [], c = firstOfMonth(new Date(minD)), e = firstOfMonth(new Date(maxD));
  while (c <= e) { axis.push(yyyymm(c)); c.setMonth(c.getMonth()+1); }
  return axis;
}

function sCurve(total, axis, k, sIdx, eIdx) {
  const N = axis.length, cum = Array(N).fill(null), L = eIdx - sIdx + 1;
  if (L <= 0 || total === 0 || sIdx < 0 || eIdx < 0) return cum;
  const s0 = 1/(1+Math.exp(-k*(0-0.5))), s1 = 1/(1+Math.exp(-k*(1-0.5)));
  for (let j = 0; j < L; j++) {
    const i = sIdx + j;
    if (i >= N) break;
    const x = j/(L-1||1), sx = 1/(1+Math.exp(-k*(x-0.5)));
    cum[i] = ((sx-s0)/(s1-s0)) * total;
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
  if (sIdx < 0) sIdx = 0;
  if (eIdx < 0) eIdx = axis.length-1;
  return [sIdx, eIdx];
}

function extractPhase(idProjet) {
  if (!idProjet) return null;
  const parts = String(idProjet).split('_');
  if (parts.length < 3) return null;
  const code = parts[2];
  const mapping = { '02':'SIA2', '03':'SIA3', '04':'SIA4', '05':'SIA5' };
  return mapping[code] || null;
}

// CORRECTION : Détection honoraires renforcée
function isHonoraires(row) {
  const task = String(row['Nom de la tâche'] || row['Nom de la tache'] || '').toLowerCase();
  const type = String(row['Type de ressource'] || '').toLowerCase();
  const desc = String(row['Description'] || '').toLowerCase();
  
  // Mots-clés honoraires
  const keywords = ['honorai', 'honor', 'études', 'etude', 'ingénieur', 'ingenieur', 
                    'bureau', 'conseil', 'expert', 'consultant', 'architecture',
                    'cpt', 'responsable projet', 'chef ee', 'technicien', 'conducteur travaux', 'dessinateur'];
  
  return keywords.some(kw => task.includes(kw) || type.includes(kw) || desc.includes(kw));
}

function isExternal(row) {
  const fournisseur = row['Nom du fournisseur'] || row['Fournisseur'] || row['Nom Fournisseur'] || '';
  const task = String(row['Nom de la tâche'] || row['Nom de la tache'] || '').toLowerCase();
  return String(fournisseur).trim() !== '' || task.includes('externe');
}

function getTaskName(row) {
  return String(row['Nom de la tâche'] || row['Nom de la tache'] || row['Tâche'] || row['Tache'] || 'Inconnu').trim();
}

function getDateReal(row) {
  return toDate(row['Date'] || row['Date du projet'] || row['date'] || null);
}

function destroyChart(key) {
  if (charts[key]) { try { charts[key].destroy(); } catch(e){} delete charts[key]; }
}

function kpiCard(label, value, unit='CHF', cls='') {
  return `<div class="kpi-card">
    <div class="kpi-label">${label}</div>
    <div class="kpi-value ${cls}">${value}</div>
    <div class="kpi-unit">${unit}</div>
  </div>`;
}

// ============================================================
// PLUGIN : Ligne "Aujourd'hui" (CORRECTION : actif partout)
// ============================================================
const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(ch, _args, opts) {
    if (!opts || !opts.label) return;
    const { ctx, chartArea: { top, bottom }, scales: { x } } = ch;
    if (!x) return; // Pas d'axe x (ex: pie)
    const labels = ch.data.labels;
    if (!labels || !labels.length) return;
    
    const toT = s => {
      const parts = String(s).match(/(\d{4})-(\d{2})/);
      if (!parts) return 0;
      return new Date(+parts[1], +parts[2]-1, 1).getTime();
    };
    
    const targetT = toT(opts.label);
    let bi = 0, bd = Infinity;
    labels.forEach((l,i) => {
      const labelStr = typeof l === 'string' ? l : String(l);
      // Extraire YYYY-MM depuis "Jan 2025" ou "2025-01"
      const match = labelStr.match(/(\d{4})/);
      if (!match) return;
      const year = match[1];
      const monthNames = ['jan','fév','fev','mar','avr','mai','juin','juil','août','aout','sep','oct','nov','déc','dec'];
      let monthIdx = 0;
      const lowerLabel = labelStr.toLowerCase();
      monthNames.forEach((mn, idx) => {
        if (lowerLabel.includes(mn)) monthIdx = idx;
      });
      const reconstructed = `${year}-${String(monthIdx+1).padStart(2,'0')}`;
      const d = Math.abs(toT(reconstructed) - targetT);
      if (d < bd) { bd = d; bi = i; }
    });
    
    const xp = x.getPixelForValue(bi);
    ctx.save();
    ctx.strokeStyle = GE_RED;
    ctx.lineWidth = 2;
    ctx.setLineDash([5,5]);
    ctx.beginPath();
    ctx.moveTo(xp, top);
    ctx.lineTo(xp, bottom);
    ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = GE_RED;
    ctx.font = 'bold 11px Arial';
    ctx.textAlign = 'center';
    ctx.fillText("Aujourd'hui", xp, top-5);
    ctx.restore();
  }
};
Chart.register(todayLinePlugin);

// ============================================================
// EVENTS
// ============================================================
document.addEventListener('DOMContentLoaded', () => {

  // Tabs
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    });
  });

  // Sous-onglets SIA 5
  document.querySelectorAll('.subtab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.subtab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.subtab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.subtab).classList.add('active');
    });
  });

  // Upload
  document.getElementById('prevFile').addEventListener('change', e => {
    if (e.target.files[0]) readExcel(e.target.files[0], rows => {
      prevRows = rows;
      document.getElementById('status').textContent = `✅ Prévisions chargées (${rows.length} lignes)`;
      maybeProcess();
    });
  });
  document.getElementById('realFile').addEventListener('change', e => {
    if (e.target.files[0]) readExcel(e.target.files[0], rows => {
      realRows = rows;
      document.getElementById('status').textContent = `✅ Réalisations chargées (${rows.length} lignes)`;
      maybeProcess();
    });
  });

  // k coefficient
  ['coefK','coefK_num'].forEach(id => {
    document.getElementById(id).addEventListener('input', function() {
      document.getElementById('coefK').value = this.value;
      document.getElementById('coefK_num').value = this.value;
      if (prevRows.length && realRows.length) processAll();
    });
  });

  // Vue Global/S
  ['viewGlobal','viewSOnly'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => {
      if (prevRows.length && realRows.length) processAll();
    });
  });

  // PDF
  document.getElementById('btnPDF').addEventListener('click', generatePDF);
});

function maybeProcess() {
  if (prevRows.length && realRows.length) {
    const s = document.getElementById('status');
    s.textContent = '⏳ Calcul en cours...';
    s.className = 'status';
    setTimeout(() => {
      processAll();
      s.textContent = '✅ Dashboard mis à jour avec succès.';
      s.className = 'status success';
    }, 50);
  }
}

function processAll() {
  processRecap();
  processPhase('SIA2','02');
  processPhase('SIA3','03');
  processPhase('SIA4','04');
  processSIA5Budget();
  processSIA5Imputations();
  processPlanningSIA5();
}

// ============================================================
// FILTRAGES PAR PHASE
// ============================================================
function filterPrevByPhase(code) {
  return prevRows.filter(r => {
    const id = String(r['ID projet'] || r['ID Projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === code;
  });
}

function filterRealByPhase(code) {
  return realRows.filter(r => {
    const id = String(r['ID projet'] || r['ID Projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === code;
  });
}

// ============================================================
// COURBES S GÉNÉRIQUES
// ============================================================
function buildSCurves(pRows, rRows) {
  const k = parseFloat(document.getElementById('coefK_num').value) || 12;
  const isSOnly = document.getElementById('viewSOnly').checked;
  const MODELS = ['Budget','Budget_Rev','AT'];
  const model = {};

  MODELS.forEach(mn => {
    const rows = pRows.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a,r) => a + toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = {
      total,
      start: dates.length ? new Date(Math.min(...dates)) : null,
      end:   dates.length ? new Date(Math.max(...dates)) : null
    };
  });

  const realDates = rRows.map(r => getDateReal(r)).filter(Boolean);
  const allDates = [];
  MODELS.forEach(mn => { if(model[mn].start) allDates.push(model[mn].start); if(model[mn].end) allDates.push(model[mn].end); });
  realDates.forEach(d => allDates.push(d));
  if (!allDates.length) return null;

  const globalMin = new Date(Math.min(...allDates));
  const globalMax = new Date(Math.max(...allDates));
  const modelStartDates = MODELS.filter(mn => model[mn].start).map(mn => model[mn].start);
  const modelEndDates   = MODELS.filter(mn => model[mn].end).map(mn => model[mn].end);
  const modelMin = modelStartDates.length ? new Date(Math.min(...modelStartDates)) : globalMin;
  const modelMax = modelEndDates.length   ? new Date(Math.max(...modelEndDates))   : globalMax;
  const axisMin  = isSOnly ? modelMin : globalMin;
  const axisMax  = isSOnly ? modelMax : globalMax;
  const axis = buildAxis(axisMin, axisMax);

  const series = {};
  MODELS.forEach(mn => {
    const filtered = pRows.filter(r => r['Modèle de prévision'] === mn);
    const [si, ei] = modelRange(filtered.length ? filtered : pRows, mn, axis);
    series[mn] = sCurve(model[mn].total, axis, k, si, ei);
  });

  const today = new Date(), mmToday = yyyymm(today);
  const realMap = new Map(axis.map(m => [m, 0]));
  let initCum = 0;
  rRows.forEach(r => {
    const d = getDateReal(r);
    if (!d) return;
    const mm = yyyymm(firstOfMonth(d));
    if (mm > mmToday) return;
    if (realMap.has(mm)) realMap.set(mm, realMap.get(mm) + toNum(r['Montant total du coût']));
    else if (isSOnly && mm < axis[0]) initCum += toNum(r['Montant total du coût']);
  });

  let acc = initCum;
  const realMonth = axis.map(m => realMap.get(m) || 0);
  const realCum   = realMonth.map(v => { acc += v; return acc; });
  const realCut   = realCum.map((v,i) => axis[i] > mmToday ? null : v);

  return { axis, series, realMonth, realCum, realCut, model, mmToday, initCum };
}

function drawSCurveChart(canvasId, data, title) {
  destroyChart(canvasId);
  if (!data) return;
  const { axis, series, realCut, mmToday } = data;

  const datasets = [];
  if (series['Budget_Rev'] && series['Budget_Rev'].some(v => v !== null && v > 0)) {
    datasets.push({
      label: 'Budget Révisé (S)',
      data: series['Budget_Rev'],
      borderColor: GE_ORANGE, backgroundColor: GE_ORANGE+'22',
      tension: 0.4, fill: false, borderWidth: 2.5,
      pointRadius: 0, pointHoverRadius: 5
    });
  }
  if (series['AT'] && series['AT'].some(v => v !== null && v > 0)) {
    datasets.push({
      label: 'AT (S)',
      data: series['AT'],
      borderColor: GE_BLUE2, backgroundColor: GE_BLUE2+'22',
      tension: 0.4, fill: false, borderWidth: 2.5,
      borderDash: [8,4],
      pointRadius: 0, pointHoverRadius: 5
    });
  }
  if (realCut && realCut.some(v => v !== null && v > 0)) {
    datasets.push({
      label: 'Réel cumulé',
      data: realCut,
      borderColor: GE_RED, backgroundColor: GE_RED+'33',
      tension: 0.1, fill: false, borderWidth: 3,
      pointRadius: 2, pointHoverRadius: 6
    });
  }
  if (!datasets.length) return;

  const ctx = document.getElementById(canvasId);
  if (!ctx) return;
  charts[canvasId] = new Chart(ctx, {
    type: 'line',
    data: { labels: axis.map(fmtMonth), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { font: { size: 12 }, usePointStyle: true } },
        title: { display: false },
        todayLine: { label: mmToday }
      },
      scales: {
        x: { ticks: { maxRotation: 45, font: { size: 10 } }, grid: { color: '#e8eef5' } },
        y: {
          ticks: { callback: v => toCHF(v), font: { size: 10 } },
          grid: { color: '#e8eef5' }
        }
      }
    }
  });
}

// ============================================================
// TABLEAU MENSUEL (CORRECTION : logique réécrite)
// ============================================================
function renderMonthlyTable(tableId, sData) {
  if (!sData) return;
  const { axis, series, realMonth, realCum, mmToday, initCum } = sData;
  const tbl = document.getElementById(tableId);
  if (!tbl) return;

  const brCum  = series['Budget_Rev'] || [];
  const atCum  = series['AT'] || [];

  // Dériver mensuel depuis cumul
  const getMonthly = cum => cum.map((v,i) => {
    if (v === null) return 0;
    const prev = i > 0 ? (cum[i-1] || 0) : 0;
    return Math.max(0, v - prev);
  });
  const brMonthly = getMonthly(brCum);

  // Cumul réel CORRIGÉ : tenir compte de initCum
  let cumReel = initCum || 0;
  const realCumArr = realMonth.map(v => { cumReel += v; return cumReel; });

  let html = `<thead><tr>
    <th>Mois</th>
    <th>BR Mensuel (S)</th><th>BR Cumulé</th>
    <th>Réel Mensuel</th><th>Réel Cumulé</th>
    <th>Écart Mensuel</th><th>Écart Cumulé</th>
  </tr></thead><tbody>`;

  axis.forEach((mm, i) => {
    const brM   = brMonthly[i] || 0;
    const brC   = brCum[i] || 0;
    const reelM = realMonth[i] || 0;
    const reelC = realCumArr[i] || 0;
    const ecM   = reelM - brM;
    const ecC   = reelC - brC;
    const isPast = mm <= mmToday;
    
    // Ne pas afficher les mois futurs sans données
    if (!isPast && brC === 0 && reelM === 0) return;
    
    const rowClass = isPast ? '' : 'style="color:#888;font-style:italic;"';

    html += `<tr ${rowClass}>
      <td><strong>${fmtMonth(mm)}</strong>${mm===mmToday?' <span class="badge en-cours">Aujourd\'hui</span>':''}</td>
      <td>${brM>0 ? toCHF(brM) : '—'}</td>
      <td>${brC>0 ? toCHF(brC) : '—'}</td>
      <td>${isPast ? (reelM>0 ? toCHF(reelM) : '<span style="color:#999">0</span>') : '—'}</td>
      <td>${isPast && reelC>0 ? toCHF(reelC) : (isPast?'<span style="color:#999">0</span>':'—')}</td>
      <td class="${isPast && (reelM>0||brM>0) ? (ecM>=0?'negative':'positive') : ''}">${isPast && (reelM>0||brM>0) ? (ecM>=0?'+':'')+toCHF(ecM) : '—'}</td>
      <td class="${isPast && (reelC>0||brC>0) ? (ecC>=0?'negative':'positive') : ''}">${isPast && (reelC>0||brC>0) ? (ecC>=0?'+':'')+toCHF(ecC) : '—'}</td>
    </tr>`;
  });
  html += `</tbody>`;
  tbl.innerHTML = html;
}

// ============================================================
// ONGLET RÉCAPITULATIF
// ============================================================
function processRecap() {
  const phases = [
    { name: 'SIA 2', code: '02', color: GE_BLUE4 },
    { name: 'SIA 3', code: '03', color: GE_BLUE3 },
    { name: 'SIA 4', code: '04', color: GE_BLUE2 },
    { name: 'SIA 5', code: '05', color: GE_BLUE  }
  ];

  const summaries = phases.map(ph => {
    const pRows = filterPrevByPhase(ph.code);
    const rRows = filterRealByPhase(ph.code);
    const budgetRev= pRows.filter(r => r['Modèle de prévision']==='Budget_Rev').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
    const at       = pRows.filter(r => r['Modèle de prévision']==='AT').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
    const reel     = rRows.reduce((a,r) => a+toNum(r['Montant total du coût']),0);
    const ecartBR  = reel - budgetRev;
    const ecartAT  = reel - at;

    const honorPRows = pRows.filter(r => isHonoraires(r));
    const honorRRows = rRows.filter(r => isHonoraires(r));
    const honorBR = honorPRows.filter(r => r['Modèle de prévision']==='Budget_Rev').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
    const honorAT = honorPRows.filter(r => r['Modèle de prévision']==='AT').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
    const honorReel = honorRRows.reduce((a,r) => a+toNum(r['Montant total du coût']),0);

    const allDates = [...pRows, ...rRows].flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin']), getDateReal(r)]).filter(Boolean);
    const planStart = allDates.length ? new Date(Math.min(...allDates)) : null;
    const planEnd   = allDates.length ? new Date(Math.max(...allDates)) : null;

    return { ...ph, budgetRev, at, reel, ecartBR, ecartAT, honorBR, honorAT, honorReel, planStart, planEnd };
  });

  // KPI Budget
  const totalBR   = summaries.reduce((a,s) => a+s.budgetRev, 0);
  const totalAT   = summaries.reduce((a,s) => a+s.at, 0);
  const totalReel = summaries.reduce((a,s) => a+s.reel, 0);
  const totalEcart= totalReel - totalAT;
  const kpiEl = document.getElementById('recapBudgetKPI');
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Budget Révisé Total', toCHF(totalBR)) +
    kpiCard('AT Total', toCHF(totalAT)) +
    kpiCard('Réel Total', toCHF(totalReel)) +
    kpiCard('Écart vs AT', (totalEcart>=0?'+':'')+toCHF(totalEcart), 'CHF', totalEcart>=0?'negative':'positive');

  // Graphique Budget
  destroyChart('recapBudgetChart');
  const rcCtx = document.getElementById('recapBudgetChart');
  if (rcCtx) {
    const nonZero = summaries.filter(s => s.budgetRev > 0 || s.at > 0 || s.reel > 0);
    if (nonZero.length) {
      charts['recapBudgetChart'] = new Chart(rcCtx, {
        type: 'bar',
        data: {
          labels: nonZero.map(s => s.name),
          datasets: [
            { label: 'Budget Révisé', data: nonZero.map(s=>s.budgetRev), backgroundColor: GE_ORANGE+'cc' },
            { label: 'AT',            data: nonZero.map(s=>s.at),        backgroundColor: GE_BLUE2+'cc' },
            { label: 'Réel',          data: nonZero.map(s=>s.reel),      backgroundColor: GE_RED+'cc' }
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { position:'top' } },
          scales: { y: { ticks: { callback: v => toCHF(v) } } }
        }
      });
    }
  }

  // Tableau Budget
  const tbl = document.getElementById('recapBudgetTable');
  if (tbl) {
    let html = `<thead><tr><th>Phase</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th><th>% AT</th></tr></thead><tbody>`;
    summaries.forEach(s => {
      const pAT = s.at>0 ? (s.reel/s.at*100).toFixed(1) : 0;
      html += `<tr>
        <td><strong>${s.name}</strong></td>
        <td>${toCHF(s.budgetRev)}</td><td>${toCHF(s.at)}</td><td>${toCHF(s.reel)}</td>
        <td class="${s.ecartBR>=0?'negative':'positive'}">${(s.ecartBR>=0?'+':'')+toCHF(s.ecartBR)}</td>
        <td class="${s.ecartAT>=0?'negative':'positive'}">${(s.ecartAT>=0?'+':'')+toCHF(s.ecartAT)}</td>
        <td>${pAT}%</td>
      </tr>`;
    });
    html += `</tbody><tfoot><tr>
      <td>TOTAL</td>
      <td>${toCHF(totalBR)}</td><td>${toCHF(totalAT)}</td><td>${toCHF(totalReel)}</td>
      <td class="${(totalReel-totalBR)>=0?'negative':'positive'}">${((totalReel-totalBR)>=0?'+':'')+toCHF(totalReel-totalBR)}</td>
      <td class="${totalEcart>=0?'negative':'positive'}">${(totalEcart>=0?'+':'')+toCHF(totalEcart)}</td>
      <td>${totalAT>0?(totalReel/totalAT*100).toFixed(1):0}%</td>
    </tr></tfoot>`;
    tbl.innerHTML = html;
  }

  // KPI Imputations
  const totHonorBR   = summaries.reduce((a,s)=>a+s.honorBR,0);
  const totHonorAT   = summaries.reduce((a,s)=>a+s.honorAT,0);
  const totHonorReel = summaries.reduce((a,s)=>a+s.honorReel,0);
  const kpiImpEl = document.getElementById('recapImputKPI');
  if (kpiImpEl) kpiImpEl.innerHTML =
    kpiCard('Honoraires BR', toCHF(totHonorBR)) +
    kpiCard('Honoraires AT', toCHF(totHonorAT)) +
    kpiCard('Honoraires Réel', toCHF(totHonorReel)) +
    kpiCard('Écart Honor. AT', ((totHonorReel-totHonorAT)>=0?'+':'')+toCHF(totHonorReel-totHonorAT), 'CHF', (totHonorReel-totHonorAT)>=0?'negative':'positive');

  // Graphique Imputations
  destroyChart('recapImputChart');
  const riCtx = document.getElementById('recapImputChart');
  if (riCtx) {
    const nonZeroH = summaries.filter(s => s.honorBR>0 || s.honorAT>0 || s.honorReel>0);
    if (nonZeroH.length) {
      charts['recapImputChart'] = new Chart(riCtx, {
        type: 'bar',
        data: {
          labels: nonZeroH.map(s=>s.name),
          datasets: [
            { label: 'Honoraires BR', data: nonZeroH.map(s=>s.honorBR),   backgroundColor: GE_ORANGE+'cc' },
            { label: 'Honoraires AT', data: nonZeroH.map(s=>s.honorAT),   backgroundColor: GE_BLUE2+'cc' },
            { label: 'Réel',          data: nonZeroH.map(s=>s.honorReel), backgroundColor: GE_GREEN+'cc'  }
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { position:'top' } },
          scales: { y: { ticks: { callback: v => toCHF(v) } } }
        }
      });
    }
  }

  // Tableau Imputations
  const tblI = document.getElementById('recapImputTable');
  if (tblI) {
    let html = `<thead><tr><th>Phase</th><th>Honoraires BR</th><th>Honoraires AT</th><th>Réel</th><th>Écart AT</th></tr></thead><tbody>`;
    summaries.forEach(s => {
      const ec = s.honorReel - s.honorAT;
      html += `<tr><td><strong>${s.name}</strong></td><td>${toCHF(s.honorBR)}</td><td>${toCHF(s.honorAT)}</td><td>${toCHF(s.honorReel)}</td>
        <td class="${ec>=0?'negative':'positive'}">${(ec>=0?'+':'')+toCHF(ec)}</td></tr>`;
    });
    html += `</tbody>`;
    tblI.innerHTML = html;
  }

  // Gantt Global
  destroyChart('recapPlanningChart');
  const ganttCtx = document.getElementById('recapPlanningChart');
  if (ganttCtx) {
    const validPhases = summaries.filter(s => s.planStart && s.planEnd);
    if (validPhases.length) {
      const allStart = new Date(Math.min(...validPhases.map(s=>s.planStart)));
      const allEnd   = new Date(Math.max(...validPhases.map(s=>s.planEnd)));
      const axis     = buildAxis(allStart, allEnd);
      const bars     = validPhases.map(s => {
        const si = axis.indexOf(yyyymm(firstOfMonth(s.planStart)));
        const ei = axis.indexOf(yyyymm(firstOfMonth(s.planEnd)));
        const base = si >= 0 ? si : 0;
        const len  = ei >= si ? ei - si + 1 : 1;
        return { label: s.name, base, len, color: s.color };
      });
      charts['recapPlanningChart'] = new Chart(ganttCtx, {
        type: 'bar',
        data: {
          labels: validPhases.map(s=>s.name),
          datasets: [
            { label:'Début', data: bars.map(b=>b.base), backgroundColor: 'transparent', borderWidth: 0 },
            { label:'Durée', data: bars.map(b=>b.len), backgroundColor: validPhases.map(s=>s.color+'cc'), borderRadius: 4 }
          ]
        },
        options: {
          indexAxis: 'y', responsive: true, maintainAspectRatio: false,
          plugins: { legend: { display: false }, todayLine: { label: yyyymm(new Date()) } },
          scales: {
            x: { stacked: true, ticks: { callback: v => axis[v] ? fmtMonth(axis[v]) : '' } },
            y: { stacked: true }
          }
        }
      });
      ganttCtx.parentElement.style.height = `${Math.max(200, validPhases.length*80)}px`;
    }
  }
}

// ============================================================
// PHASES SIA 2/3/4 — GÉNÉRIQUE
// ============================================================
function processPhase(phaseName, phaseCode) {
  const pRows = filterPrevByPhase(phaseCode);
  const rRows = filterRealByPhase(phaseCode);
  const phaseNum = phaseCode.replace('0',''); // '2', '3', '4'

  if (!pRows.length && !rRows.length) {
    const kpiEl = document.getElementById(`kpiSection${phaseName}`);
    if (kpiEl) kpiEl.innerHTML = `<div class="hint">ℹ️ Aucune donnée pour ${phaseName}</div>`;
    return;
  }

  // COURBES S
  const sData = buildSCurves(pRows, rRows);
  drawSCurveChart(`chart${phaseName}`, sData, phaseName);

  // TABLEAU MENSUEL (AJOUT PARTOUT)
  renderMonthlyTable(`table${phaseName}Monthly`, sData);

  // KPIs
  const budgetRev = pRows.filter(r => r['Modèle de prévision']==='Budget_Rev').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
  const at        = pRows.filter(r => r['Modèle de prévision']==='AT').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
  const reel      = rRows.reduce((a,r) => a+toNum(r['Montant total du coût']),0);
  const ecart     = reel - at;
  const pct       = at > 0 ? (reel/at*100).toFixed(1) : 0;

  const kpiEl = document.getElementById(`kpiSection${phaseName}`);
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Budget Révisé', toCHF(budgetRev)) +
    kpiCard('AT', toCHF(at)) +
    kpiCard('Réel', toCHF(reel)) +
    kpiCard('Écart vs AT', (ecart>=0?'+':'')+toCHF(ecart), 'CHF', ecart>=0?'negative':'positive') +
    kpiCard('Avancement', `${pct}%`, '% AT', '');

  // TÂCHES
  const taskMap = {};
  pRows.forEach(r => {
    const task   = getTaskName(r);
    const model  = r['Modèle de prévision'];
    const amount = toNum(r['Montant total du coût']);
    if (!taskMap[task]) taskMap[task] = { task, budget:0, budgetRev:0, at:0, reel:0 };
    if (model === 'Budget')     taskMap[task].budget += amount;
    if (model === 'Budget_Rev') taskMap[task].budgetRev += amount;
    if (model === 'AT')         taskMap[task].at += amount;
  });
  rRows.forEach(r => {
    const task = getTaskName(r);
    if (!taskMap[task]) taskMap[task] = { task, budget:0, budgetRev:0, at:0, reel:0 };
    taskMap[task].reel += toNum(r['Montant total du coût']);
  });
  const taskList = Object.values(taskMap).filter(t => t.budget>0 || t.budgetRev>0 || t.at>0 || t.reel>0);

  // TABLEAU BUDGETS PAR TÂCHE
  const tblBudget = document.getElementById(`table${phaseName}Budget`);
  if (tblBudget) {
    let html = `<thead><tr><th>Tâche</th><th>Budget</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th><th>% AT</th></tr></thead><tbody>`;
    taskList.forEach(t => {
      const eBR = t.reel - t.budgetRev, eAT = t.reel - t.at;
      const pAT = t.at>0 ? (t.reel/t.at*100).toFixed(1) : '—';
      html += `<tr>
        <td><strong>${t.task}</strong></td>
        <td>${toCHF(t.budget)}</td><td>${toCHF(t.budgetRev)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
        <td class="${eBR>=0?'negative':'positive'}">${eBR!==0?(eBR>=0?'+':'')+toCHF(eBR):'—'}</td>
        <td class="${eAT>=0?'negative':'positive'}">${eAT!==0?(eAT>=0?'+':'')+toCHF(eAT):'—'}</td>
        <td>${pAT}%</td>
      </tr>`;
    });
    html += `</tbody><tfoot><tr>
      <td>TOTAL</td>
      <td>${toCHF(taskList.reduce((a,t)=>a+t.budget,0))}</td>
      <td>${toCHF(budgetRev)}</td><td>${toCHF(at)}</td><td>${toCHF(reel)}</td>
      <td class="${(reel-budgetRev)>=0?'negative':'positive'}">${((reel-budgetRev)>=0?'+':'')+toCHF(reel-budgetRev)}</td>
      <td class="${ecart>=0?'negative':'positive'}">${(ecart>=0?'+':'')+toCHF(ecart)}</td>
      <td>${pct}%</td>
    </tr></tfoot>`;
    tblBudget.innerHTML = html;
  }

  // CAMEMBERT
  const pieBRId = `pie${phaseName}BR`;
  destroyChart(pieBRId);
  const pieBRCtx = document.getElementById(pieBRId);
  if (pieBRCtx) {
    const validTasks = taskList.filter(t => t.budgetRev > 0);
    if (validTasks.length) {
      charts[pieBRId] = new Chart(pieBRCtx, {
        type: 'doughnut',
        data: {
          labels: validTasks.map(t => t.task.length>22?t.task.substring(0,20)+'…':t.task),
          datasets: [{
            data: validTasks.map(t => t.budgetRev),
            backgroundColor: validTasks.map((_,i) => PIE_COLORS[i % PIE_COLORS.length]),
            borderWidth: 2, borderColor: '#fff'
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: {
            legend: { position:'right', labels: { font:{size:11}, usePointStyle:true } },
            tooltip: { callbacks: { label: ctx => ` ${ctx.label}: CHF ${toCHF(ctx.parsed)}` } }
          }
        }
      });
    }
  }

  // BARRES
  const barId = `bar${phaseName}`;
  destroyChart(barId);
  const barCtx = document.getElementById(barId);
  if (barCtx) {
    const validBar = taskList.filter(t => t.at>0 || t.reel>0);
    if (validBar.length) {
      charts[barId] = new Chart(barCtx, {
        type: 'bar',
        data: {
          labels: validBar.map(t => t.task.length>20 ? t.task.substring(0,18)+'…' : t.task),
          datasets: [
            { label:'AT',   data: validBar.map(t=>t.at),   backgroundColor: GE_BLUE2+'cc' },
            { label:'Réel', data: validBar.map(t=>t.reel), backgroundColor: GE_RED+'cc' }
          ]
        },
        options: {
          indexAxis:'y', responsive:true, maintainAspectRatio:false,
          plugins: { legend:{position:'top'} },
          scales: { x:{ ticks:{ callback: v => toCHF(v) } } }
        }
      });
    }
  }

  // IMPUTATIONS
  processImputationsPhase(phaseName, phaseCode, pRows, rRows);

  // PLANNING
  processPlanningPhase(phaseName, phaseCode, pRows, rRows);
}

// ============================================================
// IMPUTATIONS PAR PHASE
// ============================================================
function processImputationsPhase(phaseName, phaseCode, pRows, rRows) {
  const phaseNum = phaseCode.replace('0','');

  const honorPrev = pRows.filter(r => isHonoraires(r));
  const honorReal = rRows.filter(r => isHonoraires(r));

  const taskMap = {};
  honorPrev.forEach(r => {
    const task  = getTaskName(r);
    const model = r['Modèle de prévision'];
    const ext   = isExternal(r);
    const key   = `${task}__${ext?'ext':'int'}`;
    if (!taskMap[key]) taskMap[key] = { task, ext, br:0, at:0, reel:0 };
    if (model==='Budget_Rev') taskMap[key].br += toNum(r['Montant total du coût']);
    if (model==='AT')          taskMap[key].at += toNum(r['Montant total du coût']);
  });
  honorReal.forEach(r => {
    const task = getTaskName(r);
    const ext  = isExternal(r);
    const key  = `${task}__${ext?'ext':'int'}`;
    if (!taskMap[key]) taskMap[key] = { task, ext, br:0, at:0, reel:0 };
    taskMap[key].reel += toNum(r['Montant total du coût']);
  });

  const items = Object.values(taskMap).filter(t => t.br>0 || t.at>0 || t.reel>0);
  const intItems = items.filter(t => !t.ext);
  const extItems = items.filter(t => t.ext);

  const totBR   = items.reduce((a,t) => a+t.br,   0);
  const totAT   = items.reduce((a,t) => a+t.at,   0);
  const totReel = items.reduce((a,t) => a+t.reel, 0);
  const ecartAT = totReel - totAT;

  const kpiImpEl = document.getElementById(`kpiSectionImpSIA${phaseNum}`);
  if (kpiImpEl) kpiImpEl.innerHTML =
    kpiCard('Honoraires BR', toCHF(totBR)) +
    kpiCard('Honoraires AT', toCHF(totAT)) +
    kpiCard('Réel Honoraires', toCHF(totReel)) +
    kpiCard('Écart AT', (ecartAT>=0?'+':'')+toCHF(ecartAT), 'CHF', ecartAT>=0?'negative':'positive');

  const tbl = document.getElementById(`tableImputSIA${phaseNum}`) || document.getElementById(`tableImpSIA${phaseNum}`);
  if (!tbl) return;

  let html = `<thead><tr><th>Tâche</th><th>Type</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th></tr></thead><tbody>`;

  const renderSection = (list, type, cls, clsSub) => {
    if (!list.length) return '';
    let rows = `<tr class="${cls}"><td colspan="7">${type === 'int' ? '🔵 RESSOURCES INTERNES' : '🟡 RESSOURCES EXTERNES'}</td></tr>`;
    let sBR=0, sAT=0, sReel=0;
    list.forEach(t => {
      const eBR = t.reel-t.br, eAT = t.reel-t.at;
      sBR+=t.br; sAT+=t.at; sReel+=t.reel;
      rows += `<tr class="${type==='int'?'row-int':'row-ext'}">
        <td>${t.task}</td>
        <td><span class="badge ${type==='int'?'prevu':'en-cours'}">${type==='int'?'Interne':'Externe'}</span></td>
        <td>${toCHF(t.br)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
        <td class="${eBR>=0?'negative':'positive'}">${(eBR>=0?'+':'')+toCHF(eBR)}</td>
        <td class="${eAT>=0?'negative':'positive'}">${(eAT>=0?'+':'')+toCHF(eAT)}</td>
      </tr>`;
    });
    const sE = sReel-sAT;
    rows += `<tr class="${clsSub}">
      <td colspan="2"><strong>Sous-total ${type==='int'?'Interne':'Externe'}</strong></td>
      <td>${toCHF(sBR)}</td><td>${toCHF(sAT)}</td><td>${toCHF(sReel)}</td>
      <td class="${(sReel-sBR)>=0?'negative':'positive'}">${((sReel-sBR)>=0?'+':'')+toCHF(sReel-sBR)}</td>
      <td class="${sE>=0?'negative':'positive'}">${(sE>=0?'+':'')+toCHF(sE)}</td>
    </tr>`;
    return rows;
  };

  html += renderSection(intItems, 'int', 'section-header-int', 'subtotal-int');
  html += renderSection(extItems, 'ext', 'section-header-ext', 'subtotal-ext');
  const totalEcatAT = totReel - totAT;
  html += `</tbody><tfoot><tr>
    <td colspan="2"><strong>TOTAL HONORAIRES</strong></td>
    <td>${toCHF(totBR)}</td><td>${toCHF(totAT)}</td><td>${toCHF(totReel)}</td>
    <td class="${(totReel-totBR)>=0?'negative':'positive'}">${((totReel-totBR)>=0?'+':'')+toCHF(totReel-totBR)}</td>
    <td class="${totalEcatAT>=0?'negative':'positive'}">${(totalEcatAT>=0?'+':'')+toCHF(totalEcatAT)}</td>
  </tr></tfoot>`;
  tbl.innerHTML = html;
}

// ============================================================
// PLANNING PHASES SIA 2/3/4
// ============================================================
function processPlanningPhase(phaseName, phaseCode, pRows, rRows) {
  const ganttId = `gantt${phaseName}`;
  destroyChart(ganttId);
  const ctx = document.getElementById(ganttId);
  if (!ctx) return;

  const taskNames = [...new Set(pRows.map(r => getTaskName(r)))];
  if (!taskNames.length) return;

  const ganttData = taskNames.map(task => {
    const taskPrev = pRows.filter(r => getTaskName(r) === task);
    const getBR = taskPrev.filter(r => r['Modèle de prévision']==='Budget_Rev');
    const getAT = taskPrev.filter(r => r['Modèle de prévision']==='AT');
    const realT = rRows.filter(r => getTaskName(r) === task);

    const getDates = rows => {
      const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin']), getDateReal(r)]).filter(Boolean);
      return dates.length ? { start: new Date(Math.min(...dates)), end: new Date(Math.max(...dates)) } : null;
    };

    return {
      task,
      br:   getDates(getBR),
      at:   getDates(getAT),
      reel: getDates(realT)
    };
  }).filter(t => t.br || t.at || t.reel);

  if (!ganttData.length) return;

  const allDates = ganttData.flatMap(t => [t.br?.start, t.br?.end, t.at?.start, t.at?.end, t.reel?.start, t.reel?.end]).filter(Boolean);
  const axisStart = new Date(Math.min(...allDates));
  const axisEnd   = new Date(Math.max(...allDates));
  const axis = buildAxis(axisStart, axisEnd);

  const dateToIdx = d => {
    if (!d) return -1;
    const mm = yyyymm(firstOfMonth(d));
    const i = axis.indexOf(mm);
    return i >= 0 ? i : 0;
  };

  const buildBar = dateObj => {
    if (!dateObj) return [null, null];
    const si = dateToIdx(dateObj.start), ei = dateToIdx(dateObj.end);
    return [si, ei - si + 1];
  };

  const labels    = ganttData.map(t => t.task.length>22 ? t.task.substring(0,22)+'…' : t.task);
  const brBase    = ganttData.map(t => buildBar(t.br)[0]);
  const brLen     = ganttData.map(t => buildBar(t.br)[1]);
  const atBase    = ganttData.map(t => buildBar(t.at)[0]);
  const atLen     = ganttData.map(t => buildBar(t.at)[1]);
  const reelBase  = ganttData.map(t => buildBar(t.reel)[0]);
  const reelLen   = ganttData.map(t => buildBar(t.reel)[1]);

  charts[ganttId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'BR offset',   data:brBase,   backgroundColor:'transparent', borderWidth:0, stack:'br'   },
        { label:'Budget Révisé', data:brLen,  backgroundColor:GE_ORANGE+'cc', borderRadius:3, barThickness:12, stack:'br' },
        { label:'AT offset',   data:atBase,   backgroundColor:'transparent', borderWidth:0, stack:'at'   },
        { label:'AT',          data:atLen,    backgroundColor:GE_BLUE2+'cc',  borderRadius:3, barThickness:12, stack:'at' },
        { label:'Réel offset', data:reelBase, backgroundColor:'transparent', borderWidth:0, stack:'reel' },
        { label:'Réel',        data:reelLen,  backgroundColor:GE_RED+'cc',   borderRadius:3, barThickness:12, stack:'reel' }
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: {
        legend: {
          labels: {
            filter: item => !item.text.includes('offset'),
            usePointStyle:true, font:{size:11}
          }
        },
        todayLine: { label: yyyymm(new Date()) }
      },
      scales: {
        x: { stacked:true, ticks: { callback: v => axis[v] ? fmtMonth(axis[v]) : '' }, font:{size:10} },
        y: { stacked:true }
      }
    }
  });
  ctx.parentElement.style.height = `${Math.max(250, ganttData.length * 55 + 80)}px`;
}

// ============================================================
// SIA 5 — BUDGET
// ============================================================
function processSIA5Budget() {
  const pRows = filterPrevByPhase('05');
  const rRows = filterRealByPhase('05');

  if (!pRows.length && !rRows.length) return;

  const sData = buildSCurves(pRows, rRows);
  drawSCurveChart('chartSIA5', sData, 'SIA5');
  renderMonthlyTable('tableSIA5Monthly', sData);

  const budgetRev = pRows.filter(r => r['Modèle de prévision']==='Budget_Rev').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
  const at        = pRows.filter(r => r['Modèle de prévision']==='AT').reduce((a,r) => a+toNum(r['Montant total du coût']),0);
  const reel      = rRows.reduce((a,r) => a+toNum(r['Montant total du coût']),0);
  const ecart     = reel - at;
  const pct       = at > 0 ? (reel/at*100).toFixed(1) : 0;

  const kpiEl = document.getElementById('kpiSectionSIA5');
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Budget Révisé', toCHF(budgetRev)) +
    kpiCard('AT', toCHF(at)) +
    kpiCard('Réel', toCHF(reel)) +
    kpiCard('Écart vs AT', (ecart>=0?'+':'')+toCHF(ecart), 'CHF', ecart>=0?'negative':'positive') +
    kpiCard('Avancement', `${pct}%`, '% AT', '');

  const tblBudget = document.getElementById('tableSIA5Budget');
  if (tblBudget) {
    const taskMap = {};
    PROJECT_ORDER.forEach(id => { taskMap[id] = { budget:0, budgetRev:0, at:0, reel:0 }; });
    pRows.forEach(r => {
      const id = String(r['ID projet']||r['ID Projet']||'');
      if (!taskMap[id]) taskMap[id] = { budget:0, budgetRev:0, at:0, reel:0 };
      const model = r['Modèle de prévision'], amt = toNum(r['Montant total du coût']);
      if (model==='Budget')     taskMap[id].budget += amt;
      if (model==='Budget_Rev') taskMap[id].budgetRev += amt;
      if (model==='AT')         taskMap[id].at += amt;
    });
    rRows.forEach(r => {
      const id = String(r['ID projet']||r['ID Projet']||'');
      if (!taskMap[id]) taskMap[id] = { budget:0, budgetRev:0, at:0, reel:0 };
      taskMap[id].reel += toNum(r['Montant total du coût']);
    });

    let html = `<thead><tr><th>Sous-Projet</th><th>Budget</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th><th>% AT</th></tr></thead><tbody>`;

    Object.entries(taskMap).forEach(([id, t]) => {
      if (t.budget===0 && t.budgetRev===0 && t.at===0 && t.reel===0) return;
      const eBR=t.reel-t.budgetRev, eAT=t.reel-t.at;
      const pAT=t.at>0?(t.reel/t.at*100).toFixed(1):'—';
      html += `<tr>
        <td><strong>${getPN(id)}</strong></td>
        <td>${toCHF(t.budget)}</td><td>${toCHF(t.budgetRev)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
        <td class="${eBR>=0?'negative':'positive'}">${(eBR>=0?'+':'')+toCHF(eBR)}</td>
        <td class="${eAT>=0?'negative':'positive'}">${(eAT>=0?'+':'')+toCHF(eAT)}</td>
        <td>${pAT}%</td>
      </tr>`;
    });
    html += `</tbody><tfoot><tr>
      <td>TOTAL</td>
      <td>${toCHF(pRows.filter(r=>r['Modèle de prévision']==='Budget').reduce((a,r)=>a+toNum(r['Montant total du coût']),0))}</td>
      <td>${toCHF(budgetRev)}</td><td>${toCHF(at)}</td><td>${toCHF(reel)}</td>
      <td class="${(reel-budgetRev)>=0?'negative':'positive'}">${((reel-budgetRev)>=0?'+':'')+toCHF(reel-budgetRev)}</td>
      <td class="${ecart>=0?'negative':'positive'}">${(ecart>=0?'+':'')+toCHF(ecart)}</td>
      <td>${pct}%</td>
    </tr></tfoot>`;
    tblBudget.innerHTML = html;
  }

  destroyChart('pieSIA5BR');
  const pieSIA5Ctx = document.getElementById('pieSIA5BR');
  if (pieSIA5Ctx) {
    const pieData = PROJECT_ORDER.map(id => ({
      label: getPN(id),
      value: pRows.filter(r => (r['ID projet']||r['ID Projet'])===id && r['Modèle de prévision']==='Budget_Rev')
                  .reduce((a,r)=>a+toNum(r['Montant total du coût']),0)
    })).filter(d => d.value > 0);

    if (pieData.length) {
      charts['pieSIA5BR'] = new Chart(pieSIA5Ctx, {
        type: 'doughnut',
        data: {
          labels: pieData.map(d=>d.label),
          datasets: [{
            data: pieData.map(d=>d.value),
            backgroundColor: pieData.map((_,i)=>PIE_COLORS[i%PIE_COLORS.length]),
            borderWidth:2, borderColor:'#fff'
          }]
        },
        options: {
          responsive:true, maintainAspectRatio:false,
          plugins: {
            legend: { position:'right', labels:{ font:{size:11}, usePointStyle:true } },
            tooltip: { callbacks: { label: ctx => ` ${ctx.label}: CHF ${toCHF(ctx.parsed)}` } }
          }
        }
      });
    }
  }

  destroyChart('barSIA5');
  const barSIA5Ctx = document.getElementById('barSIA5');
  if (barSIA5Ctx) {
    const barData = PROJECT_ORDER.map(id => {
      const atV  = pRows.filter(r=>(r['ID projet']||r['ID Projet'])===id && r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const reelV = rRows.filter(r=>(r['ID projet']||r['ID Projet'])===id).reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      return { label: getPN(id), at: atV, reel: reelV };
    }).filter(d => d.at>0 || d.reel>0);

    if (barData.length) {
      charts['barSIA5'] = new Chart(barSIA5Ctx, {
        type:'bar',
        data:{
          labels: barData.map(d=>d.label),
          datasets:[
            { label:'AT',   data:barData.map(d=>d.at),   backgroundColor:GE_BLUE2+'cc' },
            { label:'Réel', data:barData.map(d=>d.reel), backgroundColor:GE_RED+'cc'   }
          ]
        },
        options:{
          indexAxis:'y', responsive:true, maintainAspectRatio:false,
          plugins:{ legend:{position:'top'} },
          scales:{ x:{ ticks:{ callback:v=>toCHF(v) } } }
        }
      });
    }
  }
}

// ============================================================
// SIA 5 — IMPUTATIONS
// ============================================================
function processSIA5Imputations() {
  const pRows = filterPrevByPhase('05');
  const rRows = filterRealByPhase('05');

  const honorPrev = pRows.filter(r => isHonoraires(r));
  const honorReal = rRows.filter(r => isHonoraires(r));

  const taskMap = {};
  honorPrev.forEach(r => {
    const id    = String(r['ID projet']||r['ID Projet']||'');
    const model = r['Modèle de prévision'];
    const ext   = isExternal(r);
    const key   = `${id}__${ext?'ext':'int'}`;
    if (!taskMap[key]) taskMap[key] = { id, name: getPN(id), ext, br:0, at:0, reel:0 };
    if (model==='Budget_Rev') taskMap[key].br   += toNum(r['Montant total du coût']);
    if (model==='AT')         taskMap[key].at   += toNum(r['Montant total du coût']);
  });
  honorReal.forEach(r => {
    const id  = String(r['ID projet']||r['ID Projet']||'');
    const ext = isExternal(r);
    const key = `${id}__${ext?'ext':'int'}`;
    if (!taskMap[key]) taskMap[key] = { id, name: getPN(id), ext, br:0, at:0, reel:0 };
    taskMap[key].reel += toNum(r['Montant total du coût']);
  });

  const items    = Object.values(taskMap).filter(t => t.br>0 || t.at>0 || t.reel>0);
  const intItems = items.filter(t => !t.ext);
  const extItems = items.filter(t => t.ext);

  const totBR   = items.reduce((a,t)=>a+t.br,  0);
  const totAT   = items.reduce((a,t)=>a+t.at,  0);
  const totReel = items.reduce((a,t)=>a+t.reel, 0);
  const ecartAT = totReel - totAT;

  const kpiEl = document.getElementById('kpiSectionImpSIA5');
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Honoraires BR',    toCHF(totBR)) +
    kpiCard('Honoraires AT',    toCHF(totAT)) +
    kpiCard('Réel Honoraires',  toCHF(totReel)) +
    kpiCard('Écart vs AT', (ecartAT>=0?'+':'')+toCHF(ecartAT), 'CHF', ecartAT>=0?'negative':'positive');

  const tbl = document.getElementById('tableImputSIA5');
  if (!tbl) return;

  let html = `<thead><tr><th>Sous-Projet</th><th>Type</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th></tr></thead><tbody>`;

  const renderSection = (list, type, cls, clsSub) => {
    if (!list.length) return '';
    let rows = `<tr class="${cls}"><td colspan="7">${type==='int'?'🔵 INTERNES':'🟡 EXTERNES'}</td></tr>`;
    let sBR=0, sAT=0, sReel=0;
    list.forEach(t => {
      const eBR=t.reel-t.br, eAT=t.reel-t.at;
      sBR+=t.br; sAT+=t.at; sReel+=t.reel;
      rows += `<tr class="${type==='int'?'row-int':'row-ext'}">
        <td>${t.name}</td>
        <td><span class="badge ${type==='int'?'prevu':'en-cours'}">${type==='int'?'Interne':'Externe'}</span></td>
        <td>${toCHF(t.br)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
        <td class="${eBR>=0?'negative':'positive'}">${(eBR>=0?'+':'')+toCHF(eBR)}</td>
        <td class="${eAT>=0?'negative':'positive'}">${(eAT>=0?'+':'')+toCHF(eAT)}</td>
      </tr>`;
    });
    const sE=sReel-sAT;
    rows += `<tr class="${clsSub}">
      <td colspan="2"><strong>Sous-total ${type==='int'?'Interne':'Externe'}</strong></td>
      <td>${toCHF(sBR)}</td><td>${toCHF(sAT)}</td><td>${toCHF(sReel)}</td>
      <td class="${(sReel-sBR)>=0?'negative':'positive'}">${((sReel-sBR)>=0?'+':'')+toCHF(sReel-sBR)}</td>
      <td class="${sE>=0?'negative':'positive'}">${(sE>=0?'+':'')+toCHF(sE)}</td>
    </tr>`;
    return rows;
  };

  html += renderSection(intItems, 'int', 'section-header-int', 'subtotal-int');
  html += renderSection(extItems, 'ext', 'section-header-ext', 'subtotal-ext');
  html += `</tbody><tfoot><tr>
    <td colspan="2"><strong>TOTAL HONORAIRES</strong></td>
    <td>${toCHF(totBR)}</td><td>${toCHF(totAT)}</td><td>${toCHF(totReel)}</td>
    <td class="${(totReel-totBR)>=0?'negative':'positive'}">${((totReel-totBR)>=0?'+':'')+toCHF(totReel-totBR)}</td>
    <td class="${ecartAT>=0?'negative':'positive'}">${(ecartAT>=0?'+':'')+toCHF(ecartAT)}</td>
  </tr></tfoot>`;
  tbl.innerHTML = html;
}

// ============================================================
// SIA 5 — PLANNING (éditable)
// ============================================================
function processPlanningSIA5() {
  const pRows = filterPrevByPhase('05');
  const rRows = filterRealByPhase('05');

  const planData = PROJECT_ORDER.map(id => {
    const brRows   = pRows.filter(r => (r['ID projet']||r['ID Projet'])===id && r['Modèle de prévision']==='Budget_Rev');
    const atRows   = pRows.filter(r => (r['ID projet']||r['ID Projet'])===id && r['Modèle de prévision']==='AT');
    const reelRows = rRows.filter(r => (r['ID projet']||r['ID Projet'])===id);

    const getDates = rows => {
      const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin']), getDateReal(r)]).filter(Boolean);
      return dates.length ? { start: new Date(Math.min(...dates)), end: new Date(Math.max(...dates)) } : null;
    };

    const manual = manualPlanningDates[id] || {};
    const brBase  = getDates(brRows);
    const atBase  = getDates(atRows);
    const reelBase= getDates(reelRows);

    return {
      id, name: getPN(id),
      br:   { start: manual.brStart || brBase?.start   || null, end: manual.brEnd || brBase?.end   || null },
      at:   { start: manual.atStart || atBase?.start   || null, end: manual.atEnd || atBase?.end   || null },
      reel: { start: reelBase?.start || null, end: reelBase?.end || null }
    };
  });

  const tblPlan = document.getElementById('tablePlanningSIA5');
  if (tblPlan) {
    let html = `<thead><tr>
      <th>Sous-Projet</th>
      <th>BR Début</th><th>BR Fin</th>
      <th>AT Début</th><th>AT Fin</th>
      <th>Réel Début</th><th>Réel Fin</th>
      <th>Statut</th>
    </tr></thead><tbody>`;

    planData.forEach(p => {
      const today = new Date();
      let statut = '', statusClass = '';
      if (p.reel.start && p.reel.end && p.reel.end < today) {
        statut = 'Terminé'; statusClass = 'termine';
      } else if (p.reel.start && p.reel.start <= today) {
        if (p.at.end && today > p.at.end) { statut = 'Retard'; statusClass = 'retard'; }
        else { statut = 'En cours'; statusClass = 'en-cours'; }
      } else {
        statut = 'Prévu'; statusClass = 'prevu';
      }

      html += `<tr>
        <td><strong>${p.name}</strong></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="brStart" value="${toDateInput(p.br.start)}"></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="brEnd"   value="${toDateInput(p.br.end)}"></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="atStart" value="${toDateInput(p.at.start)}"></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="atEnd"   value="${toDateInput(p.at.end)}"></td>
        <td><span style="color:#555;">${p.reel.start ? toDateInput(p.reel.start) : '—'}</span></td>
        <td><span style="color:#555;">${p.reel.end   ? toDateInput(p.reel.end)   : '—'}</span></td>
        <td><span class="badge ${statusClass}">${statut}</span></td>
      </tr>`;
    });
    html += `</tbody>`;
    tblPlan.innerHTML = html;

    // Event temps réel
    tblPlan.querySelectorAll('.plan-input').forEach(input => {
      input.addEventListener('change', () => {
        collectManualDates();
        renderGanttSIA5(planData.map(p => {
          const manual = manualPlanningDates[p.id] || {};
          return {
            ...p,
            br:   { start: manual.brStart || p.br.start,   end: manual.brEnd || p.br.end   },
            at:   { start: manual.atStart || p.at.start,   end: manual.atEnd || p.at.end   }
          };
        }));
      });
    });
  }

  renderGanttSIA5(planData);
}

function collectManualDates() {
  document.querySelectorAll('.plan-input').forEach(input => {
    const id    = input.dataset.id;
    const field = input.dataset.field;
    if (!manualPlanningDates[id]) manualPlanningDates[id] = {};
    manualPlanningDates[id][field] = input.value ? new Date(input.value) : null;
  });
}

function renderGanttSIA5(planData) {
  destroyChart('ganttSIA5');
  const ctx = document.getElementById('ganttSIA5');
  if (!ctx) return;

  const valid = planData.filter(p => p.br.start || p.at.start || p.reel.start);
  if (!valid.length) return;

  const allDates = valid.flatMap(p => [p.br.start,p.br.end,p.at.start,p.at.end,p.reel.start,p.reel.end]).filter(Boolean);
  const axisStart = new Date(Math.min(...allDates));
  const axisEnd   = new Date(Math.max(...allDates));
  const axis = buildAxis(axisStart, axisEnd);

  const toIdx = d => {
    if (!d) return null;
    const mm = yyyymm(firstOfMonth(d));
    const i = axis.indexOf(mm);
    return i >= 0 ? i : null;
  };

  const buildBar = (startD, endD) => {
    const si = toIdx(startD), ei = toIdx(endD);
    if (si === null) return [null, null];
    const len = ei !== null ? Math.max(1, ei - si + 1) : 1;
    return [si, len];
  };

  const labels = valid.map(p => p.name);
  const [brBaseArr, brLenArr]     = [valid.map(p=>buildBar(p.br.start,p.br.end)[0]),   valid.map(p=>buildBar(p.br.start,p.br.end)[1])];
  const [atBaseArr, atLenArr]     = [valid.map(p=>buildBar(p.at.start,p.at.end)[0]),   valid.map(p=>buildBar(p.at.start,p.at.end)[1])];
  const [reelBaseArr, reelLenArr] = [valid.map(p=>buildBar(p.reel.start,p.reel.end)[0]), valid.map(p=>buildBar(p.reel.start,p.reel.end)[1])];

  charts['ganttSIA5'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'BR off',   data:brBaseArr,   backgroundColor:'transparent', borderWidth:0, stack:'br',   barThickness:18 },
        { label:'Budget Révisé', data:brLenArr, backgroundColor:GE_ORANGE+'cc', borderRadius:3, stack:'br',   barThickness:18 },
        { label:'AT off',   data:atBaseArr,   backgroundColor:'transparent', borderWidth:0, stack:'at',   barThickness:18 },
        { label:'AT',       data:atLenArr,    backgroundColor:GE_BLUE2+'cc',  borderRadius:3, stack:'at',   barThickness:18 },
        { label:'Réel off', data:reelBaseArr, backgroundColor:'transparent', borderWidth:0, stack:'reel', barThickness:18 },
        { label:'Réel',     data:reelLenArr,  backgroundColor:GE_RED+'cc',   borderRadius:3, stack:'reel', barThickness:18 }
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: {
        legend: {
          labels: {
            filter: item => !item.text.includes('off'),
            usePointStyle:true, font:{size:12}
          }
        },
        todayLine: { label: yyyymm(new Date()) }
      },
      scales: {
        x: { stacked:true, ticks:{ callback: v => axis[v] ? fmtMonth(axis[v]) : '', font:{size:10} }, grid:{color:'#e8eef5'} },
        y: { stacked:true, ticks:{ font:{size:11} } }
      }
    }
  });
  ctx.parentElement.style.height = `${valid.length * 70 + 80}px`;
}

// ============================================================
// EXPORTS EXCEL / PNG (RÉINTÉGRÉS)
// ============================================================

function exportChartPNG(canvasId) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) { alert('Graphique introuvable'); return; }
  canvas.toBlob(blob => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${canvasId}_${new Date().toISOString().slice(0,10)}.png`;
    a.click();
    URL.revokeObjectURL(url);
  });
}

function exportTableExcel(tableId, filename) {
  const tbl = document.getElementById(tableId);
  if (!tbl) { alert('Tableau introuvable'); return; }
  
  // Convertir table HTML en workbook
  const wb = XLSX.utils.table_to_book(tbl, { sheet: filename });
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ============================================================
// EXPORT PDF (déjà présent, inchangé)
// ============================================================
async function generatePDF() {
  const { jsPDF } = window.jspdf;
  const scope = document.getElementById('pdfScope').value;

  const doc = new jsPDF({ orientation:'portrait', unit:'mm', format:'a4' });
  const pageW = 210, pageH = 297, margin = 12;
  const contentW = pageW - 2*margin;
  let y = margin;
  let pageCount = 1;

  const checkNewPage = (needed=20) => {
    if (y + needed > pageH - margin) {
      doc.addPage();
      pageCount++;
      y = margin;
      renderHeader();
    }
  };

  const renderHeader = () => {
    doc.setFillColor(22, 58, 95);
    doc.rect(0, 0, pageW, 18, 'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(9);
    doc.setFont('helvetica','bold');
    doc.text('GROUPE E CELSIUS — Dashboard CAD CL150096', margin, 12);
    doc.text(`Page ${pageCount}`, pageW - margin, 12, { align:'right' });
    doc.setTextColor(0,0,0);
    y = 22;
  };

  const addTitle = (text, level=1) => {
    checkNewPage(14);
    if (level === 1) {
      doc.setFillColor(22, 58, 95);
      doc.rect(margin, y, contentW, 9, 'F');
      doc.setTextColor(255,255,255);
      doc.setFontSize(10);
      doc.setFont('helvetica','bold');
      doc.text(text, margin+3, y+6.5);
      doc.setTextColor(0,0,0);
      y += 13;
    } else {
      doc.setFillColor(232, 135, 58);
      doc.rect(margin, y, 3, 7, 'F');
      doc.setFontSize(9);
      doc.setFont('helvetica','bold');
      doc.setTextColor(22,58,95);
      doc.text(text, margin+5, y+5.5);
      doc.setTextColor(0,0,0);
      y += 11;
    }
  };

  const addKPIs = (kpiSectionId) => {
    const el = document.getElementById(kpiSectionId);
    if (!el) return;
    const cards = el.querySelectorAll('.kpi-card');
    if (!cards.length) return;
    checkNewPage(20);
    const colW = contentW / Math.min(cards.length, 4);
    cards.forEach((card, i) => {
      const col = i % 4;
      const label = card.querySelector('.kpi-label')?.textContent?.trim() || '';
      const value = card.querySelector('.kpi-value')?.textContent?.trim() || '';
      const unit  = card.querySelector('.kpi-unit')?.textContent?.trim()  || '';
      const x = margin + col * colW;
      const yBox = y;
      doc.setFillColor(240,244,249);
      doc.setDrawColor(200,210,220);
      doc.roundedRect(x+1, yBox, colW-2, 16, 2, 2, 'FD');
      doc.setFontSize(6.5);
      doc.setFont('helvetica','bold');
      doc.setTextColor(90,120,155);
      doc.text(label.toUpperCase(), x + colW/2, yBox+5, { align:'center' });
      doc.setFontSize(10);
      doc.setTextColor(22,58,95);
      doc.text(value, x + colW/2, yBox+11, { align:'center' });
      doc.setFontSize(6);
      doc.setTextColor(130,150,170);
      doc.text(unit, x + colW/2, yBox+15, { align:'center' });
      if (col === 3 || i === cards.length-1) y += 19;
    });
    doc.setTextColor(0,0,0);
  };

  const addTable = (tableId, maxRows=40) => {
    const tbl = document.getElementById(tableId);
    if (!tbl) return;
    const headers = [...tbl.querySelectorAll('thead th')].map(th => th.textContent.trim());
    if (!headers.length) return;

    const colW = contentW / headers.length;
    checkNewPage(12);

    doc.setFillColor(22,58,95);
    doc.rect(margin, y, contentW, 7, 'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(6.5);
    doc.setFont('helvetica','bold');
    headers.forEach((h,i) => {
      const txt = h.length > 14 ? h.substring(0,12)+'.' : h;
      doc.text(txt, margin + i*colW + colW/2, y+5, { align:'center' });
    });
    y += 8;

    const rows = [...tbl.querySelectorAll('tbody tr, tfoot tr')].slice(0, maxRows);
    rows.forEach((row, ri) => {
      checkNewPage(7);
      const cells = [...row.querySelectorAll('td')];
      const isFooter = row.closest('tfoot') !== null;
      if (isFooter) {
        doc.setFillColor(220,230,240);
        doc.rect(margin, y, contentW, 6.5, 'F');
        doc.setFont('helvetica','bold');
      } else {
        if (ri % 2 === 0) { doc.setFillColor(248,251,255); doc.rect(margin, y, contentW, 6.5, 'F'); }
        doc.setFont('helvetica','normal');
      }
      doc.setTextColor(0,0,0);
      doc.setFontSize(6.2);
      cells.forEach((cell, ci) => {
        const txt = (cell.textContent.trim()).replace(/\s+/g,' ');
        const short = txt.length > 18 ? txt.substring(0,16)+'…' : txt;
        const align = ci === 0 ? 'left' : 'center';
        const xPos  = ci === 0 ? margin+2 : margin + ci*colW + colW/2;
        doc.text(short, xPos, y+4.5, { align });
      });
      y += 7;
    });
    y += 4;
  };

  const addChartCanvas = async (canvasId, heightMM=60) => {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    checkNewPage(heightMM + 5);
    try {
      const imgData = canvas.toDataURL('image/png', 0.92);
      if (!imgData || imgData === 'data:,') return;
      doc.addImage(imgData, 'PNG', margin, y, contentW, heightMM);
      y += heightMM + 5;
    } catch(e) { console.warn('Chart capture failed:', canvasId, e); }
  };

  const renderPhasePDF = async (phaseLabel, phaseSuffix) => {
    addTitle(`${phaseLabel} — BUDGETS`, 1);
    addKPIs(`kpiSection${phaseSuffix}`);
    await addChartCanvas(`chart${phaseSuffix}`, 55);
    addTitle('Tableau Mensuel', 2);
    addTable(`table${phaseSuffix}Monthly`, 24);
    addTitle('Tableau Budgets', 2);
    addTable(`table${phaseSuffix}Budget`);
    await addChartCanvas(`pie${phaseSuffix}BR`, 45);
    await addChartCanvas(`bar${phaseSuffix}`, 45);

    addTitle(`${phaseLabel} — IMPUTATIONS`, 1);
    addKPIs(`kpiSectionImpSIA${phaseSuffix.replace('SIA','')}`);
    addTable(`tableImputSIA${phaseSuffix.replace('SIA','')}` , 30);

    addTitle(`${phaseLabel} — PLANNING`, 1);
    await addChartCanvas(`gantt${phaseSuffix}`, 65);
  };

  renderHeader();

  // Page de garde
  doc.setFillColor(22,58,95);
  doc.rect(margin, y, contentW, 28, 'F');
  doc.setTextColor(255,255,255);
  doc.setFontSize(16);
  doc.setFont('helvetica','bold');
  doc.text('Dashboard CAD — Groupe E Celsius', pageW/2, y+12, { align:'center' });
  doc.setFontSize(10);
  doc.setFont('helvetica','normal');
  const scopeLabel = {
    recap:'Récapitulatif Global', sia2:'Phase SIA 2', sia3:'Phase SIA 3',
    sia4:'Phase SIA 4', sia5:'Phase SIA 5', all:'Rapport Complet'
  }[scope] || scope;
  doc.text(scopeLabel, pageW/2, y+20, { align:'center' });
  doc.setFontSize(8);
  doc.setTextColor(200,220,240);
  doc.text(`Généré le ${new Date().toLocaleDateString('fr-CH')} — Projet CL150096`, pageW/2, y+26, { align:'center' });
  doc.setTextColor(0,0,0);
  y += 34;

  if (scope === 'recap' || scope === 'all') {
    addTitle('RÉCAPITULATIF GLOBAL — BUDGETS', 1);
    addKPIs('recapBudgetKPI');
    await addChartCanvas('recapBudgetChart', 55);
    addTable('recapBudgetTable');
    addTitle('RÉCAPITULATIF GLOBAL — IMPUTATIONS', 1);
    addKPIs('recapImputKPI');
    await addChartCanvas('recapImputChart', 50);
    addTable('recapImputTable');
    addTitle('RÉCAPITULATIF GLOBAL — PLANNING', 1);
    await addChartCanvas('recapPlanningChart', 60);
  }

  if (scope === 'sia2' || scope === 'all') await renderPhasePDF('PHASE SIA 2', 'SIA2');
  if (scope === 'sia3' || scope === 'all') await renderPhasePDF('PHASE SIA 3', 'SIA3');
  if (scope === 'sia4' || scope === 'all') await renderPhasePDF('PHASE SIA 4', 'SIA4');

  if (scope === 'sia5' || scope === 'all') {
    addTitle('PHASE SIA 5 — BUDGETS', 1);
    addKPIs('kpiSectionSIA5');
    await addChartCanvas('chartSIA5', 55);
    addTitle('Tableau Mensuel', 2);
    addTable('tableSIA5Monthly', 24);
    addTitle('Tableau Budgets', 2);
    addTable('tableSIA5Budget');
    await addChartCanvas('pieSIA5BR', 45);
    await addChartCanvas('barSIA5', 45);

    addTitle('PHASE SIA 5 — IMPUTATIONS', 1);
    addKPIs('kpiSectionImpSIA5');
    addTable('tableImputSIA5', 25);

    addTitle('PHASE SIA 5 — PLANNING', 1);
    addTable('tablePlanningSIA5');
    await addChartCanvas('ganttSIA5', 65);
  }

  const totalPages = doc.internal.getNumberOfPages();
  for (let i=1; i<=totalPages; i++) {
    doc.setPage(i);
    doc.setFontSize(7.5);
    doc.setTextColor(120,140,160);
    doc.setFont('helvetica','normal');
    doc.text(`Groupe E Celsius — Dashboard CAD CL150096 — Page ${i}/${totalPages}`, pageW/2, pageH-5, { align:'center' });
    doc.setDrawColor(200,210,220);
    doc.line(margin, pageH-8, pageW-margin, pageH-8);
  }

  doc.save(`Dashboard_CAD_${scope}_${new Date().toISOString().slice(0,10)}.pdf`);
}

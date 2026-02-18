// ==================== VARIABLES GLOBALES ====================
let prevRows = [], realRows = [];
let chart, chartBar, chartBarTache, chartPieBudgetRev, chartPieAT;
let chartHonoMetier, chartSyntheseMetier, chartGantt;
// Stockage des dates manuelles du planning: { id: { realStart: Date|null, realEnd: Date|null } }
let manualPlanningDates = {};

const prevInput   = document.getElementById('prevFile');
const realInput   = document.getElementById('realFile');
const kSlider     = document.getElementById('coefK');
const kNumber     = document.getElementById('coefK_num');
const statusEl    = document.getElementById('status');
const chartEl     = document.getElementById('chart');
const chartBarEl  = document.getElementById('chartBar');
const chartBarTacheEl = document.getElementById('chartBarTache');
const chartPieBudgetRevEl = document.getElementById('chartPieBudgetRev');
const chartPieATEl = document.getElementById('chartPieAT');
const tablePrev   = document.getElementById('tablePreview');
const tableCmp    = document.getElementById('tableCompare');
const viewGlobalEl = document.getElementById('viewGlobal');
const viewSOnlyEl  = document.getElementById('viewSOnly');
const btnExportPNG   = document.getElementById('btnExportPNG');
const btnExportTablesPNG = document.getElementById('btnExportTablesPNG');
const btnExportXLSX  = document.getElementById('btnExportXLSX');
const btnRefresh     = document.getElementById('btnRefresh');
const kpiSection     = document.getElementById('kpiSection');

const chartHonoMetierEl    = document.getElementById('chartHonoMetier');
const tableRessources      = document.getElementById('tableRessources');
const chartSyntheseMetierEl = document.getElementById('chartSyntheseMetier');
const kpiSectionHonoraires  = document.getElementById('kpiSectionHonoraires');

const chartGanttEl    = document.getElementById('chartGantt');
const tablePlanning   = document.getElementById('tablePlanning');
const btnExportGanttPNG = document.getElementById('btnExportGanttPNG');
const btnApplyPlanning  = document.getElementById('btnApplyPlanning');

const PROJECT_NAMES = {
  'CL150096_11_05_01': 'Général - Direction',
  'CL150096_11_05_02': 'Centrale',
  'CL150096_11_05_03': 'Bâtiment',
  'CL150096_11_05_04': 'Réseau CAD',
  'CL150096_11_05_05': 'Sous-stations',
  'CL150096_11_05_06': 'Participations clients'
};

const PROJECT_ORDER = [
  'CL150096_11_05_01',
  'CL150096_11_05_02',
  'CL150096_11_05_03',
  'CL150096_11_05_04',
  'CL150096_11_05_05',
  'CL150096_11_05_06'
];

function getProjectName(id) {
  return PROJECT_NAMES[id] || id;
}

// ==================== UTILITAIRES ====================
function readExcel(file, cb) {
  const r = new FileReader();
  r.onload = e => {
    const wb   = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows  = XLSX.utils.sheet_to_json(sheet, { raw: true });
    cb(rows);
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
    // Format ISO-like avec heure: "2026-05-25 00:00:00"
    const trimmed = v.trim().replace(' 00:00:00', '').replace('T00:00:00', '');
    let d = new Date(trimmed);
    if (!isNaN(d)) return d;
    const m = trimmed.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
    if (m) {
      const dd = +m[1], MM = (+m[2]) - 1, yy = +m[3];
      return new Date(yy < 100 ? 2000 + yy : yy, MM, dd);
    }
  }
  return null;
}

function firstOfMonth(d) {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function yyyymm(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function buildAxis(minDate, maxDate) {
  const axis = [];
  const c = firstOfMonth(minDate);
  const e = firstOfMonth(maxDate);
  while (c <= e) {
    axis.push(yyyymm(c));
    c.setMonth(c.getMonth() + 1);
  }
  return axis;
}

/**
 * Calcule la courbe S cumulative pour un total donné sur un axe temporel.
 * Chaque type de budget utilise son propre range de dates (sStart, sEnd).
 */
function sCurveCumulative(total, axis, k, range) {
  const N = axis.length;
  const cum   = Array(N).fill(null);
  const month = Array(N).fill(0);
  const s0 = 1 / (1 + Math.exp(-k * (0 - 0.5)));
  const s1 = 1 / (1 + Math.exp(-k * (1 - 0.5)));
  const [s, e] = range;
  const L = e - s + 1;
  if (L <= 0 || total === 0) return { cum, month };
  for (let j = 0; j < L; j++) {
    const i = s + j;
    if (i >= N) break;
    const x  = j / (L - 1 || 1);
    const sx = 1 / (1 + Math.exp(-k * (x - 0.5)));
    const norm = (sx - s0) / (s1 - s0);
    cum[i]   = norm * total;
    month[i] = (j === 0) ? cum[i] : (cum[i] - cum[i - 1]);
  }
  return { cum, month };
}

/**
 * Calcule le range (indices min/max) d'un type de budget sur l'axe global.
 */
function computeRange(prevRows, mode, axis) {
  const rows  = prevRows.filter(r => r['Modèle de prévision'] === mode);
  const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
  if (!dates.length) return [0, axis.length - 1];
  const minD = new Date(Math.min(...dates));
  const maxD = new Date(Math.max(...dates));
  let sStart = axis.findIndex(m => {
    const [yy, mm] = m.split('-');
    return new Date(parseInt(yy), parseInt(mm) - 1, 1) >= minD;
  });
  let sEnd = axis.findIndex(m => {
    const [yy, mm] = m.split('-');
    return new Date(parseInt(yy), parseInt(mm) - 1, 1) > maxD;
  });
  if (sStart === -1) sStart = 0;
  if (sEnd === -1)   sEnd   = axis.length - 1; else sEnd = sEnd - 1;
  return [sStart, sEnd];
}

// Plugin ligne "Aujourd'hui"
const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(chart, args, opts) {
    const { ctx, chartArea: { top, bottom }, scales: { x } } = chart;
    if (!opts || !opts.label) return;
    const labels = chart.data.labels;
    const toTime = s => { const [yy, mm] = s.split('-'); return new Date(parseInt(yy), parseInt(mm) - 1, 1).getTime(); };
    const t = toTime(opts.label);
    let bestI = 0, bestD = Infinity;
    labels.forEach((l, i) => { const d = Math.abs(toTime(l) - t); if (d < bestD) { bestD = d; bestI = i; } });
    const xPos = x.getPixelForValue(bestI);
    ctx.save();
    ctx.strokeStyle = opts.color || 'red';
    ctx.lineWidth   = opts.width || 2;
    ctx.setLineDash([5, 5]);
    ctx.beginPath(); ctx.moveTo(xPos, top); ctx.lineTo(xPos, bottom); ctx.stroke();
    ctx.setLineDash([]);
    if (opts.text) {
      ctx.fillStyle  = opts.color || 'red';
      ctx.font       = 'bold 11px Arial';
      ctx.textAlign  = 'center';
      ctx.fillText(opts.text, xPos, top - 5);
    }
    ctx.restore();
  }
};

// ==================== EVENTS ====================
prevInput.addEventListener('change', e => {
  if (e.target.files[0]) readExcel(e.target.files[0], rows => { prevRows = rows; maybeProcess(); });
});
realInput.addEventListener('change', e => {
  if (e.target.files[0]) readExcel(e.target.files[0], rows => { realRows = rows; maybeProcess(); });
});
[kSlider, kNumber, viewGlobalEl, viewSOnlyEl].forEach(el => {
  el.addEventListener('input', () => {
    kSlider.value = kNumber.value = el.value || kNumber.value;
    if (prevRows.length && realRows.length) processAll();
  });
});
btnRefresh.addEventListener('click', () => { if (prevRows.length && realRows.length) processAll(); });
btnApplyPlanning.addEventListener('click', () => {
  if (prevRows.length && realRows.length) {
    // Lire les dates manuelles depuis les inputs du tableau
    PROJECT_ORDER.forEach(id => {
      const startInput = document.getElementById(`realStart_${id}`);
      const endInput   = document.getElementById(`realEnd_${id}`);
      if (!manualPlanningDates[id]) manualPlanningDates[id] = { realStart: null, realEnd: null };
      manualPlanningDates[id].realStart = startInput && startInput.value ? new Date(startInput.value) : null;
      manualPlanningDates[id].realEnd   = endInput   && endInput.value   ? new Date(endInput.value)   : null;
    });
    processPlanning();
  }
});

function maybeProcess() {
  if (prevRows.length && realRows.length) processAll();
}

function processAll() {
  processBudgets();
  processImputations();
  processPlanning();
}

// ==================== PARTIE 1: BUDGETS ====================
function processBudgets() {
  statusEl.textContent = '⏳ Calcul en cours...';
  statusEl.className   = 'status';

  const k        = parseFloat(kNumber.value);
  const viewMode = viewGlobalEl.checked ? 'global' : 's-only';

  // Totaux par type
  const totals = {};
  ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
    totals[mode] = prevRows
      .filter(r => r['Modèle de prévision'] === mode)
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  });

  // Réels
  const realDates = realRows.map(r => toDate(r['Date du projet'])).filter(Boolean);

  // Axe global = union de toutes les dates (Budget, Budget_Rev, AT, Réels)
  const allDates = [];
  ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
    prevRows.filter(r => r['Modèle de prévision'] === mode)
      .forEach(r => { [toDate(r['Date de début']), toDate(r['Date de fin'])].filter(Boolean).forEach(d => allDates.push(d)); });
  });
  realDates.forEach(d => allDates.push(d));

  if (!allDates.length) { statusEl.textContent = '⚠️ Aucune date trouvée dans les fichiers.'; return; }

  const axisMinD = new Date(Math.min(...allDates));
  const axisMaxD = new Date(Math.max(...allDates));
  let axis = buildAxis(axisMinD, axisMaxD);

  // ——— CORRECTION BUG COURBE S ———
  // Chaque type a son propre range de dates sur l'axe global
  const series = {};
  ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
    const range         = computeRange(prevRows, mode, axis);
    const { cum, month } = sCurveCumulative(totals[mode], axis, k, range);
    series[mode] = { total: totals[mode], cum, month, range };
  });

  // Réels par mois
  const realMonth = Array(axis.length).fill(0);
  realRows.forEach(r => {
    const d = toDate(r['Date du projet']);
    if (!d) return;
    const m   = yyyymm(d);
    const idx = axis.indexOf(m);
    if (idx !== -1) realMonth[idx] += toNum(r['Montant total du coût']);
  });
  const realCum = [];
  let cumR = 0;
  realMonth.forEach(v => { cumR += v; realCum.push(cumR); });

  // Vue S-only: limiter à la plage Budget_Rev
  let displayAxis    = axis;
  let realCumDisplay = realCum;
  let realMonthDisplay = realMonth;
  const sRange = series['Budget_Rev'].range;

  if (viewMode === 's-only') {
    displayAxis      = axis.slice(sRange[0], sRange[1] + 1);
    realCumDisplay   = realCum.slice(sRange[0], sRange[1] + 1);
    realMonthDisplay = realMonth.slice(sRange[0], sRange[1] + 1);
    ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
      series[mode].cumDisplay   = series[mode].cum.slice(sRange[0], sRange[1] + 1);
      series[mode].monthDisplay = series[mode].month.slice(sRange[0], sRange[1] + 1);
    });
  } else {
    ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
      series[mode].cumDisplay   = series[mode].cum;
      series[mode].monthDisplay = series[mode].month;
    });
  }

  const today = new Date();

  drawCourbesS(displayAxis, series, realCumDisplay, today);
  drawChartBarProjet();
  drawChartBarTache();
  drawPieCharts();
  buildTablePreview(displayAxis, series['Budget_Rev'], realMonthDisplay, viewMode === 's-only' && sRange[0] > 0 ? realCum[sRange[0] - 1] : 0);
  buildTableCompare();

  // KPI
  const totalBR   = totals['Budget_Rev'];
  const totalAT   = totals['AT'];
  const totalR    = realCum[realCum.length - 1] || 0;
  const ecart     = totalR - totalAT;
  const pctAvance = totalAT > 0 ? ((totalR / totalAT) * 100).toFixed(1) : 0;

  kpiSection.style.display = 'grid';
  document.getElementById('kpiBudgetRev').textContent   = toCHF(totalBR);
  document.getElementById('kpiAT').textContent           = toCHF(totalAT);
  document.getElementById('kpiReelTotal').textContent   = toCHF(totalR);
  const kpiEcartEl = document.getElementById('kpiEcart');
  kpiEcartEl.textContent  = toCHF(ecart);
  kpiEcartEl.className    = 'kpi-value ' + (ecart <= 0 ? 'positive' : 'negative');
  document.getElementById('kpiAvancement').textContent  = pctAvance;

  statusEl.textContent = '✅ Données calculées avec succès';
  statusEl.className   = 'status success';
}

function drawCourbesS(axis, series, realCum, today) {
  const data = {
    labels: axis,
    datasets: [
      { label: 'Budget (S cumul)',          data: series.Budget.cumDisplay,    borderColor: '#0ea5e9', backgroundColor: 'rgba(14,165,233,0.1)',  pointRadius: 0, tension: 0.3, spanGaps: true },
      { label: 'Budget Révisé (S cumul)',   data: series.Budget_Rev.cumDisplay, borderColor: '#f59e0b', backgroundColor: 'rgba(245,158,11,0.1)', pointRadius: 0, tension: 0.3, spanGaps: true, borderWidth: 3 },
      { label: 'AT (S cumul)',              data: series.AT.cumDisplay,         borderColor: '#6366f1', backgroundColor: 'rgba(99,102,241,0.1)', pointRadius: 0, tension: 0.3, spanGaps: true },
      { label: 'Réels cumulés',             data: realCum,                      borderColor: '#ef4444', backgroundColor: 'rgba(239,68,68,0.1)',  pointRadius: 2, pointHoverRadius: 5, tension: 0.3, spanGaps: true, borderWidth: 3 }
    ]
  };
  const options = {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top', labels: { padding: 15, font: { size: 13, weight: '600' }, usePointStyle: true } },
      tooltip: { mode: 'index', intersect: false, backgroundColor: 'rgba(26,31,54,0.95)', callbacks: { label: ctx => ctx.parsed.y !== null ? `${ctx.dataset.label}: ${toCHF(ctx.parsed.y)} CHF` : '' } },
      todayLine: { label: yyyymm(today), text: "Aujourd'hui", color: '#ef4444', width: 2 }
    },
    scales: {
      x: { title: { display: true, text: 'Mois' }, grid: { color: '#f3f4f6' }, ticks: { callback: (val, idx) => { const lab = axis[idx]; return (lab && (lab.endsWith('-01') || lab.endsWith('-04') || lab.endsWith('-07') || lab.endsWith('-10'))) ? lab : ''; }, font: { size: 11 } } },
      y: { title: { display: true, text: 'CHF (cumulé)' }, beginAtZero: true, grid: { color: '#f3f4f6' }, ticks: { callback: value => toCHF(value), font: { size: 11 } } }
    },
    interaction: { mode: 'index', intersect: false }
  };
  if (chart) chart.destroy();
  chart = new Chart(chartEl.getContext('2d'), { type: 'line', data, options, plugins: [todayLinePlugin] });
}

function buildTablePreview(axis, brSeries, realMonth, initialRealCum) {
  tablePrev.innerHTML = `<thead><tr><th>Mois</th><th>Montant BR (S)</th><th>Cumul BR (S)</th><th>Réel par mois</th><th>Réel cumulé</th><th>Écart</th></tr></thead><tbody></tbody>`;
  const tbody = tablePrev.querySelector('tbody');
  let cumBR = 0, realCumRunning = initialRealCum;
  axis.forEach((m, i) => {
    const val = brSeries.monthDisplay[i] || 0;
    cumBR += val;
    const rM = realMonth[i] || 0;
    realCumRunning += rM;
    if (val === 0 && cumBR === 0 && rM === 0 && realCumRunning === initialRealCum) return;
    const ecart = realCumRunning - cumBR;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${m}</td><td>${toCHF(val)}</td><td>${toCHF(cumBR)}</td><td>${toCHF(rM)}</td><td>${toCHF(realCumRunning)}</td><td class="${ecart <= 0 ? 'positive' : 'negative'}">${toCHF(ecart)}</td>`;
    tbody.appendChild(tr);
  });
}

function buildTableCompare() {
  tableCmp.innerHTML = `<thead><tr><th>Projet</th><th>Budget</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart Rev</th><th>Écart AT</th><th>% / AT</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb = tableCmp.querySelector('tbody');
  const tf = tableCmp.querySelector('tfoot');
  let totalB = 0, totalBR = 0, totalAT = 0, totalR = 0;
  PROJECT_ORDER.forEach(id => {
    const pred  = prevRows.filter(r => r['ID projet'] === id);
    const sumBy = mode => pred.filter(r => r['Modèle de prévision'] === mode).reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const b     = sumBy('Budget'), brv = sumBy('Budget_Rev'), at = sumBy('AT');
    const realSum = realRows.filter(r => r['ID Projet'] === id).reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    totalB += b; totalBR += brv; totalAT += at; totalR += realSum;
    const ecartRev = realSum - brv, ecartAT = realSum - at;
    const pctAT    = at > 0 ? ((realSum / at) * 100).toFixed(1) : '—';
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><strong>${getProjectName(id)}</strong></td><td>${toCHF(b)}</td><td>${toCHF(brv)}</td><td>${toCHF(at)}</td><td>${toCHF(realSum)}</td><td class="${ecartRev <= 0 ? 'positive' : 'negative'}">${toCHF(ecartRev)}</td><td class="${ecartAT <= 0 ? 'positive' : 'negative'}">${toCHF(ecartAT)}</td><td>${pctAT}${pctAT !== '—' ? '%' : ''}</td>`;
    tb.appendChild(tr);
  });
  const pctATTotal = totalAT > 0 ? ((totalR / totalAT) * 100).toFixed(1) : '—';
  tf.innerHTML = `<tr><td><strong>TOTAL</strong></td><td>${toCHF(totalB)}</td><td>${toCHF(totalBR)}</td><td>${toCHF(totalAT)}</td><td>${toCHF(totalR)}</td><td class="${totalR - totalBR <= 0 ? 'positive' : 'negative'}">${toCHF(totalR - totalBR)}</td><td class="${totalR - totalAT <= 0 ? 'positive' : 'negative'}">${toCHF(totalR - totalAT)}</td><td><strong>${pctATTotal}${pctATTotal !== '—' ? '%' : ''}</strong></td></tr>`;
}

function drawChartBarProjet() {
  const labels = [], dataBR = [], dataAT = [], dataReal = [];
  PROJECT_ORDER.forEach(id => {
    const brv     = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const at      = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const realSum = realRows.filter(r => r['ID Projet'] === id).reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    labels.push(getProjectName(id)); dataBR.push(brv); dataAT.push(at); dataReal.push(realSum);
  });
  if (chartBar) chartBar.destroy();
  chartBar = new Chart(chartBarEl.getContext('2d'), {
    type: 'bar',
    data: { labels, datasets: [
      { label: 'Budget Révisé', data: dataBR,  backgroundColor: '#f59e0b' },
      { label: 'AT',            data: dataAT,  backgroundColor: '#6366f1' },
      { label: 'Réel',          data: dataReal, backgroundColor: '#ef4444' }
    ]},
    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF` } } }, scales: { x: { ticks: { callback: value => toCHF(value) } } } }
  });
}

/**
 * ——— NOUVEAU GRAPHE: Budget Révisé / AT / Réel par Nom de la Tâche ———
 */
function drawChartBarTache() {
  // Regrouper Budget_Rev et AT par Nom de la tâche depuis prevRows
  const taskBR   = {};
  const taskAT   = {};
  const taskReal = {};

  prevRows.forEach(r => {
    const task   = r['Nom de la tâche'] || 'Non défini';
    const montant = toNum(r['Montant total du coût']);
    if (r['Modèle de prévision'] === 'Budget_Rev') taskBR[task]   = (taskBR[task]   || 0) + montant;
    if (r['Modèle de prévision'] === 'AT')         taskAT[task]   = (taskAT[task]   || 0) + montant;
  });

  realRows.forEach(r => {
    const task    = r['Nom de la tâche'] || 'Non défini';
    const montant = toNum(r['Montant total du coût']);
    taskReal[task] = (taskReal[task] || 0) + montant;
  });

  // Union de toutes les tâches connues
  const allTasks = [...new Set([...Object.keys(taskBR), ...Object.keys(taskAT), ...Object.keys(taskReal)])];

  // Raccourcir les labels (supprimer le suffixe _SIA 5)
  const shortLabel = t => t.replace(/_SIA 5$/, '').replace(/_SIA5$/, '');

  if (chartBarTache) chartBarTache.destroy();
  chartBarTache = new Chart(chartBarTacheEl.getContext('2d'), {
    type: 'bar',
    data: {
      labels: allTasks.map(shortLabel),
      datasets: [
        { label: 'Budget Révisé', data: allTasks.map(t => taskBR[t]   || 0), backgroundColor: '#f59e0b' },
        { label: 'AT',            data: allTasks.map(t => taskAT[t]   || 0), backgroundColor: '#6366f1' },
        { label: 'Réel',          data: allTasks.map(t => taskReal[t] || 0), backgroundColor: '#ef4444' }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' },
        tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF` } }
      },
      scales: { x: { ticks: { callback: value => toCHF(value) } } }
    }
  });
}

function drawPieCharts() {
  const labelsBR = [], dataBR = [], labelsAT = [], dataAT = [];
  const colors   = ['#0ea5e9','#f59e0b','#6366f1','#10b981','#ef4444','#8b5cf6'];
  PROJECT_ORDER.forEach(id => {
    const brv = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const at  = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    if (brv > 0) { labelsBR.push(getProjectName(id)); dataBR.push(brv); }
    if (at  > 0) { labelsAT.push(getProjectName(id)); dataAT.push(at); }
  });
  const pieCfg = (labels, data) => ({
    type: 'pie',
    data: { labels, datasets: [{ data, backgroundColor: colors }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right' }, tooltip: { callbacks: { label: ctx => `${ctx.label}: ${toCHF(ctx.parsed)} CHF (${((ctx.parsed / data.reduce((a,b)=>a+b,0))*100).toFixed(1)}%)` } } } }
  });
  if (chartPieBudgetRev) chartPieBudgetRev.destroy();
  chartPieBudgetRev = new Chart(chartPieBudgetRevEl.getContext('2d'), pieCfg(labelsBR, dataBR));
  if (chartPieAT) chartPieAT.destroy();
  chartPieAT = new Chart(chartPieATEl.getContext('2d'), pieCfg(labelsAT, dataAT));
}

// ==================== PARTIE 2: IMPUTATIONS ====================
/**
 * Extrait le type d'honoraires depuis le Nom de la Tâche.
 * Ex: "Honoraires conducteur travaux interne_Réseau CAD_SIA 5" → "Honoraires conducteur travaux interne"
 * Ex: "Honoraires responsable projets d'affaires_SIA 5" → "Honoraires responsable projets d'affaires"
 */
function extractTypeHonoraires(nomTache) {
  if (!nomTache || typeof nomTache !== 'string') return nomTache || 'Autre';
  // Supprimer le suffixe à partir du premier underscore suivi de "Réseau", "Sous", "Général", "Centrale", "Bâtiment", "SIA"
  const cleaned = nomTache.replace(/_(?:Réseau CAD|Sous-stations|Général|Centrale|Bâtiment|SIA).*$/, '').trim();
  return cleaned || nomTache;
}

function processImputations() {
  // ——— KPI Honoraires ———
  const honoPrevBR = prevRows.filter(r => r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires') && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const honoPrevAT = prevRows.filter(r => r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires') && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const honoReel   = realRows.filter(r => r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires')).reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const honoReste  = honoPrevAT - honoReel;

  kpiSectionHonoraires.style.display = 'grid';
  document.getElementById('kpiHonoBudgetRev').textContent = toCHF(honoPrevBR);
  document.getElementById('kpiHonoAT').textContent        = toCHF(honoPrevAT);
  document.getElementById('kpiHonoReel').textContent      = toCHF(honoReel);
  const kpiHonoResteEl = document.getElementById('kpiHonoReste');
  kpiHonoResteEl.textContent = toCHF(honoReste);
  kpiHonoResteEl.className   = 'kpi-value ' + (honoReste >= 0 ? 'positive' : 'negative');

  // ——— Graphique Honoraires par Type ———
  const typesBR   = {};
  const typesAT   = {};
  const typesReel = {};

  prevRows.forEach(r => {
    if (!r['Nom de la tâche'] || !r['Nom de la tâche'].includes('Honoraires')) return;
    const type    = extractTypeHonoraires(r['Nom de la tâche']);
    const montant = toNum(r['Montant total du coût']);
    if (r['Modèle de prévision'] === 'Budget_Rev') typesBR[type]   = (typesBR[type]   || 0) + montant;
    if (r['Modèle de prévision'] === 'AT')         typesAT[type]   = (typesAT[type]   || 0) + montant;
  });
  realRows.forEach(r => {
    if (!r['Nom de la tâche'] || !r['Nom de la tâche'].includes('Honoraires')) return;
    const type    = extractTypeHonoraires(r['Nom de la tâche']);
    const montant = toNum(r['Montant total du coût']);
    typesReel[type] = (typesReel[type] || 0) + montant;
  });

  const allTypes = [...new Set([...Object.keys(typesBR), ...Object.keys(typesAT), ...Object.keys(typesReel)])];

  if (chartHonoMetier) chartHonoMetier.destroy();
  chartHonoMetier = new Chart(chartHonoMetierEl.getContext('2d'), {
    type: 'bar',
    data: { labels: allTypes, datasets: [
      { label: 'Budget Révisé', data: allTypes.map(t => typesBR[t]   || 0), backgroundColor: '#f59e0b' },
      { label: 'AT',            data: allTypes.map(t => typesAT[t]   || 0), backgroundColor: '#6366f1' },
      { label: 'Réel',          data: allTypes.map(t => typesReel[t] || 0), backgroundColor: '#ef4444' }
    ]},
    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF` } } }, scales: { x: { ticks: { callback: v => toCHF(v) } } } }
  });

  // ——— TABLEAU RESSOURCES CORRIGÉ ———
  // Grouper par (Ressource.Nom de la ressource + Nom de la tâche)
  // Le Type d'honoraires = extractTypeHonoraires(Nom de la tâche)
  const ressMap = {};
  realRows.forEach(r => {
    const nomTache   = r['Nom de la tâche'] || '';
    const ressource  = r['Ressource.Nom de la ressource'] || (r['Nom du Fournisseur'] ? `Fourn.: ${r['Nom du Fournisseur']}` : 'Non spécifié');
    const typeHono   = extractTypeHonoraires(nomTache);
    const montant    = toNum(r['Montant total du coût']);
    const projet     = getProjectName(r['ID Projet'] || '');
    const key        = `${ressource}||${nomTache}||${projet}`;
    if (!ressMap[key]) ressMap[key] = { ressource, typeHonoraires: typeHono, nomTache, projet, montant: 0 };
    ressMap[key].montant += montant;
  });

  const ressourcesList = Object.values(ressMap)
    .filter(r => r.montant !== 0)
    .sort((a, b) => b.montant - a.montant);

  tableRessources.innerHTML = `<thead><tr><th>Ressource</th><th>Type d'Honoraires</th><th>Projet</th><th>Montant Réel (CHF)</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tbody = tableRessources.querySelector('tbody');
  const tfoot = tableRessources.querySelector('tfoot');
  let totalRessources = 0;
  ressourcesList.forEach(r => {
    totalRessources += r.montant;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.ressource}</td><td>${r.typeHonoraires}</td><td>${r.projet}</td><td class="${r.montant < 0 ? 'negative' : ''}">${toCHF(r.montant)}</td>`;
    tbody.appendChild(tr);
  });
  tfoot.innerHTML = `<tr><td colspan="3"><strong>TOTAL</strong></td><td><strong>${toCHF(totalRessources)}</strong></td></tr>`;

  // ——— Graphique Synthèse par Type (Réel) ———
  if (chartSyntheseMetier) chartSyntheseMetier.destroy();
  chartSyntheseMetier = new Chart(chartSyntheseMetierEl.getContext('2d'), {
    type: 'bar',
    data: { labels: allTypes, datasets: [{ label: 'Réel', data: allTypes.map(t => typesReel[t] || 0), backgroundColor: '#ef4444' }] },
    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `Réel: ${toCHF(ctx.parsed.x)} CHF` } } }, scales: { x: { ticks: { callback: v => toCHF(v) } } } }
  });
}

// ==================== PARTIE 3: PLANNING ====================
function processPlanning() {
  const planning = {};

  PROJECT_ORDER.forEach(id => {
    const getRange = mode => {
      const rows  = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === mode);
      const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
      return dates.length ? { start: new Date(Math.min(...dates)), end: new Date(Math.max(...dates)) } : { start: null, end: null };
    };

    const br = getRange('Budget_Rev');
    const at = getRange('AT');

    // Dates réelles: priorité aux dates manuelles si saisies
    const manual = manualPlanningDates[id] || {};
    let realStart = manual.realStart || null;
    let realEnd   = manual.realEnd   || null;

    if (!realStart || !realEnd) {
      // Calculer depuis les imputations
      const rDates = realRows.filter(r => r['ID Projet'] === id).map(r => toDate(r['Date du projet'])).filter(Boolean);
      if (rDates.length) {
        if (!realStart) realStart = new Date(Math.min(...rDates));
        if (!realEnd)   realEnd   = new Date(Math.max(...rDates));
      }
    }

    planning[id] = { name: getProjectName(id), brStart: br.start, brEnd: br.end, atStart: at.start, atEnd: at.end, realStart, realEnd };
  });

  buildTablePlanning(planning);
  drawGantt(planning);
}

function buildTablePlanning(planning) {
  tablePlanning.innerHTML = `
    <thead><tr>
      <th>Projet</th><th>Début BR</th><th>Fin BR</th>
      <th>Début AT</th><th>Fin AT</th>
      <th>Réel Début</th><th>Réel Fin</th>
      <th>Écart (j)</th><th>Statut</th>
    </tr></thead><tbody></tbody>`;

  const tbody  = tablePlanning.querySelector('tbody');
  const today  = new Date();
  const fmtDate = d => d instanceof Date && !isNaN(d) ? d.toISOString().split('T')[0] : '';
  const fmtDisp = d => d instanceof Date && !isNaN(d) ? d.toLocaleDateString('fr-CH') : '—';

  PROJECT_ORDER.forEach(id => {
    const p = planning[id];
    const ecartJours = p.atEnd && p.realEnd ? Math.round((p.realEnd - p.atEnd) / 86400000) : null;
    let statut = 'prevu', statutLabel = 'Prévu';
    if (p.realStart && !p.realEnd)         { statut = 'en-cours'; statutLabel = 'En cours'; }
    else if (p.realEnd) {
      if (ecartJours > 0) { statut = 'retard';  statutLabel = 'Retard'; }
      else               { statut = 'termine'; statutLabel = 'Terminé'; }
    }
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${p.name}</strong></td>
      <td>${fmtDisp(p.brStart)}</td><td>${fmtDisp(p.brEnd)}</td>
      <td>${fmtDisp(p.atStart)}</td><td>${fmtDisp(p.atEnd)}</td>
      <td><input type="date" id="realStart_${id}" value="${fmtDate(p.realStart)}" class="date-input"/></td>
      <td><input type="date" id="realEnd_${id}"   value="${fmtDate(p.realEnd)}"   class="date-input"/></td>
      <td>${ecartJours !== null ? ecartJours : '—'}</td>
      <td><span class="badge ${statut}">${statutLabel}</span></td>`;
    tbody.appendChild(tr);
  });
}

function drawGantt(planning) {
  const allDates = [];
  PROJECT_ORDER.forEach(id => {
    const p = planning[id];
    [p.brStart, p.brEnd, p.atStart, p.atEnd, p.realStart, p.realEnd].forEach(d => { if (d && !isNaN(d)) allDates.push(d); });
  });
  if (!allDates.length) return;

  const minDate = new Date(Math.min(...allDates));
  const toDay   = d => (d && !isNaN(d)) ? Math.round((d - minDate) / 86400000) : null;
  const datasets = [];

  PROJECT_ORDER.forEach((id, idx) => {
    const p = planning[id];
    const pushBar = (label, start, end, color) => {
      if (start && end && !isNaN(start) && !isNaN(end)) {
        datasets.push({ label, data: [{ x: [toDay(start), toDay(end)], y: idx * 3 + datasets.length % 3 }], backgroundColor: color });
      }
    };
    if (p.brStart && p.brEnd)     datasets.push({ label: `${p.name} — BR`,   data: [{ x: [toDay(p.brStart), toDay(p.brEnd)],     y: idx * 3     }], backgroundColor: '#f59e0b' });
    if (p.atStart && p.atEnd)     datasets.push({ label: `${p.name} — AT`,   data: [{ x: [toDay(p.atStart), toDay(p.atEnd)],     y: idx * 3 + 1 }], backgroundColor: '#6366f1' });
    if (p.realStart && p.realEnd) datasets.push({ label: `${p.name} — Réel`, data: [{ x: [toDay(p.realStart), toDay(p.realEnd)], y: idx * 3 + 2 }], backgroundColor: '#ef4444' });
  });

  if (chartGantt) chartGantt.destroy();
  chartGantt = new Chart(chartGanttEl.getContext('2d'), {
    type: 'bar',
    data: { datasets },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => {
          const [d0, d1] = ctx.parsed.x;
          const start = new Date(minDate.getTime() + d0 * 86400000);
          const end   = new Date(minDate.getTime() + d1 * 86400000);
          return `${ctx.dataset.label}: ${start.toLocaleDateString('fr-CH')} – ${end.toLocaleDateString('fr-CH')}`;
        }}}
      },
      scales: {
        x: { type: 'linear', ticks: { callback: value => { const d = new Date(minDate.getTime() + value * 86400000); return d.toLocaleDateString('fr-CH', { month: 'short', year: 'numeric' }); } } },
        y: { ticks: { callback: (value) => { const projIdx = Math.floor(value / 3); return PROJECT_ORDER[projIdx] ? getProjectName(PROJECT_ORDER[projIdx]) : ''; } } }
      }
    }
  });
}

// ==================== EXPORTS ====================
btnExportPNG.addEventListener('click', () => { if (!chart) return; const a = document.createElement('a'); a.download = 'courbes_cad.png'; a.href = chart.toBase64Image('image/png', 1); a.click(); });
btnExportTablesPNG.addEventListener('click', () => {
  [document.getElementById('tablePrevCard'), document.getElementById('tableCmpCard')].forEach((el, i) => {
    html2canvas(el).then(canvas => { const a = document.createElement('a'); a.download = `tableau_${i+1}.png`; a.href = canvas.toDataURL('image/png'); a.click(); });
  });
});
btnExportXLSX.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(tablePrev), 'Détails Mensuels');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(tableCmp),  'Comparatif Projets');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(tableRessources), 'Ressources');
  XLSX.writeFile(wb, 'tableaux_cad.xlsx');
});
btnExportGanttPNG.addEventListener('click', () => { if (!chartGantt) return; const a = document.createElement('a'); a.download = 'planning_gantt.png'; a.href = chartGantt.toBase64Image('image/png', 1); a.click(); });
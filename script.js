/* ============================================================
   Dashboard CAD — Groupe E Celsius
   script.js  v5.0  — 18/02/2026
   Multi-Phase: 5 onglets (Récap, SIA 2/3/4/5)
   ============================================================ */

// ---- Globals ----
let prevRows = [], realRows = [];
const charts = {}; // stockage centralisé de tous les charts

const prevInput = document.getElementById('prevFile');
const realInput = document.getElementById('realFile');
const kSlider = document.getElementById('coefK');
const kNumber = document.getElementById('coefK_num');
const statusEl = document.getElementById('status');
const viewGlobalEl = document.getElementById('viewGlobal');
const viewSOnlyEl = document.getElementById('viewSOnly');

// Palette Groupe E
const GE_BLUE = '#163a5f', GE_BLUE2 = '#1e4d8c', GE_BLUE3 = '#2e6db4', GE_BLUE4 = '#5b9bd5';
const GE_ORANGE = '#e8873a', GE_RED = '#c0392b', GE_GREEN = '#27ae60';
const PIE_COLORS = ['#2e6db4','#e8873a','#27ae60','#8e44ad','#c0392b','#16a085','#f39c12'];

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
  return parseFloat(v.replace(/[\u00A0\u202F']/g, '').replace(/[ ]/g, '').replace(',', '.')) || 0;
}
function toCHF(n) { return (n || 0).toLocaleString('fr-CH', { minimumFractionDigits: 0 }); }
function toDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  if (typeof v === 'string') {
    const cleaned = v.trim().replace(' 00:00:00', '').replace('T00:00:00', '');
    let d = new Date(cleaned);
    if (!isNaN(d)) return d;
    const m = cleaned.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})$/);
    if (m) return new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
  }
  return null;
}
function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function yyyymm(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; }
function buildAxis(minD, maxD) {
  const axis = [], c = firstOfMonth(minD), e = firstOfMonth(maxD);
  while (c <= e) { axis.push(yyyymm(c)); c.setMonth(c.getMonth() + 1); }
  return axis;
}

function sCurve(total, axis, k, sIdx, eIdx) {
  const N = axis.length, cum = Array(N).fill(null), L = eIdx - sIdx + 1;
  if (L <= 0 || total === 0 || sIdx < 0 || eIdx < 0) return cum;
  const s0 = 1 / (1 + Math.exp(-k * (0 - 0.5))), s1 = 1 / (1 + Math.exp(-k * (1 - 0.5)));
  for (let j = 0; j < L; j++) {
    const i = sIdx + j; if (i >= N) break;
    const x = j / (L - 1 || 1), sx = 1 / (1 + Math.exp(-k * (x - 0.5)));
    cum[i] = ((sx - s0) / (s1 - s0)) * total;
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
  if (eIdx < 0) { for (let i = axis.length - 1; i >= 0; i--) { if (axis[i] <= mmMax) { eIdx = i; break; } } }
  if (sIdx < 0) sIdx = 0; if (eIdx < 0) eIdx = axis.length - 1;
  return [sIdx, eIdx];
}

const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(ch, _args, opts) {
    if (!opts || !opts.label) return;
    const { ctx, chartArea: { top, bottom }, scales: { x } } = ch;
    const labels = ch.data.labels;
    const toT = s => { const [y, m] = s.split('-'); return new Date(+y, +m - 1, 1).getTime(); };
    const t = toT(opts.label);
    let bi = 0, bd = Infinity;
    labels.forEach((l, i) => { const d = Math.abs(toT(l) - t); if (d < bd) { bd = d; bi = i; } });
    const xp = x.getPixelForValue(bi);
    ctx.save();
    ctx.strokeStyle = GE_RED; ctx.lineWidth = 2; ctx.setLineDash([5, 5]);
    ctx.beginPath(); ctx.moveTo(xp, top); ctx.lineTo(xp, bottom); ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = GE_RED; ctx.font = 'bold 11px Arial'; ctx.textAlign = 'center';
    ctx.fillText("Aujourd'hui", xp, top - 5);
    ctx.restore();
  }
};

// ============================================================
// EXTRACTION PHASE depuis ID Projet
// ============================================================
function extractPhase(idProjet) {
  if (!idProjet || typeof idProjet !== 'string') return null;
  const parts = idProjet.split('_');
  if (parts.length < 3) return null;
  const phaseCode = parts[2];
  return phaseCode === '02' ? 'SIA2' : phaseCode === '03' ? 'SIA3' : phaseCode === '04' ? 'SIA4' : phaseCode === '05' ? 'SIA5' : null;
}

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

// Gestion onglets
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
});

document.getElementById('btnPDF').addEventListener('click', generatePDF);

function maybeProcess() {
  if (prevRows.length && realRows.length) {
    statusEl.textContent = '✅ Fichiers chargés — calcul en cours...';
    statusEl.className = 'status';
    processAll();
  }
}

function processAll() {
  processRecap();
  processPhase('SIA2', '02');
  processPhase('SIA3', '03');
  processPhase('SIA4', '04');
  processSIA5();
  statusEl.textContent = '✅ Calcul terminé avec succès';
  statusEl.className = 'status success';
}

// ============================================================
// RÉCAP GLOBAL
// ============================================================
function processRecap() {
  const k = parseFloat(kNumber.value) || 12;
  const isSOnly = viewSOnlyEl.checked;
  const MODELS = ['Budget_Rev', 'AT'];

  // Consolidation par phase
  const phaseData = {};
  ['SIA2', 'SIA3', 'SIA4', 'SIA5'].forEach(phase => {
    const code = phase === 'SIA2' ? '02' : phase === 'SIA3' ? '03' : phase === 'SIA4' ? '04' : '05';
    const prevPhase = prevRows.filter(r => {
      const ph = extractPhase(r['ID projet']);
      return ph === phase;
    });
    const realPhase = realRows.filter(r => {
      const ph = extractPhase(r['ID Projet'] || r['ID projet']);
      return ph === phase;
    });
    const br = prevPhase.filter(r => r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const at = prevPhase.filter(r => r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const reel = realPhase.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    phaseData[phase] = { br, at, reel };
  });

  const totalBR = Object.values(phaseData).reduce((a, p) => a + p.br, 0);
  const totalAT = Object.values(phaseData).reduce((a, p) => a + p.at, 0);
  const totalR = Object.values(phaseData).reduce((a, p) => a + p.reel, 0);
  const ecart = totalR - totalAT;
  const pct = totalAT > 0 ? (totalR / totalAT * 100).toFixed(1) : 0;

  document.getElementById('kpiSectionGlobal').style.display = 'grid';
  document.getElementById('kpiSectionGlobal').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Budget Révisé Global</div><div class="kpi-value">${toCHF(totalBR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">AT Global</div><div class="kpi-value">${toCHF(totalAT)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Coûts Réels Global</div><div class="kpi-value">${toCHF(totalR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Écart Réel / AT</div><div class="kpi-value ${ecart <= 0 ? 'positive' : 'negative'}">${toCHF(ecart)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">% Avanc. / AT</div><div class="kpi-value ${pct > 100 ? 'negative' : 'positive'}">${pct}%</div><div class="kpi-unit"></div></div>
  `;

  // Courbes S consolidées
  const model = {};
  MODELS.forEach(mn => {
    const rows = prevRows.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = { total, start: dates.length ? new Date(Math.min(...dates)) : null, end: dates.length ? new Date(Math.max(...dates)) : null };
  });

  const realDates = realRows.map(r => toDate(r['Date du projet'] || r['Date'])).filter(Boolean);
  const allDates = [];
  MODELS.forEach(mn => { if (model[mn].start) allDates.push(model[mn].start); if (model[mn].end) allDates.push(model[mn].end); });
  realDates.forEach(d => allDates.push(d));
  if (!allDates.length) return;

  const globalMin = new Date(Math.min(...allDates)), globalMax = new Date(Math.max(...allDates));
  const modelMin = new Date(Math.min(...MODELS.filter(m => model[m].start).map(m => model[m].start)));
  const modelMax = new Date(Math.max(...MODELS.filter(m => model[m].end).map(m => model[m].end)));
  const axisMin = isSOnly ? modelMin : globalMin;
  const axisMax = isSOnly ? modelMax : globalMax;
  const axis = buildAxis(axisMin, axisMax);

  const series = {};
  MODELS.forEach(mn => { const [si, ei] = modelRange(prevRows, mn, axis); series[mn] = sCurve(model[mn].total, axis, k, si, ei); });

  const today = new Date(), mmToday = yyyymm(today);
  const realMap = new Map(axis.map(m => [m, 0]));
  let initCum = 0;
  realRows.forEach(r => {
    const d = toDate(r['Date du projet'] || r['Date']); if (!d) return;
    const mm = yyyymm(firstOfMonth(d)); if (mm > mmToday) return;
    if (realMap.has(mm)) realMap.set(mm, realMap.get(mm) + toNum(r['Montant total du coût']));
    else if (isSOnly && mm < axis[0]) initCum += toNum(r['Montant total du coût']);
  });

  const realMonth = axis.map(m => realMap.get(m) || 0);
  let acc = initCum; const realCum = realMonth.map(v => { acc += v; return acc; });
  const realCut = realCum.map((v, i) => axis[i] > mmToday ? null : v);

  const ds = [
    { label: 'Budget Révisé (S cumul)', data: series['Budget_Rev'], borderColor: GE_ORANGE, backgroundColor: 'rgba(232,135,58,0.07)', pointRadius: 0, tension: 0.3, spanGaps: true, borderWidth: 3 },
    { label: 'AT (S cumul)', data: series['AT'], borderColor: GE_BLUE2, backgroundColor: 'rgba(30,77,140,0.07)', pointRadius: 0, tension: 0.3, spanGaps: true },
    { label: 'Réels cumulés', data: realCut, borderColor: GE_RED, backgroundColor: 'rgba(192,57,43,0.07)', pointRadius: 2, pointHoverRadius: 5, tension: 0.3, spanGaps: true, borderWidth: 3 }
  ];

  if (charts.chartGlobal) charts.chartGlobal.destroy();
  charts.chartGlobal = new Chart(document.getElementById('chartGlobal').getContext('2d'), {
    type: 'line', data: { labels: axis, datasets: ds },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { padding: 15, font: { size: 13, weight: '600' }, usePointStyle: true } },
        tooltip: { mode: 'index', intersect: false, backgroundColor: 'rgba(22,58,95,0.95)', callbacks: { label: ctx => ctx.parsed.y !== null ? `${ctx.dataset.label}: ${toCHF(ctx.parsed.y)} CHF` : '' } },
        todayLine: { label: yyyymm(today) }
      },
      scales: {
        x: { title: { display: true, text: 'Mois' }, grid: { color: '#eef2f7' }, ticks: { font: { size: 11 }, callback: (_, i) => { const l = axis[i]; return l && (l.endsWith('-01') || l.endsWith('-04') || l.endsWith('-07') || l.endsWith('-10')) ? l : ''; } } },
        y: { title: { display: true, text: 'CHF (cumulé)' }, beginAtZero: true, grid: { color: '#eef2f7' }, ticks: { callback: v => toCHF(v), font: { size: 11 } } }
      },
      interaction: { mode: 'index', intersect: false }
    },
    plugins: [todayLinePlugin]
  });

  // Pie phases
  const labelsPhases = ['SIA 2', 'SIA 3', 'SIA 4', 'SIA 5'];
  const dataPhases = [phaseData.SIA2.br, phaseData.SIA3.br, phaseData.SIA4.br, phaseData.SIA5.br];
  if (charts.chartPiePhases) charts.chartPiePhases.destroy();
  charts.chartPiePhases = new Chart(document.getElementById('chartPiePhases').getContext('2d'), {
    type: 'pie',
    data: { labels: labelsPhases, datasets: [{ data: dataPhases, backgroundColor: PIE_COLORS, borderWidth: 2, borderColor: '#ffffff' }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 12 }, padding: 14, usePointStyle: true, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${toCHF(ctx.parsed)} CHF (${((ctx.parsed / totalBR) * 100).toFixed(1)}%)` } }
      }
    }
  });

  // Tableau consolidé par phase
  const tablePhases = document.getElementById('tablePhases');
  tablePhases.innerHTML = `<thead><tr><th>Phase</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart / BR</th><th>Écart / AT</th><th>% / AT</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb = tablePhases.querySelector('tbody'), tf = tablePhases.querySelector('tfoot');
  ['SIA2', 'SIA3', 'SIA4', 'SIA5'].forEach(phase => {
    const p = phaseData[phase];
    const eR = p.reel - p.br, eA = p.reel - p.at, pct = p.at > 0 ? (p.reel / p.at * 100).toFixed(1) + '%' : '—';
    const phaseName = phase === 'SIA2' ? 'SIA 2 — Études Prélim.' : phase === 'SIA3' ? 'SIA 3 — Étude Projet' : phase === 'SIA4' ? 'SIA 4 — Appel Offres' : 'SIA 5 — Réalisation';
    tb.insertAdjacentHTML('beforeend', `<tr><td><strong>${phaseName}</strong></td><td>${toCHF(p.br)}</td><td>${toCHF(p.at)}</td><td>${toCHF(p.reel)}</td><td class="${eR <= 0 ? 'positive' : 'negative'}">${toCHF(eR)}</td><td class="${eA <= 0 ? 'positive' : 'negative'}">${toCHF(eA)}</td><td>${pct}</td></tr>`);
  });
  const pT = totalAT > 0 ? (totalR / totalAT * 100).toFixed(1) + '%' : '—';
  tf.innerHTML = `<tr><td><strong>TOTAL</strong></td><td>${toCHF(totalBR)}</td><td>${toCHF(totalAT)}</td><td>${toCHF(totalR)}</td><td class="${totalR - totalBR <= 0 ? 'positive' : 'negative'}">${toCHF(totalR - totalBR)}</td><td class="${totalR - totalAT <= 0 ? 'positive' : 'negative'}">${toCHF(totalR - totalAT)}</td><td><strong>${pT}</strong></td></tr>`;

  // Gantt multi-phase
  drawGanttGlobal();
}

function drawGanttGlobal() {
  const phases = ['SIA2', 'SIA3', 'SIA4', 'SIA5'];
  const plan = {};
  phases.forEach(phase => {
    const code = phase === 'SIA2' ? '02' : phase === 'SIA3' ? '03' : phase === 'SIA4' ? '04' : '05';
    const prevPhase = prevRows.filter(r => extractPhase(r['ID projet']) === phase);
    const realPhase = realRows.filter(r => extractPhase(r['ID Projet'] || r['ID projet']) === phase);

    const getR = mode => {
      const rows = prevPhase.filter(r => r['Modèle de prévision'] === mode);
      const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
      return dates.length ? { s: new Date(Math.min(...dates)), e: new Date(Math.max(...dates)) } : { s: null, e: null };
    };
    const br = getR('Budget_Rev'), at = getR('AT');
    const rd = realPhase.map(r => toDate(r['Date du projet'] || r['Date'])).filter(Boolean);
    const rS = rd.length ? new Date(Math.min(...rd)) : null;
    const rE = rd.length ? new Date(Math.max(...rd)) : null;
    const phaseName = phase === 'SIA2' ? 'SIA 2' : phase === 'SIA3' ? 'SIA 3' : phase === 'SIA4' ? 'SIA 4' : 'SIA 5';
    plan[phase] = { name: phaseName, brS: br.s, brE: br.e, atS: at.s, atE: at.e, rS, rE };
  });

  const allD = [];
  Object.values(plan).forEach(p => { [p.brS, p.brE, p.atS, p.atE, p.rS, p.rE].forEach(d => { if (d && !isNaN(d)) allD.push(d); }); });
  if (!allD.length) return;

  const minD = new Date(Math.min(...allD)), maxD = new Date(Math.max(...allD));
  const span = Math.round((maxD - minD) / 86400000) + 60;
  const toDay = d => (d && !isNaN(d)) ? Math.round((d - minD) / 86400000) : null;

  const yLabels = phases.map(ph => plan[ph].name);
  const mkPoint = (s, e, label) => {
    const s0 = toDay(s), e0 = toDay(e);
    if (s0 === null || e0 === null) return { x: [null, null], y: label };
    return { x: [s0, e0], y: label };
  };

  const dsBR = {
    label: 'Budget Révisé', data: phases.map(ph => mkPoint(plan[ph].brS, plan[ph].brE, plan[ph].name)),
    backgroundColor: GE_ORANGE + 'cc', borderColor: GE_ORANGE, borderWidth: 2, borderSkipped: false, borderRadius: 3, barThickness: 18
  };
  const dsAT = {
    label: 'AT', data: phases.map(ph => mkPoint(plan[ph].atS, plan[ph].atE, plan[ph].name)),
    backgroundColor: GE_BLUE2 + 'cc', borderColor: GE_BLUE2, borderWidth: 2, borderSkipped: false, borderRadius: 3, barThickness: 18
  };
  const dsReal = {
    label: 'Réel', data: phases.map(ph => mkPoint(plan[ph].rS, plan[ph].rE, plan[ph].name)),
    backgroundColor: GE_RED + 'cc', borderColor: GE_RED, borderWidth: 2, borderSkipped: false, borderRadius: 3, barThickness: 18
  };

  const tickStep = Math.max(30, Math.round(span / 12));
  if (charts.chartGanttGlobal) charts.chartGanttGlobal.destroy();
  charts.chartGanttGlobal = new Chart(document.getElementById('chartGanttGlobal').getContext('2d'), {
    type: 'bar', data: { labels: yLabels, datasets: [dsBR, dsAT, dsReal] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { usePointStyle: true, padding: 16, font: { size: 12, weight: '600' } } },
        tooltip: {
          callbacks: {
            title: ctx => `${ctx[0].label} — ${ctx[0].dataset.label}`,
            label: ctx => {
              const [d0, d1] = ctx.parsed.x;
              if (d0 === null || d1 === null) return 'Pas de données';
              const sD = new Date(minD.getTime() + d0 * 86400000), eD = new Date(minD.getTime() + d1 * 86400000);
              const dur = Math.round((eD - sD) / 86400000);
              return [`Début : ${sD.toLocaleDateString('fr-CH')}`, `Fin   : ${eD.toLocaleDateString('fr-CH')}`, `Durée : ${dur} jours`];
            }
          }
        }
      },
      scales: {
        x: {
          type: 'linear', min: 0, max: span, title: { display: true, text: 'Timeline' }, grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: {
            stepSize: tickStep,
            callback: value => {
              const d = new Date(minD.getTime() + value * 86400000);
              return d.toLocaleDateString('fr-CH', { month: 'short', year: '2-digit' });
            }
          }
        },
        y: { type: 'category', labels: yLabels, offset: true, grid: { color: 'rgba(0,0,0,0.08)', lineWidth: 1 }, ticks: { font: { weight: 'bold', size: 11 }, color: '#163a5f' } }
      }
    }
  });
}

// ============================================================
// PROCESS PHASE (SIA 2/3/4) — Par Tâche
// ============================================================
function processPhase(phaseName, phaseCode) {
  const k = parseFloat(kNumber.value) || 12;
  const isSOnly = viewSOnlyEl.checked;
  const MODELS = ['Budget_Rev', 'AT'];

  const prevPhase = prevRows.filter(r => extractPhase(r['ID projet']) === phaseName);
  const realPhase = realRows.filter(r => extractPhase(r['ID Projet'] || r['ID projet']) === phaseName);

  // KPI phase
  const br = prevPhase.filter(r => r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const at = prevPhase.filter(r => r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const reel = realPhase.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const ecart = reel - at, pct = at > 0 ? (reel / at * 100).toFixed(1) : 0;

  document.getElementById(`kpi${phaseName}`).style.display = 'grid';
  document.getElementById(`kpi${phaseName}`).innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Budget Révisé</div><div class="kpi-value">${toCHF(br)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">AT</div><div class="kpi-value">${toCHF(at)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Coûts Réels</div><div class="kpi-value">${toCHF(reel)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Écart Réel / AT</div><div class="kpi-value ${ecart <= 0 ? 'positive' : 'negative'}">${toCHF(ecart)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">% Avanc. / AT</div><div class="kpi-value ${pct > 100 ? 'negative' : 'positive'}">${pct}%</div><div class="kpi-unit"></div></div>
  `;

  // Courbes S
  const model = {};
  MODELS.forEach(mn => {
    const rows = prevPhase.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = { total, start: dates.length ? new Date(Math.min(...dates)) : null, end: dates.length ? new Date(Math.max(...dates)) : null };
  });

  const realDates = realPhase.map(r => toDate(r['Date du projet'] || r['Date'])).filter(Boolean);
  const allDates = [];
  MODELS.forEach(mn => { if (model[mn].start) allDates.push(model[mn].start); if (model[mn].end) allDates.push(model[mn].end); });
  realDates.forEach(d => allDates.push(d));
  if (!allDates.length) return;

  const globalMin = new Date(Math.min(...allDates)), globalMax = new Date(Math.max(...allDates));
  const modelMin = new Date(Math.min(...MODELS.filter(m => model[m].start).map(m => model[m].start)));
  const modelMax = new Date(Math.max(...MODELS.filter(m => model[m].end).map(m => model[m].end)));
  const axisMin = isSOnly ? modelMin : globalMin;
  const axisMax = isSOnly ? modelMax : globalMax;
  const axis = buildAxis(axisMin, axisMax);

  const series = {};
  MODELS.forEach(mn => { const [si, ei] = modelRange(prevPhase, mn, axis); series[mn] = sCurve(model[mn].total, axis, k, si, ei); });

  const today = new Date(), mmToday = yyyymm(today);
  const realMap = new Map(axis.map(m => [m, 0]));
  let initCum = 0;
  realPhase.forEach(r => {
    const d = toDate(r['Date du projet'] || r['Date']); if (!d) return;
    const mm = yyyymm(firstOfMonth(d)); if (mm > mmToday) return;
    if (realMap.has(mm)) realMap.set(mm, realMap.get(mm) + toNum(r['Montant total du coût']));
    else if (isSOnly && mm < axis[0]) initCum += toNum(r['Montant total du coût']);
  });

  const realMonth = axis.map(m => realMap.get(m) || 0);
  let acc = initCum; const realCum = realMonth.map(v => { acc += v; return acc; });
  const realCut = realCum.map((v, i) => axis[i] > mmToday ? null : v);

  const ds = [
    { label: 'Budget Révisé (S cumul)', data: series['Budget_Rev'], borderColor: GE_ORANGE, backgroundColor: 'rgba(232,135,58,0.07)', pointRadius: 0, tension: 0.3, spanGaps: true, borderWidth: 3 },
    { label: 'AT (S cumul)', data: series['AT'], borderColor: GE_BLUE2, backgroundColor: 'rgba(30,77,140,0.07)', pointRadius: 0, tension: 0.3, spanGaps: true },
    { label: 'Réels cumulés', data: realCut, borderColor: GE_RED, backgroundColor: 'rgba(192,57,43,0.07)', pointRadius: 2, pointHoverRadius: 5, tension: 0.3, spanGaps: true, borderWidth: 3 }
  ];

  if (charts[`chart${phaseName}`]) charts[`chart${phaseName}`].destroy();
  charts[`chart${phaseName}`] = new Chart(document.getElementById(`chart${phaseName}`).getContext('2d'), {
    type: 'line', data: { labels: axis, datasets: ds },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { padding: 15, font: { size: 13, weight: '600' }, usePointStyle: true } },
        tooltip: { mode: 'index', intersect: false, backgroundColor: 'rgba(22,58,95,0.95)', callbacks: { label: ctx => ctx.parsed.y !== null ? `${ctx.dataset.label}: ${toCHF(ctx.parsed.y)} CHF` : '' } },
        todayLine: { label: yyyymm(today) }
      },
      scales: {
        x: { title: { display: true, text: 'Mois' }, grid: { color: '#eef2f7' }, ticks: { font: { size: 11 }, callback: (_, i) => { const l = axis[i]; return l && (l.endsWith('-01') || l.endsWith('-04') || l.endsWith('-07') || l.endsWith('-10')) ? l : ''; } } },
        y: { title: { display: true, text: 'CHF (cumulé)' }, beginAtZero: true, grid: { color: '#eef2f7' }, ticks: { callback: v => toCHF(v), font: { size: 11 } } }
      },
      interaction: { mode: 'index', intersect: false }
    },
    plugins: [todayLinePlugin]
  });

  // Bar chart par tâche
  const tasks = [...new Set([...prevPhase.map(r => r['Nom de la tâche']), ...realPhase.map(r => r['Nom de la tâche'])].filter(Boolean))];
  const labels = [], dBR = [], dAT = [], dR = [];
  tasks.forEach(task => {
    labels.push(task);
    dBR.push(prevPhase.filter(r => r['Nom de la tâche'] === task && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0));
    dAT.push(prevPhase.filter(r => r['Nom de la tâche'] === task && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0));
    dR.push(realPhase.filter(r => r['Nom de la tâche'] === task).reduce((a, r) => a + toNum(r['Montant total du coût']), 0));
  });

  if (charts[`chartBar${phaseName}`]) charts[`chartBar${phaseName}`].destroy();
  charts[`chartBar${phaseName}`] = new Chart(document.getElementById(`chartBar${phaseName}`).getContext('2d'), {
    type: 'bar', data: {
      labels, datasets: [
        { label: 'Budget Révisé', data: dBR, backgroundColor: GE_ORANGE },
        { label: 'AT', data: dAT, backgroundColor: GE_BLUE2 },
        { label: 'Réel', data: dR, backgroundColor: GE_RED }
      ]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF` } } },
      scales: { x: { ticks: { callback: v => toCHF(v) } } }
    }
  });

  // Tableau comparatif par tâche
  const table = document.getElementById(`table${phaseName}`);
  table.innerHTML = `<thead><tr><th>Tâche</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart / BR</th><th>Écart / AT</th><th>% / AT</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb = table.querySelector('tbody'), tf = table.querySelector('tfoot');
  let sumBR = 0, sumAT = 0, sumR = 0;
  tasks.forEach(task => {
    const brV = prevPhase.filter(r => r['Nom de la tâche'] === task && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const atV = prevPhase.filter(r => r['Nom de la tâche'] === task && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const rV = realPhase.filter(r => r['Nom de la tâche'] === task).reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    sumBR += brV; sumAT += atV; sumR += rV;
    const eR = rV - brV, eA = rV - atV, pct = atV > 0 ? (rV / atV * 100).toFixed(1) + '%' : '—';
    tb.insertAdjacentHTML('beforeend', `<tr><td><strong>${task}</strong></td><td>${toCHF(brV)}</td><td>${toCHF(atV)}</td><td>${toCHF(rV)}</td><td class="${eR <= 0 ? 'positive' : 'negative'}">${toCHF(eR)}</td><td class="${eA <= 0 ? 'positive' : 'negative'}">${toCHF(eA)}</td><td>${pct}</td></tr>`);
  });
  const pT = sumAT > 0 ? (sumR / sumAT * 100).toFixed(1) + '%' : '—';
  tf.innerHTML = `<tr><td><strong>TOTAL</strong></td><td>${toCHF(sumBR)}</td><td>${toCHF(sumAT)}</td><td>${toCHF(sumR)}</td><td class="${sumR - sumBR <= 0 ? 'positive' : 'negative'}">${toCHF(sumR - sumBR)}</td><td class="${sumR - sumAT <= 0 ? 'positive' : 'negative'}">${toCHF(sumR - sumAT)}</td><td><strong>${pT}</strong></td></tr>`;
}

// ============================================================
// PROCESS SIA 5 (conserve logique existante sous-projets)
// ============================================================
function processSIA5() {
  const k = parseFloat(kNumber.value) || 12;
  const isSOnly = viewSOnlyEl.checked;
  const MODELS = ['Budget_Rev', 'AT'];

  const SIA5_PROJECTS = {
    'CL150096_11_05_01': 'Général - Direction',
    'CL150096_11_05_02': 'Centrale',
    'CL150096_11_05_03': 'Bâtiment',
    'CL150096_11_05_04': 'Réseau CAD',
    'CL150096_11_05_05': 'Sous-stations',
    'CL150096_11_05_06': 'Participations clients'
  };
  const PROJECT_ORDER = Object.keys(SIA5_PROJECTS);
  const getPN = id => SIA5_PROJECTS[id] || id;

  const prevSIA5 = prevRows.filter(r => extractPhase(r['ID projet']) === 'SIA5');
  const realSIA5 = realRows.filter(r => extractPhase(r['ID Projet'] || r['ID projet']) === 'SIA5');

  const totalBR = prevSIA5.filter(r => r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const totalAT = prevSIA5.filter(r => r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const totalR = realSIA5.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
  const ecart = totalR - totalAT, pct = totalAT > 0 ? (totalR / totalAT * 100).toFixed(1) : 0;

  document.getElementById('kpiSIA5').style.display = 'grid';
  document.getElementById('kpiSIA5').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Budget Révisé Total</div><div class="kpi-value">${toCHF(totalBR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">AT Total</div><div class="kpi-value">${toCHF(totalAT)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Coûts Réels</div><div class="kpi-value">${toCHF(totalR)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">Écart Réel / AT</div><div class="kpi-value ${ecart <= 0 ? 'positive' : 'negative'}">${toCHF(ecart)}</div><div class="kpi-unit">CHF</div></div>
    <div class="kpi-card"><div class="kpi-label">% Avanc. / AT</div><div class="kpi-value ${pct > 100 ? 'negative' : 'positive'}">${pct}%</div><div class="kpi-unit"></div></div>
  `;

  // Courbes S SIA5
  const model = {};
  MODELS.forEach(mn => {
    const rows = prevSIA5.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = { total, start: dates.length ? new Date(Math.min(...dates)) : null, end: dates.length ? new Date(Math.max(...dates)) : null };
  });

  const realDates = realSIA5.map(r => toDate(r['Date du projet'] || r['Date'])).filter(Boolean);
  const allDates = [];
  MODELS.forEach(mn => { if (model[mn].start) allDates.push(model[mn].start); if (model[mn].end) allDates.push(model[mn].end); });
  realDates.forEach(d => allDates.push(d));
  if (!allDates.length) return;

  const globalMin = new Date(Math.min(...allDates)), globalMax = new Date(Math.max(...allDates));
  const modelMin = new Date(Math.min(...MODELS.filter(m => model[m].start).map(m => model[m].start)));
  const modelMax = new Date(Math.max(...MODELS.filter(m => model[m].end).map(m => model[m].end)));
  const axisMin = isSOnly ? modelMin : globalMin;
  const axisMax = isSOnly ? modelMax : globalMax;
  const axis = buildAxis(axisMin, axisMax);

  const series = {};
  MODELS.forEach(mn => { const [si, ei] = modelRange(prevSIA5, mn, axis); series[mn] = sCurve(model[mn].total, axis, k, si, ei); });

  const today = new Date(), mmToday = yyyymm(today);
  const realMap = new Map(axis.map(m => [m, 0]));
  let initCum = 0;
  realSIA5.forEach(r => {
    const d = toDate(r['Date du projet'] || r['Date']); if (!d) return;
    const mm = yyyymm(firstOfMonth(d)); if (mm > mmToday) return;
    if (realMap.has(mm)) realMap.set(mm, realMap.get(mm) + toNum(r['Montant total du coût']));
    else if (isSOnly && mm < axis[0]) initCum += toNum(r['Montant total du coût']);
  });

  const realMonth = axis.map(m => realMap.get(m) || 0);
  let acc = initCum; const realCum = realMonth.map(v => { acc += v; return acc; });
  const realCut = realCum.map((v, i) => axis[i] > mmToday ? null : v);

  const ds = [
    { label: 'Budget Révisé (S cumul)', data: series['Budget_Rev'], borderColor: GE_ORANGE, backgroundColor: 'rgba(232,135,58,0.07)', pointRadius: 0, tension: 0.3, spanGaps: true, borderWidth: 3 },
    { label: 'AT (S cumul)', data: series['AT'], borderColor: GE_BLUE2, backgroundColor: 'rgba(30,77,140,0.07)', pointRadius: 0, tension: 0.3, spanGaps: true },
    { label: 'Réels cumulés', data: realCut, borderColor: GE_RED, backgroundColor: 'rgba(192,57,43,0.07)', pointRadius: 2, pointHoverRadius: 5, tension: 0.3, spanGaps: true, borderWidth: 3 }
  ];

  if (charts.chartSIA5) charts.chartSIA5.destroy();
  charts.chartSIA5 = new Chart(document.getElementById('chartSIA5').getContext('2d'), {
    type: 'line', data: { labels: axis, datasets: ds },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { padding: 15, font: { size: 13, weight: '600' }, usePointStyle: true } },
        tooltip: { mode: 'index', intersect: false, backgroundColor: 'rgba(22,58,95,0.95)', callbacks: { label: ctx => ctx.parsed.y !== null ? `${ctx.dataset.label}: ${toCHF(ctx.parsed.y)} CHF` : '' } },
        todayLine: { label: yyyymm(today) }
      },
      scales: {
        x: { title: { display: true, text: 'Mois' }, grid: { color: '#eef2f7' }, ticks: { font: { size: 11 }, callback: (_, i) => { const l = axis[i]; return l && (l.endsWith('-01') || l.endsWith('-04') || l.endsWith('-07') || l.endsWith('-10')) ? l : ''; } } },
        y: { title: { display: true, text: 'CHF (cumulé)' }, beginAtZero: true, grid: { color: '#eef2f7' }, ticks: { callback: v => toCHF(v), font: { size: 11 } } }
      },
      interaction: { mode: 'index', intersect: false }
    },
    plugins: [todayLinePlugin]
  });

  // Pie charts SIA5
  const labelsBR = [], dataBR = [], labelsAT = [], dataAT = [];
  PROJECT_ORDER.forEach(id => {
    const br = prevSIA5.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const at = prevSIA5.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    if (br > 0) { labelsBR.push(getPN(id)); dataBR.push(br); }
    if (at > 0) { labelsAT.push(getPN(id)); dataAT.push(at); }
  });

  if (charts.chartPieBudgetRevSIA5) charts.chartPieBudgetRevSIA5.destroy();
  charts.chartPieBudgetRevSIA5 = new Chart(document.getElementById('chartPieBudgetRevSIA5').getContext('2d'), {
    type: 'pie',
    data: { labels: labelsBR, datasets: [{ data: dataBR, backgroundColor: PIE_COLORS, borderWidth: 2, borderColor: '#ffffff' }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 12 }, padding: 14, usePointStyle: true, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${toCHF(ctx.parsed)} CHF (${((ctx.parsed / dataBR.reduce((a, b) => a + b, 0)) * 100).toFixed(1)}%)` } }
      }
    }
  });

  if (charts.chartPieATSIA5) charts.chartPieATSIA5.destroy();
  charts.chartPieATSIA5 = new Chart(document.getElementById('chartPieATSIA5').getContext('2d'), {
    type: 'pie',
    data: { labels: labelsAT, datasets: [{ data: dataAT, backgroundColor: PIE_COLORS, borderWidth: 2, borderColor: '#ffffff' }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 12 }, padding: 14, usePointStyle: true, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${toCHF(ctx.parsed)} CHF (${((ctx.parsed / dataAT.reduce((a, b) => a + b, 0)) * 100).toFixed(1)}%)` } }
      }
    }
  });

  // Bar chart par projet SIA5
  const labels = [], dBR = [], dAT = [], dR = [];
  PROJECT_ORDER.forEach(id => {
    labels.push(getPN(id));
    dBR.push(prevSIA5.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0));
    dAT.push(prevSIA5.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0));
    dR.push(realSIA5.filter(r => r['ID Projet'] === id).reduce((a, r) => a + toNum(r['Montant total du coût']), 0));
  });

  if (charts.chartBarSIA5) charts.chartBarSIA5.destroy();
  charts.chartBarSIA5 = new Chart(document.getElementById('chartBarSIA5').getContext('2d'), {
    type: 'bar', data: {
      labels, datasets: [
        { label: 'Budget Révisé', data: dBR, backgroundColor: GE_ORANGE },
        { label: 'AT', data: dAT, backgroundColor: GE_BLUE2 },
        { label: 'Réel', data: dR, backgroundColor: GE_RED }
      ]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF` } } },
      scales: { x: { ticks: { callback: v => toCHF(v) } } }
    }
  });

  // Tableau comparatif SIA5
  const table = document.getElementById('tableSIA5');
  table.innerHTML = `<thead><tr><th>Projet</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart / BR</th><th>Écart / AT</th><th>% / AT</th></tr></thead><tbody></tbody><tfoot></tfoot>`;
  const tb = table.querySelector('tbody'), tf = table.querySelector('tfoot');
  let sumBR = 0, sumAT = 0, sumR = 0;
  PROJECT_ORDER.forEach(id => {
    const brV = prevSIA5.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const atV = prevSIA5.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'AT').reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const rV = realSIA5.filter(r => r['ID Projet'] === id).reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    sumBR += brV; sumAT += atV; sumR += rV;
    const eR = rV - brV, eA = rV - atV, pct = atV > 0 ? (rV / atV * 100).toFixed(1) + '%' : '—';
    tb.insertAdjacentHTML('beforeend', `<tr><td><strong>${getPN(id)}</strong></td><td>${toCHF(brV)}</td><td>${toCHF(atV)}</td><td>${toCHF(rV)}</td><td class="${eR <= 0 ? 'positive' : 'negative'}">${toCHF(eR)}</td><td class="${eA <= 0 ? 'positive' : 'negative'}">${toCHF(eA)}</td><td>${pct}</td></tr>`);
  });
  const pT = sumAT > 0 ? (sumR / sumAT * 100).toFixed(1) + '%' : '—';
  tf.innerHTML = `<tr><td><strong>TOTAL</strong></td><td>${toCHF(sumBR)}</td><td>${toCHF(sumAT)}</td><td>${toCHF(sumR)}</td><td class="${sumR - sumBR <= 0 ? 'positive' : 'negative'}">${toCHF(sumR - sumBR)}</td><td class="${sumR - sumAT <= 0 ? 'positive' : 'negative'}">${toCHF(sumR - sumAT)}</td><td><strong>${pT}</strong></td></tr>`;

  // Honoraires par métier SIA5
  const extractTypeHono = nomTache => {
    if (!nomTache || typeof nomTache !== 'string') return nomTache || 'Autre';
    return nomTache.replace(/_(?:Réseau CAD|Sous-stations|Général|Centrale|Bâtiment|Participations|SIA).*$/, '').trim() || nomTache;
  };
  const isHono = r => r['Nom de la tâche'] && r['Nom de la tâche'].toLowerCase().includes('honoraires');

  const typesBR = {}, typesAT = {}, typesR = {};
  prevSIA5.forEach(r => {
    if (!isHono(r)) return;
    const t = extractTypeHono(r['Nom de la tâche']), v = toNum(r['Montant total du coût']);
    if (r['Modèle de prévision'] === 'Budget_Rev') typesBR[t] = (typesBR[t] || 0) + v;
    if (r['Modèle de prévision'] === 'AT') typesAT[t] = (typesAT[t] || 0) + v;
  });
  realSIA5.forEach(r => {
    if (!isHono(r)) return;
    const t = extractTypeHono(r['Nom de la tâche']);
    typesR[t] = (typesR[t] || 0) + toNum(r['Montant total du coût']);
  });

  const allTypes = [...new Set([...Object.keys(typesBR), ...Object.keys(typesAT), ...Object.keys(typesR)])];
  if (charts.chartHonoMetierSIA5) charts.chartHonoMetierSIA5.destroy();
  charts.chartHonoMetierSIA5 = new Chart(document.getElementById('chartHonoMetierSIA5').getContext('2d'), {
    type: 'bar', data: {
      labels: allTypes, datasets: [
        { label: 'Budget Révisé', data: allTypes.map(t => typesBR[t] || 0), backgroundColor: GE_ORANGE },
        { label: 'AT', data: allTypes.map(t => typesAT[t] || 0), backgroundColor: GE_BLUE2 },
        { label: 'Réel', data: allTypes.map(t => typesR[t] || 0), backgroundColor: GE_RED }
      ]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${toCHF(ctx.parsed.x)} CHF` } } },
      scales: { x: { ticks: { callback: v => toCHF(v) } } }
    }
  });

  // Gantt SIA5
  drawGanttSIA5();
}

function drawGanttSIA5() {
  const SIA5_PROJECTS = {
    'CL150096_11_05_01': 'Général - Direction',
    'CL150096_11_05_02': 'Centrale',
    'CL150096_11_05_03': 'Bâtiment',
    'CL150096_11_05_04': 'Réseau CAD',
    'CL150096_11_05_05': 'Sous-stations',
    'CL150096_11_05_06': 'Participations clients'
  };
  const PROJECT_ORDER = Object.keys(SIA5_PROJECTS);
  const getPN = id => SIA5_PROJECTS[id] || id;

  const prevSIA5 = prevRows.filter(r => extractPhase(r['ID projet']) === 'SIA5');
  const realSIA5 = realRows.filter(r => extractPhase(r['ID Projet'] || r['ID projet']) === 'SIA5');

  const plan = {};
  PROJECT_ORDER.forEach(id => {
    const getR = mode => {
      const rows = prevSIA5.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === mode);
      const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
      return dates.length ? { s: new Date(Math.min(...dates)), e: new Date(Math.max(...dates)) } : { s: null, e: null };
    };
    const br = getR('Budget_Rev'), at = getR('AT');
    const rd = realSIA5.filter(r => r['ID Projet'] === id).map(r => toDate(r['Date du projet'] || r['Date'])).filter(Boolean);
    const rS = rd.length ? new Date(Math.min(...rd)) : null, rE = rd.length ? new Date(Math.max(...rd)) : null;
    plan[id] = { name: getPN(id), brS: br.s, brE: br.e, atS: at.s, atE: at.e, rS, rE };
  });

  const allD = [];
  Object.values(plan).forEach(p => { [p.brS, p.brE, p.atS, p.atE, p.rS, p.rE].forEach(d => { if (d && !isNaN(d)) allD.push(d); }); });
  if (!allD.length) return;

  const minD = new Date(Math.min(...allD)), maxD = new Date(Math.max(...allD));
  const span = Math.round((maxD - minD) / 86400000) + 60;
  const toDay = d => (d && !isNaN(d)) ? Math.round((d - minD) / 86400000) : null;

  const yLabels = PROJECT_ORDER.map(id => getPN(id));
  const mkPoint = (s, e, label) => {
    const s0 = toDay(s), e0 = toDay(e);
    if (s0 === null || e0 === null) return { x: [null, null], y: label };
    return { x: [s0, e0], y: label };
  };

  const dsBR = {
    label: 'Budget Révisé', data: PROJECT_ORDER.map(id => mkPoint(plan[id].brS, plan[id].brE, plan[id].name)),
    backgroundColor: GE_ORANGE + 'cc', borderColor: GE_ORANGE, borderWidth: 2, borderSkipped: false, borderRadius: 3, barThickness: 18
  };
  const dsAT = {
    label: 'AT', data: PROJECT_ORDER.map(id => mkPoint(plan[id].atS, plan[id].atE, plan[id].name)),
    backgroundColor: GE_BLUE2 + 'cc', borderColor: GE_BLUE2, borderWidth: 2, borderSkipped: false, borderRadius: 3, barThickness: 18
  };
  const dsReal = {
    label: 'Réel', data: PROJECT_ORDER.map(id => mkPoint(plan[id].rS, plan[id].rE, plan[id].name)),
    backgroundColor: GE_RED + 'cc', borderColor: GE_RED, borderWidth: 2, borderSkipped: false, borderRadius: 3, barThickness: 18
  };

  const tickStep = Math.max(30, Math.round(span / 12));
  if (charts.chartGanttSIA5) charts.chartGanttSIA5.destroy();
  charts.chartGanttSIA5 = new Chart(document.getElementById('chartGanttSIA5').getContext('2d'), {
    type: 'bar', data: { labels: yLabels, datasets: [dsBR, dsAT, dsReal] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { usePointStyle: true, padding: 16, font: { size: 12, weight: '600' } } },
        tooltip: {
          callbacks: {
            title: ctx => `${ctx[0].label} — ${ctx[0].dataset.label}`,
            label: ctx => {
              const [d0, d1] = ctx.parsed.x;
              if (d0 === null || d1 === null) return 'Pas de données';
              const sD = new Date(minD.getTime() + d0 * 86400000), eD = new Date(minD.getTime() + d1 * 86400000);
              const dur = Math.round((eD - sD) / 86400000);
              return [`Début : ${sD.toLocaleDateString('fr-CH')}`, `Fin   : ${eD.toLocaleDateString('fr-CH')}`, `Durée : ${dur} jours`];
            }
          }
        }
      },
      scales: {
        x: {
          type: 'linear', min: 0, max: span, title: { display: true, text: 'Timeline' }, grid: { color: 'rgba(0,0,0,0.06)' },
          ticks: {
            stepSize: tickStep,
            callback: value => {
              const d = new Date(minD.getTime() + value * 86400000);
              return d.toLocaleDateString('fr-CH', { month: 'short', year: '2-digit' });
            }
          }
        },
        y: { type: 'category', labels: yLabels, offset: true, grid: { color: 'rgba(0,0,0,0.08)', lineWidth: 1 }, ticks: { font: { weight: 'bold', size: 11 }, color: '#163a5f' } }
      }
    }
  });
}

// ============================================================
// EXPORTS
// ============================================================
function exportChartPNG(canvasId, filename) {
  const canvas = document.getElementById(canvasId); if (!canvas) return;
  const a = document.createElement('a'); a.download = filename + '.png'; a.href = canvas.toDataURL('image/png', 1); a.click();
}
window.exportChartPNG = exportChartPNG;

async function exportTablePNG(cardId, filename) {
  const el = document.getElementById(cardId); if (!el) return;
  const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff', logging: false });
  canvas.toBlob(blob => { const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.download = filename + '.png'; a.href = url; a.click(); URL.revokeObjectURL(url); });
}
window.exportTablePNG = exportTablePNG;

function exportTableXLSX(tableId, sheetName) {
  const el = document.getElementById(tableId); if (!el) return;
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(el), sheetName.substring(0, 31)); XLSX.writeFile(wb, sheetName + '.xlsx');
}
window.exportTableXLSX = exportTableXLSX;

// ============================================================
// PDF GÉNÉRATION SÉLECTIVE
// ============================================================
async function generatePDF() {
  const selectedPhase = document.getElementById('pdfPhaseSelect').value;
  const { jsPDF } = window.jspdf;
  const today = new Date(), dd = String(today.getDate()).padStart(2, '0'), mm = String(today.getMonth() + 1).padStart(2, '0'), yyyy = today.getFullYear();
  const fname = `${yyyy}${mm}${dd}_Etat_Projet_${selectedPhase.toUpperCase()}.pdf`;
  const dateStr = today.toLocaleDateString('fr-CH', { day: '2-digit', month: 'long', year: 'numeric' });
  const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const W = 297, H = 210, PAD = 10, HDR = 18, FTR = 8;
  const CONTENT_H = H - HDR - FTR - PAD, CONTENT_W = W - 2 * PAD;

  function drawHeader(title, pageLabel) {
    pdf.setFillColor(22, 58, 95); pdf.rect(0, 0, W, HDR, 'F'); pdf.setFillColor(232, 135, 58); pdf.rect(0, HDR - 1.5, W, 1.5, 'F');
    pdf.setTextColor(255, 255, 255); pdf.setFontSize(10); pdf.setFont('helvetica', 'bold'); pdf.text(title, PAD, 11.5);
    pdf.setFontSize(8); pdf.setFont('helvetica', 'normal'); pdf.text(`${fname.replace('.pdf', '')}   |   ${pageLabel}`, W - PAD, 11.5, { align: 'right' });
  }
  function drawFooter(page, total) {
    pdf.setFillColor(240, 244, 249); pdf.rect(0, H - FTR, W, FTR, 'F'); pdf.setTextColor(120, 140, 160); pdf.setFontSize(7.5); pdf.setFont('helvetica', 'normal');
    pdf.text('Groupe E Celsius — Dashboard Projets CAD — Confidentiel', PAD, H - 2); pdf.text(`Page ${page} / ${total}`, W - PAD, H - 2, { align: 'right' });
  }

  // Page de garde
  pdf.setFillColor(22, 58, 95); pdf.rect(0, 0, W, H, 'F'); pdf.setFillColor(232, 135, 58); pdf.rect(0, 0, 8, H, 'F');
  pdf.setFillColor(255, 255, 255); pdf.roundedRect(30, 40, W - 60, 130, 6, 6, 'F');
  pdf.setTextColor(22, 58, 95); pdf.setFontSize(26); pdf.setFont('helvetica', 'bold'); pdf.text('Dashboard Projets CAD', W / 2, 82, { align: 'center' });
  pdf.setFontSize(15); pdf.setFont('helvetica', 'normal'); pdf.text('Groupe E Celsius', W / 2, 97, { align: 'center' });
  pdf.setDrawColor(232, 135, 58); pdf.setLineWidth(0.8); pdf.line(W / 2 - 40, 103, W / 2 + 40, 103);
  const phaseName = selectedPhase === 'recap' ? 'Récap Global' : selectedPhase === 'sia2' ? 'SIA 2' : selectedPhase === 'sia3' ? 'SIA 3' : selectedPhase === 'sia4' ? 'SIA 4' : 'SIA 5';
  pdf.setFontSize(12); pdf.setTextColor(22, 58, 95); pdf.text(`État de Projet — ${phaseName}`, W / 2, 112, { align: 'center' });
  pdf.setFontSize(10); pdf.setTextColor(90, 122, 154); pdf.text(`Généré le ${dateStr}`, W / 2, 124, { align: 'center' });
  pdf.setFontSize(9); pdf.setTextColor(180, 190, 200); pdf.text(fname, W / 2, 160, { align: 'center' });

  // Sélection des sections selon la phase
  let sections = [];
  if (selectedPhase === 'recap') {
    sections = [
      { id: 'kpiSectionGlobal', title: 'Indicateurs Clés Globaux', type: 'kpi' },
      { id: 'cardCourbesSGlobal', title: 'Courbes S Consolidées', type: 'full' },
      { id: 'cardPiePhases', title: 'Répartition par Phase', type: 'full' },
      { id: 'tablePhaseCard', title: 'Consolidation par Phase', type: 'full' },
      { id: 'cardGanttGlobal', title: 'Gantt Multi-Phase', type: 'full' }
    ];
  } else if (selectedPhase === 'sia2') {
    sections = [
      { id: 'kpiSIA2', title: 'KPI SIA 2', type: 'kpi' },
      { id: 'cardCourbesSIA2', title: 'Courbes S — SIA 2', type: 'full' },
      { id: 'cardBarSIA2', title: 'Avancement par Tâche', type: 'full' },
      { id: 'tableSIA2Card', title: 'Comparatif par Tâche', type: 'full' }
    ];
  } else if (selectedPhase === 'sia3') {
    sections = [
      { id: 'kpiSIA3', title: 'KPI SIA 3', type: 'kpi' },
      { id: 'cardCourbesSIA3', title: 'Courbes S — SIA 3', type: 'full' },
      { id: 'cardBarSIA3', title: 'Avancement par Tâche', type: 'full' },
      { id: 'tableSIA3Card', title: 'Comparatif par Tâche', type: 'full' }
    ];
  } else if (selectedPhase === 'sia4') {
    sections = [
      { id: 'kpiSIA4', title: 'KPI SIA 4', type: 'kpi' },
      { id: 'cardCourbesSIA4', title: 'Courbes S — SIA 4', type: 'full' },
      { id: 'cardBarSIA4', title: 'Avancement par Tâche', type: 'full' },
      { id: 'tableSIA4Card', title: 'Comparatif par Tâche', type: 'full' }
    ];
  } else if (selectedPhase === 'sia5') {
    sections = [
      { id: 'kpiSIA5', title: 'KPI SIA 5', type: 'kpi' },
      { id: 'cardCourbesSIA5', title: 'Courbes S — SIA 5', type: 'full' },
      { ids: ['cardPieBRSIA5', 'cardPieATSIA5'], titles: ['Répartition BR', 'Répartition AT'], type: 'double' },
      { id: 'cardBarProjetSIA5', title: 'Avancement par Projet', type: 'full' },
      { id: 'tableSIA5Card', title: 'Comparatif par Projet', type: 'full' },
      { id: 'cardHonoMetierSIA5', title: 'Honoraires par Type', type: 'full' },
      { id: 'cardGanttSIA5', title: 'Planning Gantt SIA 5', type: 'full' }
    ];
  }

  async function captureEl(id) {
    const el = document.getElementById(id); if (!el || el.style.display === 'none') return null;
    try { return await html2canvas(el, { scale: 2, backgroundColor: '#ffffff', logging: false, useCORS: true }); }
    catch (e) { console.warn('capture error', id, e); return null; }
  }

  const captured = {};
  for (const sec of sections) {
    if (sec.type === 'double') { for (const id of sec.ids) captured[id] = await captureEl(id); }
    else { captured[sec.id] = await captureEl(sec.id); }
  }

  let totalPages = 1;
  for (const sec of sections) {
    if (sec.type === 'double') { totalPages++; continue; }
    const c = captured[sec.id]; if (!c) continue; const imgH = CONTENT_W * (c.height / c.width); totalPages += Math.ceil(imgH / CONTENT_H);
  }

  let pageNum = 1; drawFooter(pageNum, totalPages);

  for (const sec of sections) {
    if (sec.type === 'double') {
      const canvases = sec.ids.map(id => captured[id]).filter(Boolean); if (!canvases.length) continue; pageNum++; pdf.addPage();
      drawHeader(sec.titles.join('  /  '), `${pageNum}/${totalPages}`); drawFooter(pageNum, totalPages);
      const colW = (CONTENT_W - 8) / 2, yTop = HDR + PAD;
      canvases.forEach((c, ci) => {
        const xOff = PAD + ci * (colW + 8), ratio = c.height / c.width, iH = Math.min(colW * ratio, CONTENT_H - 6);
        pdf.setFillColor(255, 255, 255); pdf.roundedRect(xOff - 2, yTop - 2, colW + 4, iH + 4 + 10, 3, 3, 'F');
        pdf.setDrawColor(220, 228, 240); pdf.setLineWidth(0.4); pdf.roundedRect(xOff - 2, yTop - 2, colW + 4, iH + 4 + 10, 3, 3, 'S');
        pdf.setFontSize(9); pdf.setFont('helvetica', 'bold'); pdf.setTextColor(22, 58, 95); pdf.text(sec.titles[ci] || '', xOff + colW / 2, yTop + 7, { align: 'center' });
        const imgData = c.toDataURL('image/jpeg', 0.92); pdf.addImage(imgData, 'JPEG', xOff, yTop + 10, colW, iH);
      });
      continue;
    }
    if (sec.type === 'kpi') {
      const c = captured[sec.id]; if (!c) continue; const imgH_mm = CONTENT_W * (c.height / c.width); pageNum++; pdf.addPage();
      drawHeader(sec.title, `${pageNum}/${totalPages}`); drawFooter(pageNum, totalPages); const iH = Math.min(imgH_mm, CONTENT_H);
      const imgData = c.toDataURL('image/jpeg', 0.93); pdf.addImage(imgData, 'JPEG', PAD, HDR + 4, CONTENT_W, iH); continue;
    }
    const c = captured[sec.id]; if (!c) continue; const srcW = c.width, srcH = c.height;
    const imgH_mm = CONTENT_W * (srcH / srcW), pagesNeeded = Math.ceil(imgH_mm / CONTENT_H);
    for (let pg = 0; pg < pagesNeeded; pg++) {
      pageNum++; pdf.addPage(); const titleSuffix = pagesNeeded > 1 ? ` (${pg + 1}/${pagesNeeded})` : '';
      drawHeader(sec.title + titleSuffix, `${pageNum}/${totalPages}`); drawFooter(pageNum, totalPages);
      const slicePxH = Math.ceil(srcH / pagesNeeded), srcYStart = pg * slicePxH, srcYEnd = Math.min(srcYStart + slicePxH, srcH);
      const sl = document.createElement('canvas'); sl.width = srcW; sl.height = srcYEnd - srcYStart;
      sl.getContext('2d').drawImage(c, 0, -srcYStart); const slData = sl.toDataURL('image/jpeg', 0.92);
      const slH_mm = CONTENT_W * (sl.height / sl.width); pdf.addImage(slData, 'JPEG', PAD, HDR + 4, CONTENT_W, Math.min(slH_mm, CONTENT_H));
    }
  }
  pdf.save(fname);
}
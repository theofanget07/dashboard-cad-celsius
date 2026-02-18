/* ============================================================
   Dashboard CAD — Groupe E Celsius
   script.js v5.0 — Multi-phase SIA 2/3/4/5
   Architecture: mono-fichier modulaire
   Données: Prevision_*.xlsx + Transactions-de-projet_*.xlsx
   ============================================================ */

'use strict';

// ═══════════════════════════════════════════════════════════
// GLOBALS
// ═══════════════════════════════════════════════════════════
let prevRows = [], realRows = [];
let manualPlanningDates = {};

// Registre centralisé des instances Chart.js
const charts = {};

// ═══════════════════════════════════════════════════════════
// CONSTANTES MÉTIER
// ═══════════════════════════════════════════════════════════
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

// Labels phases SIA
const PHASE_LABELS = {
  'SIA2': { label: 'SIA 2 — Avant-Projet',         color: '#2e6db4', emoji: '🔷' },
  'SIA3': { label: 'SIA 3 — Projet de Réalisation', color: '#e8873a', emoji: '🔶' },
  'SIA4': { label: 'SIA 4 — Appels d\'Offres',       color: '#5b9bd5', emoji: '🔵' },
  'SIA5': { label: 'SIA 5 — Réalisation',            color: '#27ae60', emoji: '🟢' }
};

// Palette Groupe E
const GE_BLUE   = '#163a5f';
const GE_BLUE2  = '#1e4d8c';
const GE_BLUE3  = '#2e6db4';
const GE_BLUE4  = '#5b9bd5';
const GE_ORANGE = '#e8873a';
const GE_RED    = '#c0392b';
const GE_GREEN  = '#27ae60';
const PIE_COLORS = ['#2e6db4','#e8873a','#27ae60','#8e44ad','#c0392b','#16a085','#f39c12','#2980b9'];

// ═══════════════════════════════════════════════════════════
// UTILITAIRES
// ═══════════════════════════════════════════════════════════

function readExcel(file, cb) {
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { raw: true, defval: '' });
      cb(rows, null);
    } catch(err) {
      cb(null, err);
    }
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
  return (n || 0).toLocaleString('fr-CH', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function toDate(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') {
    // Numéro de série Excel (base 1899-12-30)
    return new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  }
  if (typeof v === 'string') {
    const cleaned = v.trim().replace(' 00:00:00','').replace('T00:00:00','');
    if (!cleaned) return null;
    let d = new Date(cleaned);
    if (!isNaN(d)) return d;
    // Format DD.MM.YYYY ou DD/MM/YYYY
    const m = cleaned.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
    if (m) {
      const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
      return new Date(y, +m[2]-1, +m[1]);
    }
  }
  return null;
}

// Normalise date au 1er du mois
function firstOfMonth(d) {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

// Format YYYY-MM
function yyyymm(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
}

// Construit axe temporel mensuel [min..max]
function buildAxis(minD, maxD) {
  const axis = [];
  const c = firstOfMonth(new Date(minD));
  const e = firstOfMonth(new Date(maxD));
  while (c <= e) {
    axis.push(yyyymm(c));
    c.setMonth(c.getMonth() + 1);
  }
  return axis;
}

// Calcule courbe sigmoïde cumulée (S-curve)
function sCurve(total, axis, k, sIdx, eIdx) {
  const N = axis.length;
  const cum = Array(N).fill(null);
  const L = eIdx - sIdx + 1;
  if (L <= 0 || total === 0 || sIdx < 0 || eIdx < 0) return cum;
  const s0 = 1 / (1 + Math.exp(-k * (0 - 0.5)));
  const s1 = 1 / (1 + Math.exp(-k * (1 - 0.5)));
  for (let j = 0; j < L; j++) {
    const i = sIdx + j;
    if (i >= N) break;
    const x = j / (L - 1 || 1);
    const sx = 1 / (1 + Math.exp(-k * (x - 0.5)));
    cum[i] = ((sx - s0) / (s1 - s0)) * total;
  }
  return cum;
}

// Retourne [sIdx, eIdx] dans l'axe pour un modèle donné
function modelRange(rows, modelName, axis) {
  const filtered = rows.filter(r => r['Modèle de prévision'] === modelName);
  if (!filtered.length) return [-1, -1];
  const dates = filtered.flatMap(r => [
    toDate(r['Date de début']),
    toDate(r['Date de fin'])
  ]).filter(Boolean);
  if (!dates.length) return [-1, -1];
  const minD = new Date(Math.min(...dates.map(d => d.getTime())));
  const maxD = new Date(Math.max(...dates.map(d => d.getTime())));
  const mmMin = yyyymm(firstOfMonth(minD));
  const mmMax = yyyymm(firstOfMonth(maxD));
  let sIdx = axis.indexOf(mmMin);
  let eIdx = axis.indexOf(mmMax);
  if (sIdx < 0) sIdx = axis.findIndex(m => m >= mmMin);
  if (eIdx < 0) {
    for (let i = axis.length - 1; i >= 0; i--) {
      if (axis[i] <= mmMax) { eIdx = i; break; }
    }
  }
  if (sIdx < 0) sIdx = 0;
  if (eIdx < 0) eIdx = axis.length - 1;
  return [sIdx, eIdx];
}

// Extraction phase depuis ID projet
// "CL150096_11_02" → "SIA2", "CL150096_11_05_01" → "SIA5"
function extractPhase(idProjet) {
  if (!idProjet) return null;
  const s = String(idProjet).trim();
  const parts = s.split('_');
  if (parts.length < 3) return null;
  const code = parts[2]; // "02", "03", "04", "05"
  const mapping = { '02': 'SIA2', '03': 'SIA3', '04': 'SIA4', '05': 'SIA5' };
  return mapping[code] || null;
}

// Récupère la date d'une transaction (colonne "Date" dans Transactions, "Date du projet" sinon)
function getRealDate(row) {
  const raw = row['Date'] || row['Date du projet'] || row['date'] || '';
  return toDate(raw);
}

// Détruit un chart Chart.js existant avant recréation
function destroyChart(key) {
  if (charts[key]) {
    charts[key].destroy();
    delete charts[key];
  }
}

// Génère HTML d'une carte KPI
function kpiCard(label, value, sub, colorClass = '') {
  return `<div class="kpi-card ${colorClass}">
    <div class="kpi-label">${label}</div>
    <div class="kpi-value ${value < 0 ? 'neg' : value > 0 && colorClass === 'kpi-red' ? 'pos' : ''}">${toCHF(value)} CHF</div>
    <div class="kpi-sub">${sub || ''}</div>
  </div>`;
}

// Affiche message vide dans un conteneur
function showEmpty(containerId, msg) {
  const el = document.getElementById(containerId);
  if (el) el.innerHTML = `<div class="empty-state"><div class="icon">📭</div>${msg}</div>`;
}

// Loader overlay
function showLoader(v) {
  const el = document.getElementById('loaderOverlay');
  if (el) el.classList.toggle('active', v);
}

// ═══════════════════════════════════════════════════════════
// PLUGIN "AUJOURD'HUI" — Chart.js
// ═══════════════════════════════════════════════════════════
const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(ch, _args, opts) {
    if (!opts || !opts.label) return;
    const { ctx, chartArea: { top, bottom }, scales: { x } } = ch;
    const labels = ch.data.labels;
    const toT = s => { const [y,m] = s.split('-'); return new Date(+y, +m-1, 1).getTime(); };
    const t = toT(opts.label);
    let bi = 0, bd = Infinity;
    labels.forEach((l, i) => {
      const d = Math.abs(toT(l) - t);
      if (d < bd) { bd = d; bi = i; }
    });
    const xp = x.getPixelForValue(bi);
    ctx.save();
    ctx.strokeStyle = GE_RED;
    ctx.lineWidth = 2;
    ctx.setLineDash([5, 5]);
    ctx.beginPath(); ctx.moveTo(xp, top); ctx.lineTo(xp, bottom); ctx.stroke();
    ctx.setLineDash([]);
    ctx.fillStyle = GE_RED;
    ctx.font = 'bold 11px Arial';
    ctx.textAlign = 'center';
    ctx.fillText("Aujourd'hui", xp, top - 5);
    ctx.restore();
  }
};
Chart.register(todayLinePlugin);

// ═══════════════════════════════════════════════════════════
// INITIALISATION DOM + EVENTS
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {

  // Loader overlay dans le body
  const loaderDiv = document.createElement('div');
  loaderDiv.id = 'loaderOverlay';
  loaderDiv.className = 'loader-overlay';
  loaderDiv.innerHTML = '<div class="loader-box"><div class="loader-spinner"></div><div>Calcul en cours...</div></div>';
  document.body.appendChild(loaderDiv);

  // ── Navigation onglets principaux ──
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      const tab = document.getElementById('tab-' + btn.dataset.tab);
      if (tab) tab.classList.add('active');
    });
  });

  // ── Navigation sous-onglets SIA 5 ──
  document.querySelectorAll('.sub-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.sub-tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.sub-tab-content').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      const sub = document.getElementById(btn.dataset.subtab);
      if (sub) sub.classList.add('active');
    });
  });

  // ── Upload fichiers ──
  document.getElementById('prevFile').addEventListener('change', e => {
    const f = e.target.files[0];
    if (!f) return;
    showLoader(true);
    readExcel(f, (rows, err) => {
      showLoader(false);
      if (err || !rows) { setStatus('❌ Erreur lecture prévisions : ' + (err?.message || ''), 'error'); return; }
      prevRows = rows;
      document.getElementById('prevStatus').textContent = `✅ ${rows.length} lignes chargées`;
      setStatus(`✅ Prévisions : ${rows.length} lignes`);
      maybeProcess();
    });
  });

  document.getElementById('realFile').addEventListener('change', e => {
    const f = e.target.files[0];
    if (!f) return;
    showLoader(true);
    readExcel(f, (rows, err) => {
      showLoader(false);
      if (err || !rows) { setStatus('❌ Erreur lecture réalisations : ' + (err?.message || ''), 'error'); return; }
      realRows = rows;
      document.getElementById('realStatus').textContent = `✅ ${rows.length} lignes chargées`;
      setStatus(`✅ Réalisations : ${rows.length} lignes`);
      maybeProcess();
    });
  });

  // ── Slider coef k ──
  const kSlider = document.getElementById('coefK');
  const kNumber = document.getElementById('coefK_num');
  [kSlider, kNumber].forEach(el => {
    el.addEventListener('input', () => {
      kSlider.value = kNumber.value = el.value;
      if (prevRows.length && realRows.length) processAll();
    });
  });

  // ── Vue globale / S uniquement ──
  ['viewGlobal', 'viewSOnly'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => {
      if (prevRows.length && realRows.length) processAll();
    });
  });

  // ── Bouton recalcul Gantt ──
  document.getElementById('btnApplyPlanning').addEventListener('click', () => {
    if (!prevRows.length || !realRows.length) return;
    PROJECT_ORDER.forEach(id => {
      const s = document.getElementById(`rs_${id}`);
      const e = document.getElementById(`re_${id}`);
      if (!manualPlanningDates[id]) manualPlanningDates[id] = {};
      manualPlanningDates[id].start = s?.value ? new Date(s.value) : null;
      manualPlanningDates[id].end   = e?.value ? new Date(e.value) : null;
    });
    processPlanning();
  });

  // ── Export PDF ──
  document.getElementById('btnPDF').addEventListener('click', generatePDF);
});

// ── Helpers status bar ──
function setStatus(msg, type) {
  const el = document.getElementById('status');
  if (!el) return;
  el.textContent = msg;
  el.className = 'status-bar' + (type === 'error' ? ' status-error' : type === 'ok' ? ' status-ok' : '');
}

function maybeProcess() {
  if (prevRows.length && realRows.length) {
    setStatus('⚙️ Calcul en cours...');
    setTimeout(() => {
      try {
        processAll();
        setStatus(`✅ Données chargées — ${prevRows.length} prévisions / ${realRows.length} transactions`, 'ok');
      } catch(e) {
        setStatus('❌ Erreur : ' + e.message, 'error');
        console.error(e);
      }
    }, 50);
  }
}

// ═══════════════════════════════════════════════════════════
// ORCHESTRATEUR PRINCIPAL
// ═══════════════════════════════════════════════════════════
function processAll() {
  processRecap();
  processPhase('SIA2', '02');
  processPhase('SIA3', '03');
  processPhase('SIA4', '04');
  processSIA5Budgets();
  processImputations();
  processPlanning();
}

// ═══════════════════════════════════════════════════════════
// ONGLET RÉCAPITULATIF
// ═══════════════════════════════════════════════════════════
function processRecap() {
  const phases = ['SIA2', 'SIA3', 'SIA4', 'SIA5'];
  const phaseData = {};
  const today = new Date();
  const mmToday = yyyymm(today);

  phases.forEach(ph => {
    const code = ph.replace('SIA','').padStart(2,'0');
    const pRows = prevRows.filter(r => {
      const id = String(r['ID projet'] || '');
      const parts = id.split('_');
      return parts.length >= 3 && parts[2] === code;
    });
    const rRows = realRows.filter(r => {
      const id = String(r['ID projet'] || '');
      const parts = id.split('_');
      return parts.length >= 3 && parts[2] === code;
    });

    const totalBudget  = pRows.filter(r => r['Modèle de prévision'] === 'Budget')
                              .reduce((a,r) => a + toNum(r['Montant total du coût']), 0);
    const totalBR      = pRows.filter(r => r['Modèle de prévision'] === 'Budget_Rev')
                              .reduce((a,r) => a + toNum(r['Montant total du coût']), 0);
    const totalAT      = pRows.filter(r => r['Modèle de prévision'] === 'AT')
                              .reduce((a,r) => a + toNum(r['Montant total du coût']), 0);
    const totalReal    = rRows.reduce((a,r) => {
      const d = getRealDate(r);
      if (!d || yyyymm(d) > mmToday) return a;
      return a + toNum(r['Montant total du coût']);
    }, 0);

    phaseData[ph] = {
      label:   PHASE_LABELS[ph].label,
      budget:  totalBudget,
      br:      totalBR,
      at:      totalAT,
      reel:    totalReal,
      ecartBR: totalReal - totalBR,
      ecartAT: totalReal - totalAT,
      pctBR:   totalBR > 0 ? ((totalReal / totalBR) * 100).toFixed(1) : 0
    };
  });

  // KPIs globaux
  const totBR   = phases.reduce((a,p) => a + phaseData[p].br, 0);
  const totAT   = phases.reduce((a,p) => a + phaseData[p].at, 0);
  const totReal = phases.reduce((a,p) => a + phaseData[p].reel, 0);
  const totEcart = totReal - totAT;

  const kpiEl = document.getElementById('kpiSectionRecap');
  kpiEl.style.display = 'grid';
  kpiEl.innerHTML =
    kpiCard('Budget Révisé Total', totBR, 'Toutes phases', 'kpi-orange') +
    kpiCard('AT Total', totAT, 'Toutes phases', 'kpi-blue2') +
    kpiCard('Réel Total', totReal, 'Imputations cumulées', '') +
    kpiCard('Écart vs AT', totEcart, `${totAT > 0 ? ((totReal/totAT)*100).toFixed(1) : 0}% d'avancement`, totEcart <= 0 ? 'kpi-green' : 'kpi-red');

  // Tableau comparatif
  const table = document.getElementById('tableRecap');
  table.innerHTML = `<thead><tr>
    <th>Phase</th>
    <th class="num">Budget Initial</th>
    <th class="num">Budget Révisé</th>
    <th class="num">AT</th>
    <th class="num">Réel</th>
    <th class="num">Écart BR</th>
    <th class="num">Écart AT</th>
    <th class="num">% BR</th>
  </tr></thead><tbody>` +
    phases.map(ph => {
      const d = phaseData[ph];
      const eBRcls = d.ecartBR <= 0 ? 'green' : 'red';
      const eATcls = d.ecartAT <= 0 ? 'green' : 'red';
      return `<tr>
        <td class="bold">${PHASE_LABELS[ph].emoji} ${d.label}</td>
        <td class="num">${toCHF(d.budget)}</td>
        <td class="num">${toCHF(d.br)}</td>
        <td class="num">${toCHF(d.at)}</td>
        <td class="num">${toCHF(d.reel)}</td>
        <td class="num ${eBRcls}">${d.ecartBR >= 0 ? '+' : ''}${toCHF(d.ecartBR)}</td>
        <td class="num ${eATcls}">${d.ecartAT >= 0 ? '+' : ''}${toCHF(d.ecartAT)}</td>
        <td class="num">${d.pctBR}%</td>
      </tr>`;
    }).join('') +
    `<tr class="bold" style="background:#edf2f7;border-top:2px solid #163a5f">
      <td class="bold">TOTAL</td>
      <td class="num">${toCHF(phases.reduce((a,p)=>a+phaseData[p].budget,0))}</td>
      <td class="num">${toCHF(totBR)}</td>
      <td class="num">${toCHF(totAT)}</td>
      <td class="num">${toCHF(totReal)}</td>
      <td class="num ${totReal-totBR<=0?'green':'red'}">${totReal-totBR>=0?'+':''}${toCHF(totReal-totBR)}</td>
      <td class="num ${totEcart<=0?'green':'red'}">${totEcart>=0?'+':''}${toCHF(totEcart)}</td>
      <td class="num">${totAT>0?((totReal/totAT)*100).toFixed(1):0}%</td>
    </tr>` +
  `</tbody>`;

  // Graphique barres groupées par phase
  destroyChart('recapBar');
  const ctxBar = document.getElementById('chartRecapBar');
  if (ctxBar) {
    charts.recapBar = new Chart(ctxBar, {
      type: 'bar',
      data: {
        labels: phases.map(p => PHASE_LABELS[p].label.split('—')[0].trim()),
        datasets: [
          { label: 'Budget Révisé', data: phases.map(p => phaseData[p].br), backgroundColor: GE_ORANGE, borderRadius: 5 },
          { label: 'AT',            data: phases.map(p => phaseData[p].at), backgroundColor: GE_BLUE2, borderRadius: 5 },
          { label: 'Réel',          data: phases.map(p => phaseData[p].reel), backgroundColor: GE_GREEN, borderRadius: 5 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top' } },
        scales: {
          y: { ticks: { callback: v => toCHF(v) + ' CHF' }, grid: { color: '#e2e8f0' } },
          x: { grid: { display: false } }
        }
      }
    });
  }

  // Camembert AT global
  destroyChart('recapPie');
  const ctxPie = document.getElementById('chartRecapPie');
  const nonZeroPhases = phases.filter(p => phaseData[p].at > 0);
  if (ctxPie && nonZeroPhases.length > 0) {
    charts.recapPie = new Chart(ctxPie, {
      type: 'doughnut',
      data: {
        labels: nonZeroPhases.map(p => PHASE_LABELS[p].emoji + ' ' + PHASE_LABELS[p].label.split('—')[0].trim()),
        datasets: [{
          data: nonZeroPhases.map(p => phaseData[p].at),
          backgroundColor: nonZeroPhases.map((_,i) => PIE_COLORS[i]),
          borderWidth: 2, borderColor: '#fff'
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right' },
          tooltip: { callbacks: { label: c => ` ${toCHF(c.parsed)} CHF` } }
        }
      }
    });
  }
}

// ═══════════════════════════════════════════════════════════
// ONGLETS SIA 2 / 3 / 4 — GÉNÉRIQUE
// ═══════════════════════════════════════════════════════════
function processPhase(phaseName, phaseCode) {
  const k = parseFloat(document.getElementById('coefK_num').value) || 12;
  const isSOnly = document.getElementById('viewSOnly').checked;
  const today = new Date();
  const mmToday = yyyymm(today);
  const phaseId = phaseName.toLowerCase();

  // Filtrage par phase
  const pRows = prevRows.filter(r => {
    const id = String(r['ID projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === phaseCode;
  });
  const rRows = realRows.filter(r => {
    const id = String(r['ID projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === phaseCode;
  });

  const kpiEl = document.getElementById(`kpiSectionSIA${phaseCode.replace(/^0/,'')}`);

  // Cas sans données
  if (!pRows.length && !rRows.length) {
    if (kpiEl) {
      kpiEl.style.display = 'block';
      kpiEl.innerHTML = `<div class="empty-state"><div class="icon">📭</div>Aucune donnée pour la phase ${phaseName}</div>`;
    }
    ['table', 'pie', 'bar'].forEach(t => {
      const el = document.getElementById(`${t}SIA${phaseCode.replace(/^0/,'')}`);
      if (el) el.innerHTML = '<div class="empty-state"><div class="icon">📭</div>Aucune donnée</div>';
    });
    return;
  }

  // ── Calcul totaux ──
  const MODELS = ['Budget', 'Budget_Rev', 'AT'];
  const model = {};
  MODELS.forEach(mn => {
    const rows = pRows.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = {
      total,
      start: dates.length ? new Date(Math.min(...dates.map(d=>d.getTime()))) : null,
      end:   dates.length ? new Date(Math.max(...dates.map(d=>d.getTime()))) : null
    };
  });

  // ── Axe temporel ──
  const allDates = [];
  MODELS.forEach(mn => {
    if (model[mn].start) allDates.push(model[mn].start);
    if (model[mn].end)   allDates.push(model[mn].end);
  });
  rRows.forEach(r => { const d = getRealDate(r); if (d) allDates.push(d); });

  if (!allDates.length) return;

  const globalMin = new Date(Math.min(...allDates.map(d=>d.getTime())));
  const globalMax = new Date(Math.max(...allDates.map(d=>d.getTime())));
  const modelDates = MODELS.filter(m => model[m].start).flatMap(m => [model[m].start, model[m].end]).filter(Boolean);
  const modelMin = modelDates.length ? new Date(Math.min(...modelDates.map(d=>d.getTime()))) : globalMin;
  const modelMax = modelDates.length ? new Date(Math.max(...modelDates.map(d=>d.getTime()))) : globalMax;

  const axisMin = isSOnly ? modelMin : globalMin;
  const axisMax = isSOnly ? modelMax : globalMax;
  const axis = buildAxis(axisMin, axisMax);

  // ── Courbes S ──
  const series = {};
  MODELS.forEach(mn => {
    const [si, ei] = modelRange(pRows, mn, axis);
    series[mn] = sCurve(model[mn].total, axis, k, si, ei);
  });

  // ── Réel cumulé ──
  const realMap = new Map(axis.map(m => [m, 0]));
  let initCum = 0;
  rRows.forEach(r => {
    const d = getRealDate(r);
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

  // ── KPIs ──
  const totalBR   = model['Budget_Rev'].total;
  const totalAT   = model['AT'].total;
  const totalReal = realCum[realCum.length - 1] || 0;
  const ecartAT   = totalReal - totalAT;
  const pctAT     = totalAT > 0 ? ((totalReal / totalAT) * 100).toFixed(1) : 0;

  if (kpiEl) {
    kpiEl.style.display = 'grid';
    kpiEl.innerHTML =
      kpiCard('Budget Révisé', totalBR, PHASE_LABELS[phaseName].label, 'kpi-orange') +
      kpiCard('AT', totalAT, PHASE_LABELS[phaseName].label, 'kpi-blue2') +
      kpiCard('Réel', totalReal, 'Imputations cumulées', '') +
      kpiCard('Écart vs AT', ecartAT, `${pctAT}% avancement`, ecartAT <= 0 ? 'kpi-green' : 'kpi-red');
  }

  // ── Courbe S Chart.js ──
  const chartKey = 'chart' + phaseName;
  destroyChart(chartKey);
  const canvasId = 'chart' + phaseName;
  const ctx = document.getElementById(canvasId);
  if (ctx) {
    charts[chartKey] = new Chart(ctx, {
      type: 'line',
      data: {
        labels: axis,
        datasets: [
          {
            label: 'Budget Révisé (S)',
            data: series['Budget_Rev'],
            borderColor: GE_ORANGE, backgroundColor: GE_ORANGE + '22',
            tension: 0.4, pointRadius: 0, borderWidth: 2, fill: false, spanGaps: true
          },
          {
            label: 'AT (S)',
            data: series['AT'],
            borderColor: GE_BLUE, backgroundColor: GE_BLUE + '22',
            tension: 0.4, pointRadius: 0, borderWidth: 2, fill: false, spanGaps: true
          },
          {
            label: 'Réel cumulé',
            data: realCut,
            borderColor: GE_RED, backgroundColor: GE_RED + '22',
            tension: 0.2, pointRadius: 2, borderWidth: 2.5, fill: false, spanGaps: false
          }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top' },
          todayLine: { label: mmToday },
          tooltip: {
            callbacks: {
              label: c => ` ${c.dataset.label}: ${toCHF(c.parsed.y)} CHF`
            }
          }
        },
        scales: {
          y: {
            ticks: { callback: v => toCHF(v) + ' CHF' },
            grid: { color: '#e2e8f0' }
          },
          x: { grid: { color: '#f0f4f8' } }
        }
      }
    });
  }

  // ── Tableau budgets par ID Projet ──
  const tableEl = document.getElementById(`tableSIA${phaseCode.replace(/^0/,'')}`);
  if (tableEl) {
    const idSet = [...new Set(pRows.map(r => r['ID projet']))];
    tableEl.innerHTML = `<thead><tr>
      <th>ID Projet</th>
      <th>Nom de la tâche</th>
      <th class="num">Budget</th>
      <th class="num">Budget Révisé</th>
      <th class="num">AT</th>
      <th class="num">Réel</th>
      <th class="num">Écart BR</th>
      <th class="num">Écart AT</th>
    </tr></thead><tbody>` +
    idSet.map(id => {
      const rows = pRows.filter(r => r['ID projet'] === id);
      const nomTache = rows[0]?.['Nom de la tâche'] || id;
      const bgt = rows.filter(r => r['Modèle de prévision']==='Budget').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const br  = rows.filter(r => r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const at  = rows.filter(r => r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const reel = rRows.filter(r => String(r['ID projet']) === String(id)).reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const eBR = reel - br, eAT = reel - at;
      return `<tr>
        <td><code>${id}</code></td>
        <td>${nomTache}</td>
        <td class="num">${toCHF(bgt)}</td>
        <td class="num">${toCHF(br)}</td>
        <td class="num">${toCHF(at)}</td>
        <td class="num">${toCHF(reel)}</td>
        <td class="num ${eBR<=0?'green':'red'}">${eBR>=0?'+':''}${toCHF(eBR)}</td>
        <td class="num ${eAT<=0?'green':'red'}">${eAT>=0?'+':''}${toCHF(eAT)}</td>
      </tr>`;
    }).join('') + '</tbody>';
  }

  // ── Camembert BR ──
  const pieKey = 'pie' + phaseName + 'BR';
  destroyChart(pieKey);
  const pieBRCanvas = document.getElementById(`pieSIA${phaseCode.replace(/^0/,'')}BR`);
  const idSet2 = [...new Set(pRows.map(r => r['ID projet']))];
  const pieLabels = idSet2.map(id => {
    const rows = pRows.filter(r => r['ID projet']===id);
    return rows[0]?.['Nom de la tâche'] || id;
  });
  const pieDataBR = idSet2.map(id =>
    pRows.filter(r => r['ID projet']===id && r['Modèle de prévision']==='Budget_Rev')
         .reduce((a,r)=>a+toNum(r['Montant total du coût']),0)
  );
  if (pieBRCanvas && pieDataBR.some(v => v !== 0)) {
    charts[pieKey] = new Chart(pieBRCanvas, {
      type: 'doughnut',
      data: {
        labels: pieLabels,
        datasets: [{ data: pieDataBR, backgroundColor: PIE_COLORS, borderWidth: 2, borderColor: '#fff' }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right', labels: { font: { size: 11 } } },
          tooltip: { callbacks: { label: c => ` ${toCHF(c.parsed)} CHF` } }
        }
      }
    });
  }

  // ── Barre avancement ──
  const barKey = 'bar' + phaseName;
  destroyChart(barKey);
  const barCanvas = document.getElementById(`barSIA${phaseCode.replace(/^0/,'')}`);
  const barLabels = idSet2.map(id => {
    const rows = pRows.filter(r => r['ID projet']===id);
    const nom = rows[0]?.['Nom de la tâche'] || id;
    return nom.length > 25 ? nom.substring(0,22)+'...' : nom;
  });
  const barAT   = idSet2.map(id => pRows.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
  const barReel = idSet2.map(id => rRows.filter(r=>String(r['ID projet'])===String(id)).reduce((a,r)=>a+toNum(r['Montant total du coût']),0));
  if (barCanvas) {
    charts[barKey] = new Chart(barCanvas, {
      type: 'bar',
      data: {
        labels: barLabels,
        datasets: [
          { label: 'AT',   data: barAT,   backgroundColor: GE_BLUE2, borderRadius: 4 },
          { label: 'Réel', data: barReel, backgroundColor: GE_RED,   borderRadius: 4 }
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top' } },
        scales: {
          x: { ticks: { callback: v => toCHF(v) }, grid: { color: '#e2e8f0' } },
          y: { grid: { display: false } }
        }
      }
    });
  }
}

// ═══════════════════════════════════════════════════════════
// SIA 5 — BUDGETS (Rapport Financier)
// ═══════════════════════════════════════════════════════════
function processSIA5Budgets() {
  const k = parseFloat(document.getElementById('coefK_num').value) || 12;
  const isSOnly = document.getElementById('viewSOnly').checked;
  const today = new Date();
  const mmToday = yyyymm(today);

  // Filtrage SIA5 (code "05")
  const p5 = prevRows.filter(r => {
    const id = String(r['ID projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === '05';
  });
  const r5 = realRows.filter(r => {
    const id = String(r['ID projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === '05';
  });

  // ── Calcul modèles ──
  const MODELS = ['Budget', 'Budget_Rev', 'AT'];
  const model = {};
  MODELS.forEach(mn => {
    const rows = p5.filter(r => r['Modèle de prévision'] === mn);
    const total = rows.reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    model[mn] = {
      total,
      start: dates.length ? new Date(Math.min(...dates.map(d=>d.getTime()))) : null,
      end:   dates.length ? new Date(Math.max(...dates.map(d=>d.getTime()))) : null
    };
  });

  // ── Axe temporel ──
  const allDates = [];
  MODELS.forEach(mn => { if (model[mn].start) allDates.push(model[mn].start); if (model[mn].end) allDates.push(model[mn].end); });
  r5.forEach(r => { const d = getRealDate(r); if(d) allDates.push(d); });
  if (!allDates.length) return;

  const globalMin = new Date(Math.min(...allDates.map(d=>d.getTime())));
  const globalMax = new Date(Math.max(...allDates.map(d=>d.getTime())));
  const modelDates = MODELS.filter(m=>model[m].start).flatMap(m=>[model[m].start,model[m].end]).filter(Boolean);
  const modelMin = modelDates.length ? new Date(Math.min(...modelDates.map(d=>d.getTime()))) : globalMin;
  const modelMax = modelDates.length ? new Date(Math.max(...modelDates.map(d=>d.getTime()))) : globalMax;

  const axisMin = isSOnly ? modelMin : globalMin;
  const axisMax = isSOnly ? modelMax : globalMax;
  const axis = buildAxis(axisMin, axisMax);

  const series = {};
  MODELS.forEach(mn => {
    const [si, ei] = modelRange(p5, mn, axis);
    series[mn] = sCurve(model[mn].total, axis, k, si, ei);
  });

  // ── Réel cumulé ──
  const realMap = new Map(axis.map(m => [m, 0]));
  let initCum = 0;
  r5.forEach(r => {
    const d = getRealDate(r);
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

  // ── KPIs SIA 5 ──
  const totalBR   = model['Budget_Rev'].total;
  const totalAT   = model['AT'].total;
  const totalReal = realCum[realCum.length - 1] || 0;
  const ecartAT   = totalReal - totalAT;
  const pctAT     = totalAT > 0 ? ((totalReal / totalAT) * 100).toFixed(1) : 0;

  const kpiEl = document.getElementById('kpiSection');
  if (kpiEl) {
    kpiEl.style.display = 'grid';
    kpiEl.innerHTML =
      kpiCard('Budget Révisé SIA 5', totalBR, '6 sous-projets', 'kpi-orange') +
      kpiCard('AT SIA 5', totalAT, 'Adjudication', 'kpi-blue2') +
      kpiCard('Réel SIA 5', totalReal, 'Imputations cumulées', '') +
      kpiCard('Écart vs AT', ecartAT, `${pctAT}% avancement`, ecartAT <= 0 ? 'kpi-green' : 'kpi-red');
  }

  // ── Courbe S SIA 5 ──
  destroyChart('chart');
  const ctx = document.getElementById('chart');
  if (ctx) {
    charts.chart = new Chart(ctx, {
      type: 'line',
      data: {
        labels: axis,
        datasets: [
          {
            label: 'Budget (S)',
            data: series['Budget'],
            borderColor: GE_BLUE4, borderDash: [4,4], borderWidth: 1.5,
            pointRadius: 0, fill: false, spanGaps: true, tension: 0.4
          },
          {
            label: 'Budget Révisé (S)',
            data: series['Budget_Rev'],
            borderColor: GE_ORANGE, backgroundColor: GE_ORANGE + '15',
            tension: 0.4, pointRadius: 0, borderWidth: 2.5, fill: false, spanGaps: true
          },
          {
            label: 'AT (S)',
            data: series['AT'],
            borderColor: GE_BLUE, backgroundColor: GE_BLUE + '15',
            tension: 0.4, pointRadius: 0, borderWidth: 2.5, fill: false, spanGaps: true
          },
          {
            label: 'Réel cumulé',
            data: realCut,
            borderColor: GE_RED, backgroundColor: GE_RED + '15',
            tension: 0.2, pointRadius: 2, borderWidth: 2.5, fill: false, spanGaps: false
          }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'top' },
          todayLine: { label: mmToday },
          tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${toCHF(c.parsed.y)} CHF` } }
        },
        scales: {
          y: { ticks: { callback: v => toCHF(v) + ' CHF' }, grid: { color: '#e2e8f0' } },
          x: { grid: { color: '#f0f4f8' } }
        }
      }
    });
  }

  // ── Tableau récapitulatif par sous-projet ──
  const tablePrev = document.getElementById('tablePreview');
  if (tablePrev) {
    tablePrev.innerHTML = `<thead><tr>
      <th>Sous-projet</th>
      <th class="num">Budget</th>
      <th class="num">Budget Révisé</th>
      <th class="num">AT</th>
    </tr></thead><tbody>` +
    PROJECT_ORDER.map(id => {
      const rows = p5.filter(r => r['ID projet'] === id);
      const bgt = rows.filter(r=>r['Modèle de prévision']==='Budget').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const br  = rows.filter(r=>r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const at  = rows.filter(r=>r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      return `<tr>
        <td class="bold">${getPN(id)}</td>
        <td class="num">${toCHF(bgt)}</td>
        <td class="num">${toCHF(br)}</td>
        <td class="num">${toCHF(at)}</td>
      </tr>`;
    }).join('') +
    `<tr class="bold" style="background:#edf2f7;border-top:2px solid #163a5f">
      <td>TOTAL SIA 5</td>
      <td class="num">${toCHF(model['Budget'].total)}</td>
      <td class="num">${toCHF(totalBR)}</td>
      <td class="num">${toCHF(totalAT)}</td>
    </tr></tbody>`;
  }

  // ── Camembert BR et AT ──
  const brData = PROJECT_ORDER.map(id =>
    p5.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0)
  );
  const atData = PROJECT_ORDER.map(id =>
    p5.filter(r=>r['ID projet']===id&&r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0)
  );
  const pieLabels5 = PROJECT_ORDER.map(getPN);

  destroyChart('chartPieBudgetRev');
  const ctxPieBR = document.getElementById('chartPieBudgetRev');
  if (ctxPieBR && brData.some(v=>v!==0)) {
    charts.chartPieBudgetRev = new Chart(ctxPieBR, {
      type: 'doughnut',
      data: { labels: pieLabels5, datasets: [{ data: brData, backgroundColor: PIE_COLORS, borderWidth: 2, borderColor: '#fff' }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'right' }, tooltip: { callbacks: { label: c => ` ${toCHF(c.parsed)} CHF` } } }
      }
    });
  }

  destroyChart('chartPieAT');
  const ctxPieAT = document.getElementById('chartPieAT');
  if (ctxPieAT && atData.some(v=>v!==0)) {
    charts.chartPieAT = new Chart(ctxPieAT, {
      type: 'doughnut',
      data: { labels: pieLabels5, datasets: [{ data: atData, backgroundColor: PIE_COLORS, borderWidth: 2, borderColor: '#fff' }] },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'right' }, tooltip: { callbacks: { label: c => ` ${toCHF(c.parsed)} CHF` } } }
      }
    });
  }

  // ── Graphique barres avancement SIA 5 ──
  const realBySub = PROJECT_ORDER.map(id =>
    r5.filter(r => String(r['ID projet'])===id).reduce((a,r)=>a+toNum(r['Montant total du coût']),0)
  );

  destroyChart('chartBar');
  const ctxBar = document.getElementById('chartBar');
  if (ctxBar) {
    charts.chartBar = new Chart(ctxBar, {
      type: 'bar',
      data: {
        labels: PROJECT_ORDER.map(getPN),
        datasets: [
          { label: 'Budget Révisé', data: brData,    backgroundColor: GE_ORANGE, borderRadius: 5 },
          { label: 'AT',            data: atData,    backgroundColor: GE_BLUE2,  borderRadius: 5 },
          { label: 'Réel',          data: realBySub, backgroundColor: GE_RED,    borderRadius: 5 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top' } },
        scales: {
          y: { ticks: { callback: v => toCHF(v) + ' CHF' }, grid: { color: '#e2e8f0' } },
          x: { grid: { display: false } }
        }
      }
    });
  }

  // ── Tableau comparatif détaillé ──
  const tableCmp = document.getElementById('tableCompare');
  if (tableCmp) {
    tableCmp.innerHTML = `<thead><tr>
      <th>Sous-projet</th>
      <th class="num">Budget</th>
      <th class="num">Budget Révisé</th>
      <th class="num">AT</th>
      <th class="num">Réel</th>
      <th class="num">Écart BR</th>
      <th class="num">Écart AT</th>
      <th class="num">% BR</th>
    </tr></thead><tbody>` +
    PROJECT_ORDER.map((id, i) => {
      const rows = p5.filter(r => r['ID projet'] === id);
      const bgt  = rows.filter(r=>r['Modèle de prévision']==='Budget').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const br   = rows.filter(r=>r['Modèle de prévision']==='Budget_Rev').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const at   = rows.filter(r=>r['Modèle de prévision']==='AT').reduce((a,r)=>a+toNum(r['Montant total du coût']),0);
      const reel = realBySub[i];
      const eBR  = reel - br, eAT = reel - at;
      const pct  = br > 0 ? ((reel/br)*100).toFixed(1) : 0;
      return `<tr>
        <td class="bold">${getPN(id)}</td>
        <td class="num">${toCHF(bgt)}</td>
        <td class="num">${toCHF(br)}</td>
        <td class="num">${toCHF(at)}</td>
        <td class="num">${toCHF(reel)}</td>
        <td class="num ${eBR<=0?'green':'red'}">${eBR>=0?'+':''}${toCHF(eBR)}</td>
        <td class="num ${eAT<=0?'green':'red'}">${eAT>=0?'+':''}${toCHF(eAT)}</td>
        <td class="num">${pct}%</td>
      </tr>`;
    }).join('') +
    `<tr class="bold" style="background:#edf2f7;border-top:2px solid #163a5f">
      <td>TOTAL</td>
      <td class="num">${toCHF(model['Budget'].total)}</td>
      <td class="num">${toCHF(totalBR)}</td>
      <td class="num">${toCHF(totalAT)}</td>
      <td class="num">${toCHF(totalReal)}</td>
      <td class="num ${totalReal-totalBR<=0?'green':'red'}">${totalReal-totalBR>=0?'+':''}${toCHF(totalReal-totalBR)}</td>
      <td class="num ${ecartAT<=0?'green':'red'}">${ecartAT>=0?'+':''}${toCHF(ecartAT)}</td>
      <td class="num">${totalBR>0?((totalReal/totalBR)*100).toFixed(1):0}%</td>
    </tr></tbody>`;
  }
}

// ═══════════════════════════════════════════════════════════
// SIA 5 — IMPUTATIONS (Honoraires + Ressources)
// ═══════════════════════════════════════════════════════════
function processImputations() {
  const r5 = realRows.filter(r => {
    const id = String(r['ID projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === '05';
  });

  // ── Filtrage lignes honoraires ──
  const honoRows = r5.filter(r => {
    const nom = String(r['Nom de la tâche'] || '').toLowerCase();
    return nom.includes('honoraire');
  });

  // ── Catégorisation par métier ──
  const METIERS = [
    { key: 'Réseau CAD',        pattern: /réseau\s*cad/i },
    { key: 'Sous-stations',     pattern: /sous.station/i },
    { key: 'Centrale',          pattern: /centrale/i },
    { key: 'Bâtiment',          pattern: /bâtiment|batiment/i },
    { key: 'Général - Direction', pattern: /général|general|direction/i },
    { key: 'Participations',    pattern: /participation/i }
  ];

  const metiersData = {};
  METIERS.forEach(m => { metiersData[m.key] = { int: 0, ext: 0 }; });
  metiersData['Autres'] = { int: 0, ext: 0 };

  honoRows.forEach(r => {
    const nom = String(r['Nom de la tâche'] || '');
    const montant = toNum(r['Montant total du coût']);
    const isExt = Boolean(r['Nom du fournisseur'] && String(r['Nom du fournisseur']).trim()) ||
                  nom.toLowerCase().includes('externe');

    let matched = false;
    for (const m of METIERS) {
      if (m.pattern.test(nom)) {
        metiersData[m.key][isExt ? 'ext' : 'int'] += montant;
        matched = true; break;
      }
    }
    if (!matched) metiersData['Autres'][isExt ? 'ext' : 'int'] += montant;
  });

  // KPIs honoraires
  const totalInt = Object.values(metiersData).reduce((a, v) => a + v.int, 0);
  const totalExt = Object.values(metiersData).reduce((a, v) => a + v.ext, 0);

  const kpiHono = document.getElementById('kpiSectionHonoraires');
  if (kpiHono) {
    kpiHono.style.display = 'grid';
    kpiHono.innerHTML =
      kpiCard('Honoraires Internes', totalInt, 'Ressources Groupe E', '') +
      kpiCard('Honoraires Externes', totalExt, 'Fournisseurs tiers', 'kpi-orange') +
      kpiCard('Total Honoraires', totalInt + totalExt, 'SIA 5 toutes ressources', 'kpi-blue2') +
      kpiCard('Ratio Ext/Total', totalInt + totalExt > 0 ? (totalExt/(totalInt+totalExt)*100).toFixed(1) + '%' : '0%', 'Part externe', '');
  }
  // Correction : la 4e carte est un pourcentage, pas un CHF
  if (kpiHono) {
    const cards = kpiHono.querySelectorAll('.kpi-card');
    if (cards[3]) {
      cards[3].querySelector('.kpi-value').textContent =
        totalInt + totalExt > 0 ? ((totalExt/(totalInt+totalExt))*100).toFixed(1) + '%' : '0%';
    }
  }

  // ── Graphique barres honoraires par métier ──
  const allMetiers = Object.keys(metiersData).filter(k => metiersData[k].int + metiersData[k].ext > 0);
  destroyChart('chartHonoMetier');
  const ctxHono = document.getElementById('chartHonoMetier');
  if (ctxHono && allMetiers.length) {
    charts.chartHonoMetier = new Chart(ctxHono, {
      type: 'bar',
      data: {
        labels: allMetiers,
        datasets: [
          { label: 'Internes', data: allMetiers.map(k=>metiersData[k].int), backgroundColor: GE_BLUE3, borderRadius: 5 },
          { label: 'Externes', data: allMetiers.map(k=>metiersData[k].ext), backgroundColor: GE_ORANGE, borderRadius: 5 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top' } },
        scales: {
          y: { ticks: { callback: v => toCHF(v) + ' CHF' }, grid: { color: '#e2e8f0' } },
          x: { grid: { display: false } }
        }
      }
    });
  }

  // ── Tableau ressources internes / externes ──
  // Agrégation par nom de ressource
  const resMap = new Map();
  r5.forEach(r => {
    const nom = String(r['Nom de la tâche'] || '');
    const ressource = String(r['Nom de la ressource'] || r['Ressource'] || '');
    const fournisseur = String(r['Nom du fournisseur'] || '');
    const isExt = Boolean(fournisseur.trim()) || nom.toLowerCase().includes('externe');
    const key = ressource || fournisseur || 'Inconnu';
    const montant = toNum(r['Montant total du coût']);
    if (!resMap.has(key)) resMap.set(key, { ressource: key, fournisseur, isExt, total: 0, tache: nom });
    resMap.get(key).total += montant;
  });

  const ressources = [...resMap.values()].sort((a,b) => b.total - a.total);

  const tableRess = document.getElementById('tableRessources');
  if (tableRess && ressources.length) {
    tableRess.innerHTML = `<thead><tr>
      <th>Type</th>
      <th>Ressource / Fournisseur</th>
      <th>Tâche principale</th>
      <th class="num">Total CHF</th>
    </tr></thead><tbody>` +
    ressources.map(r => `<tr class="${r.isExt ? 'tr-external' : 'tr-internal'}">
      <td><span class="${r.isExt ? 'badge-ext' : 'badge-int'}">${r.isExt ? 'EXT' : 'INT'}</span></td>
      <td>${r.ressource}</td>
      <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis">${r.tache}</td>
      <td class="num">${toCHF(r.total)}</td>
    </tr>`).join('') +
    '</tbody>';
  }
}

// ═══════════════════════════════════════════════════════════
// SIA 5 — PLANNING & GANTT
// ═══════════════════════════════════════════════════════════
function processPlanning() {
  const p5 = prevRows.filter(r => {
    const id = String(r['ID projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === '05';
  });
  const r5 = realRows.filter(r => {
    const id = String(r['ID projet'] || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === '05';
  });

  // Calcul plages par sous-projet
  const planning = {};
  PROJECT_ORDER.forEach(id => {
    const pRows = p5.filter(r => r['ID projet'] === id);
    const rRows = r5.filter(r => String(r['ID projet']) === id);

    const getBudgetDates = mn => {
      const rows = pRows.filter(r => r['Modèle de prévision'] === mn);
      const dates = rows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
      return dates.length ? {
        start: new Date(Math.min(...dates.map(d=>d.getTime()))),
        end:   new Date(Math.max(...dates.map(d=>d.getTime())))
      } : null;
    };

    const realDates = rRows.map(r => getRealDate(r)).filter(Boolean);

    // Dates manuelles si définies
    const manual = manualPlanningDates[id];
    const realStart = manual?.start || (realDates.length ? new Date(Math.min(...realDates.map(d=>d.getTime()))) : null);
    const realEnd   = manual?.end   || (realDates.length ? new Date(Math.max(...realDates.map(d=>d.getTime()))) : null);

    planning[id] = {
      name:      getPN(id),
      budget:    getBudgetDates('Budget'),
      budgetRev: getBudgetDates('Budget_Rev'),
      at:        getBudgetDates('AT'),
      reel:      realStart && realEnd ? { start: realStart, end: realEnd } : null
    };
  });

  // ── Formulaire inputs planning ──
  const planInputsEl = document.getElementById('planningInputs');
  if (planInputsEl) {
    planInputsEl.innerHTML = PROJECT_ORDER.map(id => {
      const p = planning[id];
      const rs = p.reel?.start ? p.reel.start.toISOString().split('T')[0] : '';
      const re = p.reel?.end   ? p.reel.end.toISOString().split('T')[0]   : '';
      return `<div class="planning-row">
        <label>${p.name}</label>
        <span style="font-size:.75rem;color:#718096">Réel Début</span>
        <input type="date" id="rs_${id}" value="${rs}">
        <span style="font-size:.75rem;color:#718096">Fin</span>
        <input type="date" id="re_${id}" value="${re}">
      </div>`;
    }).join('');
  }

  // ── Gantt ──
  // On construit les données pour un chart de type 'bar' horizontal flottant
  const today = new Date();
  const allGanttDates = [];
  PROJECT_ORDER.forEach(id => {
    const p = planning[id];
    ['budget','budgetRev','at','reel'].forEach(k => {
      if (p[k]) { allGanttDates.push(p[k].start); allGanttDates.push(p[k].end); }
    });
  });
  if (!allGanttDates.length) return;
  allGanttDates.push(today);

  const ganttMin = new Date(Math.min(...allGanttDates.map(d=>d.getTime())));
  const ganttMax = new Date(Math.max(...allGanttDates.map(d=>d.getTime())));
  const toMs = d => d ? d.getTime() : null;
  const names = PROJECT_ORDER.map(id => planning[id].name);

  // Convertir en jours depuis ganttMin pour l'axe
  const toDays = d => d ? Math.round((d.getTime() - ganttMin.getTime()) / 86400000) : null;
  const totalDays = Math.round((ganttMax.getTime() - ganttMin.getTime()) / 86400000) + 10;

  // Génère dataset flottant [start, end] en jours
  const makeDataset = (colorFill, colorBorder, label, key) => ({
    label,
    data: PROJECT_ORDER.map(id => {
      const p = planning[id][key];
      if (!p) return null;
      return [toDays(p.start), toDays(p.end)];
    }),
    backgroundColor: colorFill,
    borderColor: colorBorder,
    borderWidth: 1.5,
    borderRadius: 4,
    barThickness: 14
  });

  destroyChart('chartGantt');
  const ctxGantt = document.getElementById('chartGantt');
  if (ctxGantt) {
    charts.chartGantt = new Chart(ctxGantt, {
      type: 'bar',
      data: {
        labels: names,
        datasets: [
          makeDataset(GE_BLUE4 + '88', GE_BLUE4, 'Budget', 'budget'),
          makeDataset(GE_ORANGE + '88', GE_ORANGE, 'Budget Révisé', 'budgetRev'),
          makeDataset(GE_BLUE2 + '88', GE_BLUE2, 'AT', 'at'),
          makeDataset(GE_RED + '88', GE_RED, 'Réel', 'reel')
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: true, maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top' },
          tooltip: {
            callbacks: {
              label: c => {
                if (!c.raw) return '';
                const s = new Date(ganttMin.getTime() + c.raw[0] * 86400000);
                const e = new Date(ganttMin.getTime() + c.raw[1] * 86400000);
                return ` ${c.dataset.label}: ${s.toLocaleDateString('fr-CH')} → ${e.toLocaleDateString('fr-CH')}`;
              }
            }
          }
        },
        scales: {
          x: {
            min: 0, max: totalDays,
            ticks: {
              callback: v => {
                const d = new Date(ganttMin.getTime() + v * 86400000);
                return d.toLocaleDateString('fr-CH', { month: 'short', year: '2-digit' });
              },
              maxTicksLimit: 12
            },
            grid: { color: '#e2e8f0' }
          },
          y: { grid: { display: false } }
        }
      }
    });
  }

  // ── Tableau planning ──
  const tablePlan = document.getElementById('tablePlanning');
  if (tablePlan) {
    const fmt = d => d ? d.toLocaleDateString('fr-CH') : '—';
    tablePlan.innerHTML = `<thead><tr>
      <th>Sous-projet</th>
      <th>Budget Début</th><th>Budget Fin</th>
      <th>Budget Rév. Début</th><th>Budget Rév. Fin</th>
      <th>AT Début</th><th>AT Fin</th>
      <th>Réel Début</th><th>Réel Fin</th>
    </tr></thead><tbody>` +
    PROJECT_ORDER.map(id => {
      const p = planning[id];
      return `<tr>
        <td class="bold">${p.name}</td>
        <td>${fmt(p.budget?.start)}</td><td>${fmt(p.budget?.end)}</td>
        <td>${fmt(p.budgetRev?.start)}</td><td>${fmt(p.budgetRev?.end)}</td>
        <td>${fmt(p.at?.start)}</td><td>${fmt(p.at?.end)}</td>
        <td>${fmt(p.reel?.start)}</td><td>${fmt(p.reel?.end)}</td>
      </tr>`;
    }).join('') + '</tbody>';
  }
}

// ═══════════════════════════════════════════════════════════
// EXPORT PDF
// ═══════════════════════════════════════════════════════════
async function generatePDF() {
  const { jsPDF } = window.jspdf;
  const scope = document.getElementById('pdfScope').value;

  setStatus('⏳ Génération PDF en cours...', '');
  showLoader(true);

  const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const W = 210, H = 297, M = 12;
  let y = M;

  const addHeader = (title) => {
    pdf.setFillColor(22, 58, 95);
    pdf.rect(0, 0, W, 22, 'F');
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(13);
    pdf.setFont('helvetica', 'bold');
    pdf.text('📊 Dashboard CAD — Groupe E Celsius', M, 10);
    pdf.setFontSize(9);
    pdf.setFont('helvetica', 'normal');
    pdf.text(title, M, 17);
    pdf.setTextColor(0, 0, 0);
    y = 28;
  };

  const addCanvas = async (canvasId, label) => {
    const el = document.getElementById(canvasId);
    if (!el) return;
    try {
      const imgData = el.toDataURL('image/png', 0.92);
      const ratio = el.height / el.width;
      const imgW = W - 2*M;
      const imgH = Math.min(imgW * ratio, 80);
      if (y + imgH + 15 > H - M) { pdf.addPage(); addHeader(label); }
      pdf.setFontSize(9); pdf.setFont('helvetica', 'bold');
      pdf.setTextColor(22, 58, 95);
      pdf.text(label, M, y + 5);
      y += 8;
      pdf.addImage(imgData, 'PNG', M, y, imgW, imgH);
      y += imgH + 8;
    } catch(e) { console.warn('PDF: canvas skip', canvasId, e); }
  };

  const scopeMap = {
    recap: { title: 'Récapitulatif Global', tabs: ['tab-recap'], charts: ['chartRecapBar','chartRecapPie'] },
    sia2:  { title: 'Phase SIA 2 — Avant-Projet',          tabs: ['tab-sia2'], charts: ['chartSIA2','pieSIA2BR','barSIA2'] },
    sia3:  { title: 'Phase SIA 3 — Projet de Réalisation',  tabs: ['tab-sia3'], charts: ['chartSIA3','pieSIA3BR','barSIA3'] },
    sia4:  { title: 'Phase SIA 4 — Appels d\'Offres',       tabs: ['tab-sia4'], charts: ['chartSIA4','pieSIA4BR','barSIA4'] },
    sia5:  { title: 'Phase SIA 5 — Réalisation',            tabs: ['tab-sia5'], charts: ['chart','chartPieBudgetRev','chartPieAT','chartBar','chartHonoMetier','chartGantt'] }
  };

  const scopes = scope === 'all' ? Object.keys(scopeMap).filter(k=>k!=='all') : [scope];

  for (let i = 0; i < scopes.length; i++) {
    const sc = scopeMap[scopes[i]];
    if (!sc) continue;
    if (i > 0) pdf.addPage();
    addHeader(sc.title);
    pdf.setFontSize(11); pdf.setFont('helvetica', 'bold');
    pdf.text(sc.title, M, y);
    pdf.setFontSize(8); pdf.setFont('helvetica', 'normal');
    pdf.setTextColor(120, 120, 120);
    pdf.text(`Généré le ${new Date().toLocaleDateString('fr-CH')} — Projet CL150096`, M, y + 5);
    pdf.setTextColor(0,0,0);
    y += 12;
    for (const cid of sc.charts) {
      await addCanvas(cid, cid);
    }
  }

  pdf.save(`Dashboard_CAD_${scope}_${new Date().toISOString().split('T')[0]}.pdf`);
  showLoader(false);
  setStatus('✅ PDF généré avec succès', 'ok');
}

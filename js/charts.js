/* ============================================================
   Dashboard CAD — Groupe E Celsius
   charts.js — Graphiques & Courbes S v5.3
   ============================================================ */

// ============================================================
// PALETTE GROUPE E
// ============================================================
const GE_BLUE   = '#163a5f';
const GE_BLUE2  = '#1e4d8c';
const GE_BLUE3  = '#2e6db4';
const GE_BLUE4  = '#5b9bd5';
const GE_ORANGE = '#e8873a';
const GE_RED    = '#c0392b';
const GE_GREEN  = '#27ae60';

const PIE_COLORS = [
  '#2e6db4','#e8873a','#27ae60','#8e44ad',
  '#c0392b','#16a085','#d4ac0d','#2980b9',
  '#e74c3c','#1abc9c','#f39c12','#7f8c8d'
];

// ============================================================
// COURBE S LOGISTIQUE
// ============================================================

function sCurve(total, axis, k, sIdx, eIdx) {
  const N = axis.length;
  const cum = Array(N).fill(null);
  const L = eIdx - sIdx + 1;
  
  if (L <= 0 || total === 0 || sIdx < 0 || eIdx < 0) return cum;
  
  const s0 = 1/(1+Math.exp(-k*(0-0.5)));
  const s1 = 1/(1+Math.exp(-k*(1-0.5)));
  
  for (let j = 0; j < L; j++) {
    const i = sIdx + j;
    if (i >= N) break;
    const x = j/(L-1||1);
    const sx = 1/(1+Math.exp(-k*(x-0.5)));
    cum[i] = ((sx-s0)/(s1-s0)) * total;
  }
  return cum;
}

function modelRange(rows, modelName, axis) {
  const filtered = rows.filter(r => col(r, 'Modèle de prévision', 'Modele de prevision', 'Modèle') === modelName);
  const dates = filtered.flatMap(r => [
    toDate(col(r, 'Date de début', 'Date de debut', 'Date début')),
    toDate(col(r, 'Date de fin', 'Date fin'))
  ]).filter(Boolean);
  
  if (!dates.length) return [-1, -1];
  
  const minD = new Date(Math.min(...dates));
  const maxD = new Date(Math.max(...dates));
  const mmMin = yyyymm(firstOfMonth(minD));
  const mmMax = yyyymm(firstOfMonth(maxD));
  
  let sIdx = axis.indexOf(mmMin);
  let eIdx = axis.indexOf(mmMax);
  
  if (sIdx < 0) sIdx = axis.findIndex(m => m >= mmMin);
  if (eIdx < 0) {
    for (let i = axis.length-1; i >= 0; i--) {
      if (axis[i] <= mmMax) { eIdx = i; break; }
    }
  }
  
  if (sIdx < 0) sIdx = 0;
  if (eIdx < 0) eIdx = axis.length-1;
  
  return [sIdx, eIdx];
}

// ============================================================
// BUILD S-CURVES (CORRECTION v5.3 : utilisation col())
// ============================================================

function buildSCurves(pRows, rRows) {
  const k = parseFloat(document.getElementById('coefK_num').value) || 12;
  const isSOnly = document.getElementById('viewSOnly').checked;
  const MODELS = ['Budget','Budget_Rev','AT'];
  const model = {};

  MODELS.forEach(mn => {
    const rows = pRows.filter(r => col(r, 'Modèle de prévision', 'Modele de prevision') === mn);
    const total = rows.reduce((a,r) => a + toNum(col(r, 'Montant total du coût', 'Montant total du cout', 'Montant') || 0), 0);
    const dates = rows.flatMap(r => [
      toDate(col(r, 'Date de début', 'Date de debut')),
      toDate(col(r, 'Date de fin', 'Date fin'))
    ]).filter(Boolean);
    
    model[mn] = {
      total,
      start: dates.length ? new Date(Math.min(...dates)) : null,
      end:   dates.length ? new Date(Math.max(...dates)) : null
    };
  });

  const realDates = rRows.map(r => getDateReal(r)).filter(Boolean);
  const allDates = [];
  MODELS.forEach(mn => {
    if(model[mn].start) allDates.push(model[mn].start);
    if(model[mn].end) allDates.push(model[mn].end);
  });
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
    const filtered = pRows.filter(r => col(r, 'Modèle de prévision', 'Modele de prevision') === mn);
    const [si, ei] = modelRange(filtered.length ? filtered : pRows, mn, axis);
    series[mn] = sCurve(model[mn].total, axis, k, si, ei);
  });

  const today = new Date();
  const mmToday = yyyymm(today);
  const realMap = new Map(axis.map(m => [m, 0]));
  let initCum = 0;
  
  rRows.forEach(r => {
    const d = getDateReal(r);
    if (!d) return;
    const mm = yyyymm(firstOfMonth(d));
    if (mm > mmToday) return; // Exclure futures
    
    const amt = toNum(col(r, 'Montant total du coût', 'Montant total du cout', 'Montant') || 0);
    if (realMap.has(mm)) {
      realMap.set(mm, realMap.get(mm) + amt);
    } else if (isSOnly && mm < axis[0]) {
      initCum += amt;
    }
  });

  let acc = initCum;
  const realMonth = axis.map(m => realMap.get(m) || 0);
  const realCum   = realMonth.map(v => { acc += v; return acc; });
  const realCut   = realCum.map((v,i) => axis[i] > mmToday ? null : v);

  return { axis, series, realMonth, realCum, realCut, model, mmToday, initCum };
}

// ============================================================
// DESSIN COURBE S (Chart.js)
// ============================================================

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
// TABLEAU MENSUEL (CORRECTION v5.3 : suppression double cumul)
// ============================================================

function renderMonthlyTable(tableId, sData) {
  if (!sData) return;
  const { axis, series, realMonth, realCum, mmToday, initCum } = sData;
  const tbl = document.getElementById(tableId);
  if (!tbl) return;

  const brCum  = series['Budget_Rev'] || [];
  const atCum  = series['AT'] || [];

  const getMonthly = cum => cum.map((v,i) => {
    if (v === null) return 0;
    const prev = i > 0 ? (cum[i-1] || 0) : 0;
    return Math.max(0, v - prev);
  });
  const brMonthly = getMonthly(brCum);

  // CORRECTION v5.3 : utiliser directement realCum (déjà calculé dans buildSCurves)
  // Suppression du recalcul local qui causait le double cumul

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
    const reelC = realCum[i] || 0; // <-- Utilisation directe (pas de recalcul)
    const ecM   = reelM - brM;
    const ecC   = reelC - brC;
    const isPast = mm <= mmToday;

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
// PLUGIN TODAY LINE
// ============================================================

const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(ch, _args, opts) {
    if (!opts || !opts.label) return;
    const { ctx, chartArea: { top, bottom }, scales: { x } } = ch;
    if (!x) return;
    const labels = ch.data.labels;
    if (!labels || !labels.length) return;

    const targetMM = opts.label; // Format yyyymm (ex: "2026-02")

    let bestIdx = -1;
    labels.forEach((l, i) => {
      const labelStr = String(l);
      if (labelStr === targetMM) { bestIdx = i; return; }
      
      const monthMap = {
        'jan':  '01', 'fév':  '02', 'fev':  '02', 'mar':  '03',
        'avr':  '04', 'mai':  '05', 'juin': '06', 'juil': '07',
        'août': '08', 'aout': '08', 'sep':  '09', 'oct':  '10',
        'nov':  '11', 'déc':  '12', 'dec':  '12'
      };
      const yearMatch  = labelStr.match(/(\d{4})/);
      const lowerLabel = labelStr.toLowerCase();
      if (!yearMatch) return;
      const year = yearMatch[1];
      let month = null;
      for (const [abbr, num] of Object.entries(monthMap)) {
        if (lowerLabel.startsWith(abbr) || lowerLabel.includes(' ' + abbr) || lowerLabel.includes(abbr + ' ')) {
          month = num;
          break;
        }
      }
      if (month && `${year}-${month}` === targetMM) bestIdx = i;
    });

    if (bestIdx < 0) return;
    const xp = x.getPixelForValue(bestIdx);
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
    ctx.fillText("Aujourd'hui", xp, top - 5);
    ctx.restore();
  }
};

// Enregistrer le plugin au chargement
if (typeof Chart !== 'undefined') {
  Chart.register(todayLinePlugin);
}

// ============================================================
// DESTROY CHART
// ============================================================

function destroyChart(key) {
  if (charts[key]) {
    try { charts[key].destroy(); } catch(e){}
    delete charts[key];
  }
}

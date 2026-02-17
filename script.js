// ==================== VARIABLES GLOBALES ====================
let prevRows = [], realRows = [];
let chart, chartBar, chartPieBudgetRev, chartPieAT;
let chartHonoMetier, chartSyntheseMetier, chartGantt;

const prevInput = document.getElementById('prevFile');
const realInput = document.getElementById('realFile');
const kSlider = document.getElementById('coefK');
const kNumber = document.getElementById('coefK_num');
const statusEl = document.getElementById('status');
const chartEl = document.getElementById('chart');
const chartBarEl = document.getElementById('chartBar');
const chartPieBudgetRevEl = document.getElementById('chartPieBudgetRev');
const chartPieATEl = document.getElementById('chartPieAT');
const tablePrev = document.getElementById('tablePreview');
const tableCmp = document.getElementById('tableCompare');
const viewGlobalEl = document.getElementById('viewGlobal');
const viewSOnlyEl = document.getElementById('viewSOnly');
const btnExportPNG = document.getElementById('btnExportPNG');
const btnExportTablesPNG = document.getElementById('btnExportTablesPNG');
const btnExportXLSX = document.getElementById('btnExportXLSX');
const btnRefresh = document.getElementById('btnRefresh');
const kpiSection = document.getElementById('kpiSection');

// Onglet 2
const chartHonoMetierEl = document.getElementById('chartHonoMetier');
const tableRessources = document.getElementById('tableRessources');
const chartSyntheseMetierEl = document.getElementById('chartSyntheseMetier');
const kpiSectionHonoraires = document.getElementById('kpiSectionHonoraires');

// Onglet 3
const chartGanttEl = document.getElementById('chartGantt');
const tablePlanning = document.getElementById('tablePlanning');
const btnExportGanttPNG = document.getElementById('btnExportGanttPNG');

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

// ==================== GESTION ONGLETS ====================
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
    
    // Reprocess pour onglet 2 et 3
    if (btn.dataset.tab === 'tab2' && prevRows.length && realRows.length) {
      processTab2();
    }
    if (btn.dataset.tab === 'tab3' && prevRows.length && realRows.length) {
      processTab3();
    }
  });
});

// ==================== UTILITAIRES ====================
function readExcel(file, cb) {
  const r = new FileReader();
  r.onload = e => {
    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { raw: true });
    cb(rows);
  };
  r.readAsArrayBuffer(file);
}

function toNum(v) {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string') return 0;
  return parseFloat(v.replace(/\u00A0/g, ' ').replace(/[ '\u202F']/g, '').replace(',', '.')) || 0;
}

function toCHF(n) {
  return (n || 0).toLocaleString('fr-CH', { minimumFractionDigits: 0 });
}

function toDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + v * 86400000);
  if (typeof v === 'string') {
    let d = new Date(v);
    if (!isNaN(d)) return d;
    const m = v.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
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

function sCurveCumulative(total, axis, k, range) {
  const N = axis.length;
  const cum = Array(N).fill(null);
  const month = Array(N).fill(0);
  const s0 = 1 / (1 + Math.exp(-k * (0 - 0.5)));
  const s1 = 1 / (1 + Math.exp(-k * (1 - 0.5)));
  const [s, e] = range;
  const L = e - s + 1;
  for (let j = 0; j < L; j++) {
    const i = s + j;
    if (i >= N) break;
    const x = j / (L - 1 || 1);
    const sx = 1 / (1 + Math.exp(-k * (x - 0.5)));
    const norm = (sx - s0) / (s1 - s0);
    cum[i] = norm * total;
    month[i] = (j === 0) ? cum[i] : (cum[i] - cum[i - 1]);
  }
  return { cum, month };
}

// Plugin ligne "Aujourd'hui"
const todayLinePlugin = {
  id: 'todayLine',
  afterDatasetsDraw(chart, args, opts) {
    const { ctx, chartArea: { top, bottom }, scales: { x } } = chart;
    if (!opts.label) return;
    const labels = chart.data.labels;
    const toTime = s => {
      const [yy, mm] = s.split('-');
      return new Date(parseInt(yy), parseInt(mm) - 1, 1).getTime();
    };
    const t = toTime(opts.label);
    let bestI = 0, bestD = Infinity;
    labels.forEach((l, i) => {
      const d = Math.abs(toTime(l) - t);
      if (d < bestD) { bestD = d; bestI = i; }
    });
    const xPos = x.getPixelForValue(bestI);
    ctx.save();
    ctx.strokeStyle = opts.color || 'red';
    ctx.lineWidth = opts.width || 2;
    ctx.setLineDash([5, 5]);
    ctx.beginPath();
    ctx.moveTo(xPos, top);
    ctx.lineTo(xPos, bottom);
    ctx.stroke();
    ctx.setLineDash([]);
    if (opts.text) {
      ctx.fillStyle = opts.color || 'red';
      ctx.font = 'bold 11px Arial';
      ctx.textAlign = 'center';
      ctx.fillText(opts.text, xPos, top - 5);
    }
    ctx.restore();
  }
};

// ==================== EVENTS FICHIERS ====================
prevInput.addEventListener('change', e => {
  if (e.target.files[0]) {
    readExcel(e.target.files[0], rows => {
      prevRows = rows;
      maybeProcess();
    });
  }
});

realInput.addEventListener('change', e => {
  if (e.target.files[0]) {
    readExcel(e.target.files[0], rows => {
      realRows = rows;
      maybeProcess();
    });
  }
});

[kSlider, kNumber, viewGlobalEl, viewSOnlyEl].forEach(el => {
  el.addEventListener('input', () => {
    kSlider.value = kNumber.value = el.value || kNumber.value;
    if (prevRows.length && realRows.length) processTab1();
  });
});

btnRefresh.addEventListener('click', () => {
  if (prevRows.length && realRows.length) processTab1();
});

function maybeProcess() {
  if (prevRows.length && realRows.length) {
    processTab1();
  }
}

// ==================== ONGLET 1: RAPPORT FINANCIER ====================
function processTab1() {
  statusEl.textContent = '⏳ Calcul en cours...';
  statusEl.className = 'status';

  const k = parseFloat(kNumber.value);
  const viewMode = viewGlobalEl.checked ? 'global' : 's-only';

  // Séries Budget, Budget_Rev, AT
  const series = {};
  ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
    const total = prevRows
      .filter(r => r['Modèle de prévision'] === mode)
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    const dates = prevRows
      .filter(r => r['Modèle de prévision'] === mode)
      .flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])])
      .filter(Boolean);

    const minD = dates.length ? new Date(Math.min(...dates)) : new Date();
    const maxD = dates.length ? new Date(Math.max(...dates)) : new Date();

    series[mode] = { total, minD, maxD };
  });

  // Axe global
  const allDates = Object.values(series).flatMap(s => [s.minD, s.maxD]);
  const globalMinD = new Date(Math.min(...allDates));
  const globalMaxD = new Date(Math.max(...allDates));

  // Réels
  const realDates = realRows
    .map(r => toDate(r['Date du projet']))
    .filter(Boolean);
  const realMinD = realDates.length ? new Date(Math.min(...realDates)) : globalMinD;
  const realMaxD = realDates.length ? new Date(Math.max(...realDates)) : globalMaxD;

  const axisMinD = new Date(Math.min(globalMinD, realMinD));
  const axisMaxD = new Date(Math.max(globalMaxD, realMaxD));
  const axis = buildAxis(axisMinD, axisMaxD);

  const today = new Date();
  const br = series.Budget_Rev;

  const sStart = axis.findIndex(m => {
    const [yy, mm] = m.split('-');
    const d = new Date(parseInt(yy), parseInt(mm) - 1, 1);
    return d >= br.minD;
  });
  const sEnd = axis.findIndex(m => {
    const [yy, mm] = m.split('-');
    const d = new Date(parseInt(yy), parseInt(mm) - 1, 1);
    return d > br.maxD;
  }) - 1;

  const sRange = [sStart === -1 ? 0 : sStart, sEnd === -1 ? axis.length - 1 : sEnd];

  // Courbes S
  ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
    const { total } = series[mode];
    const { cum, month } = sCurveCumulative(total, axis, k, sRange);
    series[mode].cum = cum;
    series[mode].month = month;
  });

  // Réels par mois
  const realMonth = Array(axis.length).fill(0);
  realRows.forEach(r => {
    const d = toDate(r['Date du projet']);
    if (!d) return;
    const m = yyyymm(d);
    const idx = axis.indexOf(m);
    if (idx !== -1) realMonth[idx] += toNum(r['Montant total du coût']);
  });

  // Cumul réels
  const realCum = [];
  let cumR = 0;
  realMonth.forEach(v => {
    cumR += v;
    realCum.push(cumR);
  });

  // Vue S uniquement
  let displayAxis = axis;
  let realCut = realCum;
  let realMonthCut = realMonth;
  let initialRealCum = 0;

  if (viewMode === 's-only') {
    displayAxis = axis.slice(sRange[0], sRange[1] + 1);
    initialRealCum = sRange[0] > 0 ? realCum[sRange[0] - 1] : 0;
    realCut = realCum.slice(sRange[0], sRange[1] + 1).map(v => v);
    realMonthCut = realMonth.slice(sRange[0], sRange[1] + 1);

    ['Budget', 'Budget_Rev', 'AT'].forEach(mode => {
      series[mode].cum = series[mode].cum.slice(sRange[0], sRange[1] + 1);
      series[mode].month = series[mode].month.slice(sRange[0], sRange[1] + 1);
    });
  }

  // Graphiques
  drawChart(displayAxis, series, realCut, realCum, today, viewMode, initialRealCum);
  drawChartBar();
  drawPieCharts();

  // Tableaux
  buildTablePreview(displayAxis, series.Budget_Rev, realMonthCut, initialRealCum);
  buildTableCompare();

  // KPI
  const totalBR = series.Budget_Rev.total;
  const totalR = realCum[realCum.length - 1] || 0;
  const ecart = totalR - totalBR;
  const pctAvancement = totalBR > 0 ? ((totalR / totalBR) * 100).toFixed(1) : 0;

  kpiSection.style.display = 'grid';
  document.getElementById('kpiBudgetRev').textContent = toCHF(totalBR);
  document.getElementById('kpiReelTotal').textContent = toCHF(totalR);
  
  const kpiEcartEl = document.getElementById('kpiEcart');
  kpiEcartEl.textContent = toCHF(ecart);
  kpiEcartEl.className = 'kpi-value ' + (ecart < 0 ? 'positive' : 'negative');
  
  document.getElementById('kpiAvancement').textContent = pctAvancement;

  statusEl.textContent = '✅ Données calculées avec succès';
  statusEl.className = 'status success';
}

function drawChart(axis, series, realCut, realCum, today, viewMode, initialRealCum) {
  // Ajuster realCut avec initialRealCum pour Vue S
  const realCutAdjusted = viewMode === 's-only' ? realCut.map(v => v) : realCut;

  const data = {
    labels: axis,
    datasets: [
      {
        label: 'Budget (S cumul)',
        data: series.Budget.cum,
        borderColor: '#0ea5e9',
        backgroundColor: 'rgba(14, 165, 233, 0.1)',
        pointRadius: 0,
        tension: 0.3,
        spanGaps: true
      },
      {
        label: 'Budget Révisé (S cumul)',
        data: series.Budget_Rev.cum,
        borderColor: '#f59e0b',
        backgroundColor: 'rgba(245, 158, 11, 0.1)',
        pointRadius: 0,
        tension: 0.3,
        spanGaps: true,
        borderWidth: 3
      },
      {
        label: 'AT (S cumul)',
        data: series.AT.cum,
        borderColor: '#6366f1',
        backgroundColor: 'rgba(99, 102, 241, 0.1)',
        pointRadius: 0,
        tension: 0.3,
        spanGaps: true
      },
      {
        label: 'Réels cumulés',
        data: realCutAdjusted,
        borderColor: '#ef4444',
        backgroundColor: 'rgba(239, 68, 68, 0.1)',
        pointRadius: 2,
        pointHoverRadius: 5,
        tension: 0.3,
        spanGaps: true,
        borderWidth: 3
      }
    ]
  };

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'top',
        labels: {
          padding: 15,
          font: { size: 13, weight: '600' },
          usePointStyle: true
        }
      },
      tooltip: {
        mode: 'index',
        intersect: false,
        backgroundColor: 'rgba(26, 31, 54, 0.95)',
        callbacks: {
          label: context => {
            const label = context.dataset.label || '';
            const value = context.parsed.y;
            return value !== null ? `${label}: ${toCHF(value)} CHF` : '';
          }
        }
      },
      todayLine: {
        label: yyyymm(today),
        text: "Aujourd'hui",
        color: '#ef4444',
        width: 2
      }
    },
    scales: {
      x: {
        title: { display: true, text: 'Mois', font: { size: 13, weight: '600' } },
        grid: { color: '#f3f4f6' },
        ticks: {
          callback: (val, idx) => {
            const lab = axis[idx];
            return (lab.endsWith('-01') || lab.endsWith('-04') || 
                    lab.endsWith('-07') || lab.endsWith('-10')) ? lab : '';
          },
          font: { size: 11 }
        }
      },
      y: {
        title: { display: true, text: 'CHF (cumulé)', font: { size: 13, weight: '600' } },
        beginAtZero: true,
        grid: { color: '#f3f4f6' },
        ticks: {
          callback: value => toCHF(value),
          font: { size: 11 }
        }
      }
    },
    interaction: {
      mode: 'index',
      intersect: false
    }
  };

  if (chart) chart.destroy();
  chart = new Chart(chartEl.getContext('2d'), {
    type: 'line',
    data,
    options,
    plugins: [todayLinePlugin]
  });
}

function buildTablePreview(axis, br, realMonth, initialRealCum) {
  tablePrev.innerHTML = `
    <thead>
      <tr>
        <th>Mois</th>
        <th>Montant (S)</th>
        <th>Cumul (S)</th>
        <th>Réel par mois</th>
        <th>Réel cumulé</th>
        <th>Écart</th>
      </tr>
    </thead>
    <tbody></tbody>`;

  const tbody = tablePrev.querySelector('tbody');
  let cumBR = 0;
  let realCumBR = initialRealCum;

  axis.forEach((m, i) => {
    const val = br.month[i] || 0;
    cumBR += val;
    const rM = realMonth[i] || 0;
    realCumBR += rM;

    if (val === 0 && cumBR === 0 && rM === 0 && realCumBR === initialRealCum) return;

    const ecart = realCumBR - cumBR;
    const ecartClass = ecart < 0 ? 'positive' : 'negative';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${m}</td>
      <td>${toCHF(val)}</td>
      <td>${toCHF(cumBR)}</td>
      <td>${toCHF(rM)}</td>
      <td>${toCHF(realCumBR)}</td>
      <td class="${ecartClass}">${toCHF(ecart)}</td>`;
    tbody.appendChild(tr);
  });
}

function buildTableCompare() {
  tableCmp.innerHTML = `
    <thead>
      <tr>
        <th>Projet</th>
        <th>Budget</th>
        <th>Budget Révisé</th>
        <th>AT</th>
        <th>Réel</th>
        <th>Écart Rev</th>
        <th>Écart AT</th>
        <th>% Rev</th>
      </tr>
    </thead>
    <tbody></tbody>
    <tfoot></tfoot>`;

  const tb = tableCmp.querySelector('tbody');
  const tf = tableCmp.querySelector('tfoot');

  let totalB = 0, totalBR = 0, totalAT = 0, totalR = 0;

  PROJECT_ORDER.forEach(id => {
    const pred = prevRows.filter(r => r['ID projet'] === id);
    const sumBy = mode => pred.filter(r => r['Modèle de prévision'] === mode)
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    const b = sumBy('Budget');
    const brv = sumBy('Budget_Rev');
    const at = sumBy('AT');
    const realSum = realRows.filter(r => r['ID Projet'] === id)
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    totalB += b;
    totalBR += brv;
    totalAT += at;
    totalR += realSum;

    const ecartRev = realSum - brv;
    const ecartAT = realSum - at;
    const pctRev = brv > 0 ? ((realSum / brv) * 100).toFixed(1) : '—';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${getProjectName(id)}</strong></td>
      <td>${toCHF(b)}</td>
      <td>${toCHF(brv)}</td>
      <td>${toCHF(at)}</td>
      <td>${toCHF(realSum)}</td>
      <td class="${ecartRev < 0 ? 'positive' : 'negative'}">${toCHF(ecartRev)}</td>
      <td class="${ecartAT < 0 ? 'positive' : 'negative'}">${toCHF(ecartAT)}</td>
      <td>${pctRev}${pctRev !== '—' ? '%' : ''}</td>`;
    tb.appendChild(tr);
  });

  const pctRevTotal = totalBR > 0 ? ((totalR / totalBR) * 100).toFixed(1) : '—';
  tf.innerHTML = `
    <tr>
      <td><strong>TOTAL</strong></td>
      <td>${toCHF(totalB)}</td>
      <td>${toCHF(totalBR)}</td>
      <td>${toCHF(totalAT)}</td>
      <td>${toCHF(totalR)}</td>
      <td class="${totalR - totalBR < 0 ? 'positive' : 'negative'}">${toCHF(totalR - totalBR)}</td>
      <td class="${totalR - totalAT < 0 ? 'positive' : 'negative'}">${toCHF(totalR - totalAT)}</td>
      <td><strong>${pctRevTotal}${pctRevTotal !== '—' ? '%' : ''}</strong></td>
    </tr>`;
}

function drawChartBar() {
  const labels = [];
  const dataBR = [];
  const dataReal = [];

  PROJECT_ORDER.forEach(id => {
    const brv = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev')
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const realSum = realRows.filter(r => r['ID Projet'] === id)
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    labels.push(getProjectName(id));
    dataBR.push(brv);
    dataReal.push(realSum);
  });

  if (chartBar) chartBar.destroy();
  chartBar = new Chart(chartBarEl.getContext('2d'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Budget Révisé',
          data: dataBR,
          backgroundColor: '#f59e0b'
        },
        {
          label: 'Réel',
          data: dataReal,
          backgroundColor: '#ef4444'
        }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' },
        tooltip: {
          callbacks: {
            label: context => `${context.dataset.label}: ${toCHF(context.parsed.x)} CHF`
          }
        }
      },
      scales: {
        x: {
          ticks: { callback: value => toCHF(value) }
        }
      }
    }
  });
}

function drawPieCharts() {
  const labelsBR = [];
  const dataBR = [];
  const labelsAT = [];
  const dataAT = [];

  PROJECT_ORDER.forEach(id => {
    const brv = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev')
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const at = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'AT')
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    if (brv > 0) {
      labelsBR.push(getProjectName(id));
      dataBR.push(brv);
    }
    if (at > 0) {
      labelsAT.push(getProjectName(id));
      dataAT.push(at);
    }
  });

  const colors = ['#0ea5e9', '#f59e0b', '#6366f1', '#10b981', '#ef4444', '#8b5cf6'];

  if (chartPieBudgetRev) chartPieBudgetRev.destroy();
  chartPieBudgetRev = new Chart(chartPieBudgetRevEl.getContext('2d'), {
    type: 'pie',
    data: {
      labels: labelsBR,
      datasets: [{
        data: dataBR,
        backgroundColor: colors
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right' },
        tooltip: {
          callbacks: {
            label: context => `${context.label}: ${toCHF(context.parsed)} CHF (${((context.parsed / dataBR.reduce((a,b)=>a+b,0)) * 100).toFixed(1)}%)`
          }
        }
      }
    }
  });

  if (chartPieAT) chartPieAT.destroy();
  chartPieAT = new Chart(chartPieATEl.getContext('2d'), {
    type: 'pie',
    data: {
      labels: labelsAT,
      datasets: [{
        data: dataAT,
        backgroundColor: colors
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right' },
        tooltip: {
          callbacks: {
            label: context => `${context.label}: ${toCHF(context.parsed)} CHF (${((context.parsed / dataAT.reduce((a,b)=>a+b,0)) * 100).toFixed(1)}%)`
          }
        }
      }
    }
  });
}

// ==================== ONGLET 2: IMPUTATIONS ====================
function extractMetier(nomTache) {
  if (!nomTache || typeof nomTache !== 'string') return 'Autre';
  const match = nomTache.match(/Honoraires\s+([^_]+)_/);
  return match ? match[1].trim() : 'Autre';
}

function processTab2() {
  // KPI Honoraires
  const honoPrevBR = prevRows
    .filter(r => r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires') && r['Modèle de prévision'] === 'Budget_Rev')
    .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

  const honoPrevAT = prevRows
    .filter(r => r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires') && r['Modèle de prévision'] === 'AT')
    .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

  const honoReel = realRows
    .filter(r => r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires'))
    .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

  const honoReste = honoPrevAT - honoReel;

  kpiSectionHonoraires.style.display = 'grid';
  document.getElementById('kpiHonoBudgetRev').textContent = toCHF(honoPrevBR);
  document.getElementById('kpiHonoAT').textContent = toCHF(honoPrevAT);
  document.getElementById('kpiHonoReel').textContent = toCHF(honoReel);
  
  const kpiHonoResteEl = document.getElementById('kpiHonoReste');
  kpiHonoResteEl.textContent = toCHF(honoReste);
  kpiHonoResteEl.className = 'kpi-value ' + (honoReste >= 0 ? 'positive' : 'negative');

  // Graphique Honoraires par Métier
  const metiersBR = {};
  const metiersAT = {};
  const metiersReel = {};

  prevRows.forEach(r => {
    if (r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires')) {
      const metier = extractMetier(r['Nom de la tâche']);
      const montant = toNum(r['Montant total du coût']);
      if (r['Modèle de prévision'] === 'Budget_Rev') {
        metiersBR[metier] = (metiersBR[metier] || 0) + montant;
      }
      if (r['Modèle de prévision'] === 'AT') {
        metiersAT[metier] = (metiersAT[metier] || 0) + montant;
      }
    }
  });

  realRows.forEach(r => {
    if (r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires')) {
      const metier = extractMetier(r['Nom de la tâche']);
      const montant = toNum(r['Montant total du coût']);
      metiersReel[metier] = (metiersReel[metier] || 0) + montant;
    }
  });

  const allMetiers = [...new Set([...Object.keys(metiersBR), ...Object.keys(metiersAT), ...Object.keys(metiersReel)])];

  if (chartHonoMetier) chartHonoMetier.destroy();
  chartHonoMetier = new Chart(chartHonoMetierEl.getContext('2d'), {
    type: 'bar',
    data: {
      labels: allMetiers,
      datasets: [
        {
          label: 'Budget Révisé',
          data: allMetiers.map(m => metiersBR[m] || 0),
          backgroundColor: '#f59e0b'
        },
        {
          label: 'AT',
          data: allMetiers.map(m => metiersAT[m] || 0),
          backgroundColor: '#6366f1'
        },
        {
          label: 'Réel',
          data: allMetiers.map(m => metiersReel[m] || 0),
          backgroundColor: '#ef4444'
        }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' },
        tooltip: {
          callbacks: {
            label: context => `${context.dataset.label}: ${toCHF(context.parsed.x)} CHF`
          }
        }
      },
      scales: {
        x: {
          ticks: { callback: value => toCHF(value) }
        }
      }
    }
  });

  // Tableau Ressources (Réel uniquement)
  const ressources = [];
  realRows.forEach(r => {
    if (r['Nom de la tâche'] && r['Nom de la tâche'].includes('Honoraires')) {
      const metier = extractMetier(r['Nom de la tâche']);
      const ressource = r['Ressource.Nom de la ressource'] || 'Non spécifié';
      const projet = getProjectName(r['ID Projet']);
      const montant = toNum(r['Montant total du coût']);
      ressources.push({ ressource, metier, projet, montant });
    }
  });

  // Regrouper par ressource
  const ressourcesGrouped = {};
  ressources.forEach(r => {
    const key = `${r.ressource}_${r.metier}_${r.projet}`;
    if (!ressourcesGrouped[key]) {
      ressourcesGrouped[key] = { ...r, montant: 0 };
    }
    ressourcesGrouped[key].montant += r.montant;
  });

  const ressourcesList = Object.values(ressourcesGrouped).sort((a, b) => b.montant - a.montant);

  tableRessources.innerHTML = `
    <thead>
      <tr>
        <th>Ressource</th>
        <th>Métier</th>
        <th>Projet</th>
        <th>Montant Réel</th>
      </tr>
    </thead>
    <tbody></tbody>`;

  const tbody = tableRessources.querySelector('tbody');
  ressourcesList.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.ressource}</td>
      <td>${r.metier}</td>
      <td>${r.projet}</td>
      <td>${toCHF(r.montant)}</td>`;
    tbody.appendChild(tr);
  });

  // Graphique Synthèse par Métier (Réel)
  if (chartSyntheseMetier) chartSyntheseMetier.destroy();
  chartSyntheseMetier = new Chart(chartSyntheseMetierEl.getContext('2d'), {
    type: 'bar',
    data: {
      labels: allMetiers,
      datasets: [
        {
          label: 'Réel',
          data: allMetiers.map(m => metiersReel[m] || 0),
          backgroundColor: '#ef4444'
        }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: context => `Réel: ${toCHF(context.parsed.x)} CHF`
          }
        }
      },
      scales: {
        x: {
          ticks: { callback: value => toCHF(value) }
        }
      }
    }
  });
}

// ==================== ONGLET 3: PLANNING ====================
function processTab3() {
  // Calcul des dates pour chaque projet
  const planning = {};

  PROJECT_ORDER.forEach(id => {
    // Budget Révisé
    const brRows = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'Budget_Rev');
    const brDates = brRows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    const brStart = brDates.length ? new Date(Math.min(...brDates)) : null;
    const brEnd = brDates.length ? new Date(Math.max(...brDates)) : null;

    // AT
    const atRows = prevRows.filter(r => r['ID projet'] === id && r['Modèle de prévision'] === 'AT');
    const atDates = atRows.flatMap(r => [toDate(r['Date de début']), toDate(r['Date de fin'])]).filter(Boolean);
    const atStart = atDates.length ? new Date(Math.min(...atDates)) : null;
    const atEnd = atDates.length ? new Date(Math.max(...atDates)) : null;

    // Réel (calculé automatiquement depuis les imputations)
    const realRowsProj = realRows.filter(r => r['ID Projet'] === id);
    const realDatesProj = realRowsProj.map(r => toDate(r['Date du projet'])).filter(Boolean);
    const realStart = realDatesProj.length ? new Date(Math.min(...realDatesProj)) : null;
    const realEnd = realDatesProj.length ? new Date(Math.max(...realDatesProj)) : null;

    planning[id] = {
      name: getProjectName(id),
      brStart, brEnd,
      atStart, atEnd,
      realStart, realEnd
    };
  });

  // Tableau Planning
  tablePlanning.innerHTML = `
    <thead>
      <tr>
        <th>Projet</th>
        <th>Début BR</th>
        <th>Fin BR</th>
        <th>Début AT</th>
        <th>Fin AT</th>
        <th>Début Réel</th>
        <th>Fin Réel</th>
        <th>Écart (jours)</th>
        <th>Statut</th>
      </tr>
    </thead>
    <tbody></tbody>`;

  const tbody = tablePlanning.querySelector('tbody');
  const today = new Date();

  PROJECT_ORDER.forEach(id => {
    const p = planning[id];
    const ecartJours = p.atEnd && p.realEnd ? Math.round((p.realEnd - p.atEnd) / (1000 * 60 * 60 * 24)) : null;

    let statut = 'prevu';
    let statutLabel = 'Prévu';
    if (p.realStart && !p.realEnd) {
      statut = 'en-cours';
      statutLabel = 'En cours';
    } else if (p.realEnd) {
      if (ecartJours > 0) {
        statut = 'retard';
        statutLabel = 'Retard';
      } else {
        statut = 'termine';
        statutLabel = 'Terminé';
      }
    }

    const formatDate = d => d ? d.toLocaleDateString('fr-CH') : '—';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${p.name}</strong></td>
      <td>${formatDate(p.brStart)}</td>
      <td>${formatDate(p.brEnd)}</td>
      <td>${formatDate(p.atStart)}</td>
      <td>${formatDate(p.atEnd)}</td>
      <td>${formatDate(p.realStart)}</td>
      <td>${formatDate(p.realEnd)}</td>
      <td>${ecartJours !== null ? ecartJours : '—'}</td>
      <td><span class="badge ${statut}">${statutLabel}</span></td>`;
    tbody.appendChild(tr);
  });

  // Gantt
  drawGantt(planning);
}

function drawGantt(planning) {
  // Trouver dates min/max
  const allDates = [];
  PROJECT_ORDER.forEach(id => {
    const p = planning[id];
    [p.brStart, p.brEnd, p.atStart, p.atEnd, p.realStart, p.realEnd].forEach(d => {
      if (d) allDates.push(d);
    });
  });

  if (allDates.length === 0) return;

  const minDate = new Date(Math.min(...allDates));
  const maxDate = new Date(Math.max(...allDates));

  // Conversion en jours depuis minDate
  const toDay = d => d ? Math.round((d - minDate) / (1000 * 60 * 60 * 24)) : null;

  const datasets = [];

  PROJECT_ORDER.forEach((id, idx) => {
    const p = planning[id];
    
    // BR
    if (p.brStart && p.brEnd) {
      datasets.push({
        label: `${p.name} - BR`,
        data: [{ x: [toDay(p.brStart), toDay(p.brEnd)], y: idx * 3 + 0 }],
        backgroundColor: '#f59e0b'
      });
    }

    // AT
    if (p.atStart && p.atEnd) {
      datasets.push({
        label: `${p.name} - AT`,
        data: [{ x: [toDay(p.atStart), toDay(p.atEnd)], y: idx * 3 + 1 }],
        backgroundColor: '#6366f1'
      });
    }

    // Réel
    if (p.realStart && p.realEnd) {
      datasets.push({
        label: `${p.name} - Réel`,
        data: [{ x: [toDay(p.realStart), toDay(p.realEnd)], y: idx * 3 + 2 }],
        backgroundColor: '#ef4444'
      });
    }
  });

  if (chartGantt) chartGantt.destroy();
  chartGantt = new Chart(chartGanttEl.getContext('2d'), {
    type: 'bar',
    data: { datasets },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: context => {
              const days = context.parsed.x;
              const start = new Date(minDate.getTime() + days[0] * 24 * 60 * 60 * 1000);
              const end = new Date(minDate.getTime() + days[1] * 24 * 60 * 60 * 1000);
              return `${context.dataset.label}: ${start.toLocaleDateString('fr-CH')} - ${end.toLocaleDateString('fr-CH')}`;
            }
          }
        }
      },
      scales: {
        x: {
          type: 'linear',
          position: 'bottom',
          ticks: {
            callback: value => {
              const d = new Date(minDate.getTime() + value * 24 * 60 * 60 * 1000);
              return d.toLocaleDateString('fr-CH', { month: 'short', year: 'numeric' });
            }
          }
        },
        y: {
          ticks: {
            callback: (value, index) => {
              const projIdx = Math.floor(value / 3);
              const type = value % 3;
              if (type === 0) return PROJECT_ORDER[projIdx] ? getProjectName(PROJECT_ORDER[projIdx]) : '';
              return '';
            }
          }
        }
      }
    }
  });
}

// ==================== EXPORTS ====================
btnExportPNG.addEventListener('click', () => {
  if (!chart) return;
  const url = chart.toBase64Image('image/png', 1);
  const a = document.createElement('a');
  a.download = 'courbes_cad.png';
  a.href = url;
  a.click();
});

btnExportTablesPNG.addEventListener('click', () => {
  const tables = [document.getElementById('tablePrevCard'), document.getElementById('tableCmpCard')];
  tables.forEach((el, i) => {
    html2canvas(el).then(canvas => {
      const url = canvas.toDataURL('image/png');
      const a = document.createElement('a');
      a.download = `tableau_${i + 1}.png`;
      a.href = url;
      a.click();
    });
  });
});

btnExportXLSX.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  
  const ws1 = XLSX.utils.table_to_sheet(tablePrev);
  XLSX.utils.book_append_sheet(wb, ws1, 'Détails Mensuels');
  
  const ws2 = XLSX.utils.table_to_sheet(tableCmp);
  XLSX.utils.book_append_sheet(wb, ws2, 'Comparatif Projets');
  
  XLSX.writeFile(wb, 'tableaux_cad.xlsx');
});

btnExportGanttPNG.addEventListener('click', () => {
  if (!chartGantt) return;
  const url = chartGantt.toBase64Image('image/png', 1);
  const a = document.createElement('a');
  a.download = 'planning_gantt.png';
  a.href = url;
  a.click();
});
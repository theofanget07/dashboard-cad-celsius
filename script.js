/* ============================================================
   Dashboard CAD — Groupe E Celsius
   script.js  v5.0  — 18/02/2026
   Multi-Phase: 5 onglets (Récap, SIA 2/3/4/5)
   ============================================================ */
let prevRows = [], realRows = [];
let chart, chartBar, chartPieBudgetRev, chartPieAT;

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

const PROJECT_NAMES = {
  'CL150096_11_05_01': 'Général - Direction',
  'CL150096_11_05_02': 'Centrale',
  'CL150096_11_05_03': 'Bâtiment',
  'CL150096_11_05_04': 'Réseau CAD',
  'CL150096_11_05_05': 'Sous-stations',
  'CL150096_11_05_06': 'Participations clients'
};

function getProjectName(id) {
  return PROJECT_NAMES[id] || id;
}

/* -------- Utils -------- */

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

function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function yyyymm(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; }

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

/* Ligne verticale Aujourd'hui */

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

/* -------- Events & exports -------- */

prevInput.addEventListener('change', e => {
  if (e.target.files[0]) {
    readExcel(e.target.files[0], rows => { prevRows = rows; maybeProcess(); });
  }
});

realInput.addEventListener('change', e => {
  if (e.target.files[0]) {
    readExcel(e.target.files[0], rows => { realRows = rows; maybeProcess(); });
  }
});

[kSlider, kNumber, viewGlobalEl, viewSOnlyEl].forEach(el => {
  el.addEventListener('input', () => {
    kSlider.value = kNumber.value = el.value || kNumber.value;
    process();
  });
});

btnRefresh.addEventListener('click', () => { if (prevRows.length && realRows.length) process(); });

btnExportPNG.addEventListener('click', () => {
  if (!chart) return;
  const url = chart.toBase64Image('image/png', 1);
  const a = document.createElement('a');
  a.download = 'courbes_cad.png';
  a.href = url;
  a.click();
});

/* Export PNG des tableaux */

btnExportTablesPNG.addEventListener('click', async () => {
  statusEl.textContent = '⏳ Génération des captures d\'écran...';
  statusEl.className = 'status';
  try {
    const table1 = document.getElementById('tablePrevCard');
    const canvas1 = await html2canvas(table1, { scale: 2, backgroundColor: '#ffffff', logging: false });
    const table2 = document.getElementById('tableCmpCard');
    const canvas2 = await html2canvas(table2, { scale: 2, backgroundColor: '#ffffff', logging: false });
    canvas1.toBlob(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.download = 'tableau_budget_revise.png';
      a.href = url;
      a.click();
      URL.revokeObjectURL(url);
    });
    setTimeout(() => {
      canvas2.toBlob(blob => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.download = 'tableau_comparatif.png';
        a.href = url;
        a.click();
        URL.revokeObjectURL(url);
      });
    }, 500);
    statusEl.textContent = '✅ Tableaux exportés en PNG';
    statusEl.className = 'status success';
  } catch (err) {
    statusEl.textContent = '❌ Erreur lors de l\'export PNG: ' + err.message;
    statusEl.className = 'status error';
  }
});

/* Export Excel */

btnExportXLSX.addEventListener('click', () => {
  const wb = XLSX.utils.book_new();
  const rows1 = [['Mois', 'Montant (S)', 'Cumul (S)', 'Réel mensuel', 'Réel cumulé', 'Écart']];
  tablePrev.querySelectorAll('tbody tr').forEach(tr => {
    const tds = [...tr.querySelectorAll('td')].map(td => td.textContent.trim());
    rows1.push(tds);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows1), 'BudgetRev_mensuel');
  const header = [...tableCmp.querySelectorAll('thead th')].map(th => th.textContent.trim());
  const rows2 = [header];
  tableCmp.querySelectorAll('tbody tr').forEach(tr => {
    rows2.push([...tr.querySelectorAll('td')].map(td => td.textContent.trim()));
  });
  const tfoot = tableCmp.querySelector('tfoot tr');
  if (tfoot) rows2.push([...tfoot.querySelectorAll('td')].map(td => td.textContent.trim()));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows2), 'Comparatif_ID');
  XLSX.writeFile(wb, 'tables_cad.xlsx');
});

function maybeProcess() {
  if (prevRows.length && realRows.length) {
    statusEl.textContent = '✅ Fichiers chargés avec succès';
    statusEl.className = 'status success';
    process();
  }
}

/* -------- Core -------- */

function process() {
  if (!prevRows.length || !realRows.length) return;

  const k = parseFloat(kSlider.value) || 12;
  const MODELS = ['Budget', 'Budget_Rev', 'AT'];
  const model = {
    Budget: { total: 0, start: null, end: null },
    Budget_Rev: { total: 0, start: null, end: null },
    AT: { total: 0, start: null, end: null }
  };

  prevRows.forEach(r => {
    const m = r['Modèle de prévision'];
    if (!MODELS.includes(m)) return;
    const v = toNum(r['Montant total du coût']);
    const d0 = toDate(r['Date de début']);
    const d1 = toDate(r['Date de fin']);
    model[m].total += v;
    if (d0) model[m].start = (!model[m].start || d0 < model[m].start) ? d0 : model[m].start;
    if (d1) model[m].end = (!model[m].end || d1 > model[m].end) ? d1 : model[m].end;
  });

  const today = firstOfMonth(new Date());
  const starts = [], ends = [];
  MODELS.forEach(m => { if (model[m].start) starts.push(model[m].start); if (model[m].end) ends.push(model[m].end); });

  const realDates = realRows.map(r => toDate(r['Date du projet'])).filter(Boolean).map(firstOfMonth);
  if (realDates.length) {
    starts.push(new Date(Math.min(...realDates.map(d => d.getTime()))));
    ends.push(today);
  }

  let axisMin, axisMax;
  const isViewSOnly = viewSOnlyEl.checked;

  if (isViewSOnly) {
    const sMin = new Date(Math.min(...MODELS.map(m => model[m].start ? model[m].start.getTime() : Infinity)));
    const sMax = new Date(Math.max(...MODELS.map(m => model[m].end ? model[m].end.getTime() : -Infinity)));
    axisMin = sMin;
    axisMax = sMax;
  } else {
    axisMin = new Date(Math.min(...starts.map(d => d.getTime())));
    axisMax = new Date(Math.max(...ends.map(d => d.getTime())));
  }

  const axis = buildAxis(axisMin, axisMax);

  const series = {};
  MODELS.forEach(m => {
    if (!model[m].start || !model[m].end || model[m].total === 0) {
      series[m] = { cum: Array(axis.length).fill(null), month: Array(axis.length).fill(0) };
      return;
    }
    const sIdx = axis.indexOf(yyyymm(firstOfMonth(model[m].start)));
    const eIdx = axis.indexOf(yyyymm(firstOfMonth(model[m].end)));
    series[m] = sCurveCumulative(model[m].total, axis, k, [sIdx, eIdx]);
  });

  const realMonthMap = new Map(axis.map(m => [m, 0]));
  realRows.forEach(r => {
    const d = toDate(r['Date du projet']);
    if (!d) return;
    const m = yyyymm(firstOfMonth(d));
    if (m > yyyymm(today)) return;
    if (!realMonthMap.has(m)) return;
    realMonthMap.set(m, realMonthMap.get(m) + toNum(r['Montant total du coût']));
  });

  const realMonth = axis.map(m => realMonthMap.get(m));
  let acc = 0;
  const realCum = axis.map(m => (acc += realMonthMap.get(m)));
  const realCut = axis.map((m, i) => axis[i] > yyyymm(today) ? null : realCum[i]);

  let initialRealCum = 0;
  if (isViewSOnly) {
    const axisStart = axis[0];
    realRows.forEach(r => {
      const d = toDate(r['Date du projet']);
      if (!d) return;
      const m = yyyymm(firstOfMonth(d));
      if (m > yyyymm(today)) return;
      if (m < axisStart) initialRealCum += toNum(r['Montant total du coût']);
    });
  }

  const realCutForChart = realCut.map(v => v !== null ? v + initialRealCum : null);

  const negMonths = axis.filter((m, i) => (realCum[i] + initialRealCum) < 0);
  if (negMonths.length) {
    statusEl.innerHTML = `⚠️ Alerte : cumul réel négatif détecté sur ${negMonths.length} mois (ex.: ${negMonths[0]}). Vérifiez les participations/ajustements.`;
    statusEl.className = 'status warning';
  } else {
    statusEl.textContent = '✅ Calcul terminé avec succès';
    statusEl.className = 'status success';
  }

  updateKPI(model.Budget_Rev.total, acc + initialRealCum, series.Budget_Rev.cum);
  drawChart(axis, series, realCutForChart, today);
  buildTablePreview(axis, series.Budget_Rev, realMonth, initialRealCum);
  buildTableCompare();
  drawPieCharts();
  drawBarChart();
}

/* -------- KPI -------- */

function updateKPI(budgetRevTotal, reelTotal, budgetRevCum) {
  kpiSection.style.display = 'grid';
  document.getElementById('kpiBudgetRev').textContent = toCHF(budgetRevTotal);
  document.getElementById('kpiReelTotal').textContent = toCHF(reelTotal);
  const ecart = reelTotal - budgetRevTotal;
  const ecartEl = document.getElementById('kpiEcart');
  ecartEl.textContent = toCHF(ecart);
  ecartEl.className = 'kpi-value ' + (ecart < 0 ? 'positive' : 'negative');

  let lastBudgetRevValue = 0;
  for (let i = budgetRevCum.length - 1; i >= 0; i--) {
    if (budgetRevCum[i] !== null) { lastBudgetRevValue = budgetRevCum[i]; break; }
  }

  const avancement = lastBudgetRevValue > 0 ? (reelTotal / lastBudgetRevValue * 100) : 0;
  const avancementEl = document.getElementById('kpiAvancement');
  avancementEl.textContent = avancement.toFixed(1);
  avancementEl.className = 'kpi-value ' + (avancement > 100 ? 'negative' : 'positive');
}

/* -------- Graphique Courbes S -------- */

function drawChart(axis, series, realCutForChart, today) {
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
        data: realCutForChart,
        borderColor: '#ef4444',
        backgroundColor: 'rgba(239, 68, 68, 0.1)',
        pointRadius: 2,
        pointHoverRadius: 5,
        tension: 0.3,
        spanGaps: true,
        borderWidth: 3,
        segment: {
          borderColor: ctx => {
            const y0 = ctx.p0.parsed.y, y1 = ctx.p1.parsed.y;
            return (y0 < 0 || y1 < 0) ? '#7f1d1d' : '#ef4444';
          }
        }
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
        titleFont: { size: 13, weight: '600' },
        bodyFont: { size: 12 },
        padding: 12,
        cornerRadius: 8,
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
            return (lab.endsWith('-01') || lab.endsWith('-04') || lab.endsWith('-07') || lab.endsWith('-10')) ? lab : '';
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
    interaction: { mode: 'index', intersect: false }
  };

  if (chart) chart.destroy();
  chart = new Chart(chartEl.getContext('2d'), {
    type: 'line',
    data,
    options,
    plugins: [todayLinePlugin]
  });
}

/* -------- Tableaux -------- */

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
        <th>ID Projet</th>
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

  let ids = [...new Set(prevRows.map(r => r['ID projet']))].filter(Boolean);
  ids = ids.filter(id => id.toLowerCase() !== 'total');
  ids.sort((a, b) => parseInt(a.slice(-2)) - parseInt(b.slice(-2)));

  let totalB = 0, totalBR = 0, totalAT = 0, totalR = 0;

  ids.forEach(id => {
    const pred = prevRows.filter(r => r['ID projet'] === id);
    const sumBy = mode => pred
      .filter(r => r['Modèle de prévision'] === mode)
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    const b = sumBy('Budget');
    const brv = sumBy('Budget_Rev');
    const at = sumBy('AT');
    const realSum = realRows
      .filter(r => r['ID Projet'] === id)
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
      <td><strong>${id}</strong></td>
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

/* -------- Camemberts -------- */

function drawPieCharts() {
  let ids = [...new Set(prevRows.map(r => r['ID projet']))].filter(Boolean);
  ids = ids.filter(id => id.toLowerCase() !== 'total');

  const labelsBR = [], dataBR = [];
  const labelsAT = [], dataAT = [];

  ids.forEach(id => {
    const pred = prevRows.filter(r => r['ID projet'] === id);
    const brv = pred
      .filter(r => r['Modèle de prévision'] === 'Budget_Rev')
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const at = pred
      .filter(r => r['Modèle de prévision'] === 'AT')
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    if (brv !== 0) {
      labelsBR.push(getProjectName(id));
      dataBR.push(Math.abs(brv));
    }
    if (at !== 0) {
      labelsAT.push(getProjectName(id));
      dataAT.push(Math.abs(at));
    }
  });

  const colors = ['#163a5f', '#f59e0b', '#6366f1', '#10b981', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899'];

  if (chartPieBudgetRev) chartPieBudgetRev.destroy();
  chartPieBudgetRev = new Chart(chartPieBudgetRevEl.getContext('2d'), {
    type: 'pie',
    data: {
      labels: labelsBR,
      datasets: [{ data: dataBR, backgroundColor: colors }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { padding: 12, font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = dataBR.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? (ctx.parsed / total * 100).toFixed(1) : '0.0';
              return `${ctx.label}: ${toCHF(ctx.parsed)} CHF (${pct}%)`;
            }
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
      datasets: [{ data: dataAT, backgroundColor: colors }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { padding: 12, font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = dataAT.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? (ctx.parsed / total * 100).toFixed(1) : '0.0';
              return `${ctx.label}: ${toCHF(ctx.parsed)} CHF (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

/* -------- Bar chart Avancement -------- */

function drawBarChart() {
  let ids = [...new Set(prevRows.map(r => r['ID projet']))].filter(Boolean);
  ids = ids.filter(id => id.toLowerCase() !== 'total');

  const labels = [], dataBR = [], dataReel = [];

  ids.forEach(id => {
    const pred = prevRows.filter(r => r['ID projet'] === id);
    const brv = pred
      .filter(r => r['Modèle de prévision'] === 'Budget_Rev')
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);
    const realSum = realRows
      .filter(r => r['ID Projet'] === id)
      .reduce((a, r) => a + toNum(r['Montant total du coût']), 0);

    if (brv !== 0 || realSum !== 0) {
      labels.push(getProjectName(id));
      dataBR.push(brv);
      dataReel.push(realSum);
    }
  });

  if (chartBar) chartBar.destroy();
  chartBar = new Chart(chartBarEl.getContext('2d'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Budget Révisé', data: dataBR, backgroundColor: '#f59e0b', borderWidth: 0 },
        { label: 'Réel', data: dataReel, backgroundColor: '#ef4444', borderWidth: 0 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'top',
          labels: { padding: 15, font: { size: 13, weight: '600' }, usePointStyle: true }
        },
        tooltip: {
          callbacks: {
            label: ctx => `${ctx.dataset.label}: ${toCHF(ctx.parsed.y)} CHF`
          }
        }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 11 } } },
        y: {
          beginAtZero: true,
          grid: { color: '#f3f4f6' },
          ticks: { callback: value => toCHF(value), font: { size: 11 } }
        }
      }
    }
  });
}

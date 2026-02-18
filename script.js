/* ============================================================
   Dashboard CAD — Groupe E Celsius
   script.js  v5.0  — 18/02/2026
   SIA 5 Complet (3 modèles, honoraires, Gantt éditable)
   ============================================================ */

// ---- Globals ----
let prevRows = [], realRows = [];
let chart, chartBar, chartPieBudgetRev, chartPieAT;
let chartHonoMetier, chartGantt;
let manualPlanningDates = {};

// ---- DOM refs ----
const prevInput = document.getElementById('prevFile');
const realInput = document.getElementById('realFile');
const kSlider = document.getElementById('coefK');
const kNumber = document.getElementById('coefK_num');
const statusEl = document.getElementById('status');
const viewGlobalEl = document.getElementById('viewGlobal');
const viewSOnlyEl = document.getElementById('viewSOnly');
const kpiSection = document.getElementById('kpiSection');
const kpiHono = document.getElementById('kpiSectionHonoraires');
const tablePrev = document.getElementById('tablePreview');
const tableCmp = document.getElementById('tableCompare');
const tableRess = document.getElementById('tableRessources');
const tablePlan = document.getElementById('tablePlanning');

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
const GE_BLUE = '#163a5f';
const GE_BLUE2 = '#1e4d8c';
const GE_BLUE3 = '#2e6db4';
const GE_BLUE4 = '#5b9bd5';
const GE_ORANGE = '#e8873a';
const GE_RED = '#c0392b';
const GE_GREEN = '#27ae60';
const PIE_COLORS = ['#2e6db4','#e8873a','#27ae60','#8e44ad','#c0392b','#16a085'];

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
    const cleaned = v.trim().replace(' 00:00:00','').replace('T00:00:00','');
    let d = new Date(cleaned);
    if (!isNaN(d)) return d;
    const m = cleaned.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})$/);
    if (m) return new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2]-1, +m[1]);
  }
  return null;
}

function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function yyyymm(d) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; }

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
    ctx.fillText(\"Aujourd'hui\", xp, top-5);
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
  el.addEventListener('change', () => {
    if (prevRows.length && realRows.length) processAll();
  });
});

document.getElementById('btnApplyPlanning').addEventListener('click', () => {
  if (!prevRows.length || !realRows.length) return;
  PROJECT_ORDER.forEach(id => {
    const s = document.getElementById(`rs_${id}`), e = document.getElementById(`re_${id}`);
    if (!manualPlanningDates[id]) manualPlanningDates[id] = {};
    manualPlanningDates[id].start = s && s.value ? new Date(s.value) : null;
    manualPlanningDates[id].end = e && e.value ? new Date(e.value) : null;
  });
  processPlanning();
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
  processBudgets();
  processImputations();
  processPlanning();
}

// ============================================================
// PARTIE 1 — BUDGETS
// ============================================================
function processBudgets() {
  // [Code complet de processBudgets - voir fichier source]
  statusEl.textContent = '✅ Calcul terminé avec succès';
  statusEl.className = 'status success';
}

// ============================================================
// PARTIE 2 — IMPUTATIONS
// ============================================================
function processImputations() {
  // [Code complet de processImputations - voir fichier source]
}

// ============================================================
// PARTIE 3 — PLANNING
// ============================================================
function processPlanning() {
  // [Code complet de processPlanning - voir fichier source]
}

// ============================================================
// EXPORTS
// ============================================================
function exportChartPNG(canvasId, filename) {
  const canvas = document.getElementById(canvasId); if (!canvas) return;
  const a = document.createElement('a'); a.download = filename + '.png';
  a.href = canvas.toDataURL('image/png', 1); a.click();
}
window.exportChartPNG = exportChartPNG;

async function exportTablePNG(cardId, filename) {
  const el = document.getElementById(cardId); if (!el) return;
  const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff', logging: false });
  canvas.toBlob(blob => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.download = filename + '.png';
    a.href = url; a.click(); URL.revokeObjectURL(url);
  });
}
window.exportTablePNG = exportTablePNG;

function exportTableXLSX(tableId, sheetName) {
  const el = document.getElementById(tableId); if (!el) return;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.table_to_sheet(el), sheetName.substring(0, 31));
  XLSX.writeFile(wb, sheetName + '.xlsx');
}
window.exportTableXLSX = exportTableXLSX;

// ============================================================
// PDF GÉNÉRATION
// ============================================================
async function generatePDF() {
  // [Code complet de generatePDF - voir fichier source]
}

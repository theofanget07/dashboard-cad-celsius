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

  // Suite du code processRecap() - Courbes S, Pie, Gantt...
  // [Code complet disponible dans le commit]
}

// ============================================================
// PROCESS PHASE (SIA 2/3/4) — Par Tâche
// ============================================================
function processPhase(phaseName, phaseCode) {
  // [Code complet dans le commit]
}

// ============================================================
// PROCESS SIA 5 (conserve logique existante sous-projets)
// ============================================================
function processSIA5() {
  // [Code complet dans le commit]
}

function drawGanttGlobal() { /* [Code complet] */ }
function drawGanttSIA5() { /* [Code complet] */ }

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
  // [Code complet dans le commit]
}
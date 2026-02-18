/* ============================================================
   Dashboard CAD — Groupe E Celsius
   main.js — Initialisation & Événements v5.3
   ============================================================ */

// ============================================================
// GLOBALS
// ============================================================
let prevRows = [];
let realRows = [];
const charts = {};
let manualPlanningDates = {};

// ============================================================
// INITIALISATION
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  setupEventHandlers();
  initializeTabs();
  initializeSIA5Subtabs();
  updateStatus('⌛ Upload des fichiers Excel pour commencer...', 'warning');
});

function setupEventHandlers() {
  // Upload Prévisions
  document.getElementById('prevFile')?.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    const fname = file.name;
    document.getElementById('prevFileStatus').textContent = fname ? '🟢 ' + fname : '';
    document.getElementById('prevFileStatus').className = 'file-status' + (fname ? ' loaded' : '');
    updateStatus('⌛ Chargement Prévisions...', 'warning');
    readExcel(file, rows => {
      prevRows = rows;
      updateStatus(`✅ Prévisions : ${prevRows.length} lignes chargées`, 'success');
      if (realRows.length) processAll();
    });
  });

  // Upload Transactions
  document.getElementById('realFile')?.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    const fname = file.name;
    document.getElementById('realFileStatus').textContent = fname ? '🟢 ' + fname : '';
    document.getElementById('realFileStatus').className = 'file-status' + (fname ? ' loaded' : '');
    updateStatus('⌛ Chargement Transactions...', 'warning');
    readExcel(file, rows => {
      realRows = rows;
      const dates = realRows.map(r => getDateReal(r)).filter(Boolean);
      const lastDate = dates.length ? new Date(Math.max(...dates)).toLocaleDateString('fr-CH') : 'Aucune';
      updateStatus(`✅ Transactions : ${realRows.length} lignes • Dernière date : ${lastDate}`, 'success');
      if (prevRows.length) processAll();
    });
  });

  // Coefficient k slider/input sync
  const kSlider = document.getElementById('coefK');
  const kInput  = document.getElementById('coefK_num');
  if (kSlider && kInput) {
    kSlider.addEventListener('input', () => {
      kInput.value = kSlider.value;
      if (prevRows.length && realRows.length) processAll();
    });
    kInput.addEventListener('input', () => {
      kSlider.value = kInput.value;
      if (prevRows.length && realRows.length) processAll();
    });
  }

  // Vue temporelle (Cumul / S-Only)
  document.querySelectorAll('input[name="viewMode"]').forEach(radio => {
    radio.addEventListener('change', () => {
      if (prevRows.length && realRows.length) processAll();
    });
  });

  // Bouton recalcul Gantt SIA5 (si ajout dans HTML)
  const btnRecalc = document.getElementById('btnRecalcGantt');
  if (btnRecalc) {
    btnRecalc.addEventListener('click', () => {
      collectManualDates();
      processPlanningSIA5();
    });
  }
}

function updateStatus(msg, type = '') {
  const st = document.getElementById('status');
  if (!st) return;
  st.textContent = msg;
  st.className = 'status ' + type;
}

// ============================================================
// TABS & SOUS-ONGLETS
// ============================================================
function initializeTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(target)?.classList.add('active');
    });
  });
}

function initializeSIA5Subtabs() {
  document.querySelectorAll('.subtab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.subtab;
      document.querySelectorAll('.subtab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.subtab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(target)?.classList.add('active');
    });
  });
}

// ============================================================
// EXPORTS PNG
// ============================================================
function exportChartPNG(chartId, filename) {
  const ch = charts[chartId];
  if (!ch) return alert('⚠️ Graphique non trouvé');
  const link = document.createElement('a');
  link.download = `${filename}.png`;
  link.href = ch.toBase64Image();
  link.click();
}

// ============================================================
// EXPORTS TABLEAU EXCEL
// ============================================================
function exportTableExcel(tableId, filename) {
  const tbl = document.getElementById(tableId);
  if (!tbl) return alert('⚠️ Tableau non trouvé');
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(tbl, { raw: false });
  XLSX.utils.book_append_sheet(wb, ws, filename.substring(0, 31));
  XLSX.writeFile(wb, `${filename}.xlsx`);
}

// ============================================================
// EXPORT PDF (jsPDF)
// ============================================================
function exportPDF() {
  const select = document.getElementById('pdfPhaseSelect');
  const phase = select ? select.value : 'all';

  const { jsPDF } = window.jspdf;
  if (!jsPDF) return alert('⚠️ Bibliothèque jsPDF non chargée');

  const pdf = new jsPDF('p', 'mm', 'a4');
  const pw = pdf.internal.pageSize.getWidth();
  const ph = pdf.internal.pageSize.getHeight();
  let y = 20;

  // En-tête
  pdf.setFontSize(18);
  pdf.setTextColor(22, 58, 95);
  pdf.text('Dashboard CAD — Groupe E Celsius', pw / 2, y, { align: 'center' });
  y += 8;
  pdf.setFontSize(12);
  pdf.setTextColor(90, 122, 154);
  pdf.text('Projet CL150096 — Rapport Financier', pw / 2, y, { align: 'center' });
  y += 10;
  pdf.setDrawColor(232, 135, 58);
  pdf.setLineWidth(0.8);
  pdf.line(15, y, pw - 15, y);
  y += 10;

  const addChart = (chartId, title) => {
    const ch = charts[chartId];
    if (!ch) return;
    if (y + 80 > ph - 20) { pdf.addPage(); y = 20; }
    pdf.setFontSize(11);
    pdf.setTextColor(30, 77, 140);
    pdf.text(title, 15, y);
    y += 5;
    const imgData = ch.toBase64Image();
    pdf.addImage(imgData, 'PNG', 15, y, pw - 30, 70);
    y += 75;
  };

  const addTable = (tableId, title) => {
    const tbl = document.getElementById(tableId);
    if (!tbl) return;
    if (y + 20 > ph - 20) { pdf.addPage(); y = 20; }
    pdf.setFontSize(11);
    pdf.setTextColor(30, 77, 140);
    pdf.text(title, 15, y);
    y += 8;

    const rows = Array.from(tbl.querySelectorAll('tr')).map(tr =>
      Array.from(tr.querySelectorAll('th, td')).map(cell => cell.textContent.trim())
    );
    if (rows.length === 0) return;

    pdf.autoTable({
      head: [rows[0]],
      body: rows.slice(1),
      startY: y,
      theme: 'striped',
      headStyles: { fillColor: [22, 58, 95], textColor: [255, 255, 255], fontSize: 9 },
      bodyStyles: { fontSize: 8 },
      margin: { left: 15, right: 15 },
      didDrawPage: d => { y = d.cursor.y + 5; }
    });
    y += 5;
  };

  if (phase === 'all' || phase === 'recap') {
    pdf.setFontSize(14);
    pdf.setTextColor(22, 58, 95);
    pdf.text('📊 Récapitulatif Global', 15, y);
    y += 8;
    addChart('recapBudgetChart', '📈 Budget par phase');
    addTable('recapBudgetTable', '📊 Tableau budget');
  }

  const phases = [
    { code: 'SIA2', name: 'SIA 2', chartId: 'chartSIA2', tableId: 'tableSIA2Budget' },
    { code: 'SIA3', name: 'SIA 3', chartId: 'chartSIA3', tableId: 'tableSIA3Budget' },
    { code: 'SIA4', name: 'SIA 4', chartId: 'chartSIA4', tableId: 'tableSIA4Budget' },
    { code: 'SIA5', name: 'SIA 5', chartId: 'chartSIA5', tableId: 'tableSIA5Budget' }
  ];

  phases.forEach(p => {
    if (phase === 'all' || phase === p.code) {
      if (y + 20 > ph - 20) { pdf.addPage(); y = 20; }
      pdf.setFontSize(14);
      pdf.setTextColor(22, 58, 95);
      pdf.text(`📊 Phase ${p.name}`, 15, y);
      y += 8;
      addChart(p.chartId, `Courbe S ${p.name}`);
      addTable(p.tableId, `Tableau budget ${p.name}`);
    }
  });

  const filename = phase === 'all' ? 'Dashboard_CAD_Complet.pdf' : `Dashboard_CAD_${phase}.pdf`;
  pdf.save(filename);
}

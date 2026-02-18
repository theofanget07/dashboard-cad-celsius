/* ============================================================
   Dashboard CAD — Groupe E Celsius
   utils.js — Fonctions utilitaires v5.3
   ============================================================ */

// ============================================================
// ROBUSTESSE COLONNES EXCEL — CORRECTION FONDAMENTALE
// ============================================================

/**
 * Normalise une chaîne : lowercase, trim, suppression accents, espaces multiples
 */
function normalize(s) {
  return String(s).toLowerCase().trim()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

/**
 * Accède à une colonne Excel de manière robuste
 * Tente chaque variante fournie (ordre, casse, accents insensibles)
 * @param {Object} row - Ligne Excel
 * @param {...string} keys - Liste des variantes de nom de colonne
 * @returns {*} Valeur de la colonne ou undefined
 */
function col(row, ...keys) {
  for (const k of keys) {
    // Essai exact
    if (row[k] !== undefined) return row[k];
    // Essai normalisé
    const kn = normalize(k);
    for (const rk of Object.keys(row)) {
      if (normalize(rk) === kn) return row[rk];
    }
  }
  return undefined;
}

// ============================================================
// CONVERSIONS
// ============================================================

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

// ============================================================
// DATES & AXES
// ============================================================

function firstOfMonth(d) { 
  return new Date(d.getFullYear(), d.getMonth(), 1); 
}

function yyyymm(d) { 
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; 
}

function fmtMonth(mm) {
  const [y, m] = mm.split('-');
  const months = ['Jan','Fév','Mar','Avr','Mai','Juin','Juil','Août','Sep','Oct','Nov','Déc'];
  return `${months[+m-1]} ${y}`;
}

function buildAxis(minD, maxD) {
  const axis = [];
  const c = firstOfMonth(new Date(minD));
  const e = firstOfMonth(new Date(maxD));
  while (c <= e) { 
    axis.push(yyyymm(c)); 
    c.setMonth(c.getMonth()+1); 
  }
  return axis;
}

// ============================================================
// EXTRACTION DONNÉES PROJET
// ============================================================

function getTaskName(row) {
  return String(col(row, 'Nom de la tâche', 'Nom de la tache', 'Tâche', 'Tache') || 'Inconnu').trim();
}

function getDateReal(row) {
  return toDate(col(row, 'Date', 'Date du projet', 'date') || null);
}

function extractPhase(idProjet) {
  if (!idProjet) return null;
  const parts = String(idProjet).split('_');
  if (parts.length < 3) return null;
  const code = parts[2];
  const mapping = { '02':'SIA2', '03':'SIA3', '04':'SIA4', '05':'SIA5' };
  return mapping[code] || null;
}

// ============================================================
// DÉTECTION HONORAIRES & EXTERNE (CORRECTIONS v5.3)
// ============================================================

/**
 * CORRECTION 3 : Filtre strict — honoraires uniquement si le mot apparaît dans le nom de tâche
 */
function isHonoraires(row) {
  const task = String(col(row, 'Nom de la tâche', 'Nom de la tache', 'Tâche', 'Tache') || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return task.includes('honoraires') || task.includes('honoraire');
}

/**
 * CORRECTION 1 : isExternal — TOUT interne par défaut
 * Externe UNIQUEMENT si le mot "externe" ou "(ext)" apparaît explicitement
 */
function isExternal(row) {
  const task = String(col(row, 'Nom de la tâche', 'Nom de la tache', 'Tâche', 'Tache') || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const desc = String(col(row, 'Description') || '').toLowerCase();
  
  // CPT toujours interne (correction explicite)
  if (task.includes('cpt')) return false;
  
  // Externe uniquement si mot explicite
  return task.includes('externe') || task.includes('(ext)') || task.includes('ext)') ||
         desc.includes('externe') || desc.includes('(ext)');
}

// ============================================================
// MÉTIERS HONORAIRES (SIA5 transversal)
// ============================================================

const METIER_KEYWORDS = [
  { key: 'Responsable Projet', patterns: ['responsable projet', 'resp. projet', 'resp projet', 'responsable de projet'] },
  { key: 'Chef de projet EE',  patterns: ['chef ee', 'chef de projet ee', 'chef ee electrique', 'chef projet ee'] },
  { key: 'Technicien',         patterns: ['technicien', 'technicienne'] },
  { key: 'Conducteur Travaux', patterns: ['conducteur travaux', 'cdt travaux', 'conducteur de travaux'] },
  { key: 'Dessinateur',        patterns: ['dessinateur', 'dessinatrice', 'dessin'] },
  { key: 'CPT / Contrôle',     patterns: ['cpt', 'contrôle des travaux', 'controle des travaux', 'contrôle travaux'] },
  { key: 'Études / Bureau',    patterns: ['études', 'etudes', 'bureau d\'études', 'ingénieur études', 'ingenieur etudes'] },
  { key: 'Expert / Conseil',   patterns: ['expert', 'conseil', 'consultant', 'consultant externe'] },
  { key: 'Architecture',       patterns: ['architecte', 'architecture'] }
];

function getMetier(row) {
  const task = String(col(row, 'Nom de la tâche', 'Nom de la tache', 'Tâche', 'Tache') || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const desc = String(col(row, 'Description') || '').toLowerCase();
  
  for (const m of METIER_KEYWORDS) {
    if (m.patterns.some(p => task.includes(p) || desc.includes(p))) return m.key;
  }
  
  // Fallback : nom de tâche tronqué
  const raw = String(col(row, 'Nom de la tâche', 'Nom de la tache', 'Tâche', 'Tache') || 'Honoraires').trim();
  return raw.length > 35 ? raw.substring(0, 33) + '…' : raw;
}

// ============================================================
// EXCEL READ
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

// ============================================================
// KPI HTML
// ============================================================

function kpiCard(label, value, unit='CHF', cls='') {
  return `<div class="kpi-card">
    <div class="kpi-label">${label}</div>
    <div class="kpi-value ${cls}">${value}</div>
    <div class="kpi-unit">${unit}</div>
  </div>`;
}

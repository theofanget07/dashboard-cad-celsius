/* ============================================================
   Dashboard CAD — Groupe E Celsius
   phases.js — Traitement des phases SIA 2/3/4/5 v5.3
   Corrections v5.3 appliquées:
   - Utilisation systématique de col() pour colonnes
   - Filtre honoraires STRICT (nom de tâche uniquement)
   - Correction isExternal (externe uniquement si explicite)
   - Tableaux ressources ajoutés
   ============================================================ */

// ============================================================
// CONFIGURATION SIA 5
// ============================================================
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

// ============================================================
// FILTRAGE PAR PHASE
// ============================================================
function filterPrevByPhase(code) {
  return prevRows.filter(r => {
    const id = String(col(r, 'ID projet', 'ID Projet', 'id projet') || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === code;
  });
}

function filterRealByPhase(code) {
  return realRows.filter(r => {
    const id = String(col(r, 'ID projet', 'ID Projet', 'id projet') || '');
    const parts = id.split('_');
    return parts.length >= 3 && parts[2] === code;
  });
}

// ============================================================
// PROCESS ALL
// ============================================================
function processAll() {
  processRecap();
  processPhase('SIA2','02');
  processPhase('SIA3','03');
  processPhase('SIA4','04');
  processSIA5Budget();
  processSIA5Imputations();
  processPlanningSIA5();
}

// ============================================================
// RÉCAPITULATIF GLOBAL
// ============================================================
function processRecap() {
  const phases = [
    { name: 'SIA 2', code: '02', color: GE_BLUE4 },
    { name: 'SIA 3', code: '03', color: GE_BLUE3 },
    { name: 'SIA 4', code: '04', color: GE_BLUE2 },
    { name: 'SIA 5', code: '05', color: GE_BLUE  }
  ];

  const summaries = phases.map(ph => {
    const pRows = filterPrevByPhase(ph.code);
    const rRows = filterRealByPhase(ph.code);
    const budgetRev= pRows.filter(r => col(r,'Modèle de prévision','Modele de prevision')==='Budget_Rev')
      .reduce((a,r) => a+toNum(col(r,'Montant total du coût','Montant')||0),0);
    const at = pRows.filter(r => col(r,'Modèle de prévision','Modele')==='AT')
      .reduce((a,r) => a+toNum(col(r,'Montant total du coût','Montant')||0),0);
    const reel = rRows.reduce((a,r) => a+toNum(col(r,'Montant total du coût','Montant')||0),0);
    const ecartBR = reel - budgetRev;
    const ecartAT = reel - at;

    const honorPRows = pRows.filter(r => isHonoraires(r));
    const honorRRows = rRows.filter(r => isHonoraires(r));
    const honorBR = honorPRows.filter(r => col(r,'Modèle de prévision')==='Budget_Rev')
      .reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
    const honorAT = honorPRows.filter(r => col(r,'Modèle')==='AT')
      .reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
    const honorReel = honorRRows.reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);

    const allDates = [...pRows, ...rRows].flatMap(r => [
      toDate(col(r,'Date de début')),
      toDate(col(r,'Date de fin')),
      getDateReal(r)
    ]).filter(Boolean);
    const planStart = allDates.length ? new Date(Math.min(...allDates)) : null;
    const planEnd   = allDates.length ? new Date(Math.max(...allDates)) : null;

    return { ...ph, budgetRev, at, reel, ecartBR, ecartAT, honorBR, honorAT, honorReel, planStart, planEnd };
  });

  const totalBR   = summaries.reduce((a,s) => a+s.budgetRev, 0);
  const totalAT   = summaries.reduce((a,s) => a+s.at, 0);
  const totalReel = summaries.reduce((a,s) => a+s.reel, 0);
  const totalEcart= totalReel - totalAT;
  
  const kpiEl = document.getElementById('recapBudgetKPI');
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Budget Révisé Total', toCHF(totalBR)) +
    kpiCard('AT Total', toCHF(totalAT)) +
    kpiCard('Réel Total', toCHF(totalReel)) +
    kpiCard('Écart vs AT', (totalEcart>=0?'+':'')+toCHF(totalEcart), 'CHF', totalEcart>=0?'negative':'positive');

  destroyChart('recapBudgetChart');
  const rcCtx = document.getElementById('recapBudgetChart');
  if (rcCtx) {
    const nonZero = summaries.filter(s => s.budgetRev > 0 || s.at > 0 || s.reel > 0);
    if (nonZero.length) {
      charts['recapBudgetChart'] = new Chart(rcCtx, {
        type: 'bar',
        data: {
          labels: nonZero.map(s => s.name),
          datasets: [
            { label: 'Budget Révisé', data: nonZero.map(s=>s.budgetRev), backgroundColor: GE_ORANGE+'cc' },
            { label: 'AT',            data: nonZero.map(s=>s.at),        backgroundColor: GE_BLUE2+'cc' },
            { label: 'Réel',          data: nonZero.map(s=>s.reel),      backgroundColor: GE_RED+'cc' }
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { position:'top' } },
          scales: { y: { ticks: { callback: v => toCHF(v) } } }
        }
      });
    }
  }

  const tbl = document.getElementById('recapBudgetTable');
  if (tbl) {
    let html = `<thead><tr><th>Phase</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th><th>% AT</th></tr></thead><tbody>`;
    summaries.forEach(s => {
      const pAT = s.at>0 ? (s.reel/s.at*100).toFixed(1) : 0;
      html += `<tr>
        <td><strong>${s.name}</strong></td>
        <td>${toCHF(s.budgetRev)}</td><td>${toCHF(s.at)}</td><td>${toCHF(s.reel)}</td>
        <td class="${s.ecartBR>=0?'negative':'positive'}">${(s.ecartBR>=0?'+':'')+toCHF(s.ecartBR)}</td>
        <td class="${s.ecartAT>=0?'negative':'positive'}">${(s.ecartAT>=0?'+':'')+toCHF(s.ecartAT)}</td>
        <td>${pAT}%</td>
      </tr>`;
    });
    html += `</tbody><tfoot><tr>
      <td>TOTAL</td>
      <td>${toCHF(totalBR)}</td><td>${toCHF(totalAT)}</td><td>${toCHF(totalReel)}</td>
      <td class="${(totalReel-totalBR)>=0?'negative':'positive'}">${((totalReel-totalBR)>=0?'+':'')+toCHF(totalReel-totalBR)}</td>
      <td class="${totalEcart>=0?'negative':'positive'}">${(totalEcart>=0?'+':'')+toCHF(totalEcart)}</td>
      <td>${totalAT>0?(totalReel/totalAT*100).toFixed(1):0}%</td>
    </tr></tfoot>`;
    tbl.innerHTML = html;
  }

  // Imputations récap
  const totHonorBR   = summaries.reduce((a,s)=>a+s.honorBR,0);
  const totHonorAT   = summaries.reduce((a,s)=>a+s.honorAT,0);
  const totHonorReel = summaries.reduce((a,s)=>a+s.honorReel,0);
  const kpiImpEl = document.getElementById('recapImputKPI');
  if (kpiImpEl) kpiImpEl.innerHTML =
    kpiCard('Honoraires BR', toCHF(totHonorBR)) +
    kpiCard('Honoraires AT', toCHF(totHonorAT)) +
    kpiCard('Honoraires Réel', toCHF(totHonorReel)) +
    kpiCard('Écart Honor. AT', ((totHonorReel-totHonorAT)>=0?'+':'')+toCHF(totHonorReel-totHonorAT), 'CHF', (totHonorReel-totHonorAT)>=0?'negative':'positive');

  destroyChart('recapImputChart');
  const riCtx = document.getElementById('recapImputChart');
  if (riCtx) {
    const nonZeroH = summaries.filter(s => s.honorBR>0 || s.honorAT>0 || s.honorReel>0);
    if (nonZeroH.length) {
      charts['recapImputChart'] = new Chart(riCtx, {
        type: 'bar',
        data: {
          labels: nonZeroH.map(s=>s.name),
          datasets: [
            { label: 'Honoraires BR', data: nonZeroH.map(s=>s.honorBR),   backgroundColor: GE_ORANGE+'cc' },
            { label: 'Honoraires AT', data: nonZeroH.map(s=>s.honorAT),   backgroundColor: GE_BLUE2+'cc' },
            { label: 'Réel',          data: nonZeroH.map(s=>s.honorReel), backgroundColor: GE_GREEN+'cc'  }
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend: { position:'top' } },
          scales: { y: { ticks: { callback: v => toCHF(v) } } }
        }
      });
    }
  }

  const tblI = document.getElementById('recapImputTable');
  if (tblI) {
    let html = `<thead><tr><th>Phase</th><th>Honoraires BR</th><th>Honoraires AT</th><th>Réel</th><th>Écart AT</th></tr></thead><tbody>`;
    summaries.forEach(s => {
      const ec = s.honorReel - s.honorAT;
      html += `<tr><td><strong>${s.name}</strong></td><td>${toCHF(s.honorBR)}</td><td>${toCHF(s.honorAT)}</td><td>${toCHF(s.honorReel)}</td>
        <td class="${ec>=0?'negative':'positive'}">${(ec>=0?'+':'')+toCHF(ec)}</td></tr>`;
    });
    html += `</tbody>`;
    tblI.innerHTML = html;
  }

  // Planning récap (Gantt global)
  destroyChart('recapPlanningChart');
  const ganttCtx = document.getElementById('recapPlanningChart');
  if (ganttCtx) {
    const validPhases = summaries.filter(s => s.planStart && s.planEnd);
    if (validPhases.length) {
      const allStart = new Date(Math.min(...validPhases.map(s=>s.planStart)));
      const allEnd   = new Date(Math.max(...validPhases.map(s=>s.planEnd)));
      const axis     = buildAxis(allStart, allEnd);
      const bars     = validPhases.map(s => {
        const si = axis.indexOf(yyyymm(firstOfMonth(s.planStart)));
        const ei = axis.indexOf(yyyymm(firstOfMonth(s.planEnd)));
        const base = si >= 0 ? si : 0;
        const len  = ei >= si ? ei - si + 1 : 1;
        return { label: s.name, base, len, color: s.color };
      });
      charts['recapPlanningChart'] = new Chart(ganttCtx, {
        type: 'bar',
        data: {
          labels: validPhases.map(s=>s.name),
          datasets: [
            { label:'Début', data: bars.map(b=>b.base), backgroundColor: 'transparent', borderWidth: 0 },
            { label:'Durée', data: bars.map(b=>b.len), backgroundColor: validPhases.map(s=>s.color+'cc'), borderRadius: 4 }
          ]
        },
        options: {
          indexAxis: 'y', responsive: true, maintainAspectRatio: false,
          plugins: { legend: { display: false }, todayLine: { label: yyyymm(new Date()) } },
          scales: {
            x: { stacked: true, ticks: { callback: v => axis[v] ? fmtMonth(axis[v]) : '' } },
            y: { stacked: true }
          }
        }
      });
      ganttCtx.parentElement.style.height = `${Math.max(200, validPhases.length*80)}px`;
    }
  }
}

// ============================================================
// PHASES SIA 2/3/4 — GÉNÉRIQUE
// ============================================================
function processPhase(phaseName, phaseCode) {
  const pRows = filterPrevByPhase(phaseCode);
  const rRows = filterRealByPhase(phaseCode);
  const phaseNum = phaseCode.replace('0','');

  if (!pRows.length && !rRows.length) {
    const kpiEl = document.getElementById(`kpiSection${phaseName}`);
    if (kpiEl) kpiEl.innerHTML = `<div class="hint">ℹ️ Aucune donnée pour ${phaseName}</div>`;
    return;
  }

  const sData = buildSCurves(pRows, rRows);
  drawSCurveChart(`chart${phaseName}`, sData, phaseName);
  renderMonthlyTable(`table${phaseName}Monthly`, sData);

  const budgetRev = pRows.filter(r => col(r,'Modèle de prévision')==='Budget_Rev')
    .reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
  const at = pRows.filter(r => col(r,'Modèle de prévision')==='AT')
    .reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
  const reel = rRows.reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
  const ecart = reel - at;
  const pct = at > 0 ? (reel/at*100).toFixed(1) : 0;

  const kpiEl = document.getElementById(`kpiSection${phaseName}`);
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Budget Révisé', toCHF(budgetRev)) +
    kpiCard('AT', toCHF(at)) +
    kpiCard('Réel', toCHF(reel)) +
    kpiCard('Écart vs AT', (ecart>=0?'+':'')+toCHF(ecart), 'CHF', ecart>=0?'negative':'positive') +
    kpiCard('Avancement', `${pct}%`, '% AT', '');

  const taskMap = {};
  pRows.forEach(r => {
    const task   = getTaskName(r);
    const model  = col(r,'Modèle de prévision');
    const amount = toNum(col(r,'Montant total du coût')||0);
    if (!taskMap[task]) taskMap[task] = { task, budget:0, budgetRev:0, at:0, reel:0 };
    if (model === 'Budget')     taskMap[task].budget += amount;
    if (model === 'Budget_Rev') taskMap[task].budgetRev += amount;
    if (model === 'AT')         taskMap[task].at += amount;
  });
  rRows.forEach(r => {
    const task = getTaskName(r);
    if (!taskMap[task]) taskMap[task] = { task, budget:0, budgetRev:0, at:0, reel:0 };
    taskMap[task].reel += toNum(col(r,'Montant total du coût')||0);
  });
  const taskList = Object.values(taskMap).filter(t => t.budget>0 || t.budgetRev>0 || t.at>0 || t.reel>0);

  const tblBudget = document.getElementById(`table${phaseName}Budget`);
  if (tblBudget) {
    let html = `<thead><tr><th>Tâche</th><th>Budget</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th><th>% AT</th></tr></thead><tbody>`;
    taskList.forEach(t => {
      const eBR = t.reel - t.budgetRev, eAT = t.reel - t.at;
      const pAT = t.at>0 ? (t.reel/t.at*100).toFixed(1) : '—';
      html += `<tr>
        <td><strong>${t.task}</strong></td>
        <td>${toCHF(t.budget)}</td><td>${toCHF(t.budgetRev)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
        <td class="${eBR>=0?'negative':'positive'}">${eBR!==0?(eBR>=0?'+':'')+toCHF(eBR):'—'}</td>
        <td class="${eAT>=0?'negative':'positive'}">${eAT!==0?(eAT>=0?'+':'')+toCHF(eAT):'—'}</td>
        <td>${pAT}%</td>
      </tr>`;
    });
    html += `</tbody><tfoot><tr>
      <td>TOTAL</td>
      <td>${toCHF(taskList.reduce((a,t)=>a+t.budget,0))}</td>
      <td>${toCHF(budgetRev)}</td><td>${toCHF(at)}</td><td>${toCHF(reel)}</td>
      <td class="${(reel-budgetRev)>=0?'negative':'positive'}">${((reel-budgetRev)>=0?'+':'')+toCHF(reel-budgetRev)}</td>
      <td class="${ecart>=0?'negative':'positive'}">${(ecart>=0?'+':'')+toCHF(ecart)}</td>
      <td>${pct}%</td>
    </tr></tfoot>`;
    tblBudget.innerHTML = html;
  }

  const pieBRId = `pie${phaseName}BR`;
  destroyChart(pieBRId);
  const pieBRCtx = document.getElementById(pieBRId);
  if (pieBRCtx) {
    const validTasks = taskList.filter(t => t.budgetRev > 0);
    if (validTasks.length) {
      charts[pieBRId] = new Chart(pieBRCtx, {
        type: 'doughnut',
        data: {
          labels: validTasks.map(t => t.task.length>22?t.task.substring(0,20)+'…':t.task),
          datasets: [{
            data: validTasks.map(t => t.budgetRev),
            backgroundColor: validTasks.map((_,i) => PIE_COLORS[i % PIE_COLORS.length]),
            borderWidth: 2, borderColor: '#fff'
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: {
            legend: { position:'right', labels: { font:{size:11}, usePointStyle:true } },
            tooltip: { callbacks: { label: ctx => ` ${ctx.label}: CHF ${toCHF(ctx.parsed)}` } }
          }
        }
      });
    }
  }

  const barId = `bar${phaseName}`;
  destroyChart(barId);
  const barCtx = document.getElementById(barId);
  if (barCtx) {
    const validBar = taskList.filter(t => t.at>0 || t.reel>0);
    if (validBar.length) {
      charts[barId] = new Chart(barCtx, {
        type: 'bar',
        data: {
          labels: validBar.map(t => t.task.length>20 ? t.task.substring(0,18)+'…' : t.task),
          datasets: [
            { label:'AT',   data: validBar.map(t=>t.at),   backgroundColor: GE_BLUE2+'cc' },
            { label:'Réel', data: validBar.map(t=>t.reel), backgroundColor: GE_RED+'cc' }
          ]
        },
        options: {
          indexAxis:'y', responsive:true, maintainAspectRatio:false,
          plugins: { legend:{position:'top'} },
          scales: { x:{ ticks:{ callback: v => toCHF(v) } } }
        }
      });
    }
  }

  processImputationsPhase(phaseName, phaseCode, pRows, rRows);
  processPlanningPhase(phaseName, phaseCode, pRows, rRows);
}

// ============================================================
// IMPUTATIONS PAR PHASE (SIA2/3/4)
// CORRECTION v5.3 : filtre honoraires STRICT
// ============================================================
function processImputationsPhase(phaseName, phaseCode, pRows, rRows) {
  const phaseNum = phaseCode.replace('0','');

  // CORRECTION : Filtre strict honoraires
  const honorPrev = pRows.filter(r => isHonoraires(r));
  const honorReal = rRows.filter(r => isHonoraires(r));

  const taskMap = {};
  honorPrev.forEach(r => {
    const task  = getTaskName(r);
    const model = col(r,'Modèle de prévision');
    const ext   = isExternal(r);
    const key   = `${task}__${ext?'ext':'int'}`;
    if (!taskMap[key]) taskMap[key] = { task, ext, br:0, at:0, reel:0 };
    if (model==='Budget_Rev') taskMap[key].br += toNum(col(r,'Montant total du coût')||0);
    if (model==='AT')         taskMap[key].at += toNum(col(r,'Montant total du coût')||0);
  });
  honorReal.forEach(r => {
    const task = getTaskName(r);
    const ext  = isExternal(r);
    const key  = `${task}__${ext?'ext':'int'}`;
    if (!taskMap[key]) taskMap[key] = { task, ext, br:0, at:0, reel:0 };
    taskMap[key].reel += toNum(col(r,'Montant total du coût')||0);
  });

  const items = Object.values(taskMap).filter(t => t.br>0 || t.at>0 || t.reel>0);
  const intItems = items.filter(t => !t.ext);
  const extItems = items.filter(t => t.ext);

  const totBR   = items.reduce((a,t) => a+t.br,   0);
  const totAT   = items.reduce((a,t) => a+t.at,   0);
  const totReel = items.reduce((a,t) => a+t.reel, 0);
  const ecartAT = totReel - totAT;

  const kpiImpEl = document.getElementById(`kpiSectionImpSIA${phaseNum}`);
  if (kpiImpEl) kpiImpEl.innerHTML =
    kpiCard('Honoraires BR', toCHF(totBR)) +
    kpiCard('Honoraires AT', toCHF(totAT)) +
    kpiCard('Réel Honoraires', toCHF(totReel)) +
    kpiCard('Écart AT', (ecartAT>=0?'+':'')+toCHF(ecartAT), 'CHF', ecartAT>=0?'negative':'positive');

  const tbl = document.getElementById(`tableImputSIA${phaseNum}`) || document.getElementById(`tableImpSIA${phaseNum}`);
  if (!tbl) return;

  let html = `<thead><tr><th>Tâche</th><th>Type</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th></tr></thead><tbody>`;

  const renderSection = (list, type, cls, clsSub) => {
    if (!list.length) return '';
    let rows = `<tr class="${cls}"><td colspan="7">${type === 'int' ? '🔵 RESSOURCES INTERNES' : '🟡 RESSOURCES EXTERNES'}</td></tr>`;
    let sBR=0, sAT=0, sReel=0;
    list.forEach(t => {
      const eBR = t.reel-t.br, eAT = t.reel-t.at;
      sBR+=t.br; sAT+=t.at; sReel+=t.reel;
      rows += `<tr class="${type==='int'?'row-int':'row-ext'}">
        <td>${t.task}</td>
        <td><span class="badge ${type==='int'?'prevu':'en-cours'}">${type==='int'?'Interne':'Externe'}</span></td>
        <td>${toCHF(t.br)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
        <td class="${eBR>=0?'negative':'positive'}">${(eBR>=0?'+':'')+toCHF(eBR)}</td>
        <td class="${eAT>=0?'negative':'positive'}">${(eAT>=0?'+':'')+toCHF(eAT)}</td>
      </tr>`;
    });
    const sE = sReel-sAT;
    rows += `<tr class="${clsSub}">
      <td colspan="2"><strong>Sous-total ${type==='int'?'Interne':'Externe'}</strong></td>
      <td>${toCHF(sBR)}</td><td>${toCHF(sAT)}</td><td>${toCHF(sReel)}</td>
      <td class="${(sReel-sBR)>=0?'negative':'positive'}">${((sReel-sBR)>=0?'+':'')+toCHF(sReel-sBR)}</td>
      <td class="${sE>=0?'negative':'positive'}">${(sE>=0?'+':'')+toCHF(sE)}</td>
    </tr>`;
    return rows;
  };

  html += renderSection(intItems, 'int', 'section-header-int', 'subtotal-int');
  html += renderSection(extItems, 'ext', 'section-header-ext', 'subtotal-ext');
  const totalEcatAT = totReel - totAT;
  html += `</tbody><tfoot><tr>
    <td colspan="2"><strong>TOTAL HONORAIRES</strong></td>
    <td>${toCHF(totBR)}</td><td>${toCHF(totAT)}</td><td>${toCHF(totReel)}</td>
    <td class="${(totReel-totBR)>=0?'negative':'positive'}">${((totReel-totBR)>=0?'+':'')+toCHF(totReel-totBR)}</td>
    <td class="${totalEcatAT>=0?'negative':'positive'}">${(totalEcatAT>=0?'+':'')+toCHF(totalEcatAT)}</td>
  </tr></tfoot>`;
  tbl.innerHTML = html;
}

// ============================================================
// PLANNING PHASES SIA 2/3/4
// ============================================================
function processPlanningPhase(phaseName, phaseCode, pRows, rRows) {
  const ganttId = `gantt${phaseName}`;
  destroyChart(ganttId);
  const ctx = document.getElementById(ganttId);
  if (!ctx) return;

  const taskNames = [...new Set(pRows.map(r => getTaskName(r)))];
  if (!taskNames.length) return;

  const ganttData = taskNames.map(task => {
    const taskPrev = pRows.filter(r => getTaskName(r) === task);
    const getBR = taskPrev.filter(r => col(r,'Modèle de prévision')==='Budget_Rev');
    const getAT = taskPrev.filter(r => col(r,'Modèle')==='AT');
    const realT = rRows.filter(r => getTaskName(r) === task);

    const getDates = rows => {
      const dates = rows.flatMap(r => [
        toDate(col(r,'Date de début')),
        toDate(col(r,'Date de fin')),
        getDateReal(r)
      ]).filter(Boolean);
      return dates.length ? { start: new Date(Math.min(...dates)), end: new Date(Math.max(...dates)) } : null;
    };

    return {
      task,
      br:   getDates(getBR),
      at:   getDates(getAT),
      reel: getDates(realT)
    };
  }).filter(t => t.br || t.at || t.reel);

  if (!ganttData.length) return;

  const allDates = ganttData.flatMap(t => [t.br?.start, t.br?.end, t.at?.start, t.at?.end, t.reel?.start, t.reel?.end]).filter(Boolean);
  const axisStart = new Date(Math.min(...allDates));
  const axisEnd   = new Date(Math.max(...allDates));
  const axis = buildAxis(axisStart, axisEnd);

  const dateToIdx = d => {
    if (!d) return -1;
    const mm = yyyymm(firstOfMonth(d));
    const i = axis.indexOf(mm);
    return i >= 0 ? i : 0;
  };

  const buildBar = dateObj => {
    if (!dateObj) return [null, null];
    const si = dateToIdx(dateObj.start), ei = dateToIdx(dateObj.end);
    return [si, ei - si + 1];
  };

  const labels    = ganttData.map(t => t.task.length>22 ? t.task.substring(0,22)+'…' : t.task);
  const brBase    = ganttData.map(t => buildBar(t.br)[0]);
  const brLen     = ganttData.map(t => buildBar(t.br)[1]);
  const atBase    = ganttData.map(t => buildBar(t.at)[0]);
  const atLen     = ganttData.map(t => buildBar(t.at)[1]);
  const reelBase  = ganttData.map(t => buildBar(t.reel)[0]);
  const reelLen   = ganttData.map(t => buildBar(t.reel)[1]);

  charts[ganttId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'BR offset',   data:brBase,   backgroundColor:'transparent', borderWidth:0, stack:'br'   },
        { label:'Budget Révisé', data:brLen,  backgroundColor:GE_ORANGE+'cc', borderRadius:3, barThickness:12, stack:'br' },
        { label:'AT offset',   data:atBase,   backgroundColor:'transparent', borderWidth:0, stack:'at'   },
        { label:'AT',          data:atLen,    backgroundColor:GE_BLUE2+'cc',  borderRadius:3, barThickness:12, stack:'at' },
        { label:'Réel offset', data:reelBase, backgroundColor:'transparent', borderWidth:0, stack:'reel' },
        { label:'Réel',        data:reelLen,  backgroundColor:GE_RED+'cc',   borderRadius:3, barThickness:12, stack:'reel' }
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: {
        legend: {
          labels: {
            filter: item => !item.text.includes('offset'),
            usePointStyle:true, font:{size:11}
          }
        },
        todayLine: { label: yyyymm(new Date()) }
      },
      scales: {
        x: { stacked:true, ticks: { callback: v => axis[v] ? fmtMonth(axis[v]) : '' }, font:{size:10} },
        y: { stacked:true }
      }
    }
  });
  ctx.parentElement.style.height = `${Math.max(250, ganttData.length * 55 + 80)}px`;
}

// ============================================================
// SIA 5 — BUDGET
// ============================================================
function processSIA5Budget() {
  const pRows = filterPrevByPhase('05');
  const rRows = filterRealByPhase('05');

  if (!pRows.length && !rRows.length) return;

  const sData = buildSCurves(pRows, rRows);
  drawSCurveChart('chartSIA5', sData, 'SIA5');
  renderMonthlyTable('tableSIA5Monthly', sData);

  const budgetRev = pRows.filter(r => col(r,'Modèle de prévision')==='Budget_Rev')
    .reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
  const at = pRows.filter(r => col(r,'Modèle')==='AT')
    .reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
  const reel = rRows.reduce((a,r) => a+toNum(col(r,'Montant total du coût')||0),0);
  const ecart = reel - at;
  const pct = at > 0 ? (reel/at*100).toFixed(1) : 0;

  const kpiEl = document.getElementById('kpiSectionSIA5');
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Budget Révisé', toCHF(budgetRev)) +
    kpiCard('AT', toCHF(at)) +
    kpiCard('Réel', toCHF(reel)) +
    kpiCard('Écart vs AT', (ecart>=0?'+':'')+toCHF(ecart), 'CHF', ecart>=0?'negative':'positive') +
    kpiCard('Avancement', `${pct}%`, '% AT', '');

  const tblBudget = document.getElementById('tableSIA5Budget');
  if (tblBudget) {
    const taskMap = {};
    PROJECT_ORDER.forEach(id => { taskMap[id] = { budget:0, budgetRev:0, at:0, reel:0 }; });
    pRows.forEach(r => {
      const id = String(col(r,'ID projet','ID Projet')||'');
      if (!taskMap[id]) taskMap[id] = { budget:0, budgetRev:0, at:0, reel:0 };
      const model = col(r,'Modèle de prévision'), amt = toNum(col(r,'Montant total du coût')||0);
      if (model==='Budget')     taskMap[id].budget += amt;
      if (model==='Budget_Rev') taskMap[id].budgetRev += amt;
      if (model==='AT')         taskMap[id].at += amt;
    });
    rRows.forEach(r => {
      const id = String(col(r,'ID projet','ID Projet')||'');
      if (!taskMap[id]) taskMap[id] = { budget:0, budgetRev:0, at:0, reel:0 };
      taskMap[id].reel += toNum(col(r,'Montant total du coût')||0);
    });

    let html = `<thead><tr><th>Sous-Projet</th><th>Budget</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th><th>% AT</th></tr></thead><tbody>`;

    Object.entries(taskMap).forEach(([id, t]) => {
      if (t.budget===0 && t.budgetRev===0 && t.at===0 && t.reel===0) return;
      const eBR=t.reel-t.budgetRev, eAT=t.reel-t.at;
      const pAT=t.at>0?(t.reel/t.at*100).toFixed(1):'—';
      html += `<tr>
        <td><strong>${getPN(id)}</strong></td>
        <td>${toCHF(t.budget)}</td><td>${toCHF(t.budgetRev)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
        <td class="${eBR>=0?'negative':'positive'}">${(eBR>=0?'+':'')+toCHF(eBR)}</td>
        <td class="${eAT>=0?'negative':'positive'}">${(eAT>=0?'+':'')+toCHF(eAT)}</td>
        <td>${pAT}%</td>
      </tr>`;
    });
    html += `</tbody><tfoot><tr>
      <td>TOTAL</td>
      <td>${toCHF(pRows.filter(r=>col(r,'Modèle')==='Budget').reduce((a,r)=>a+toNum(col(r,'Montant')||0),0))}</td>
      <td>${toCHF(budgetRev)}</td><td>${toCHF(at)}</td><td>${toCHF(reel)}</td>
      <td class="${(reel-budgetRev)>=0?'negative':'positive'}">${((reel-budgetRev)>=0?'+':'')+toCHF(reel-budgetRev)}</td>
      <td class="${ecart>=0?'negative':'positive'}">${(ecart>=0?'+':'')+toCHF(ecart)}</td>
      <td>${pct}%</td>
    </tr></tfoot>`;
    tblBudget.innerHTML = html;
  }

  destroyChart('pieSIA5BR');
  const pieSIA5Ctx = document.getElementById('pieSIA5BR');
  if (pieSIA5Ctx) {
    const pieData = PROJECT_ORDER.map(id => ({
      label: getPN(id),
      value: pRows.filter(r => (col(r,'ID projet')||col(r,'ID Projet'))===id && col(r,'Modèle')==='Budget_Rev')
                  .reduce((a,r)=>a+toNum(col(r,'Montant')||0),0)
    })).filter(d => d.value > 0);

    if (pieData.length) {
      charts['pieSIA5BR'] = new Chart(pieSIA5Ctx, {
        type: 'doughnut',
        data: {
          labels: pieData.map(d=>d.label),
          datasets: [{
            data: pieData.map(d=>d.value),
            backgroundColor: pieData.map((_,i)=>PIE_COLORS[i%PIE_COLORS.length]),
            borderWidth:2, borderColor:'#fff'
          }]
        },
        options: {
          responsive:true, maintainAspectRatio:false,
          plugins: {
            legend: { position:'right', labels:{ font:{size:11}, usePointStyle:true } },
            tooltip: { callbacks: { label: ctx => ` ${ctx.label}: CHF ${toCHF(ctx.parsed)}` } }
          }
        }
      });
    }
  }

  destroyChart('barSIA5');
  const barSIA5Ctx = document.getElementById('barSIA5');
  if (barSIA5Ctx) {
    const barData = PROJECT_ORDER.map(id => {
      const atV  = pRows.filter(r=>(col(r,'ID projet'))===id && col(r,'Modèle')==='AT').reduce((a,r)=>a+toNum(col(r,'Montant')||0),0);
      const reelV = rRows.filter(r=>(col(r,'ID projet'))===id).reduce((a,r)=>a+toNum(col(r,'Montant')||0),0);
      return { label: getPN(id), at: atV, reel: reelV };
    }).filter(d => d.at>0 || d.reel>0);

    if (barData.length) {
      charts['barSIA5'] = new Chart(barSIA5Ctx, {
        type:'bar',
        data:{
          labels: barData.map(d=>d.label),
          datasets:[
            { label:'AT',   data:barData.map(d=>d.at),   backgroundColor:GE_BLUE2+'cc' },
            { label:'Réel', data:barData.map(d=>d.reel), backgroundColor:GE_RED+'cc'   }
          ]
        },
        options:{
          indexAxis:'y', responsive:true, maintainAspectRatio:false,
          plugins:{ legend:{position:'top'} },
          scales:{ x:{ ticks:{ callback:v=>toCHF(v) } } }
        }
      });
    }
  }
}

// ============================================================
// SIA 5 — IMPUTATIONS (CORRECTION filtre strict)
// ============================================================
function processSIA5Imputations() {
  const pRows = filterPrevByPhase('05');
  const rRows = filterRealByPhase('05');

  // CORRECTION : filtre strict sur nom de tâche
  const honorPrev = pRows.filter(r => isHonoraires(r));
  const honorReal = rRows.filter(r => isHonoraires(r));

  const metierMap = {};
  honorPrev.forEach(r => {
    const metier = getMetier(r);
    const model  = col(r,'Modèle de prévision');
    const ext    = isExternal(r);
    const key    = `${metier}__${ext?'ext':'int'}`;
    if (!metierMap[key]) metierMap[key] = { metier, ext, br:0, at:0, reel:0 };
    if (model==='Budget_Rev') metierMap[key].br   += toNum(col(r,'Montant total du coût')||0);
    if (model==='AT')         metierMap[key].at   += toNum(col(r,'Montant total du coût')||0);
  });
  honorReal.forEach(r => {
    const metier = getMetier(r);
    const ext    = isExternal(r);
    const key    = `${metier}__${ext?'ext':'int'}`;
    if (!metierMap[key]) metierMap[key] = { metier, ext, br:0, at:0, reel:0 };
    metierMap[key].reel += toNum(col(r,'Montant total du coût')||0);
  });

  const items    = Object.values(metierMap).filter(t => t.br>0 || t.at>0 || t.reel>0);
  const intItems = items.filter(t => !t.ext);
  const extItems = items.filter(t => t.ext);

  const totBR   = items.reduce((a,t)=>a+t.br,   0);
  const totAT   = items.reduce((a,t)=>a+t.at,   0);
  const totReel = items.reduce((a,t)=>a+t.reel,  0);
  const ecartAT = totReel - totAT;

  const kpiEl = document.getElementById('kpiSectionImpSIA5');
  if (kpiEl) kpiEl.innerHTML =
    kpiCard('Honoraires BR',    toCHF(totBR)) +
    kpiCard('Honoraires AT',    toCHF(totAT)) +
    kpiCard('Réel Honoraires',  toCHF(totReel)) +
    kpiCard('Écart vs AT', (ecartAT>=0?'+':'')+toCHF(ecartAT), 'CHF', ecartAT>=0?'negative':'positive');

  const tbl = document.getElementById('tableImputSIA5');
  if (tbl) {
    let html = `<thead><tr><th>Métier</th><th>Type</th><th>Budget Révisé</th><th>AT</th><th>Réel</th><th>Écart BR</th><th>Écart AT</th></tr></thead><tbody>`;

    const renderSection = (list, type, cls, clsSub) => {
      if (!list.length) return '';
      let rows = `<tr class="${cls}"><td colspan="7">${type==='int'?'🔵 INTERNES':'🟡 EXTERNES'}</td></tr>`;
      let sBR=0, sAT=0, sReel=0;
      list.forEach(t => {
        const eBR=t.reel-t.br, eAT=t.reel-t.at;
        sBR+=t.br; sAT+=t.at; sReel+=t.reel;
        rows += `<tr class="${type==='int'?'row-int':'row-ext'}">
          <td>${t.metier}</td>
          <td><span class="badge ${type==='int'?'prevu':'en-cours'}">${type==='int'?'Interne':'Externe'}</span></td>
          <td>${toCHF(t.br)}</td><td>${toCHF(t.at)}</td><td>${toCHF(t.reel)}</td>
          <td class="${eBR>=0?'negative':'positive'}">${(eBR>=0?'+':'')+toCHF(eBR)}</td>
          <td class="${eAT>=0?'negative':'positive'}">${(eAT>=0?'+':'')+toCHF(eAT)}</td>
        </tr>`;
      });
      const sE=sReel-sAT;
      rows += `<tr class="${clsSub}">
        <td colspan="2"><strong>Sous-total ${type==='int'?'Interne':'Externe'}</strong></td>
        <td>${toCHF(sBR)}</td><td>${toCHF(sAT)}</td><td>${toCHF(sReel)}</td>
        <td class="${(sReel-sBR)>=0?'negative':'positive'}">${((sReel-sBR)>=0?'+':'')+toCHF(sReel-sBR)}</td>
        <td class="${sE>=0?'negative':'positive'}">${(sE>=0?'+':'')+toCHF(sE)}</td>
      </tr>`;
      return rows;
    };

    html += renderSection(intItems, 'int', 'section-header-int', 'subtotal-int');
    html += renderSection(extItems, 'ext', 'section-header-ext', 'subtotal-ext');
    html += `</tbody><tfoot><tr>
      <td colspan="2"><strong>TOTAL HONORAIRES</strong></td>
      <td>${toCHF(totBR)}</td><td>${toCHF(totAT)}</td><td>${toCHF(totReel)}</td>
      <td class="${(totReel-totBR)>=0?'negative':'positive'}">${((totReel-totBR)>=0?'+':'')+toCHF(totReel-totBR)}</td>
      <td class="${ecartAT>=0?'negative':'positive'}">${(ecartAT>=0?'+':'')+toCHF(ecartAT)}</td>
    </tr></tfoot>`;
    tbl.innerHTML = html;
  }

  destroyChart('barImputSIA5');
  const barImputCtx = document.getElementById('barImputSIA5');
  if (barImputCtx && items.length) {
    const sorted = [...items].sort((a,b) => (b.at||b.br) - (a.at||a.br));
    charts['barImputSIA5'] = new Chart(barImputCtx, {
      type: 'bar',
      data: {
        labels: sorted.map(t => `${t.metier}${t.ext?' (ext)':''}`),
        datasets: [
          { label: 'Budget Révisé', data: sorted.map(t=>t.br),   backgroundColor: GE_ORANGE+'cc', borderRadius: 3 },
          { label: 'AT',            data: sorted.map(t=>t.at),   backgroundColor: GE_BLUE2+'cc',  borderRadius: 3 },
          { label: 'Réel',          data: sorted.map(t=>t.reel), backgroundColor: GE_RED+'cc',    borderRadius: 3 }
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { usePointStyle: true, font: { size: 11 } } },
          tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: CHF ${toCHF(ctx.parsed.x)}` } }
        },
        scales: {
          x: { ticks: { callback: v => toCHF(v), font: { size: 10 } }, grid: { color: '#e8eef5' } },
          y: { ticks: { font: { size: 11 } } }
        }
      }
    });
    barImputCtx.parentElement.style.height = `${Math.max(200, sorted.length * 50 + 80)}px`;
  }
}

// ============================================================
// SIA 5 — PLANNING
// ============================================================
function processPlanningSIA5() {
  const pRows = filterPrevByPhase('05');
  const rRows = filterRealByPhase('05');

  const planData = PROJECT_ORDER.map(id => {
    const brRows   = pRows.filter(r => (col(r,'ID projet')||col(r,'ID Projet'))===id && col(r,'Modèle')==='Budget_Rev');
    const atRows   = pRows.filter(r => (col(r,'ID projet'))===id && col(r,'Modèle')==='AT');
    const reelRows = rRows.filter(r => (col(r,'ID projet'))===id);

    const getDates = rows => {
      const dates = rows.flatMap(r => [
        toDate(col(r,'Date de début')),
        toDate(col(r,'Date de fin')),
        getDateReal(r)
      ]).filter(Boolean);
      return dates.length ? { start: new Date(Math.min(...dates)), end: new Date(Math.max(...dates)) } : null;
    };

    const manual = manualPlanningDates[id] || {};
    const brBase   = getDates(brRows);
    const atBase   = getDates(atRows);
    const reelBase = getDates(reelRows);

    return {
      id, name: getPN(id),
      br:   { start: manual.brStart || brBase?.start   || null, end: manual.brEnd || brBase?.end   || null },
      at:   { start: manual.atStart || atBase?.start   || null, end: manual.atEnd || atBase?.end   || null },
      reel: { start: reelBase?.start || null, end: reelBase?.end || null }
    };
  });

  const tblPlan = document.getElementById('tablePlanningSIA5');
  if (tblPlan) {
    let html = `<thead><tr>
      <th>Sous-Projet</th>
      <th>Projeté BR Début</th><th>Projeté BR Fin</th>
      <th>Projeté AT Début</th><th>Projeté AT Fin</th>
      <th>Réel Début</th><th>Réel Fin</th>
      <th>Statut</th>
    </tr></thead><tbody>`;

    planData.forEach(p => {
      const today = new Date();
      let statut = '', statusClass = '';
      if (p.reel.start && p.reel.end && p.reel.end < today) {
        statut = 'Terminé'; statusClass = 'termine';
      } else if (p.reel.start && p.reel.start <= today) {
        if (p.at.end && today > p.at.end) { statut = 'Retard'; statusClass = 'retard'; }
        else { statut = 'En cours'; statusClass = 'en-cours'; }
      } else {
        statut = 'Prévu'; statusClass = 'prevu';
      }

      html += `<tr>
        <td><strong>${p.name}</strong></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="brStart" value="${toDateInput(p.br.start)}" title="Projeté BR Début"></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="brEnd"   value="${toDateInput(p.br.end)}"   title="Projeté BR Fin"></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="atStart" value="${toDateInput(p.at.start)}" title="Projeté AT Début"></td>
        <td><input type="date" class="date-input plan-input" data-id="${p.id}" data-field="atEnd"   value="${toDateInput(p.at.end)}"   title="Projeté AT Fin"></td>
        <td><span class="readonly-date" style="color:#444;font-size:0.9em;">${p.reel.start ? toDateInput(p.reel.start) : '<em style="color:#bbb">—</em>'}</span></td>
        <td><span class="readonly-date" style="color:#444;font-size:0.9em;">${p.reel.end   ? toDateInput(p.reel.end)   : '<em style="color:#bbb">—</em>'}</span></td>
        <td><span class="badge ${statusClass}">${statut}</span></td>
      </tr>`;
    });
    html += `</tbody>`;
    tblPlan.innerHTML = html;

    tblPlan.querySelectorAll('.plan-input').forEach(input => {
      input.addEventListener('change', () => {
        collectManualDates();
        renderGanttSIA5(planData.map(p => {
          const manual = manualPlanningDates[p.id] || {};
          return {
            ...p,
            br: { start: manual.brStart || p.br.start, end: manual.brEnd || p.br.end },
            at: { start: manual.atStart || p.at.start, end: manual.atEnd || p.at.end }
          };
        }));
      });
    });
  }

  renderGanttSIA5(planData);
}

function collectManualDates() {
  document.querySelectorAll('.plan-input').forEach(input => {
    const id    = input.dataset.id;
    const field = input.dataset.field;
    if (!manualPlanningDates[id]) manualPlanningDates[id] = {};
    manualPlanningDates[id][field] = input.value ? new Date(input.value) : null;
  });
}

function renderGanttSIA5(planData) {
  destroyChart('ganttSIA5');
  const ctx = document.getElementById('ganttSIA5');
  if (!ctx) return;

  const valid = planData.filter(p => p.br.start || p.at.start || p.reel.start);
  if (!valid.length) return;

  const allDates = valid.flatMap(p => [p.br.start,p.br.end,p.at.start,p.at.end,p.reel.start,p.reel.end]).filter(Boolean);
  const axisStart = new Date(Math.min(...allDates));
  const axisEnd   = new Date(Math.max(...allDates));
  const axis = buildAxis(axisStart, axisEnd);

  const toIdx = d => {
    if (!d) return null;
    const mm = yyyymm(firstOfMonth(d));
    const i = axis.indexOf(mm);
    return i >= 0 ? i : null;
  };

  const buildBar = (startD, endD) => {
    const si = toIdx(startD), ei = toIdx(endD);
    if (si === null) return [null, null];
    const len = ei !== null ? Math.max(1, ei - si + 1) : 1;
    return [si, len];
  };

  const labels = valid.map(p => p.name);
  const [brBaseArr, brLenArr]     = [valid.map(p=>buildBar(p.br.start,p.br.end)[0]),     valid.map(p=>buildBar(p.br.start,p.br.end)[1])];
  const [atBaseArr, atLenArr]     = [valid.map(p=>buildBar(p.at.start,p.at.end)[0]),     valid.map(p=>buildBar(p.at.start,p.at.end)[1])];
  const [reelBaseArr, reelLenArr] = [valid.map(p=>buildBar(p.reel.start,p.reel.end)[0]), valid.map(p=>buildBar(p.reel.start,p.reel.end)[1])];

  charts['ganttSIA5'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'BR off',          data:brBaseArr,   backgroundColor:'transparent', borderWidth:0, stack:'br',   barThickness:18 },
        { label:'Projeté BR',      data:brLenArr,    backgroundColor:GE_ORANGE+'cc', borderRadius:3, stack:'br',   barThickness:18 },
        { label:'AT off',          data:atBaseArr,   backgroundColor:'transparent', borderWidth:0, stack:'at',   barThickness:18 },
        { label:'Projeté AT',      data:atLenArr,    backgroundColor:GE_BLUE2+'cc',  borderRadius:3, stack:'at',   barThickness:18 },
        { label:'Réel off',        data:reelBaseArr, backgroundColor:'transparent', borderWidth:0, stack:'reel', barThickness:18 },
        { label:'Réel (constaté)', data:reelLenArr,  backgroundColor:GE_RED+'cc',   borderRadius:3, stack:'reel', barThickness:18 }
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: {
        legend: {
          labels: {
            filter: item => !item.text.includes('off'),
            usePointStyle:true, font:{size:12}
          }
        },
        todayLine: { label: yyyymm(new Date()) }
      },
      scales: {
        x: { stacked:true, ticks:{ callback: v => axis[v] ? fmtMonth(axis[v]) : '', font:{size:10} }, grid:{color:'#e8eef5'} },
        y: { stacked:true, ticks:{ font:{size:11} } }
      }
    }
  });
  ctx.parentElement.style.height = `${valid.length * 70 + 80}px`;
}

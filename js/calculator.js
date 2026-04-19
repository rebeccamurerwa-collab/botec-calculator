// BOTEC Calculator JS
// Handles: beneficiary rows, cost rows, calculations, save/load, exports

const HEADS = ['Internal Consulting','Travel Costs','Premix Costs','Equipment Costs',
  'M&E Costs','Logistics Costs','Packaging Costs','Event / Admin Costs','Other'];

const PRESETS = {
  c1:        { h:'Internal Consulting', d:'Partnerships Manager',           u:'Person-month', cpu:105000, y1:9.6,       y2:8.4,       y3:7.2 },
  c2:        { h:'Internal Consulting', d:'Senior Partnerships Officer',    u:'Person-month', cpu:85000,  y1:10.8,      y2:9.6,       y3:8.4 },
  travel:    { h:'Travel Costs',        d:'Travel – Partnerships Manager',  u:'Trip',         cpu:25000,  y1:12,        y2:6,         y3:4   },
  premix:    { h:'Premix Costs',        d:'NaFeEDTA premix',                u:'KG',           cpu:400,    y1:3690,      y2:3690,      y3:3690 },
  equip:     { h:'Equipment Costs',     d:'Microdoser',                     u:'Device',       cpu:200000, y1:5,         y2:0,         y3:0   },
  mae:       { h:'M&E Costs',           d:'Iron spot test kit',             u:'Kit',          cpu:1750,   y1:10,        y2:10,        y3:10  },
  transport: { h:'Logistics Costs',     d:'Transportation cost',            u:'Per KG atta',  cpu:1,      y1:18450000,  y2:18450000,  y3:18450000 },
  grinding:  { h:'Logistics Costs',     d:'Grinding cost',                  u:'Per KG wheat', cpu:3,      y1:18450000,  y2:18450000,  y3:18450000 },
  packaging: { h:'Packaging Costs',     d:'Packaging cost',           u:'Per KG wheat flour', cpu:0.5,    y1:18450000,  y2:18450000,  y3:18450000 }
};

const BEN_LABELS = [
  '<Major Unit> Number',
  '<Minor Unit> Number',
  'Number of Beneficiaries per <Minor Unit>',
  '<Sub-unit> Number',
  'Number of Beneficiaries per <Sub-unit>',
  'Custom'
];

let logSet = new Set(['Logistics Costs']);
let chart = null;
let lastCalc = {};
let docId = null;
let currentUser = null;
let benID = 0, cID = 0;

// ---- AUTH GUARD ----
async function init() {
  const { data: { session } } = await sb.auth.getSession();
  if (!session) { window.location.href = 'login.html'; return; }
  currentUser = session.user;

  renderLogFlags();

  // Check if opening existing doc
  const params = new URLSearchParams(window.location.search);
  docId = params.get('id');

  if (docId) {
    await loadDocument(docId);
  }
  // else blank calculator — no default data
}

// ---- SAVE / LOAD ----
function setSaveStatus(msg, colour) {
  const el = document.getElementById('save-status');
  el.textContent = msg;
  el.style.color = colour || 'var(--color-text-secondary)';
}

async function saveDocument() {
  setSaveStatus('Saving…');
  calcAll();

  const name = document.getElementById('doc-title').value.trim() || 'Untitled BOTEC';
  const programme = document.getElementById('programme').value.trim();

  // Serialise full state
  const data = serialiseState();

  if (docId) {
    const { error } = await sb.from('botec_documents')
      .update({ name, programme, data })
      .eq('id', docId);
    if (error) { setSaveStatus('Save failed', '#c0392b'); alert(error.message); return; }
  } else {
    const { data: inserted, error } = await sb.from('botec_documents')
      .insert({ user_id: currentUser.id, name, programme, data })
      .select('id')
      .single();
    if (error) { setSaveStatus('Save failed', '#c0392b'); alert(error.message); return; }
    docId = inserted.id;
    window.history.replaceState({}, '', `calculator.html?id=${docId}`);
  }

  document.title = `${name} — BOTEC`;
  setSaveStatus('Saved', 'var(--color-text-success, #1D9E75)');
  setTimeout(() => setSaveStatus(''), 3000);
}

async function loadDocument(id) {
  const { data, error } = await sb.from('botec_documents').select('*').eq('id', id).single();
  if (error) { alert('Could not load document: ' + error.message); return; }
  document.getElementById('doc-title').value = data.name;
  document.title = `${data.name} — BOTEC`;
  deserialiseState(data.data);
  setSaveStatus('');
}

function serialiseState() {
  calcAll();
  return {
    // setup
    projName: document.getElementById('projName').value,
    programme: document.getElementById('programme').value,
    prepBy: document.getElementById('prepBy').value,
    prepDate: document.getElementById('prepDate').value,
    reviewBy: document.getElementById('reviewBy').value,
    reviewDate: document.getElementById('reviewDate').value,
    currency: document.getElementById('currency').value,
    numYears: document.getElementById('numYears').value,
    bufferPct: document.getElementById('bufferPct').value,
    mgrMult: document.getElementById('mgrMult').value,
    purposeNote: document.getElementById('purposeNote').value,
    logSet: [...logSet],
    // rows
    benRows: getBenRows(),
    costRows: getCostData(),
    // calc results (for dashboard preview)
    cpbAvg: lastCalc.cpbAvg,
    totalBen: lastCalc.totalBen,
    totalAll: lastCalc.totalAll
  };
}

function deserialiseState(s) {
  if (!s) return;
  // setup
  const set = (id, v) => { const el = document.getElementById(id); if (el && v != null) el.value = v; };
  set('projName', s.projName); set('programme', s.programme);
  set('prepBy', s.prepBy); set('prepDate', s.prepDate);
  set('reviewBy', s.reviewBy); set('reviewDate', s.reviewDate);
  set('currency', s.currency); set('numYears', s.numYears);
  set('bufferPct', s.bufferPct); set('mgrMult', s.mgrMult);
  set('purposeNote', s.purposeNote);

  if (s.logSet) { logSet = new Set(s.logSet); renderLogFlags(); }

  // beneficiary rows
  document.getElementById('benBody').innerHTML = '';
  benID = 0;
  (s.benRows || []).forEach(r => addBen(r));

  // cost rows
  document.getElementById('costBody').innerHTML = '';
  cID = 0;
  (s.costRows || []).forEach(r => addCost(r));

  calcAll();
}

// ---- LOG FLAGS ----
function renderLogFlags() {
  const el = document.getElementById('logFlags');
  el.innerHTML = HEADS.map(h => `
    <label style="display:flex;align-items:center;gap:5px;font-size:13px;cursor:pointer;padding:4px 10px;
      border:0.5px solid var(--color-border-tertiary);border-radius:20px;
      background:${logSet.has(h) ? 'var(--color-background-info,#e6f1fb)' : 'var(--color-background-secondary)'}">
      <input type="checkbox" ${logSet.has(h) ? 'checked' : ''} onchange="toggleLog('${h}',this.checked)"
        style="width:auto;margin:0"> ${h}
    </label>`).join('');
}
function toggleLog(h, v) { v ? logSet.add(h) : logSet.delete(h); renderLogFlags(); calcAll(); }

// ---- TAB SWITCHING ----
function go(id) {
  document.querySelectorAll('.tab').forEach((t, i) => {
    t.classList.toggle('on', ['setup','ben','costs','results'][i] === id);
  });
  document.querySelectorAll('.pane').forEach(p => p.classList.remove('on'));
  document.getElementById('p-' + id).classList.add('on');
  if (id === 'results') { calcAll(); renderResults(); }
}

// ---- BENEFICIARY ROWS ----
function addBen(d = {}) {
  const id = ++benID;
  const tr = document.createElement('tr');
  tr.id = 'bn' + id;
  const labelSel = BEN_LABELS.map(l =>
    `<option value="${l}" ${(d.label || BEN_LABELS[0]) === l ? 'selected' : ''}>${l}</option>`
  ).join('');
  tr.innerHTML = `
    <td><select style="font-size:12px" onchange="calcBens()">${labelSel}</select></td>
    <td><input type="text" placeholder="e.g. Districts" value="${d.name || ''}" style="font-style:italic"></td>
    <td><input type="text" placeholder="Source / assumption" value="${d.notes || ''}"></td>
    <td><input type="number" value="${d.y1 != null && !isNaN(d.y1) ? d.y1 : ''}" style="text-align:right" oninput="calcBens()"></td>
    <td><input type="number" value="${d.y2 != null && !isNaN(d.y2) ? d.y2 : ''}" style="text-align:right" oninput="calcBens()"></td>
    <td><input type="number" value="${d.y3 != null && !isNaN(d.y3) ? d.y3 : ''}" style="text-align:right" oninput="calcBens()"></td>
    <td><button class="del" onclick="document.getElementById('bn${id}').remove();calcBens()">×</button></td>`;
  document.getElementById('benBody').appendChild(tr);
  calcBens();
}

function getBenRows() {
  return Array.from(document.querySelectorAll('#benBody tr')).map(tr => {
    const ins = tr.querySelectorAll('input,select');
    return {
      label: ins[0].value, name: ins[1].value, notes: ins[2].value,
      y1: parseFloat(ins[3].value), y2: parseFloat(ins[4].value), y3: parseFloat(ins[5].value)
    };
  });
}

function calcBens() {
  const rows = getBenRows();
  let p1 = null, p2 = null, p3 = null;
  rows.forEach(r => {
    if (!isNaN(r.y1)) p1 = p1 === null ? r.y1 : p1 * r.y1;
    if (!isNaN(r.y2)) p2 = p2 === null ? r.y2 : p2 * r.y2;
    if (!isNaN(r.y3)) p3 = p3 === null ? r.y3 : p3 * r.y3;
  });
  const fmt = n => n !== null ? `<strong>${Math.round(n).toLocaleString()}</strong>` : '—';
  document.getElementById('tot1').innerHTML = fmt(p1);
  document.getElementById('tot2').innerHTML = fmt(p2);
  document.getElementById('tot3').innerHTML = fmt(p3);

  const names = rows.map(r => r.name || r.label.replace(/<|>/g, ''));
  const vals1 = rows.map(r => isNaN(r.y1) ? '?' : r.y1);
  const el = document.getElementById('benExplain');
  if (el && rows.length > 0) {
    el.innerHTML = `<strong>Formula (Year 1):</strong> ${names.map((n, i) => `${n || 'unit'} (${vals1[i]})`).join(' × ')} = <strong>${p1 !== null ? Math.round(p1).toLocaleString() : '?'}</strong> beneficiaries`;
  }
  calcAll();
}

function getBenTotals() {
  const rows = getBenRows();
  let p1 = null, p2 = null, p3 = null;
  rows.forEach(r => {
    if (!isNaN(r.y1)) p1 = p1 === null ? r.y1 : p1 * r.y1;
    if (!isNaN(r.y2)) p2 = p2 === null ? r.y2 : p2 * r.y2;
    if (!isNaN(r.y3)) p3 = p3 === null ? r.y3 : p3 * r.y3;
  });
  return { b1: p1 || 0, b2: p2 || 0, b3: p3 || 0 };
}

// ---- COST ROWS ----
function addCost(d = {}) {
  const id = ++cID;
  const tr = document.createElement('tr');
  tr.id = 'cr' + id;
  const sel = HEADS.map(h => `<option value="${h}" ${h === (d.h || d.head || HEADS[0]) ? 'selected' : ''}>${h}</option>`).join('');
  const cpu = d.cpu || 0, u1 = d.y1 || d.u1 || 0, u2 = d.y2 || d.u2 || 0, u3 = d.y3 || d.u3 || 0;
  tr.innerHTML = `
    <td><select onchange="calcAll()">${sel}</select></td>
    <td><input type="text" value="${d.d || d.desc || ''}"></td>
    <td><input type="text" value="${d.u || d.unit || ''}"></td>
    <td><input type="number" value="${cpu}" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" value="${u1}" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" value="${u2}" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" value="${u3}" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td id="cy1_${id}" class="cost-computed">—</td>
    <td id="cy2_${id}" class="cost-computed">—</td>
    <td id="cy3_${id}" class="cost-computed">—</td>
    <td><button class="del" onclick="document.getElementById('cr${id}').remove();calcAll()">×</button></td>`;
  document.getElementById('costBody').appendChild(tr);
  updateCostRow(tr.querySelector('input'));
  calcAll();
}
function preset(k) { addCost(PRESETS[k]); }

function updateCostRow(inp) {
  const tr = inp.closest('tr');
  const ins = tr.querySelectorAll('input[type=number]');
  const cpu = parseFloat(ins[0].value) || 0;
  const u1 = parseFloat(ins[1].value) || 0;
  const u2 = parseFloat(ins[2].value) || 0;
  const u3 = parseFloat(ins[3].value) || 0;
  const id = tr.id.replace('cr', '');
  const sym = getSym();
  const fmtC = n => sym + Math.round(n).toLocaleString();
  ['cy1_', 'cy2_', 'cy3_'].forEach((p, i) => {
    const el = document.getElementById(p + id);
    if (el) el.textContent = fmtC(cpu * [u1, u2, u3][i]);
  });
}

function getCostData() {
  return Array.from(document.querySelectorAll('#costBody tr')).map(tr => {
    const ins = tr.querySelectorAll('input,select');
    const cpu = parseFloat(ins[3].value) || 0;
    const u1 = parseFloat(ins[4].value) || 0;
    const u2 = parseFloat(ins[5].value) || 0;
    const u3 = parseFloat(ins[6].value) || 0;
    return { head: ins[0].value, desc: ins[1].value, unit: ins[2].value, cpu, u1, u2, u3, cy1: cpu*u1, cy2: cpu*u2, cy3: cpu*u3 };
  });
}

// ---- HELPERS ----
const getSym = () => ({ INR:'₹', USD:'$', GBP:'£', EUR:'€' })[document.getElementById('currency').value] || '₹';
const nv = id => parseFloat(document.getElementById(id).value) || 0;
const fC  = n => getSym() + Math.round(n).toLocaleString();
const fCd = (n, d=2) => getSym() + n.toFixed(d);

// ---- CORE CALCULATIONS ----
function calcAll() {
  const rows = getCostData();
  const { b1, b2, b3 } = getBenTotals();
  const buf = nv('bufferPct') / 100;
  const mgr = nv('mgrMult') / 100;
  const NY  = parseInt(document.getElementById('numYears').value) || 3;

  const byHead = {};
  HEADS.forEach(h => { byHead[h] = { y1:0, y2:0, y3:0 }; });
  rows.forEach(r => {
    const k = byHead[r.head] !== undefined ? r.head : 'Other';
    byHead[k].y1 += r.cy1; byHead[k].y2 += r.cy2; byHead[k].y3 += r.cy3;
  });

  const cons   = byHead['Internal Consulting'];
  const mgrVal = { y1: cons.y1*mgr, y2: cons.y2*mgr, y3: cons.y3*mgr };

  const subY = { y1:0, y2:0, y3:0 };
  HEADS.forEach(h => { subY.y1 += byHead[h].y1; subY.y2 += byHead[h].y2; subY.y3 += byHead[h].y3; });
  subY.y1 += mgrVal.y1; subY.y2 += mgrVal.y2; subY.y3 += mgrVal.y3;

  const bufVal = { y1: subY.y1*buf, y2: subY.y2*buf, y3: subY.y3*buf };
  const totY   = { y1: subY.y1+bufVal.y1, y2: subY.y2+bufVal.y2, y3: subY.y3+bufVal.y3 };

  const yrTotals = [totY.y1, totY.y2, totY.y3].slice(0, NY);
  const bens     = [b1, b2, b3].slice(0, NY);
  const totalAll = yrTotals.reduce((s, v) => s + v, 0);
  const totalBen = bens.reduce((s, v) => s + v, 0);
  const avgCost  = totalAll / NY;

  const cpbY   = [b1>0?totY.y1/b1:0, b2>0?totY.y2/b2:0, b3>0?totY.y3/b3:0];
  const cpbAvg = totalBen > 0 ? totalAll / totalBen : 0;

  const logY = { y1:0, y2:0, y3:0 };
  logSet.forEach(h => { if (byHead[h]) { logY.y1 += byHead[h].y1; logY.y2 += byHead[h].y2; logY.y3 += byHead[h].y3; } });

  const nonLogBase = { y1:0, y2:0, y3:0 };
  HEADS.forEach(h => {
    if (!logSet.has(h)) { nonLogBase.y1 += byHead[h].y1; nonLogBase.y2 += byHead[h].y2; nonLogBase.y3 += byHead[h].y3; }
  });
  nonLogBase.y1 += mgrVal.y1; nonLogBase.y2 += mgrVal.y2; nonLogBase.y3 += mgrVal.y3;

  const cpbExclY = [
    b1>0 ? nonLogBase.y1*(1+buf)/b1 : 0,
    b2>0 ? nonLogBase.y2*(1+buf)/b2 : 0,
    b3>0 ? nonLogBase.y3*(1+buf)/b3 : 0
  ];
  const cpbExclAvg = totalBen > 0 ? (nonLogBase.y1+nonLogBase.y2+nonLogBase.y3)*(1+buf)/totalBen : 0;

  lastCalc = { rows, byHead, mgrVal, bufVal, subY, totY, totalAll, totalBen, avgCost,
               bens, yrTotals, cpbY, cpbAvg, logY, cpbExclY, cpbExclAvg, buf, mgr, NY, b1, b2, b3 };
}

// ---- RENDER RESULTS ----
function renderResults() {
  calcAll();
  const d = lastCalc;
  if (!d.yrTotals) return;
  const NY = d.NY || 3;

  document.getElementById('summCards').innerHTML = `
    <div class="card"><div class="lbl">CPB (avg)</div><div class="val">${fCd(d.cpbAvg)}</div><div class="sub">total cost ÷ total beneficiaries</div></div>
    <div class="card"><div class="lbl">CPB excl. logistics</div><div class="val">${fCd(d.cpbExclAvg)}</div></div>
    <div class="card"><div class="lbl">Total ${NY}-year cost</div><div class="val">${fC(d.totalAll)}</div></div>
    <div class="card"><div class="lbl">Avg yearly cost</div><div class="val">${fC(d.avgCost)}</div><div class="sub">total ÷ ${NY}</div></div>
    <div class="card"><div class="lbl">Total beneficiaries</div><div class="val">${d.totalBen.toLocaleString()}</div></div>
  `;

  const yrs = Array.from({ length: NY }, (_, i) => `Year ${i + 1}`);
  document.getElementById('summHead').innerHTML =
    `<th>Cost head</th>${yrs.map(y => `<th style="text-align:right">${y}</th>`).join('')}<th style="text-align:right">Avg/year</th><th style="text-align:right">${NY}-year total</th>`;

  const bd = document.getElementById('summBody');
  bd.innerHTML = '';
  const ah = HEADS.filter(h => d.byHead[h] && d.byHead[h].y1 + d.byHead[h].y2 + d.byHead[h].y3 > 0);
  ah.forEach(h => {
    const v = d.byHead[h], vals = [v.y1, v.y2, v.y3].slice(0, NY), t = vals.reduce((s, x) => s + x, 0);
    bd.innerHTML += `<tr><td>${h}</td>${vals.map(x => `<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(t/NY)}</td><td style="text-align:right">${fC(t)}</td></tr>`;
  });

  const mv = [d.mgrVal.y1, d.mgrVal.y2, d.mgrVal.y3].slice(0, NY);
  const mt = mv.reduce((s, x) => s + x, 0);
  if (mt > 0) bd.innerHTML += `<tr class="derived-row"><td>${(d.mgr*100).toFixed(0)}% managerial multiplier</td>${mv.map(x => `<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(mt/NY)}</td><td style="text-align:right">${fC(mt)}</td></tr>`;

  const bv = [d.bufVal.y1, d.bufVal.y2, d.bufVal.y3].slice(0, NY);
  const bt = bv.reduce((s, x) => s + x, 0);
  if (bt > 0) bd.innerHTML += `<tr class="derived-row"><td>${(d.buf*100).toFixed(0)}% buffer</td>${bv.map(x => `<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(bt/NY)}</td><td style="text-align:right">${fC(bt)}</td></tr>`;

  const tv = d.yrTotals;
  bd.innerHTML += `<tr class="grand-row"><td>Total costs</td>${tv.map(x => `<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(d.avgCost)}</td><td style="text-align:right">${fC(d.totalAll)}</td></tr>`;
  bd.innerHTML += `<tr class="derived-row"><td>Number of beneficiaries</td>${d.bens.map(x => `<td style="text-align:right">${Math.round(x).toLocaleString()}</td>`).join('')}<td style="text-align:right">${Math.round(d.totalBen/NY).toLocaleString()}</td><td style="text-align:right">${Math.round(d.totalBen).toLocaleString()}</td></tr>`;
  bd.innerHTML += `<tr class="cpb-row"><td>Cost per beneficiary</td>${d.cpbY.slice(0,NY).map(x => `<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbAvg)}</td><td style="text-align:right">${fCd(d.cpbAvg)}</td></tr>`;
  if (logSet.size > 0) {
    bd.innerHTML += `<tr class="cpb-row" style="font-style:italic"><td>CPB excl. ${[...logSet].join(' + ')}</td>${d.cpbExclY.slice(0,NY).map(x => `<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbExclAvg)}</td><td style="text-align:right">${fCd(d.cpbExclAvg)}</td></tr>`;
  }

  // Chart
  if (chart) chart.destroy();
  const colors = ['#378ADD','#1D9E75','#BA7517','#D85A30','#D4537E','#7F77DD','#639922','#B4B2A9'];
  const datasets = ah.map((h, i) => ({
    label: h,
    data: [d.byHead[h].y1, d.byHead[h].y2, d.byHead[h].y3].slice(0, NY),
    backgroundColor: colors[i % colors.length],
    stack: 's'
  }));
  if (mt > 0) datasets.push({ label: 'Mgr multiplier', data: mv, backgroundColor: '#CCC', stack: 's' });
  if (bt > 0) datasets.push({ label: 'Buffer', data: bv, backgroundColor: '#E0E0E0', stack: 's' });

  chart = new Chart(document.getElementById('chartC').getContext('2d'), {
    type: 'bar',
    data: { labels: yrs, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fC(ctx.parsed.y)}` } } },
      scales: {
        x: { stacked: true },
        y: { stacked: true, ticks: { callback: v => getSym() + (v >= 1e7 ? (v/1e7).toFixed(1)+'Cr' : v >= 1e5 ? (v/1e5).toFixed(0)+'L' : v >= 1000 ? (v/1000).toFixed(0)+'K' : v) } }
      }
    }
  });

  // Formula box
  const b = nv('bufferPct'), m = nv('mgrMult');
  document.getElementById('fBox').innerHTML =
    `<strong>Beneficiaries [Yr]</strong> = Row₁[Yr] × Row₂[Yr] × … (product of all unit rows)<br>` +
    `<strong>Row cost [Yr]</strong> = Cost/unit × Units [Yr]<br>` +
    `<strong>Cost head total [Yr]</strong> = Σ all rows with that head<br>` +
    `<strong>Managerial multiplier [Yr]</strong> = Internal Consulting [Yr] × ${m}%<br>` +
    `<strong>Sub-total [Yr]</strong> = Σ(all cost heads [Yr]) + managerial multiplier [Yr]<br>` +
    `<strong>Buffer [Yr]</strong> = Sub-total [Yr] × ${b}%<br>` +
    `<strong>Total cost [Yr]</strong> = Sub-total [Yr] + Buffer [Yr]<br>` +
    `<strong>CPB [Yr]</strong> = Total cost [Yr] ÷ Beneficiaries [Yr]<br>` +
    `<strong>CPB (avg)</strong> = Σ(Total cost) ÷ Σ(Beneficiaries) — not average of averages<br>` +
    `<strong>CPB excl. logistics [Yr]</strong> = (Non-logistics sub-total + mgr multiplier) × (1 + ${b}%) ÷ Beneficiaries [Yr]`;
}

// ---- EXPORTS ----
function dlCSV() {
  calcAll(); const d = lastCalc;
  const pn = document.getElementById('projName').value || 'BOTEC';
  let c = `BOTEC Cost Per Beneficiary Estimate\nProject,${pn}\n\n`;
  c += `BENEFICIARY CALCULATION\nTemplate Label,Unit Name,Notes,Year 1,Year 2,Year 3\n`;
  getBenRows().forEach(r => { c += `"${r.label}","${r.name}","${r.notes}",${isNaN(r.y1)?'':r.y1},${isNaN(r.y2)?'':r.y2},${isNaN(r.y3)?'':r.y3}\n`; });
  c += `Total Beneficiaries,,,${Math.round(d.b1)},${Math.round(d.b2)},${Math.round(d.b3)}\n\n`;
  c += `COST ITEMS\nCost Head,Description,Unit,Cost/unit,Units Y1,Cost Y1,Units Y2,Cost Y2,Units Y3,Cost Y3\n`;
  d.rows.forEach(r => { c += `${r.head},"${r.desc}","${r.unit}",${r.cpu},${r.u1},${r.cy1},${r.u2},${r.cy2},${r.u3},${r.cy3}\n`; });
  c += `\nSUMMARY\nTotal,${d.totY.y1},${d.totY.y2},${d.totY.y3},${d.avgCost},${d.totalAll}\n`;
  c += `CPB,${d.cpbY.join(',')},${d.cpbAvg}\nCPB excl logistics,${d.cpbExclY.join(',')},${d.cpbExclAvg}\n`;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([c], { type: 'text/csv' }));
  a.download = `BOTEC_${pn.replace(/\s+/g,'_')}.csv`; a.click();
}

function dlJSON() {
  calcAll();
  const blob = new Blob([JSON.stringify(serialiseState(), null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `BOTEC_${(document.getElementById('projName').value||'export').replace(/\s+/g,'_')}.json`;
  a.click();
}

function dlXLSX() {
  calcAll(); const d = lastCalc;
  const WB = XLSX.utils.book_new();
  const pn = document.getElementById('projName').value || 'BOTEC';
  const cur = document.getElementById('currency').value;
  const NY = d.NY || 3;

  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet([
    ['READ-ME COST PER BENEFICIARY ESTIMATE'], [''],
    ['Purpose', document.getElementById('purposeNote').value], [''],
    ['Tabs','Details'],
    ['Summary','Lays out the Cost per Beneficiary details'],
    ['Beneficiary Calculation','Estimated number of beneficiaries'],
    ['Cost Calculation','Estimated cost calculation'],
    ['Unit Costs','Total units of resources used']
  ]), 'Read Me');

  const sumRows = [
    ['COST PER BENEFICIARY ESTIMATE'], [''],
    ['Project Name:', pn],
    ['Prepared By:', document.getElementById('prepBy').value],
    ['Preparation Date:', document.getElementById('prepDate').value], [''],
    ['Reviewed By:', document.getElementById('reviewBy').value], [''],
    ['','','Year 1','Year 2','Year 3','Average Cost','Total Cost'],
    ['Number of Beneficiaries','', d.b1, d.b2, d.b3, d.totalBen/NY, d.totalBen], [''],
    [`Costs (in ${cur}):`]
  ];
  HEADS.filter(h => d.byHead[h] && d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3>0).forEach(h => {
    const v = d.byHead[h], t = v.y1+v.y2+v.y3;
    sumRows.push(['', h, v.y1, v.y2, v.y3, t/NY, t]);
  });
  const mt = d.mgrVal.y1+d.mgrVal.y2+d.mgrVal.y3;
  if (mt > 0) sumRows.push(['', `${(d.mgr*100).toFixed(0)}% multiplier`, d.mgrVal.y1, d.mgrVal.y2, d.mgrVal.y3, mt/NY, mt]);
  const bt = d.bufVal.y1+d.bufVal.y2+d.bufVal.y3;
  sumRows.push(['', `${(d.buf*100).toFixed(0)}% Buffer`, d.bufVal.y1, d.bufVal.y2, d.bufVal.y3, bt/NY, bt], ['']);
  sumRows.push(['Total Costs','', d.totY.y1, d.totY.y2, d.totY.y3, d.avgCost, d.totalAll]);
  sumRows.push(['Cost per Beneficiary','', d.cpbY[0], d.cpbY[1]||'', d.cpbY[2]||'', d.cpbAvg, d.cpbAvg]);
  if (logSet.size > 0) sumRows.push([`CPB excl. ${[...logSet].join('+')}`, cur, d.cpbExclY[0], d.cpbExclY[1]||'', d.cpbExclY[2]||'', d.cpbExclAvg, d.cpbExclAvg]);
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(sumRows), 'Summary');

  const benData = [['BENEFICIARY CALCULATIONS'],[''],
    ['Notes:','Break down into hierarchical units. Total = product of all rows per year.'],[''],[''],
    ['','Notes','Links','YEAR 1','YEAR 2','YEAR 3'],['']
  ];
  getBenRows().forEach(r => { benData.push([r.name||r.label, r.notes,'', isNaN(r.y1)?'':r.y1, isNaN(r.y2)?'':r.y2, isNaN(r.y3)?'':r.y3]); });
  benData.push(['Number of Beneficiaries','','', Math.round(d.b1), Math.round(d.b2), Math.round(d.b3)]);
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(benData), 'Beneficiary Calculation');

  const cc = [['COST CALCULATION'],[''],
    ['COST HEAD','UNITS DESCRIPTION','DESCRIPTION',`COST PER UNIT (${cur})`,'# Units Y1','# Units Y2','# Units Y3','Cost Y1','Cost Y2','Cost Y3']
  ];
  d.rows.forEach(r => { cc.push([r.head, r.desc,'', r.cpu, r.u1, r.u2, r.u3, r.cy1, r.cy2, r.cy3]); });
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(cc), 'Cost Calculation');

  const uc = [['Cost Head','Unit Description',`Cost per Unit (${cur})`,'Unit label']];
  const seen = new Set();
  d.rows.forEach(r => { const k = r.head+'|'+r.desc; if (!seen.has(k)) { seen.add(k); uc.push([r.head, r.desc, r.cpu, r.unit]); } });
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(uc), 'Unit Costs');

  XLSX.writeFile(WB, `BOTEC_${pn.replace(/\s+/g,'_')}.xlsx`);
}

// Mark unsaved on any input change
document.addEventListener('input', () => setSaveStatus('Unsaved'));

// Ctrl/Cmd+S to save
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); saveDocument(); }
});

init();

// BOTEC Calculator JS

const HEADS = ['Internal Consulting','Travel Costs','Premix Costs','Equipment Costs',
  'M&E Costs','Logistics Costs','Packaging Costs','Event / Admin Costs','Other'];

// Unit costs keep their defaults; units per year are empty by default
const PRESETS = {
  c1: {
    h:'Internal Consulting', d:'Partnerships Manager', u:'Staff member', cpu:105000, y1:'', y2:'', y3:'',
    guide:'How many months does this person work on the project each year? E.g. if they spend 80% of their time over 12 months, enter 9.6.'
  },
  c2: {
    h:'Internal Consulting', d:'Senior Partnerships Officer', u:'Staff member', cpu:85000, y1:'', y2:'', y3:'',
    guide:'How many months does this person work on the project each year? E.g. 90% of 12 months = 10.8 months.'
  },
  travel: {
    h:'Travel Costs', d:'Travel – Partnerships Manager', u:'Trip', cpu:25000, y1:'', y2:'', y3:'',
    guide:'How many field trips are taken each year? Each trip is costed at the rate above. E.g. one trip per month = 12.'
  },
  premix: {
    h:'Premix Costs', d:'NaFeEDTA premix', u:'KG', cpu:400, y1:'', y2:'', y3:'',
    guide:'Total KGs of premix needed each year. Calculate as: beneficiaries × daily consumption (g) × serving days ÷ 1,000,000.'
  },
  equip: {
    h:'Equipment Costs', d:'Microdoser', u:'Device', cpu:200000, y1:'', y2:'', y3:'',
    guide:'Number of devices to purchase. Usually a one-time purchase in Year 1 only — enter 0 for Years 2 and 3.'
  },
  mae: {
    h:'M&E Costs', d:'Iron spot test kit', u:'Kit', cpu:1750, y1:'', y2:'', y3:'',
    guide:'Number of test kits used per year. Typically one kit per mill per testing round (e.g. 10 mills × 1 test = 10).'
  },
  transport: {
    h:'Logistics Costs', d:'Transportation cost', u:'KG atta', cpu:1, y1:'', y2:'', y3:'',
    guide:'Total KGs of atta transported per year. Convert from MT: monthly consumption (MT) × 1,000 × 12 months.'
  },
  grinding: {
    h:'Logistics Costs', d:'Grinding cost', u:'KG wheat', cpu:3, y1:'', y2:'', y3:'',
    guide:'Total KGs of wheat to be ground per year. Typically the same volume as your annual atta consumption.'
  },
  packaging: {
    h:'Packaging Costs', d:'Packaging cost', u:'KG wheat flour', cpu:0.5, y1:'', y2:'', y3:'',
    guide:'Total KGs of wheat flour packaged per year. Same as your annual atta consumption volume.'
  }
};

let logSet = new Set(['Logistics Costs']);
let chart = null;
let lastCalc = {};
let docId = null;
let currentUser = null;
let cID = 0;

// ---- AUTH GUARD ----
async function init() {
  const { data: { session } } = await sb.auth.getSession();
  if (!session) { window.location.href = 'login.html'; return; }
  currentUser = session.user;
  renderLogFlags();
  const params = new URLSearchParams(window.location.search);
  docId = params.get('id');
  if (docId) await loadDocument(docId);
}

// ---- SAVE / LOAD ----
function setSaveStatus(msg, colour) {
  const el = document.getElementById('save-status');
  el.textContent = msg;
  el.style.color = colour || 'var(--text3)';
}

async function saveDocument() {
  setSaveStatus('Saving…');
  calcAll();
  const name = document.getElementById('doc-title').value.trim() || 'Untitled BOTEC';
  const programme = document.getElementById('programme').value.trim();
  const data = serialiseState();
  if (docId) {
    const { error } = await sb.from('botec_documents').update({ name, programme, data }).eq('id', docId);
    if (error) { setSaveStatus('Save failed', '#c0392b'); alert(error.message); return; }
  } else {
    const { data: inserted, error } = await sb.from('botec_documents')
      .insert({ user_id: currentUser.id, name, programme, data }).select('id').single();
    if (error) { setSaveStatus('Save failed', '#c0392b'); alert(error.message); return; }
    docId = inserted.id;
    window.history.replaceState({}, '', `calculator.html?id=${docId}`);
  }
  document.title = `${name} — BOTEC`;
  setSaveStatus('Saved', '#059669');
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
    benY1: document.getElementById('benY1').value,
    benGrowth: document.getElementById('benGrowth').value,
    benNotes: document.getElementById('benNotes').value,
    costRows: getCostData(),
    cpbAvg: lastCalc.cpbAvg,
    totalBen: lastCalc.totalBen,
    totalAll: lastCalc.totalAll
  };
}

function deserialiseState(s) {
  if (!s) return;
  const set = (id, v) => { const el = document.getElementById(id); if (el && v != null) el.value = v; };
  set('projName', s.projName); set('programme', s.programme);
  set('prepBy', s.prepBy); set('prepDate', s.prepDate);
  set('reviewBy', s.reviewBy); set('reviewDate', s.reviewDate);
  set('currency', s.currency); set('numYears', s.numYears);
  set('bufferPct', s.bufferPct); set('mgrMult', s.mgrMult);
  set('purposeNote', s.purposeNote);
  if (s.logSet) { logSet = new Set(s.logSet); renderLogFlags(); }
  set('benY1', s.benY1);
  set('benGrowth', s.benGrowth || '0');
  set('benNotes', s.benNotes);
  calcBens();
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
      border:0.5px solid var(--border);border-radius:20px;
      background:${logSet.has(h) ? '#e0f7fa' : 'var(--surface2)'}">
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

// ---- BENEFICIARIES (simplified) ----
function calcBens() {
  const { b1, b2, b3 } = getBenTotals();
  const growth = parseFloat(document.getElementById('benGrowth').value) || 0;
  const NY = parseInt(document.getElementById('numYears').value) || 3;
  const fmt = n => n > 0 ? Math.round(n).toLocaleString() : '—';

  document.getElementById('ben-show-1').textContent = fmt(b1);
  document.getElementById('ben-show-2').textContent = NY >= 2 ? fmt(b2) : '—';
  document.getElementById('ben-show-3').textContent = NY >= 3 ? fmt(b3) : '—';

  const g2 = document.getElementById('ben-growth-2');
  const g3 = document.getElementById('ben-growth-3');
  if (growth !== 0 && b1 > 0) {
    g2.textContent = `${growth > 0 ? '+' : ''}${growth}% vs Year 1`;
    g3.textContent = `${growth > 0 ? '+' : ''}${(growth * 2).toFixed(1)}% vs Year 1`;
  } else {
    g2.textContent = ''; g3.textContent = '';
  }

  const total = b1 + (NY >= 2 ? b2 : 0) + (NY >= 3 ? b3 : 0);
  document.getElementById('ben-total').textContent = total > 0 ? Math.round(total).toLocaleString() : '—';
  calcAll();
}

function getBenTotals() {
  const y1 = parseFloat(document.getElementById('benY1').value) || 0;
  const growth = parseFloat(document.getElementById('benGrowth').value) || 0;
  const g = 1 + growth / 100;
  return { b1: y1, b2: y1 * g, b3: y1 * g * g };
}

// ---- COST ROWS ----
function addCost(d = {}) {
  const id = ++cID;
  const tr = document.createElement('tr');
  tr.id = 'cr' + id;
  const sel = HEADS.map(h => `<option value="${h}" ${h === (d.h || d.head || HEADS[0]) ? 'selected' : ''}>${h}</option>`).join('');
  const cpu = d.cpu != null ? d.cpu : '';
  const u1 = (d.u1 != null && d.u1 !== '') ? d.u1 : (d.y1 !== '' && d.y1 != null ? d.y1 : '');
  const u2 = (d.u2 != null && d.u2 !== '') ? d.u2 : (d.y2 !== '' && d.y2 != null ? d.y2 : '');
  const u3 = (d.u3 != null && d.u3 !== '') ? d.u3 : (d.y3 !== '' && d.y3 != null ? d.y3 : '');
  const guide = d.guide || '';
  tr.innerHTML = `
    <td><select onchange="calcAll()">${sel}</select></td>
    <td>
      <input type="text" value="${d.d || d.desc || ''}">
      ${guide ? `<div class="cost-guide">${guide}</div>` : ''}
    </td>
    <td><input type="text" value="${d.u || d.unit || ''}" placeholder="e.g. staff member, trip, KG"></td>
    <td><input type="number" value="${cpu}" placeholder="Monthly cost" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" value="${u1}" placeholder="0" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" value="${u2}" placeholder="0" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" value="${u3}" placeholder="0" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td id="cy1_${id}" class="cost-computed">—</td>
    <td id="cy2_${id}" class="cost-computed">—</td>
    <td id="cy3_${id}" class="cost-computed">—</td>
    <td><button class="del" onclick="document.getElementById('cr${id}').remove();calcAll()">×</button></td>`;
  document.getElementById('costBody').appendChild(tr);
  updateCostRow(tr.querySelector('input[type=number]'));
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
  const fmtC = n => getSym() + Math.round(n).toLocaleString();
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

  const abbr = n => {
    if (n >= 1e7) return getSym() + (n/1e7).toFixed(2) + ' Cr';
    if (n >= 1e5) return getSym() + (n/1e5).toFixed(2) + ' L';
    if (n >= 1000) return getSym() + (n/1000).toFixed(1) + 'K';
    return fCd(n);
  };

  document.getElementById('summCards').innerHTML = `
    <div class="card"><div class="lbl">CPB — avg</div><div class="val">${fCd(d.cpbAvg)}</div><div class="sub">total cost ÷ total beneficiaries</div></div>
    <div class="card"><div class="lbl">CPB excl. logistics</div><div class="val">${fCd(d.cpbExclAvg)}</div></div>
    <div class="card"><div class="lbl">Total ${NY}-yr cost</div><div class="val">${abbr(d.totalAll)}</div></div>
    <div class="card"><div class="lbl">Avg yearly cost</div><div class="val">${abbr(d.avgCost)}</div><div class="sub">total ÷ ${NY}</div></div>
    <div class="card"><div class="lbl">Total beneficiaries</div><div class="val">${d.totalBen >= 1e5 ? (d.totalBen/1000).toFixed(1)+'K' : Math.round(d.totalBen).toLocaleString()}</div></div>
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

  const mv = [d.mgrVal.y1, d.mgrVal.y2, d.mgrVal.y3].slice(0, NY), mt = mv.reduce((s, x) => s + x, 0);
  if (mt > 0) bd.innerHTML += `<tr class="derived-row"><td>${(d.mgr*100).toFixed(0)}% managerial multiplier</td>${mv.map(x => `<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(mt/NY)}</td><td style="text-align:right">${fC(mt)}</td></tr>`;

  const bv = [d.bufVal.y1, d.bufVal.y2, d.bufVal.y3].slice(0, NY), bt = bv.reduce((s, x) => s + x, 0);
  if (bt > 0) bd.innerHTML += `<tr class="derived-row"><td>${(d.buf*100).toFixed(0)}% buffer</td>${bv.map(x => `<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(bt/NY)}</td><td style="text-align:right">${fC(bt)}</td></tr>`;

  bd.innerHTML += `<tr class="grand-row"><td>Total costs</td>${d.yrTotals.map(x => `<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(d.avgCost)}</td><td style="text-align:right">${fC(d.totalAll)}</td></tr>`;
  bd.innerHTML += `<tr class="derived-row"><td>Number of beneficiaries</td>${d.bens.map(x => `<td style="text-align:right">${Math.round(x).toLocaleString()}</td>`).join('')}<td style="text-align:right">${Math.round(d.totalBen/NY).toLocaleString()}</td><td style="text-align:right">${Math.round(d.totalBen).toLocaleString()}</td></tr>`;
  bd.innerHTML += `<tr class="cpb-row"><td>Cost per beneficiary</td>${d.cpbY.slice(0,NY).map(x => `<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbAvg)}</td><td style="text-align:right">${fCd(d.cpbAvg)}</td></tr>`;
  if (logSet.size > 0) {
    bd.innerHTML += `<tr class="cpb-row" style="font-style:italic"><td>CPB excl. ${[...logSet].join(' + ')}</td>${d.cpbExclY.slice(0,NY).map(x => `<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbExclAvg)}</td><td style="text-align:right">${fCd(d.cpbExclAvg)}</td></tr>`;
  }

  if (chart) chart.destroy();
  const colors = ['#0097a7','#dc6059','#ff8dcb','#00bcd4','#f06292','#4dd0e1','#e57373','#80deea'];
  const datasets = ah.map((h, i) => ({
    label: h, data: [d.byHead[h].y1, d.byHead[h].y2, d.byHead[h].y3].slice(0, NY),
    backgroundColor: colors[i % colors.length], stack: 's'
  }));
  if (mt > 0) datasets.push({ label: 'Mgr multiplier', data: mv, backgroundColor: '#CCC', stack: 's' });
  if (bt > 0) datasets.push({ label: 'Buffer', data: bv, backgroundColor: '#E0E0E0', stack: 's' });

  chart = new Chart(document.getElementById('chartC').getContext('2d'), {
    type: 'bar', data: { labels: yrs, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fC(ctx.parsed.y)}` } } },
      scales: {
        x: { stacked: true },
        y: { stacked: true, ticks: { callback: v => getSym() + (v >= 1e7 ? (v/1e7).toFixed(1)+'Cr' : v >= 1e5 ? (v/1e5).toFixed(0)+'L' : v >= 1000 ? (v/1000).toFixed(0)+'K' : v) } }
      }
    }
  });
}

// ---- EXPORTS ----
function dlCSV() {
  calcAll(); const d = lastCalc;
  const pn = document.getElementById('projName').value || 'BOTEC';
  let c = `BOTEC Cost Per Beneficiary Estimate\nProject,${pn}\n\n`;
  c += `BENEFICIARIES\nYear 1,${Math.round(d.b1)}\nAnnual growth %,${document.getElementById('benGrowth').value||0}\nYear 2,${Math.round(d.b2)}\nYear 3,${Math.round(d.b3)}\nTotal,${Math.round(d.totalBen)}\n\n`;
  c += `COST ITEMS\nCost Head,Description,Unit,Monthly Cost/unit,Units Y1,Cost Y1,Units Y2,Cost Y2,Units Y3,Cost Y3\n`;
  d.rows.forEach(r => { c += `${r.head},"${r.desc}","${r.unit}",${r.cpu},${r.u1},${r.cy1},${r.u2},${r.cy2},${r.u3},${r.cy3}\n`; });
  c += `\nSUMMARY\nTotal,${d.totY.y1},${d.totY.y2},${d.totY.y3},${d.avgCost},${d.totalAll}\nCPB,${d.cpbY.join(',')},${d.cpbAvg}\n`;
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
    ['READ-ME COST PER BENEFICIARY ESTIMATE'],[''],
    ['Purpose', document.getElementById('purposeNote').value],[''],
    ['Tabs','Details'],['Summary','Cost per Beneficiary details'],
    ['Beneficiary Calculation','Estimated beneficiaries'],['Cost Calculation','Estimated costs'],['Unit Costs','Unit reference']
  ]), 'Read Me');
  const sumRows = [
    ['COST PER BENEFICIARY ESTIMATE'],[''],
    ['Project Name:', pn],['Prepared By:', document.getElementById('prepBy').value],
    ['Preparation Date:', document.getElementById('prepDate').value],[''],
    ['Reviewed By:', document.getElementById('reviewBy').value],[''],
    ['','','Year 1','Year 2','Year 3','Average Cost','Total Cost'],
    ['Number of Beneficiaries','', d.b1, d.b2, d.b3, d.totalBen/NY, d.totalBen],[''],['Costs (in '+cur+'):']
  ];
  HEADS.filter(h => d.byHead[h] && d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3>0).forEach(h => {
    const v=d.byHead[h],t=v.y1+v.y2+v.y3; sumRows.push(['',h,v.y1,v.y2,v.y3,t/NY,t]);
  });
  const mt=d.mgrVal.y1+d.mgrVal.y2+d.mgrVal.y3;
  if(mt>0) sumRows.push(['',`${(d.mgr*100).toFixed(0)}% multiplier`,d.mgrVal.y1,d.mgrVal.y2,d.mgrVal.y3,mt/NY,mt]);
  const bt=d.bufVal.y1+d.bufVal.y2+d.bufVal.y3;
  sumRows.push(['',`${(d.buf*100).toFixed(0)}% Buffer`,d.bufVal.y1,d.bufVal.y2,d.bufVal.y3,bt/NY,bt],['']);
  sumRows.push(['Total Costs','',d.totY.y1,d.totY.y2,d.totY.y3,d.avgCost,d.totalAll]);
  sumRows.push(['Cost per Beneficiary','',d.cpbY[0],d.cpbY[1]||'',d.cpbY[2]||'',d.cpbAvg,d.cpbAvg]);
  if(logSet.size>0) sumRows.push([`CPB excl. ${[...logSet].join('+')}`,cur,d.cpbExclY[0],d.cpbExclY[1]||'',d.cpbExclY[2]||'',d.cpbExclAvg,d.cpbExclAvg]);
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(sumRows), 'Summary');
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet([
    ['BENEFICIARY CALCULATIONS'],[''],
    ['Year 1 Beneficiaries', Math.round(d.b1)],
    ['Annual Growth %', document.getElementById('benGrowth').value||0],
    ['Year 2 Beneficiaries', Math.round(d.b2)],
    ['Year 3 Beneficiaries', Math.round(d.b3)],
    ['Notes', document.getElementById('benNotes').value]
  ]), 'Beneficiary Calculation');
  const cc=[['COST CALCULATION'],[''],['COST HEAD','DESCRIPTION','UNIT',`MONTHLY COST/UNIT (${cur})`,'UNITS Y1','UNITS Y2','UNITS Y3','Cost Y1','Cost Y2','Cost Y3']];
  d.rows.forEach(r=>{cc.push([r.head,r.desc,r.unit,r.cpu,r.u1,r.u2,r.u3,r.cy1,r.cy2,r.cy3]);});
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(cc), 'Cost Calculation');
  const uc=[['Cost Head','Unit Description',`Monthly Cost/Unit (${cur})`,'Unit type']];
  const seen=new Set();
  d.rows.forEach(r=>{const k=r.head+'|'+r.desc;if(!seen.has(k)){seen.add(k);uc.push([r.head,r.desc,r.cpu,r.unit]);}});
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(uc), 'Unit Costs');
  XLSX.writeFile(WB, `BOTEC_${pn.replace(/\s+/g,'_')}.xlsx`);
}

function dlPDF() {
  calcAll(); const d = lastCalc;
  if (!d.rows) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pn = document.getElementById('projName').value || 'BOTEC';
  const cur = document.getElementById('currency').value;
  const NY = d.NY || 3;
  const s = getSym();
  const fmtN = n => s + Math.round(n).toLocaleString();
  const fmtD = (n, dp=2) => s + n.toFixed(dp);

  doc.setFillColor(0,151,167); doc.rect(0,0,297,18,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(13); doc.setFont('helvetica','bold');
  doc.text('BOTEC Cost Per Beneficiary Estimate', 14, 12);
  doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.text(pn, 200, 12);
  doc.setTextColor(80,80,80); doc.setFontSize(8);
  doc.text(`Prepared by: ${document.getElementById('prepBy').value||'—'}   Date: ${document.getElementById('prepDate').value||'—'}   Programme: ${document.getElementById('programme').value||'—'}`, 14, 26);

  const cards=[
    {label:'Cost per beneficiary (avg)',val:fmtD(d.cpbAvg),color:[220,96,89]},
    {label:`Total ${NY}-year cost`,val:fmtN(d.totalAll),color:[0,151,167]},
    {label:'Avg yearly cost',val:fmtN(d.avgCost),color:[255,141,203]},
    {label:'Total beneficiaries',val:Math.round(d.totalBen).toLocaleString(),color:[0,151,167]}
  ];
  cards.forEach((c,i)=>{
    doc.setFillColor(...c.color); doc.roundedRect(14+i*71,30,66,18,2,2,'F');
    doc.setTextColor(255,255,255); doc.setFontSize(7); doc.setFont('helvetica','normal');
    doc.text(c.label, 18+i*71, 36); doc.setFontSize(11); doc.setFont('helvetica','bold');
    doc.text(c.val, 18+i*71, 44);
  });

  const yrs=Array.from({length:NY},(_,i)=>`Year ${i+1}`);
  const ah=HEADS.filter(h=>d.byHead[h]&&d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3>0);
  const tableRows=[];
  ah.forEach(h=>{const v=d.byHead[h],vals=[v.y1,v.y2,v.y3].slice(0,NY),t=vals.reduce((s,x)=>s+x,0);tableRows.push([h,...vals.map(x=>fmtN(x)),fmtN(t/NY),fmtN(t)]);});
  const mv=[d.mgrVal.y1,d.mgrVal.y2,d.mgrVal.y3].slice(0,NY),mt=mv.reduce((s,x)=>s+x,0);
  if(mt>0) tableRows.push([`${(d.mgr*100).toFixed(0)}% mgr multiplier`,...mv.map(x=>fmtN(x)),fmtN(mt/NY),fmtN(mt)]);
  const bv=[d.bufVal.y1,d.bufVal.y2,d.bufVal.y3].slice(0,NY),bt=bv.reduce((s,x)=>s+x,0);
  if(bt>0) tableRows.push([`${(d.buf*100).toFixed(0)}% buffer`,...bv.map(x=>fmtN(x)),fmtN(bt/NY),fmtN(bt)]);
  tableRows.push(['TOTAL COSTS',...d.yrTotals.map(x=>fmtN(x)),fmtN(d.avgCost),fmtN(d.totalAll)]);
  tableRows.push(['Beneficiaries',...d.bens.map(x=>Math.round(x).toLocaleString()),Math.round(d.totalBen/NY).toLocaleString(),Math.round(d.totalBen).toLocaleString()]);
  tableRows.push(['Cost per beneficiary',...d.cpbY.slice(0,NY).map(x=>fmtD(x)),fmtD(d.cpbAvg),fmtD(d.cpbAvg)]);
  if(logSet.size>0) tableRows.push([`CPB excl. ${[...logSet].join('+')}`,...d.cpbExclY.slice(0,NY).map(x=>fmtD(x)),fmtD(d.cpbExclAvg),fmtD(d.cpbExclAvg)]);

  doc.autoTable({
    startY:55, head:[['Cost head',...yrs,'Avg / year',`${NY}-year total`]], body:tableRows,
    styles:{fontSize:8,cellPadding:3}, headStyles:{fillColor:[0,151,167],textColor:255,fontStyle:'bold'},
    alternateRowStyles:{fillColor:[245,245,245]},
    didParseCell:data=>{
      if(['TOTAL COSTS','Cost per beneficiary'].includes(data.row.raw[0])) data.cell.styles.fontStyle='bold';
      if(data.row.raw[0].startsWith('Cost per')||data.row.raw[0].startsWith('CPB excl')) data.cell.styles.textColor=[220,96,89];
    }, margin:{left:14,right:14}
  });

  doc.addPage();
  doc.setFillColor(0,151,167); doc.rect(0,0,297,18,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(13); doc.setFont('helvetica','bold');
  doc.text('Cost Line Items', 14, 12);
  doc.autoTable({
    startY:25, head:[['Cost head','Description','Unit',`Monthly cost/unit (${cur})`,`Units Y1`,`Cost Y1`,`Units Y2`,`Cost Y2`,`Units Y3`,`Cost Y3`]],
    body:d.rows.map(r=>[r.head,r.desc,r.unit,fmtN(r.cpu),r.u1,fmtN(r.cy1),r.u2,fmtN(r.cy2),r.u3,fmtN(r.cy3)]),
    styles:{fontSize:7,cellPadding:2}, headStyles:{fillColor:[0,151,167],textColor:255,fontStyle:'bold'},
    alternateRowStyles:{fillColor:[245,245,245]}, margin:{left:14,right:14}
  });

  const pageCount=doc.internal.getNumberOfPages();
  for(let i=1;i<=pageCount;i++){
    doc.setPage(i); doc.setFontSize(7); doc.setTextColor(150,150,150); doc.setFont('helvetica','normal');
    doc.text(`${pn} — BOTEC Estimate`,14,205); doc.text(`Page ${i} of ${pageCount}`,270,205);
  }
  doc.save(`BOTEC_${pn.replace(/\s+/g,'_')}.pdf`);
}

document.addEventListener('input', () => setSaveStatus('Unsaved'));
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); saveDocument(); }
});

init();

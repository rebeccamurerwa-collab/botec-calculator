// BOTEC Calculator JS

const HEADS = ['Internal Consulting','Travel Costs','Premix Costs','Equipment Costs',
  'M&E Costs','Logistics Costs','Packaging Costs','Event / Admin Costs','Other'];

const PRESETS = {
  c1:        { h:'Internal Consulting', d:'Partnerships Manager',          u:'Staff member',    cpu:105000, y1:'',y2:'',y3:'', guide:'How many months does this person work on the project each year? E.g. 80% of 12 months = 9.6.' },
  c2:        { h:'Internal Consulting', d:'Senior Partnerships Officer',   u:'Staff member',    cpu:85000,  y1:'',y2:'',y3:'', guide:'How many months does this person work on the project each year? E.g. 90% of 12 months = 10.8.' },
  travel:    { h:'Travel Costs',        d:'Travel – Partnerships Manager', u:'Trip',            cpu:25000,  y1:'',y2:'',y3:'', guide:'How many field trips are taken each year? One trip per month = 12.' },
  premix:    { h:'Premix Costs',        d:'NaFeEDTA premix',               u:'KG',              cpu:400,    y1:'',y2:'',y3:'', guide:'Total KGs needed per year: beneficiaries × daily consumption (g) × serving days ÷ 1,000,000.' },
  equip:     { h:'Equipment Costs',     d:'Microdoser',                    u:'Device',          cpu:200000, y1:'',y2:'',y3:'', guide:'Number of devices to purchase. Usually Year 1 only — enter 0 for Years 2 and 3.' },
  mae:       { h:'M&E Costs',           d:'Iron spot test kit',            u:'Kit',             cpu:1750,   y1:'',y2:'',y3:'', guide:'Kits used per year. E.g. 10 mills × 1 test per year = 10.' },
  transport: { h:'Logistics Costs',     d:'Transportation cost',           u:'KG atta',         cpu:1,      y1:'',y2:'',y3:'', guide:'Total KGs of atta transported per year. E.g. monthly MT × 1,000 × 12.' },
  grinding:  { h:'Logistics Costs',     d:'Grinding cost',                 u:'KG wheat',        cpu:3,      y1:'',y2:'',y3:'', guide:'Total KGs of wheat ground per year — typically same as atta consumption volume.' },
  packaging: { h:'Packaging Costs',     d:'Packaging cost',                u:'KG wheat flour',  cpu:0.5,    y1:'',y2:'',y3:'', guide:'Total KGs packaged per year — same as annual atta consumption volume.' }
};

// ---- STATE ----
let logSet        = new Set(['Logistics Costs']);
let chart         = null;
let lastCalc      = {};
let docId         = null;
let currentUser   = null;
let cID           = 0;
let benMode       = 'growth'; // 'growth' or 'manual'
const isReviewMode = new URLSearchParams(window.location.search).has('review');

// ---- INIT ----
async function init() {
  const { data: { session } } = await sb.auth.getSession();

  if (isReviewMode) {
    // Reviewer doesn't need an account — load doc by ID only
    initReviewMode();
  } else {
    if (!session) { window.location.href = 'login.html'; return; }
    currentUser = session.user;
    renderLogFlags();
    const params = new URLSearchParams(window.location.search);
    docId = params.get('id');
    if (docId) {
      await loadDocument(docId);
      await loadComments();
    }
  }
}

// ============================================================
//  REVIEW MODE
// ============================================================
async function initReviewMode() {
  const params = new URLSearchParams(window.location.search);
  docId = params.get('id');
  if (!docId) { document.body.innerHTML = '<p style="padding:2rem">No document ID in URL.</p>'; return; }

  // Show review UI
  document.getElementById('creator-nav').classList.add('hidden');
  document.getElementById('reviewer-nav').classList.remove('hidden');
  document.getElementById('comments-fab').classList.remove('hidden');

  // Load the shared document
  const { data, error } = await sb.from('botec_documents').select('*').eq('id', docId).single();
  if (error || !data || !data.is_shared) {
    document.body.innerHTML = '<p style="padding:2rem;font-family:sans-serif">This document is not available for review, or the link has expired.</p>';
    return;
  }

  document.getElementById('reviewer-doc-title').textContent = data.name;
  document.title = `Reviewing: ${data.name}`;

  // Load data into form (all inputs will be disabled)
  deserialiseState(data.data);
  disableAllInputs();
  renderLogFlags();
}

function disableAllInputs() {
  document.querySelectorAll('input, select, textarea, button.addbtn, button.del').forEach(el => {
    el.disabled = true;
    el.style.opacity = '0.75';
  });
}

function getReviewerName() {
  let name = sessionStorage.getItem('reviewerName');
  if (!name) {
    name = prompt('Please enter your name for the review comments:');
    if (!name) name = 'Anonymous';
    sessionStorage.setItem('reviewerName', name);
  }
  return name;
}

// ---- SHARE FOR REVIEW ----
async function shareForReview() {
  if (!docId) {
    alert('Please save the document first before sharing.');
    return;
  }
  // Mark document as shared
  const { error } = await sb.from('botec_documents').update({ is_shared: true }).eq('id', docId);
  if (error) { alert('Error sharing: ' + error.message); return; }

  const reviewUrl = `${window.location.origin}${window.location.pathname}?id=${docId}&review`;
  await navigator.clipboard.writeText(reviewUrl);
  showToast('Review link copied to clipboard! Share it with your reviewer.');
}

async function stopSharing() {
  const { error } = await sb.from('botec_documents').update({ is_shared: false }).eq('id', docId);
  if (error) { alert('Error: ' + error.message); return; }
  showToast('Sharing disabled. The review link will no longer work.');
  document.getElementById('share-btn').textContent = '🔗 Share for review';
  document.getElementById('share-btn').onclick = shareForReview;
}

// ============================================================
//  COMMENTS
// ============================================================
async function loadComments() {
  if (!docId || isReviewMode) return;
  const { data, error } = await sb.from('botec_comments').select('*').eq('doc_id', docId).order('created_at');
  if (error) return;
  renderComments(data || []);
}

function renderComments(comments) {
  const panel = document.getElementById('comments-list');
  const badge = document.getElementById('comments-badge');
  const open  = comments.filter(c => !c.resolved).length;
  badge.textContent  = open > 0 ? open : '';
  badge.style.display = open > 0 ? 'flex' : 'none';

  if (comments.length === 0) {
    panel.innerHTML = '<p style="color:var(--text3);font-size:13px;text-align:center;padding:20px 0">No comments yet.</p>';
    return;
  }
  panel.innerHTML = comments.map(c => `
    <div class="comment-item ${c.resolved ? 'resolved' : ''}" id="cm-${c.id}">
      <div class="comment-meta">
        <span class="comment-author">${c.reviewer_name}</span>
        <span class="comment-section">${c.section}</span>
        <span class="comment-time">${new Date(c.created_at).toLocaleDateString('en-GB',{day:'numeric',month:'short'})}</span>
      </div>
      <p class="comment-text">${c.comment}</p>
      ${!c.resolved ? `<button class="resolve-btn" onclick="resolveComment('${c.id}')">✓ Mark resolved</button>` : '<span class="resolved-label">✓ Resolved</span>'}
    </div>`).join('');
}

async function resolveComment(id) {
  const { error } = await sb.from('botec_comments').update({ resolved: true }).eq('id', id);
  if (error) { alert('Error: ' + error.message); return; }
  await loadComments();
}

async function submitComment() {
  const name    = isReviewMode ? getReviewerName() : (currentUser?.user_metadata?.full_name || currentUser?.email || 'Creator');
  const section = document.getElementById('comment-section').value;
  const text    = document.getElementById('comment-text').value.trim();
  if (!text) { alert('Please enter a comment before submitting.'); return; }

  const { error } = await sb.from('botec_comments').insert({ doc_id: docId, reviewer_name: name, section, comment: text });
  if (error) { alert('Error saving comment: ' + error.message); return; }

  document.getElementById('comment-text').value = '';
  showToast('Comment submitted!');

  if (!isReviewMode) await loadComments();
  else {
    // Reviewer sees confirmation only
    document.getElementById('comment-form-area').innerHTML =
      '<p style="color:#059669;font-size:14px;padding:10px 0">✓ Comment submitted. You can add more comments or close this panel.</p>' +
      '<button class="addbtn" onclick="resetCommentForm()">Add another comment</button>';
  }
}

function resetCommentForm() {
  document.getElementById('comment-form-area').innerHTML = `
    <div class="field" style="margin-bottom:8px">
      <label>Section</label>
      <select id="comment-section">
        <option>General</option><option>Project Setup</option>
        <option>Beneficiaries</option><option>Cost Items</option><option>Results</option>
      </select>
    </div>
    <div class="field" style="margin-bottom:8px">
      <label>Comment</label>
      <textarea id="comment-text" rows="3" placeholder="Enter your comment here…"></textarea>
    </div>
    <button class="btn-primary" onclick="submitComment()">Submit comment</button>`;
}

function toggleCommentsPanel() {
  const panel = document.getElementById('comments-panel');
  panel.classList.toggle('hidden');
  if (!panel.classList.contains('hidden') && !isReviewMode) loadComments();
}

// ============================================================
//  SAVE / LOAD
// ============================================================
function setSaveStatus(msg, colour) {
  const el = document.getElementById('save-status');
  if (el) { el.textContent = msg; el.style.color = colour || 'var(--text3)'; }
}

async function saveDocument() {
  setSaveStatus('Saving…');
  calcAll();
  const name      = document.getElementById('doc-title').value.trim() || 'Untitled BOTEC';
  const programme = document.getElementById('programme').value.trim();
  const data      = serialiseState();
  if (docId) {
    const { error } = await sb.from('botec_documents').update({ name, programme, data }).eq('id', docId);
    if (error) { setSaveStatus('Save failed', '#c0392b'); alert(error.message); return; }
  } else {
    const { data: ins, error } = await sb.from('botec_documents')
      .insert({ user_id: currentUser.id, name, programme, data }).select('id').single();
    if (error) { setSaveStatus('Save failed', '#c0392b'); alert(error.message); return; }
    docId = ins.id;
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
  if (data.is_shared) {
    document.getElementById('share-btn').textContent = '🔗 Sharing (click to stop)';
    document.getElementById('share-btn').onclick = stopSharing;
  }
  deserialiseState(data.data);
  setSaveStatus('');
}

function serialiseState() {
  calcAll();
  return {
    projName: v('projName'), programme: v('programme'),
    prepBy: v('prepBy'), prepDate: v('prepDate'),
    reviewBy: v('reviewBy'), reviewDate: v('reviewDate'),
    currency: v('currency'), numYears: v('numYears'),
    bufferPct: v('bufferPct'), mgrMult: v('mgrMult'),
    purposeNote: v('purposeNote'), logSet: [...logSet],
    benMode,
    benY1: v('benY1'), benGrowth: v('benGrowth'),
    benY2manual: v('benY2manual'), benY3manual: v('benY3manual'),
    benNotes: v('benNotes'),
    costRows: getCostData(),
    cpbAvg: lastCalc.cpbAvg, totalBen: lastCalc.totalBen, totalAll: lastCalc.totalAll
  };
}

function deserialiseState(s) {
  if (!s) return;
  const set2 = (id, val) => { const el = document.getElementById(id); if (el && val != null) el.value = val; };
  set2('projName', s.projName); set2('programme', s.programme);
  set2('prepBy', s.prepBy); set2('prepDate', s.prepDate);
  set2('reviewBy', s.reviewBy); set2('reviewDate', s.reviewDate);
  set2('currency', s.currency); set2('numYears', s.numYears);
  set2('bufferPct', s.bufferPct); set2('mgrMult', s.mgrMult);
  set2('purposeNote', s.purposeNote);
  if (s.logSet) { logSet = new Set(s.logSet); renderLogFlags(); }
  if (s.benMode) setBenMode(s.benMode, true);
  set2('benY1', s.benY1); set2('benGrowth', s.benGrowth || '0');
  set2('benY2manual', s.benY2manual); set2('benY3manual', s.benY3manual);
  set2('benNotes', s.benNotes);
  calcBens();
  document.getElementById('costBody').innerHTML = '';
  cID = 0;
  (s.costRows || []).forEach(r => addCost(r));
  calcAll();
}

const v = id => { const el = document.getElementById(id); return el ? el.value : ''; };

// ============================================================
//  LOG FLAGS
// ============================================================
function renderLogFlags() {
  const el = document.getElementById('logFlags');
  if (!el) return;
  el.innerHTML = HEADS.map(h => `
    <label style="display:flex;align-items:center;gap:5px;font-size:13px;cursor:pointer;padding:4px 10px;
      border:0.5px solid var(--border);border-radius:20px;
      background:${logSet.has(h) ? '#e0f7fa' : 'var(--surface2)'}">
      <input type="checkbox" ${logSet.has(h) ? 'checked' : ''} onchange="toggleLog('${h}',this.checked)"
        style="width:auto;margin:0"> ${h}
    </label>`).join('');
}
function toggleLog(h, val) { val ? logSet.add(h) : logSet.delete(h); renderLogFlags(); calcAll(); }

// ============================================================
//  TAB SWITCHING
// ============================================================
function go(id) {
  document.querySelectorAll('.tab').forEach((t, i) => {
    t.classList.toggle('on', ['setup','ben','costs','results'][i] === id);
  });
  document.querySelectorAll('.pane').forEach(p => p.classList.remove('on'));
  document.getElementById('p-' + id).classList.add('on');
  if (id === 'results') { calcAll(); renderResults(); }
}

// ============================================================
//  BENEFICIARIES
// ============================================================
function setBenMode(mode, silent) {
  benMode = mode;
  document.getElementById('btn-mode-growth').classList.toggle('mode-active', mode === 'growth');
  document.getElementById('btn-mode-manual').classList.toggle('mode-active', mode === 'manual');
  document.getElementById('ben-growth-row').classList.toggle('hidden', mode !== 'growth');
  document.getElementById('ben-manual-row').classList.toggle('hidden', mode !== 'manual');
  if (!silent) calcBens();
}

function calcBens() {
  const { b1, b2, b3 } = getBenTotals();
  const NY  = parseInt(v('numYears')) || 3;
  const growth = parseFloat(v('benGrowth')) || 0;
  const fmt = n => n > 0 ? Math.round(n).toLocaleString() : '—';

  const s1 = document.getElementById('ben-show-1');
  const s2 = document.getElementById('ben-show-2');
  const s3 = document.getElementById('ben-show-3');
  const g2 = document.getElementById('ben-growth-2');
  const g3 = document.getElementById('ben-growth-3');
  const tot = document.getElementById('ben-total');

  if (s1) s1.textContent = fmt(b1);
  if (s2) s2.textContent = NY >= 2 ? fmt(b2) : '—';
  if (s3) s3.textContent = NY >= 3 ? fmt(b3) : '—';

  if (g2) g2.textContent = (benMode === 'growth' && growth !== 0 && b1 > 0) ? `${growth > 0 ? '+' : ''}${growth}% vs Y1` : '';
  if (g3) g3.textContent = (benMode === 'growth' && growth !== 0 && b1 > 0) ? `${growth > 0 ? '+' : ''}${(growth*2).toFixed(1)}% vs Y1` : '';

  const total = b1 + (NY >= 2 ? b2 : 0) + (NY >= 3 ? b3 : 0);
  if (tot) tot.textContent = total > 0 ? Math.round(total).toLocaleString() : '—';
  calcAll();
}

function getBenTotals() {
  const y1 = parseFloat(v('benY1')) || 0;
  if (benMode === 'manual') {
    return { b1: y1, b2: parseFloat(v('benY2manual')) || 0, b3: parseFloat(v('benY3manual')) || 0 };
  }
  const g = 1 + (parseFloat(v('benGrowth')) || 0) / 100;
  return { b1: y1, b2: y1 * g, b3: y1 * g * g };
}

// ============================================================
//  COST ROWS
// ============================================================
function addCost(d = {}) {
  const id  = ++cID;
  const tr  = document.createElement('tr');
  tr.id     = 'cr' + id;
  const sel = HEADS.map(h => `<option value="${h}" ${h === (d.h || d.head || HEADS[0]) ? 'selected' : ''}>${h}</option>`).join('');
  const cpu = d.cpu != null ? d.cpu : '';
  const u1  = (d.u1 != null && d.u1 !== '') ? d.u1 : (d.y1 !== '' && d.y1 != null ? d.y1 : '');
  const u2  = (d.u2 != null && d.u2 !== '') ? d.u2 : (d.y2 !== '' && d.y2 != null ? d.y2 : '');
  const u3  = (d.u3 != null && d.u3 !== '') ? d.u3 : (d.y3 !== '' && d.y3 != null ? d.y3 : '');
  const guide  = d.guide  || '';
  const justif = d.justif || d.justification || '';

  tr.innerHTML = `
    <td><select onchange="calcAll()">${sel}</select></td>
    <td>
      <input type="text" value="${escHtml(d.d || d.desc || '')}">
      ${guide ? `<div class="cost-guide">${guide}</div>` : ''}
    </td>
    <td>
      <textarea rows="2" placeholder="Why this cost? What is the source or assumption?" style="font-size:12px;resize:vertical;min-height:36px">${escHtml(justif)}</textarea>
    </td>
    <td><input type="text" value="${escHtml(d.u || d.unit || '')}" placeholder="e.g. staff member, KG"></td>
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
function escHtml(s) { return String(s).replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function updateCostRow(inp) {
  const tr  = inp.closest('tr');
  const ins = tr.querySelectorAll('input[type=number]');
  const cpu = parseFloat(ins[0].value) || 0;
  const u1  = parseFloat(ins[1].value) || 0;
  const u2  = parseFloat(ins[2].value) || 0;
  const u3  = parseFloat(ins[3].value) || 0;
  const id  = tr.id.replace('cr', '');
  const fmt = n => getSym() + Math.round(n).toLocaleString();
  ['cy1_','cy2_','cy3_'].forEach((p, i) => {
    const el = document.getElementById(p + id);
    if (el) el.textContent = fmt(cpu * [u1, u2, u3][i]);
  });
}

function getCostData() {
  return Array.from(document.querySelectorAll('#costBody tr')).map(tr => {
    const sels    = tr.querySelectorAll('select');
    const inputs  = tr.querySelectorAll('input');
    const textareas = tr.querySelectorAll('textarea');
    const numIns  = tr.querySelectorAll('input[type=number]');
    const cpu = parseFloat(numIns[0]?.value) || 0;
    const u1  = parseFloat(numIns[1]?.value) || 0;
    const u2  = parseFloat(numIns[2]?.value) || 0;
    const u3  = parseFloat(numIns[3]?.value) || 0;
    return {
      head:   sels[0]?.value || '',
      desc:   inputs[0]?.value || '',
      justif: textareas[0]?.value || '',
      unit:   inputs[1]?.value || '',
      cpu, u1, u2, u3, cy1: cpu*u1, cy2: cpu*u2, cy3: cpu*u3
    };
  });
}

// ============================================================
//  HELPERS
// ============================================================
const getSym   = () => ({ INR:'₹', USD:'$', GBP:'£', EUR:'€' })[v('currency')] || '₹';
const nv       = id => parseFloat(v(id)) || 0;
const fC       = n => getSym() + Math.round(n).toLocaleString();
const fCd      = (n, d=2) => getSym() + n.toFixed(d);

function showToast(msg) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  t.classList.remove('hidden');
  setTimeout(() => t.classList.add('hidden'), 3500);
}

// ============================================================
//  CORE CALCULATIONS
// ============================================================
function calcAll() {
  const rows = getCostData();
  const { b1, b2, b3 } = getBenTotals();
  const buf = nv('bufferPct') / 100;
  const mgr = nv('mgrMult')   / 100;
  const NY  = parseInt(v('numYears')) || 3;

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
  const totalAll = yrTotals.reduce((s, x) => s + x, 0);
  const totalBen = bens.reduce((s, x) => s + x, 0);
  const avgCost  = totalAll / NY;

  const cpbY   = [b1>0?totY.y1/b1:0, b2>0?totY.y2/b2:0, b3>0?totY.y3/b3:0];
  const cpbAvg = totalBen > 0 ? totalAll / totalBen : 0;

  const nonLogBase = { y1:0, y2:0, y3:0 };
  HEADS.forEach(h => {
    if (!logSet.has(h)) { nonLogBase.y1 += byHead[h].y1; nonLogBase.y2 += byHead[h].y2; nonLogBase.y3 += byHead[h].y3; }
  });
  nonLogBase.y1 += mgrVal.y1; nonLogBase.y2 += mgrVal.y2; nonLogBase.y3 += mgrVal.y3;

  const cpbExclY   = [b1>0?nonLogBase.y1*(1+buf)/b1:0, b2>0?nonLogBase.y2*(1+buf)/b2:0, b3>0?nonLogBase.y3*(1+buf)/b3:0];
  const cpbExclAvg = totalBen > 0 ? (nonLogBase.y1+nonLogBase.y2+nonLogBase.y3)*(1+buf)/totalBen : 0;

  lastCalc = { rows, byHead, mgrVal, bufVal, subY, totY, totalAll, totalBen, avgCost,
               bens, yrTotals, cpbY, cpbAvg, cpbExclY, cpbExclAvg, buf, mgr, NY, b1, b2, b3 };
}

// ============================================================
//  RENDER RESULTS
// ============================================================
function renderResults() {
  calcAll();
  const d  = lastCalc;
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
    <div class="card"><div class="lbl">Total beneficiaries</div><div class="val">${d.totalBen >= 1e5 ? (d.totalBen/1000).toFixed(1)+'K' : Math.round(d.totalBen).toLocaleString()}</div></div>`;

  const yrs = Array.from({ length: NY }, (_, i) => `Year ${i+1}`);
  document.getElementById('summHead').innerHTML =
    `<th>Cost head</th>${yrs.map(y=>`<th style="text-align:right">${y}</th>`).join('')}<th style="text-align:right">Avg/year</th><th style="text-align:right">${NY}-yr total</th>`;

  const bd = document.getElementById('summBody');
  bd.innerHTML = '';
  const ah = HEADS.filter(h => d.byHead[h] && d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3 > 0);
  ah.forEach(h => {
    const vv = d.byHead[h], vals = [vv.y1,vv.y2,vv.y3].slice(0,NY), t = vals.reduce((s,x)=>s+x,0);
    bd.innerHTML += `<tr><td>${h}</td>${vals.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(t/NY)}</td><td style="text-align:right">${fC(t)}</td></tr>`;
  });

  const mv = [d.mgrVal.y1,d.mgrVal.y2,d.mgrVal.y3].slice(0,NY), mt = mv.reduce((s,x)=>s+x,0);
  if (mt>0) bd.innerHTML += `<tr class="derived-row"><td>${(d.mgr*100).toFixed(0)}% managerial multiplier</td>${mv.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(mt/NY)}</td><td style="text-align:right">${fC(mt)}</td></tr>`;

  const bv = [d.bufVal.y1,d.bufVal.y2,d.bufVal.y3].slice(0,NY), bt = bv.reduce((s,x)=>s+x,0);
  if (bt>0) bd.innerHTML += `<tr class="derived-row"><td>${(d.buf*100).toFixed(0)}% buffer</td>${bv.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(bt/NY)}</td><td style="text-align:right">${fC(bt)}</td></tr>`;

  bd.innerHTML += `<tr class="grand-row"><td>Total costs</td>${d.yrTotals.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(d.avgCost)}</td><td style="text-align:right">${fC(d.totalAll)}</td></tr>`;
  bd.innerHTML += `<tr class="derived-row"><td>Beneficiaries</td>${d.bens.map(x=>`<td style="text-align:right">${Math.round(x).toLocaleString()}</td>`).join('')}<td style="text-align:right">${Math.round(d.totalBen/NY).toLocaleString()}</td><td style="text-align:right">${Math.round(d.totalBen).toLocaleString()}</td></tr>`;
  bd.innerHTML += `<tr class="cpb-row"><td>Cost per beneficiary</td>${d.cpbY.slice(0,NY).map(x=>`<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbAvg)}</td><td style="text-align:right">${fCd(d.cpbAvg)}</td></tr>`;
  if (logSet.size > 0) bd.innerHTML += `<tr class="cpb-row" style="font-style:italic"><td>CPB excl. ${[...logSet].join(' + ')}</td>${d.cpbExclY.slice(0,NY).map(x=>`<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbExclAvg)}</td><td style="text-align:right">${fCd(d.cpbExclAvg)}</td></tr>`;

  if (chart) chart.destroy();
  const colors = ['#0097a7','#dc6059','#ff8dcb','#00bcd4','#f06292','#4dd0e1','#e57373','#80deea'];
  const datasets = ah.map((h, i) => ({ label:h, data:[d.byHead[h].y1,d.byHead[h].y2,d.byHead[h].y3].slice(0,NY), backgroundColor:colors[i%colors.length], stack:'s' }));
  if (mt > 0) datasets.push({ label:'Mgr multiplier', data:mv, backgroundColor:'#CCC', stack:'s' });
  if (bt > 0) datasets.push({ label:'Buffer',          data:bv, backgroundColor:'#E0E0E0', stack:'s' });
  chart = new Chart(document.getElementById('chartC').getContext('2d'), {
    type:'bar', data:{ labels:yrs, datasets },
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label:ctx=>`${ctx.dataset.label}: ${fC(ctx.parsed.y)}` } } },
      scales:{ x:{ stacked:true }, y:{ stacked:true, ticks:{ callback:val=>getSym()+(val>=1e7?(val/1e7).toFixed(1)+'Cr':val>=1e5?(val/1e5).toFixed(0)+'L':val>=1000?(val/1000).toFixed(0)+'K':val) } } } }
  });
}

// ============================================================
//  EXPORTS
// ============================================================
function dlCSV() {
  calcAll(); const d = lastCalc;
  const pn = v('projName') || 'BOTEC';
  let c = `BOTEC Cost Per Beneficiary Estimate\nProject,${pn}\n\n`;
  c += `BENEFICIARIES\nMode,${benMode}\nYear 1,${Math.round(d.b1)}\nAnnual growth %,${v('benGrowth')||0}\nYear 2,${Math.round(d.b2)}\nYear 3,${Math.round(d.b3)}\nTotal,${Math.round(d.totalBen)}\n\n`;
  c += `COST ITEMS\nCost Head,Description,Justification,Unit,Monthly Cost/unit,Units Y1,Cost Y1,Units Y2,Cost Y2,Units Y3,Cost Y3\n`;
  d.rows.forEach(r => { c += `${r.head},"${r.desc}","${r.justif}","${r.unit}",${r.cpu},${r.u1},${r.cy1},${r.u2},${r.cy2},${r.u3},${r.cy3}\n`; });
  c += `\nSUMMARY\nTotal,${d.totY.y1},${d.totY.y2},${d.totY.y3},${d.avgCost},${d.totalAll}\nCPB,${d.cpbY.join(',')},${d.cpbAvg}\n`;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([c], { type:'text/csv' }));
  a.download = `BOTEC_${pn.replace(/\s+/g,'_')}.csv`; a.click();
}

function dlJSON() {
  calcAll();
  const blob = new Blob([JSON.stringify(serialiseState(), null, 2)], { type:'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `BOTEC_${(v('projName')||'export').replace(/\s+/g,'_')}.json`;
  a.click();
}

function dlXLSX() {
  calcAll(); const d = lastCalc;
  const WB  = XLSX.utils.book_new();
  const pn  = v('projName') || 'BOTEC';
  const cur = v('currency');
  const NY  = d.NY || 3;

  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet([
    ['READ-ME COST PER BENEFICIARY ESTIMATE'],[''],
    ['Purpose', v('purposeNote')],[''],
    ['Tabs','Details'],['Summary','Cost per Beneficiary'],['Beneficiary Calculation','Beneficiary numbers'],
    ['Cost Calculation','All cost line items with justifications'],['Unit Costs','Unit cost reference']
  ]), 'Read Me');

  const sumRows = [
    ['COST PER BENEFICIARY ESTIMATE'],[''],
    ['Project Name:', pn],['Prepared By:', v('prepBy')],['Preparation Date:', v('prepDate')],[''],
    ['Reviewed By:', v('reviewBy')],[''],
    ['','','Year 1','Year 2','Year 3','Average Cost','Total Cost'],
    ['Number of Beneficiaries','', d.b1, d.b2, d.b3, d.totalBen/NY, d.totalBen],[''],['Costs (in '+cur+'):']
  ];
  HEADS.filter(h=>d.byHead[h]&&d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3>0).forEach(h=>{
    const vv=d.byHead[h],t=vv.y1+vv.y2+vv.y3; sumRows.push(['',h,vv.y1,vv.y2,vv.y3,t/NY,t]);
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
    ['Mode', benMode === 'growth' ? 'Annual % growth' : 'Manual entry'],
    ['Year 1 Beneficiaries', Math.round(d.b1)],
    ['Annual Growth %', v('benGrowth')||0],
    ['Year 2 Beneficiaries', Math.round(d.b2)],
    ['Year 3 Beneficiaries', Math.round(d.b3)],
    ['Notes', v('benNotes')]
  ]), 'Beneficiary Calculation');

  const cc=[['COST CALCULATION'],[''],
    ['COST HEAD','DESCRIPTION','JUSTIFICATION / NOTES','UNIT',`MONTHLY COST/UNIT (${cur})`,'UNITS Y1','UNITS Y2','UNITS Y3','Cost Y1','Cost Y2','Cost Y3']];
  d.rows.forEach(r=>{cc.push([r.head,r.desc,r.justif,r.unit,r.cpu,r.u1,r.u2,r.u3,r.cy1,r.cy2,r.cy3]);});
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(cc), 'Cost Calculation');

  const uc=[['Cost Head','Unit Description','Justification',`Monthly Cost/Unit (${cur})`,'Unit type']];
  const seen=new Set();
  d.rows.forEach(r=>{const k=r.head+'|'+r.desc;if(!seen.has(k)){seen.add(k);uc.push([r.head,r.desc,r.justif,r.cpu,r.unit]);}});
  XLSX.utils.book_append_sheet(WB, XLSX.utils.aoa_to_sheet(uc), 'Unit Costs');

  XLSX.writeFile(WB, `BOTEC_${pn.replace(/\s+/g,'_')}.xlsx`);
}

function dlPDF() {
  calcAll(); const d = lastCalc;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
  const pn  = v('projName') || 'BOTEC';
  const cur = v('currency');
  const NY  = d.NY || 3;
  const s   = getSym();
  const fmtN = n => s + Math.round(n).toLocaleString();
  const fmtD = (n, dp=2) => s + n.toFixed(dp);

  doc.setFillColor(0,151,167); doc.rect(0,0,297,18,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(13); doc.setFont('helvetica','bold');
  doc.text('BOTEC Cost Per Beneficiary Estimate', 14, 12);
  doc.setFontSize(9); doc.setFont('helvetica','normal'); doc.text(pn, 220, 12);
  doc.setTextColor(80,80,80); doc.setFontSize(8);
  doc.text(`Prepared by: ${v('prepBy')||'—'}   Date: ${v('prepDate')||'—'}   Programme: ${v('programme')||'—'}`, 14, 26);

  const cards=[
    {label:'Cost per beneficiary (avg)',val:fmtD(d.cpbAvg),color:[220,96,89]},
    {label:`Total ${NY}-year cost`,val:fmtN(d.totalAll),color:[0,151,167]},
    {label:'Avg yearly cost',val:fmtN(d.avgCost),color:[255,141,203]},
    {label:'Total beneficiaries',val:Math.round(d.totalBen).toLocaleString(),color:[0,151,167]}
  ];
  cards.forEach((c,i)=>{
    doc.setFillColor(...c.color); doc.roundedRect(14+i*71,30,66,18,2,2,'F');
    doc.setTextColor(255,255,255); doc.setFontSize(7); doc.setFont('helvetica','normal');
    doc.text(c.label,18+i*71,36); doc.setFontSize(11); doc.setFont('helvetica','bold');
    doc.text(c.val,18+i*71,44);
  });

  const yrs=Array.from({length:NY},(_,i)=>`Year ${i+1}`);
  const ah=HEADS.filter(h=>d.byHead[h]&&d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3>0);
  const tableRows=[];
  ah.forEach(h=>{const vv=d.byHead[h],vals=[vv.y1,vv.y2,vv.y3].slice(0,NY),t=vals.reduce((s,x)=>s+x,0);tableRows.push([h,...vals.map(x=>fmtN(x)),fmtN(t/NY),fmtN(t)]);});
  const mv=[d.mgrVal.y1,d.mgrVal.y2,d.mgrVal.y3].slice(0,NY),mt=mv.reduce((s,x)=>s+x,0);
  if(mt>0) tableRows.push([`${(d.mgr*100).toFixed(0)}% mgr`,...mv.map(x=>fmtN(x)),fmtN(mt/NY),fmtN(mt)]);
  const bv=[d.bufVal.y1,d.bufVal.y2,d.bufVal.y3].slice(0,NY),bt=bv.reduce((s,x)=>s+x,0);
  if(bt>0) tableRows.push([`${(d.buf*100).toFixed(0)}% buffer`,...bv.map(x=>fmtN(x)),fmtN(bt/NY),fmtN(bt)]);
  tableRows.push(['TOTAL COSTS',...d.yrTotals.map(x=>fmtN(x)),fmtN(d.avgCost),fmtN(d.totalAll)]);
  tableRows.push(['Beneficiaries',...d.bens.map(x=>Math.round(x).toLocaleString()),Math.round(d.totalBen/NY).toLocaleString(),Math.round(d.totalBen).toLocaleString()]);
  tableRows.push(['Cost per beneficiary',...d.cpbY.slice(0,NY).map(x=>fmtD(x)),fmtD(d.cpbAvg),fmtD(d.cpbAvg)]);
  if(logSet.size>0) tableRows.push([`CPB excl. ${[...logSet].join('+')}`,...d.cpbExclY.slice(0,NY).map(x=>fmtD(x)),fmtD(d.cpbExclAvg),fmtD(d.cpbExclAvg)]);

  doc.autoTable({ startY:55, head:[['Cost head',...yrs,'Avg/year',`${NY}-yr total`]], body:tableRows,
    styles:{fontSize:8,cellPadding:3}, headStyles:{fillColor:[0,151,167],textColor:255,fontStyle:'bold'},
    alternateRowStyles:{fillColor:[245,245,245]},
    didParseCell:data=>{ if(['TOTAL COSTS','Cost per beneficiary'].includes(data.row.raw[0])) data.cell.styles.fontStyle='bold'; if(data.row.raw[0].startsWith('Cost per')||data.row.raw[0].startsWith('CPB excl')) data.cell.styles.textColor=[220,96,89]; },
    margin:{left:14,right:14} });

  doc.addPage();
  doc.setFillColor(0,151,167); doc.rect(0,0,297,18,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(13); doc.setFont('helvetica','bold');
  doc.text('Cost Line Items & Justifications', 14, 12);
  doc.autoTable({ startY:25,
    head:[['Cost head','Description','Justification','Unit',`Monthly cost/unit`,`Units Y1`,`Cost Y1`,`Units Y2`,`Cost Y2`,`Units Y3`,`Cost Y3`]],
    body:d.rows.map(r=>[r.head,r.desc,r.justif||'',r.unit,fmtN(r.cpu),r.u1,fmtN(r.cy1),r.u2,fmtN(r.cy2),r.u3,fmtN(r.cy3)]),
    styles:{fontSize:7,cellPadding:2}, headStyles:{fillColor:[0,151,167],textColor:255,fontStyle:'bold'},
    alternateRowStyles:{fillColor:[245,245,245]}, margin:{left:14,right:14} });

  const pc=doc.internal.getNumberOfPages();
  for(let i=1;i<=pc;i++){ doc.setPage(i); doc.setFontSize(7); doc.setTextColor(150,150,150); doc.setFont('helvetica','normal'); doc.text(`${pn} — BOTEC Estimate`,14,205); doc.text(`Page ${i} of ${pc}`,270,205); }
  doc.save(`BOTEC_${pn.replace(/\s+/g,'_')}.pdf`);
}

// ---- KEYBOARD SHORTCUT ----
document.addEventListener('input', () => { if (!isReviewMode) setSaveStatus('Unsaved'); });
document.addEventListener('keydown', e => {
  if ((e.ctrlKey||e.metaKey) && e.key==='s') { e.preventDefault(); if (!isReviewMode) saveDocument(); }
});

init();

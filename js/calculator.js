// BOTEC Calculator JS — full rewrite

const HEADS = ['Internal Consulting','Travel Costs','Premix Costs','Equipment Costs',
  'M&E Costs','Logistics Costs','Packaging Costs','Event / Admin Costs','Other'];

const PRESETS = {
  c1:        { h:'Internal Consulting', d:'Partnerships Manager',          u:'Person-month', cpu:105000, consulting:true,
               guide:'Hours/month varies by year. Fill in the hours calculator below.' },
  c2:        { h:'Internal Consulting', d:'Senior Partnerships Officer',   u:'Person-month', cpu:85000,  consulting:true,
               guide:'Hours/month varies by year. Fill in the hours calculator below.' },
  c3:        { h:'Internal Consulting', d:'Senior Programs Officer',       u:'Person-month', cpu:85000,  consulting:true,
               guide:'Hours/month varies by year. Fill in the hours calculator below.' },
  travel:    { h:'Travel Costs',        d:'Travel Partnerships Manager',   u:'Person-month', cpu:25000,  travelRow:true,
               guide:'Enter months travelling per year and number of people.' },
  travel2:   { h:'Travel Costs',        d:'Travel Senior Partnerships Officer', u:'Person-month', cpu:15000, travelRow:true,
               guide:'Enter months travelling per year and number of people.' },
  travel3:   { h:'Travel Costs',        d:'Travel Senior Programs Officer',u:'Person-month', cpu:15000,  travelRow:true,
               guide:'Enter months travelling per year and number of people.' },
  premix:    { h:'Premix Costs',        d:'NaFeEDTA premix',               u:'KG',           cpu:400,    attaType:'premix',
               guide:'Auto-calculated from atta consumption (yearly MT ÷ 5). Ratio of premix to flour is 1:5.' },
  equip:     { h:'Equipment Costs',     d:'Microdoser',                    u:'Device',       cpu:200000 },
  nabl:      { h:'M&E Costs',           d:'NABL testing',                  u:'Test',         cpu:4000,   maeType:'nabl' },
  icheck:    { h:'M&E Costs',           d:'I-check testing',               u:'Test',         cpu:1500,   maeType:'icheck' },
  ironspot:  { h:'M&E Costs',           d:'Iron spot test kit',            u:'Kit',          cpu:1750,
               guide:'Iron Spot Test (1L per kit, 10ml/day). 1 mill per district × 5 districts = 5 mills. 10 IST kits per year across 5 mills (tested daily). Enter total kits per year.' },
  transport: { h:'Logistics Costs',     d:'Transportation cost',           u:'KG atta',      cpu:1,      attaType:'transport',
               guide:'Auto-calculated from atta consumption (yearly KG). ₹1.00 per KG of wheat flour transported.' },
  grinding:  { h:'Logistics Costs',     d:'Grinding cost',                 u:'KG wheat',     cpu:3,      attaType:'grinding',
               guide:'Auto-calculated from atta consumption (yearly KG). ₹3.00 per KG of wheat ground.' },
  packaging: { h:'Packaging Costs',     d:'Packaging cost',                u:'KG wheat flour',cpu:0.5,   attaType:'packaging',
               guide:'Auto-calculated from atta consumption (yearly KG). ₹0.50 per KG of wheat flour packaged.' },
  event:     { h:'Event / Admin Costs', d:'Awareness programme',           u:'Event',        cpu:21000,
               guide:'Enter number of events per year.' }
};

// ── State ────────────────────────────────────────────────────
let logSet      = new Set(['Logistics Costs']);
let chart       = null;
let lastCalc    = {};
let lastAtta    = { mt1:0, mt2:0, mt3:0, kg1:0, kg2:0, kg3:0 };
let docId       = null;
let currentUser = null;
let cID         = 0;
let benMode     = 'growth';
let isReviewer  = false;
const params    = new URLSearchParams(window.location.search);

// ── Init ─────────────────────────────────────────────────────
async function init() {
  const { data: { session } } = await sb.auth.getSession();
  if (!session) { window.location.href = 'login.html'; return; }
  currentUser = session.user;
  renderLogFlags();

  docId = params.get('id');
  if (docId) {
    await loadDocument(docId);
    await loadNotifications();
    await loadComments();
  }

  sb.channel('notif-' + currentUser.id)
    .on('postgres_changes', { event:'INSERT', schema:'public', table:'notifications', filter:`user_id=eq.${currentUser.id}` },
      () => loadNotifications())
    .subscribe();
}

// ── Notifications ─────────────────────────────────────────────
async function loadNotifications() {
  const { data } = await sb.from('notifications').select('*')
    .eq('user_id', currentUser.id).order('created_at', { ascending: false }).limit(20);
  if (!data) return;
  const unread = data.filter(n => !n.read).length;
  const badge  = document.getElementById('notif-badge');
  if (badge) { badge.textContent = unread; badge.style.display = unread > 0 ? 'flex' : 'none'; }
  renderNotifications(data);
}

function renderNotifications(notifs) {
  const el = document.getElementById('notif-list');
  if (!el) return;
  if (!notifs.length) { el.innerHTML = '<p class="empty-notif">No notifications yet.</p>'; return; }
  el.innerHTML = notifs.map(n => `
    <div class="notif-item ${n.read ? 'read' : 'unread'}" onclick="openNotif('${n.id}','${n.doc_id}')">
      <div class="notif-icon">${n.type === 'review_requested' ? '📋' : n.type === 'comment_added' ? '💬' : '✅'}</div>
      <div class="notif-body">
        <p class="notif-msg">${n.message}</p>
        <p class="notif-time">${new Date(n.created_at).toLocaleDateString('en-GB',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'})}</p>
      </div>
      ${!n.read ? '<div class="notif-dot"></div>' : ''}
    </div>`).join('');
}

async function openNotif(notifId, docId) {
  await sb.from('notifications').update({ read: true }).eq('id', notifId);
  if (docId) window.location.href = `calculator.html?id=${docId}`;
  toggleNotifPanel();
}

function toggleNotifPanel() {
  document.getElementById('notif-panel').classList.toggle('hidden');
  loadNotifications();
}

async function createNotification(userId, docId, type, message) {
  await sb.from('notifications').insert({ user_id: userId, doc_id: docId, type, message });
}

// ── Reviewer system ───────────────────────────────────────────
function toggleReviewerPanel() {
  document.getElementById('reviewer-panel').classList.toggle('hidden');
  document.getElementById('reviewer-search-input').value = '';
  document.getElementById('reviewer-search-results').innerHTML = '';
  updateReviewerStatus();
}

async function updateReviewerStatus() {
  if (!docId) return;
  const { data } = await sb.from('botec_documents')
    .select('reviewer_id, reviewer_status, profiles:reviewer_id(full_name,email)')
    .eq('id', docId).single();
  if (!data) return;

  const el = document.getElementById('current-reviewer-info');
  if (!el) return;
  if (data.reviewer_id && data.profiles) {
    const statusLabel = { none:'—', requested:'Review requested', in_review:'In review', complete:'Review complete' };
    el.innerHTML = `
      <div class="reviewer-assigned">
        <div class="reviewer-avatar">${(data.profiles.full_name||'?')[0].toUpperCase()}</div>
        <div>
          <div style="font-weight:500;font-size:14px">${data.profiles.full_name || data.profiles.email}</div>
          <div style="font-size:12px;color:var(--text3)">${data.profiles.email}</div>
          <div class="reviewer-status-badge status-${data.reviewer_status}">${statusLabel[data.reviewer_status]||data.reviewer_status}</div>
        </div>
        <button class="btn-ghost btn-sm danger" onclick="removeReviewer()" style="margin-left:auto">Remove</button>
      </div>`;
  } else {
    el.innerHTML = '<p style="font-size:13px;color:var(--text3)">No reviewer assigned yet.</p>';
  }
}

async function searchReviewer() {
  const query = document.getElementById('reviewer-search-input').value.trim();
  if (query.length < 2) return;
  const { data } = await sb.from('profiles').select('id,full_name,email')
    .ilike('email', `%${query}%`).neq('id', currentUser.id).limit(5);
  const el = document.getElementById('reviewer-search-results');
  if (!data || !data.length) { el.innerHTML = '<p style="font-size:13px;color:var(--text3);padding:8px 0">No users found with that email.</p>'; return; }
  el.innerHTML = data.map(p => `
    <div class="reviewer-result" onclick="assignReviewer('${p.id}','${escHtml(p.full_name||p.email)}','${escHtml(p.email)}')">
      <div class="reviewer-avatar sm">${(p.full_name||p.email||'?')[0].toUpperCase()}</div>
      <div>
        <div style="font-size:13px;font-weight:500">${p.full_name || p.email}</div>
        <div style="font-size:12px;color:var(--text3)">${p.email}</div>
      </div>
      <button class="btn-primary btn-sm" style="margin-left:auto">Assign</button>
    </div>`).join('');
}

async function assignReviewer(reviewerId, reviewerName, reviewerEmail) {
  if (!docId) { alert('Please save the document first.'); return; }
  const { error } = await sb.from('botec_documents')
    .update({ reviewer_id: reviewerId, reviewer_status: 'requested' }).eq('id', docId);
  if (error) { alert('Error assigning reviewer: ' + error.message); return; }

  const docName = document.getElementById('doc-title').value || 'Untitled BOTEC';
  const creatorName = currentUser.user_metadata?.full_name || currentUser.email;
  await createNotification(reviewerId, docId, 'review_requested',
    `${creatorName} has asked you to review "${docName}"`);

  showToast(`Review request sent to ${reviewerName}`);
  document.getElementById('reviewer-search-results').innerHTML = '';
  document.getElementById('reviewer-search-input').value = '';
  await updateReviewerStatus();
}

async function removeReviewer() {
  const { error } = await sb.from('botec_documents')
    .update({ reviewer_id: null, reviewer_status: 'none' }).eq('id', docId);
  if (error) { alert('Error: ' + error.message); return; }
  await updateReviewerStatus();
  showToast('Reviewer removed.');
}

async function markReviewComplete() {
  const { data: doc } = await sb.from('botec_documents').select('user_id,name').eq('id', docId).single();
  await sb.from('botec_documents').update({ reviewer_status: 'complete' }).eq('id', docId);
  const reviewerName = currentUser.user_metadata?.full_name || currentUser.email;
  await createNotification(doc.user_id, docId, 'review_complete',
    `${reviewerName} has completed their review of "${doc.name}"`);
  showToast('Review marked as complete. The document owner has been notified.');
}

// ── Comments ─────────────────────────────────────────────────
async function loadComments() {
  const { data } = await sb.from('botec_comments').select('*').eq('doc_id', docId).order('created_at');
  if (!data) return;
  renderComments(data);
  const unresolved = data.filter(c => !c.resolved).length;
  const badge = document.getElementById('comments-badge');
  if (badge) { badge.textContent = unresolved; badge.style.display = unresolved > 0 ? 'flex' : 'none'; }
}

function renderComments(comments) {
  const el = document.getElementById('comments-list');
  if (!el) return;
  if (!comments.length) { el.innerHTML = '<p class="empty-notif">No comments yet.</p>'; return; }
  el.innerHTML = comments.map(c => `
    <div class="comment-item ${c.resolved ? 'resolved' : ''}">
      <div class="comment-meta">
        <span class="comment-author">${c.reviewer_name}</span>
        <span class="comment-section">${c.section}</span>
        <span class="comment-time">${new Date(c.created_at).toLocaleDateString('en-GB',{day:'numeric',month:'short'})}</span>
      </div>
      <p class="comment-text">${c.comment}</p>
      ${!c.resolved && !isReviewer ? `<button class="resolve-btn" onclick="resolveComment('${c.id}')">✓ Mark resolved</button>` : c.resolved ? '<span class="resolved-label">✓ Resolved</span>' : ''}
    </div>`).join('');
}

async function submitComment() {
  const name    = currentUser.user_metadata?.full_name || currentUser.email;
  const section = document.getElementById('comment-section').value;
  const text    = document.getElementById('comment-text').value.trim();
  if (!text) { alert('Please enter a comment.'); return; }

  const { error } = await sb.from('botec_comments').insert({ doc_id: docId, reviewer_name: name, section, comment: text });
  if (error) { alert('Error: ' + error.message); return; }

  if (isReviewer) {
    const { data: doc } = await sb.from('botec_documents').select('user_id,name').eq('id', docId).single();
    await createNotification(doc.user_id, docId, 'comment_added',
      `${name} left a comment on "${doc.name}"`);
  }

  document.getElementById('comment-text').value = '';
  await loadComments();
  showToast('Comment added.');
}

async function resolveComment(id) {
  await sb.from('botec_comments').update({ resolved: true }).eq('id', id);
  await loadComments();
}

function toggleCommentsPanel() {
  document.getElementById('comments-panel').classList.toggle('hidden');
  loadComments();
}

// ── Save / Load ───────────────────────────────────────────────
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
  const { data, error } = await sb.from('botec_documents')
    .select('*').eq('id', id).single();
  if (error) { alert('Could not load document: ' + error.message); return; }

  isReviewer = data.reviewer_id === currentUser.id;

  document.getElementById('doc-title').value = data.name;
  document.title = `${data.name} — BOTEC`;
  deserialiseState(data.data);

  if (isReviewer) {
    document.getElementById('reviewer-mode-bar').classList.remove('hidden');
    document.getElementById('save-btn').classList.add('hidden');
    document.getElementById('reviewer-btn').classList.add('hidden');
    disableAllInputs();
    if (data.reviewer_status === 'requested') {
      await sb.from('botec_documents').update({ reviewer_status: 'in_review' }).eq('id', id);
    }
  } else {
    if (data.reviewer_id) {
      const { data: rProfile } = await sb.from('profiles').select('full_name,email').eq('id', data.reviewer_id).single();
      updateReviewerBadge(data.reviewer_status, rProfile);
    }
  }
  setSaveStatus('');
}

function updateReviewerBadge(status, profile) {
  const badge = document.getElementById('reviewer-badge');
  if (!badge) return;
  if (!profile) { badge.classList.add('hidden'); return; }
  badge.classList.remove('hidden');
  const labels = { none:'', requested:'Review requested', in_review:'In review', complete:'Review complete ✓' };
  badge.textContent = (profile.full_name || profile.email) + ' — ' + (labels[status] || status);
  badge.className = 'reviewer-badge status-' + status;
  const costsLabel = document.getElementById('costs-reviewer-label');
  if (costsLabel) costsLabel.textContent = profile ? `${profile.full_name || profile.email} assigned` : 'Assign reviewer';
}

function disableAllInputs() {
  document.querySelectorAll('input, select, textarea, .addbtn, .del').forEach(el => {
    el.disabled = true; el.style.opacity = '0.7';
  });
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
  ['projName','programme','prepBy','prepDate','reviewBy','reviewDate',
   'currency','numYears','bufferPct','mgrMult','purposeNote'].forEach(k => set2(k, s[k]));
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

const v   = id => { const el = document.getElementById(id); return el ? el.value : ''; };
const nv  = id => parseFloat(v(id)) || 0;
const getSym = () => ({ INR:'₹', USD:'$', GBP:'£', EUR:'€' })[v('currency')] || '₹';
const fC  = n => getSym() + Math.round(n).toLocaleString();
const fCd = (n, d=2) => getSym() + n.toFixed(d);
const escHtml = s => String(s).replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function showToast(msg) {
  const t = document.getElementById('toast');
  if (!t) return; t.textContent = msg; t.classList.remove('hidden');
  setTimeout(() => t.classList.add('hidden'), 3500);
}

// ── Log flags ─────────────────────────────────────────────────
function renderLogFlags() {
  const el = document.getElementById('logFlags');
  if (!el) return;
  el.innerHTML = HEADS.map(h => `
    <label style="display:flex;align-items:center;gap:5px;font-size:13px;cursor:pointer;padding:4px 10px;
      border:0.5px solid var(--border);border-radius:20px;
      background:${logSet.has(h)?'#e0f7fa':'var(--surface2)'}">
      <input type="checkbox" ${logSet.has(h)?'checked':''} onchange="toggleLog('${h}',this.checked)" style="width:auto;margin:0"> ${h}
    </label>`).join('');
}
function toggleLog(h, val) { val ? logSet.add(h) : logSet.delete(h); renderLogFlags(); calcAll(); }

// ── Tabs ──────────────────────────────────────────────────────
function go(id) {
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('on', ['checklist','setup','ben','costs','results'][i] === id));
  document.querySelectorAll('.pane').forEach(p => p.classList.remove('on'));
  document.getElementById('p-' + id).classList.add('on');
  if (id === 'results')   { calcAll(); renderResults(); }
  if (id === 'costs')     { calcAtta(); }
  if (id === 'checklist') { updateChecklistCount(); }
}

function updateChecklistCount() {
  const boxes = document.querySelectorAll('#p-checklist input[type=checkbox]');
  const checked = [...boxes].filter(b => b.checked).length;
  const el = document.getElementById('checklist-count');
  if (el) el.textContent = `${checked} of ${boxes.length} items checked`;
}

// Attach checklist counter to all checkboxes after DOM ready
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('#p-checklist input[type=checkbox]').forEach(cb => {
    cb.addEventListener('change', updateChecklistCount);
  });
});

// ── Beneficiaries ─────────────────────────────────────────────
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
  const NY     = parseInt(v('numYears')) || 3;
  const growth = parseFloat(v('benGrowth')) || 0;
  const fmt    = n => n > 0 ? Math.round(n).toLocaleString() : '—';
  const el     = (id) => document.getElementById(id);

  if (el('ben-show-1')) el('ben-show-1').textContent = fmt(b1);
  if (el('ben-show-2')) el('ben-show-2').textContent = NY >= 2 ? fmt(b2) : '—';
  if (el('ben-show-3')) el('ben-show-3').textContent = NY >= 3 ? fmt(b3) : '—';

  const g2 = el('ben-growth-2'), g3 = el('ben-growth-3');
  if (g2) g2.textContent = (benMode==='growth' && growth!==0 && b1>0) ? `${growth>0?'+':''}${growth}% vs Y1` : '';
  if (g3) g3.textContent = (benMode==='growth' && growth!==0 && b1>0) ? `${growth>0?'+':''}${(growth*2).toFixed(1)}% vs Y1` : '';

  const total = b1 + (NY>=2?b2:0) + (NY>=3?b3:0);
  if (el('ben-total')) el('ben-total').textContent = total > 0 ? Math.round(total).toLocaleString() : '—';

  calcAll();
  calcAtta();
}

function getBenTotals() {
  const y1 = parseFloat(v('benY1')) || 0;
  if (benMode === 'manual') return { b1:y1, b2:parseFloat(v('benY2manual'))||0, b3:parseFloat(v('benY3manual'))||0 };
  const g = 1 + (parseFloat(v('benGrowth'))||0) / 100;
  return { b1:y1, b2:y1*g, b3:y1*g*g };
}

// ── Atta consumption calculator ───────────────────────────────
// ── Atta programme presets ───────────────────────────────────
function applyAttaPreset() {
  const type = document.getElementById('atta-programme-type')?.value;
  const presets = {
    mdm:    { gpd:125, sdm:16,  ratio:5, label:'MDM: 125g · 16 days/month · 1:5 ratio' },
    twh:    { gpd:250, sdm:26,  ratio:5, label:'Tribal Welfare Hostel: 250g · 26 days/month · 1:5 ratio' },
    pds:    { gpd:417, sdm:30,  ratio:5, label:'PDS: 417g · 30 days/month · 1:5 ratio' },
    custom: { label:'Custom — edit fields as needed' }
  };
  const p = presets[type];
  if (!p) return;
  if (type !== 'custom') {
    const gpd   = document.getElementById('atta-gpd');
    const sdm   = document.getElementById('atta-sdm');
    const ratio = document.getElementById('atta-ratio');
    if (gpd)   gpd.value   = p.gpd;
    if (sdm)   sdm.value   = p.sdm;
    if (ratio) ratio.value = p.ratio;
  }
  const lbl = document.getElementById('atta-preset-label');
  if (lbl) lbl.textContent = p.label;
  calcAtta();
}

function calcAtta() {
  const { b1, b2, b3 } = getBenTotals();
  const NY = parseInt(v('numYears')) || 3;

  // Read editable assumptions — fall back to MDM defaults
  const gpd   = parseFloat(document.getElementById('atta-gpd')?.value)   || 125;
  const sdm   = parseFloat(document.getElementById('atta-sdm')?.value)   || 16;
  const ratio = parseFloat(document.getElementById('atta-ratio')?.value) || 5;

  const calcYear = ben => {
    const daily   = ben * gpd / 1e6;
    const monthly = daily * sdm;
    const yearMT  = monthly * 12;
    const yearKG  = yearMT * 1000;
    return { daily, monthly, yearMT, yearKG };
  };

  const a1 = calcYear(b1), a2 = calcYear(b2), a3 = calcYear(b3);
  lastAtta = { mt1:a1.yearMT, mt2:a2.yearMT, mt3:a3.yearMT, kg1:a1.yearKG, kg2:a2.yearKG, kg3:a3.yearKG };

  const fmtMT = n => n > 0 ? n.toFixed(2) + ' MT' : '—';
  const fmtKG = n => n > 0 ? Math.round(n).toLocaleString() + ' KG' : '—';

  const el = id => document.getElementById(id);
  if (el('atta-daily-1'))   el('atta-daily-1').textContent   = a1.daily > 0 ? a1.daily.toFixed(2) + ' MT/day' : '—';
  if (el('atta-monthly-1')) el('atta-monthly-1').textContent = fmtMT(a1.monthly);
  if (el('atta-yearly-1'))  el('atta-yearly-1').textContent  = fmtMT(a1.yearMT);
  if (el('atta-kg-1'))      el('atta-kg-1').textContent      = fmtKG(a1.yearKG);
  if (el('atta-daily-2'))   el('atta-daily-2').textContent   = NY>=2 ? (a2.daily.toFixed(2)+' MT/day') : '—';
  if (el('atta-monthly-2')) el('atta-monthly-2').textContent = NY>=2 ? fmtMT(a2.monthly) : '—';
  if (el('atta-yearly-2'))  el('atta-yearly-2').textContent  = NY>=2 ? fmtMT(a2.yearMT)  : '—';
  if (el('atta-kg-2'))      el('atta-kg-2').textContent      = NY>=2 ? fmtKG(a2.yearKG)  : '—';
  if (el('atta-daily-3'))   el('atta-daily-3').textContent   = NY>=3 ? (a3.daily.toFixed(2)+' MT/day') : '—';
  if (el('atta-monthly-3')) el('atta-monthly-3').textContent = NY>=3 ? fmtMT(a3.monthly) : '—';
  if (el('atta-yearly-3'))  el('atta-yearly-3').textContent  = NY>=3 ? fmtMT(a3.yearMT)  : '—';
  if (el('atta-kg-3'))      el('atta-kg-3').textContent      = NY>=3 ? fmtKG(a3.yearKG)  : '—';

  const ratioVal = parseFloat(document.getElementById('atta-ratio')?.value) || 5;
  if (el('atta-premix-1')) el('atta-premix-1').textContent = a1.yearMT>0?(a1.yearMT/ratioVal).toFixed(0)+' KG':'—';
  if (el('atta-premix-2')) el('atta-premix-2').textContent = a2.yearMT>0?(a2.yearMT/ratioVal).toFixed(0)+' KG':'—';
  if (el('atta-premix-3')) el('atta-premix-3').textContent = a3.yearMT>0?(a3.yearMT/ratioVal).toFixed(0)+' KG':'—';

  updateAttaRows();
}

function updateAttaRows() {
  document.querySelectorAll('#costBody tr[data-atta-type]').forEach(tr => {
    const type = tr.dataset.attaType;
    const id   = tr.id.replace('cr','');
    const u1   = document.getElementById('u1_'+id);
    const u2   = document.getElementById('u2_'+id);
    const u3   = document.getElementById('u3_'+id);
    let v1=0, v2=0, v3=0;
    const ratioForRow = parseFloat(document.getElementById('atta-ratio')?.value) || 5;
    if (type === 'premix') { v1=lastAtta.mt1/ratioForRow; v2=lastAtta.mt2/ratioForRow; v3=lastAtta.mt3/ratioForRow; }
    else                   { v1=lastAtta.kg1;   v2=lastAtta.kg2;   v3=lastAtta.kg3;   }
    if (u1) u1.value = v1 > 0 ? Math.round(v1) : '';
    if (u2) u2.value = v2 > 0 ? Math.round(v2) : '';
    if (u3) u3.value = v3 > 0 ? Math.round(v3) : '';
    updateCostRow(tr.querySelector('input[type=number]'));
  });
  calcAll();
}

// ── Cost rows ─────────────────────────────────────────────────
function addCost(d = {}) {
  const id      = ++cID;
  const tr      = document.createElement('tr');
  tr.id         = 'cr' + id;
  const headVal = d.h || d.head || HEADS[0];
  const isConsulting = d.consulting || headVal === 'Internal Consulting' && (d.consulting !== false);
  const isTravelRow  = d.travelRow;
  const attaType     = d.attaType || d.atta_type || null;
  const maeType      = d.maeType || null;

  if (attaType) tr.dataset.attaType = attaType;

  const sel  = HEADS.map(h => `<option value="${h}" ${h===headVal?'selected':''}>${h}</option>`).join('');
  const cpu  = d.cpu != null ? d.cpu : '';
  const u1   = (d.u1 != null && d.u1 !== '') ? d.u1 : (d.y1 !== '' && d.y1 != null ? d.y1 : '');
  const u2   = (d.u2 != null && d.u2 !== '') ? d.u2 : (d.y2 !== '' && d.y2 != null ? d.y2 : '');
  const u3   = (d.u3 != null && d.u3 !== '') ? d.u3 : (d.y3 !== '' && d.y3 != null ? d.y3 : '');
  const guide  = d.guide  || '';
  const justif = d.justif || d.justification || '';

  const unitReadonly = attaType ? 'readonly style="background:#e0f7fa;color:#006064;text-align:right"' : 'style="text-align:right"';

  // CHANGE: delete button uses position:sticky to stay visible on right edge
  tr.innerHTML = `
    <td><select onchange="onHeadChange(this,${id})">${sel}</select></td>
    <td>
      <input type="text" value="${escHtml(d.d||d.desc||'')}">
      ${guide ? `<div class="cost-guide">${guide}</div>` : ''}
    </td>
    <td><textarea rows="2" placeholder="Notes for reviewer…" style="font-size:12px;resize:vertical;min-height:34px">${escHtml(justif)}</textarea></td>
    <td><input type="text" value="${escHtml(d.u||d.unit||'')}" placeholder="e.g. KG, trip"></td>
    <td><input type="number" value="${cpu}" placeholder="Monthly cost" style="text-align:right" oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" id="u1_${id}" value="${u1}" placeholder="0" ${unitReadonly} oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" id="u2_${id}" value="${u2}" placeholder="0" ${unitReadonly} oninput="updateCostRow(this);calcAll()"></td>
    <td><input type="number" id="u3_${id}" value="${u3}" placeholder="0" ${unitReadonly} oninput="updateCostRow(this);calcAll()"></td>
    <td id="cy1_${id}" class="cost-computed">—</td>
    <td id="cy2_${id}" class="cost-computed">—</td>
    <td id="cy3_${id}" class="cost-computed">—</td>
    <td class="del-cell"><button class="del" onclick="deleteRow(${id})" title="Delete row">×</button></td>`;

  document.getElementById('costBody').appendChild(tr);

  if (isConsulting) addHoursHelper(id, d);
  else if (isTravelRow) addTravelHelper(id, d);
  else if (maeType) addMaeHelper(id, d, maeType);

  if (attaType) { setTimeout(() => { updateAttaRows(); updateCostRow(tr.querySelector('input[type=number]')); }, 0); }
  else { updateCostRow(tr.querySelector('input[type=number]')); }
  calcAll();
}

function addHoursHelper(id, d = {}) {
  const existing = document.getElementById('ch'+id);
  if (existing) existing.remove();
  const hr  = document.createElement('tr');
  hr.id     = 'ch' + id;
  hr.className = 'helper-row';
  const people = d.hPeople != null ? d.hPeople : '';
  const h1 = d.hHrs1 != null ? d.hHrs1 : '';
  const h2 = d.hHrs2 != null ? d.hHrs2 : '';
  const h3 = d.hHrs3 != null ? d.hHrs3 : '';
  hr.innerHTML = `<td colspan="12" class="helper-cell">
    <div class="helper-inner">
      <div class="helper-toggle" onclick="toggleHelper('hbody_${id}')">
        <span class="helper-label">Hours calculator</span>
        <span class="helper-chevron" id="hchev_${id}">▾</span>
      </div>
      <div id="hbody_${id}">
      <div class="helper-fixed">Fixed: 160 standard hrs/month &nbsp;·&nbsp; 12 months/year</div>
      <div class="helper-rows">
        <div class="helper-person-row">
          <span class="helper-field-lbl">Number of people</span>
          <input type="number" id="hPpl_${id}" value="${people}" placeholder="e.g. 2" style="width:70px" oninput="calcHoursRow(${id})">
          <span class="helper-fixed" style="margin-left:4px">(same across all years)</span>
        </div>
        <div class="helper-year-row">
          <span class="hyr-lbl">Year 1</span>
          <input type="number" id="hH1_${id}" value="${h1}" placeholder="hrs/month" style="width:90px" oninput="calcHoursRow(${id})">
          <span class="hyr-sep">hrs/month ÷ 160 × 12 × people</span>
          <span class="hyr-result" id="hR1_${id}">= — person-months</span>
        </div>
        <div class="helper-year-row">
          <span class="hyr-lbl">Year 2</span>
          <input type="number" id="hH2_${id}" value="${h2}" placeholder="hrs/month" style="width:90px" oninput="calcHoursRow(${id})">
          <span class="hyr-sep">hrs/month ÷ 160 × 12 × people</span>
          <span class="hyr-result" id="hR2_${id}">= — person-months</span>
        </div>
        <div class="helper-year-row">
          <span class="hyr-lbl">Year 3</span>
          <input type="number" id="hH3_${id}" value="${h3}" placeholder="hrs/month" style="width:90px" oninput="calcHoursRow(${id})">
          <span class="hyr-sep">hrs/month ÷ 160 × 12 × people</span>
          <span class="hyr-result" id="hR3_${id}">= — person-months</span>
        </div>
      </div>
      <div class="helper-hint">Units above are filled in automatically. You can still type directly to override.</div>
      </div>
      </div>
    </div></td>`;
  document.getElementById('cr'+id).insertAdjacentElement('afterend', hr);
  if (h1 || h2 || h3) calcHoursRow(id);
}

function calcHoursRow(id) {
  const ppl = parseFloat(document.getElementById('hPpl_'+id)?.value) || 0;
  const hs  = [1,2,3].map(y => parseFloat(document.getElementById(`hH${y}_${id}`)?.value) || 0);
  hs.forEach((h, i) => {
    const units = h > 0 ? parseFloat(((h/160)*12*ppl).toFixed(2)) : 0;
    const res   = document.getElementById(`hR${i+1}_${id}`);
    if (res) res.textContent = h > 0 ? `= ${units} person-months` : '= —';
    const uEl   = document.getElementById(`u${i+1}_${id}`);
    if (uEl) uEl.value = units || '';
  });
  const mainRow = document.getElementById('cr'+id);
  if (mainRow) updateCostRow(mainRow.querySelector('input[type=number]'));
  calcAll();
}

function addTravelHelper(id, d = {}) {
  const existing = document.getElementById('th'+id);
  if (existing) existing.remove();
  const tr2     = document.createElement('tr');
  tr2.id        = 'th' + id;
  tr2.className = 'helper-row';
  const people  = d.tPeople  != null ? d.tPeople  : '';
  const m1      = d.tM1      != null ? d.tM1      : '';
  const m2      = d.tM2      != null ? d.tM2      : '';
  const m3      = d.tM3      != null ? d.tM3      : '';
  const dpr     = d.tDayRate != null ? d.tDayRate : '';
  const td1     = d.tDays1   != null ? d.tDays1   : '';
  const td2     = d.tDays2   != null ? d.tDays2   : '';
  const td3     = d.tDays3   != null ? d.tDays3   : '';

  tr2.innerHTML = `<td colspan="12" class="helper-cell">
    <div class="helper-inner">
      <div class="helper-toggle" onclick="toggleHelper('tbody_${id}')">
        <span class="helper-label">Travel calculator</span>
        <span class="helper-chevron" id="tchev_${id}">▾</span>
      </div>
      <div id="tbody_${id}">
      <div class="helper-fixed" style="margin-bottom:8px">
        <strong>Person-month</strong> = the cost of one person travelling for one month.
        Use the <em>months calculator</em> if you think in months, or the <em>days converter</em>
        if you have a number of travel days — it will convert to person-months automatically (÷ 22 working days).
      </div>

      <div style="display:flex;gap:24px;flex-wrap:wrap">
        <!-- Left: months calculator (existing) -->
        <div style="flex:1;min-width:260px">
          <div class="helper-field-lbl" style="font-weight:600;margin-bottom:6px">Option A — enter months directly</div>
          <div class="helper-rows">
            <div class="helper-person-row">
              <span class="helper-field-lbl">Number of people travelling</span>
              <input type="number" id="tPpl_${id}" value="${people}" placeholder="e.g. 2" style="width:70px" oninput="calcTravelRow(${id})">
            </div>
            <div class="helper-year-row">
              <span class="hyr-lbl">Year 1</span>
              <input type="number" id="tM1_${id}" value="${m1}" placeholder="months" style="width:80px" oninput="calcTravelRow(${id})">
              <span class="hyr-sep">months × people</span>
              <span class="hyr-result" id="tR1_${id}">= — person-months</span>
            </div>
            <div class="helper-year-row">
              <span class="hyr-lbl">Year 2</span>
              <input type="number" id="tM2_${id}" value="${m2}" placeholder="months" style="width:80px" oninput="calcTravelRow(${id})">
              <span class="hyr-sep">months × people</span>
              <span class="hyr-result" id="tR2_${id}">= — person-months</span>
            </div>
            <div class="helper-year-row">
              <span class="hyr-lbl">Year 3</span>
              <input type="number" id="tM3_${id}" value="${m3}" placeholder="months" style="width:80px" oninput="calcTravelRow(${id})">
              <span class="hyr-sep">months × people</span>
              <span class="hyr-result" id="tR3_${id}">= — person-months</span>
            </div>
          </div>
        </div>

        <!-- Right: days converter -->
        <div style="flex:1;min-width:260px;border-left:0.5px solid var(--border);padding-left:20px">
          <div class="helper-field-lbl" style="font-weight:600;margin-bottom:6px">Option B — convert from days of travel</div>
          <div class="helper-rows">
            <div class="helper-person-row">
              <span class="helper-field-lbl">Daily rate (₹/day)</span>
              <input type="number" id="tDayRate_${id}" value="${dpr}" placeholder="e.g. 1136" style="width:90px" oninput="calcTravelDays(${id})">
              <span class="helper-fixed" style="margin-left:6px">÷ 22 working days = monthly rate</span>
            </div>
            <div class="helper-year-row">
              <span class="hyr-lbl">Year 1</span>
              <input type="number" id="tDays1_${id}" value="${td1}" placeholder="days" style="width:70px" oninput="calcTravelDays(${id})">
              <span class="hyr-sep">days × people ÷ 22</span>
              <span class="hyr-result" id="tDR1_${id}">= — person-months</span>
            </div>
            <div class="helper-year-row">
              <span class="hyr-lbl">Year 2</span>
              <input type="number" id="tDays2_${id}" value="${td2}" placeholder="days" style="width:70px" oninput="calcTravelDays(${id})">
              <span class="hyr-sep">days × people ÷ 22</span>
              <span class="hyr-result" id="tDR2_${id}">= — person-months</span>
            </div>
            <div class="helper-year-row">
              <span class="hyr-lbl">Year 3</span>
              <input type="number" id="tDays3_${id}" value="${td3}" placeholder="days" style="width:70px" oninput="calcTravelDays(${id})">
              <span class="hyr-sep">days × people ÷ 22</span>
              <span class="hyr-result" id="tDR3_${id}">= — person-months</span>
            </div>
          </div>
          <div class="helper-hint">Filling in Option B will update the monthly cost/unit field and the units above automatically.</div>
        </div>
      </div>
      </div>
    </div></td>`;
  document.getElementById('cr'+id).insertAdjacentElement('afterend', tr2);
  if (m1 || m2 || m3) calcTravelRow(id);
  if (td1 || td2 || td3) calcTravelDays(id);
}

function calcTravelRow(id) {
  const ppl = parseFloat(document.getElementById('tPpl_'+id)?.value) || 0;
  [1,2,3].forEach(y => {
    const m     = parseFloat(document.getElementById(`tM${y}_${id}`)?.value) || 0;
    const units = parseFloat((m * ppl).toFixed(2));
    const res   = document.getElementById(`tR${y}_${id}`);
    if (res) res.textContent = m > 0 ? `= ${units} person-months` : '= —';
    const uEl   = document.getElementById(`u${y}_${id}`);
    if (uEl) uEl.value = units || '';
  });
  const mainRow = document.getElementById('cr'+id);
  if (mainRow) updateCostRow(mainRow.querySelector('input[type=number]'));
  calcAll();
}

function calcTravelDays(id) {
  const dayRate = parseFloat(document.getElementById(`tDayRate_${id}`)?.value) || 0;
  const ppl     = parseFloat(document.getElementById(`tPpl_${id}`)?.value) || 1;
  // Update the monthly cost/unit field from daily rate
  if (dayRate > 0) {
    const monthlyRate = parseFloat((dayRate * 22).toFixed(0));
    const cpuEl = document.getElementById('cr'+id)?.querySelector('input[type=number]');
    if (cpuEl) { cpuEl.value = monthlyRate; }
  }
  [1,2,3].forEach(y => {
    const days  = parseFloat(document.getElementById(`tDays${y}_${id}`)?.value) || 0;
    const units = parseFloat(((days * ppl) / 22).toFixed(2));
    const res   = document.getElementById(`tDR${y}_${id}`);
    if (res) res.textContent = days > 0 ? `= ${units} person-months` : '= —';
    // Also fill the units fields
    const uEl = document.getElementById(`u${y}_${id}`);
    if (uEl && days > 0) uEl.value = units || '';
  });
  const mainRow = document.getElementById('cr'+id);
  if (mainRow) updateCostRow(mainRow.querySelector('input[type=number]'));
  calcAll();
}

function onHeadChange(sel, id) {
  const head = sel.value;
  ['ch','th','mh'].forEach(p => { const e=document.getElementById(p+id); if(e) e.remove(); });
  if (head === 'Internal Consulting') addHoursHelper(id, {});
  else if (head === 'Travel Costs') addTravelHelper(id, {});
  const tr = document.getElementById('cr'+id);
  delete tr.dataset.attaType;
  const attaInputs = tr.querySelectorAll('input[type=number][readonly]');
  attaInputs.forEach(i => { i.removeAttribute('readonly'); i.style.background=''; i.style.color=''; });
  calcAll();
}

function deleteRow(id) {
  ['cr','ch','th','mh'].forEach(p => { const e=document.getElementById(p+id); if(e) e.remove(); });
  calcAll();
}

function preset(k) { addCost(PRESETS[k]); }

function updateCostRow(inp) {
  if (!inp) return;
  const tr  = inp.closest('tr');
  if (!tr) return;
  const ins = tr.querySelectorAll('input[type=number]');
  const cpu = parseFloat(ins[0]?.value) || 0;
  const u1  = parseFloat(ins[1]?.value) || 0;
  const u2  = parseFloat(ins[2]?.value) || 0;
  const u3  = parseFloat(ins[3]?.value) || 0;
  const id  = tr.id.replace('cr','');
  ['cy1_','cy2_','cy3_'].forEach((p,i) => {
    const el = document.getElementById(p+id);
    if (el) el.textContent = fC(cpu * [u1,u2,u3][i]);
  });
}

function getCostData() {
  return Array.from(document.querySelectorAll('#costBody tr[id^="cr"]')).map(tr => {
    const id      = tr.id.replace('cr','');
    const sels    = tr.querySelectorAll('select');
    const txtIns  = tr.querySelectorAll('input[type=text]');
    const numIns  = tr.querySelectorAll('input[type=number]');
    const txts    = tr.querySelectorAll('textarea');
    const cpu = parseFloat(numIns[0]?.value)||0;
    const u1  = parseFloat(numIns[1]?.value)||0;
    const u2  = parseFloat(numIns[2]?.value)||0;
    const u3  = parseFloat(numIns[3]?.value)||0;
    const hPpl  = document.getElementById('hPpl_'+id)?.value;
    const hH1   = document.getElementById('hH1_'+id)?.value;
    const hH2   = document.getElementById('hH2_'+id)?.value;
    const hH3   = document.getElementById('hH3_'+id)?.value;
    const tPpl  = document.getElementById('tPpl_'+id)?.value;
    const tM1   = document.getElementById('tM1_'+id)?.value;
    const tM2   = document.getElementById('tM2_'+id)?.value;
    const tM3   = document.getElementById('tM3_'+id)?.value;
    return {
      head:   sels[0]?.value||'',
      desc:   txtIns[0]?.value||'',
      justif: txts[0]?.value||'',
      unit:   txtIns[1]?.value||'',
      cpu, u1, u2, u3, cy1:cpu*u1, cy2:cpu*u2, cy3:cpu*u3,
      attaType: tr.dataset.attaType || null,
      consulting: !!document.getElementById('hPpl_'+id),
      travelRow:  !!document.getElementById('tPpl_'+id),
      maeType: document.getElementById('mae_mills_'+id) ? (document.getElementById('mae_freq_'+id) ? 'nabl' : 'icheck') : null,
      ...(document.getElementById('mae_mills_'+id) && {
        maeMills:  document.getElementById('mae_mills_'+id)?.value,
        maeFreq:   document.getElementById('mae_freq_'+id)?.value,
        maeMonths: document.getElementById('mae_months_'+id)?.value
      }),
      ...(hPpl!=null && {hPeople:hPpl, hHrs1:hH1, hHrs2:hH2, hHrs3:hH3}),
      ...(tPpl!=null && {tPeople:tPpl, tM1, tM2, tM3,
        tDayRate: document.getElementById('tDayRate_'+id)?.value,
        tDays1:   document.getElementById('tDays1_'+id)?.value,
        tDays2:   document.getElementById('tDays2_'+id)?.value,
        tDays3:   document.getElementById('tDays3_'+id)?.value,
      }),
    };
  });
}

// ── Core calculations ─────────────────────────────────────────
function calcAll() {
  const rows = getCostData();
  const { b1, b2, b3 } = getBenTotals();
  const buf = nv('bufferPct') / 100;
  const mgr = nv('mgrMult')   / 100;
  const NY  = parseInt(v('numYears')) || 3;

  const byHead = {};
  HEADS.forEach(h => { byHead[h] = {y1:0,y2:0,y3:0}; });
  rows.forEach(r => {
    const k = byHead[r.head] !== undefined ? r.head : 'Other';
    byHead[k].y1 += r.cy1; byHead[k].y2 += r.cy2; byHead[k].y3 += r.cy3;
  });

  const cons   = byHead['Internal Consulting'];
  const mgrVal = { y1:cons.y1*mgr, y2:cons.y2*mgr, y3:cons.y3*mgr };

  const subY = {y1:0,y2:0,y3:0};
  HEADS.forEach(h => { subY.y1+=byHead[h].y1; subY.y2+=byHead[h].y2; subY.y3+=byHead[h].y3; });
  subY.y1+=mgrVal.y1; subY.y2+=mgrVal.y2; subY.y3+=mgrVal.y3;

  const bufVal = { y1:subY.y1*buf, y2:subY.y2*buf, y3:subY.y3*buf };
  const totY   = { y1:subY.y1+bufVal.y1, y2:subY.y2+bufVal.y2, y3:subY.y3+bufVal.y3 };

  const yrTotals = [totY.y1,totY.y2,totY.y3].slice(0,NY);
  const bens     = [b1,b2,b3].slice(0,NY);
  const totalAll = yrTotals.reduce((s,x)=>s+x,0);
  const totalBen = bens.reduce((s,x)=>s+x,0);
  const avgCost  = totalAll / NY;

  const cpbY   = [b1>0?totY.y1/b1:0, b2>0?totY.y2/b2:0, b3>0?totY.y3/b3:0];
  const cpbAvg = totalBen > 0 ? totalAll/totalBen : 0;

  const nonLogBase = {y1:0,y2:0,y3:0};
  HEADS.forEach(h => {
    if (!logSet.has(h)) { nonLogBase.y1+=byHead[h].y1; nonLogBase.y2+=byHead[h].y2; nonLogBase.y3+=byHead[h].y3; }
  });
  nonLogBase.y1+=mgrVal.y1; nonLogBase.y2+=mgrVal.y2; nonLogBase.y3+=mgrVal.y3;
  const cpbExclY   = [b1>0?nonLogBase.y1*(1+buf)/b1:0, b2>0?nonLogBase.y2*(1+buf)/b2:0, b3>0?nonLogBase.y3*(1+buf)/b3:0];
  const cpbExclAvg = totalBen>0?(nonLogBase.y1+nonLogBase.y2+nonLogBase.y3)*(1+buf)/totalBen:0;

  lastCalc = { rows,byHead,mgrVal,bufVal,subY,totY,totalAll,totalBen,avgCost,
               bens,yrTotals,cpbY,cpbAvg,cpbExclY,cpbExclAvg,buf,mgr,NY,b1,b2,b3 };
}

// ── Render results ────────────────────────────────────────────
function renderResults() {
  calcAll();
  const d  = lastCalc;
  const NY = d.NY||3;
  const abbr = n => {
    if (n>=1e7) return getSym()+(n/1e7).toFixed(2)+' Cr';
    if (n>=1e5) return getSym()+(n/1e5).toFixed(2)+' L';
    if (n>=1000) return getSym()+(n/1000).toFixed(1)+'K';
    return fCd(n);
  };

  document.getElementById('summCards').innerHTML = `
    <div class="card"><div class="lbl">CPB — avg</div><div class="val">${fCd(d.cpbAvg)}</div><div class="sub">total cost ÷ total beneficiaries</div></div>
    <div class="card"><div class="lbl">CPB excl. logistics</div><div class="val">${fCd(d.cpbExclAvg)}</div></div>
    <div class="card"><div class="lbl">Total ${NY}-yr cost</div><div class="val">${abbr(d.totalAll)}</div></div>
    <div class="card"><div class="lbl">Avg yearly cost</div><div class="val">${abbr(d.avgCost)}</div></div>
    <div class="card"><div class="lbl">Total beneficiaries</div><div class="val">${d.totalBen>=1e5?(d.totalBen/1000).toFixed(1)+'K':Math.round(d.totalBen).toLocaleString()}</div></div>`;

  const yrs = Array.from({length:NY},(_,i)=>`Year ${i+1}`);
  document.getElementById('summHead').innerHTML =
    `<th>Cost head</th>${yrs.map(y=>`<th style="text-align:right">${y}</th>`).join('')}<th style="text-align:right">Avg/year</th><th style="text-align:right">${NY}-yr total</th>`;

  const bd = document.getElementById('summBody');
  bd.innerHTML = '';
  const ah = HEADS.filter(h=>d.byHead[h]&&d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3>0);
  ah.forEach(h => {
    const vv=d.byHead[h], vals=[vv.y1,vv.y2,vv.y3].slice(0,NY), t=vals.reduce((s,x)=>s+x,0);
    bd.innerHTML += `<tr><td>${h}</td>${vals.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(t/NY)}</td><td style="text-align:right">${fC(t)}</td></tr>`;
  });

  const mv=[d.mgrVal.y1,d.mgrVal.y2,d.mgrVal.y3].slice(0,NY), mt=mv.reduce((s,x)=>s+x,0);
  if(mt>0) bd.innerHTML+=`<tr class="derived-row"><td>${(d.mgr*100).toFixed(0)}% managerial multiplier</td>${mv.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(mt/NY)}</td><td style="text-align:right">${fC(mt)}</td></tr>`;

  const bv=[d.bufVal.y1,d.bufVal.y2,d.bufVal.y3].slice(0,NY), bt=bv.reduce((s,x)=>s+x,0);
  if(bt>0) bd.innerHTML+=`<tr class="derived-row"><td>${(d.buf*100).toFixed(0)}% buffer</td>${bv.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(bt/NY)}</td><td style="text-align:right">${fC(bt)}</td></tr>`;

  bd.innerHTML+=`<tr class="grand-row"><td>Total costs</td>${d.yrTotals.map(x=>`<td style="text-align:right">${fC(x)}</td>`).join('')}<td style="text-align:right">${fC(d.avgCost)}</td><td style="text-align:right">${fC(d.totalAll)}</td></tr>`;
  bd.innerHTML+=`<tr class="derived-row"><td>Beneficiaries</td>${d.bens.map(x=>`<td style="text-align:right">${Math.round(x).toLocaleString()}</td>`).join('')}<td style="text-align:right">${Math.round(d.totalBen/NY).toLocaleString()}</td><td style="text-align:right">${Math.round(d.totalBen).toLocaleString()}</td></tr>`;
  bd.innerHTML+=`<tr class="cpb-row"><td>Cost per beneficiary</td>${d.cpbY.slice(0,NY).map(x=>`<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbAvg)}</td><td style="text-align:right">${fCd(d.cpbAvg)}</td></tr>`;
  if(logSet.size>0) bd.innerHTML+=`<tr class="cpb-row" style="font-style:italic"><td>CPB excl. ${[...logSet].join(' + ')}</td>${d.cpbExclY.slice(0,NY).map(x=>`<td style="text-align:right">${fCd(x)}</td>`).join('')}<td style="text-align:right">${fCd(d.cpbExclAvg)}</td><td style="text-align:right">${fCd(d.cpbExclAvg)}</td></tr>`;

  if(chart) chart.destroy();
  const colors=['#0097a7','#dc6059','#ff8dcb','#00bcd4','#f06292','#4dd0e1','#e57373','#80deea'];
  const datasets=ah.map((h,i)=>({label:h,data:[d.byHead[h].y1,d.byHead[h].y2,d.byHead[h].y3].slice(0,NY),backgroundColor:colors[i%colors.length],stack:'s'}));
  if(mt>0) datasets.push({label:'Mgr multiplier',data:mv,backgroundColor:'#CCC',stack:'s'});
  if(bt>0) datasets.push({label:'Buffer',data:bv,backgroundColor:'#E0E0E0',stack:'s'});
  chart=new Chart(document.getElementById('chartC').getContext('2d'),{
    type:'bar',data:{labels:yrs,datasets},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${fC(ctx.parsed.y)}`}}},
      scales:{x:{stacked:true},y:{stacked:true,ticks:{callback:val=>getSym()+(val>=1e7?(val/1e7).toFixed(1)+'Cr':val>=1e5?(val/1e5).toFixed(0)+'L':val>=1000?(val/1000).toFixed(0)+'K':val)}}}}
  });
}

// ── Excel export (only export format) ────────────────────────
function dlXLSX() {
  calcAll(); const d=lastCalc;
  const WB=XLSX.utils.book_new(), pn=v('projName')||'BOTEC', cur=v('currency'), NY=d.NY||3;
  const s=getSym();
  const n0 = n => Math.round(n);
  const n2 = n => parseFloat((n||0).toFixed(2));
  const yrs = Array.from({length:NY},(_,i)=>`Year ${i+1}`);
  const ah  = HEADS.filter(h=>d.byHead[h]&&d.byHead[h].y1+d.byHead[h].y2+d.byHead[h].y3>0);

  // ── SUMMARY SHEET ─────────────────────────────────────────────
  const sum = [];
  sum.push(['COST PER BENEFICIARY ESTIMATE']);
  sum.push([]);
  sum.push(['Project:', pn]);
  sum.push(['Prepared by:', v('prepBy')]);
  sum.push(['Preparation date:', v('prepDate')]);
  sum.push(['Reviewed by:', v('reviewBy')]);
  sum.push(['Review date:', v('reviewDate')]);
  sum.push(['Programme / state:', v('programme')]);
  sum.push([]);
  sum.push(['', ...yrs, 'Average / yr', `${NY}-yr Total`]);
  sum.push(['BENEFICIARIES', n0(d.b1), NY>=2?n0(d.b2):'', NY>=3?n0(d.b3):'', n0(d.totalBen/NY), n0(d.totalBen)]);
  sum.push([]);
  sum.push(['COSTS ('+cur+')', ...yrs, 'Average / yr', `${NY}-yr Total`]);

  ah.forEach(h => {
    const vv=d.byHead[h], vals=[vv.y1,vv.y2||0,vv.y3||0].slice(0,NY), t=vals.reduce((a,b)=>a+b,0);
    const row = [h, ...vals];
    while(row.length < 2+NY) row.push('');
    row.push(n0(t/NY), n0(t));
    sum.push(row);
  });

  const mt=d.mgrVal.y1+(d.mgrVal.y2||0)+(d.mgrVal.y3||0);
  if(mt>0) sum.push([`${(d.mgr*100).toFixed(0)}% Managerial multiplier`, n0(d.mgrVal.y1), n0(d.mgrVal.y2||0), n0(d.mgrVal.y3||0), n0(mt/NY), n0(mt)].slice(0, 2+NY+2));

  const bt=d.bufVal.y1+(d.bufVal.y2||0)+(d.bufVal.y3||0);
  if(bt>0) sum.push([`${(d.buf*100).toFixed(0)}% Buffer / contingency`, n0(d.bufVal.y1), n0(d.bufVal.y2||0), n0(d.bufVal.y3||0), n0(bt/NY), n0(bt)].slice(0, 2+NY+2));

  sum.push([]);
  sum.push(['TOTAL COSTS', n0(d.totY.y1), n0(d.totY.y2||0), n0(d.totY.y3||0), n0(d.avgCost), n0(d.totalAll)].slice(0, 2+NY+2));
  sum.push(['Beneficiaries', n0(d.b1), n0(d.b2||0), n0(d.b3||0), n0(d.totalBen/NY), n0(d.totalBen)].slice(0, 2+NY+2));
  sum.push(['COST PER BENEFICIARY (avg)', n2(d.cpbY[0]||0), n2(d.cpbY[1]||0), n2(d.cpbY[2]||0), n2(d.cpbAvg), n2(d.cpbAvg)].slice(0, 2+NY+2));
  if(logSet.size>0) {
    sum.push([`CPB excl. ${[...logSet].join(' + ')}`, n2(d.cpbExclY[0]||0), n2(d.cpbExclY[1]||0), n2(d.cpbExclY[2]||0), n2(d.cpbExclAvg), n2(d.cpbExclAvg)].slice(0, 2+NY+2));
  }

  const sumWS = XLSX.utils.aoa_to_sheet(sum);
  sumWS['!cols'] = [28,18,18,18,18,18].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(WB, sumWS, 'Summary');

  // ── COST CALCULATION SHEET ────────────────────────────────────
  const cc = [];
  cc.push(['COST CALCULATION']);
  cc.push([]);
  cc.push(['COST HEAD','DESCRIPTION','NOTES / JUSTIFICATION','UNIT',`MONTHLY COST/UNIT (${cur})`,'UNITS Y1','UNITS Y2','UNITS Y3',`COST Y1 (${cur})`,`COST Y2 (${cur})`,`COST Y3 (${cur})`]);
  d.rows.forEach(r => {
    cc.push([r.head, r.desc, r.justif||'', r.unit, r.cpu, r.u1||0, r.u2||0, r.u3||0, n0(r.cy1), n0(r.cy2), n0(r.cy3)]);
  });
  cc.push([]);
  cc.push(['TOTALS','','','','','','','', n0(d.totY.y1), n0(d.totY.y2||0), n0(d.totY.y3||0)]);

  const ccWS = XLSX.utils.aoa_to_sheet(cc);
  ccWS['!cols'] = [22,26,34,14,18,10,10,10,16,16,16].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(WB, ccWS, 'Cost Calculation');

  // ── BENEFICIARY SHEET ─────────────────────────────────────────
  const ben = [];
  ben.push(['BENEFICIARY CALCULATIONS']);
  ben.push([]);
  ben.push(['Year 1', n0(d.b1)]);
  if(NY>=2) ben.push(['Year 2', n0(d.b2||0)]);
  if(NY>=3) ben.push(['Year 3', n0(d.b3||0)]);
  ben.push([]);
  ben.push(['TOTAL', n0(d.totalBen)]);
  ben.push([]);
  ben.push(['Notes / source', v('benNotes')]);

  const benWS = XLSX.utils.aoa_to_sheet(ben);
  benWS['!cols'] = [22,22].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(WB, benWS, 'Beneficiary Calculation');

  // ── UNIT COSTS SHEET ──────────────────────────────────────────
  const uc = [];
  uc.push(['UNIT COSTS']);
  uc.push([]);
  uc.push(['COST HEAD','DESCRIPTION',`MONTHLY COST/UNIT (${cur})`,'UNIT']);
  const seen = new Set();
  d.rows.forEach(r => {
    const k = r.head+'|'+r.desc;
    if(!seen.has(k)){ seen.add(k); uc.push([r.head, r.desc, r.cpu, r.unit]); }
  });

  const ucWS = XLSX.utils.aoa_to_sheet(uc);
  ucWS['!cols'] = [24,30,22,16].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(WB, ucWS, 'Unit Costs');

  XLSX.writeFile(WB, `BOTEC_${pn.replace(/\s+/g,'_')}.xlsx`);
}
// ── M&E helpers ───────────────────────────────────────────────
function addMaeHelper(id, d = {}, type = 'nabl') {
  const existing = document.getElementById('mh'+id);
  if (existing) existing.remove();
  const mr = document.createElement('tr');
  mr.id = 'mh' + id;
  mr.className = 'helper-row';

  const mills  = d.maeMills  != null ? d.maeMills  : '';
  const freq   = d.maeFreq   != null ? d.maeFreq   : '';
  const months = d.maeMonths != null ? d.maeMonths : '';

  if (type === 'nabl') {
    mr.innerHTML = `<td colspan="12" class="helper-cell">
      <div class="helper-inner">
        <div class="helper-toggle" onclick="toggleHelper('maebody_${id}')">
          <span class="helper-label">NABL testing calculator</span>
          <span class="helper-chevron" id="maechev_${id}">▾</span>
        </div>
        <div id="maebody_${id}">
          <div class="helper-fixed">NABL testing is performed every 6 months. One mill per district = 5 mills total. Each mill gets 2 tests per year = 10 quality tests per year.</div>
          <div class="helper-rows">
            <div class="helper-year-row">
              <span class="helper-field-lbl" style="min-width:120px">Number of mills</span>
              <input type="number" id="mae_mills_${id}" value="${mills}" placeholder="e.g. 5" style="width:70px" oninput="calcMaeRow(${id},'nabl')">
              <span class="hyr-sep">mills ×</span>
              <input type="number" id="mae_freq_${id}" value="${freq}" placeholder="tests/year e.g. 2" style="width:120px" oninput="calcMaeRow(${id},'nabl')">
              <span class="hyr-sep">tests/year/mill =</span>
              <span class="hyr-result" id="mae_result_${id}">— tests/year</span>
            </div>
          </div>
          <div class="helper-hint">Enter the same number for all 3 years, or override units directly above.</div>
        </div>
      </div></td>`;
  } else {
    mr.innerHTML = `<td colspan="12" class="helper-cell">
      <div class="helper-inner">
        <div class="helper-toggle" onclick="toggleHelper('maebody_${id}')">
          <span class="helper-label">I-check testing calculator</span>
          <span class="helper-chevron" id="maechev_${id}">▾</span>
        </div>
        <div id="maebody_${id}">
          <div class="helper-fixed">I-check testing is performed monthly at each mill.</div>
          <div class="helper-rows">
            <div class="helper-year-row">
              <span class="helper-field-lbl" style="min-width:120px">Number of mills</span>
              <input type="number" id="mae_mills_${id}" value="${mills}" placeholder="e.g. 5" style="width:70px" oninput="calcMaeRow(${id},'icheck')">
              <span class="hyr-sep">mills ×</span>
              <input type="number" id="mae_months_${id}" value="${months}" placeholder="months/year e.g. 12" style="width:130px" oninput="calcMaeRow(${id},'icheck')">
              <span class="hyr-sep">months/year =</span>
              <span class="hyr-result" id="mae_result_${id}">— tests/year</span>
            </div>
          </div>
          <div class="helper-hint">Units above are filled in automatically. You can still type directly to override.</div>
        </div>
      </div></td>`;
  }
  document.getElementById('cr'+id).insertAdjacentElement('afterend', mr);
  if (mills || freq || months) calcMaeRow(id, type);
}

function calcMaeRow(id, type) {
  const mills  = parseFloat(document.getElementById('mae_mills_'+id)?.value)  || 0;
  const freq   = parseFloat(document.getElementById('mae_freq_'+id)?.value)   || 0;
  const months = parseFloat(document.getElementById('mae_months_'+id)?.value) || 0;
  const units  = type === 'nabl' ? mills * freq : mills * months;
  const res = document.getElementById('mae_result_'+id);
  if (res) res.textContent = units > 0 ? `${units} tests/year` : '— tests/year';
  ['u1_','u2_','u3_'].forEach(p => {
    const el = document.getElementById(p+id);
    if (el) el.value = units || '';
  });
  const mainRow = document.getElementById('cr'+id);
  if (mainRow) updateCostRow(mainRow.querySelector('input[type=number]'));
  calcAll();
}

function toggleHelper(bodyId) {
  const body = document.getElementById(bodyId);
  if (!body) return;
  const isHidden = body.style.display === 'none';
  body.style.display = isHidden ? '' : 'none';
  const chevId = bodyId.replace('body_','chev_').replace('hbody_','hchev_').replace('tbody_','tchev_');
  const chev = document.getElementById(chevId);
  if (chev) chev.textContent = isHidden ? '▾' : '▸';
}

document.addEventListener('input', () => { if(!isReviewer) setSaveStatus('Unsaved'); });
document.addEventListener('keydown', e => { if((e.ctrlKey||e.metaKey)&&e.key==='s'){e.preventDefault();if(!isReviewer)saveDocument();} });

init();

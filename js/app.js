/* WeRoCon Lab — Procurement & Inventory System
   Main app logic. Single source of truth (no layered "override scripts"
   like the old prototype had — that was the cause of the delete/restore
   bug). Everything reads/writes through DB (Supabase) so data is real,
   permanent, and shared across devices.
*/

const $ = (s) => document.querySelector(s);
let currentTab = 'dashboard';
let componentsCache = [];
let activeDocComponentIds = [];
let unsubscribeRealtime = null;

// ---------------- AUTH GATE ----------------
function renderGate() {
  $('#gate').innerHTML = `
    <div class="gate-box">
      <img src="assets/logo.png" alt="WeRoCon Lab logo" class="gate-logo">
      <h1>WeRoCon Lab</h1>
      <p class="muted">Procurement &amp; Inventory System</p>
      <input id="gate-email" type="email" placeholder="Email" autocomplete="username">
      <input id="gate-password" type="password" placeholder="Password" autocomplete="current-password">
      <div id="gate-error" class="gate-error"></div>
      <button class="btn primary" style="width:100%" id="gate-submit" onclick="tryGateUnlock()">Sign in</button>
      <p class="gate-note">Sign in with your lab account. Ask the lab admin if you need one created.</p>
    </div>`;
  const onEnter = (e) => { if (e.key === 'Enter') tryGateUnlock(); };
  $('#gate-email').addEventListener('keydown', onEnter);
  $('#gate-password').addEventListener('keydown', onEnter);
}

async function tryGateUnlock() {
  const email = $('#gate-email').value;
  const password = $('#gate-password').value;
  const btn = $('#gate-submit');
  const errBox = $('#gate-error');
  errBox.textContent = '';
  btn.disabled = true;
  btn.textContent = 'Signing in...';

  const result = await Auth.tryUnlock(email, password);

  btn.disabled = false;
  btn.textContent = 'Sign in';

  if (result.ok) {
    boot();
  } else {
    errBox.textContent = result.message || 'Sign in failed.';
  }
}

async function logout() {
  if (unsubscribeRealtime) unsubscribeRealtime();
  await Auth.lock();
  location.reload();
}

// ---------------- BOOT ----------------
async function boot() {
  $('#gate').classList.add('hidden');
  $('#shell').classList.remove('hidden');
  $('#role-label').textContent = Auth.email() || Auth.role();
  if (Auth.isViewOnly()) document.body.classList.add('view-only');

  await refreshComponents();
  showTab('dashboard');

  unsubscribeRealtime = DB.subscribeToChanges(async () => {
    await refreshComponents();
    showTab(currentTab, true);
  });
}

async function refreshComponents() {
  componentsCache = await DB.listComponents();
}

function fmtRs(n) { return '₹' + Number(n || 0).toLocaleString('en-IN'); }
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.style.display = 'block';
  clearTimeout(window._toastTimer);
  window._toastTimer = setTimeout(() => t.style.display = 'none', 2800);
}

// ---------------- TABS ----------------
function showTab(tab, silent = false) {
  currentTab = tab;
  document.querySelectorAll('.nav button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
  const renderers = { dashboard: renderDashboard, add: renderAddComponent, inventory: renderInventory, tracking: renderTracking, documents: renderDocuments };
  renderers[tab] && renderers[tab]();
}
document.addEventListener('click', (e) => {
  const btn = e.target.closest('.nav button');
  if (btn) showTab(btn.dataset.tab);
});

// ---------------- DASHBOARD ----------------
function pillForPayment(s) {
  if (s === 'Got the reimbursement' || (s || '').includes('Post-delivery') || s === 'Paid') return `<span class="pill green">${s}</span>`;
  if ((s || '').includes('Pending')) return `<span class="pill amber">${s}</span>`;
  return `<span class="pill blue">${s}</span>`;
}
function pillForGem(s) {
  if (s === 'Available on GeM') return `<span class="pill green">${s}</span>`;
  if (s === 'Non-GeM certified') return `<span class="pill blue">${s}</span>`;
  return `<span class="pill gray">${s}</span>`;
}
function pillForTracking(t) {
  if (!t) return `<span class="pill gray">No tracking</span>`;
  if (t.status === 'Delivered') return `<span class="pill green">Delivered</span>`;
  if (t.status === 'Out for delivery') return `<span class="pill blue">Out for delivery</span>`;
  if (t.status === 'In transit') return `<span class="pill amber">In transit</span>`;
  return `<span class="pill gray">${t.status}</span>`;
}

async function renderDashboard() {
  const comps = componentsCache;
  const total = comps.reduce((s, c) => s + c.qty * c.unit_price, 0);
  const paid = comps.filter(c => (c.payment_status || '').toLowerCase().includes('paid') || c.payment_status === 'Got the reimbursement').length;
  const trackingRows = await DB.listAllTracking();
  const trackMap = {};
  trackingRows.forEach(t => trackMap[t.component_id] = t);

  $('#app').innerHTML = `
    <section class="hero">
      <div><h1>Automatic Inventory System</h1><p>Enter a component once — fund approval, documents, shipment tracking and stock all stay linked to it.</p></div>
      <button class="btn dark" onclick="showTab('add')">+ Enter component</button>
    </section>
    <div class="cards">
      <div class="card metric"><b>${comps.length}</b><span>Components entered</span></div>
      <div class="card metric"><b>${fmtRs(total)}</b><span>Procurement value</span></div>
      <div class="card metric"><b>${paid}</b><span>Payments / reimbursements done</span></div>
      <div class="card metric"><b>${trackingRows.filter(t => t.status === 'Delivered').length}</b><span>Delivered to inventory</span></div>
    </div>
    <section class="section">
      <div class="head"><h2>Component procurement status</h2></div>
      ${comps.length === 0 ? `<div class="card empty">No components yet. Click "Enter component" to add your first one.</div>` :
        `<div class="steps">${comps.map((c, i) => `
          <div class="step">
            <div class="num">${i + 1}</div>
            <div>
              <h3>${c.name}</h3>
              <p>Qty ${c.qty} · ${c.vendor || 'No vendor'} · ${fmtRs(c.qty * c.unit_price)}</p>
            </div>
            <div style="display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end">
              ${pillForGem(c.gem_status)}
              ${pillForPayment(c.payment_status)}
              ${pillForTracking(trackMap[c.id])}
            </div>
          </div>`).join('')}</div>`}
    </section>
    ${renderReimbursementSection()}`;
}

// ---------------- REIMBURSEMENT ----------------
function reimbursementCandidates() {
  return componentsCache.filter(c => c.payment_mode === 'Paid by Saurav');
}

function renderReimbursementSection() {
  const list = reimbursementCandidates();
  if (list.length === 0) return '';

  return `
    <section class="section">
      <div class="head"><h2>Reimbursement status</h2></div>
      <div class="steps">
        ${list.map(c => {
          const status = c.payment_status;
          let body = '';

          if (status === 'Not applied yet') {
            body = `
              <span class="pill amber">Reminder: reimbursement not applied</span>
              <button class="btn" onclick="setReimbursementStatus(${c.id}, 'Applied for reimbursement')">Mark as applied</button>`;
          } else if (status === 'Applied for reimbursement') {
            body = `
              <span class="pill blue">Applied — awaiting confirmation</span>
              <span style="font-size:12px;color:var(--muted,#666)">Got the reimbursement?</span>
              <button class="btn primary" onclick="setReimbursementStatus(${c.id}, 'Got the reimbursement')">Yes</button>
              <button class="btn" onclick="setReimbursementStatus(${c.id}, 'Not applied yet')">Not yet</button>`;
          } else if (status === 'Got the reimbursement') {
            body = `<span class="pill green">Reimbursement received</span>`;
          } else {
            body = `<span class="pill gray">${status || '—'}</span>`;
          }

          return `
            <div class="step">
              <div class="num">₹</div>
              <div>
                <h3>${c.name}</h3>
                <p>${fmtRs(c.qty * c.unit_price)} · ${c.vendor || 'No vendor'}</p>
              </div>
              <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;justify-content:flex-end">${body}</div>
            </div>`;
        }).join('')}
      </div>
    </section>`;
}

async function setReimbursementStatus(id, newStatus) {
  await DB.updateComponent(id, { paymentStatus: newStatus });
  await refreshComponents();
  toast(newStatus === 'Got the reimbursement' ? 'Marked as reimbursed 🎉' : 'Status updated');
  showTab('dashboard');
}

// ---------------- ADD COMPONENT ----------------
function rdProjectNo(itemType) {
  return `IITJ/R&D/FS/${itemType}/2025-26/`;
}

function renderAddComponent() {
  $('#app').innerHTML = `
    <section class="hero">
      <div><h1>Enter component</h1><p>Add purchase details and upload the invoice. You'll generate the fund approval next.</p></div>
    </section>
    <div class="card">
      <div class="formgrid">
        <div class="field"><label>Name of PI</label><input id="f-pi" value="Dr. Saurav Kumar" readonly></div>
        <div class="field"><label>Project title</label><input id="f-project" value="Wearable Robotics &amp; Control Laboratory" readonly></div>
        <div class="field"><label>R&D project no.</label><input id="f-projectno" value="${rdProjectNo('Non-Recurring')}" readonly></div>

        <div class="field"><label>Item classification</label>
          <select id="f-type" onchange="onItemTypeChange()"><option value="Non-Recurring">Non Recurring (NR)</option><option value="Recurring">Recurring (R)</option></select>
        </div>
        <div class="field"><label>Component / item name *</label><input id="f-name" placeholder="e.g. BLDC motor"></div>
        <div class="field"><label>Category</label>
          <select id="f-category"><option>Motor</option><option>Bearing</option><option>Sensor</option><option>Cable</option><option>Drive / VFD</option><option>Other</option></select>
        </div>

        <div class="field"><label>Quantity *</label><input id="f-qty" type="number" min="1" value="1"></div>
        <div class="field"><label>Estimated unit cost (incl. all taxes) *</label><input id="f-price" type="number" min="0" placeholder="₹"></div>
        <div class="field"><label>Proposed vendor</label><input id="f-vendor"></div>

        <div class="field"><label>GeM status</label>
          <select id="f-gem"><option>Not checked</option><option>Available on GeM</option><option>Non-GeM certified</option></select>
        </div>
        <div class="field"><label>Payment mode</label>
          <select id="f-paymode" onchange="onPayModeChange()">
            <option value="Paid by Saurav">Paid by Saurav</option>
            <option value="By Purchase Order (P.O.)">By Purchase Order (P.O.)</option>
            <option value="By email confirmation">By email confirmation</option>
          </select>
        </div>
        <div class="field" id="f-paymethod-wrap"><label>Payment method</label>
          <select id="f-paymethod"><option>Cash</option><option>Credit card</option><option>Debit card</option><option>UPI</option><option>Bank transfer</option></select>
        </div>

        <div class="field"><label>Payment status</label>
          <select id="f-paystatus"></select>
        </div>

        <div class="field full"><label>Remarks / specifications</label><textarea id="f-remarks" placeholder="Technical specs, purpose, notes..."></textarea></div>
        <div class="field full"><label>Invoice file</label><input id="f-invoice-file" type="file" accept=".pdf,image/*"></div>
      </div>
      <div class="actions">
        <button class="btn" onclick="showTab('dashboard')">Cancel</button>
        <button class="btn primary" onclick="saveComponent()">Save component</button>
      </div>
    </div>`;
  onPayModeChange();
}

function onItemTypeChange() {
  $('#f-projectno').value = rdProjectNo($('#f-type').value);
}

const PAY_STATUS_OPTIONS = {
  'Paid by Saurav': ['Applied for reimbursement', 'Got the reimbursement', 'Not applied yet'],
  'By Purchase Order (P.O.)': ['Payment done by R&D', 'Not done'],
  'By email confirmation': ['Payment done by R&D', 'Not done']
};

function onPayModeChange() {
  const mode = $('#f-paymode').value;
  const methodWrap = $('#f-paymethod-wrap');
  methodWrap.style.display = (mode === 'Paid by Saurav') ? '' : 'none';

  const statusSel = $('#f-paystatus');
  statusSel.innerHTML = PAY_STATUS_OPTIONS[mode].map(s => `<option>${s}</option>`).join('');
}

async function saveComponent() {
  const required = ['f-name', 'f-price'];
  for (const id of required) {
    if (!$('#' + id).value.trim()) { toast('Please complete all required (*) fields'); return; }
  }
  const payMode = $('#f-paymode').value;
  const comp = await DB.addComponent({
    name: $('#f-name').value.trim(),
    category: $('#f-category').value,
    itemType: $('#f-type').value,
    qty: $('#f-qty').value,
    unitPrice: $('#f-price').value,
    vendor: $('#f-vendor').value.trim(),
    gemStatus: $('#f-gem').value,
    paymentMode: payMode,
    paymentMethod: payMode === 'Paid by Saurav' ? $('#f-paymethod').value : '',
    paymentStatus: $('#f-paystatus').value,
    remarks: $('#f-remarks').value.trim(),
    piName: $('#f-pi').value.trim(),
    projectTitle: $('#f-project').value.trim(),
    projectNo: $('#f-projectno').value.trim(),
    createdBy: Auth.email()
  });

  const invoiceFile = $('#f-invoice-file').files[0];
  if (invoiceFile) {
    await DB.uploadFile(comp.id, 'invoice', invoiceFile, Auth.email());
  }

  await refreshComponents();
  toast('Component saved');
  showTab('documents');
}

// ---------------- INVENTORY ----------------
async function renderInventory() {
  const viewOnly = Auth.isViewOnly();
  $('#app').innerHTML = `
    <div class="head">
      <h2>Inventory</h2>
      <div>${viewOnly ? '<span class="pill gray">View only</span>' : `<button class="btn primary" onclick="showTab('add')">+ Add component</button>`}</div>
    </div>
    <div class="notice">Data is stored in Supabase — shared live across every device, and deletions are permanent.</div>
    <div class="tablewrap section">
      <table>
        <thead><tr>
          <th>Component</th><th>Qty</th><th>Vendor</th><th>Unit price</th><th>Total</th>
          <th>GeM</th><th>Payment</th><th>Invoice</th><th>Remarks</th>${viewOnly ? '' : '<th></th>'}
        </tr></thead>
        <tbody>${componentsCache.map(c => `
          <tr>
            <td><b>${c.name}</b><br><small class="muted">${c.category}</small></td>
            <td>${c.qty}</td>
            <td>${c.vendor || '—'}</td>
            <td>${fmtRs(c.unit_price)}</td>
            <td><b>${fmtRs(c.qty * c.unit_price)}</b></td>
            <td>${pillForGem(c.gem_status)}</td>
            <td>${pillForPayment(c.payment_status)}</td>
            <td>${c.invoice_no || '<span class="muted">Pending</span>'}</td>
            <td class="muted" style="font-size:11px;max-width:160px">${c.remarks || '—'}</td>
            ${viewOnly ? '' : `<td><button class="btn danger" onclick="confirmDeleteComponent(${c.id}, '${c.name.replace(/'/g, "\\'")}')">Delete</button></td>`}
          </tr>`).join('') || `<tr><td colspan="9" class="empty">No components yet</td></tr>`}
        </tbody>
      </table>
    </div>`;
}

async function confirmDeleteComponent(id, name) {
  if (!confirm(`Permanently delete "${name}"? This removes its documents and tracking too. This cannot be undone.`)) return;
  await DB.deleteComponent(id);
  await refreshComponents();
  toast('Deleted permanently');
  showTab('inventory');
}

// ---------------- TRACKING ----------------
async function renderTracking() {
  const trackingRows = await DB.listAllTracking();
  const trackMap = {};
  trackingRows.forEach(t => trackMap[t.component_id] = t);
  const liveEnabled = !!APP_CONFIG.TRACKING_FUNCTION_URL;

  $('#app').innerHTML = `
    <section class="hero">
      <div><h1>Shipment tracking</h1><p>${liveEnabled ? 'Auto-synced with TrackCourier.io. Click "Sync now" to refresh live status.' : 'Manual tracking — deploy the Edge Function (see README) to enable auto-sync.'}</p></div>
    </section>
    <div class="tablewrap section">
      <table>
        <thead><tr><th>Component</th><th>Courier</th><th>Tracking ID</th><th>Status</th><th>Expected</th><th>Last synced</th><th></th></tr></thead>
        <tbody>${componentsCache.map(c => {
          const t = trackMap[c.id];
          return `<tr>
            <td><b>${c.name}</b></td>
            <td>${t?.courier || '—'}</td>
            <td>${t?.tracking_id || '—'}</td>
            <td>${pillForTracking(t)}</td>
            <td>${t?.expected_delivery || '—'}</td>
            <td class="muted" style="font-size:11px">${t?.last_synced_at ? new Date(t.last_synced_at).toLocaleString('en-IN') : '—'}</td>
            <td style="display:flex;gap:6px">
              <button class="btn" onclick="openTrackModal(${c.id})">${t ? 'Update' : 'Add tracking'}</button>
              ${liveEnabled && t?.tracking_id ? `<button class="btn" onclick="syncOne(${c.id}, '${t.tracking_id}', '${t.courier}')">Sync now</button>` : ''}
            </td>
          </tr>`;
        }).join('') || `<tr><td colspan="7" class="empty">No components yet</td></tr>`}
        </tbody>
      </table>
    </div>`;
}

function openTrackModal(componentId) {
  const comp = componentsCache.find(c => c.id === componentId);
  $('#modal-host').innerHTML = `
    <div class="modal open"><div class="modalbox">
      <h2>Update shipment — ${comp.name}</h2>
      <div class="formgrid">
        <div class="field"><label>Courier</label>
          <select id="t-courier">${Tracking.COURIERS.map(c => `<option>${c}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Tracking ID (AWB)</label><input id="t-id"></div>
        <div class="field"><label>Status</label>
          <select id="t-status">${Tracking.STEPS.map(s => `<option>${s}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Expected delivery</label><input type="date" id="t-expected"></div>
      </div>
      <div class="actions">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn primary" onclick="saveTracking(${componentId})">Save</button>
      </div>
    </div></div>`;
}
function closeModal() { $('#modal-host').innerHTML = ''; }

async function saveTracking(componentId) {
  const trackingId = $('#t-id').value.trim();
  if (!trackingId) { toast('Enter a tracking ID'); return; }
  await DB.upsertTracking(componentId, {
    courier: $('#t-courier').value,
    trackingId,
    status: $('#t-status').value,
    expected: $('#t-expected').value
  });
  closeModal();
  toast('Tracking saved');
  showTab('tracking');
}

async function syncOne(componentId, trackingId, courier) {
  toast('Syncing with TrackCourier.io...');
  const result = await DB.syncTrackingLive(componentId, trackingId, courier);
  if (result && result.current_status) {
    toast('Synced: ' + result.current_status);
  } else {
    const msg = (result && result.error) ? result.error : 'Sync failed — unknown error, check browser console';
    toast(msg);
    console.error('Sync failure detail:', result);
  }
  showTab('tracking');
}

// ---------------- DOCUMENTS ----------------
function renderDocuments() {
  if (componentsCache.length === 0) {
    $('#app').innerHTML = `<div class="card empty">Add a component first to generate documents.</div>`;
    return;
  }
  if (activeDocComponentIds.length === 0) activeDocComponentIds = componentsCache.map(c => c.id);

  $('#app').innerHTML = `
    <div class="head">
      <h2>Documents</h2>
      <button class="btn green" onclick="downloadBundleAll()">Download complete procurement pack (.zip)</button>
    </div>
    <div class="card">
      <div class="head"><h2 style="font-size:13px">Select components to include</h2></div>
      <div class="checkrow">${componentsCache.map(c => `
        <label><input type="checkbox" ${activeDocComponentIds.includes(c.id) ? 'checked' : ''} onchange="toggleDocComponent(${c.id}, this.checked)"> ${c.name}</label>`).join('')}
      </div>
    </div>
    <div class="tabs">
      <button class="btn primary" onclick="downloadSingle('fund-approval')">Fund Approval</button>
      <button class="btn" onclick="downloadSingle('quotation')">Quotation Comparison</button>
      <button class="btn" onclick="downloadSingle('non-gem')">Non-GeM Certificate</button>
      <button class="btn" onclick="downloadSingle('payment-receipt')">Payment Receipt</button>
    </div>
    <div class="notice">All four documents are generated client-side from your component data and download as .docx — editable in Word. The bundle button packages all applicable documents plus your uploaded quotation/invoice files into one .zip.</div>`;
}

function toggleDocComponent(id, checked) {
  if (checked) { if (!activeDocComponentIds.includes(id)) activeDocComponentIds.push(id); }
  else { activeDocComponentIds = activeDocComponentIds.filter(x => x !== id); }
}

function selectedComponents() {
  return componentsCache.filter(c => activeDocComponentIds.includes(c.id));
}

function metaFromSelection() {
  const sel = selectedComponents();
  const first = sel[0] || {};
  return {
    labName: APP_CONFIG.LAB_NAME,
    piName: first.pi_name || '',
    projectTitle: first.project_title || APP_CONFIG.PROJECT_TITLE_DEFAULT,
    projectNo: first.project_no || ''
  };
}

function toDocGenShape(c) {
  return {
    name: c.name, qty: c.qty, unitPrice: c.unit_price, vendor: c.vendor,
    invoiceNo: c.invoice_no, paymentStatus: c.payment_status,
    gemStatus: c.gem_status, gemSearchRef: c.gem_search_ref,
    itemType: c.item_type, remarks: c.remarks,
    quotationFiles: [], invoiceFile: null
  };
}

async function downloadSingle(type) {
  const sel = selectedComponents();
  if (sel.length === 0) { toast('Select at least one component'); return; }
  const meta = metaFromSelection();
  const stamp = new Date().toISOString().split('T')[0];
  const names = { 'fund-approval': 'Fund_Approval', quotation: 'Quotation_Comparison', 'non-gem': 'Non_GeM_Certificate', 'payment-receipt': 'Payment_Receipt' };
  await DocGen.downloadOne(type, meta, sel.map(toDocGenShape), `${names[type]}_${stamp}.docx`);
  await DB.logDocument(type, DocGen.refNo(type.slice(0, 2).toUpperCase()), activeDocComponentIds, Auth.email());
  toast('Downloaded');
}

async function downloadBundleAll() {
  const sel = selectedComponents();
  if (sel.length === 0) { toast('Select at least one component'); return; }
  const meta = metaFromSelection();
  await DocGen.downloadBundle(meta, sel.map(toDocGenShape));
  await DB.logDocument('bundle', 'BUNDLE', activeDocComponentIds, Auth.email());
  toast('Bundle downloaded');
}

// ---------------- INIT ----------------
window.addEventListener('DOMContentLoaded', async () => {
  await Auth.init();
  if (Auth.isUnlocked()) {
    boot();
  } else {
    renderGate();
  }
});

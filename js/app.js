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
let inventorySearchTerm = '';
let inventorySelectedIds = new Set();

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

// Items added via "Add existing item" (already purchased, fully settled) skip
// tracking, documents/billing, and the reimbursement pipeline entirely.
function isQuickAddItem(c) {
  return !!c.is_quick_add;
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
  // Quick-added (already purchased & settled) items are excluded from the
  // dashboard entirely — they only live in Inventory.
  const comps = componentsCache.filter(c => !isQuickAddItem(c));
  const total = comps.reduce((s, c) => s + c.qty * c.unit_price, 0);
  const paid = comps.filter(c => (c.payment_status || '').toLowerCase().includes('paid') || c.payment_status === 'Got the reimbursement').length;
  const trackingRows = (await DB.listAllTracking()).filter(t => comps.some(c => c.id === t.component_id));
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
  // Quick-added (already purchased & settled) items are excluded — they
  // never enter the reimbursement pipeline.
  return componentsCache.filter(c => c.payment_mode === 'Paid by Saurav' && !isQuickAddItem(c));
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
        <div class="field"><label>Invoice file</label><input id="f-invoice-file" type="file" accept=".pdf,image/*"></div>
        <div class="field"><label>Proof of payment (image only)</label><input id="f-payment-proof-file" type="file" accept="image/*"></div>
        <div class="field full"><label>Non-GeM certificate (if applicable)</label><input id="f-nongem-file" type="file" accept=".pdf,image/*"></div>
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
    createdBy: Auth.email(),
    isQuickAdd: false
  });

  const invoiceFile = $('#f-invoice-file').files[0];
  if (invoiceFile) {
    await DB.uploadFile(comp.id, 'invoice', invoiceFile, Auth.email());
    toast('Scanning invoice for bank details...');
    const extracted = await DB.extractInvoiceDetails(invoiceFile);
    if (extracted && !extracted.error) {
      const patch = {};
      if (extracted.invoiceNumber) patch.invoiceNo = extracted.invoiceNumber;
      if (extracted.invoiceDate) patch.invoiceDate = extracted.invoiceDate;
      if (extracted.vendorBankAccount) patch.vendorBankAccount = extracted.vendorBankAccount;
      if (extracted.vendorIfsc) patch.vendorIfsc = extracted.vendorIfsc;
      if (Object.keys(patch).length) {
        await DB.updateComponent(comp.id, patch);
        toast('Invoice scanned — bank details extracted');
      } else {
        toast('Invoice scanned — no bank details found, you can add them manually');
      }
    } else if (extracted && extracted.error) {
      console.warn('Invoice extraction failed:', extracted.error);
      toast('Could not auto-read the invoice — you can fill bank details manually later');
    }
  }
  const paymentProofFile = $('#f-payment-proof-file').files[0];
  if (paymentProofFile) {
    await DB.uploadFile(comp.id, 'payment_proof', paymentProofFile, Auth.email());
  }
  const nonGemFile = $('#f-nongem-file').files[0];
  if (nonGemFile) {
    await DB.uploadFile(comp.id, 'non_gem_certificate', nonGemFile, Auth.email());
  }

  await refreshComponents();
  toast('Component saved');
  showTab('documents');
}

// ---------------- INVENTORY ----------------
function filteredInventory() {
  const term = inventorySearchTerm.trim().toLowerCase();
  if (!term) return componentsCache;
  return componentsCache.filter(c =>
    (c.name || '').toLowerCase().includes(term) ||
    (c.vendor || '').toLowerCase().includes(term) ||
    (c.category || '').toLowerCase().includes(term) ||
    (c.remarks || '').toLowerCase().includes(term) ||
    (c.invoice_no || '').toLowerCase().includes(term)
  );
}

function onInventorySearch(value) {
  inventorySearchTerm = value;
  renderInventory();
}

async function renderInventory() {
  const viewOnly = Auth.isViewOnly();
  const rows = filteredInventory();
  const selectedCount = inventorySelectedIds.size;
  $('#app').innerHTML = `
    <div class="head">
      <h2>Inventory</h2>
      <div style="display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end">
        ${viewOnly ? '<span class="pill gray">View only</span>' : `
          <button class="btn" onclick="openQuickAddModal()">+ Add existing item (already purchased)</button>
          <button class="btn primary" onclick="showTab('add')">+ Add component</button>`}
      </div>
    </div>
    <div class="notice">Data is stored in Supabase — shared live across every device, and deletions are permanent. Use "Add existing item" for stock that's already been bought and fully settled — it skips fund approval &amp; document generation, tracking, and reimbursement, and lands straight in Inventory. Use "Edit" on any row to change details or remarks at any time.</div>
    <div class="section" style="margin-bottom:12px;display:flex;gap:10px;flex-wrap:wrap;align-items:center">
      <input id="inv-search" type="text" placeholder="Search by name, vendor, category, remarks or invoice no..." value="${inventorySearchTerm.replace(/"/g, '&quot;')}" oninput="onInventorySearch(this.value)" style="width:100%;max-width:420px">
      ${!viewOnly && selectedCount > 0 ? `
        <span class="pill blue">${selectedCount} selected</span>
        <button class="btn" onclick="bulkSetQuickAdd(true)">Mark as already purchased (hide from Dashboard/Tracking/Docs)</button>
        <button class="btn" onclick="bulkSetQuickAdd(false)">Mark as active procurement (show everywhere)</button>
        <button class="btn" onclick="clearInventorySelection()">Clear selection</button>` : ''}
    </div>
    <div class="tablewrap section">
      <table>
        <thead><tr>
          ${viewOnly ? '' : '<th></th>'}<th>Component</th><th>Qty</th><th>Vendor</th><th>Unit price</th><th>Total</th>
          <th>GeM</th><th>Payment</th><th>Invoice</th><th>Remarks</th><th>Status</th>${viewOnly ? '' : '<th></th>'}
        </tr></thead>
        <tbody>${rows.map(c => `
          <tr>
            ${viewOnly ? '' : `<td><input type="checkbox" ${inventorySelectedIds.has(c.id) ? 'checked' : ''} onchange="toggleInventorySelect(${c.id}, this.checked)"></td>`}
            <td><b>${c.name}</b><br><small class="muted">${c.category}</small></td>
            <td>${c.qty}</td>
            <td>${c.vendor || '—'}</td>
            <td>${fmtRs(c.unit_price)}</td>
            <td><b>${fmtRs(c.qty * c.unit_price)}</b></td>
            <td>${pillForGem(c.gem_status)}</td>
            <td>${pillForPayment(c.payment_status)}</td>
            <td>${c.invoice_no || '<span class="muted">Pending</span>'}</td>
            <td class="muted" style="font-size:11px;max-width:160px">${c.remarks || '—'}</td>
            <td>${isQuickAddItem(c) ? '<span class="pill gray">Already purchased</span>' : '<span class="pill blue">Active procurement</span>'}</td>
            ${viewOnly ? '' : `<td style="display:flex;gap:6px;white-space:nowrap"><button class="btn" onclick="openEditComponentModal(${c.id})">Edit</button><button class="btn danger" onclick="confirmDeleteComponent(${c.id}, '${c.name.replace(/'/g, "\\'")}')">Delete</button></td>`}
          </tr>`).join('') || `<tr><td colspan="10" class="empty">${inventorySearchTerm ? 'No components match your search' : 'No components yet'}</td></tr>`}
        </tbody>
      </table>
    </div>`;

  const searchInput = $('#inv-search');
  if (searchInput && inventorySearchTerm) {
    searchInput.focus();
    const pos = searchInput.value.length;
    searchInput.setSelectionRange(pos, pos);
  }
}

function toggleInventorySelect(id, checked) {
  if (checked) inventorySelectedIds.add(id);
  else inventorySelectedIds.delete(id);
  renderInventory();
}

function clearInventorySelection() {
  inventorySelectedIds.clear();
  renderInventory();
}

async function bulkSetQuickAdd(flag) {
  const ids = Array.from(inventorySelectedIds);
  if (ids.length === 0) return;
  for (const id of ids) {
    await DB.updateComponent(id, { isQuickAdd: flag });
  }
  inventorySelectedIds.clear();
  await refreshComponents();
  toast(flag ? 'Marked as already purchased' : 'Marked as active procurement');
  showTab('inventory');
}

async function confirmDeleteComponent(id, name) {
  if (!confirm(`Permanently delete "${name}"? This removes its documents and tracking too. This cannot be undone.`)) return;
  await DB.deleteComponent(id);
  await refreshComponents();
  toast('Deleted permanently');
  showTab('inventory');
}

// ---------------- EDIT COMPONENT (available any time, from Inventory) ----------------
function openEditComponentModal(id) {
  const c = componentsCache.find(x => x.id === id);
  if (!c) return;
  $('#modal-host').innerHTML = `
    <div class="modal open"><div class="modalbox">
      <h2>Edit — ${c.name}</h2>
      <div class="formgrid">
        <div class="field"><label>Component / item name *</label><input id="e-name" value="${(c.name || '').replace(/"/g, '&quot;')}"></div>
        <div class="field"><label>Category</label>
          <select id="e-category">${['Motor', 'Bearing', 'Sensor', 'Cable', 'Drive / VFD', 'Other'].map(o => `<option ${c.category === o ? 'selected' : ''}>${o}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Quantity *</label><input id="e-qty" type="number" min="1" value="${c.qty}"></div>
        <div class="field"><label>Unit price (incl. all taxes) *</label><input id="e-price" type="number" min="0" value="${c.unit_price}"></div>
        <div class="field"><label>Vendor</label><input id="e-vendor" value="${(c.vendor || '').replace(/"/g, '&quot;')}"></div>
        <div class="field"><label>GeM status</label>
          <select id="e-gem">${['Not checked', 'Available on GeM', 'Non-GeM certified'].map(o => `<option ${c.gem_status === o ? 'selected' : ''}>${o}</option>`).join('')}</select>
        </div>
        <div class="field"><label>Payment mode</label>
          <select id="e-paymode" onchange="onEditPayModeChange()">
            <option value="Paid by Saurav" ${c.payment_mode === 'Paid by Saurav' ? 'selected' : ''}>Paid by Saurav</option>
            <option value="By Purchase Order (P.O.)" ${c.payment_mode === 'By Purchase Order (P.O.)' ? 'selected' : ''}>By Purchase Order (P.O.)</option>
            <option value="By email confirmation" ${c.payment_mode === 'By email confirmation' ? 'selected' : ''}>By email confirmation</option>
          </select>
        </div>
        <div class="field"><label>Payment status</label><select id="e-paystatus"></select></div>
        <div class="field"><label>Invoice date</label><input id="e-invoice-date" value="${(c.invoice_date || '').replace(/"/g, '&quot;')}" placeholder="DD/MM/YYYY"></div>
        <div class="field"><label>Vendor bank A/c no.</label><input id="e-bank-acc" value="${(c.vendor_bank_account || '').replace(/"/g, '&quot;')}"></div>
        <div class="field"><label>Vendor IFSC</label><input id="e-bank-ifsc" value="${(c.vendor_ifsc || '').replace(/"/g, '&quot;')}"></div>
        <div class="field full"><label>Remarks / specifications</label><textarea id="e-remarks">${c.remarks || ''}</textarea></div>
      </div>
      <div class="actions">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn primary" onclick="saveEditComponent(${id})">Save changes</button>
      </div>
    </div></div>`;
  onEditPayModeChange(c.payment_status);
}

function onEditPayModeChange(preselect) {
  const mode = $('#e-paymode').value;
  const sel = $('#e-paystatus');
  sel.innerHTML = PAY_STATUS_OPTIONS[mode].map(s => `<option ${s === preselect ? 'selected' : ''}>${s}</option>`).join('');
}

async function saveEditComponent(id) {
  if (!$('#e-name').value.trim()) { toast('Name is required'); return; }
  await DB.updateComponent(id, {
    name: $('#e-name').value.trim(),
    category: $('#e-category').value,
    qty: $('#e-qty').value,
    unitPrice: $('#e-price').value,
    vendor: $('#e-vendor').value.trim(),
    gemStatus: $('#e-gem').value,
    paymentMode: $('#e-paymode').value,
    paymentStatus: $('#e-paystatus').value,
    invoiceDate: $('#e-invoice-date').value.trim(),
    vendorBankAccount: $('#e-bank-acc').value.trim(),
    vendorIfsc: $('#e-bank-ifsc').value.trim(),
    remarks: $('#e-remarks').value.trim()
  });
  closeModal();
  await refreshComponents();
  toast('Component updated');
  showTab('inventory');
}

// ---------------- QUICK ADD (already-purchased item — straight into inventory, no procurement flow) ----------------
function openQuickAddModal() {
  $('#modal-host').innerHTML = `
    <div class="modal open"><div class="modalbox">
      <h2>Add existing item to inventory</h2>
      <p class="muted" style="margin-top:-4px;font-size:12px">For stock that's already been bought and fully settled. This skips fund approval / documents, tracking, and reimbursement, and lands straight in Inventory.</p>
      <div class="formgrid">
        <div class="field"><label>Component / item name *</label><input id="q-name" placeholder="e.g. BLDC motor"></div>
        <div class="field"><label>Category</label>
          <select id="q-category"><option>Motor</option><option>Bearing</option><option>Sensor</option><option>Cable</option><option>Drive / VFD</option><option>Other</option></select>
        </div>
        <div class="field"><label>Quantity *</label><input id="q-qty" type="number" min="1" value="1"></div>
        <div class="field"><label>Unit price (incl. all taxes) *</label><input id="q-price" type="number" min="0" placeholder="₹"></div>
        <div class="field"><label>Vendor</label><input id="q-vendor"></div>
        <div class="field"><label>GeM status</label>
          <select id="q-gem"><option>Not checked</option><option>Available on GeM</option><option>Non-GeM certified</option></select>
        </div>
        <div class="field"><label>Payment mode</label>
          <select id="q-paymode" onchange="onQuickPayModeChange()">
            <option value="Paid by Saurav">Paid by Saurav</option>
            <option value="By Purchase Order (P.O.)">By Purchase Order (P.O.)</option>
            <option value="By email confirmation">By email confirmation</option>
          </select>
        </div>
        <div class="field"><label>Payment status</label><select id="q-paystatus"></select></div>
        <div class="field full"><label>Remarks / specifications</label><textarea id="q-remarks" placeholder="Editable anytime later from Inventory → Edit"></textarea></div>
      </div>
      <div class="actions">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn primary" onclick="saveQuickAdd()">Add to inventory</button>
      </div>
    </div></div>`;
  onQuickPayModeChange();
}

function onQuickPayModeChange() {
  const mode = $('#q-paymode').value;
  const sel = $('#q-paystatus');
  // default straight to the "done" end of whichever pipeline is selected,
  // since the item is already fully bought and settled
  const doneStatus = {
    'Paid by Saurav': 'Got the reimbursement',
    'By Purchase Order (P.O.)': 'Payment done by R&D',
    'By email confirmation': 'Payment done by R&D'
  }[mode];
  sel.innerHTML = PAY_STATUS_OPTIONS[mode].map(s => `<option ${s === doneStatus ? 'selected' : ''}>${s}</option>`).join('');
}

async function saveQuickAdd() {
  if (!$('#q-name').value.trim() || !$('#q-price').value) { toast('Please complete all required (*) fields'); return; }
  await DB.addComponent({
    name: $('#q-name').value.trim(),
    category: $('#q-category').value,
    itemType: 'Non-Recurring',
    qty: $('#q-qty').value,
    unitPrice: $('#q-price').value,
    vendor: $('#q-vendor').value.trim(),
    gemStatus: $('#q-gem').value,
    paymentMode: $('#q-paymode').value,
    paymentMethod: '',
    paymentStatus: $('#q-paystatus').value,
    remarks: $('#q-remarks').value.trim(),
    piName: 'Dr. Saurav Kumar',
    projectTitle: 'Wearable Robotics & Control Laboratory',
    projectNo: rdProjectNo('Non-Recurring'),
    createdBy: Auth.email(),
    isQuickAdd: true
  });
  closeModal();
  await refreshComponents();
  toast('Added directly to inventory');
  showTab('inventory');
}

// ---------------- TRACKING ----------------
async function renderTracking() {
  const trackingRows = await DB.listAllTracking();
  const trackMap = {};
  trackingRows.forEach(t => trackMap[t.component_id] = t);
  const liveEnabled = !!APP_CONFIG.TRACKING_FUNCTION_URL;

  // Quick-added (already purchased & settled) items never appear here.
  const trackableComponents = componentsCache.filter(c => !isQuickAddItem(c));

  $('#app').innerHTML = `
    <section class="hero">
      <div><h1>Shipment tracking</h1><p>${liveEnabled ? 'Auto-synced with TrackCourier.io. Click "Sync now" to refresh live status.' : 'Manual tracking — deploy the Edge Function (see README) to enable auto-sync.'}</p></div>
    </section>
    <div class="tablewrap section">
      <table>
        <thead><tr><th>Component</th><th>Courier</th><th>Tracking ID</th><th>Status</th><th>Expected</th><th>Last synced</th><th></th></tr></thead>
        <tbody>${trackableComponents.map(c => {
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
  // Quick-added (already purchased & settled) items are excluded from
  // billing/document generation — fund approval, quotation comparison,
  // non-GeM certificate and payment receipts don't apply to them.
  const documentableComponents = componentsCache.filter(c => !isQuickAddItem(c));

  if (documentableComponents.length === 0) {
    $('#app').innerHTML = `<div class="card empty">Add a component first to generate documents.</div>`;
    return;
  }
  if (activeDocComponentIds.length === 0 || !activeDocComponentIds.some(id => documentableComponents.some(c => c.id === id))) {
    activeDocComponentIds = documentableComponents.map(c => c.id);
  }

  $('#app').innerHTML = `
    <div class="head">
      <h2>Documents</h2>
      <button class="btn green" onclick="downloadBundleAll()">Download complete procurement pack (.zip)</button>
    </div>
    <div class="card">
      <div class="head"><h2 style="font-size:13px">Select components to include</h2></div>
      <div class="checkrow">${documentableComponents.map(c => `
        <label><input type="checkbox" ${activeDocComponentIds.includes(c.id) ? 'checked' : ''} onchange="toggleDocComponent(${c.id}, this.checked)"> ${c.name}</label>`).join('')}
      </div>
    </div>
    <div class="tabs">
      <button class="btn primary" onclick="downloadSingle('fund-approval')">Fund Approval</button>
      <button class="btn" onclick="downloadSingle('quotation')">Quotation Comparison</button>
      <button class="btn" onclick="downloadSingle('non-gem')">Non-GeM Certificate</button>
      <button class="btn" onclick="downloadSingle('payment-receipt')">Payment Receipt</button>
      <button class="btn" onclick="downloadSingle('proof-of-payment')">Proof of Payment (combined)</button>
      <button class="btn" onclick="openPaymentFormModal()">Payment Reimbursement Form</button>
    </div>
    <div class="notice">All documents are generated client-side from your component data and download as .docx — editable in Word. "Proof of Payment (combined)" lays every uploaded payment-proof image out on as few pages as possible. "Payment Reimbursement Form" auto-fills the vendor's bank account/IFSC from the invoice you uploaded (scanned automatically when you added the component) — you just fill in the fund approval reference and a couple of other details. The bundle button packages every applicable document plus your actual uploaded invoice, payment-proof and non-GeM files into one .zip. Items added via "Add existing item" (already purchased) don't appear here.</div>`;
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
    quotationFiles: [], invoiceFile: null, paymentProofFile: null, nonGemFile: null
  };
}

// Pulls each selected component's actually-uploaded invoice / payment-proof /
// non-GeM-certificate files out of Storage (as base64) so they can be
// embedded in generated docs or dropped straight into the bundle .zip.
async function enrichWithFiles(components) {
  const enriched = [];
  for (const c of components) {
    const files = await DB.listFiles(c.id);
    const invoiceRec = files.find(f => f.file_type === 'invoice');
    const proofRec = files.find(f => f.file_type === 'payment_proof');
    const nonGemRec = files.find(f => f.file_type === 'non_gem_certificate');

    const invoiceFile = invoiceRec ? { name: invoiceRec.file_name, dataUrl: await DB.downloadFileAsDataUrl(invoiceRec.storage_path) } : null;
    const paymentProofFile = proofRec ? { name: proofRec.file_name, dataUrl: await DB.downloadFileAsDataUrl(proofRec.storage_path) } : null;
    const nonGemFile = nonGemRec ? { name: nonGemRec.file_name, dataUrl: await DB.downloadFileAsDataUrl(nonGemRec.storage_path) } : null;

    enriched.push({ ...toDocGenShape(c), invoiceFile, paymentProofFile, nonGemFile });
  }
  return enriched;
}

async function downloadSingle(type) {
  const sel = selectedComponents();
  if (sel.length === 0) { toast('Select at least one component'); return; }
  toast('Preparing document...');
  const meta = metaFromSelection();
  const stamp = new Date().toISOString().split('T')[0];
  const names = {
    'fund-approval': 'Fund_Approval', quotation: 'Quotation_Comparison',
    'non-gem': 'Non_GeM_Certificate', 'payment-receipt': 'Payment_Receipt',
    'proof-of-payment': 'Proof_of_Payment'
  };
  const enriched = await enrichWithFiles(sel);
  await DocGen.downloadOne(type, meta, enriched, `${names[type]}_${stamp}.docx`);
  await DB.logDocument(type, DocGen.refNo(type.slice(0, 2).toUpperCase()), activeDocComponentIds, Auth.email());
  toast('Downloaded');
}

// ---- Payment Reimbursement Form: bank details come from the invoice scan,
// everything else (fund approval ref, procurement route, etc.) is asked here. ----
function openPaymentFormModal() {
  const sel = selectedComponents().filter(c => c.invoice_no);
  if (sel.length === 0) { toast('Select at least one component with an invoice'); return; }
  const first = sel[0];
  const missingBank = sel.filter(c => !c.vendor_bank_account || !c.vendor_ifsc);

  $('#modal-host').innerHTML = `
    <div class="modal open"><div class="modalbox">
      <h2>Payment Reimbursement Form — details</h2>
      <p class="muted" style="margin-top:-4px;font-size:12px">Bank account fields below are auto-filled from the scanned invoice where available. Check them, then fill in the rest.</p>
      ${missingBank.length ? `<div class="notice">No bank details could be read for: ${missingBank.map(c => c.name).join(', ')}. Enter them manually below, or fix via Inventory → Edit.</div>` : ''}
      <div class="formgrid">
        <div class="field"><label>Fund Approval Reference No.</label><input id="pf-ref" placeholder="e.g. IITJ/RD/2026/045"></div>
        <div class="field"><label>Fund Approval Date</label><input type="date" id="pf-date"></div>
        <div class="field"><label>Procurement route</label>
          <select id="pf-route"><option>GeM</option><option>Non-GeM</option></select>
        </div>
        <div class="field"><label>Budget Head</label>
          <select id="pf-budget"><option ${first.item_type !== 'Recurring' ? 'selected' : ''}>Non-Recurring</option><option ${first.item_type === 'Recurring' ? 'selected' : ''}>Recurring</option></select>
        </div>
        <div class="field"><label>Payee name (vendor)</label><input id="pf-payee" value="${(first.vendor || '').replace(/"/g, '&quot;')}"></div>
        <div class="field"><label>Bank A/c No.</label><input id="pf-acc" value="${(first.vendor_bank_account || '').replace(/"/g, '&quot;')}"></div>
        <div class="field"><label>IFSC</label><input id="pf-ifsc" value="${(first.vendor_ifsc || '').replace(/"/g, '&quot;')}"></div>
        <div class="field full"><label>Justification (optional — only needed for urgent procurement)</label><textarea id="pf-justification"></textarea></div>
      </div>
      <div class="actions">
        <button class="btn" onclick="closeModal()">Cancel</button>
        <button class="btn primary" onclick="generatePaymentForm()">Generate</button>
      </div>
    </div></div>`;
}

async function generatePaymentForm() {
  const sel = selectedComponents().filter(c => c.invoice_no);
  const meta = {
    ...metaFromSelection(),
    fundApprovalRef: $('#pf-ref').value.trim(),
    fundApprovalDate: $('#pf-date').value,
    procurementRoute: $('#pf-route').value,
    budgetHead: $('#pf-budget').value,
    payeeName: $('#pf-payee').value.trim(),
    bankAccountNo: $('#pf-acc').value.trim(),
    bankIfsc: $('#pf-ifsc').value.trim(),
    justification: $('#pf-justification').value.trim()
  };
  const bills = sel.map(c => ({
    invoiceNo: c.invoice_no,
    date: c.invoice_date || '',
    itemDetails: c.name,
    relevancy: 'Project requirement',
    amount: c.qty * c.unit_price,
    stockRegisterPage: ''
  }));
  closeModal();
  toast('Generating...');
  const stamp = new Date().toISOString().split('T')[0];
  await DocGen.downloadOne('payment-form', meta, bills, `Payment_Reimbursement_Form_${stamp}.docx`);
  await DB.logDocument('payment-form', DocGen.refNo('PF'), activeDocComponentIds, Auth.email());
  toast('Downloaded');
}

async function downloadBundleAll() {
  const sel = selectedComponents();
  if (sel.length === 0) { toast('Select at least one component'); return; }
  toast('Preparing bundle...');
  const meta = metaFromSelection();
  const enriched = await enrichWithFiles(sel);
  await DocGen.downloadBundle(meta, enriched);
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

/* Codex fix: make Inventory edit + save reliable. */
(() => {
  const STORAGE_KEY = 'werocon_inventory_items';
  let editingInventoryId = null;
  const q = (selectors, root = document) => { for (const selector of selectors) { const el = root.querySelector(selector); if (el) return el; } return null; };
  const norm = (value) => String(value ?? '').trim();
  const toNumber = (value) => { const n = Number(String(value ?? '').replace(/[^0-9.-]/g, '')); return Number.isFinite(n) ? n : 0; };
  const getInventory = () => {
    const candidates = [window.inventory, window.inventoryItems, window.Inventory, window.state?.inventory, window.appState?.inventory, window.AppState?.inventory, window.data?.inventory];
    const found = candidates.find(Array.isArray);
    if (found) return found;
    try { const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]'); if (Array.isArray(saved)) return saved; } catch {}
    return [];
  };
  const setInventory = (items) => {
    if (Array.isArray(window.inventory)) window.inventory = items;
    if (Array.isArray(window.inventoryItems)) window.inventoryItems = items;
    if (window.state && Array.isArray(window.state.inventory)) window.state.inventory = items;
    if (window.appState && Array.isArray(window.appState.inventory)) window.appState.inventory = items;
    if (window.AppState && Array.isArray(window.AppState.inventory)) window.AppState.inventory = items;
    if (window.data && Array.isArray(window.data.inventory)) window.data.inventory = items;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
  };
  const itemKeys = (item) => [item?.id, item?.itemId, item?.inventoryId, item?.sku, item?.code, item?.name, item?.itemName];
  const findItem = (id) => { const wanted = norm(id); return getInventory().find((item) => itemKeys(item).some((value) => norm(value) === wanted)); };
  const inventoryForm = () => q(['#inventoryForm', 'form[data-form="inventory"]', 'form[data-module="inventory"]', 'form[data-section="inventory"]', '#inventory form', '[data-page="inventory"] form']);
  const field = (form, names) => {
    for (const name of names) {
      const selectors = ['[name="' + name + '"]', '#' + name, '[data-field="' + name + '"]', '[name$=".' + name + '"]'];
      const el = q(selectors, form || document);
      if (el) return el;
    }
    return null;
  };
  const setField = (form, names, value) => { const el = field(form, names); if (!el) return; el.value = value ?? ''; el.dispatchEvent(new Event('input', { bubbles: true })); el.dispatchEvent(new Event('change', { bubbles: true })); };
  const readForm = (form) => {
    const fd = new FormData(form);
    const raw = Object.fromEntries(fd.entries());
    const read = (...names) => { for (const name of names) { if (raw[name] != null && raw[name] !== '') return raw[name]; const el = field(form, [name]); if (el && el.value !== '') return el.value; } return ''; };
    return {
      id: editingInventoryId || read('id', 'itemId', 'inventoryId', 'sku', 'code') || Date.now().toString(),
      name: read('name', 'itemName', 'item', 'description'),
      category: read('category', 'type'),
      quantity: toNumber(read('quantity', 'qty', 'stock', 'availableQty')),
      unit: read('unit', 'uom'),
      location: read('location', 'store', 'rack'),
      reorderLevel: toNumber(read('reorderLevel', 'reorder', 'minStock', 'minimumStock')),
      supplier: read('supplier', 'vendor', 'supplierName'),
      notes: read('notes', 'remarks', 'comments'),
    };
  };
  const fillForm = (item) => {
    const form = inventoryForm();
    if (!form || !item) return false;
    editingInventoryId = norm(item.id ?? item.itemId ?? item.inventoryId ?? item.sku ?? item.code ?? item.name ?? item.itemName);
    setField(form, ['id', 'itemId', 'inventoryId', 'sku', 'code'], editingInventoryId);
    setField(form, ['name', 'itemName', 'item', 'description'], item.name ?? item.itemName ?? item.description ?? '');
    setField(form, ['category', 'type'], item.category ?? item.type ?? '');
    setField(form, ['quantity', 'qty', 'stock', 'availableQty'], item.quantity ?? item.qty ?? item.stock ?? item.availableQty ?? '');
    setField(form, ['unit', 'uom'], item.unit ?? item.uom ?? '');
    setField(form, ['location', 'store', 'rack'], item.location ?? item.store ?? item.rack ?? '');
    setField(form, ['reorderLevel', 'reorder', 'minStock', 'minimumStock'], item.reorderLevel ?? item.reorder ?? item.minStock ?? item.minimumStock ?? '');
    setField(form, ['supplier', 'vendor', 'supplierName'], item.supplier ?? item.vendor ?? item.supplierName ?? '');
    setField(form, ['notes', 'remarks', 'comments'], item.notes ?? item.remarks ?? item.comments ?? '');
    form.dataset.editingId = editingInventoryId;
    form.scrollIntoView({ behavior: 'smooth', block: 'start' });
    return true;
  };
  const renderInventory = () => { for (const fn of ['renderInventory', 'renderInventoryTable', 'loadInventory', 'refreshInventory', 'showInventory']) { if (typeof window[fn] === 'function') { try { window[fn](); return; } catch {} } } };
  const saveItem = async (form) => {
    const next = readForm(form);
    const items = getInventory();
    const id = norm(next.id);
    const index = items.findIndex((item) => [item?.id, item?.itemId, item?.inventoryId, item?.sku, item?.code].some((value) => norm(value) === id));
    const previous = index >= 0 ? items[index] : {};
    const merged = { ...previous, ...next, id: previous.id ?? next.id };
    const updated = index >= 0 ? items.map((item, i) => (i === index ? merged : item)) : [...items, merged];
    setInventory(updated);
    if (window.db?.saveInventoryItem) await window.db.saveInventoryItem(merged);
    else if (window.DB?.saveInventoryItem) await window.DB.saveInventoryItem(merged);
    else if (window.saveInventoryItem) await window.saveInventoryItem(merged);
    else if (window.updateInventoryItem && index >= 0) await window.updateInventoryItem(merged);
    else if (window.addInventoryItem && index < 0) await window.addInventoryItem(merged);
    editingInventoryId = null;
    delete form.dataset.editingId;
    form.reset();
    renderInventory();
  };
  document.addEventListener('click', (event) => {
    const button = event.target.closest('button, a, [role="button"], [data-action]');
    if (!button) return;
    const label = norm(button.dataset.action || button.dataset.mode || button.textContent).toLowerCase();
    if (!/edit/.test(label)) return;
    const row = button.closest('tr, [data-inventory-id], [data-item-id], [data-id]');
    const id = button.dataset.id || button.dataset.inventoryId || button.dataset.itemId || row?.dataset.inventoryId || row?.dataset.itemId || row?.dataset.id || row?.children?.[0]?.textContent;
    const item = findItem(id);
    if (item && fillForm(item)) { event.preventDefault(); event.stopPropagation(); }
  }, true);
  document.addEventListener('submit', (event) => {
    const form = event.target;
    if (!(form instanceof HTMLFormElement)) return;
    const looksInventory = form.matches('#inventoryForm, form[data-form="inventory"], form[data-module="inventory"], form[data-section="inventory"]') || !!form.closest('#inventory, [data-page="inventory"]') || /inventory|stock|item/i.test(form.id + ' ' + form.className + ' ' + form.dataset.form + ' ' + form.dataset.module);
    if (!looksInventory) return;
    event.preventDefault();
    saveItem(form).catch((error) => { console.error('Inventory save failed:', error); alert('Inventory could not be saved. Please check the fields and try again.'); });
  }, true);
})();

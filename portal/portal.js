// ========================================
// CLIENT IDENTITY (read from session cookie)
// ========================================
function getClientId() {
  const cookie = document.cookie.split(';').find(c => c.trim().startsWith('portal_client='));
  return cookie ? cookie.split('=')[1] : null;
}

const CLIENT_ID = getClientId();

// Redirect to login if no valid session
if (!CLIENT_ID) {
  window.location.href = '/portal/';
}

function clientHeaders(extra = {}) {
  return { 'X-Client-Id': CLIENT_ID, ...extra };
}

function portalLogout() {
  document.cookie = 'portal_client=;path=/;max-age=0';
  window.location.href = '/portal/';
}


// ========================================
// CLIENT CONFIG & DEMO DATA
// ========================================

let CLIENT_CONFIG = {
  billingFrequency: 'monthly', // default fallback
};

async function loadClientConfig() {
  try {
    const res = await fetch('/api/client/config', { headers: clientHeaders() });
    const data = await res.json();
    CLIENT_CONFIG = { ...CLIENT_CONFIG, ...data };
  } catch (e) {
    console.warn('Could not load client config, using defaults');
  }
}

// TODO: In production, fetch from GET /api/qbo/accounts
let qboAccounts = [
  { id: 'acct-001', name: 'TD Business Chequing', type: 'bank', last4: '4521' },
  { id: 'acct-002', name: 'RBC Visa Business', type: 'credit_card', last4: '8834' },
  { id: 'acct-003', name: 'Scotiabank Savings', type: 'bank', last4: '1190' },
];

// Demo missing receipts (used in Step 3 after matching)
const DEMO_MISSING_RECEIPTS = [
  { id: 'txn-001', date: '2026-03-15', payee: 'Staples Business Depot', amount: 247.83, category: 'Office Supplies' },
  { id: 'txn-002', date: '2026-03-14', payee: 'Amazon Web Services', amount: 156.42, category: 'Software & Cloud' },
  { id: 'txn-003', date: '2026-03-12', payee: 'WestJet Airlines', amount: 489.00, category: 'Travel' },
  { id: 'txn-004', date: '2026-03-10', payee: 'Bell Canada', amount: 89.99, category: 'Telecommunications' },
  { id: 'txn-005', date: '2026-03-08', payee: 'Uber Eats — Client Meeting', amount: 62.15, category: 'Meals & Entertainment' },
  { id: 'txn-006', date: '2026-03-05', payee: 'Canva Pro', amount: 16.99, category: 'Software & Cloud' },
  { id: 'txn-007', date: '2026-03-03', payee: 'Petro-Canada', amount: 78.40, category: 'Vehicle Expenses' },
];


// ========================================
// WORKFLOW STATE (persisted to localStorage)
// ========================================
const STORAGE_KEY = `jackdoes-workflow-${CLIENT_ID}`;

let workflowState = {
  step1: { complete: false, uploads: {} },
  step2: { complete: false, uploads: {}, physicalCopies: {} },
  step3: { complete: false, resolved: {} },
};

function saveWorkflowState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(workflowState));
}

function loadWorkflowState() {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      workflowState = { ...workflowState, ...parsed };
    }
  } catch (e) { /* ignore */ }
}

loadWorkflowState();


// ========================================
// PERIOD CALCULATION
// ========================================
function getRequiredPeriods(billingFrequency) {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth(); // 0-indexed
  const periods = [];

  if (billingFrequency === 'monthly') {
    // Show all months from January through current month of current year
    for (let m = 0; m <= currentMonth; m++) {
      const d = new Date(currentYear, m, 1);
      periods.push({
        key: `${currentYear}-${String(m + 1).padStart(2, '0')}`,
        label: d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' }),
        startDate: new Date(currentYear, m, 1),
        endDate: new Date(currentYear, m + 1, 0),
      });
    }
  } else if (billingFrequency === 'quarterly') {
    // Show completed quarters of current year
    const quarters = [
      { months: [0, 1, 2], label: 'Q1' },
      { months: [1, 2, 3], label: 'Q2' },  // Placeholder — fixed below
      { months: [3, 4, 5], label: 'Q2' },
      { months: [6, 7, 8], label: 'Q3' },
      { months: [9, 10, 11], label: 'Q4' },
    ];
    const qDefs = [
      { start: 0, end: 2, label: 'Q1 (Jan–Mar)' },
      { start: 3, end: 5, label: 'Q2 (Apr–Jun)' },
      { start: 6, end: 8, label: 'Q3 (Jul–Sep)' },
      { start: 9, end: 11, label: 'Q4 (Oct–Dec)' },
    ];
    for (const q of qDefs) {
      if (currentMonth >= q.start) {
        periods.push({
          key: `${currentYear}-${q.label.substring(0, 2)}`,
          label: `${q.label} ${currentYear}`,
          startDate: new Date(currentYear, q.start, 1),
          endDate: new Date(currentYear, q.end + 1, 0),
        });
      }
    }
  } else if (billingFrequency === 'annual') {
    periods.push({
      key: `${currentYear}`,
      label: `${currentYear}`,
      startDate: new Date(currentYear, 0, 1),
      endDate: new Date(currentYear, 11, 31),
    });
  }

  return periods;
}

// Get periods, filtering out any already completed in prior workflow runs
function getOutstandingPeriods() {
  return getRequiredPeriods(CLIENT_CONFIG.billingFrequency);
}


// ========================================
// QUICKBOOKS CONNECTION STATUS
// ========================================
async function checkQBOStatus() {
  const banner = document.getElementById('qbo-banner');
  if (!banner) return;

  // Update the connect link to include this client's ID
  const connectBtn = banner.querySelector('.btn-qbo');
  if (connectBtn && CLIENT_ID) {
    connectBtn.href = `/api/qbo/connect?clientId=${CLIENT_ID}`;
  }

  try {
    const res = await fetch(`/api/qbo/status?clientId=${CLIENT_ID}`);
    const data = await res.json();
    if (data.connected) {
      banner.classList.add('connected');
      banner.querySelector('.qbo-banner-text').innerHTML =
        '<strong>quickbooks connected</strong> — jack has access to your financial data';
      if (connectBtn) {
        connectBtn.textContent = 'connected';
        connectBtn.classList.add('connected');
      }
    }
  } catch (e) { /* server not running */ }

  const params = new URLSearchParams(window.location.search);
  if (params.get('qbo') === 'connected') {
    banner.classList.add('connected');
    banner.querySelector('.qbo-banner-text').innerHTML =
      '<strong>quickbooks connected</strong> — jack has access to your financial data';
    if (connectBtn) {
      connectBtn.textContent = 'connected';
      connectBtn.classList.add('connected');
    }
    window.history.replaceState({}, '', window.location.pathname);
  }
}

checkQBOStatus();


// ========================================
// VIEW SWITCHING (with locked step gating)
// ========================================
const sidebarLinks = document.querySelectorAll('.sidebar-link[data-view]');
const views = document.querySelectorAll('.view');

sidebarLinks.forEach(link => {
  link.addEventListener('click', (e) => {
    e.preventDefault();

    // Block locked steps
    if (link.classList.contains('locked')) return;

    const viewId = link.dataset.view;

    sidebarLinks.forEach(l => l.classList.remove('active'));
    link.classList.add('active');

    views.forEach(v => v.classList.remove('active'));
    const target = document.getElementById(`view-${viewId}`);
    if (target) target.classList.add('active');

    // Trigger view-specific rendering
    if (viewId === 'step1') renderStep1();
    if (viewId === 'step2') renderStep2();
    if (viewId === 'step3') renderStep3();
    if (viewId === 'files') loadUploadedFiles();
    if (viewId === 'notifications') { loadNotifications(); markNotificationsRead(); }
  });
});


// ========================================
// WORKFLOW STEP MANAGEMENT
// ========================================
function updateWorkflowUI() {
  const step1Link = document.querySelector('[data-view="step1"]');
  const step2Link = document.querySelector('[data-view="step2"]');
  const step3Link = document.querySelector('[data-view="step3"]');

  // Step 1 is always accessible
  if (step1Link) step1Link.classList.remove('locked');

  // Update step indicators and lock states
  updateStepIndicator(1, workflowState.step1.complete);
  updateStepIndicator(2, workflowState.step2.complete);
  updateStepIndicator(3, workflowState.step3.complete);

  if (workflowState.step1.complete) {
    if (step2Link) step2Link.classList.remove('locked');
  } else {
    if (step2Link) step2Link.classList.add('locked');
  }

  if (workflowState.step2.complete) {
    if (step3Link) step3Link.classList.remove('locked');
  } else {
    if (step3Link) step3Link.classList.add('locked');
  }

  // Update progress bars in each step view
  document.querySelectorAll('.workflow-progress').forEach(bar => {
    const steps = bar.querySelectorAll('.wf-step');
    steps.forEach(s => {
      const num = parseInt(s.dataset.step);
      s.classList.remove('active', 'complete', 'locked');
      if (num === 1 && workflowState.step1.complete) s.classList.add('complete');
      else if (num === 2 && workflowState.step2.complete) s.classList.add('complete');
      else if (num === 3 && workflowState.step3.complete) s.classList.add('complete');
      else if (num === 1) s.classList.add('active');
      else if (num === 2 && workflowState.step1.complete) s.classList.add('active');
      else if (num === 3 && workflowState.step2.complete) s.classList.add('active');
      else s.classList.add('locked');
    });

    // Update connectors
    const connectors = bar.querySelectorAll('.wf-connector');
    if (connectors[0]) connectors[0].classList.toggle('done', workflowState.step1.complete);
    if (connectors[1]) connectors[1].classList.toggle('done', workflowState.step2.complete);
  });
}

function updateStepIndicator(stepNum, isComplete) {
  const link = document.querySelector(`[data-view="step${stepNum}"]`);
  if (!link) return;
  const indicator = link.querySelector('.step-indicator');
  if (!indicator) return;

  if (isComplete) {
    indicator.classList.add('complete');
    indicator.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>';
  } else {
    indicator.classList.remove('complete');
    indicator.textContent = stepNum;
  }
}

// Initialize: fetch config then render workflow and populate user info
(async () => {
  await loadClientConfig();
  updateWorkflowUI();

  // Populate sidebar user info
  const name = CLIENT_CONFIG.name || CLIENT_ID;
  const nameEl = document.getElementById('user-name');
  const emailEl = document.getElementById('user-email');
  const avatarEl = document.getElementById('user-avatar');
  if (nameEl) nameEl.textContent = name;
  if (emailEl) emailEl.textContent = CLIENT_CONFIG.email || '';
  if (avatarEl) {
    const initials = name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
    avatarEl.textContent = initials;
  }
})();


// ========================================
// CHAT FUNCTIONALITY (Ask Jack)
// ========================================
const chatMessages = document.getElementById('chat-messages');
const chatInput = document.getElementById('chat-input');
const chatSend = document.getElementById('chat-send');

if (chatInput) {
  chatInput.addEventListener('input', () => {
    chatInput.style.height = 'auto';
    chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
  });

  chatInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
}

if (chatSend) chatSend.addEventListener('click', sendMessage);

const sessionId = 'session_' + Math.random().toString(36).substring(2, 15);

async function sendMessage() {
  const message = chatInput.value.trim();
  if (!message) return;

  appendMessage('user', message);
  chatInput.value = '';
  chatInput.style.height = 'auto';

  const typingEl = appendTyping();

  try {
    const response = await fetch('/api/chat', {
      method: 'POST',
      headers: clientHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ message, sessionId }),
    });
    const data = await response.json();
    typingEl.remove();
    if (data.error) {
      appendMessage('jack', `sorry, something went wrong: ${data.error}`);
    } else {
      appendMessage('jack', data.response);
    }
  } catch (error) {
    typingEl.remove();
    appendMessage('jack', "sorry, i'm having trouble connecting right now. please try again in a moment.");
  }
}

function escapeHtml(text) {
  const el = document.createElement('span');
  el.textContent = text;
  return el.innerHTML;
}

function appendMessage(sender, text) {
  const div = document.createElement('div');
  div.className = `chat-message ${sender}`;
  if (sender === 'jack') {
    div.innerHTML = `<img src="../jack-avatar-clean.png" alt="jack" class="chat-avatar"><div class="chat-bubble">${text}</div>`;
  } else {
    div.innerHTML = `<div class="chat-bubble"><p>${escapeHtml(text)}</p></div>`;
  }
  chatMessages.appendChild(div);
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

function appendTyping() {
  const div = document.createElement('div');
  div.className = 'chat-message jack';
  div.innerHTML = `<img src="../jack-avatar-clean.png" alt="jack" class="chat-avatar"><div class="chat-bubble"><div class="chat-typing"><span></span><span></span><span></span></div></div>`;
  chatMessages.appendChild(div);
  chatMessages.scrollTop = chatMessages.scrollHeight;
  return div;
}


// ========================================
// REUSABLE DRAG-DROP HELPER
// ========================================
function initDropzone(dropzoneEl, fileInputEl, onFiles) {
  if (!dropzoneEl || !fileInputEl) return;

  dropzoneEl.addEventListener('click', (e) => {
    if (e.target.closest('button') || e.target.closest('label')) return;
    fileInputEl.click();
  });

  dropzoneEl.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzoneEl.classList.add('dragover');
  });

  dropzoneEl.addEventListener('dragleave', () => {
    dropzoneEl.classList.remove('dragover');
  });

  dropzoneEl.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzoneEl.classList.remove('dragover');
    onFiles(e.dataTransfer.files);
  });

  fileInputEl.addEventListener('change', () => {
    onFiles(fileInputEl.files);
    fileInputEl.value = '';
  });
}


// ========================================
// STEP 1: UPLOAD BANK & CREDIT CARD STATEMENTS
// ========================================
function renderStep1() {
  const container = document.getElementById('step1-accounts');
  if (!container) return;

  const periods = getOutstandingPeriods();

  let html = '';

  // Account cards
  qboAccounts.forEach(acct => {
    const typeBadge = acct.type === 'bank'
      ? '<span class="account-type-badge bank">bank</span>'
      : '<span class="account-type-badge cc">credit card</span>';

    html += `
      <div class="account-card" data-account-id="${acct.id}">
        <div class="account-card-header">
          <div class="account-card-info">
            <h3>${acct.name}</h3>
            <span class="account-last4">····${acct.last4}</span>
          </div>
          ${typeBadge}
        </div>
        <div class="period-slots">
          ${periods.map(p => {
            const uploadKey = `${acct.id}-${p.key}`;
            const isUploaded = workflowState.step1.uploads[uploadKey];
            return `
              <div class="period-slot ${isUploaded ? 'uploaded' : ''}">
                <div class="period-slot-info">
                  <span class="period-label">${p.label}</span>
                </div>
                ${isUploaded ? `
                  <span class="period-status uploaded">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                    uploaded
                  </span>
                ` : `
                  <label class="btn-period-upload">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
                    upload statement
                    <input type="file" accept=".pdf,.csv,.xlsx,.xls,.jpg,.jpeg,.png" capture="environment" hidden data-account-id="${acct.id}" data-period-key="${p.key}">
                  </label>
                `}
              </div>
            `;
          }).join('')}
        </div>
      </div>
    `;
  });

  // Add account button
  html += `
    <div class="account-card account-card-add" id="add-account-card">
      <button class="btn-add-account" id="btn-add-account">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
        add new bank or credit card account
      </button>
      <div class="add-account-form" id="add-account-form" style="display:none;">
        <div class="add-account-fields">
          <input type="text" id="new-account-name" placeholder="account name (e.g. TD Business Chequing)" class="form-input">
          <select id="new-account-type" class="form-input">
            <option value="bank">bank account</option>
            <option value="credit_card">credit card</option>
          </select>
          <input type="text" id="new-account-last4" placeholder="last 4 digits" maxlength="4" class="form-input">
        </div>
        <div class="add-account-actions">
          <button class="btn-add-confirm" id="btn-add-confirm">add account</button>
          <button class="btn-add-cancel" id="btn-add-cancel">cancel</button>
        </div>
      </div>
    </div>
  `;

  container.innerHTML = html;

  // Bind upload inputs
  container.querySelectorAll('input[type="file"][data-account-id]').forEach(input => {
    input.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleStep1Upload(e.target.dataset.accountId, e.target.dataset.periodKey, e.target.files[0]);
      }
    });
  });

  // Bind add account
  const addBtn = document.getElementById('btn-add-account');
  const addForm = document.getElementById('add-account-form');
  if (addBtn) {
    addBtn.addEventListener('click', () => {
      addBtn.style.display = 'none';
      addForm.style.display = 'block';
    });
  }

  const confirmBtn = document.getElementById('btn-add-confirm');
  const cancelBtn = document.getElementById('btn-add-cancel');

  if (confirmBtn) {
    confirmBtn.addEventListener('click', () => {
      const name = document.getElementById('new-account-name').value.trim();
      const type = document.getElementById('new-account-type').value;
      const last4 = document.getElementById('new-account-last4').value.trim();
      if (!name || !last4 || last4.length !== 4) return;

      const newAcct = {
        id: 'acct-' + Date.now(),
        name,
        type,
        last4,
      };
      qboAccounts.push(newAcct);
      // TODO: POST /api/qbo/accounts to create GL account in QuickBooks
      renderStep1();
    });
  }

  if (cancelBtn) {
    cancelBtn.addEventListener('click', () => {
      addForm.style.display = 'none';
      addBtn.style.display = 'flex';
    });
  }
}

function handleStep1Upload(accountId, periodKey, file) {
  const uploadKey = `${accountId}-${periodKey}`;
  workflowState.step1.uploads[uploadKey] = true;
  saveWorkflowState();
  checkStep1Complete();
  renderStep1();
  updateWorkflowUI();
}

function checkStep1Complete() {
  const periods = getOutstandingPeriods();
  const allDone = qboAccounts.every(acct =>
    periods.every(p => workflowState.step1.uploads[`${acct.id}-${p.key}`])
  );
  workflowState.step1.complete = allDone;
  saveWorkflowState();
}


// ========================================
// STEP 2: UPLOAD INVOICES & RECEIPTS
// ========================================
function renderStep2() {
  const container = document.getElementById('step2-periods');
  if (!container) return;

  const periods = getOutstandingPeriods();
  const categories = [
    { key: 'purchase-invoices', label: 'purchase invoices' },
    { key: 'purchase-receipts', label: 'purchase receipts' },
    { key: 'sales-invoices', label: 'sales invoices' },
  ];

  let html = '';

  // Physical copies banner
  html += `
    <div class="physical-copies-banner">
      <div class="physical-copies-banner-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="2" width="14" height="20" rx="2" ry="2"/><line x1="12" y1="18" x2="12.01" y2="18"/></svg>
      </div>
      <div class="physical-copies-banner-text">
        <h3>have physical copies?</h3>
        <p>if you have receipts or invoices in physical format, you can choose to mail them to us. select the "mail physical copies" option for any period below.</p>
      </div>
    </div>
  `;

  periods.forEach(p => {
    const hasPhysical = workflowState.step2.physicalCopies[p.key];
    const uploadsForPeriod = categories.some(c => workflowState.step2.uploads[`${p.key}-${c.key}`]);
    const periodDone = hasPhysical || uploadsForPeriod;

    html += `
      <div class="period-section ${periodDone ? 'period-done' : ''}">
        <div class="period-section-header">
          <h3>${p.label}</h3>
          ${periodDone ? '<span class="period-check"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg></span>' : ''}
        </div>
        <div class="category-upload-zones">
          ${categories.map(c => {
            const uploadKey = `${p.key}-${c.key}`;
            const isUploaded = workflowState.step2.uploads[uploadKey];
            return `
              <div class="category-upload-zone ${isUploaded ? 'zone-done' : ''}">
                <h4>${c.label}</h4>
                ${isUploaded ? `
                  <div class="zone-uploaded">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                    <span>uploaded</span>
                  </div>
                ` : `
                  <label class="zone-dropzone">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
                    <span>drag & drop or click</span>
                    <input type="file" accept=".pdf,.csv,.xlsx,.xls,.jpg,.jpeg,.png,.doc,.docx" capture="environment" multiple hidden data-period-key="${p.key}" data-category="${c.key}">
                  </label>
                `}
              </div>
            `;
          }).join('')}
        </div>
        <div class="period-physical-option">
          ${hasPhysical ? `
            <span class="physical-status active">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
              physical copies requested — see mailing instructions below
            </span>
          ` : `
            <button class="btn-physical-copies" data-period-key="${p.key}">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
              mail physical copies
            </button>
          `}
        </div>
      </div>
    `;
  });

  // Mailing instructions (shown if any period has physical copies)
  const anyPhysical = Object.values(workflowState.step2.physicalCopies).some(v => v);
  if (anyPhysical) {
    html += `
      <div class="mailing-instructions">
        <div class="mailing-instructions-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
        </div>
        <div>
          <h3>mailing instructions</h3>
          <p>please mail your physical documents to:</p>
          <address>
            <strong>jack does</strong><br>
            Greater Vancouver Area, BC<br>
            Canada
          </address>
          <p class="mailing-note">a prepaid postage label feature is coming soon. for now, please affix standard postage to your envelope.</p>
        </div>
      </div>
    `;
  }

  container.innerHTML = html;

  // Bind file uploads
  container.querySelectorAll('input[type="file"][data-period-key][data-category]').forEach(input => {
    input.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleStep2Upload(e.target.dataset.periodKey, e.target.dataset.category);
      }
    });
  });

  // Bind physical copies buttons
  container.querySelectorAll('.btn-physical-copies[data-period-key]').forEach(btn => {
    btn.addEventListener('click', () => {
      pendingPhysicalPeriod = btn.dataset.periodKey;
      document.getElementById('physical-modal').classList.add('active');
    });
  });
}

let pendingPhysicalPeriod = null;

function handleStep2Upload(periodKey, category) {
  const uploadKey = `${periodKey}-${category}`;
  workflowState.step2.uploads[uploadKey] = true;
  saveWorkflowState();
  checkStep2Complete();
  renderStep2();
  updateWorkflowUI();
}

function checkStep2Complete() {
  const periods = getOutstandingPeriods();
  const allDone = periods.every(p => {
    const hasUpload = ['purchase-invoices', 'purchase-receipts', 'sales-invoices']
      .some(c => workflowState.step2.uploads[`${p.key}-${c}`]);
    const hasPhysical = workflowState.step2.physicalCopies[p.key];
    return hasUpload || hasPhysical;
  });
  workflowState.step2.complete = allDone;
  saveWorkflowState();
}

// Physical copies modal handlers
const physicalModal = document.getElementById('physical-modal');
const physicalCancel = document.getElementById('physical-modal-cancel');
const physicalConfirm = document.getElementById('physical-modal-confirm');

if (physicalCancel) {
  physicalCancel.addEventListener('click', () => {
    pendingPhysicalPeriod = null;
    physicalModal.classList.remove('active');
  });
}

if (physicalConfirm) {
  physicalConfirm.addEventListener('click', () => {
    if (pendingPhysicalPeriod) {
      workflowState.step2.physicalCopies[pendingPhysicalPeriod] = true;
      saveWorkflowState();
      checkStep2Complete();
      renderStep2();
      updateWorkflowUI();
    }
    pendingPhysicalPeriod = null;
    physicalModal.classList.remove('active');
  });
}

if (physicalModal) {
  physicalModal.addEventListener('click', (e) => {
    if (e.target === physicalModal) {
      pendingPhysicalPeriod = null;
      physicalModal.classList.remove('active');
    }
  });
}


// ========================================
// STEP 3: REVIEW & MATCH (MISSING RECEIPTS)
// ========================================
let pendingNotAvailableId = null;

function renderStep3() {
  const container = document.getElementById('step3-missing-receipts');
  if (!container) return;

  // Initialize resolved state from workflowState
  const receipts = DEMO_MISSING_RECEIPTS.map(r => ({
    ...r,
    status: workflowState.step3.resolved[r.id] || 'missing',
  }));

  const pending = receipts.filter(r => r.status === 'missing');

  let html = '';

  // Mail physical copies button
  html += `
    <div class="step3-top-actions">
      <button class="btn-physical-copies step3-mail-btn" id="step3-mail-btn">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
        mail physical copies
      </button>
    </div>
  `;

  // Phone upload hint
  html += `
    <div class="phone-upload-banner">
      <div class="phone-upload-banner-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="2" width="14" height="20" rx="2" ry="2"/><line x1="12" y1="18" x2="12.01" y2="18"/></svg>
      </div>
      <div class="phone-upload-banner-text">
        <h3>snap a photo from your phone</h3>
        <p>add this portal to your phone's home screen for quick access. just take a photo of each receipt or invoice and upload it directly — no scanning needed.</p>
      </div>
    </div>
  `;

  // Info banner
  html += `
    <div class="missing-receipts-intro">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>
      <span>jack compared your bank transactions against the receipts and invoices you've uploaded. the items below are missing documentation. bank fees and interest charges have been excluded.</span>
    </div>
  `;

  if (pending.length === 0) {
    html += `
      <div class="missing-receipts-empty">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
        <p>all caught up!</p>
        <span>all bank transactions have been matched or resolved</span>
      </div>
    `;
  }

  receipts.forEach(txn => {
    const resolved = txn.status !== 'missing';
    html += `
      <div class="missing-receipt-card ${resolved ? 'resolved' : ''}" data-txn-id="${txn.id}">
        <div class="missing-receipt-icon">
          ${resolved
            ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'
            : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/><line x1="9" y1="15" x2="15" y2="15"/></svg>'
          }
        </div>
        <div class="missing-receipt-details">
          <div class="missing-receipt-payee">${txn.payee}</div>
          <div class="missing-receipt-meta">
            <span class="receipt-date">${formatReceiptDate(txn.date)}</span>
            <span class="receipt-category">${txn.category}</span>
          </div>
        </div>
        <div class="missing-receipt-amount">$${txn.amount.toFixed(2)}</div>
        <div class="missing-receipt-actions">
          ${txn.status === 'missing' ? `
            <label class="btn-upload-receipt">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              upload
              <input type="file" accept="image/*,.pdf" capture="environment" hidden data-txn-id="${txn.id}">
            </label>
            <button class="btn-not-available" data-txn-id="${txn.id}">not available</button>
          ` : txn.status === 'uploaded' ? `
            <span class="missing-receipt-status uploaded">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
              uploaded
            </span>
          ` : `
            <span class="missing-receipt-status not-available">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              not available
            </span>
          `}
        </div>
      </div>
    `;
  });

  container.innerHTML = html;

  // Bind receipt uploads
  container.querySelectorAll('input[type="file"][data-txn-id]').forEach(input => {
    input.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        workflowState.step3.resolved[e.target.dataset.txnId] = 'uploaded';
        saveWorkflowState();
        renderStep3();
      }
    });
  });

  // Bind "not available" buttons
  container.querySelectorAll('.btn-not-available[data-txn-id]').forEach(btn => {
    btn.addEventListener('click', () => {
      pendingNotAvailableId = btn.dataset.txnId;
      document.getElementById('cra-modal').classList.add('active');
    });
  });

  // Bind mail button
  const mailBtn = document.getElementById('step3-mail-btn');
  if (mailBtn) {
    mailBtn.addEventListener('click', () => {
      pendingPhysicalPeriod = '__step3__';
      document.getElementById('physical-modal').classList.add('active');
    });
  }
}

function formatReceiptDate(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

// CRA modal handlers
const craModal = document.getElementById('cra-modal');
const craCancel = document.getElementById('cra-modal-cancel');
const craConfirm = document.getElementById('cra-modal-confirm');

if (craCancel) {
  craCancel.addEventListener('click', () => {
    pendingNotAvailableId = null;
    craModal.classList.remove('active');
  });
}

if (craConfirm) {
  craConfirm.addEventListener('click', () => {
    if (pendingNotAvailableId) {
      workflowState.step3.resolved[pendingNotAvailableId] = 'not-available';
      saveWorkflowState();
      renderStep3();
    }
    pendingNotAvailableId = null;
    craModal.classList.remove('active');
  });
}

if (craModal) {
  craModal.addEventListener('click', (e) => {
    if (e.target === craModal) {
      pendingNotAvailableId = null;
      craModal.classList.remove('active');
    }
  });
}


// ========================================
// MY UPLOADED FILES
// ========================================
async function loadUploadedFiles() {
  const filesList = document.getElementById('files-list');
  if (!filesList) return;

  try {
    const res = await fetch('/api/files', { headers: clientHeaders() });
    const data = await res.json();

    if (!data.files || data.files.length === 0) {
      filesList.innerHTML = `
        <div class="files-empty">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" class="files-empty-icon"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/></svg>
          <p>no files uploaded yet</p>
          <span class="files-empty-hint">files uploaded through the workflow steps will appear here</span>
        </div>`;
      return;
    }

    filesList.innerHTML = `
      <table class="files-table">
        <thead><tr><th></th><th>file name</th><th>category</th><th>size</th><th>uploaded</th></tr></thead>
        <tbody>
          ${data.files.map(f => `
            <tr>
              <td><div class="file-icon-sm"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/></svg></div></td>
              <td class="file-name-cell">${f.name}</td>
              <td><span class="file-category-tag">${f.category}</span></td>
              <td class="file-size-cell">${formatFileSize(f.size)}</td>
              <td class="file-date-cell">${formatFileDate(f.uploadedAt)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>`;
  } catch (e) {
    filesList.innerHTML = `<div class="files-empty"><p>no files uploaded yet</p><span class="files-empty-hint">files uploaded through the workflow steps will appear here</span></div>`;
  }
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

function formatFileDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: 'numeric', minute: '2-digit' });
}


// ========================================
// NOTIFICATIONS
// ========================================
let cachedNotifications = [];

async function loadNotificationBadge() {
  try {
    const res = await fetch('/api/notifications', { headers: clientHeaders() });
    const data = await res.json();
    cachedNotifications = data.notifications || [];
    updateBadge(data.unreadCount || 0);
  } catch (e) { /* server not running */ }
}

function updateBadge(count) {
  const badge = document.getElementById('notification-badge');
  if (!badge) return;
  if (count > 0) {
    badge.textContent = count;
    badge.classList.add('visible');
  } else {
    badge.classList.remove('visible');
  }
}

async function loadNotifications() {
  const list = document.getElementById('notifications-list');
  if (!list) return;

  try {
    const res = await fetch('/api/notifications', { headers: clientHeaders() });
    const data = await res.json();
    cachedNotifications = data.notifications || [];
    updateBadge(data.unreadCount || 0);

    if (cachedNotifications.length === 0) {
      list.innerHTML = `
        <div class="notifications-empty">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>
          <p>no notifications</p>
          <span>when jack needs your attention on a document, it'll show up here</span>
        </div>`;
      return;
    }

    list.innerHTML = cachedNotifications.map(n => `
      <div class="notification-card ${n.read ? '' : 'unread'}" data-id="${n.id}">
        <div class="notification-card-header">
          <div class="notification-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          </div>
          <div class="notification-info">
            <div class="notification-filename">${n.fileName || 'unknown file'}</div>
            <div class="notification-meta">
              <span class="status-rejected">rejected</span>
              <span>${formatFileDate(n.reviewedAt)}</span>
            </div>
          </div>
        </div>
        ${n.rejectReason ? `<div class="notification-reason"><strong>reason:</strong> ${n.rejectReason}</div>` : `<div class="notification-reason"><strong>reason:</strong> no reason provided</div>`}
        ${n.summary ? `<div class="notification-summary">${n.summary}</div>` : ''}
      </div>
    `).join('');
  } catch (e) {
    list.innerHTML = `<div class="notifications-empty"><p>could not load notifications</p></div>`;
  }
}

async function markNotificationsRead() {
  const unread = cachedNotifications.filter(n => !n.read);
  if (unread.length === 0) return;
  try {
    await fetch('/api/notifications/read', {
      method: 'POST',
      headers: clientHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ ids: unread.map(n => n.id) }),
    });
    updateBadge(0);
    cachedNotifications.forEach(n => n.read = true);
    document.querySelectorAll('.notification-card.unread').forEach(el => el.classList.remove('unread'));
  } catch (e) { /* silently fail */ }
}

loadNotificationBadge();
setInterval(loadNotificationBadge, 30000);


// ========================================
// LOGIN FORM (placeholder)
// ========================================
const loginForm = document.getElementById('login-form');
if (loginForm) {
  loginForm.addEventListener('submit', (e) => {
    e.preventDefault();
    window.location.href = 'dashboard.html';
  });
}

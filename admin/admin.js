// ========================================
// AUTH
// ========================================
function getAuth() {
  const cookie = document.cookie.split(';').find(c => c.trim().startsWith('admin_auth='));
  return cookie ? cookie.split('=')[1] : '';
}

function logout() {
  document.cookie = 'admin_auth=;path=/;max-age=0';
  window.location.href = '/admin';
}

// ========================================
// STATE
// ========================================
let currentFilter = 'pending';
let allAnalyses = [];
let qboAccounts = []; // Chart of accounts from QuickBooks
let prepaidState = { prepaidAccount: null, items: [], amortizationRuns: [], scanThreshold: 500 };
let prepaidPreview = null; // cached preview for the current close period
let accruedLiabState = { accruedLiabilitiesAccount: null, analysisRuns: [], materialityThreshold: 10 };
let accruedLiabAnalysis = null; // latest analysis result for the panel
let shiState = { shareholderLoanAccount: null, invoices: [] };

/**
 * Show a brief "settings saved ✓" confirmation next to a button.
 */
function showSettingsSaved(btnId) {
  const btn = document.getElementById(btnId);
  if (!btn) return;
  const orig = btn.textContent;
  btn.textContent = '✓ saved';
  btn.style.background = 'var(--green-600)';
  btn.style.color = '#fff';
  setTimeout(() => {
    btn.textContent = orig;
    btn.style.background = '';
    btn.style.color = '';
  }, 2000);
}

/** Format a GL account label — shows "1234 - Account Name" if number exists, otherwise just name */
function fmtAcct(acct) {
  if (!acct) return '';
  const num = acct.acctNum || acct.AcctNum || '';
  const name = acct.name || acct.Name || '';
  return num ? `${num} - ${name}` : name;
}

/** Build an <option> tag for a GL account */
function acctOption(a, selectedId) {
  const label = fmtAcct(a);
  return `<option value="${a.id}" data-name="${(a.name || '').replace(/"/g, '&quot;')}" data-acctnum="${a.acctNum || ''}" data-type="${a.type || ''}" ${a.id === selectedId ? 'selected' : ''}>${label}</option>`;
}

// ========================================
// LOAD DATA
// ========================================

/**
 * Fetch the QBO chart of accounts for dropdown menus
 */
/**
 * Load QBO chart of accounts. If a clientId is provided, fetches accounts for that client's QBO connection.
 */
async function loadAccounts(clientId) {
  try {
    const url = clientId ? `/api/admin/accounts?clientId=${clientId}` : '/api/admin/accounts';
    const res = await fetch(url, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    qboAccounts = data.accounts || [];
    console.log(`Loaded ${qboAccounts.length} QBO accounts${clientId ? ` for client ${clientId}` : ''}`);
  } catch (e) {
    console.error('Failed to load QBO accounts:', e);
    qboAccounts = [];
  }
}

/**
 * Build an account <select> dropdown, grouped by account type
 * @param {string} selectedName - pre-selected account name (for matching Claude's suggestion)
 * @param {string} selectedId - pre-selected account ID
 * @param {number} entryIndex
 * @param {number} lineIndex
 */
function renderAccountSelect(selectedName, selectedId, entryIndex, lineIndex) {
  if (qboAccounts.length === 0) {
    // Fallback to text input if no accounts loaded
    return `<input type="text" class="edit-cell" value="${(selectedName || '').replace(/"/g, '&quot;')}" placeholder="account name" data-field="accountName" data-entry="${entryIndex}" data-line="${lineIndex}">`;
  }

  // Group accounts by type
  const groups = {};
  qboAccounts.forEach(a => {
    if (!groups[a.type]) groups[a.type] = [];
    groups[a.type].push(a);
  });

  // Try to find the best match for Claude's suggested account
  let matchedId = '';
  if (selectedId) {
    // Claude provided an accountId — use it directly if it exists in our list
    const byId = qboAccounts.find(a => a.id === String(selectedId));
    if (byId) matchedId = byId.id;
  }
  if (!matchedId && selectedName) {
    const nameLower = selectedName.toLowerCase();
    // Exact match first
    const exact = qboAccounts.find(a => a.name.toLowerCase() === nameLower);
    if (exact) {
      matchedId = exact.id;
    } else {
      // Fuzzy: check if account name contains the suggestion or vice versa
      const partial = qboAccounts.find(a =>
        a.name.toLowerCase().includes(nameLower) || nameLower.includes(a.name.toLowerCase())
      );
      if (partial) matchedId = partial.id;
    }
  }

  // Sort group keys for consistent ordering
  const typeOrder = ['Bank', 'Accounts Receivable', 'Other Current Asset', 'Fixed Asset', 'Other Asset',
    'Accounts Payable', 'Other Current Liability', 'Long Term Liability',
    'Equity', 'Income', 'Cost of Goods Sold', 'Expense', 'Other Income', 'Other Expense'];

  const sortedTypes = Object.keys(groups).sort((a, b) => {
    const ai = typeOrder.indexOf(a);
    const bi = typeOrder.indexOf(b);
    return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
  });

  let options = `<option value="">— select account —</option>`;
  sortedTypes.forEach(type => {
    options += `<optgroup label="${type}">`;
    groups[type].forEach(a => {
      const selected = a.id === matchedId ? ' selected' : '';
      options += `<option value="${a.id}" data-name="${a.name.replace(/"/g, '&quot;')}" data-acctnum="${a.acctNum || ''}"${selected}>${fmtAcct(a)}</option>`;
    });
    options += `</optgroup>`;
  });

  // If Claude suggested an account that doesn't match anything, add it as a custom option
  if (!matchedId && selectedName) {
    options = `<option value="" data-name="${selectedName.replace(/"/g, '&quot;')}" selected>⚠ ${selectedName} (not matched)</option>` + options.replace('<option value="">— select account —</option>', '');
  }

  return `<select class="edit-cell edit-account-select" data-field="accountName" data-entry="${entryIndex}" data-line="${lineIndex}">${options}</select>`;
}

async function loadStats() {
  try {
    const res = await fetch('/api/admin/stats', {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();

    document.getElementById('stat-pending').textContent = data.pending;
    document.getElementById('stat-approved').textContent = data.approved;
    document.getElementById('stat-rejected').textContent = data.rejected;

    // Top bar QBO status — show count of connected clients
    const qboEl = document.getElementById('qbo-status');
    const dot = qboEl.querySelector('.status-dot');
    const statusText = document.getElementById('qbo-status-text');
    const connections = data.qboConnections || {};
    const connCount = Object.keys(connections).length;

    if (connCount > 0) {
      dot.classList.remove('disconnected');
      dot.classList.add('connected');
      statusText.textContent = `qbo: ${connCount} client${connCount !== 1 ? 's' : ''} connected`;
    } else {
      dot.classList.remove('connected');
      dot.classList.add('disconnected');
      statusText.textContent = 'qbo: no clients connected';
    }
  } catch (e) {
    console.error('Failed to load stats:', e);
  }
}

async function loadAnalyses() {
  try {
    const res = await fetch('/api/admin/analyses', {
      headers: { 'Authorization': getAuth() },
    });

    if (res.status === 401) {
      window.location.href = '/admin';
      return;
    }

    const data = await res.json();
    allAnalyses = data.analyses || [];
    renderAnalyses();
  } catch (e) {
    console.error('Failed to load analyses:', e);
  }
}

function renderAnalyses() {
  const content = document.getElementById('admin-content');
  const filtered = currentFilter === 'all'
    ? allAnalyses
    : allAnalyses.filter(a => a.status === currentFilter);

  if (filtered.length === 0) {
    const label = currentFilter === 'all' ? '' : currentFilter + ' ';
    content.innerHTML = `
      <div class="empty-state">
        <p>no ${label}documents to review</p>
        <span>when clients upload documents, they'll appear here</span>
      </div>`;
    return;
  }

  content.innerHTML = filtered.map(a => renderAnalysisCard(a)).join('');
}

function renderAnalysisCard(a) {
  const analysis = a.analysis || {};
  const entries = analysis.entries || [];
  const statusClass = a.status === 'partial' ? 'partial' : a.status;

  return `
    <div class="analysis-card status-${statusClass}" id="card-${a.id}">
      <div class="analysis-header" onclick="toggleCard('${a.id}')">
        <div class="analysis-header-left">
          <div class="analysis-file-icon ${statusClass}">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/></svg>
          </div>
          <div class="analysis-info">
            <span class="analysis-filename">${a.fileName || 'unknown file'}</span>
            <div class="analysis-meta">
              <span class="analysis-client-name">${a.clientName || 'unknown client'}</span>
              <span>${a.category || 'general'}</span>
              <span>${formatDate(a.createdAt)}</span>
              ${analysis.totalAmount ? `<span>$${analysis.totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>` : ''}
            </div>
          </div>
        </div>
        <div class="analysis-header-right">
          <span class="analysis-status-badge ${statusClass}">${a.status}</span>
          <svg class="analysis-expand" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="18" height="18"><polyline points="6 9 12 15 18 9"/></svg>
        </div>
      </div>
      <div class="analysis-detail">
        <div class="analysis-summary">
          <div class="analysis-summary-row">
            ${analysis.documentType ? `<div class="analysis-summary-item"><span class="label">type</span><span class="value">${escapeHtml(analysis.documentType)}</span></div>` : ''}
            ${analysis.vendor ? `<div class="analysis-summary-item"><span class="label">vendor</span><span class="value">${escapeHtml(analysis.vendor)}</span></div>` : ''}
            ${analysis.customer ? `<div class="analysis-summary-item"><span class="label">customer</span><span class="value">${escapeHtml(analysis.customer)}</span></div>` : ''}
            ${analysis.date ? `<div class="analysis-summary-item"><span class="label">date</span><span class="value">${escapeHtml(analysis.date)}</span></div>` : ''}
            ${analysis.confidence ? `<div class="analysis-summary-item"><span class="label">confidence</span><span class="value">${escapeHtml(analysis.confidence)}</span></div>` : ''}
          </div>
          ${analysis.summary ? `<p>${escapeHtml(analysis.summary)}</p>` : ''}
        </div>

        <div id="entries-${a.id}">
        ${entries.map((entry, i) => `
          <div class="admin-entry-block" data-analysis-id="${a.id}" data-entry-index="${i}">
            <div class="admin-entry-header">
              <span class="admin-entry-type">${escapeHtml(entry.type || 'journal entry')}</span>
              <span class="admin-entry-date">${a.status === 'pending' ? `<input type="date" class="edit-date" value="${escapeHtml(entry.date || '')}" data-field="date" data-entry="${i}">` : escapeHtml(entry.date || '')}</span>
            </div>
            <div class="admin-entry-memo">${a.status === 'pending' ? `<input type="text" class="edit-memo" value="${escapeHtml(entry.memo || '')}" placeholder="memo / description" data-field="memo" data-entry="${i}">` : escapeHtml(entry.memo || '')}</div>
            <table class="admin-entry-table">
              <thead>
                <tr>
                  <th>account</th>
                  <th>description</th>
                  <th style="text-align:right">debit</th>
                  <th style="text-align:right">credit</th>
                  ${a.status === 'pending' ? '<th></th>' : ''}
                </tr>
              </thead>
              <tbody>
                ${entry.lines.map((line, li) => `
                  <tr data-line-index="${li}">
                    <td>${a.status === 'pending' ? renderAccountSelect(line.accountName, line.accountId, i, li) : escapeHtml(line.accountName)}</td>
                    <td>${a.status === 'pending' ? `<input type="text" class="edit-cell" value="${escapeHtml(line.description || '')}" data-field="description" data-entry="${i}" data-line="${li}">` : escapeHtml(line.description || '')}</td>
                    <td style="text-align:right">${a.status === 'pending' ? `<input type="number" step="0.01" class="edit-cell edit-amount" value="${line.type === 'debit' ? Math.abs(line.amount).toFixed(2) : ''}" placeholder="0.00" data-field="debit" data-entry="${i}" data-line="${li}">` : (line.type === 'debit' ? '$' + Math.abs(line.amount).toFixed(2) : '')}</td>
                    <td style="text-align:right">${a.status === 'pending' ? `<input type="number" step="0.01" class="edit-cell edit-amount" value="${line.type === 'credit' ? Math.abs(line.amount).toFixed(2) : ''}" placeholder="0.00" data-field="credit" data-entry="${i}" data-line="${li}">` : (line.type === 'credit' ? '$' + Math.abs(line.amount).toFixed(2) : '')}</td>
                    ${a.status === 'pending' ? `<td><button class="btn-remove-line" onclick="removeLine('${a.id}', ${i}, ${li})" title="remove line">&times;</button></td>` : ''}
                  </tr>
                `).join('')}
              </tbody>
            </table>
            ${a.status === 'pending' ? `
              <div class="entry-table-actions">
                <button class="btn-add-line" onclick="addLine('${a.id}', ${i})">
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  add line
                </button>
              </div>
            ` : ''}
          </div>
        `).join('')}
        </div>

        ${a.status === 'pending' ? `
          <button class="btn-add-entry" onclick="addEntry('${a.id}')">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
            add another entry
          </button>
        ` : ''}

        ${analysis.needsReview && analysis.needsReview.length > 0 ? `
          <div class="admin-review-notes">
            <strong>needs review:</strong>
            <ul>${analysis.needsReview.map(n => `<li>${escapeHtml(n)}</li>`).join('')}</ul>
          </div>
        ` : ''}

        ${analysis.notes ? `<div class="admin-notes">${escapeHtml(analysis.notes)}</div>` : ''}

        ${a.status === 'pending' ? `
          <div class="admin-actions" id="actions-${a.id}">
            <button class="btn-admin-approve" onclick="approveEntry('${a.id}')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><polyline points="20 6 9 17 4 12"/></svg>
              approve & post to quickbooks
            </button>
            <button class="btn-admin-reject" onclick="rejectEntry('${a.id}')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
              reject
            </button>
          </div>
        ` : `
          <div class="admin-reviewed-info">
            ${a.status === 'approved' ? 'approved' : a.status === 'rejected' ? 'rejected' : escapeHtml(a.status)}
            ${a.reviewedAt ? ` on ${formatDate(a.reviewedAt)}` : ''}
            ${a.rejectReason ? ` — reason: ${escapeHtml(a.rejectReason)}` : ''}
          </div>
        `}
      </div>
    </div>
  `;
}

// ========================================
// ACTIONS
// ========================================
function toggleCard(id) {
  const card = document.getElementById(`card-${id}`);
  card.classList.toggle('expanded');
}

/**
 * Read edited entry values from the inline form fields
 */
function getEditedEntries(id) {
  const entriesContainer = document.getElementById(`entries-${id}`);
  if (!entriesContainer) return null;

  const entryBlocks = entriesContainer.querySelectorAll('.admin-entry-block');
  const entries = [];

  entryBlocks.forEach(block => {
    const dateInput = block.querySelector('.edit-date');
    const memoInput = block.querySelector('.edit-memo');
    const typeEl = block.querySelector('.admin-entry-type');
    const rows = block.querySelectorAll('tbody tr');

    const lines = [];
    rows.forEach(row => {
      // Account: could be a <select> or <input>
      const accountSelect = row.querySelector('.edit-account-select');
      const accountInput = row.querySelector('[data-field="accountName"]:not(select)');
      let accountName = '';
      let accountId = '';

      if (accountSelect) {
        accountId = accountSelect.value;
        const selectedOption = accountSelect.options[accountSelect.selectedIndex];
        accountName = selectedOption?.dataset?.name || selectedOption?.textContent?.trim() || '';
      } else if (accountInput) {
        accountName = accountInput.value?.trim() || '';
      }

      const descInput = row.querySelector('[data-field="description"]');
      const description = descInput?.value?.trim() || '';
      const debitInput = row.querySelector('[data-field="debit"]');
      const creditInput = row.querySelector('[data-field="credit"]');
      const debitVal = parseFloat(debitInput?.value) || 0;
      const creditVal = parseFloat(creditInput?.value) || 0;

      if (!accountName && debitVal === 0 && creditVal === 0) return;

      if (debitVal > 0) {
        lines.push({ accountName, accountId, description, amount: debitVal, type: 'debit' });
      } else if (creditVal > 0) {
        lines.push({ accountName, accountId, description, amount: creditVal, type: 'credit' });
      }
    });

    if (lines.length > 0) {
      entries.push({
        type: typeEl?.textContent?.trim() || 'journal_entry',
        date: dateInput?.value || '',
        memo: memoInput?.value?.trim() || '',
        lines,
      });
    }
  });

  return entries;
}

async function approveEntry(id) {
  const actionsDiv = document.getElementById(`actions-${id}`);

  // Read edited values from the form
  const editedEntries = getEditedEntries(id);

  // Validate debits = credits for each entry
  if (editedEntries) {
    for (const entry of editedEntries) {
      const totalDebit = entry.lines.filter(l => l.type === 'debit').reduce((s, l) => s + l.amount, 0);
      const totalCredit = entry.lines.filter(l => l.type === 'credit').reduce((s, l) => s + l.amount, 0);
      if (Math.abs(totalDebit - totalCredit) > 0.01) {
        actionsDiv.innerHTML = `
          <span class="admin-action-result error">
            debits ($${totalDebit.toFixed(2)}) don't equal credits ($${totalCredit.toFixed(2)}) — please fix before approving
          </span>
          <button class="btn-admin-approve" onclick="approveEntry('${id}')" style="margin-left:10px;">retry</button>
          <button class="btn-admin-reject" onclick="rejectEntry('${id}')">reject</button>
        `;
        return;
      }
    }
  }

  actionsDiv.innerHTML = '<span class="admin-action-result">posting to quickbooks...</span>';

  try {
    const res = await fetch('/api/admin/approve', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': getAuth(),
      },
      body: JSON.stringify({ analysisId: id, entries: editedEntries }),
    });

    const data = await res.json();

    if (data.success) {
      const successCount = data.results.filter(r => r.success).length;
      const total = data.results.length;
      const failures = data.results.filter(r => !r.success);
      let resultHTML = '';
      if (successCount === total) {
        resultHTML = `
          <span class="admin-action-result success">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><polyline points="20 6 9 17 4 12"/></svg>
            ${successCount} of ${total} entries posted to quickbooks
          </span>`;
      } else {
        resultHTML = `
          <span class="admin-action-result error">
            ${successCount} of ${total} entries posted. Errors:<br>
            ${failures.map(f => `• ${f.error}`).join('<br>')}
          </span>
          <button class="btn-admin-approve" onclick="approveEntry('${id}')" style="margin-top:8px;">retry</button>`;
      }
      actionsDiv.innerHTML = resultHTML;

      const card = document.getElementById(`card-${id}`);
      card.classList.remove('status-pending');
      card.classList.add('status-approved');
      card.querySelector('.analysis-status-badge').textContent = 'approved';
      card.querySelector('.analysis-status-badge').className = 'analysis-status-badge approved';
      card.querySelector('.analysis-file-icon').className = 'analysis-file-icon approved';

      loadStats();
    } else {
      actionsDiv.innerHTML = `<span class="admin-action-result error">error: ${data.error}</span>`;
    }
  } catch (e) {
    actionsDiv.innerHTML = '<span class="admin-action-result error">failed to connect</span>';
  }
}

async function rejectEntry(id) {
  const reason = prompt('Reason for rejection (optional):') || '';
  const actionsDiv = document.getElementById(`actions-${id}`);

  try {
    await fetch('/api/admin/reject', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': getAuth(),
      },
      body: JSON.stringify({ analysisId: id, reason }),
    });

    const card = document.getElementById(`card-${id}`);
    card.classList.remove('status-pending');
    card.classList.add('status-rejected');
    card.querySelector('.analysis-status-badge').textContent = 'rejected';
    card.querySelector('.analysis-status-badge').className = 'analysis-status-badge rejected';
    card.querySelector('.analysis-file-icon').className = 'analysis-file-icon rejected';
    actionsDiv.innerHTML = `<div class="admin-reviewed-info">rejected${reason ? ' — ' + reason : ''}</div>`;

    loadStats();
  } catch (e) {
    actionsDiv.innerHTML = '<span class="admin-action-result error">failed to reject</span>';
  }
}

// ========================================
// EDIT ENTRIES — Add / Remove lines
// ========================================

function addLine(analysisId, entryIndex) {
  const block = document.querySelector(`[data-analysis-id="${analysisId}"][data-entry-index="${entryIndex}"]`);
  if (!block) return;

  const tbody = block.querySelector('tbody');
  const lineIndex = tbody.querySelectorAll('tr').length;
  const tr = document.createElement('tr');
  tr.dataset.lineIndex = lineIndex;
  tr.innerHTML = `
    <td>${renderAccountSelect('', '', entryIndex, lineIndex)}</td>
    <td><input type="text" class="edit-cell" value="" placeholder="description" data-field="description" data-entry="${entryIndex}" data-line="${lineIndex}"></td>
    <td style="text-align:right"><input type="number" step="0.01" class="edit-cell edit-amount" value="" placeholder="0.00" data-field="debit" data-entry="${entryIndex}" data-line="${lineIndex}"></td>
    <td style="text-align:right"><input type="number" step="0.01" class="edit-cell edit-amount" value="" placeholder="0.00" data-field="credit" data-entry="${entryIndex}" data-line="${lineIndex}"></td>
    <td><button class="btn-remove-line" onclick="this.closest('tr').remove()" title="remove line">&times;</button></td>
  `;
  tbody.appendChild(tr);

  // Focus the account select
  const select = tr.querySelector('.edit-account-select');
  if (select) select.focus();
  else tr.querySelector('input')?.focus();
}

function removeLine(analysisId, entryIndex, lineIndex) {
  const block = document.querySelector(`[data-analysis-id="${analysisId}"][data-entry-index="${entryIndex}"]`);
  if (!block) return;

  const rows = block.querySelectorAll('tbody tr');
  if (rows.length <= 1) return;

  const row = block.querySelector(`tbody tr[data-line-index="${lineIndex}"]`);
  if (row) {
    row.style.transition = 'opacity 0.2s';
    row.style.opacity = '0';
    setTimeout(() => row.remove(), 200);
  }
}

function addEntry(analysisId) {
  const entriesContainer = document.getElementById(`entries-${analysisId}`);
  if (!entriesContainer) return;

  const entryIndex = entriesContainer.querySelectorAll('.admin-entry-block').length;
  const today = new Date().toISOString().split('T')[0];

  const div = document.createElement('div');
  div.className = 'admin-entry-block';
  div.dataset.analysisId = analysisId;
  div.dataset.entryIndex = entryIndex;
  div.innerHTML = `
    <div class="admin-entry-header">
      <span class="admin-entry-type">journal_entry</span>
      <span class="admin-entry-date"><input type="date" class="edit-date" value="${today}" data-field="date" data-entry="${entryIndex}"></span>
    </div>
    <div class="admin-entry-memo"><input type="text" class="edit-memo" value="" placeholder="memo / description" data-field="memo" data-entry="${entryIndex}"></div>
    <table class="admin-entry-table">
      <thead>
        <tr>
          <th>account</th>
          <th>description</th>
          <th style="text-align:right">debit</th>
          <th style="text-align:right">credit</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        <tr data-line-index="0">
          <td>${renderAccountSelect('', '', entryIndex, 0)}</td>
          <td><input type="text" class="edit-cell" value="" placeholder="description" data-field="description" data-entry="${entryIndex}" data-line="0"></td>
          <td style="text-align:right"><input type="number" step="0.01" class="edit-cell edit-amount" value="" placeholder="0.00" data-field="debit" data-entry="${entryIndex}" data-line="0"></td>
          <td style="text-align:right"><input type="number" step="0.01" class="edit-cell edit-amount" value="" placeholder="0.00" data-field="credit" data-entry="${entryIndex}" data-line="0"></td>
          <td><button class="btn-remove-line" onclick="this.closest('tr').remove()" title="remove line">&times;</button></td>
        </tr>
        <tr data-line-index="1">
          <td>${renderAccountSelect('', '', entryIndex, 1)}</td>
          <td><input type="text" class="edit-cell" value="" placeholder="description" data-field="description" data-entry="${entryIndex}" data-line="1"></td>
          <td style="text-align:right"><input type="number" step="0.01" class="edit-cell edit-amount" value="" placeholder="0.00" data-field="debit" data-entry="${entryIndex}" data-line="1"></td>
          <td style="text-align:right"><input type="number" step="0.01" class="edit-cell edit-amount" value="" placeholder="0.00" data-field="credit" data-entry="${entryIndex}" data-line="1"></td>
          <td><button class="btn-remove-line" onclick="this.closest('tr').remove()" title="remove line">&times;</button></td>
        </tr>
      </tbody>
    </table>
    <div class="entry-table-actions">
      <button class="btn-add-line" onclick="addLine('${analysisId}', ${entryIndex})">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
        add line
      </button>
    </div>
  `;

  entriesContainer.appendChild(div);
  const select = div.querySelector('.edit-account-select');
  if (select) select.focus();
  div.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// ========================================
// FILTERS
// ========================================
document.querySelectorAll('.filter-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentFilter = btn.dataset.filter;
    renderAnalyses();
  });
});

// ========================================
// HELPERS
// ========================================
function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: 'numeric', minute: '2-digit',
  });
}

// ========================================
// VIEW SWITCHING
// ========================================
let currentAdminView = 'documents';

document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentAdminView = btn.dataset.adminView;
    document.querySelectorAll('.admin-view').forEach(v => v.style.display = 'none');
    document.getElementById(`view-${currentAdminView}`).style.display = '';
    if (currentAdminView === 'clients') {
      loadClients();
      showClientsList();
    }
    if (currentAdminView === 'reporting') {
      initReportingView();
    }
  });
});


// ========================================
// CLIENT MANAGEMENT
// ========================================
let allClients = [];
let editingClientId = null;
let selectedClientId = null;
let allFixedAssets = [];
let fixedAssetRuns = [];
let assetClasses = [];
let editingAssetId = null;
let editingClassId = null;

async function loadClients() {
  try {
    const res = await fetch('/api/admin/clients', { headers: { 'Authorization': getAuth() } });
    const data = await res.json();
    allClients = data.clients || [];
    renderClients();
  } catch (e) { console.error('Failed to load clients:', e); }
}

async function renderClients() {
  const container = document.getElementById('clients-list');
  if (allClients.length === 0) {
    container.innerHTML = '<div class="empty-state"><p>no clients yet</p><span>click "+ add client" to create one</span></div>';
    return;
  }

  // Fetch QBO connection status for all clients
  let qboConnections = {};
  try {
    const res = await fetch('/api/qbo/status', { headers: { 'Authorization': getAuth() } });
    const data = await res.json();
    qboConnections = data.connections || {};
  } catch (e) { /* ignore */ }

  container.innerHTML = allClients.map(c => {
    const qboConnected = !!qboConnections[c.id];
    return `
    <div class="client-card client-card-clickable" data-client-id="${c.id}" onclick="openClientDetail('${c.id}')">
      <div class="client-card-info">
        <span class="client-card-name">${c.name}</span>
        <span class="client-card-email">${c.email}</span>
      </div>
      <div class="client-card-meta">
        <span class="status-dot ${qboConnected ? 'connected' : 'disconnected'}" title="QBO ${qboConnected ? 'connected' : 'not connected'}" style="width:8px;height:8px;"></span>
        <span class="client-billing-badge ${c.billingFrequency}">${c.billingFrequency}</span>
        <button class="btn-edit-client" onclick="event.stopPropagation(); openEditClient('${c.id}')">edit</button>
      </div>
    </div>`;
  }).join('');
}

function showClientsList() {
  document.getElementById('clients-list-view').style.display = '';
  document.getElementById('client-detail-view').style.display = 'none';
  selectedClientId = null;
}

async function openClientDetail(clientId) {
  selectedClientId = clientId;
  const client = allClients.find(c => c.id === clientId);
  if (!client) return;

  document.getElementById('clients-list-view').style.display = 'none';
  document.getElementById('client-detail-view').style.display = '';
  document.getElementById('client-detail-name').textContent = client.name;

  // Show info tab by default
  switchClientTab('info');

  // Load QBO status for this client
  loadClientQboStatus(clientId);
  loadQBOCloseDate();

  // Load accounts for this client's QBO connection (await so dropdowns populate before settings render)
  await loadAccounts(clientId);

  // Render client info
  const fyeDisplay = formatFiscalYearEnd(client.fiscalYearEnd);
  document.getElementById('client-info-content').innerHTML = `
    <div style="display:grid;gap:12px;max-width:400px;padding:16px 0;">
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">email</span><br><strong>${client.email}</strong></div>
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">billing</span><br><span class="client-billing-badge ${client.billingFrequency}">${client.billingFrequency}</span></div>
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">fiscal year-end</span><br><strong>${fyeDisplay}</strong></div>
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">client id</span><br><code style="font-size:0.82rem;">${clientId}</code></div>
    </div>`;
}

function formatFiscalYearEnd(mmdd) {
  if (!mmdd || !/^\d{2}-\d{2}$/.test(mmdd)) return '<span style="color:var(--gray-400);">not set</span>';
  const [m, d] = mmdd.split('-').map(Number);
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${monthNames[m - 1]} ${d}`;
}

// ========================================
// PER-CLIENT QBO CONNECTION
// ========================================
async function loadClientQboStatus(clientId) {
  const dot = document.getElementById('client-qbo-dot');
  const text = document.getElementById('client-qbo-text');
  const connectBtn = document.getElementById('btn-client-qbo-connect');
  const disconnectBtn = document.getElementById('btn-client-qbo-disconnect');

  text.textContent = 'checking quickbooks...';
  connectBtn.style.display = 'none';
  disconnectBtn.style.display = 'none';

  try {
    const res = await fetch(`/api/admin/clients/${clientId}/qbo-status`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();

    if (data.connected) {
      dot.classList.remove('disconnected');
      dot.classList.add('connected');
      text.textContent = 'quickbooks connected';
      disconnectBtn.style.display = '';
      connectBtn.style.display = 'none';
      loadQBOCloseDate();
    } else {
      dot.classList.remove('connected');
      dot.classList.add('disconnected');
      text.textContent = 'quickbooks not connected';
      connectBtn.style.display = '';
      disconnectBtn.style.display = 'none';
    }
  } catch (e) {
    text.textContent = 'could not check qbo status';
    connectBtn.style.display = '';
  }
}

function connectClientQbo() {
  if (!selectedClientId) return;
  // Redirect to QBO OAuth with this client's ID
  window.location.href = `/api/qbo/connect?from=admin&clientId=${selectedClientId}`;
}

async function disconnectClientQbo() {
  if (!selectedClientId) return;
  if (!confirm('Disconnect QuickBooks for this client?')) return;

  try {
    await fetch('/api/qbo/disconnect', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ clientId: selectedClientId }),
    });
    loadClientQboStatus(selectedClientId);
    loadStats();
  } catch (e) {
    alert('Failed to disconnect QuickBooks');
  }
}

// Wire up QBO connect/disconnect buttons
document.getElementById('btn-client-qbo-connect').addEventListener('click', connectClientQbo);
document.getElementById('btn-client-qbo-disconnect').addEventListener('click', disconnectClientQbo);

function switchClientTab(tab) {
  document.querySelectorAll('.sub-nav-btn').forEach(b => b.classList.remove('active'));
  document.querySelector(`.sub-nav-btn[data-client-tab="${tab}"]`)?.classList.add('active');
  document.querySelectorAll('.client-tab').forEach(t => t.style.display = 'none');
  document.getElementById(`client-tab-${tab}`).style.display = '';
  // Both 'close' and 'settings' need the fixed-asset data loaded (close uses
  // it to build step cards, settings uses it for the asset classes list).
  if ((tab === 'close' || tab === 'settings') && selectedClientId) {
    loadClientFixedAssets();
  }
}

document.querySelectorAll('.sub-nav-btn').forEach(btn => {
  btn.addEventListener('click', () => switchClientTab(btn.dataset.clientTab));
});

document.getElementById('btn-back-to-clients').addEventListener('click', showClientsList);

function openAddClient() {
  editingClientId = null;
  document.getElementById('client-modal-title').textContent = 'add client';
  document.getElementById('client-name').value = '';
  document.getElementById('client-email').value = '';
  document.getElementById('client-password').value = '';
  document.getElementById('client-password').placeholder = 'set a password for portal login';
  document.getElementById('client-billing').value = 'monthly';
  document.getElementById('client-fye').value = '';
  document.getElementById('client-modal-error').style.display = 'none';
  document.getElementById('client-modal').style.display = '';
}

function openEditClient(id) {
  const client = allClients.find(c => c.id === id);
  if (!client) return;
  editingClientId = id;
  document.getElementById('client-modal-title').textContent = 'edit client';
  document.getElementById('client-name').value = client.name;
  document.getElementById('client-email').value = client.email;
  document.getElementById('client-password').value = '';
  document.getElementById('client-password').placeholder = 'leave blank to keep current password';
  document.getElementById('client-billing').value = client.billingFrequency || 'monthly';
  document.getElementById('client-fye').value = client.fiscalYearEnd || '';
  document.getElementById('client-modal-error').style.display = 'none';
  document.getElementById('client-modal').style.display = '';
}

function closeClientModal() { document.getElementById('client-modal').style.display = 'none'; }

async function saveClient() {
  const name = document.getElementById('client-name').value.trim();
  const email = document.getElementById('client-email').value.trim();
  const password = document.getElementById('client-password').value;
  const billingFrequency = document.getElementById('client-billing').value;
  const fiscalYearEnd = document.getElementById('client-fye').value.trim();
  const errorEl = document.getElementById('client-modal-error');

  if (!name || !email) { errorEl.textContent = 'name and email are required'; errorEl.style.display = ''; return; }
  if (!editingClientId && !password) { errorEl.textContent = 'password is required for new clients'; errorEl.style.display = ''; return; }
  if (fiscalYearEnd && !/^\d{2}-\d{2}$/.test(fiscalYearEnd)) {
    errorEl.textContent = 'fiscal year-end must be MM-DD (e.g. 12-31)'; errorEl.style.display = ''; return;
  }

  const body = { name, email, billingFrequency, fiscalYearEnd: fiscalYearEnd || null };
  if (password) body.password = password;

  try {
    const res = await fetch(editingClientId ? `/api/admin/clients/${editingClientId}` : '/api/admin/clients', {
      method: editingClientId ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (data.error) { errorEl.textContent = data.error; errorEl.style.display = ''; return; }
    closeClientModal();
    loadClients();
  } catch (e) { errorEl.textContent = 'failed to save'; errorEl.style.display = ''; }
}

document.getElementById('btn-add-client').addEventListener('click', () => openAddClient());
document.getElementById('client-modal-cancel').addEventListener('click', () => closeClientModal());
document.getElementById('client-modal-save').addEventListener('click', () => saveClient());
document.getElementById('client-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeClientModal(); });


// ========================================
// FIXED ASSETS (per-client)
// ========================================

// Read-only view of the QBO book close date. QBO's Preferences API accepts
// PUT requests for BookCloseDate and returns 200 OK, but silently echoes the
// old value back instead of applying the change — meaning there's no way for
// a third-party app to actually set this field. Users have to set it in QBO's
// own settings UI, then click refresh here to re-pull the value.
async function loadQBOCloseDate() {
  if (!selectedClientId) return;
  const displayEl = document.getElementById('qbo-close-date-display');
  const openBtn = document.getElementById('btn-open-qbo-close-settings');
  const refreshBtn = document.getElementById('btn-refresh-close-date');
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/book-close-date`, { headers: { 'Authorization': getAuth() } });
    const data = await res.json();
    if (!data.connected) {
      displayEl.textContent = 'QuickBooks not connected — connect QBO to see the close date';
      openBtn.style.display = 'none';
      refreshBtn.style.display = 'none';
      return;
    }
    if (data.closeDate) {
      displayEl.innerHTML = `books closed through: <strong style="color:var(--gray-900);">${data.closeDate}</strong>`;
    } else {
      displayEl.textContent = 'no close date set in QBO';
    }
    openBtn.style.display = '';
    refreshBtn.style.display = '';
  } catch (e) {
    displayEl.textContent = 'failed to load QBO close date';
  }
}

async function loadClientFixedAssets() {
  if (!selectedClientId) return;
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets`, { headers: { 'Authorization': getAuth() } });
    const data = await res.json();
    allFixedAssets = data.assets || [];
    fixedAssetRuns = data.amortizationRuns || [];
    assetClasses = data.assetClasses || [];

    // Auto-create missing GL account policies for any assets that don't have one
    const missingGLAccounts = new Map();
    for (const asset of allFixedAssets) {
      if (!asset.active) continue;
      const glKey = asset.glAccountName || asset.assetAccountName || asset.name;
      if (!glKey) continue;
      const hasClass = assetClasses.some(c =>
        (asset.assetAccountId && c.glAccountId === asset.assetAccountId) ||
        c.glAccountName === glKey
      );
      if (!hasClass && !missingGLAccounts.has(glKey)) {
        missingGLAccounts.set(glKey, {
          glAccountId: asset.assetAccountId || '',
          glAccountName: glKey,
          // Grab expense/accum from the first asset that has them
          expenseAccountId: asset.expenseAccountId || '',
          expenseAccountName: asset.expenseAccountName || '',
          accumAccountId: asset.accumAccountId || '',
          accumAccountName: asset.accumAccountName || '',
        });
      }
    }

    if (missingGLAccounts.size > 0) {
      const newlyCreatedClassIds = [];
      for (const [glName, info] of missingGLAccounts) {
        try {
          const createRes = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/classes`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
            body: JSON.stringify({
              glAccountId: info.glAccountId,
              glAccountName: info.glAccountName,
              method: 'straight-line',
              usefulLifeMonths: 36,
              expenseAccountId: info.expenseAccountId,
              expenseAccountName: info.expenseAccountName,
              accumAccountId: info.accumAccountId,
              accumAccountName: info.accumAccountName,
            }),
          });
          const createData = await createRes.json();
          if (createData.assetClass) {
            assetClasses.push(createData.assetClass);
            newlyCreatedClassIds.push(createData.assetClass.id);
          }
        } catch (e) { console.error('Failed to auto-create class:', e); }
      }
      // Run AI suggestions ONLY for the classes we just created — never re-assess
      // existing policies. The suggestion is saved as advisory data; the user has
      // to open the edit-policy modal and click "apply" to actually adopt it.
      if (newlyCreatedClassIds.length > 0) {
        runAISuggestionsForClasses(newlyCreatedClassIds);
      }
    }

    // Backfill expense/accum accounts on existing classes that are missing them
    for (const cls of assetClasses) {
      if (cls.expenseAccountId && cls.accumAccountId) continue;
      // Find an asset in this class that has expense/accum set
      const donor = allFixedAssets.find(a => a.active &&
        (a.assetAccountId === cls.glAccountId || a.glAccountName === cls.glAccountName) &&
        (a.expenseAccountId || a.accumAccountId));
      if (donor) {
        const patch = {};
        if (!cls.expenseAccountId && donor.expenseAccountId) {
          patch.expenseAccountId = donor.expenseAccountId;
          patch.expenseAccountName = donor.expenseAccountName || '';
        }
        if (!cls.accumAccountId && donor.accumAccountId) {
          patch.accumAccountId = donor.accumAccountId;
          patch.accumAccountName = donor.accumAccountName || '';
        }
        if (Object.keys(patch).length > 0) {
          try {
            await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/classes/${cls.id}`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
              body: JSON.stringify(patch),
            });
            Object.assign(cls, patch);
          } catch (e) { console.error('Failed to backfill class accounts:', e); }
        }
      }
    }

    renderAssetClasses();
    // Load module state — renderCloseSteps depends on all of these
    await Promise.all([loadClientPrepaid(), loadClientAccruedLiab(), loadClientShareholderInvoices()]);
    await renderFixedAssets();
  } catch (e) { console.error('Failed to load fixed assets:', e); }
}

function getAssetClass(asset) {
  return assetClasses.find(c =>
    (asset.assetAccountId && c.glAccountId === asset.assetAccountId) ||
    c.glAccountName === (asset.glAccountName || asset.assetAccountName)
  ) || null;
}

/**
 * Calculate the monthly amortization amount for an asset based on its class policy.
 * Salvage value is read from the asset (it's an asset-level concept, not a class-level
 * concept). Returns null if no policy is set.
 */
function calcMonthlyAmort(asset, cls) {
  if (!cls || !asset.originalCost) return null;
  const cost = parseFloat(asset.originalCost) || 0;
  const salvage = parseFloat(asset.salvageValue) || 0;

  if (cls.method === 'declining-balance') {
    const rate = parseFloat(cls.decliningRate) || 0;
    if (!rate) return null;
    return (cost - salvage) * rate / 12;
  } else {
    const months = parseInt(cls.usefulLifeMonths) || 0;
    if (!months) return null;
    return (cost - salvage) / months;
  }
}

/**
 * Sum life-to-date amortization runs for a single asset (matched by id, then by name).
 * Used to compute accumulated amortization at the asset and class level for the UI.
 */
function ltdAmortForAsset(asset) {
  let total = 0;
  for (const run of fixedAssetRuns || []) {
    for (const a of run.assets || []) {
      if ((a.assetId && a.assetId === asset.id) ||
          (a.id && a.id === asset.id) ||
          (a.name && a.name === asset.name) ||
          (a.assetName && a.assetName === asset.name)) {
        total += parseFloat(a.amount) || 0;
      }
    }
  }
  return total;
}

function renderAssetClasses() {
  const list = document.getElementById('asset-classes-list');
  if (!list) return;

  if (assetClasses.length === 0) {
    list.innerHTML = '<p style="color:var(--gray-400);text-align:center;padding:20px;font-size:0.85rem;">no GL accounts yet — sync from QBO to import assets and their GL accounts</p>';
    return;
  }

  const fmtCurrency = (n) => '$' + (n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  list.innerHTML = assetClasses.map(c => {
    const classAssets = allFixedAssets.filter(a => a.active && (a.assetAccountId === c.glAccountId || a.glAccountName === c.glAccountName));
    const assetCount = classAssets.length;
    let totalCost = 0, totalAccum = 0;
    for (const a of classAssets) {
      totalCost += parseFloat(a.originalCost) || 0;
      totalAccum += ltdAmortForAsset(a);
    }
    const totalNbv = totalCost - totalAccum;

    const methodLabel = c.method === 'declining-balance'
      ? `declining balance @ ${((c.decliningRate || 0) * 100).toFixed(0)}%/yr`
      : `straight-line, ${c.usefulLifeMonths || '—'} months`;
    const ai = c.aiSuggestion;
    const needsAccounts = !c.expenseAccountId || !c.accumAccountId;

    return `
      <div class="asset-class-card ${needsAccounts ? 'asset-needs-setup' : ''}">
        <div class="asset-class-main">
          <div class="asset-class-header">
            <span class="asset-class-name">${c.glAccountName}</span>
            <span class="asset-class-method-badge ${c.method}">${c.method === 'declining-balance' ? 'DB' : 'SL'}</span>
            <span class="asset-class-count">${assetCount} asset${assetCount !== 1 ? 's' : ''}</span>
            ${needsAccounts ? '<span class="asset-setup-badge">needs expense/accum accounts</span>' : ''}
          </div>
          <div class="asset-class-params">
            <span>${methodLabel}</span>
            ${c.expenseAccountName ? `<span>expense acct: ${c.expenseAccountName}</span>` : ''}
            ${c.accumAccountName ? `<span>accum acct: ${c.accumAccountName}</span>` : ''}
          </div>
          <div class="asset-class-totals" style="display:flex;gap:20px;margin-top:8px;padding-top:8px;border-top:1px solid var(--gray-200);font-size:0.82rem;">
            <span style="color:var(--gray-600);">cost: <strong style="color:var(--gray-900);">${fmtCurrency(totalCost)}</strong></span>
            <span style="color:var(--gray-600);">accum amort: <strong style="color:var(--gray-900);">${fmtCurrency(totalAccum)}</strong></span>
            <span style="color:var(--gray-600);">net book value: <strong style="color:var(--gray-900);">${fmtCurrency(totalNbv)}</strong></span>
          </div>
          ${ai ? `<div class="asset-class-ai" style="margin-top:6px;font-size:0.78rem;color:var(--blue-400);">AI suggests: ${ai.method === 'declining-balance' ? `declining balance @ ${((ai.decliningRate || 0) * 100).toFixed(0)}%/yr` : `straight-line, ${ai.usefulLifeMonths || '—'} months`}${ai.ccaClass ? ` — CCA Class ${ai.ccaClass} (${ai.ccaRate})` : ''}${ai.reasoning ? ` — ${ai.reasoning}` : ''}</div>` : ''}
        </div>
        <div class="asset-card-actions">
          <button class="btn-edit-client" onclick="openEditClass('${c.id}')">edit policy</button>
          <button class="btn-delete-asset" onclick="deleteClass('${c.id}')">delete</button>
        </div>
      </div>`;
  }).join('');
}

// Rerender anything in the close-workflow tab that depends on fixedAssetRuns
// or allFixedAssets. Previously this touched a per-asset list; now it just
// re-renders the step cards so status badges stay in sync. Returns the
// renderCloseSteps promise so callers can await the DOM update if they need
// to write into step-owned elements (e.g. the reconciliation panel).
function renderFixedAssets() {
  return renderCloseSteps();
}

// ========================================
// MONTH-END CLOSE WORKFLOW
// ========================================
//
// The close workflow is an ordered list of step cards. Each step has a status
// (complete / ready / locked / skipped / running) and optionally an action.
// Steps are locked by default if any prior non-skipped step is not complete,
// so users work through the list in order.
//
// Adding a new module later (prepaid, accrued liabilities, etc.) = add a new step object
// to the list in buildCloseSteps() and implement its completion check + action
// handler. No UI wiring needed.

// Tracks the current closing period so the step cards can reference it.
let currentClosePeriod = null;

/**
 * Ask the server what period we should be closing. Uses the existing preview-
 * amortization endpoint because it already runs the "next month after last
 * close" logic we need.
 */
async function fetchCurrentClosePeriod() {
  if (!selectedClientId) return null;
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/close-period`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    return {
      month: data.month || null,
      reason: data.reason || '',
      closeDate: data.closeDate || null,
    };
  } catch (e) {
    console.error('fetchCurrentClosePeriod failed', e);
    return null;
  }
}

function formatPeriodLabel(yyyymm) {
  if (!yyyymm) return '—';
  const [y, m] = yyyymm.split('-').map(Number);
  const d = new Date(y, m - 1, 1);
  return d.toLocaleString('en-US', { month: 'long', year: 'numeric' });
}

/**
 * Build the ordered list of workflow steps with status + lock state.
 * Each step: { id, num, title, desc, status, statusLabel, action, actionLabel, meta }
 *
 *   status values:
 *     'complete' — done for this period (green ✓)
 *     'ready'    — can be run right now (blue)
 *     'locked'   — previous required step not complete (gray)
 *     'skipped'  — placeholder module, always counts as complete for deps
 *     'running'  — action in progress (not currently used, reserved)
 */
function buildCloseSteps(period) {
  const month = period?.month;
  const hasAssets = allFixedAssets.some(a => a.active);
  const assetCount = allFixedAssets.filter(a => a.active).length;
  const runForThisMonth = fixedAssetRuns.find(r => r.month === month);
  const lastRun = fixedAssetRuns[fixedAssetRuns.length - 1];

  const steps = [
    {
      id: 'sync',
      num: 1,
      title: 'sync from QuickBooks',
      desc: 'pull the latest fixed-asset accounts from QBO and import any new ones.',
      status: hasAssets ? 'complete' : 'ready',
      statusLabel: hasAssets ? `${assetCount} asset${assetCount !== 1 ? 's' : ''} loaded` : 'not synced',
      action: 'sync',
      actionLabel: hasAssets ? 're-sync' : 'sync now',
      meta: '',
    },
    (() => {
      // Shareholder-paid invoices
      const shiConfigured = !!shiState.shareholderLoanAccount;
      const pendingInvoices = (shiState.invoices || []).filter(i => i.status === 'pending' && i.closeMonth === month);
      const postedInvoices = (shiState.invoices || []).filter(i => i.status === 'posted' && i.closeMonth === month);
      let status, statusLabel, action, actionLabel, meta;
      if (!shiConfigured) {
        status = 'skipped';
        statusLabel = 'not configured';
        action = null;
        actionLabel = 'configure in settings';
        meta = '';
      } else if (postedInvoices.length > 0 && pendingInvoices.length === 0) {
        const total = postedInvoices.reduce((s, i) => s + (i.journalEntry?.totalAmount || 0), 0);
        status = 'complete';
        statusLabel = `posted ${fmtMoney(total)} for ${postedInvoices.length} invoice${postedInvoices.length !== 1 ? 's' : ''}`;
        action = 'upload-shi';
        actionLabel = 'upload more';
        meta = '';
      } else {
        status = 'ready';
        statusLabel = pendingInvoices.length > 0
          ? `${pendingInvoices.length} invoice${pendingInvoices.length !== 1 ? 's' : ''} pending review`
          : 'upload invoices to process';
        action = 'upload-shi';
        actionLabel = 'upload invoice';
        meta = postedInvoices.length > 0 ? `${postedInvoices.length} already posted` : '';
      }
      return {
        id: 'shareholder-invoices',
        num: 2,
        title: 'shareholder-paid invoices',
        desc: 'upload invoices paid by the shareholder personally. AI reads each invoice and suggests an expense.',
        status, statusLabel, action, actionLabel, meta,
      };
    })(),
    (() => {
      // Prepaid expenses — status driven by prepaidPreview + prepaidState
      const configured = !!prepaidState.prepaidAccount;
      const hasItems = (prepaidState.items || []).length > 0;
      const prepaidRun = (prepaidState.amortizationRuns || []).find(r => r.month === month);
      const monthScanned = (prepaidState.scannedMonths || []).includes(month);
      let status, statusLabel, action, actionLabel, meta;
      if (!configured) {
        status = 'skipped';
        statusLabel = 'not configured';
        action = null;
        actionLabel = 'configure in settings';
        meta = '';
      } else if (prepaidRun) {
        status = 'complete';
        statusLabel = `posted $${Number(prepaidRun.totalAmount || 0).toFixed(2)} for ${prepaidRun.itemCount} item${prepaidRun.itemCount !== 1 ? 's' : ''}`;
        action = 'run-prepaid';
        actionLabel = 'view run';
        meta = '';
      } else if (monthScanned && !hasItems) {
        // Scanned but nothing needed prepaid treatment
        status = 'complete';
        statusLabel = 'scanned — no prepaid expenses found';
        action = 'scan-prepaids';
        actionLabel = 're-scan';
        meta = '';
      } else {
        const eligibleCount = prepaidPreview?.eligibleCount || 0;
        const eligibleTotal = Number(prepaidPreview?.totalAmount || 0);
        status = hasItems && eligibleCount > 0 ? 'ready' : (monthScanned ? 'complete' : 'ready');
        statusLabel = eligibleCount > 0
          ? `${eligibleCount} item${eligibleCount !== 1 ? 's' : ''} ready — $${eligibleTotal.toFixed(2)}`
          : (monthScanned ? 'scanned — nothing to amortize' : 'nothing to amortize this period');
        action = eligibleCount > 0 ? 'run-prepaid' : 'scan-prepaids';
        actionLabel = eligibleCount > 0 ? 'run amortization' : 're-scan';
        meta = '';
      }
      // Offer export & Claude review once the period has been processed
      // (either posted a run or scanned and confirmed nothing to amortize).
      const canExport = status === 'complete';
      return {
        id: 'prepaid',
        num: 3,
        title: 'prepaid expenses',
        desc: 'amortize prepaid expenses for the period.',
        status, statusLabel, action, actionLabel, meta,
        secondaryAction: canExport ? 'export-prepaid' : null,
        secondaryActionLabel: canExport ? 'export & review' : null,
      };
    })(),
    (() => {
      const configured = !!accruedLiabState.accruedLiabilitiesAccount;
      const alRun = (accruedLiabState.analysisRuns || []).find(r => r.month === month);
      const posted = !!alRun?.accrualJE;
      let status, statusLabel, action, actionLabel, meta;
      if (!configured) {
        status = 'skipped';
        statusLabel = 'not configured';
        action = null;
        actionLabel = 'configure in settings';
        meta = '';
      } else if (posted) {
        status = 'complete';
        statusLabel = `posted ${fmtMoney(alRun.accrualJE.totalAmount)} (${alRun.accrualJE.lineCount} line${alRun.accrualJE.lineCount !== 1 ? 's' : ''})`;
        action = 'analyze-accrued-liab';
        actionLabel = 'view analysis';
        meta = alRun.reversalJE ? `reversal JE dated ${alRun.reversalJE.date}` : '';
      } else if (alRun) {
        // Analysis ran — check if there's anything to post
        const aAccts = (alRun.partA?.accounts || []).filter(a => a.status !== 'dismissed');
        const bTxns = (alRun.partB?.transactions || []).filter(t => t.status !== 'dismissed');
        const totalA = aAccts.reduce((s, a) => s + (Number(a.accrualAmount) || 0), 0);
        const totalB = bTxns.filter(t => !t.overlapWithPartA || !aAccts.some(a => a.accountId === t.accountId)).reduce((s, t) => s + (Number(t.accrualAmount) || 0), 0);
        const grandTotal = Math.round((totalA + totalB) * 100) / 100;
        if (grandTotal <= 0) {
          status = 'complete';
          statusLabel = 'analyzed — no accrued liabilities detected';
          action = 'analyze-accrued-liab';
          actionLabel = 're-analyze';
          meta = '';
        } else {
          status = 'ready';
          statusLabel = `analysis complete — ${fmtMoney(grandTotal)} to accrue`;
          action = 'analyze-accrued-liab';
          actionLabel = 're-analyze';
          meta = '';
        }
      } else {
        status = 'ready';
        statusLabel = 'ready to analyze';
        action = 'analyze-accrued-liab';
        actionLabel = 'analyze';
        meta = '';
      }
      return {
        id: 'accrued-liabilities',
        num: 4,
        title: 'accrued liabilities',
        desc: 'analyze expense patterns and subsequent transactions to identify missing accruals.',
        status, statusLabel, action, actionLabel, meta,
      };
    })(),
    {
      id: 'receivables',
      num: 5,
      title: 'receivables & deferred revenues',
      desc: 'reconcile AR, defer unearned revenue, recognize earned revenue.',
      status: 'skipped',
      statusLabel: 'not configured',
      action: null,
      actionLabel: 'coming soon',
      meta: '',
    },
    {
      id: 'fixed-assets',
      num: 6,
      title: 'fixed asset amortization',
      desc: 'compute monthly amortization, reconcile to QBO, post the journal entry.',
      status: runForThisMonth ? 'complete' : 'ready',
      statusLabel: runForThisMonth
        ? `posted $${runForThisMonth.totalAmount.toFixed(2)} for ${runForThisMonth.assetCount} asset${runForThisMonth.assetCount !== 1 ? 's' : ''}`
        : 'ready to run',
      action: 'run-amort',
      actionLabel: runForThisMonth ? 'view run' : 'run amortization',
      // Show secondary "export & review" button once amortization has been posted
      // so the user can run the Excel export / Claude review immediately without
      // waiting for step 9.
      secondaryAction: runForThisMonth ? 'export' : null,
      secondaryActionLabel: runForThisMonth ? 'export & review' : null,
      meta: lastRun && !runForThisMonth ? `last run was ${lastRun.month}` : '',
    },
    {
      id: 'income-taxes',
      num: 7,
      title: 'current income taxes',
      desc: 'estimate and accrue current-period income tax expense.',
      status: 'skipped',
      statusLabel: 'not configured',
      action: null,
      actionLabel: 'coming soon',
      meta: '',
    },
    {
      id: 'reconciliation',
      num: 8,
      title: 'overall reconciliation',
      desc: 'tie every module schedule back to the QBO trial balance.',
      status: 'skipped',
      statusLabel: 'per-module only (for now)',
      action: null,
      actionLabel: 'coming soon',
      meta: '',
    },
    {
      id: 'export',
      num: 9,
      title: 'export & review',
      desc: 'generate the Excel workbook and run a Claude review pass over it.',
      status: 'ready',
      statusLabel: 'ready',
      action: 'export',
      actionLabel: 'export to Excel',
      meta: '',
    },
  ];

  // Dependency locking: the export/reconciliation steps at the end should
  // wait until the module steps above are complete (or skipped). Individual
  // module steps (prepaid, accrued liabilities, fixed assets) are independent
  // of each other and can be run in any order.
  const moduleStepIds = new Set(['sync', 'shareholder-invoices', 'prepaid', 'accrued-liabilities', 'receivables', 'fixed-assets', 'income-taxes']);
  const allModulesDone = steps
    .filter(s => moduleStepIds.has(s.id))
    .every(s => s.status === 'complete' || s.status === 'skipped');

  for (const step of steps) {
    if (!moduleStepIds.has(step.id) && step.id !== 'export' && step.id !== 'reconciliation') continue;
    if ((step.id === 'export' || step.id === 'reconciliation') && !allModulesDone && step.status !== 'skipped') {
      step.status = 'locked';
      step.statusLabel = 'complete all modules first';
    }
  }

  return steps;
}

const STEP_STATUS_STYLES = {
  complete: { bg: 'var(--green-50)', border: 'var(--green-500)', badgeBg: 'var(--green-500)', badgeFg: '#ffffff', badgeIcon: '✓' },
  ready:    { bg: 'var(--blue-50)',  border: 'var(--blue-500)',  badgeBg: 'var(--blue-500)',  badgeFg: '#ffffff', badgeIcon: '●' },
  locked:   { bg: 'var(--gray-50)',  border: 'var(--gray-300)',  badgeBg: 'var(--gray-300)',  badgeFg: 'var(--gray-600)', badgeIcon: '🔒' },
  skipped:  { bg: 'var(--gray-50)',  border: 'var(--gray-200)',  badgeBg: 'var(--gray-200)',  badgeFg: 'var(--gray-500)', badgeIcon: '—' },
  running:  { bg: 'var(--blue-50)',  border: 'var(--blue-500)',  badgeBg: 'var(--blue-500)',  badgeFg: '#ffffff', badgeIcon: '…' },
};

function renderStepCard(step) {
  const s = STEP_STATUS_STYLES[step.status] || STEP_STATUS_STYLES.skipped;
  const canAct = step.action && (step.status === 'ready' || step.status === 'complete');
  let actionBtn;
  if (step.id === 'shareholder-invoices') {
    // Shareholder invoices step: file upload button (label wrapping hidden input)
    const disabled = step.status === 'locked' || step.status === 'skipped';
    actionBtn = disabled
      ? `<span class="btn-step-action-placeholder">${step.actionLabel}</span>`
      : `<label class="btn-step-action" style="cursor:pointer;text-align:center;display:inline-block;">
           ${step.actionLabel}
           <input type="file" id="shi-file-input" accept=".pdf,.jpg,.jpeg,.png,.gif,.webp" style="display:none;" onchange="handleShiUpload(this)">
         </label>`;
  } else if (step.id === 'prepaid') {
    // Primary action: scan-prepaids (or "view run" if already posted, via step.action)
    const scanDisabled = step.status === 'locked' ? 'disabled' : '';
    const primaryAction = step.action || 'scan-prepaids';
    const primaryLabel = step.actionLabel || 'scan expenses';
    actionBtn = `<button class="btn-step-action" data-step-action="${primaryAction}" ${scanDisabled}>${primaryLabel}</button>`;
  } else {
    actionBtn = step.action
      ? `<button class="btn-step-action" data-step-action="${step.action}" ${canAct ? '' : 'disabled'}>${step.actionLabel}</button>`
      : `<span class="btn-step-action-placeholder">${step.actionLabel}</span>`;
  }

  // Optional secondary action button (e.g. "export & review" on the fixed-assets
  // step after amortization has been posted).
  if (step.secondaryAction && step.secondaryActionLabel) {
    actionBtn += `<button class="btn-step-action btn-step-action-secondary" data-step-action="${step.secondaryAction}">${step.secondaryActionLabel}</button>`;
  }

  // The fixed-assets step embeds the reconciliation panel AND the Claude review
  // panel (so results appear right under the step when user clicks "export &
  // review"). The export step (9) is kept as a fallback entry point.
  let extras = '';
  if (step.id === 'shareholder-invoices') {
    extras = `<div id="shi-panel" style="margin-top:12px;"></div>`;
  } else if (step.id === 'prepaid') {
    extras = `
      <div id="prepaid-scan-panel" style="margin-top:12px;"></div>
      <div id="prepaid-review-panel" style="display:none;margin-top:12px;"></div>`;
  } else if (step.id === 'accrued-liabilities') {
    extras = `<div id="accrued-liab-panel" style="margin-top:12px;"></div>`;
  } else if (step.id === 'fixed-assets') {
    extras = `
      <div id="reconciliation-panel" style="display:none;margin-top:12px;"></div>
      <div id="claude-review-panel" style="display:none;margin-top:12px;"></div>`;
  }

  return `
    <div class="close-step" data-step-id="${step.id}" style="background:${s.bg};border-left:4px solid ${s.border};">
      <div class="close-step-num">${step.num}</div>
      <div class="close-step-body">
        <div class="close-step-title">${step.title}</div>
        <div class="close-step-desc">${step.desc}</div>
        <div class="close-step-status">
          <span class="close-step-badge" style="background:${s.badgeBg};color:${s.badgeFg};">${s.badgeIcon} ${step.statusLabel}</span>
          ${step.meta ? `<span class="close-step-meta">${step.meta}</span>` : ''}
        </div>
        ${extras}
      </div>
      <div class="close-step-action-col">${actionBtn}</div>
    </div>`;
}

async function renderCloseSteps() {
  const container = document.getElementById('close-steps');
  if (!container || !selectedClientId) return;

  // Resolve current period first (network). Do this once and cache.
  if (!currentClosePeriod || currentClosePeriod._clientId !== selectedClientId) {
    const period = await fetchCurrentClosePeriod();
    if (period) period._clientId = selectedClientId;
    currentClosePeriod = period;
  }

  const monthEl = document.getElementById('close-period-month');
  const reasonEl = document.getElementById('close-period-reason');
  if (monthEl) monthEl.textContent = formatPeriodLabel(currentClosePeriod?.month);
  if (reasonEl) reasonEl.textContent = currentClosePeriod?.reason ? `(${currentClosePeriod.reason})` : '';

  const steps = buildCloseSteps(currentClosePeriod);
  container.innerHTML = steps.map(renderStepCard).join('');

  // Auto-render shareholder invoice panel if there are pending invoices
  const shiPending = (shiState.invoices || []).filter(i => i.closeMonth === currentClosePeriod?.month);
  if (shiPending.length > 0) renderShiPanel();

  // Render fiscal year calendar (non-blocking)
  renderCloseCalendar();
}

// ── Close Calendar (fiscal year overview table) ──────────────────────
async function renderCloseCalendar() {
  const container = document.getElementById('close-calendar');
  if (!container || !selectedClientId) return;

  container.innerHTML = '<div style="color:var(--gray-400);font-size:0.82rem;">loading calendar...</div>';

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/close-calendar`, {
      headers: { 'Authorization': getAuth() },
    });
    if (!res.ok) { container.innerHTML = ''; return; }
    const data = await res.json();

    const statusIcon = (status) => {
      if (status === 'complete') return '<span class="cal-status cal-complete">✓</span>';
      if (status === 'closed')   return '<span class="cal-status cal-closed">—</span>';
      if (status === 'pending')  return '<span class="cal-status cal-pending">○</span>';
      if (status === 'skipped')  return '<span class="cal-skipped">n/a</span>';
      return '<span class="cal-future-dot">·</span>';
    };

    const formatLabel = (monthStr) => {
      // monthStr is "YYYY-MM", convert to "Mar 2026"
      const [y, m] = monthStr.split('-').map(Number);
      const d = new Date(y, m - 1, 1);
      return d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
    };

    const rows = data.months.map(m => {
      const rowClass = m.isCurrent ? 'cal-current' : (m.isFuture ? 'cal-future' : '');
      const label = formatLabel(m.month) + (m.isCurrent ? ' ← current' : '');
      return `
        <tr class="${rowClass}">
          <td>${label}</td>
          <td style="text-align:center;">${statusIcon(m.modules.shareholderInvoices)}</td>
          <td style="text-align:center;">${statusIcon(m.modules.fixedAssets)}</td>
          <td style="text-align:center;">${statusIcon(m.modules.prepaidExpenses)}</td>
          <td style="text-align:center;">${statusIcon(m.modules.accruedLiabilities)}</td>
        </tr>`;
    }).join('');

    // Count completed modules for summary
    const completedMonths = data.months.filter(m =>
      m.modules.fixedAssets === 'complete' ||
      m.modules.prepaidExpenses === 'complete' ||
      m.modules.accruedLiabilities === 'complete'
    ).length;
    const summary = completedMonths > 0
      ? `${completedMonths} of 12 months with activity`
      : 'no months completed yet';

    const isOpen = container.querySelector('.close-calendar-body.open') ? true : false;

    container.innerHTML = `
      <div class="close-calendar">
        <div class="close-calendar-header" onclick="toggleCloseCalendar()">
          <span class="close-calendar-toggle ${isOpen ? 'open' : ''}">&#9654;</span>
          <span class="close-calendar-title">fiscal year overview — ${formatLabel(data.months[0].month)} to ${formatLabel(data.months[data.months.length - 1].month)}</span>
          <span class="close-calendar-summary">${summary}</span>
        </div>
        <div class="close-calendar-body ${isOpen ? 'open' : ''}">
          <table>
            <thead>
              <tr>
                <th>period</th>
                <th style="text-align:center;">SH invoices</th>
                <th style="text-align:center;">fixed assets</th>
                <th style="text-align:center;">prepaid expenses</th>
                <th style="text-align:center;">accrued liabilities</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>`;
  } catch (e) {
    console.error('renderCloseCalendar error:', e);
    container.innerHTML = '';
  }
}

function toggleCloseCalendar() {
  const body = document.querySelector('.close-calendar-body');
  const arrow = document.querySelector('.close-calendar-toggle');
  if (body) body.classList.toggle('open');
  if (arrow) arrow.classList.toggle('open');
}

// Event delegation — all step-card action buttons funnel through this one
// handler, so adding a new step later just needs a new case here.
document.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-step-action]');
  if (!btn || btn.disabled) return;
  const action = btn.dataset.stepAction;
  if (action === 'sync') {
    autoSyncAndImportFromQbo().then(() => { currentClosePeriod = null; renderCloseSteps(); });
  } else if (action === 'run-amort') {
    openRunAmortization();
  } else if (action === 'analyze-accrued-liab') {
    runAccruedLiabAnalysis();
  } else if (action === 'post-accrued-liab') {
    postAccruedLiabilities();
  } else if (action === 'scan-prepaids') {
    runPrepaidScan();
  } else if (action === 'run-prepaid') {
    openRunPrepaid();
  } else if (action === 'export') {
    exportFixedAssetsExcel();
  } else if (action === 'export-prepaid') {
    exportPrepaidExcel();
  }
});

// Named export handler so the step-card button can call it directly (and so
// it's easier to reason about than an anonymous listener).
async function exportFixedAssetsExcel() {
  if (!selectedClientId) return;
  // Handle all "export" buttons on the page (primary on step 9, secondary on fixed-assets card)
  const btns = Array.from(document.querySelectorAll('[data-step-action="export"]'));
  const originalLabels = btns.map(b => b.textContent);
  btns.forEach(b => { b.textContent = 'exporting & reviewing...'; b.disabled = true; });
  // Scroll the review panel into view so the user sees the result as it loads
  const panel = document.getElementById('claude-review-panel');
  if (panel) panel.scrollIntoView({ behavior: 'smooth', block: 'center' });
  renderReviewPanel({ status: 'loading' });
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/export-excel`, {
      headers: { 'Authorization': getAuth() },
    });
    if (!res.ok) { alert('export failed'); renderReviewPanel(null); return; }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const client = allClients.find(c => c.id === selectedClientId);
    a.download = `${client?.name || selectedClientId} - Fixed Assets.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    try {
      const reviewRes = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/last-review`, {
        headers: { 'Authorization': getAuth() },
      });
      const data = await reviewRes.json();
      renderReviewPanel(data.review);
    } catch (e) {
      console.error('Failed to fetch review:', e);
      renderReviewPanel({ status: 'error', summary: 'Could not load review result.', findings: [] });
    }
  } catch (e) { alert('export failed: ' + e.message); renderReviewPanel(null); }
  btns.forEach((b, i) => { b.textContent = originalLabels[i]; b.disabled = false; });
}

// Export the prepaid schedule and run Claude review for the current close period.
// Mirrors exportFixedAssetsExcel but targets the prepaid endpoints and panel.
async function exportPrepaidExcel() {
  if (!selectedClientId) return;
  const btns = Array.from(document.querySelectorAll('[data-step-action="export-prepaid"]'));
  const originalLabels = btns.map(b => b.textContent);
  btns.forEach(b => { b.textContent = 'exporting & reviewing...'; b.disabled = true; });
  const panel = document.getElementById('prepaid-review-panel');
  if (panel) panel.scrollIntoView({ behavior: 'smooth', block: 'center' });
  renderPrepaidReviewPanel({ status: 'loading' });
  try {
    const month = currentClosePeriod?.month;
    const exportUrl = `/api/admin/clients/${selectedClientId}/prepaid-expenses/export-excel${month ? `?month=${month}` : ''}`;
    const res = await fetch(exportUrl, { headers: { 'Authorization': getAuth() } });
    if (!res.ok) { alert('export failed'); renderPrepaidReviewPanel(null); return; }
    const blob = await res.blob();
    const dlUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = dlUrl;
    const client = allClients.find(c => c.id === selectedClientId);
    a.download = `${client?.name || selectedClientId} - Prepaid Expenses.xlsx`;
    a.click();
    URL.revokeObjectURL(dlUrl);
    try {
      const reviewRes = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/last-review`, {
        headers: { 'Authorization': getAuth() },
      });
      const data = await reviewRes.json();
      renderPrepaidReviewPanel(data.review);
    } catch (e) {
      console.error('Failed to fetch prepaid review:', e);
      renderPrepaidReviewPanel({ status: 'error', summary: 'Could not load review result.', findings: [] });
    }
  } catch (e) { alert('export failed: ' + e.message); renderPrepaidReviewPanel(null); }
  btns.forEach((b, i) => { b.textContent = originalLabels[i]; b.disabled = false; });
}

// ========================================
// PREPAID EXPENSES MODULE
// ========================================

async function loadClientPrepaid() {
  if (!selectedClientId) return;
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    prepaidState = {
      prepaidAccount: data.prepaidAccount || null,
      items: data.items || [],
      amortizationRuns: data.amortizationRuns || [],
      scannedMonths: data.scannedMonths || [],
      scanThreshold: data.scanThreshold ?? 500,
    };
    // Also fetch preview for current close period so step card can show eligible total
    try {
      const pvRes = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/preview-amortization`, {
        headers: { 'Authorization': getAuth() },
      });
      prepaidPreview = await pvRes.json();
    } catch (e) { prepaidPreview = null; }
  } catch (e) {
    console.error('Failed to load prepaid expenses:', e);
    prepaidState = { prepaidAccount: null, items: [], amortizationRuns: [] };
    prepaidPreview = null;
  }
  renderPrepaidSettings();
}

function fmtMoney(n) {
  return '$' + (Number(n) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function renderPrepaidSettings() {
  // Account picker — show asset-type accounts plus anything with "prepaid" in the name.
  // Always include the currently-saved account even if it doesn't match the filter.
  const acctSelect = document.getElementById('prepaid-account-select');
  if (acctSelect) {
    const savedId = prepaidState.prepaidAccount?.id;
    const eligible = qboAccounts.filter(a =>
      ['Other Current Asset', 'Other Asset', 'Fixed Asset'].includes(a.type) ||
      (a.name || '').toLowerCase().includes('prepaid')
    );
    // Ensure the saved account is in the list
    if (savedId && !eligible.find(a => a.id === savedId)) {
      const saved = qboAccounts.find(a => a.id === savedId);
      if (saved) eligible.unshift(saved);
      else eligible.unshift({ id: savedId, name: prepaidState.prepaidAccount.name || '(saved account)', type: '?' });
    }
    acctSelect.innerHTML = '<option value="">select prepaid expenses account…</option>' +
      eligible.map(a => acctOption(a, savedId)).join('');
  }

  // Threshold field
  const thresholdInput = document.getElementById('prepaid-scan-threshold');
  if (thresholdInput) {
    thresholdInput.value = prepaidState.scanThreshold ?? 500;
  }

  // Items list
  const list = document.getElementById('prepaid-items-list');
  if (!list) return;

  if (!prepaidState.items.length) {
    list.innerHTML = '<p style="color:var(--gray-400);text-align:center;padding:20px;font-size:0.85rem;">no prepaid items yet — add one manually or import from Excel</p>';
    return;
  }

  list.innerHTML = prepaidState.items.map(item => {
    const completed = !!item.completedAt;
    const totalMonths = (() => {
      const s = new Date(item.startDate);
      const e = new Date(item.endDate);
      return (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth()) + 1;
    })();
    const amortizable = Number(item.openingBalance ?? item.totalAmount) || 0;
    const perMonth = totalMonths > 0 ? amortizable / totalMonths : 0;
    return `
      <div class="asset-class-card ${completed ? 'completed' : ''}">
        <div class="asset-class-main">
          <div class="asset-class-header">
            <span class="asset-class-name">${item.vendor}</span>
            ${completed ? '<span class="asset-class-method-badge">done</span>' : ''}
          </div>
          <div class="asset-class-params">
            ${item.description ? `<span>${item.description}</span>` : ''}
            <span>${item.startDate} → ${item.endDate}</span>
            <span>${totalMonths} month${totalMonths !== 1 ? 's' : ''}</span>
            ${item.expenseAccountName ? `<span>expense acct: ${item.expenseAccountName}</span>` : ''}
          </div>
          <div class="asset-class-totals" style="display:flex;gap:20px;margin-top:8px;padding-top:8px;border-top:1px solid var(--gray-200);font-size:0.82rem;">
            <span style="color:var(--gray-600);">total: <strong style="color:var(--gray-900);">${fmtMoney(item.totalAmount)}</strong></span>
            <span style="color:var(--gray-600);">opening bal: <strong style="color:var(--gray-900);">${fmtMoney(amortizable)}</strong></span>
            <span style="color:var(--gray-600);">per month: <strong style="color:var(--gray-900);">${fmtMoney(perMonth)}</strong></span>
          </div>
        </div>
        <div class="asset-card-actions">
          <button class="btn-edit-client" onclick="openEditPrepaid('${item.id}')">edit</button>
          <button class="btn-delete-asset" onclick="deletePrepaid('${item.id}')">delete</button>
        </div>
      </div>`;
  }).join('');
}

async function savePrepaidSettings() {
  const select = document.getElementById('prepaid-account-select');
  const thresholdInput = document.getElementById('prepaid-scan-threshold');
  const threshold = thresholdInput ? (parseFloat(thresholdInput.value) || 500) : 500;

  try {
    // Save threshold
    await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/settings`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ scanThreshold: threshold }),
    });

    // Save account if selected
    if (select && select.value) {
      const id = select.value;
      const name = select.options[select.selectedIndex].dataset.name;
      const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/account`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
        body: JSON.stringify({ id, name }),
      });
      if (!res.ok) { alert('save failed'); return; }
    }

    await loadClientPrepaid();
    renderCloseSteps();
    showSettingsSaved('btn-save-prepaid-settings');
  } catch (e) { alert('save failed: ' + e.message); }
}

// ---- Add/Edit modal ----

let editingPrepaidId = null;

function openAddPrepaid() {
  editingPrepaidId = null;
  document.getElementById('prepaid-modal-title').textContent = 'add prepaid expense';
  document.getElementById('prepaid-vendor').value = '';
  document.getElementById('prepaid-description').value = '';
  document.getElementById('prepaid-total').value = '';
  document.getElementById('prepaid-opening').value = '';
  document.getElementById('prepaid-start').value = '';
  document.getElementById('prepaid-end').value = '';
  populatePrepaidExpenseAccountSelect('');
  document.getElementById('prepaid-modal').style.display = 'flex';
}

function openEditPrepaid(id) {
  const item = prepaidState.items.find(i => i.id === id);
  if (!item) return;
  editingPrepaidId = id;
  document.getElementById('prepaid-modal-title').textContent = 'edit prepaid expense';
  document.getElementById('prepaid-vendor').value = item.vendor || '';
  document.getElementById('prepaid-description').value = item.description || '';
  document.getElementById('prepaid-total').value = item.totalAmount || '';
  document.getElementById('prepaid-opening').value = item.openingBalance ?? item.totalAmount ?? '';
  document.getElementById('prepaid-start').value = item.startDate || '';
  document.getElementById('prepaid-end').value = item.endDate || '';
  populatePrepaidExpenseAccountSelect(item.expenseAccountId);
  document.getElementById('prepaid-modal').style.display = 'flex';
}

function populatePrepaidExpenseAccountSelect(selectedId) {
  const select = document.getElementById('prepaid-expense-account');
  if (!select) return;
  const expenseAccounts = qboAccounts.filter(a => ['Expense', 'Other Expense', 'Cost of Goods Sold'].includes(a.type));
  select.innerHTML = '<option value="">select expense account…</option>' +
    expenseAccounts.map(a => acctOption(a, selectedId)).join('');
}

function closePrepaidModal() {
  document.getElementById('prepaid-modal').style.display = 'none';
  editingPrepaidId = null;
}

async function savePrepaid() {
  const vendor = document.getElementById('prepaid-vendor').value.trim();
  const description = document.getElementById('prepaid-description').value.trim();
  const totalAmount = parseFloat(document.getElementById('prepaid-total').value);
  const openingRaw = document.getElementById('prepaid-opening').value;
  const openingBalance = openingRaw === '' ? totalAmount : parseFloat(openingRaw);
  const startDate = document.getElementById('prepaid-start').value;
  const endDate = document.getElementById('prepaid-end').value;
  const select = document.getElementById('prepaid-expense-account');
  const expenseAccountId = select.value;
  const expenseAccountName = select.options[select.selectedIndex]?.dataset.name || '';

  if (!vendor || !totalAmount || !startDate || !endDate || !expenseAccountId) {
    alert('vendor, amount, start/end dates, and expense account are required');
    return;
  }

  const payload = { vendor, description, totalAmount, openingBalance, startDate, endDate, expenseAccountId, expenseAccountName };

  try {
    const url = editingPrepaidId
      ? `/api/admin/clients/${selectedClientId}/prepaid-expenses/items/${editingPrepaidId}`
      : `/api/admin/clients/${selectedClientId}/prepaid-expenses/items`;
    const method = editingPrepaidId ? 'PUT' : 'POST';
    const res = await fetch(url, {
      method,
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify(payload),
    });
    if (!res.ok) { const e = await res.json().catch(() => ({})); alert('save failed: ' + (e.error || res.status)); return; }
    closePrepaidModal();
    await loadClientPrepaid();
    renderCloseSteps();
  } catch (e) { alert('save failed: ' + e.message); }
}

async function deletePrepaid(id) {
  if (!confirm('delete this prepaid item?')) return;
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/items/${id}`, {
      method: 'DELETE',
      headers: { 'Authorization': getAuth() },
    });
    if (!res.ok) { alert('delete failed'); return; }
    await loadClientPrepaid();
    renderCloseSteps();
  } catch (e) { alert('delete failed: ' + e.message); }
}

// ---- Excel import/export ----

async function downloadPrepaidTemplate() {
  const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/export-template`, {
    headers: { 'Authorization': getAuth() },
  });
  if (!res.ok) { alert('download failed'); return; }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'prepaid-expenses-template.xlsx';
  a.click();
  URL.revokeObjectURL(url);
}

async function importPrepaidExcel(file) {
  const formData = new FormData();
  formData.append('file', file);
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/import-excel`, {
      method: 'POST',
      headers: { 'Authorization': getAuth() },
      body: formData,
    });
    const data = await res.json();
    if (!res.ok) { alert('import failed: ' + (data.error || res.status)); return; }
    let msg = `imported ${data.itemCount} item${data.itemCount !== 1 ? 's' : ''}`;
    if (data.warnings?.length) msg += '\n\nwarnings:\n' + data.warnings.join('\n');
    alert(msg);
    await loadClientPrepaid();
    renderCloseSteps();
  } catch (e) { alert('import failed: ' + e.message); }
}

// ---- Run amortization ----

async function openRunPrepaid() {
  if (!selectedClientId) return;
  try {
    // Refresh preview
    const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/preview-amortization`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();

    if (data.alreadyRun) {
      alert(`Prepaid amortization already run for ${data.month}.\n\nTotal: ${fmtMoney(data.runDetails.totalAmount)}\nItems: ${data.runDetails.itemCount}`);
      return;
    }
    if (data.notConfigured) {
      alert('Prepaid Expenses GL account is not set. Configure it in the settings tab first.');
      return;
    }
    if (!data.lines || data.lines.length === 0) {
      alert(`No prepaid items to amortize for ${data.month}.`);
      return;
    }

    const linesText = data.lines.map(l =>
      `  ${l.vendor}${l.description ? ' - ' + l.description : ''}: ${fmtMoney(l.amount)} → ${l.expenseAccountName || '(no account)'}`
    ).join('\n');
    const confirmMsg = `Post prepaid amortization for ${data.month}?\n\n${linesText}\n\nTotal: ${fmtMoney(data.totalAmount)}\nCredit: ${data.prepaidAccount.name}\n\nThis will post a journal entry to QBO.`;
    if (!confirm(confirmMsg)) return;

    const postRes = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/run-amortization`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ month: data.month }),
    });
    const postData = await postRes.json();
    if (!postRes.ok) { alert('run failed: ' + (postData.error || postRes.status)); return; }
    alert(`Posted prepaid amortization for ${postData.run.month}\nTotal: ${fmtMoney(postData.run.totalAmount)}\nJE ID: ${postData.run.journalEntryId || '(unknown)'}`);
    await loadClientPrepaid();
    renderCloseSteps();
  } catch (e) { alert('error: ' + e.message); }
}

// ========================================
// ACCRUED LIABILITIES MODULE
// ========================================

async function loadClientAccruedLiab() {
  if (!selectedClientId) return;
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/accrued-liabilities`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    accruedLiabState = {
      accruedLiabilitiesAccount: data.accruedLiabilitiesAccount || null,
      analysisRuns: data.analysisRuns || [],
      materialityThreshold: data.materialityThreshold ?? 10,
      excludedAccountIds: data.excludedAccountIds || [],
    };
  } catch (e) {
    console.error('Failed to load accrued liabilities:', e);
    accruedLiabState = { accruedLiabilitiesAccount: null, analysisRuns: [], materialityThreshold: 10 };
  }
  renderAccruedLiabSettings();
}

function renderAccruedLiabSettings() {
  const acctSelect = document.getElementById('accrued-liab-account-select');
  if (acctSelect) {
    const savedId = accruedLiabState.accruedLiabilitiesAccount?.id;
    const eligible = qboAccounts.filter(a =>
      ['Other Current Liability', 'Long Term Liability', 'Accounts Payable', 'Other Liability'].includes(a.type) ||
      (a.name || '').toLowerCase().includes('accrued')
    );
    if (savedId && !eligible.find(a => a.id === savedId)) {
      const saved = qboAccounts.find(a => a.id === savedId);
      if (saved) eligible.unshift(saved);
      else eligible.unshift({ id: savedId, name: accruedLiabState.accruedLiabilitiesAccount.name || '(saved)', type: '?' });
    }
    acctSelect.innerHTML = '<option value="">select accrued liabilities account…</option>' +
      eligible.map(a => acctOption(a, savedId)).join('');
  }
  const thresholdInput = document.getElementById('accrued-liab-threshold');
  if (thresholdInput) thresholdInput.value = accruedLiabState.materialityThreshold ?? 10;
}

async function saveAccruedLiabSettings() {
  const select = document.getElementById('accrued-liab-account-select');
  const thresholdInput = document.getElementById('accrued-liab-threshold');
  const threshold = thresholdInput ? (parseFloat(thresholdInput.value) || 10) : 10;

  try {
    await fetch(`/api/admin/clients/${selectedClientId}/accrued-liabilities/settings`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ materialityThreshold: threshold }),
    });
    if (select && select.value) {
      await fetch(`/api/admin/clients/${selectedClientId}/accrued-liabilities/account`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
        body: JSON.stringify({ id: select.value, name: select.options[select.selectedIndex].dataset.name }),
      });
    }
    await loadClientAccruedLiab();
    renderCloseSteps();
    showSettingsSaved('btn-save-accrued-liab-settings');
  } catch (e) { alert('save failed: ' + e.message); }
}

// ---- Analysis ----

async function runAccruedLiabAnalysis() {
  if (!selectedClientId) return;
  const panel = document.getElementById('accrued-liab-panel');
  if (!panel) return;

  const analyzeBtn = document.querySelector('[data-step-action="analyze-accrued-liab"]');
  if (analyzeBtn) { analyzeBtn.textContent = 'analyzing…'; analyzeBtn.disabled = true; }
  panel.innerHTML = `
    <div style="padding:16px;background:var(--blue-50);border-radius:8px;font-size:0.85rem;color:var(--gray-700);">
      <strong>analyzing expense patterns…</strong><br>
      pulling 18-month P&L from QBO and scanning subsequent transactions. this may take 10–20 seconds.
    </div>`;

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/accrued-liabilities/analyze`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ month: currentClosePeriod?.month || undefined }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'analysis failed');
    accruedLiabAnalysis = data;
    await loadClientAccruedLiab();
    await renderCloseSteps();
    // Re-render results into the freshly created panel (renderCloseSteps rebuilds the DOM)
    const freshPanel = document.getElementById('accrued-liab-panel');
    if (freshPanel) renderAccruedLiabResults(freshPanel, data);
  } catch (e) {
    panel.innerHTML = `<div style="padding:12px;background:var(--red-50);border-radius:8px;color:var(--red-600);">analysis failed: ${e.message}</div>`;
  }
  if (analyzeBtn) { analyzeBtn.textContent = 'analyze'; analyzeBtn.disabled = false; }
}

function renderAccruedLiabResults(panel, data) {
  const partA = data.partA || { accounts: [] };
  const partB = data.partB || { transactions: [] };
  const flaggedA = partA.accounts.filter(a => a.status !== 'dismissed');
  const flaggedB = partB.transactions.filter(t => t.status !== 'dismissed');

  const diag = partA.diagnostics || null;
  const partBStatus = partB.skipped
    ? `<span style="color:var(--gray-500);">skipped — ${partB.note || 'window invalid'}</span>`
    : `${partB.transactions.length} subsequent transaction${partB.transactions.length !== 1 ? 's' : ''} (${partB.windowStart} → ${partB.windowEnd})`;

  let html = `
    <div style="padding:12px;background:var(--gray-50);border-radius:8px;margin-bottom:10px;font-size:0.85rem;">
      <strong>analysis for ${data.month}</strong> —
      Part A: ${partA.accounts.length} expense account${partA.accounts.length !== 1 ? 's' : ''} flagged (${partA.lookbackMonths}-month lookback) |
      Part B: ${partBStatus}
      ${data.alreadyPosted ? '<br><span style="color:var(--green-600);">✓ accrual JE already posted</span>' : ''}
      ${diag ? `
        <div style="margin-top:6px;font-size:0.75rem;color:var(--gray-600);">
          P&amp;L ${diag.pnlRangeStart || '?'} → ${diag.pnlRangeEnd || '?'} returned ${diag.monthsReturned ?? '?'} month column${diag.monthsReturned === 1 ? '' : 's'} (${diag.priorMonthsCount ?? '?'} prior to ${data.month}).
          ${diag.monthsReturned === 0 ? '<span style="color:var(--red-600);font-weight:600;">⚠ P&amp;L column parsing returned zero months — analysis cannot run until this is fixed.</span><br>' : ''}
          scanned ${diag.totalAccounts} P&amp;L account${diag.totalAccounts !== 1 ? 's' : ''}:
          ${diag.evaluated} evaluated,
          ${diag.negligibleHistory} skipped (no history),
          ${diag.excluded} excluded by settings,
          ${diag.withinTolerance} within materiality tolerance,
          ${diag.zeroCurrentBelowFrequency} zero this month but below frequency threshold
          ${diag.monthsReturned === 0 && Array.isArray(diag.rawColumns) && diag.rawColumns.length ? `
            <details style="margin-top:6px;">
              <summary style="cursor:pointer;color:var(--red-600);font-weight:600;">show raw QBO columns (${diag.rawColumns.length})</summary>
              <pre style="margin:6px 0 0;padding:8px;background:#fff;border:1px solid var(--gray-300);border-radius:4px;font-size:0.7rem;white-space:pre-wrap;max-height:200px;overflow:auto;">${JSON.stringify(diag.rawColumns, null, 2)}</pre>
            </details>` : ''}
        </div>` : ''}
    </div>`;

  // Near-miss panel: accounts that appeared last month but didn't meet the flag threshold
  if (partA.accounts.length === 0 && diag && diag.nearMiss && diag.nearMiss.length > 0) {
    html += `
      <div style="padding:10px 12px;background:#fffbeb;border:1px solid #fde68a;border-radius:8px;margin-bottom:10px;font-size:0.8rem;">
        <div style="font-weight:600;color:#854d0e;margin-bottom:4px;">⚠ possible missed accruals — these accounts had activity last month but didn't meet the recurring-pattern threshold:</div>
        <ul style="margin:4px 0 0 18px;padding:0;color:#713f12;">
          ${diag.nearMiss.map(nm => `<li>${nm.accountName} — last month: ${fmtMoney(nm.priorMonthAmount)} (seen in ${nm.frequency} of prior ${partA.lookbackMonths} months)</li>`).join('')}
        </ul>
        <div style="margin-top:6px;font-size:0.75rem;color:#713f12;">If any of these are recurring, review them manually or record a one-off accrual via journal entry.</div>
      </div>`;
  }

  // Part A table
  if (partA.accounts.length > 0) {
    html += `<div style="margin-bottom:6px;font-size:0.85rem;font-weight:600;">part A — expense pattern gaps</div>`;
    html += `<div style="overflow-x:auto;margin-bottom:12px;">
      <table style="width:100%;border-collapse:collapse;font-size:0.8rem;">
        <thead>
          <tr style="border-bottom:2px solid var(--gray-300);text-align:left;">
            <th style="padding:6px 8px;width:24px;"></th>
            <th style="padding:6px 8px;">account</th>
            <th style="padding:6px 8px;text-align:right;" title="Average across months that had activity (excludes empty months) — better estimate of recurring cost than the window average">avg / active mo</th>
            <th style="padding:6px 8px;text-align:center;" title="Months with activity / total prior months in the lookback window">active mos</th>
            <th style="padding:6px 8px;text-align:right;">this month</th>
            <th style="padding:6px 8px;">reason</th>
            <th style="padding:6px 8px;text-align:right;">accrual amt</th>
            <th style="padding:6px 8px;text-align:center;">action</th>
          </tr>
        </thead>
        <tbody>
          ${partA.accounts.map((a, i) => {
            const dismissed = a.status === 'dismissed';
            const rowStyle = dismissed ? 'opacity:0.4;text-decoration:line-through;' : '';
            const activeAvg = a.activeMonthAverage ?? a.average ?? 0;
            const freqLabel = `${a.frequency || 0}/${a.monthsInLookback ?? partA.lookbackMonths}`;
            return `<tr style="border-bottom:1px solid var(--gray-200);${rowStyle}">
              <td style="padding:6px 4px;text-align:center;vertical-align:middle;"><button type="button" class="al-detail-toggle" data-al-detail-index="${i}" style="background:none;border:1px solid var(--gray-300);border-radius:4px;width:20px;height:20px;cursor:pointer;font-size:0.7rem;line-height:1;padding:0;" title="show calculation details">▸</button></td>
              <td style="padding:6px 8px;">${a.accountName}</td>
              <td style="padding:6px 8px;text-align:right;">${fmtMoney(activeAvg)}</td>
              <td style="padding:6px 8px;text-align:center;font-variant-numeric:tabular-nums;color:var(--gray-600);">${freqLabel}</td>
              <td style="padding:6px 8px;text-align:right;">${fmtMoney(a.currentMonth)}</td>
              <td style="padding:6px 8px;"><span style="font-size:0.72rem;padding:1px 6px;border-radius:4px;background:${String(a.reason).startsWith('missing') ? '#fef2f2' : '#fef9c3'};color:${String(a.reason).startsWith('missing') ? '#991b1b' : '#854d0e'};" title="${a.reason === 'missing_new_recurring' ? 'Appeared last month but not this month — likely a new recurring subscription.' : a.reason === 'missing_recurring' ? 'Recurring pattern (including last month) but no current-month activity.' : a.reason === 'missing' ? 'Account with established history has no current-month activity.' : 'Current-month activity is below the materiality-adjusted active-month average.'}">${String(a.reason).replace(/_/g, ' ')}</span></td>
              <td style="padding:6px 8px;text-align:right;"><input type="number" step="0.01" value="${a.accrualAmount}" style="width:90px;padding:2px 4px;text-align:right;border:1px solid var(--gray-300);border-radius:3px;" data-al-part="A" data-al-index="${i}" onchange="updateAccruedLiabItem(this)" ${dismissed ? 'disabled' : ''}></td>
              <td style="padding:6px 8px;text-align:center;">
                ${dismissed
                  ? `<button class="btn btn-secondary" style="font-size:0.72rem;padding:2px 8px;" onclick="toggleAccruedLiabDismiss('A',${i},false)">restore</button>`
                  : `<button class="btn btn-secondary" style="font-size:0.72rem;padding:2px 8px;" onclick="toggleAccruedLiabDismiss('A',${i},true)">dismiss</button>`
                }
              </td>
            </tr>
            <tr data-al-detail-row="${i}" style="display:none;background:#f8fafc;">
              <td colspan="8" style="padding:10px 14px;font-size:0.78rem;color:var(--gray-700);">
                ${renderAccruedLiabDetail(a, partA.lookbackMonths, data.month, partA.priorMonth)}
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>`;
  }

  // Part B table
  if (partB.transactions.length > 0) {
    html += `<div style="margin-bottom:6px;font-size:0.85rem;font-weight:600;">part B — subsequent transactions</div>`;
    html += `<div style="overflow-x:auto;margin-bottom:12px;">
      <table style="width:100%;border-collapse:collapse;font-size:0.8rem;">
        <thead>
          <tr style="border-bottom:2px solid var(--gray-300);text-align:left;">
            <th style="padding:6px 8px;">date</th>
            <th style="padding:6px 8px;">vendor</th>
            <th style="padding:6px 8px;text-align:right;">amount</th>
            <th style="padding:6px 8px;">account</th>
            <th style="padding:6px 8px;">memo</th>
            <th style="padding:6px 8px;text-align:right;">accrual amt</th>
            <th style="padding:6px 8px;text-align:center;">action</th>
          </tr>
        </thead>
        <tbody>
          ${partB.transactions.map((t, i) => {
            const dismissed = t.status === 'dismissed';
            const rowStyle = dismissed ? 'opacity:0.4;text-decoration:line-through;' : '';
            const overlap = t.overlapWithPartA ? '<span title="Also flagged in Part A" style="color:var(--blue-500);font-size:0.72rem;">↑A</span> ' : '';
            return `<tr style="border-bottom:1px solid var(--gray-200);${rowStyle}">
              <td style="padding:6px 8px;white-space:nowrap;">${t.date}</td>
              <td style="padding:6px 8px;">${overlap}${t.vendor}</td>
              <td style="padding:6px 8px;text-align:right;">${fmtMoney(t.amount)}</td>
              <td style="padding:6px 8px;font-size:0.78rem;">${t.accountName}</td>
              <td style="padding:6px 8px;font-size:0.78rem;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${(t.memo || '').replace(/"/g, '&quot;')}">${t.memo || '—'}</td>
              <td style="padding:6px 8px;text-align:right;"><input type="number" step="0.01" value="${t.accrualAmount}" style="width:90px;padding:2px 4px;text-align:right;border:1px solid var(--gray-300);border-radius:3px;" data-al-part="B" data-al-index="${i}" onchange="updateAccruedLiabItem(this)" ${dismissed ? 'disabled' : ''}></td>
              <td style="padding:6px 8px;text-align:center;">
                ${dismissed
                  ? `<button class="btn btn-secondary" style="font-size:0.72rem;padding:2px 8px;" onclick="toggleAccruedLiabDismiss('B',${i},false)">restore</button>`
                  : `<button class="btn btn-secondary" style="font-size:0.72rem;padding:2px 8px;" onclick="toggleAccruedLiabDismiss('B',${i},true)">dismiss</button>`
                }
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>`;
  }

  // Summary + post button
  const totalA = flaggedA.reduce((s, a) => s + (Number(a.accrualAmount) || 0), 0);
  const totalB = flaggedB.filter(t => !t.overlapWithPartA || !flaggedA.some(a => a.accountId === t.accountId)).reduce((s, t) => s + (Number(t.accrualAmount) || 0), 0);
  const grandTotal = Math.round((totalA + totalB) * 100) / 100;

  if (!data.alreadyPosted) {
    if (grandTotal <= 0) {
      html += `
        <div style="display:flex;align-items:center;gap:8px;padding:12px;background:var(--green-50);border-radius:8px;font-size:0.85rem;color:var(--green-700);">
          <span style="font-size:1.1rem;">✓</span>
          <strong>no accrued liabilities detected — nothing to post</strong>
        </div>`;
    } else {
      html += `
        <div style="display:flex;justify-content:space-between;align-items:center;padding:12px;background:var(--blue-50);border-radius:8px;font-size:0.85rem;">
          <div>
            <strong>total accrual: ${fmtMoney(grandTotal)}</strong>
            <span style="color:var(--gray-500);margin-left:8px;">(Part A: ${fmtMoney(totalA)} + Part B: ${fmtMoney(totalB)})</span>
          </div>
          <button class="btn btn-primary" onclick="showAccrualPostModal()">review & post accrual JE</button>
        </div>`;
    }
  }

  panel.innerHTML = html;

  // Wire up the per-item detail toggles
  panel.querySelectorAll('.al-detail-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = btn.dataset.alDetailIndex;
      const row = panel.querySelector(`tr[data-al-detail-row="${idx}"]`);
      if (!row) return;
      const isHidden = row.style.display === 'none';
      row.style.display = isHidden ? 'table-row' : 'none';
      btn.textContent = isHidden ? '▾' : '▸';
      btn.title = isHidden ? 'hide calculation details' : 'show calculation details';
    });
  });
}

/**
 * Renders the expandable detail block for a Part A flagged account.
 * Three sections: (1) plain-language reason, (2) full math breakdown
 * exposing every variable in the calculation, (3) vendor itemization
 * from the prior month's transactions, (4) recent monthly history.
 */
function renderAccruedLiabDetail(a, lookbackMonths, targetMonth, priorMonthKey) {
  const reasonExplain = {
    missing: `<strong>Missing</strong> — this account has an established history (${a.frequency} of prior ${a.monthsInLookback ?? lookbackMonths} months had activity) but no activity was recorded in ${targetMonth}.`,
    missing_recurring: `<strong>Missing — recurring</strong> — appeared last month and in ${a.frequency} of the prior ${a.monthsInLookback ?? lookbackMonths} months. Likely a recurring expense that hasn't been entered yet for ${targetMonth}.`,
    missing_new_recurring: `<strong>Missing — new recurring</strong> — first appeared in ${a.firstActiveMonth || 'a recent month'} (${a.monthsSinceFirstActivity ?? 1} month${(a.monthsSinceFirstActivity ?? 1) === 1 ? '' : 's'} of history). Likely a new subscription that should recur this month.`,
    below_average: `<strong>Below average</strong> — this account had ${fmtMoney(a.currentMonth)} this month vs the average of ${fmtMoney(a.activeMonthAverage ?? a.average)} per active month, which is below the materiality threshold.`,
  };

  // ---- Math breakdown ----
  // All values come from the server's `calculation` object plus the headline numbers.
  const calc = a.calculation || {};
  const totalSpend = a.totalPriorSpend ?? 0;
  const monthsInLookback = a.monthsInLookback ?? lookbackMonths ?? 0;
  const monthsSinceFirst = a.monthsSinceFirstActivity ?? 0;
  const freq = a.frequency ?? 0;

  const rowsHtml = [
    { label: 'Σ prior spend (all months in lookback)', value: fmtMoney(totalSpend), note: `sum across the ${monthsInLookback}-month window` },
    { label: 'Months in lookback window', value: monthsInLookback, note: 'P&L monthly columns returned for the prior-period window' },
    { label: 'Months since first activity', value: monthsSinceFirst, note: `first seen ${a.firstActiveMonth || '—'}` },
    { label: 'Months with activity (frequency)', value: freq, note: 'months with non-zero activity' },
    { label: 'Window average (Σ ÷ lookback months)', value: fmtMoney(totalSpend / Math.max(1, monthsInLookback)), note: 'understates recurring cost when history is short' },
    { label: 'History average (Σ ÷ months since first activity)', value: fmtMoney(monthsSinceFirst > 0 ? totalSpend / monthsSinceFirst : 0), note: 'better when activity is consistent month-to-month' },
    { label: 'Active-month average (Σ ÷ frequency)', value: fmtMoney(freq > 0 ? totalSpend / freq : 0), note: '★ best estimate of per-occurrence cost', highlight: calc.basis === 'active_month_average' },
    { label: `Prior month actual (${a.priorMonthKey || priorMonthKey || '—'})`, value: fmtMoney(a.priorMonthAmount ?? 0), note: 'most direct signal when only 1-2 months of history exist', highlight: calc.basis === 'prior_month_actual' },
    { label: `This month actual (${targetMonth})`, value: fmtMoney(a.currentMonth ?? 0), note: 'what is currently in QBO for this month', highlight: calc.basis === 'gap_to_active_average' },
  ];
  const mathTable = `
    <table style="width:100%;border-collapse:collapse;font-size:0.72rem;">
      <tbody>
        ${rowsHtml.map(r => `
          <tr style="border-bottom:1px solid var(--gray-200);${r.highlight ? 'background:#dcfce7;font-weight:600;' : ''}">
            <td style="padding:3px 6px;">${r.label}</td>
            <td style="padding:3px 6px;text-align:right;font-variant-numeric:tabular-nums;">${r.value}</td>
            <td style="padding:3px 6px;color:var(--gray-500);font-size:0.68rem;">${r.note}</td>
          </tr>`).join('')}
      </tbody>
    </table>
    <div style="margin-top:6px;padding:6px 8px;background:#dbeafe;border-radius:4px;font-size:0.72rem;">
      <strong>Suggested accrual = ${fmtMoney(a.accrualAmount)}</strong><br>
      <span style="color:var(--gray-700);">${calc.formula || '—'}</span>
    </div>`;

  // ---- Vendor breakdown ----
  const vb = Array.isArray(a.vendorBreakdown) ? a.vendorBreakdown : [];
  const vbTotal = vb.reduce((s, v) => s + (Number(v.amount) || 0), 0);
  const reconciles = Math.abs(vbTotal - (a.accrualAmount || 0)) <= 0.5;
  const vendorTable = vb.length
    ? `<table style="width:100%;border-collapse:collapse;font-size:0.72rem;">
         <thead>
           <tr style="border-bottom:1px solid var(--gray-300);text-align:left;">
             <th style="padding:3px 6px;">vendor</th>
             <th style="padding:3px 6px;text-align:center;">#</th>
             <th style="padding:3px 6px;text-align:right;">prior-month total</th>
             <th style="padding:3px 6px;">latest description</th>
           </tr>
         </thead>
         <tbody>
           ${vb.map(v => `
             <tr style="border-bottom:1px solid var(--gray-200);">
               <td style="padding:3px 6px;font-weight:500;">${v.vendor}</td>
               <td style="padding:3px 6px;text-align:center;color:var(--gray-600);">${v.count}</td>
               <td style="padding:3px 6px;text-align:right;font-variant-numeric:tabular-nums;">${fmtMoney(v.amount)}</td>
               <td style="padding:3px 6px;color:var(--gray-600);font-size:0.68rem;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${((v.transactions?.[0]?.memo) || '').replace(/"/g, '&quot;')}">${(v.transactions?.[0]?.docNumber ? `#${v.transactions[0].docNumber} — ` : '')}${v.transactions?.[0]?.memo || '—'}</td>
             </tr>`).join('')}
           <tr style="font-weight:600;${reconciles ? 'background:#dcfce7;' : 'background:#fef3c7;'}">
             <td style="padding:3px 6px;">total</td>
             <td style="padding:3px 6px;"></td>
             <td style="padding:3px 6px;text-align:right;font-variant-numeric:tabular-nums;">${fmtMoney(vbTotal)}</td>
             <td style="padding:3px 6px;font-size:0.68rem;color:${reconciles ? '#166534' : '#854d0e'};">${reconciles ? '✓ reconciles to suggested accrual' : `differs from suggested accrual by ${fmtMoney(Math.abs(vbTotal - (a.accrualAmount || 0)))}`}</td>
           </tr>
         </tbody>
       </table>
       <div style="margin-top:4px;color:var(--gray-500);font-size:0.68rem;">JE will be itemized per vendor when the breakdown reconciles. Otherwise a single rolled-up line is posted.</div>`
    : `<div style="color:var(--gray-500);font-style:italic;">No prior-month bills/expenses found for this account in QBO. JE will be posted as a single rolled-up line.</div>`;

  // ---- Monthly history (mini chart) ----
  const monthly = a.monthlyTotals || {};
  const sortedMonths = Object.keys(monthly).sort();
  const recent = sortedMonths.slice(-12);
  const maxAmt = recent.reduce((m, k) => Math.max(m, Math.abs(monthly[k] || 0)), 0) || 1;
  const pmKey = a.priorMonthKey || priorMonthKey;
  const historyHtml = recent.length
    ? `<table style="width:auto;border-collapse:collapse;font-size:0.7rem;">
         <tbody>
           ${recent.map(mk => {
             const v = monthly[mk] || 0;
             const isPriorMonth = mk === pmKey;
             const barWidth = Math.max(1, Math.round((Math.abs(v) / maxAmt) * 100));
             const highlight = isPriorMonth ? 'background:#fef3c7;' : '';
             return `<tr style="${highlight}">
               <td style="padding:2px 8px;color:var(--gray-600);">${mk}${isPriorMonth ? ' <span style="color:#854d0e;font-weight:600;">(prior month)</span>' : ''}</td>
               <td style="padding:2px 8px;text-align:right;font-variant-numeric:tabular-nums;">${fmtMoney(v)}</td>
               <td style="padding:2px 8px;width:100px;"><div style="height:5px;background:var(--gray-200);border-radius:3px;overflow:hidden;"><div style="width:${barWidth}%;height:100%;background:${isPriorMonth ? '#f59e0b' : '#94a3b8'};"></div></div></td>
             </tr>`;
           }).join('')}
         </tbody>
       </table>`
    : `<div style="color:var(--gray-500);font-style:italic;">No prior month activity recorded.</div>`;

  return `
    <div style="display:grid;grid-template-columns:1fr;gap:14px;">
      <div>
        <div style="font-weight:600;margin-bottom:4px;font-size:0.82rem;">why flagged</div>
        <div>${reasonExplain[a.reason] || `<strong>${a.reason}</strong>`}</div>
      </div>
      <div>
        <div style="font-weight:600;margin-bottom:4px;font-size:0.82rem;">how the suggestion was calculated</div>
        ${mathTable}
      </div>
      <div>
        <div style="font-weight:600;margin-bottom:4px;font-size:0.82rem;">vendor breakdown — prior month (${pmKey || '—'}) transactions on this account</div>
        ${vendorTable}
      </div>
      <div>
        <div style="font-weight:600;margin-bottom:4px;font-size:0.82rem;">recent monthly history (last ${recent.length} of ${sortedMonths.length} active months)</div>
        ${historyHtml}
        <div style="margin-top:4px;color:var(--gray-500);font-size:0.68rem;">Only months with non-zero activity are shown. Empty months were skipped in the active-month average.</div>
      </div>
    </div>`;
}

async function updateAccruedLiabItem(input) {
  const part = input.dataset.alPart;
  const index = parseInt(input.dataset.alIndex);
  const amount = parseFloat(input.value) || 0;
  const month = accruedLiabAnalysis?.month;
  if (!month) return;

  try {
    await fetch(`/api/admin/clients/${selectedClientId}/accrued-liabilities/runs/${month}/items`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ part, index, accrualAmount: amount }),
    });
    // Update local state
    if (part === 'A' && accruedLiabAnalysis.partA?.accounts[index]) {
      accruedLiabAnalysis.partA.accounts[index].accrualAmount = amount;
    } else if (part === 'B' && accruedLiabAnalysis.partB?.transactions[index]) {
      accruedLiabAnalysis.partB.transactions[index].accrualAmount = amount;
    }
  } catch (e) { console.error('update failed:', e); }
}

async function toggleAccruedLiabDismiss(part, index, dismiss) {
  const month = accruedLiabAnalysis?.month;
  if (!month) return;

  try {
    await fetch(`/api/admin/clients/${selectedClientId}/accrued-liabilities/runs/${month}/items`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ part, index, status: dismiss ? 'dismissed' : 'pending' }),
    });
    // Update local + re-render
    if (part === 'A' && accruedLiabAnalysis.partA?.accounts[index]) {
      accruedLiabAnalysis.partA.accounts[index].status = dismiss ? 'dismissed' : 'pending';
    } else if (part === 'B' && accruedLiabAnalysis.partB?.transactions[index]) {
      accruedLiabAnalysis.partB.transactions[index].status = dismiss ? 'dismissed' : 'pending';
    }
    const panel = document.getElementById('accrued-liab-panel');
    if (panel) renderAccruedLiabResults(panel, accruedLiabAnalysis);
  } catch (e) { alert('update failed: ' + e.message); }
}

function showAccrualPostModal() {
  if (!accruedLiabAnalysis) return;
  const data = accruedLiabAnalysis;
  const month = data.month;
  const alAccount = accruedLiabState.accruedLiabilitiesAccount;

  // Gather non-dismissed items with amounts > 0
  const aItems = (data.partA?.accounts || []).filter(a => a.status !== 'dismissed' && Number(a.accrualAmount) > 0);
  const bItems = (data.partB?.transactions || []).filter(t => t.status !== 'dismissed' && Number(t.accrualAmount) > 0);
  // Exclude Part B items that overlap with Part A
  const bFiltered = bItems.filter(t => !t.overlapWithPartA || !aItems.some(a => a.accountId === t.accountId));
  const allItems = [...aItems.map(a => ({ accountName: a.accountName, amount: Number(a.accrualAmount), source: 'A' })),
                    ...bFiltered.map(t => ({ accountName: t.accountName, amount: Number(t.accrualAmount), source: 'B' }))];
  const grandTotal = Math.round(allItems.reduce((s, i) => s + i.amount, 0) * 100) / 100;

  // Build JE lines table
  let linesHtml = '';
  for (const item of allItems) {
    linesHtml += `<tr style="border-bottom:1px solid var(--gray-200);">
      <td style="padding:6px 8px;">${item.accountName}</td>
      <td style="padding:6px 8px;text-align:center;">Dr</td>
      <td style="padding:6px 8px;text-align:right;">${fmtMoney(item.amount)}</td>
      <td style="padding:6px 8px;text-align:right;">—</td>
      <td style="padding:6px 8px;font-size:0.72rem;color:var(--gray-500);">Part ${item.source}</td>
    </tr>`;
  }
  // Credit line
  linesHtml += `<tr style="border-bottom:1px solid var(--gray-200);font-weight:600;">
    <td style="padding:6px 8px;">${alAccount?.name || 'Accrued Liabilities'}</td>
    <td style="padding:6px 8px;text-align:center;">Cr</td>
    <td style="padding:6px 8px;text-align:right;">—</td>
    <td style="padding:6px 8px;text-align:right;">${fmtMoney(grandTotal)}</td>
    <td style="padding:6px 8px;"></td>
  </tr>`;

  // Parse period for dates
  const [y, m] = month.split('-').map(Number);
  const lastDay = new Date(y, m, 0);
  const accrualDate = lastDay.toISOString().split('T')[0];
  const reversalDate = new Date(y, m, 1).toISOString().split('T')[0];

  const overlay = document.createElement('div');
  overlay.id = 'accrual-post-modal';
  overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.4);display:flex;align-items:center;justify-content:center;z-index:1000;';
  overlay.innerHTML = `
    <div style="background:var(--white);border-radius:12px;padding:28px;width:100%;max-width:680px;max-height:80vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,0.2);">
      <h3 style="margin:0 0 4px;font-size:1.05rem;">accrual journal entry — ${formatPeriodLabel(month)}</h3>
      <p style="margin:0 0 16px;font-size:0.82rem;color:var(--gray-500);">This will create two entries in QBO:</p>

      <div style="margin-bottom:16px;">
        <div style="font-size:0.82rem;font-weight:600;margin-bottom:6px;">1. Accrual JE — dated ${accrualDate}</div>
        <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
          <thead>
            <tr style="border-bottom:2px solid var(--gray-300);text-align:left;">
              <th style="padding:6px 8px;">account</th>
              <th style="padding:6px 8px;text-align:center;">type</th>
              <th style="padding:6px 8px;text-align:right;">debit</th>
              <th style="padding:6px 8px;text-align:right;">credit</th>
              <th style="padding:6px 8px;">source</th>
            </tr>
          </thead>
          <tbody>${linesHtml}</tbody>
          <tfoot>
            <tr style="border-top:2px solid var(--gray-400);font-weight:700;">
              <td style="padding:6px 8px;">total</td>
              <td></td>
              <td style="padding:6px 8px;text-align:right;">${fmtMoney(grandTotal)}</td>
              <td style="padding:6px 8px;text-align:right;">${fmtMoney(grandTotal)}</td>
              <td></td>
            </tr>
          </tfoot>
        </table>
      </div>

      <div style="margin-bottom:20px;font-size:0.82rem;color:var(--gray-600);">
        <strong>2. Auto-reversal JE</strong> — dated ${reversalDate} (same lines, debits/credits reversed)
      </div>

      <div style="display:flex;justify-content:flex-end;gap:10px;">
        <button class="btn btn-secondary" onclick="document.getElementById('accrual-post-modal').remove()" style="padding:8px 18px;">cancel</button>
        <button class="btn btn-primary" onclick="postAccruedLiabilities()" style="padding:8px 18px;">post to QBO</button>
      </div>
    </div>`;
  overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
  document.body.appendChild(overlay);
}

async function postAccruedLiabilities() {
  if (!selectedClientId || !accruedLiabAnalysis?.month) return;
  const month = accruedLiabAnalysis.month;

  // Close modal if open
  const modal = document.getElementById('accrual-post-modal');

  // Find the post button in modal or step card
  const postBtn = modal?.querySelector('.btn-primary') || document.querySelector('[data-step-action="post-accrued-liab"]');
  if (postBtn) { postBtn.textContent = 'posting…'; postBtn.disabled = true; }

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/accrued-liabilities/post`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ month }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'post failed');

    if (modal) modal.remove();
    alert(`Posted accrual JE for ${month}\n\nAccrual: ${fmtMoney(data.accrualJE.totalAmount)} (JE #${data.accrualJE.journalEntryId})\nReversal dated ${data.reversalJE.date} (JE #${data.reversalJE.journalEntryId})`);
    await loadClientAccruedLiab();
    currentClosePeriod = null;
    renderCloseSteps();
  } catch (e) {
    alert('post failed: ' + e.message);
    if (postBtn) { postBtn.textContent = 'post to QBO'; postBtn.disabled = false; }
  }
}

// Wire settings save button
document.getElementById('btn-save-accrued-liab-settings').addEventListener('click', saveAccruedLiabSettings);

// ========================================
// SHAREHOLDER-PAID INVOICES MODULE
// ========================================

async function loadClientShareholderInvoices() {
  if (!selectedClientId) return;
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/shareholder-invoices`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    shiState = {
      shareholderLoanAccount: data.shareholderLoanAccount || null,
      invoices: data.invoices || [],
    };
  } catch (e) {
    console.error('Failed to load shareholder invoices:', e);
    shiState = { shareholderLoanAccount: null, invoices: [] };
  }
  renderShiSettings();
}

function renderShiSettings() {
  const acctSelect = document.getElementById('shi-account-select');
  if (!acctSelect) return;
  const savedId = shiState.shareholderLoanAccount?.id;
  const eligible = qboAccounts.filter(a =>
    ['Other Current Liability', 'Long Term Liability', 'Equity', 'Other Liability'].includes(a.type) ||
    (a.name || '').toLowerCase().includes('shareholder') ||
    (a.name || '').toLowerCase().includes('loan') ||
    (a.name || '').toLowerCase().includes('due to')
  );
  if (savedId && !eligible.find(a => a.id === savedId)) {
    const saved = qboAccounts.find(a => a.id === savedId);
    if (saved) eligible.unshift(saved);
    else eligible.unshift({ id: savedId, name: shiState.shareholderLoanAccount.name || '(saved)', type: '?' });
  }
  acctSelect.innerHTML = '<option value="">select shareholder loan account…</option>' +
    eligible.map(a => acctOption(a, savedId)).join('');
}

async function saveShiSettings() {
  const select = document.getElementById('shi-account-select');
  try {
    await fetch(`/api/admin/clients/${selectedClientId}/shareholder-invoices/settings`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({
        shareholderLoanAccount: select.value
          ? { id: select.value, name: select.options[select.selectedIndex]?.dataset.name || '', type: select.options[select.selectedIndex]?.dataset.type || '' }
          : null,
      }),
    });
    showSettingsSaved('btn-save-shi-settings');
    await loadClientShareholderInvoices();
    renderCloseSteps();
  } catch (e) { alert('save failed: ' + e.message); }
}

document.getElementById('btn-save-shi-settings').addEventListener('click', saveShiSettings);

async function handleShiUpload(input) {
  if (!input.files?.length || !selectedClientId) return;
  const file = input.files[0];
  input.value = ''; // reset so same file can be re-selected

  const panel = document.getElementById('shi-panel');
  if (panel) {
    panel.innerHTML = `
      <div style="padding:16px;background:var(--blue-50);border-radius:8px;font-size:0.85rem;color:var(--gray-700);">
        <strong>analyzing ${file.name}…</strong><br>
        uploading and sending to AI for review. this may take 10–20 seconds.
      </div>`;
  }

  try {
    const formData = new FormData();
    formData.append('invoice', file);

    const res = await fetch(`/api/admin/clients/${selectedClientId}/shareholder-invoices/upload`, {
      method: 'POST',
      headers: { 'Authorization': getAuth() },
      body: formData,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'upload failed');

    await loadClientShareholderInvoices();
    renderCloseSteps();
    // After re-render, show the review panel with all pending invoices
    setTimeout(() => renderShiPanel(), 100);
  } catch (e) {
    if (panel) panel.innerHTML = `<div style="padding:12px;background:var(--red-50);border-radius:8px;color:var(--red-600);">upload failed: ${e.message}</div>`;
  }
}

function renderShiPanel() {
  const panel = document.getElementById('shi-panel');
  if (!panel) return;

  const month = currentClosePeriod?.month;
  const pending = shiState.invoices.filter(i => i.status === 'pending' && i.closeMonth === month);
  const posted = shiState.invoices.filter(i => i.status === 'posted' && i.closeMonth === month);
  if (pending.length === 0 && posted.length === 0) {
    panel.innerHTML = '';
    return;
  }

  let html = '';

  if (pending.length > 0) {
    html += pending.map(inv => renderShiInvoiceCard(inv)).join('');
  }

  if (posted.length > 0) {
    html += `<details style="margin-top:8px;"><summary style="font-size:0.82rem;color:var(--gray-500);cursor:pointer;">${posted.length} posted invoice${posted.length !== 1 ? 's' : ''}</summary>`;
    html += posted.map(inv => renderShiInvoiceCard(inv)).join('');
    html += `</details>`;
  }

  panel.innerHTML = html;

  // Initialize totals and wire up live amount updates
  panel.querySelectorAll('.shi-invoice-card').forEach(card => {
    shiUpdateTotals(card);
    card.querySelectorAll('.shi-amt-input').forEach(inp => {
      inp.addEventListener('input', () => shiUpdateTotals(card));
    });
  });
}

function renderShiInvoiceCard(inv) {
  const a = inv.analysis || {};
  const isPosted = inv.status === 'posted';
  const borderColor = isPosted ? 'var(--green-500)' : 'var(--blue-500)';
  const bgColor = isPosted ? '#f0fdf4' : '#eff6ff';

  const confidenceColors = {
    high: { bg: '#dcfce7', fg: '#166534' },
    medium: { bg: '#fef9c3', fg: '#854d0e' },
    low: { bg: '#f3f4f6', fg: '#6b7280' },
  };
  const cc = confidenceColors[a.confidence] || confidenceColors.low;

  const eligibleAccounts = qboAccounts.filter(ac =>
    ['Expense', 'Other Expense', 'Cost of Goods Sold', 'Fixed Asset', 'Other Current Asset', 'Other Current Liability', 'Long Term Liability', 'Accounts Payable'].includes(ac.type)
  );

  // Fuzzy-match AI suggested account name to actual QBO accounts
  function matchAccount(line) {
    // Try suggestedAccount first (AI should return exact name from chart), then category
    const candidates = [line.suggestedAccount, line.suggestedCategory, line.description].filter(Boolean);
    for (const raw of candidates) {
      const suggested = raw.toLowerCase().trim();
      if (!suggested) continue;

      // Exact name match
      const exact = eligibleAccounts.find(ac => (ac.name || '').toLowerCase() === suggested);
      if (exact) return exact.id;

      // Account number match (if AI returned a number)
      if (/^\d+$/.test(suggested.split(/\s/)[0])) {
        const numMatch = eligibleAccounts.find(ac => ac.acctNum === suggested.split(/\s/)[0]);
        if (numMatch) return numMatch.id;
      }

      // Exact substring match (account name contains suggestion or vice versa)
      const contains = eligibleAccounts.find(ac => {
        const acName = (ac.name || '').toLowerCase();
        return acName.includes(suggested) || suggested.includes(acName);
      });
      if (contains) return contains.id;
    }

    // Word overlap scoring across all candidates
    const allWords = candidates.join(' ').toLowerCase().split(/[\s/&,()-]+/).filter(w => w.length > 2);
    // Remove very common words that cause false matches
    const stopWords = new Set(['and', 'the', 'for', 'from', 'with', 'fees', 'fee', 'expense', 'expenses', 'other', 'cost', 'general']);
    const words = allWords.filter(w => !stopWords.has(w));
    let bestMatch = null, bestScore = 0;
    for (const ac of eligibleAccounts) {
      const acName = (ac.name || '').toLowerCase();
      const score = words.filter(w => acName.includes(w)).length;
      if (score > bestScore) { bestScore = score; bestMatch = ac; }
    }
    return bestScore > 0 ? bestMatch.id : '';
  }

  const linesHtml = (a.lines || []).map((line, idx) => {
    const matchedId = !isPosted ? matchAccount(line) : '';
    return `
    <div class="shi-je-line" data-line-idx="${idx}" style="display:flex;gap:6px;align-items:center;margin-bottom:4px;font-size:0.8rem;">
      <span style="flex:0 0 16px;color:var(--gray-400);font-size:0.72rem;">${idx + 1}.</span>
      ${!isPosted ? `
        <input type="text" class="shi-desc-input" value="${(line.description || '').replace(/"/g, '&quot;')}" style="flex:1;padding:3px 6px;border:1px solid var(--gray-300);border-radius:4px;font-size:0.78rem;">
        <input type="number" class="shi-amt-input" step="0.01" value="${Number(line.amount).toFixed(2)}" style="flex:0 0 90px;padding:3px 6px;border:1px solid var(--gray-300);border-radius:4px;font-size:0.78rem;text-align:right;">
        <select class="shi-acct-select" data-line-idx="${idx}" style="flex:0 0 220px;padding:3px 6px;border:1px solid var(--gray-300);border-radius:4px;font-size:0.78rem;">
          <option value="">select GL account…</option>
          ${eligibleAccounts.map(ac => acctOption(ac, matchedId)).join('')}
        </select>
        <button onclick="shiRemoveLine('${inv.id}',${idx})" style="flex:0 0 24px;background:none;border:none;color:var(--red-400);cursor:pointer;font-size:1rem;padding:0;" title="remove line">×</button>
      ` : `
        <span style="flex:1;">${line.description}</span>
        <span style="flex:0 0 80px;text-align:right;font-weight:500;">$${Number(line.amount).toFixed(2)}</span>
      `}
    </div>`;
  }).join('');

  const addLineBtn = !isPosted ? `
    <div style="margin-top:4px;">
      <button onclick="shiAddLine('${inv.id}')" style="background:none;border:1px dashed var(--gray-300);border-radius:4px;padding:4px 12px;font-size:0.76rem;color:var(--gray-500);cursor:pointer;">+ add line</button>
    </div>
  ` : '';

  const actions = !isPosted ? `
    <div style="display:flex;gap:8px;margin-top:8px;">
      <button class="btn btn-primary" onclick="postShiInvoice('${inv.id}')" style="font-size:0.78rem;padding:5px 12px;">post expense</button>
      <button class="btn btn-secondary" onclick="deleteShiInvoice('${inv.id}')" style="font-size:0.78rem;padding:5px 10px;color:var(--red-600);">delete</button>
    </div>
  ` : `
    <div style="margin-top:6px;font-size:0.78rem;color:var(--green-600);">
      ✓ posted ${inv.journalEntry?.date || ''} — $${Number(inv.journalEntry?.totalAmount || 0).toFixed(2)}${inv.journalEntry?.jeId ? ` (Expense #${inv.journalEntry.jeId})` : ''}
    </div>
  `;

  return `
    <div class="shi-invoice-card" data-invoice-id="${inv.id}" style="border-left:3px solid ${borderColor};background:${bgColor};padding:10px 12px;margin-bottom:6px;border-radius:4px;font-size:0.82rem;">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;">
        <div>
          <strong>${a.vendor || 'Unknown Vendor'}</strong>
          <span style="margin-left:8px;font-weight:600;">$${Number(a.totalAmount || 0).toFixed(2)}</span>
          <span style="margin-left:8px;color:var(--gray-500);">${a.invoiceDate || '—'}</span>
        </div>
        <div style="display:flex;gap:6px;align-items:center;">
          <span style="padding:2px 8px;border-radius:10px;font-size:0.72rem;font-weight:500;background:${cc.bg};color:${cc.fg};">${a.confidence || 'low'} confidence</span>
          <span style="font-size:0.72rem;color:var(--gray-400);" title="${inv.fileName}">📄 ${inv.fileName?.length > 20 ? inv.fileName.slice(0, 20) + '…' : inv.fileName}</span>
        </div>
      </div>
      <div style="margin-top:4px;color:var(--gray-600);">${a.description || ''}</div>
      ${a.reasoning ? `<div style="margin-top:4px;font-style:italic;color:var(--gray-500);font-size:0.78rem;">${a.reasoning}</div>` : ''}
      <div style="margin-top:8px;padding-top:6px;border-top:1px solid rgba(0,0,0,0.06);">
        ${linesHtml}
        ${addLineBtn}
        <div class="shi-totals" data-inv-total="${Number(a.totalAmount || 0).toFixed(2)}" style="display:flex;justify-content:flex-end;gap:16px;margin-top:6px;font-size:0.78rem;font-weight:600;">
        </div>
        ${!isPosted ? `
          <div style="display:flex;align-items:center;gap:8px;margin-top:6px;font-size:0.8rem;">
            <label style="color:var(--gray-500);">date:</label>
            <input type="date" class="shi-date-input" value="${a.invoiceDate || ''}" style="padding:3px 6px;border:1px solid var(--gray-300);border-radius:4px;font-size:0.78rem;">
          </div>
        ` : ''}
      </div>
      ${actions}
    </div>`;
}

async function postShiInvoice(invoiceId) {
  const card = document.querySelector(`.shi-invoice-card[data-invoice-id="${invoiceId}"]`);
  if (!card) return;

  // Gather line data from the editable inputs in the card
  const lineEls = card.querySelectorAll('.shi-je-line');
  const inv = shiState.invoices.find(i => i.id === invoiceId);
  if (!inv) return;

  const lines = [];
  for (let i = 0; i < lineEls.length; i++) {
    const lineEl = lineEls[i];
    const sel = lineEl.querySelector('.shi-acct-select');
    const descInput = lineEl.querySelector('.shi-desc-input');
    const amtInput = lineEl.querySelector('.shi-amt-input');
    if (!sel?.value) {
      alert(`Please select a GL account for line ${i + 1}`);
      return;
    }
    const amount = parseFloat(amtInput?.value) || 0;
    if (amount <= 0) {
      alert(`Line ${i + 1} has an invalid amount`);
      return;
    }
    lines.push({
      accountId: sel.value,
      accountName: sel.options[sel.selectedIndex]?.dataset.name || '',
      description: descInput?.value || '',
      amount,
    });
  }

  // Read the editable date from the card
  const dateInput = card.querySelector('.shi-date-input');
  const jeDate = dateInput?.value || inv.analysis.invoiceDate;

  const totalAmt = lines.reduce((s, l) => s + l.amount, 0);
  if (!confirm(`Post expense for $${totalAmt.toFixed(2)} dated ${jeDate}?\n\nLines: ${lines.map(l => `${l.accountName} $${l.amount.toFixed(2)}`).join(', ')}\nPaid from: ${shiState.shareholderLoanAccount?.name || 'Shareholder Loan'} $${totalAmt.toFixed(2)}`)) return;

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/shareholder-invoices/${invoiceId}/post`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ lines, date: jeDate }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'post failed');

    alert(`Expense posted!${data.journalEntry?.jeId ? ` Expense #${data.journalEntry.jeId}` : ''}\n$${data.journalEntry?.totalAmount?.toFixed(2)}`);
    await loadClientShareholderInvoices();
    currentClosePeriod = null;
    renderCloseSteps();
    setTimeout(() => renderShiPanel(), 100);
  } catch (e) { alert('post failed: ' + e.message); }
}

async function deleteShiInvoice(invoiceId) {
  if (!confirm('Delete this invoice? This cannot be undone.')) return;
  try {
    await fetch(`/api/admin/clients/${selectedClientId}/shareholder-invoices/${invoiceId}`, {
      method: 'DELETE',
      headers: { 'Authorization': getAuth() },
    });
    await loadClientShareholderInvoices();
    renderCloseSteps();
    setTimeout(() => renderShiPanel(), 100);
  } catch (e) { alert('delete failed: ' + e.message); }
}

function shiUpdateTotals(card) {
  const totalsEl = card.querySelector('.shi-totals');
  if (!totalsEl) return;
  const invTotal = parseFloat(totalsEl.dataset.invTotal) || 0;
  const amtInputs = card.querySelectorAll('.shi-amt-input');
  let lineTotal = 0;
  amtInputs.forEach(inp => { lineTotal += parseFloat(inp.value) || 0; });
  lineTotal = Math.round(lineTotal * 100) / 100;
  const diff = Math.round((invTotal - lineTotal) * 100) / 100;
  const balanced = Math.abs(diff) < 0.01;
  totalsEl.innerHTML = `
    <span style="color:var(--gray-500);">lines total: $${lineTotal.toFixed(2)}</span>
    <span style="color:var(--gray-500);">invoice total: $${invTotal.toFixed(2)}</span>
    ${balanced
      ? '<span style="color:var(--green-600);">✓ balanced</span>'
      : `<span style="color:var(--red-600);">⚠ difference: $${diff.toFixed(2)}</span>`}`;
}

function shiAddLine(invoiceId) {
  const inv = shiState.invoices.find(i => i.id === invoiceId);
  if (!inv || !inv.analysis) return;
  inv.analysis.lines.push({ description: '', amount: 0, suggestedAccount: '' });
  renderShiPanel();
}

function shiRemoveLine(invoiceId, lineIdx) {
  const inv = shiState.invoices.find(i => i.id === invoiceId);
  if (!inv || !inv.analysis?.lines) return;
  if (inv.analysis.lines.length <= 1) {
    alert('Cannot remove the last line');
    return;
  }
  inv.analysis.lines.splice(lineIdx, 1);
  renderShiPanel();
}

// ========================================
// PREPAID EXPENSE SCANNER (Part B)
// ========================================

let lastScanResults = null;

async function runPrepaidScan() {
  if (!selectedClientId) return;
  const panel = document.getElementById('prepaid-scan-panel');
  if (!panel) return;

  // Show loading state
  const scanBtn = document.querySelector('[data-step-action="scan-prepaids"]');
  if (scanBtn) { scanBtn.textContent = 'scanning…'; scanBtn.disabled = true; }
  panel.innerHTML = `
    <div style="padding:16px;background:var(--blue-50);border-radius:8px;font-size:0.85rem;color:var(--gray-700);">
      <strong>scanning expense transactions…</strong><br>
      pulling transactions from QBO, checking for invoice attachments, and sending each to Claude for review. this may take 30–60 seconds depending on the number of transactions.
    </div>`;

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/scan`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ threshold: prepaidState.scanThreshold || 500, month: currentClosePeriod?.month || undefined }),
    });
    const data = await res.json();
    if (!res.ok) { throw new Error(data.error || 'scan failed'); }
    lastScanResults = data;
    renderScanResults(panel, data);
  } catch (e) {
    panel.innerHTML = `<div style="padding:12px;background:var(--red-50);border-radius:8px;color:var(--red-600);">scan failed: ${e.message}</div>`;
  }
  if (scanBtn) { scanBtn.textContent = 'scan expenses'; scanBtn.disabled = false; }
}

function renderScanResults(panel, data) {
  if (!data.results || data.results.length === 0) {
    panel.innerHTML = `
      <div style="padding:12px;background:var(--gray-50);border-radius:8px;font-size:0.85rem;">
        <strong>no transactions found</strong> above $${data.threshold} for ${data.month}.
      </div>`;
    return;
  }

  const flagged = data.results.filter(r => r.claudeReview?.isPrepaid === true);
  const clean = data.results.filter(r => r.claudeReview?.isPrepaid === false);
  const errors = data.results.filter(r => r.error || r.claudeReview?.error);

  let html = `
    <div style="padding:12px;background:var(--gray-50);border-radius:8px;margin-bottom:8px;font-size:0.85rem;">
      <strong>scan complete for ${data.month}</strong> —
      ${data.candidateCount} transaction${data.candidateCount !== 1 ? 's' : ''} scanned (≥ $${data.threshold}),
      <span style="color:${flagged.length > 0 ? 'var(--red-600)' : 'var(--green-600)'};">
        ${flagged.length} potential prepaid${flagged.length !== 1 ? 's' : ''} detected
      </span>
    </div>`;

  // Show flagged items first (these need action)
  if (flagged.length > 0) {
    html += `<div style="margin-bottom:8px;"><strong style="color:var(--red-600);font-size:0.85rem;">⚠ potential prepaids found:</strong></div>`;
    html += flagged.map(r => renderScanResultCard(r, data, 'flagged')).join('');
  }

  // Show clean items (collapsed)
  if (clean.length > 0) {
    html += `
      <details style="margin-top:8px;">
        <summary style="cursor:pointer;font-size:0.85rem;color:var(--gray-600);padding:4px 0;">
          ✓ ${clean.length} clean transaction${clean.length !== 1 ? 's' : ''} (click to expand)
        </summary>
        <div style="margin-top:4px;">
          ${clean.map(r => renderScanResultCard(r, data, 'clean')).join('')}
        </div>
      </details>`;
  }

  // Errors
  if (errors.length > 0) {
    html += `
      <details style="margin-top:8px;">
        <summary style="cursor:pointer;font-size:0.85rem;color:var(--gray-400);">
          ${errors.length} error${errors.length !== 1 ? 's' : ''}
        </summary>
        <div style="margin-top:4px;">${errors.map(r => `<div style="font-size:0.8rem;padding:4px;color:var(--gray-500);">${r.txn.vendor} — ${r.error || 'parse error'}</div>`).join('')}</div>
      </details>`;
  }

  panel.innerHTML = html;
}

function renderScanResultCard(result, scanData, type) {
  const t = result.txn;
  const r = result.claudeReview || {};
  const isFlagged = type === 'flagged';
  const borderColor = isFlagged ? 'var(--red-400)' : 'var(--green-400)';
  const bgColor = isFlagged ? '#fef2f2' : '#f0fdf4';

  const confidenceBadge = r.confidence
    ? `<span style="display:inline-block;padding:1px 6px;border-radius:4px;font-size:0.72rem;background:${r.confidence === 'high' ? '#dcfce7' : r.confidence === 'medium' ? '#fef9c3' : '#f3f4f6'};color:${r.confidence === 'high' ? '#166534' : r.confidence === 'medium' ? '#854d0e' : '#6b7280'};">${r.confidence} confidence</span>`
    : '';

  const attachBadge = result.hasAttachment
    ? `<span style="font-size:0.72rem;color:var(--blue-500);">📎 ${result.attachment?.fileName || 'attachment'}</span>`
    : `<span style="font-size:0.72rem;color:var(--gray-400);">no attachment</span>`;

  let actionHtml = '';
  if (isFlagged) {
    // "Accept as prepaid" button — adds to the schedule
    const encoded = encodeURIComponent(JSON.stringify({
      vendor: t.vendor,
      description: r.description || '',
      totalAmount: r.totalAmount || t.amount,
      prepaidAmount: r.prepaidAmount || r.totalAmount || t.amount,
      startDate: r.servicePeriodStart || t.date,
      endDate: r.servicePeriodEnd || '',
      expenseAccountId: t.accountId,
      expenseAccountName: t.accountName,
      sourceTxnId: t.id,
      sourceTxnType: t.type,
    }));
    actionHtml = `
      <div style="margin-top:6px;display:flex;gap:8px;">
        <button class="btn btn-primary" style="font-size:0.78rem;padding:4px 10px;" onclick="acceptScanResult(decodeURIComponent('${encoded}'))">add to prepaid schedule</button>
        <button class="btn btn-secondary" style="font-size:0.78rem;padding:4px 10px;" onclick="this.closest('.scan-result-card').style.display='none'">dismiss</button>
      </div>`;
  }

  return `
    <div class="scan-result-card" style="border-left:3px solid ${borderColor};background:${bgColor};padding:10px 12px;margin-bottom:6px;border-radius:4px;font-size:0.82rem;">
      <div style="display:flex;justify-content:space-between;align-items:start;">
        <div>
          <strong>${t.vendor}</strong> — ${fmtMoney(t.amount)} — ${t.date}
          ${t.accountName ? `<span style="color:var(--gray-500);margin-left:6px;">${t.accountName}</span>` : ''}
        </div>
        <div style="display:flex;gap:6px;align-items:center;">
          ${confidenceBadge} ${attachBadge}
        </div>
      </div>
      <div style="margin-top:4px;color:var(--gray-600);font-size:0.8rem;">${r.reasoning || ''}</div>
      ${r.servicePeriodStart && r.servicePeriodEnd ? `<div style="margin-top:2px;font-size:0.78rem;color:var(--gray-500);">service period: ${r.servicePeriodStart} → ${r.servicePeriodEnd}${r.prepaidAmount ? ` | prepaid portion: ${fmtMoney(r.prepaidAmount)}` : ''}</div>` : ''}
      ${actionHtml}
    </div>`;
}

async function acceptScanResult(encodedJson) {
  try {
    const data = JSON.parse(encodedJson);
    if (!data.endDate) {
      const endDate = prompt('Service period end date not detected.\nPlease enter the end date (YYYY-MM-DD):');
      if (!endDate) return;
      data.endDate = endDate;
    }
    const res = await fetch(`/api/admin/clients/${selectedClientId}/prepaid-expenses/scan/accept`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify(data),
    });
    const result = await res.json();
    if (!res.ok) { alert('failed: ' + (result.error || res.status)); return; }
    alert(`Added "${data.vendor}" to the prepaid schedule.\n\nIt will be amortized ${fmtMoney(data.prepaidAmount)} from ${data.startDate} to ${data.endDate}.`);
    await loadClientPrepaid();
    renderCloseSteps();
  } catch (e) { alert('error: ' + e.message); }
}

// ========================================
// ASSET CLASS MODAL FUNCTIONS
// ========================================

function populateClassDropdowns(cls) {
  const selects = {
    'class-account-expense': { types: ['Expense', 'Other Expense'], selected: cls?.expenseAccountId },
    'class-account-accum': { types: ['Fixed Asset', 'Other Asset', 'Other Current Asset'], selected: cls?.accumAccountId },
  };
  Object.entries(selects).forEach(([id, cfg]) => {
    const select = document.getElementById(id);
    select.innerHTML = '<option value="">select account...</option>';
    qboAccounts.filter(a => cfg.types.includes(a.type)).forEach(a => {
      const opt = document.createElement('option');
      opt.value = a.id;
      opt.textContent = `${a.name} (${a.type})`;
      opt.dataset.name = a.name;
      if (a.id === cfg.selected) opt.selected = true;
      select.appendChild(opt);
    });
  });
}

function toggleClassMethod() {
  const method = document.getElementById('class-method').value;
  document.getElementById('class-sl-group').style.display = method === 'straight-line' ? '' : 'none';
  document.getElementById('class-db-group').style.display = method === 'declining-balance' ? '' : 'none';
}

function openEditClass(id) {
  const cls = assetClasses.find(c => c.id === id);
  if (!cls) return;
  editingClassId = id;
  document.getElementById('asset-class-modal-title').textContent = `amortization policy: ${cls.glAccountName}`;
  document.getElementById('class-gl-name').value = cls.glAccountName || '';
  document.getElementById('class-method').value = cls.method || 'straight-line';
  document.getElementById('class-useful-life').value = cls.usefulLifeMonths || '';
  document.getElementById('class-declining-rate').value = cls.decliningRate ? (cls.decliningRate * 100) : '';
  document.getElementById('class-modal-error').style.display = 'none';

  // Show AI suggestion if available, with full recommended values + apply button
  const aiEl = document.getElementById('class-ai-suggestion');
  const applyBtn = document.getElementById('btn-apply-class-ai');
  if (cls.aiSuggestion) {
    const ai = cls.aiSuggestion;
    const methodLabel = ai.method === 'declining-balance'
      ? `declining balance @ ${((ai.decliningRate || 0) * 100).toFixed(0)}%/yr`
      : `straight-line, ${ai.usefulLifeMonths || '—'} months`;
    let html = `<strong>AI suggests: ${methodLabel}</strong>`;
    if (ai.ccaClass) html += ` — CCA Class ${ai.ccaClass} (${ai.ccaRate})`;
    if (ai.reasoning) html += `<br>${ai.reasoning}`;
    html += `<br><em style="color:var(--gray-500);">This is advisory only — click "apply AI suggestion" to copy these values into the form.</em>`;
    aiEl.innerHTML = html;
    aiEl.style.display = '';
    applyBtn.style.display = '';
  } else {
    aiEl.style.display = 'none';
    applyBtn.style.display = 'none';
  }

  populateClassDropdowns(cls);
  toggleClassMethod();
  document.getElementById('asset-class-modal').style.display = 'flex';
}

// Apply the saved AI suggestion to the form fields. Does NOT save — the user
// still has to click "save" to persist.
function applyClassAiSuggestion() {
  const cls = assetClasses.find(c => c.id === editingClassId);
  if (!cls || !cls.aiSuggestion) return;
  const ai = cls.aiSuggestion;
  const method = ai.method || 'straight-line';
  document.getElementById('class-method').value = method;
  toggleClassMethod();
  if (method === 'declining-balance' && ai.decliningRate) {
    document.getElementById('class-declining-rate').value = (ai.decliningRate * 100).toFixed(0);
    document.getElementById('class-useful-life').value = '';
  } else if (ai.usefulLifeMonths) {
    document.getElementById('class-useful-life').value = ai.usefulLifeMonths;
    document.getElementById('class-declining-rate').value = '';
  }
}

function closeClassModal() {
  document.getElementById('asset-class-modal').style.display = 'none';
  editingClassId = null;
}

async function saveClass() {
  const errorEl = document.getElementById('class-modal-error');
  errorEl.style.display = 'none';

  const method = document.getElementById('class-method').value;
  const usefulLifeMonths = document.getElementById('class-useful-life').value;
  const decliningRatePercent = document.getElementById('class-declining-rate').value;
  const expenseEl = document.getElementById('class-account-expense');
  const accumEl = document.getElementById('class-account-accum');

  if (method === 'straight-line' && !usefulLifeMonths) {
    errorEl.textContent = 'useful life is required for straight-line method';
    errorEl.style.display = ''; return;
  }
  if (method === 'declining-balance' && !decliningRatePercent) {
    errorEl.textContent = 'declining rate is required for declining balance method';
    errorEl.style.display = ''; return;
  }

  const body = {
    method,
    usefulLifeMonths: parseInt(usefulLifeMonths) || null,
    decliningRate: decliningRatePercent ? parseFloat(decliningRatePercent) / 100 : null,
    expenseAccountId: expenseEl.value || '',
    expenseAccountName: expenseEl.selectedOptions[0]?.dataset.name || '',
    accumAccountId: accumEl.value || '',
    accumAccountName: accumEl.selectedOptions[0]?.dataset.name || '',
  };

  try {
    const url = `/api/admin/clients/${selectedClientId}/fixed-assets/classes/${editingClassId}`;
    const res = await fetch(url, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (data.error) { errorEl.textContent = data.error; errorEl.style.display = ''; return; }
    closeClassModal();
    loadClientFixedAssets();
  } catch (e) { errorEl.textContent = 'failed to save class'; errorEl.style.display = ''; }
}

async function deleteClass(id) {
  const cls = assetClasses.find(c => c.id === id);
  const assetCount = allFixedAssets.filter(a => a.active && (a.assetAccountId === cls?.glAccountId || a.glAccountName === cls?.glAccountName)).length;
  if (!confirm(`delete amortization policy for "${cls?.glAccountName}"? ${assetCount} asset${assetCount !== 1 ? 's' : ''} will lose their policy.`)) return;
  try {
    await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/classes/${id}`, {
      method: 'DELETE', headers: { 'Authorization': getAuth() },
    });
    loadClientFixedAssets();
  } catch (e) { alert('failed to delete class'); }
}

async function suggestClassAmortization() {
  const glName = document.getElementById('class-gl-name').value.trim();
  const btn = document.getElementById('btn-suggest-class-amort');
  const aiEl = document.getElementById('class-ai-suggestion');

  if (!glName) { alert('no GL account name to suggest for'); return; }

  btn.textContent = 'thinking...';
  btn.disabled = true;
  aiEl.style.display = 'none';

  try {
    const res = await fetch('/api/admin/fixed-assets/suggest-amortization', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ assetName: glName, originalCost: 10000 }),
    });
    const data = await res.json();

    if (data.usefulLifeMonths) {
      // Apply suggestion to form fields
      const suggestedMethod = data.method || 'straight-line';
      document.getElementById('class-method').value = suggestedMethod;
      toggleClassMethod();

      if (suggestedMethod === 'declining-balance' && data.decliningRate) {
        document.getElementById('class-declining-rate').value = (data.decliningRate * 100).toFixed(0);
      } else {
        document.getElementById('class-useful-life').value = data.usefulLifeMonths;
      }
      let html = `<strong>AI: ${suggestedMethod === 'declining-balance' ? 'declining balance' : 'straight-line'}</strong>`;
      if (data.ccaClass) html += ` — CCA Class ${data.ccaClass} (${data.ccaRate})`;
      if (data.reasoning) html += `<br>${data.reasoning}`;
      aiEl.innerHTML = html;
      aiEl.style.display = '';
    }
  } catch (e) {
    aiEl.textContent = 'failed to get suggestion';
    aiEl.style.display = '';
  }
  btn.textContent = 'AI suggest policy';
  btn.disabled = false;
}

// Class modal event listeners
document.getElementById('class-method').addEventListener('change', toggleClassMethod);
document.getElementById('class-modal-save').addEventListener('click', saveClass);
document.getElementById('class-modal-cancel').addEventListener('click', closeClassModal);
document.getElementById('asset-class-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeClassModal(); });
document.getElementById('btn-suggest-class-amort').addEventListener('click', suggestClassAmortization);
document.getElementById('btn-apply-class-ai').addEventListener('click', applyClassAiSuggestion);

// Prepaid expenses settings wiring
document.getElementById('btn-save-prepaid-settings').addEventListener('click', savePrepaidSettings);
document.getElementById('btn-add-prepaid').addEventListener('click', openAddPrepaid);
document.getElementById('btn-download-prepaid-template').addEventListener('click', downloadPrepaidTemplate);
document.getElementById('prepaid-import-file').addEventListener('change', (e) => {
  if (e.target.files && e.target.files[0]) {
    importPrepaidExcel(e.target.files[0]);
    e.target.value = '';
  }
});
document.getElementById('prepaid-modal-save').addEventListener('click', savePrepaid);
document.getElementById('prepaid-modal-cancel').addEventListener('click', closePrepaidModal);
document.getElementById('prepaid-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closePrepaidModal(); });

// Temporary: QBO attachment pipeline smoke test
document.getElementById('btn-test-attachments').addEventListener('click', async () => {
  if (!selectedClientId) {
    alert('select a client first');
    return;
  }
  const btn = document.getElementById('btn-test-attachments');
  const output = document.getElementById('attachment-test-output');
  btn.disabled = true;
  btn.textContent = 'testing…';
  output.style.display = 'block';
  output.textContent = 'scanning recent bills/purchases for an attachment…';
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/qbo/attachments/test`);
    const data = await res.json();
    output.textContent = JSON.stringify(data, null, 2);
  } catch (e) {
    output.textContent = 'error: ' + e.message;
  } finally {
    btn.disabled = false;
    btn.textContent = 'test QBO attachments';
  }
});

// Silent sync-and-import path used right after a fresh QBO OAuth connect.
// Skips the review modal entirely — hits the sync endpoint, imports every
// account it returns via the shared importAssetsFromQboAccounts helper, then
// refreshes the class summary. Shows a brief status message on the
// amortization-status banner so the user knows something happened.
async function autoSyncAndImportFromQbo() {
  if (!selectedClientId) return;
  const statusEl = document.getElementById('amortization-status');
  statusEl.textContent = 'syncing fixed assets from QuickBooks...';
  statusEl.style.display = '';
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/sync-qbo`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
    });
    const data = await res.json();
    if (data.error) {
      statusEl.textContent = `sync failed: ${data.error}`;
      return;
    }
    const accounts = data.accounts || [];
    if (accounts.length === 0) {
      statusEl.textContent = 'no new fixed asset accounts found in QBO';
      await loadClientFixedAssets();
      return;
    }
    const newClassIds = await importAssetsFromQboAccounts(accounts);
    await loadClientFixedAssets();
    statusEl.textContent = `imported ${accounts.length} fixed asset${accounts.length !== 1 ? 's' : ''} from QuickBooks`;
    if (newClassIds.length > 0) {
      runAISuggestionsForClasses(newClassIds);
    }
  } catch (e) {
    console.error('auto sync failed:', e);
    statusEl.textContent = 'sync failed — try clicking "sync from QBO" manually';
  }
}

// QBO Sync
async function syncFromQBO() {
  if (!selectedClientId) return;
  const loadingEl = document.getElementById('qbo-sync-loading');
  const listEl = document.getElementById('qbo-sync-list');
  const importBtn = document.getElementById('qbo-sync-import');
  const errorEl = document.getElementById('qbo-sync-error');

  loadingEl.style.display = '';
  listEl.innerHTML = '';
  importBtn.style.display = 'none';
  errorEl.style.display = 'none';
  document.getElementById('qbo-sync-modal').style.display = 'flex';

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/sync-qbo`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
    });
    const data = await res.json();
    loadingEl.style.display = 'none';

    if (data.error) { errorEl.textContent = data.error; errorEl.style.display = ''; return; }

    if (!data.accounts || data.accounts.length === 0) {
      listEl.innerHTML = '<p style="color:var(--gray-500);text-align:center;padding:20px;">no new fixed asset accounts found in QBO, or all have already been imported.</p>';
      return;
    }

    listEl.innerHTML = data.accounts.map(a => `
      <label class="qbo-sync-item">
        <input type="checkbox" class="qbo-sync-check" data-account='${JSON.stringify(a).replace(/'/g, "&#39;")}' checked>
        <div class="qbo-sync-item-info">
          <div class="qbo-sync-item-name">${a.name}</div>
          <div class="qbo-sync-item-meta">
            ${a.glAccountName && a.glAccountName !== a.name ? `GL: ${a.glAccountName} — ` : ''}
            $${(a.originalCost || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            ${a.txnDate ? ` — ${a.txnDate}` : ''}
            ${a.vendorName ? ` — ${a.vendorName}` : ''}
          </div>
        </div>
      </label>
    `).join('');
    importBtn.style.display = '';
  } catch (e) {
    loadingEl.style.display = 'none';
    errorEl.textContent = 'failed to fetch QBO data';
    errorEl.style.display = '';
  }
}

async function importSelectedAssets() {
  const checkboxes = document.querySelectorAll('.qbo-sync-check:checked');
  const errorEl = document.getElementById('qbo-sync-error');
  const importBtn = document.getElementById('qbo-sync-import');
  errorEl.style.display = 'none';
  importBtn.textContent = 'importing...';
  importBtn.disabled = true;

  const accounts = Array.from(checkboxes).map(cb => JSON.parse(cb.dataset.account));
  const newClassIds = await importAssetsFromQboAccounts(accounts);

  // Close modal and show assets immediately
  document.getElementById('qbo-sync-modal').style.display = 'none';
  importBtn.textContent = 'import selected';
  importBtn.disabled = false;
  await loadClientFixedAssets();

  if (newClassIds.length > 0) {
    runAISuggestionsForClasses(newClassIds);
  }
}

/**
 * Import a list of QBO accounts (as returned by /sync-qbo) into this client's
 * fixed-asset roster. Creates the asset records, auto-creates any missing
 * asset classes, and returns the IDs of newly-created classes so the caller
 * can run AI suggestions on them.
 *
 * Pure function — no DOM dependency — so it can be called from the manual
 * modal flow OR from the pre-amortization auto-sync.
 */
async function importAssetsFromQboAccounts(accounts) {
  const importedAssets = [];

  for (const account of accounts) {
    const body = {
      name: account.name,
      description: account.description || '',
      glAccountName: account.glAccountName || '',
      originalCost: account.originalCost || account.currentBalance || 0,
      usefulLifeMonths: 60, // default — will be updated by AI suggestion
      salvageValue: 0,
      acquisitionDate: account.txnDate || new Date().toISOString().split('T')[0],
      assetAccountId: account.qboAccountId,
      assetAccountName: account.glAccountName || account.name,
      expenseAccountId: account.suggestedExpenseAccountId || '',
      expenseAccountName: account.suggestedExpenseAccountName || '',
      accumAccountId: account.suggestedAccumAccountId || '',
      accumAccountName: account.suggestedAccumAccountName || '',
      qboAccountId: account.qboAccountId,
      txnKey: account.txnKey || null,
      vendorName: account.vendorName || '',
      fromSync: true, // skip strict validation
    };

    try {
      const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (data.asset) importedAssets.push(data.asset);
    } catch (e) { console.error('Failed to import asset:', e); }
  }

  // Auto-create asset classes for each unique GL account, carrying over expense/accum from sync
  const uniqueGLAccounts = new Map();
  for (const asset of importedAssets) {
    if (asset.assetAccountId && !uniqueGLAccounts.has(asset.assetAccountId)) {
      uniqueGLAccounts.set(asset.assetAccountId, {
        glAccountId: asset.assetAccountId,
        glAccountName: asset.glAccountName || asset.assetAccountName || asset.name,
        expenseAccountId: asset.expenseAccountId || '',
        expenseAccountName: asset.expenseAccountName || '',
        accumAccountId: asset.accumAccountId || '',
        accumAccountName: asset.accumAccountName || '',
      });
    }
  }

  // Create classes for any GL accounts that don't already have one. Track the IDs
  // of classes we actually create here so we can run AI suggestions on those alone
  // — never on classes that already had a policy.
  const newlyCreatedClassIds = [];
  for (const [glId, info] of uniqueGLAccounts) {
    const exists = assetClasses.some(c => c.glAccountId === glId);
    if (!exists) {
      try {
        const createRes = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/classes`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
          body: JSON.stringify({
            glAccountId: info.glAccountId,
            glAccountName: info.glAccountName,
            method: 'straight-line',
            usefulLifeMonths: 36,
            expenseAccountId: info.expenseAccountId,
            expenseAccountName: info.expenseAccountName,
            accumAccountId: info.accumAccountId,
            accumAccountName: info.accumAccountName,
          }),
        });
        const createData = await createRes.json();
        if (createData.assetClass) newlyCreatedClassIds.push(createData.assetClass.id);
      } catch (e) { console.error('Failed to create asset class:', e); }
    }
  }

  return newlyCreatedClassIds;
}

/**
 * Run AI amortization suggestions for the given list of asset class IDs.
 *
 * The suggestion is saved as ADVISORY data on the class (aiSuggestion field) but
 * the class's actual policy fields (method, usefulLifeMonths, decliningRate) are
 * NOT modified. The user has to open the edit-policy modal and click "apply
 * suggestion" to actually adopt the AI's recommendation. This was a deliberate
 * product decision: AI advises, the user decides.
 */
async function runAISuggestionsForClasses(classIds = []) {
  if (!Array.isArray(classIds) || classIds.length === 0) return;
  const statusEl = document.getElementById('amortization-status');
  const targets = assetClasses.filter(c => classIds.includes(c.id) && !c.aiSuggestion);
  if (targets.length === 0) return;

  statusEl.textContent = `getting AI amortization suggestions for ${targets.length} new class${targets.length !== 1 ? 'es' : ''}...`;
  statusEl.style.display = '';

  let updated = 0;
  for (const cls of targets) {
    try {
      const sugRes = await fetch('/api/admin/fixed-assets/suggest-amortization', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
        body: JSON.stringify({ assetName: cls.glAccountName, originalCost: 10000 }),
      });
      const suggestion = await sugRes.json();

      if (suggestion.method || suggestion.usefulLifeMonths || suggestion.decliningRate) {
        const suggestedMethod = suggestion.method || 'straight-line';
        // Save the AI's full recommendation as advisory data ONLY. Do not touch
        // the class's actual method/usefulLifeMonths/decliningRate fields.
        const updateBody = {
          aiSuggestion: {
            method: suggestedMethod,
            usefulLifeMonths: suggestion.usefulLifeMonths || null,
            decliningRate: suggestion.decliningRate || null,
            ccaClass: suggestion.ccaClass || null,
            ccaRate: suggestion.ccaRate || null,
            reasoning: suggestion.reasoning || '',
            suggestedAt: new Date().toISOString(),
          },
        };

        await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/classes/${cls.id}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
          body: JSON.stringify(updateBody),
        });
        updated++;
      }
    } catch (e) {
      console.error(`AI suggestion failed for class "${cls.glAccountName}":`, e);
    }
    statusEl.textContent = `AI suggestions: ${updated} of ${targets.length} classes...`;
  }

  statusEl.textContent = `AI suggestions complete — ${updated} new class${updated !== 1 ? 'es' : ''} now have advisory recommendations. Open "edit policy" to review and apply.`;
  setTimeout(() => { statusEl.style.display = 'none'; }, 8000);

  await loadClientFixedAssets();
}

// AI amortization suggestion
async function suggestAmortization() {
  const name = document.getElementById('asset-name').value.trim();
  const cost = document.getElementById('asset-cost').value;
  const suggestionEl = document.getElementById('ai-suggestion');
  const btn = document.getElementById('btn-suggest-amort');

  if (!name) { alert('enter an asset name first'); return; }

  btn.textContent = 'thinking...';
  btn.disabled = true;
  suggestionEl.style.display = 'none';

  try {
    const res = await fetch('/api/admin/fixed-assets/suggest-amortization', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ assetName: name, originalCost: cost }),
    });
    const data = await res.json();

    if (data.usefulLifeMonths) {
      document.getElementById('asset-useful-life').value = data.usefulLifeMonths;
      document.getElementById('asset-salvage').value = data.salvageValue || 0;
      let html = `<strong>suggested: ${data.usefulLifeMonths} months</strong>`;
      if (data.ccaClass) html += ` (CCA Class ${data.ccaClass}, ${data.ccaRate})`;
      html += `<br>${data.reasoning}`;
      suggestionEl.innerHTML = html;
      suggestionEl.style.display = '';
    }
  } catch (e) {
    suggestionEl.textContent = 'failed to get suggestion';
    suggestionEl.style.display = '';
  }
  btn.textContent = 'suggest';
  btn.disabled = false;
}

function populateAssetDropdowns(asset) {
  const selects = {
    'asset-account-asset': { types: ['Fixed Asset', 'Other Asset'], selected: asset?.assetAccountId },
    'asset-account-expense': { types: ['Expense', 'Other Expense'], selected: asset?.expenseAccountId },
    'asset-account-accum': { types: ['Fixed Asset', 'Other Asset', 'Other Current Asset'], selected: asset?.accumAccountId },
  };
  Object.entries(selects).forEach(([id, cfg]) => {
    const select = document.getElementById(id);
    select.innerHTML = '<option value="">select account...</option>';
    // qboAccounts uses normalized format: { id, name, type } from the server
    qboAccounts.filter(a => cfg.types.includes(a.type)).forEach(a => {
      const opt = document.createElement('option');
      opt.value = a.id;
      opt.textContent = `${a.name} (${a.type})`;
      opt.dataset.name = a.name;
      if (a.id === cfg.selected) opt.selected = true;
      select.appendChild(opt);
    });
  });
}

function openAddAsset() {
  editingAssetId = null;
  document.getElementById('asset-modal-title').textContent = 'add fixed asset';
  document.getElementById('asset-name').value = '';
  document.getElementById('asset-cost').value = '';
  document.getElementById('asset-useful-life').value = '';
  document.getElementById('asset-salvage').value = '0';
  document.getElementById('asset-acq-date').value = '';
  document.getElementById('asset-modal-error').style.display = 'none';
  document.getElementById('ai-suggestion').style.display = 'none';
  populateAssetDropdowns(null);
  document.getElementById('asset-modal').style.display = 'flex';
}

function openEditAsset(id) {
  const asset = allFixedAssets.find(a => a.id === id);
  if (!asset) return;
  editingAssetId = id;
  document.getElementById('asset-modal-title').textContent = asset.name;
  document.getElementById('asset-name').value = asset.name;
  document.getElementById('asset-cost').value = asset.originalCost;
  document.getElementById('asset-useful-life').value = asset.usefulLifeMonths;
  document.getElementById('asset-salvage').value = asset.salvageValue;
  document.getElementById('asset-acq-date').value = asset.acquisitionDate;
  document.getElementById('asset-modal-error').style.display = 'none';

  // Show description if available
  const descEl = document.getElementById('asset-description-display');
  if (descEl) {
    if (asset.description) {
      descEl.textContent = asset.description;
      descEl.style.display = '';
    } else {
      descEl.style.display = 'none';
    }
  }

  // Show class policy info
  const cls = getAssetClass(asset);
  const suggestionEl = document.getElementById('ai-suggestion');
  if (cls) {
    const methodLabel = cls.method === 'declining-balance'
      ? `declining balance @ ${((cls.decliningRate || 0) * 100).toFixed(0)}%/yr`
      : `straight-line, ${cls.usefulLifeMonths || '—'} months`;
    let html = `<strong>GL account policy (${cls.glAccountName}):</strong> ${methodLabel}`;
    if (cls.aiSuggestion?.ccaClass) html += ` — CCA Class ${cls.aiSuggestion.ccaClass} (${cls.aiSuggestion.ccaRate})`;
    suggestionEl.innerHTML = html;
    suggestionEl.style.display = '';
  } else {
    suggestionEl.style.display = 'none';
  }

  populateAssetDropdowns(asset);
  document.getElementById('asset-modal').style.display = 'flex';
}

function closeAssetModal() { document.getElementById('asset-modal').style.display = 'none'; }

async function saveAsset() {
  const errorEl = document.getElementById('asset-modal-error');
  errorEl.style.display = 'none';
  const name = document.getElementById('asset-name').value.trim();
  const originalCost = document.getElementById('asset-cost').value;
  const usefulLifeMonths = document.getElementById('asset-useful-life').value;
  const salvageValue = document.getElementById('asset-salvage').value || '0';
  const acquisitionDate = document.getElementById('asset-acq-date').value;
  const assetAcctEl = document.getElementById('asset-account-asset');
  const expenseAcctEl = document.getElementById('asset-account-expense');
  const accumAcctEl = document.getElementById('asset-account-accum');

  if (!name || !originalCost || !usefulLifeMonths || !acquisitionDate) {
    errorEl.textContent = 'name, cost, useful life, and acquisition date are required';
    errorEl.style.display = ''; return;
  }
  if (!assetAcctEl.value || !expenseAcctEl.value || !accumAcctEl.value) {
    errorEl.textContent = 'all three QBO accounts must be selected';
    errorEl.style.display = ''; return;
  }

  const body = { name, originalCost, usefulLifeMonths, salvageValue, acquisitionDate,
    assetAccountId: assetAcctEl.value, assetAccountName: assetAcctEl.selectedOptions[0]?.dataset.name || '',
    expenseAccountId: expenseAcctEl.value, expenseAccountName: expenseAcctEl.selectedOptions[0]?.dataset.name || '',
    accumAccountId: accumAcctEl.value, accumAccountName: accumAcctEl.selectedOptions[0]?.dataset.name || '' };

  try {
    const url = editingAssetId
      ? `/api/admin/clients/${selectedClientId}/fixed-assets/${editingAssetId}`
      : `/api/admin/clients/${selectedClientId}/fixed-assets`;
    const res = await fetch(url, {
      method: editingAssetId ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (data.error) { errorEl.textContent = data.error; errorEl.style.display = ''; return; }
    closeAssetModal();
    loadClientFixedAssets();
  } catch (e) { errorEl.textContent = 'failed to save asset'; errorEl.style.display = ''; }
}

async function deleteAsset(id) {
  if (!confirm('are you sure you want to delete this asset?')) return;
  try {
    await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/${id}`, {
      method: 'DELETE', headers: { 'Authorization': getAuth() },
    });
    loadClientFixedAssets();
  } catch (e) { alert('failed to delete asset'); }
}

let currentAmortMonth = null;

async function loadAmortizationPreview(monthOverride) {
  const errorEl = document.getElementById('amortization-modal-error');
  const previewEl = document.getElementById('amortization-preview');
  const closeEl = document.getElementById('amortization-close-date');
  const reasonEl = document.getElementById('amortization-month');
  const monthInput = document.getElementById('amortization-month-input');
  const confirmBtn = document.getElementById('amortization-confirm');
  errorEl.style.display = 'none';
  previewEl.innerHTML = '<p style="color:var(--gray-500);">loading...</p>';
  try {
    const url = `/api/admin/clients/${selectedClientId}/fixed-assets/preview-amortization${monthOverride ? `?month=${monthOverride}` : ''}`;
    const res = await fetch(url, { headers: { 'Authorization': getAuth() } });
    const data = await res.json();

    currentAmortMonth = data.month;
    monthInput.value = data.month;
    monthInput.disabled = false;
    closeEl.textContent = data.closeDate ? `qbo books closed through: ${data.closeDate}` : 'no qbo close date set';
    reasonEl.textContent = data.reason ? `(${data.reason})` : '';

    if (data.alreadyRun) {
      monthInput.disabled = true;
      const rd = data.runDetails;
      previewEl.innerHTML = `
        <div style="padding:12px;background:var(--green-50);border-radius:8px;border-left:4px solid var(--green-500);">
          <strong style="color:var(--green-700);">✓ amortization already posted for ${data.month}</strong>
          ${rd ? `<div style="margin-top:8px;font-size:0.85rem;color:var(--gray-700);">
            <div>posted: ${new Date(rd.ranAt).toLocaleDateString()}</div>
            <div>total: $${rd.totalAmount.toFixed(2)} (${rd.assetCount} asset${rd.assetCount !== 1 ? 's' : ''})</div>
            ${rd.journalEntryId ? `<div>JE #${rd.journalEntryId}</div>` : ''}
          </div>` : ''}
          <p style="margin-top:8px;font-size:0.82rem;color:var(--gray-500);">the closing period must advance (by updating the close date in QBO) before amortization can be run for the next month.</p>
        </div>`;
      confirmBtn.style.display = 'none';
      return;
    }
    if (data.eligibleCount === 0) {
      previewEl.innerHTML = '<p style="color:var(--gray-500);">no eligible assets for amortization this month.</p>';
      confirmBtn.style.display = 'none';
      return;
    }

    previewEl.innerHTML = `
      <table class="amortization-preview-table">
        <thead><tr><th>asset</th><th>dr expense</th><th>cr accum</th><th style="text-align:right;">amount</th></tr></thead>
        <tbody>${data.lines.map(l => `<tr><td>${l.assetName}</td><td>${l.expenseAccountName}</td><td>${l.accumAccountName}</td><td style="text-align:right;">$${l.amount.toFixed(2)}</td></tr>`).join('')}</tbody>
        <tfoot><tr><td colspan="3">total</td><td style="text-align:right;">$${data.totalAmount.toFixed(2)}</td></tr></tfoot>
      </table>`;
    confirmBtn.style.display = '';
  } catch (e) {
    previewEl.innerHTML = '';
    errorEl.textContent = 'failed to preview amortization';
    errorEl.style.display = '';
  }
}

async function openRunAmortization() {
  document.getElementById('amortization-modal').style.display = 'flex';
  // Always use the current closing period so we don't jump ahead to a future month
  const closingMonth = currentClosePeriod?.month || null;
  await loadAmortizationPreview(closingMonth);
}

function closeAmortizationModal() { document.getElementById('amortization-modal').style.display = 'none'; }

async function confirmRunAmortization() {
  const errorEl = document.getElementById('amortization-modal-error');
  const btn = document.getElementById('amortization-confirm');
  errorEl.style.display = 'none'; btn.textContent = 'syncing from QBO...'; btn.disabled = true;

  // Always pull the latest fixed assets from QBO before running amortization so
  // we never miss an asset that was added in QBO since the last sync. Silent —
  // no modal. Any new accounts get auto-imported and the preview is refreshed
  // so the user sees what they're actually about to post.
  try {
    const syncRes = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/sync-qbo`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
    });
    const syncData = await syncRes.json();
    if (syncData.accounts && syncData.accounts.length > 0) {
      const newClassIds = await importAssetsFromQboAccounts(syncData.accounts);
      await loadClientFixedAssets();
      await loadAmortizationPreview(currentAmortMonth);
      if (newClassIds.length > 0) runAISuggestionsForClasses(newClassIds);
    }
  } catch (e) {
    console.error('pre-amortization sync failed:', e);
    // Non-fatal — fall through and run amortization with what we have
  }

  btn.textContent = 'posting...';

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/run-amortization`, {
      method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ month: currentAmortMonth }),
    });
    const data = await res.json();
    if (data.error) { errorEl.textContent = data.error; errorEl.style.display = ''; btn.textContent = 'post to quickbooks'; btn.disabled = false; return; }

    closeAmortizationModal();
    const statusEl = document.getElementById('amortization-status');
    statusEl.textContent = `amortization posted to quickbooks — $${data.run.totalAmount.toFixed(2)} for ${data.run.assetCount} assets`;
    statusEl.style.display = '';
    // Invalidate the cached close period so the step-card statuses pick up
    // the freshly-posted run. Re-renders happen via loadClientFixedAssets.
    currentClosePeriod = null;
    await loadClientFixedAssets();
    if (data.reconciliation) showReconciliationResult(data.reconciliation);
  } catch (e) { errorEl.textContent = 'failed to post amortization'; errorEl.style.display = ''; }
  btn.textContent = 'post to quickbooks'; btn.disabled = false;
}

function showReconciliationResult(rec) {
  const panel = document.getElementById('reconciliation-panel');
  if (!panel) return;
  if (rec.error) {
    panel.innerHTML = `<div style="padding:12px;border-radius:6px;background:#fef3c7;border:1px solid #f59e0b;color:#92400e;">reconciliation error: ${rec.error}</div>`;
    panel.style.display = '';
    return;
  }
  const checks = rec.checks || [];
  if (checks.length === 0) {
    panel.innerHTML = '<div style="padding:12px;border-radius:6px;background:#f3f4f6;color:#6b7280;">no reconciliation checks performed</div>';
    panel.style.display = '';
    return;
  }
  const fmt = (n) => n === null || n === undefined ? '—' : `$${Number(n).toFixed(2)}`;
  const statusBadge = (s) => {
    if (s === 'pass') return '<span style="color:#16a34a;font-weight:600;">✓ pass</span>';
    if (s === 'fail') return '<span style="color:#dc2626;font-weight:600;">✗ fail</span>';
    return '<span style="color:#d97706;font-weight:600;">⚠ missing</span>';
  };
  const headerColor = rec.allPassed ? '#16a34a' : '#dc2626';
  const headerText = rec.allPassed ? '✓ reconciliation passed' : '✗ reconciliation has differences';
  const rows = checks.map(c => `
    <tr>
      <td>${c.type === 'cost' ? 'Cost' : 'Accum'}</td>
      <td>${c.glAccount}${c.accumAccount ? ` <span style="color:#9ca3af;">→ ${c.accumAccount}</span>` : ''}</td>
      <td style="text-align:right;">${fmt(c.schedule)}</td>
      <td style="text-align:right;">${fmt(c.qbo)}</td>
      <td style="text-align:right;">${fmt(c.difference)}</td>
      <td>${statusBadge(c.status)}</td>
    </tr>
    ${c.note ? `<tr><td colspan="6" style="font-size:0.78rem;color:#9ca3af;padding-left:24px;">${c.note}</td></tr>` : ''}
  `).join('');
  panel.innerHTML = `
    <div style="border:1px solid ${headerColor};border-radius:8px;overflow:hidden;margin-top:16px;">
      <div style="padding:10px 14px;background:${headerColor};color:white;font-weight:600;">${headerText} <span style="opacity:0.85;font-weight:400;">(as of ${rec.asOf || ''})</span></div>
      <table class="amortization-preview-table" style="margin:0;">
        <thead><tr><th>type</th><th>account</th><th style="text-align:right;">schedule</th><th style="text-align:right;">QBO</th><th style="text-align:right;">diff</th><th>status</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
  panel.style.display = '';
}

// Fixed asset event listeners. The old top-of-tab action buttons
// (btn-sync-qbo-assets, btn-run-amortization, btn-export-excel) have been
// replaced by step-card action buttons in the month-end close workflow,
// wired up via event delegation in the data-step-action handler above.
document.getElementById('asset-modal-cancel').addEventListener('click', closeAssetModal);
document.getElementById('asset-modal-save').addEventListener('click', saveAsset);
document.getElementById('asset-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeAssetModal(); });
document.getElementById('btn-refresh-close-date').addEventListener('click', async () => {
  const btn = document.getElementById('btn-refresh-close-date');
  const display = document.getElementById('qbo-close-date-display');
  if (btn) btn.textContent = 'refreshing…';
  if (display) display.textContent = 'refreshing from QBO…';
  await loadQBOCloseDate();
  // Also invalidate the close period cache so the workflow picks up the new date
  currentClosePeriod = null;
  await renderCloseSteps();
  if (btn) { btn.textContent = '✓ refreshed'; setTimeout(() => { btn.textContent = 'refresh'; }, 2000); }
});
document.getElementById('amortization-cancel').addEventListener('click', closeAmortizationModal);
document.getElementById('amortization-confirm').addEventListener('click', confirmRunAmortization);
document.getElementById('amortization-month-input').addEventListener('change', (e) => {
  const val = e.target.value;
  if (val && /^\d{4}-\d{2}$/.test(val)) loadAmortizationPreview(val);
});
document.getElementById('amortization-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeAmortizationModal(); });
document.getElementById('qbo-sync-cancel').addEventListener('click', () => { document.getElementById('qbo-sync-modal').style.display = 'none'; });
document.getElementById('qbo-sync-import').addEventListener('click', importSelectedAssets);
document.getElementById('qbo-sync-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) document.getElementById('qbo-sync-modal').style.display = 'none'; });
document.getElementById('btn-suggest-amort').addEventListener('click', suggestAmortization);

// Render the Claude review traffic-light panel into a given panel element.
// Used by both the fixed-asset and prepaid review panels.
// Status can be: clean / warnings / errors / error / skipped / loading / null
function _renderReviewPanelInto(panelId, review) {
  const panel = document.getElementById(panelId);
  if (!panel) return;
  _renderReviewPanelCore(panel, review, panelId);
}

function renderPrepaidReviewPanel(review) {
  _renderReviewPanelInto('prepaid-review-panel', review);
}

function renderReviewPanel(review) {
  _renderReviewPanelInto('claude-review-panel', review);
}

// Shared renderer used by both fixed-asset and prepaid review panels.
// `panelId` is used to scope dismiss-button handlers so multiple panels can co-exist.
function _renderReviewPanelCore(panel, review, panelId) {
  if (!review) {
    panel.style.display = 'none';
    panel.innerHTML = '';
    return;
  }
  if (review.status === 'loading') {
    panel.style.display = '';
    panel.innerHTML = `
      <div style="background:var(--gray-800);border-left:4px solid var(--gray-500);padding:12px 16px;border-radius:6px;color:var(--gray-300);font-size:0.9rem;">
        <strong>Claude is reviewing the schedule…</strong>
        <div style="font-size:0.82rem;color:var(--gray-400);margin-top:4px;">Comparing the workbook to QBO and running calculation/reasonability checks.</div>
      </div>`;
    return;
  }
  const colors = {
    clean:    { bar: '#22c55e', label: 'CLEAN' },
    warnings: { bar: '#f59e0b', label: 'WARNINGS' },
    errors:   { bar: '#ef4444', label: 'ERRORS' },
    error:    { bar: '#ef4444', label: 'REVIEW FAILED' },
    skipped:  { bar: '#6b7280', label: 'SKIPPED' },
  };
  const c = colors[review.status] || { bar: '#6b7280', label: (review.status || 'unknown').toUpperCase() };
  const sevBg = { error: '#7f1d1d', warning: '#78350f', info: '#1e3a8a' };
  const findingRows = (review.findings || []).map(f => `
    <tr>
      <td style="padding:6px 10px;font-size:0.78rem;font-weight:600;background:${sevBg[f.severity] || '#374151'};color:#fff;text-transform:uppercase;">${f.severity || ''}</td>
      <td style="padding:6px 10px;font-size:0.82rem;color:var(--gray-300);">${escapeHtml(f.category || '')}</td>
      <td style="padding:6px 10px;font-size:0.82rem;color:var(--gray-200);">${escapeHtml(f.assetName || '—')}</td>
      <td style="padding:6px 10px;font-size:0.82rem;color:var(--gray-200);">${escapeHtml(f.message || '')}</td>
    </tr>`).join('');
  const findingsTable = (review.findings && review.findings.length)
    ? `<table style="width:100%;border-collapse:collapse;margin-top:10px;background:var(--gray-900);border-radius:4px;overflow:hidden;">
         <thead><tr>
           <th style="padding:6px 10px;text-align:left;font-size:0.72rem;color:var(--gray-400);text-transform:uppercase;">Severity</th>
           <th style="padding:6px 10px;text-align:left;font-size:0.72rem;color:var(--gray-400);text-transform:uppercase;">Category</th>
           <th style="padding:6px 10px;text-align:left;font-size:0.72rem;color:var(--gray-400);text-transform:uppercase;">Item</th>
           <th style="padding:6px 10px;text-align:left;font-size:0.72rem;color:var(--gray-400);text-transform:uppercase;">Message</th>
         </tr></thead>
         <tbody>${findingRows}</tbody>
       </table>`
    : '<div style="font-size:0.82rem;color:var(--gray-400);margin-top:8px;font-style:italic;">No findings.</div>';

  const dismissBtnId = `btn-dismiss-${panelId}`;
  panel.style.display = '';
  panel.innerHTML = `
    <div style="background:var(--gray-800);border-left:4px solid ${c.bar};padding:12px 16px;border-radius:6px;">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:6px;">
        <span style="background:${c.bar};color:#fff;font-weight:700;font-size:0.78rem;padding:3px 10px;border-radius:3px;letter-spacing:0.5px;">${c.label}</span>
        <span style="color:var(--gray-300);font-size:0.9rem;">Claude review — ${escapeHtml(review.asOfMonth || 'most recent period')}</span>
        <button id="${dismissBtnId}" style="margin-left:auto;background:transparent;border:1px solid var(--gray-600);color:var(--gray-400);font-size:0.75rem;padding:3px 10px;border-radius:3px;cursor:pointer;">dismiss</button>
      </div>
      <div style="color:var(--gray-200);font-size:0.88rem;line-height:1.4;">${escapeHtml(review.summary || '')}</div>
      ${findingsTable}
    </div>`;
  document.getElementById(dismissBtnId).addEventListener('click', () => _renderReviewPanelInto(panelId, null));
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Excel import is intentionally hidden from the UI — sync-from-QBO is the canonical
// source of truth for fixed assets and amortization runs. The /import-excel server
// endpoint is preserved for future onboarding (clients migrating from other systems)
// and disaster recovery, but can only be reached by calling the endpoint directly.


// ========================================
// SETTINGS — CHANGE PASSWORD
// ========================================
document.getElementById('btn-change-pw').addEventListener('click', async () => {
  const currentPw = document.getElementById('settings-current-pw').value;
  const newPw = document.getElementById('settings-new-pw').value;
  const confirmPw = document.getElementById('settings-confirm-pw').value;
  const errorEl = document.getElementById('settings-pw-error');
  const successEl = document.getElementById('settings-pw-success');

  errorEl.style.display = 'none';
  successEl.style.display = 'none';

  if (!currentPw || !newPw) {
    errorEl.textContent = 'all fields are required';
    errorEl.style.display = '';
    return;
  }
  if (newPw.length < 6) {
    errorEl.textContent = 'new password must be at least 6 characters';
    errorEl.style.display = '';
    return;
  }
  if (newPw !== confirmPw) {
    errorEl.textContent = 'new passwords do not match';
    errorEl.style.display = '';
    return;
  }

  try {
    const res = await fetch('/api/admin/change-password', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ currentPassword: currentPw, newPassword: newPw }),
    });
    const data = await res.json();
    if (data.error) {
      errorEl.textContent = data.error;
      errorEl.style.display = '';
      return;
    }
    // Update the auth cookie with the new password
    document.cookie = `admin_auth=${newPw};path=/;max-age=86400`;
    successEl.textContent = 'password updated successfully';
    successEl.style.display = '';
    document.getElementById('settings-current-pw').value = '';
    document.getElementById('settings-new-pw').value = '';
    document.getElementById('settings-confirm-pw').value = '';
  } catch (e) {
    errorEl.textContent = 'failed to update password';
    errorEl.style.display = '';
  }
});


// ========================================
// SETTINGS — QBO TOKENS
// ========================================
async function loadQboTokens() {
  const statusEl = document.getElementById('qbo-tokens-status');
  const valueEl = document.getElementById('qbo-tokens-value');
  const copyBtn = document.getElementById('btn-copy-tokens');
  try {
    const res = await fetch('/api/admin/qbo-tokens', {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    if (data.connected && data.tokens) {
      statusEl.textContent = 'qbo connected — tokens available';
      statusEl.style.color = 'var(--green-600)';
      valueEl.value = data.tokens;
      valueEl.style.display = '';
      copyBtn.style.display = '';
    } else if (data.connected) {
      statusEl.textContent = 'qbo connected but token file not found';
      statusEl.style.color = 'var(--amber-600)';
    } else {
      statusEl.textContent = 'qbo not connected — click "connect quickbooks" in the top bar';
      statusEl.style.color = 'var(--gray-500)';
      valueEl.style.display = 'none';
      copyBtn.style.display = 'none';
    }
  } catch (e) {
    statusEl.textContent = 'could not fetch token status';
  }
}

document.getElementById('btn-copy-tokens').addEventListener('click', () => {
  const valueEl = document.getElementById('qbo-tokens-value');
  navigator.clipboard.writeText(valueEl.value);
  const btn = document.getElementById('btn-copy-tokens');
  btn.textContent = 'copied!';
  setTimeout(() => { btn.textContent = 'copy to clipboard'; }, 2000);
});

document.getElementById('btn-refresh-tokens').addEventListener('click', loadQboTokens);

// ========================================
// FINANCIAL REPORTING
// ========================================
let reportingClients = [];
let reportingSelectedClientId = null;
let reportingSnapshots = [];
let reportingActiveTab = 'snapshots';
let reportingDetailSnapshot = null;       // currently open snapshot in the fullscreen modal
let reportingDetailSelectedFY = null;     // FY label being viewed
// Mapping state — populated lazily when the user opens the mapping tab.
let mappingDimensions = null;            // { "1": { id, name, description, isPrimary }, ... }
let mappingAccountValues = {};           // keyed by `id:<id>` or `name:<name>` — server canonical
let mappingQboAccounts = [];             // chart of accounts from /api/admin/accounts
let mappingFilter = '';

function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function setReportingStatus(msg, kind) {
  const el = document.getElementById('reporting-status');
  if (!el) return;
  el.textContent = msg || '';
  el.style.color = kind === 'error' ? 'var(--red-600, #c62828)'
    : kind === 'success' ? 'var(--green-600)'
    : 'var(--gray-500)';
}

function setMappingStatus(msg, kind) {
  const el = document.getElementById('mapping-status');
  if (!el) return;
  el.textContent = msg || '';
  el.style.color = kind === 'error' ? 'var(--red-600, #c62828)'
    : kind === 'success' ? 'var(--green-600)'
    : 'var(--gray-500)';
}

async function initReportingView() {
  try {
    const res = await fetch('/api/admin/clients', { headers: { 'Authorization': getAuth() } });
    const data = await res.json();
    reportingClients = data.clients || [];
  } catch (e) {
    console.error('Failed to load clients for reporting:', e);
    reportingClients = [];
  }

  const sel = document.getElementById('reporting-client-select');
  sel.innerHTML = '<option value="">— select client —</option>'
    + reportingClients.map(c => `<option value="${c.id}">${escapeHtml(c.name)}</option>`).join('');

  // Restore previously chosen client if we have one
  if (reportingSelectedClientId && reportingClients.find(c => c.id === reportingSelectedClientId)) {
    sel.value = reportingSelectedClientId;
    await onReportingClientChanged();
  } else {
    renderReportingSnapshots([]);
    renderMappingTable();
    renderDimensionsSummary();
  }
}

function switchReportingTab(tab) {
  reportingActiveTab = tab;
  document.querySelectorAll('[data-reporting-tab]').forEach(b => b.classList.toggle('active', b.dataset.reportingTab === tab));
  document.querySelectorAll('.reporting-tab').forEach(el => el.style.display = 'none');
  const target = document.getElementById('reporting-tab-' + tab);
  if (target) target.style.display = '';
  if (tab === 'mapping' && reportingSelectedClientId) {
    loadMappingsAndAccounts();
  }
  if (tab === 'dimensions' && reportingSelectedClientId) {
    loadDimensionsConfig();
  }
}

document.querySelectorAll('[data-reporting-tab]').forEach(btn => {
  btn.addEventListener('click', () => switchReportingTab(btn.dataset.reportingTab));
});

async function onReportingClientChanged() {
  hideReportingDetail();
  if (!reportingSelectedClientId) {
    renderReportingSnapshots([]);
    mappingDimensions = null;
    mappingAccountValues = {};
    mappingQboAccounts = [];
    renderMappingTable();
    renderDimensionsSummary();
    return;
  }
  // Always refresh snapshots so the snapshots tab is current.
  await loadReportingSnapshots();
  if (reportingActiveTab === 'mapping') {
    await loadMappingsAndAccounts();
  } else if (reportingActiveTab === 'dimensions') {
    await loadDimensionsConfig();
  } else {
    // Reset cached mapping state — it'll reload when the user opens the tab.
    mappingDimensions = null;
    mappingAccountValues = {};
    mappingQboAccounts = [];
  }
}

document.getElementById('reporting-client-select')?.addEventListener('change', async (e) => {
  reportingSelectedClientId = e.target.value || null;
  await onReportingClientChanged();
});

async function loadReportingSnapshots() {
  if (!reportingSelectedClientId) return;
  try {
    const res = await fetch(`/api/admin/clients/${reportingSelectedClientId}/reporting/tb-snapshots`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    reportingSnapshots = data.snapshots || [];
    renderReportingSnapshots(reportingSnapshots);
  } catch (e) {
    console.error('Failed to load snapshots:', e);
    renderReportingSnapshots([]);
  }
}

function renderReportingSnapshots(list) {
  const el = document.getElementById('reporting-snapshots-list');
  if (!reportingSelectedClientId) {
    el.innerHTML = '<div class="empty-state"><p>select a client above to see saved snapshots</p></div>';
    return;
  }
  if (!list.length) {
    el.innerHTML = '<div class="empty-state"><p>no snapshots yet</p><span>pull one using the button above</span></div>';
    return;
  }
  el.innerHTML = `
    <table style="width:100%;border-collapse:collapse;font-size:0.88rem;">
      <thead>
        <tr style="border-bottom:1px solid var(--gray-200);text-align:left;">
          <th style="padding:8px 6px;">coverage</th>
          <th style="padding:8px 6px;">close date</th>
          <th style="padding:8px 6px;text-align:right;">accounts</th>
          <th style="padding:8px 6px;text-align:right;">months</th>
          <th style="padding:8px 6px;">pulled</th>
          <th style="padding:8px 6px;"></th>
        </tr>
      </thead>
      <tbody>
        ${list.map(s => {
          const isMulti = s.type === 'multi-period';
          const coverage = isMulti
            ? `${escapeHtml((s.fiscalYearLabels || []).join(' · '))}`
            : `${escapeHtml(s.startDate || '')} → ${escapeHtml(s.endDate || '')} <span style="color:var(--gray-500);font-size:0.78rem;">(YTD legacy)</span>`;
          const cutoff = isMulti ? escapeHtml(s.closeDate || s.cutoffMonthEnd || '') : '—';
          const months = isMulti ? s.periodCount : '—';
          return `
            <tr style="border-bottom:1px solid var(--gray-100);">
              <td style="padding:8px 6px;">${coverage}</td>
              <td style="padding:8px 6px;">${cutoff}</td>
              <td style="padding:8px 6px;text-align:right;">${s.accountCount}</td>
              <td style="padding:8px 6px;text-align:right;">${months}</td>
              <td style="padding:8px 6px;font-size:0.8rem;color:var(--gray-500);">${formatDate(s.createdAt)}</td>
              <td style="padding:8px 6px;text-align:right;white-space:nowrap;">
                <button class="btn-edit-client" data-snapshot-view="${s.id}">view</button>
                <button class="btn-edit-client" data-snapshot-delete="${s.id}" style="color:var(--red-600,#c62828);">delete</button>
              </td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table>
  `;
  el.querySelectorAll('[data-snapshot-view]').forEach(b => {
    b.addEventListener('click', () => viewReportingSnapshot(b.dataset.snapshotView));
  });
  el.querySelectorAll('[data-snapshot-delete]').forEach(b => {
    b.addEventListener('click', () => deleteReportingSnapshot(b.dataset.snapshotDelete));
  });
}

async function viewReportingSnapshot(snapshotId) {
  if (!reportingSelectedClientId) return;
  try {
    const res = await fetch(`/api/admin/clients/${reportingSelectedClientId}/reporting/tb-snapshots/${snapshotId}`, {
      headers: { 'Authorization': getAuth() },
    });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    renderReportingDetail(data.snapshot);
  } catch (e) {
    alert('Failed to load snapshot: ' + e.message);
  }
}

function fmtCell(n) {
  if (n === null || n === undefined || n === 0) return '';
  const v = Number(n) || 0;
  const formatted = Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return v < 0 ? `(${formatted})` : formatted;
}

function renderReportingDetail(snap) {
  reportingDetailSnapshot = snap;
  if (snap.type === 'multi-period') {
    const fys = snap.fiscalYears || [];
    const current = fys.find(fy => fy.isCurrent && (fy.months || []).length > 0);
    const firstWithData = fys.find(fy => (fy.months || []).length > 0);
    reportingDetailSelectedFY = (current || firstWithData || fys[0])?.label || null;
  } else {
    reportingDetailSelectedFY = null;
  }
  document.getElementById('reporting-detail-modal').style.display = 'flex';
  drawReportingDetail();
}

function drawReportingDetail() {
  const snap = reportingDetailSnapshot;
  if (!snap) return;
  const title = document.getElementById('reporting-detail-title');
  const tabsEl = document.getElementById('reporting-detail-fy-tabs');
  const table = document.getElementById('reporting-detail-table');

  if (snap.type === 'multi-period') {
    title.textContent = `trial balance · close date ${snap.closeDate}`;
    const fys = snap.fiscalYears || [];
    tabsEl.innerHTML = fys.map(fy => {
      const isActive = fy.label === reportingDetailSelectedFY;
      const hasData = (fy.months || []).length > 0;
      const bg = isActive ? 'var(--blue-600, #2563eb)' : '#fff';
      const color = isActive ? '#fff' : 'var(--gray-700)';
      const border = isActive ? 'var(--blue-600, #2563eb)' : 'var(--gray-300)';
      const opacity = hasData ? 1 : 0.5;
      return `<button class="reporting-fy-tab" data-fy-label="${escapeHtml(fy.label)}" ${hasData ? '' : 'disabled'} style="padding:5px 12px;border:1px solid ${border};border-radius:14px;background:${bg};color:${color};font-size:0.78rem;font-weight:500;cursor:${hasData ? 'pointer' : 'not-allowed'};opacity:${opacity};">${escapeHtml(fy.label)}${fy.isCurrent ? ' · current' : ''}</button>`;
    }).join('');
    tabsEl.querySelectorAll('.reporting-fy-tab').forEach(b => {
      b.addEventListener('click', () => {
        reportingDetailSelectedFY = b.dataset.fyLabel;
        drawReportingDetail();
      });
    });
    table.innerHTML = renderSingleFYTable(snap, reportingDetailSelectedFY);
  } else {
    title.textContent = `trial balance — ${snap.startDate} → ${snap.endDate}${snap.currency ? ' (' + snap.currency + ')' : ''}`;
    tabsEl.innerHTML = '';
    const rows = (snap.accounts || []).map(a => `
      <tr style="border-bottom:1px solid var(--gray-100);">
        <td style="padding:6px 8px;font-size:0.85rem;">${escapeHtml(a.accountName)}</td>
        <td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums;">${a.debit ? fmtMoney(a.debit) : ''}</td>
        <td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums;">${a.credit ? fmtMoney(a.credit) : ''}</td>
        <td style="padding:6px 8px;text-align:right;font-variant-numeric:tabular-nums;color:${a.net < 0 ? 'var(--red-600,#c62828)' : 'var(--gray-700)'};">${fmtMoney(a.net)}</td>
      </tr>
    `).join('');
    table.innerHTML = `
      <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
        <thead>
          <tr style="border-bottom:1px solid var(--gray-300);text-align:left;">
            <th style="padding:8px;">account</th>
            <th style="padding:8px;text-align:right;">debit</th>
            <th style="padding:8px;text-align:right;">credit</th>
            <th style="padding:8px;text-align:right;">net</th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td colspan="4" style="padding:16px;color:var(--gray-500);">no account rows</td></tr>'}</tbody>
        <tfoot>
          <tr style="border-top:1px solid var(--gray-300);font-weight:600;">
            <td style="padding:8px;">totals</td>
            <td style="padding:8px;text-align:right;">${fmtMoney(snap.totalDebit)}</td>
            <td style="padding:8px;text-align:right;">${fmtMoney(snap.totalCredit)}</td>
            <td style="padding:8px;text-align:right;">${fmtMoney((snap.totalDebit || 0) - (snap.totalCredit || 0))}</td>
          </tr>
        </tfoot>
      </table>
    `;
  }
}

// Build the canonical 12-month axis for a fiscal year using its start/end
// dates so we always show the full FY layout, even if QBO returned no
// activity for some months (those columns just appear blank).
function fyMonthAxis(fy) {
  if (!fy?.startDate || !fy?.endDate) return fy?.months || [];
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const out = [];
  const [sy, sm] = fy.startDate.split('-').map(Number);
  const [ey, em] = fy.endDate.split('-').map(Number);
  let y = sy, m = sm;
  while (y < ey || (y === ey && m <= em)) {
    out.push({
      period: `${y}-${String(m).padStart(2, '0')}`,
      label: `${monthNames[m - 1]} ${y}`,
    });
    m++;
    if (m > 12) { m = 1; y++; }
  }
  return out;
}

// Single-FY table: account rows × 12 month columns. Pulled-but-empty
// months show a blank cell; not-yet-pulled months (current FY) also show
// blank. Sticky first column for account name.
function renderSingleFYTable(snap, fyLabel) {
  const fys = snap.fiscalYears || [];
  const fy = fys.find(f => f.label === fyLabel) || fys.find(f => (f.months || []).length > 0) || fys[0];
  if (!fy) return '<div class="empty-state"><p>no fiscal years in this snapshot</p></div>';

  const months = fyMonthAxis(fy);
  const pulledPeriods = new Set((fy.months || []).map(m => m.period));

  const accountRows = (snap.accounts || []).filter(a => {
    // Only show accounts with any activity in this FY — keeps the table
    // readable when an account is brand new or fully retired.
    return months.some(m => (a.nets?.[m.period] || 0) !== 0);
  }).map(a => {
    const cells = months.map(m => {
      const v = a.nets?.[m.period];
      const wasPulled = pulledPeriods.has(m.period);
      const isNeg = (v || 0) < 0;
      const bg = wasPulled ? '#fff' : 'var(--gray-50)';
      return `<td style="padding:4px 6px;text-align:right;font-variant-numeric:tabular-nums;font-size:0.82rem;color:${isNeg ? 'var(--red-600,#c62828)' : 'var(--gray-700)'};background:${bg};">${fmtCell(v)}</td>`;
    }).join('');
    return `
      <tr style="border-bottom:1px solid var(--gray-100);">
        <td style="padding:6px 10px;font-size:0.85rem;position:sticky;left:0;background:#fff;border-right:1px solid var(--gray-200);min-width:280px;">${escapeHtml(a.accountName)}</td>
        ${cells}
      </tr>
    `;
  }).join('');

  const totalCells = months.map(m => {
    const v = snap.totals?.[m.period];
    const wasPulled = pulledPeriods.has(m.period);
    const bg = wasPulled ? 'var(--gray-50)' : '#f3f4f6';
    return `<td style="padding:6px;text-align:right;font-variant-numeric:tabular-nums;font-size:0.82rem;font-weight:600;background:${bg};">${fmtCell(v)}</td>`;
  }).join('');

  const monthHeaderCells = months.map(m => {
    const wasPulled = pulledPeriods.has(m.period);
    return `<th style="padding:8px 6px;font-size:0.78rem;font-weight:500;color:${wasPulled ? 'var(--gray-700)' : 'var(--gray-400)'};text-align:right;min-width:96px;background:var(--gray-50);">${escapeHtml(m.label)}${wasPulled ? '' : ' <span style="font-size:0.65rem;font-weight:400;">(unpulled)</span>'}</th>`;
  }).join('');

  return `
    <div style="overflow:auto;border:1px solid var(--gray-200);border-radius:6px;">
      <table style="border-collapse:collapse;font-size:0.82rem;width:max-content;min-width:100%;">
        <thead style="position:sticky;top:0;z-index:2;">
          <tr>
            <th style="padding:8px 10px;text-align:left;position:sticky;left:0;background:var(--gray-50);border-right:1px solid var(--gray-200);min-width:280px;z-index:3;">account · ${escapeHtml(fy.label)}${fy.isCurrent ? ' (current)' : ''}</th>
            ${monthHeaderCells}
          </tr>
        </thead>
        <tbody>${accountRows || `<tr><td colspan="${months.length + 1}" style="padding:16px;color:var(--gray-500);">no accounts with activity in this fiscal year</td></tr>`}</tbody>
        <tfoot>
          <tr style="border-top:2px solid var(--gray-300);">
            <td style="padding:6px 10px;font-weight:600;position:sticky;left:0;background:var(--gray-50);border-right:1px solid var(--gray-200);">net total</td>
            ${totalCells}
          </tr>
        </tfoot>
      </table>
    </div>
    <p style="font-size:0.78rem;color:var(--gray-500);margin-top:8px;">net change per period (debit − credit). negatives in parentheses. blank cells with grey background = month not yet pulled (current FY only goes through last closed period).</p>
  `;
}

function hideReportingDetail() {
  const modal = document.getElementById('reporting-detail-modal');
  if (modal) modal.style.display = 'none';
  reportingDetailSnapshot = null;
  reportingDetailSelectedFY = null;
}

document.getElementById('btn-close-reporting-detail')?.addEventListener('click', hideReportingDetail);

async function deleteReportingSnapshot(snapshotId) {
  if (!reportingSelectedClientId) return;
  if (!confirm('Delete this snapshot? This cannot be undone.')) return;
  try {
    const res = await fetch(`/api/admin/clients/${reportingSelectedClientId}/reporting/tb-snapshots/${snapshotId}`, {
      method: 'DELETE',
      headers: { 'Authorization': getAuth() },
    });
    if (!res.ok) throw new Error(await res.text());
    hideReportingDetail();
    await loadReportingSnapshots();
  } catch (e) {
    alert('Delete failed: ' + e.message);
  }
}

// ---- GL mapping ----

async function loadMappingsAndAccounts() {
  if (!reportingSelectedClientId) return;
  setMappingStatus('loading...');
  try {
    const [mapRes, acctRes] = await Promise.all([
      fetch(`/api/admin/clients/${reportingSelectedClientId}/reporting/mappings`, { headers: { 'Authorization': getAuth() } }),
      fetch(`/api/admin/accounts?clientId=${reportingSelectedClientId}`, { headers: { 'Authorization': getAuth() } }),
    ]);
    const mapData = await mapRes.json();
    const acctData = await acctRes.json();
    mappingDimensions = mapData.dimensions || {};
    mappingAccountValues = mapData.accounts || {};
    mappingQboAccounts = acctData.accounts || [];
    renderDimensionsSummary();
    renderMappingTable();
    setMappingStatus(mappingQboAccounts.length === 0 ? 'no QBO accounts — connect QuickBooks to load the chart of accounts' : '');
  } catch (e) {
    console.error('Failed to load mappings:', e);
    setMappingStatus('failed to load mappings: ' + e.message, 'error');
  }
}

function dimensionList() {
  if (!mappingDimensions) return [];
  return Array.from({ length: 10 }, (_, i) => mappingDimensions[String(i + 1)] || { id: i + 1, name: '', description: '', isPrimary: i === 0 });
}

function activeDimensions() {
  // Visible dimension columns = those with a non-empty name. Primary is always shown.
  return dimensionList().filter(d => d.isPrimary || (d.name && d.name.trim()));
}

function renderDimensionsSummary() {
  const el = document.getElementById('dimensions-summary');
  if (!el) return;
  if (!reportingSelectedClientId) { el.textContent = 'select a client to view its dimensions'; return; }
  if (!mappingDimensions) { el.textContent = ''; return; }
  const dims = dimensionList();
  el.innerHTML = dims.map(d => {
    const label = d.name || `(unnamed)`;
    const tag = d.isPrimary ? '<span style="background:var(--blue-50);color:var(--blue-600);padding:1px 6px;border-radius:4px;font-size:0.72rem;font-weight:600;margin-left:4px;">PRIMARY</span>' : '';
    const opts = Array.isArray(d.options) ? d.options.map(optionShape).filter(Boolean) : [];
    const optHint = opts.length ? ` <span style="color:var(--gray-500);font-size:0.78rem;">(${opts.length} option${opts.length === 1 ? '' : 's'})</span>` : '';
    return `<span style="display:inline-block;margin:2px 8px 2px 0;"><strong>${d.id}.</strong> ${escapeHtml(label)}${tag}${optHint}</span>`;
  }).join('');
}

function mappingKeyFor(acct) {
  return acct.id ? `id:${acct.id}` : `name:${acct.name}`;
}

// Coerce a stored option (string from old shape, or {name, statementType})
// into a canonical { name, statementType } pair. Renamed accounts hold
// onto their value across this normalization.
function optionShape(o) {
  if (typeof o === 'string') return { name: o, statementType: null };
  if (o && typeof o === 'object') return { name: String(o.name || ''), statementType: o.statementType || null };
  return null;
}

const STMT_TYPE_LABEL = { balance_sheet: 'Balance Sheet', income_statement: 'Income Statement' };

// Build the mapping-table cell for one (account × dimension):
//   - <select> with <optgroup> grouping when the dimension has options
//   - <input type="text"> when the dimension is free-text
// Pre-existing mapped values that no longer match any option are kept
// as a "⚠ legacy" entry so renames/removals don't silently drop data.
function renderMappingSelectOrInput(d, cur, dataAttrs) {
  const opts = (Array.isArray(d.options) ? d.options : []).map(optionShape).filter(Boolean);
  if (opts.length === 0) {
    return `<input type="text" class="mapping-cell" ${dataAttrs} value="${escapeHtml(cur)}" style="width:100%;padding:4px 6px;border:1px solid var(--gray-200);border-radius:3px;font-size:0.82rem;background:#fff;">`;
  }
  const inOpts = cur && opts.some(o => o.name.toLowerCase() === cur.toLowerCase());
  // Group by statement type for the primary dimension; for non-primary,
  // statement type is meaningless so collapse all options into a single group.
  const groups = { balance_sheet: [], income_statement: [], unclassified: [] };
  for (const o of opts) {
    const key = o.statementType && groups[o.statementType] ? o.statementType : 'unclassified';
    groups[key].push(o);
  }
  function renderOptionTag(o) {
    const sel = o.name.toLowerCase() === cur.toLowerCase() ? ' selected' : '';
    return `<option value="${escapeHtml(o.name)}"${sel}>${escapeHtml(o.name)}</option>`;
  }
  const groupHtml = d.isPrimary
    ? [
        groups.balance_sheet.length     ? `<optgroup label="${STMT_TYPE_LABEL.balance_sheet}">${groups.balance_sheet.map(renderOptionTag).join('')}</optgroup>` : '',
        groups.income_statement.length  ? `<optgroup label="${STMT_TYPE_LABEL.income_statement}">${groups.income_statement.map(renderOptionTag).join('')}</optgroup>` : '',
        groups.unclassified.length      ? `<optgroup label="Unclassified">${groups.unclassified.map(renderOptionTag).join('')}</optgroup>` : '',
      ].join('')
    : opts.map(renderOptionTag).join('');
  const legacyOpt = cur && !inOpts
    ? `<option value="${escapeHtml(cur)}" selected>⚠ ${escapeHtml(cur)} (not in options)</option>`
    : '';
  return `
    <select class="mapping-cell" ${dataAttrs}
      style="width:100%;padding:4px 6px;border:1px solid var(--gray-200);border-radius:3px;font-size:0.82rem;background:#fff;">
      <option value=""${cur ? '' : ' selected'}>— unmapped —</option>
      ${legacyOpt}
      ${groupHtml}
    </select>
  `;
}

function renderMappingTable() {
  const el = document.getElementById('mapping-table');
  if (!el) return;
  if (!reportingSelectedClientId) {
    el.className = 'empty-state';
    el.innerHTML = '<p>select a client above to map GL accounts</p>';
    return;
  }
  if (!mappingQboAccounts.length) {
    el.className = 'empty-state';
    el.innerHTML = '<p>no QBO accounts loaded</p><span>connect QuickBooks for this client, or click "refresh accounts from QBO"</span>';
    return;
  }
  const dims = activeDimensions();
  const filter = mappingFilter.trim().toLowerCase();
  const rows = mappingQboAccounts.filter(a => {
    if (!filter) return true;
    return (a.name || '').toLowerCase().includes(filter) || (a.type || '').toLowerCase().includes(filter) || (a.acctNum || '').toLowerCase().includes(filter);
  });

  el.className = '';
  el.innerHTML = `
    <table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <thead>
        <tr style="border-bottom:1px solid var(--gray-300);text-align:left;background:var(--gray-50);">
          <th style="padding:8px 6px;position:sticky;left:0;background:var(--gray-50);z-index:1;min-width:240px;">account</th>
          <th style="padding:8px 6px;min-width:140px;">type</th>
          ${dims.map(d => `<th style="padding:8px 6px;min-width:140px;">${escapeHtml(d.name || `dim ${d.id}`)}${d.isPrimary ? ' *' : ''}</th>`).join('')}
        </tr>
      </thead>
      <tbody>
        ${rows.map(a => {
          const key = mappingKeyFor(a);
          const stored = mappingAccountValues[key]?.values || {};
          return `
            <tr data-mapping-row="${escapeHtml(key)}" style="border-bottom:1px solid var(--gray-100);">
              <td style="padding:6px;position:sticky;left:0;background:#fff;">
                <div style="font-weight:500;">${escapeHtml(fmtAcct(a))}</div>
              </td>
              <td style="padding:6px;color:var(--gray-600);">${escapeHtml(a.type || '')}</td>
              ${dims.map(d => {
                const cur = stored[String(d.id)] || '';
                const dataAttrs = `data-account-id="${escapeHtml(a.id || '')}" data-account-name="${escapeHtml(a.name || '')}" data-account-type="${escapeHtml(a.type || '')}" data-dim="${d.id}"`;
                const cellHtml = renderMappingSelectOrInput(d, cur, dataAttrs);
                return `<td style="padding:4px;">${cellHtml}</td>`;
              }).join('')}
            </tr>
          `;
        }).join('')}
      </tbody>
    </table>
    <p style="font-size:0.78rem;color:var(--gray-500);margin-top:8px;">* primary dimension. ${rows.length} of ${mappingQboAccounts.length} accounts shown.</p>
  `;

  // Auto-save per cell. Inputs save on blur (lets the user finish typing);
  // selects save on change (single-click commit). All cells skip the save
  // when value didn't actually change.
  el.querySelectorAll('input.mapping-cell').forEach(input => {
    input.dataset.original = input.value;
    input.addEventListener('blur', () => {
      if (input.value === input.dataset.original) return;
      saveMappingCell(input);
    });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') input.blur();
    });
  });
  el.querySelectorAll('select.mapping-cell').forEach(sel => {
    sel.dataset.original = sel.value;
    sel.addEventListener('change', () => {
      if (sel.value === sel.dataset.original) return;
      saveMappingCell(sel);
    });
  });
}

document.getElementById('mapping-account-filter')?.addEventListener('input', (e) => {
  mappingFilter = e.target.value;
  renderMappingTable();
});

async function saveMappingCell(input) {
  const accountId = input.dataset.accountId;
  const accountName = input.dataset.accountName;
  const accountType = input.dataset.accountType;
  const dimId = input.dataset.dim;
  const value = input.value;

  // Build full values object from the row's other cells (inputs and selects)
  // so the server doesn't need to merge — the row holds the canonical state
  // for this account.
  const row = input.closest('[data-mapping-row]');
  const values = {};
  row.querySelectorAll('.mapping-cell').forEach(cell => {
    values[cell.dataset.dim] = cell.value;
  });

  // Optimistic UI: tint the cell while saving.
  input.style.background = '#fffbe6';
  try {
    const res = await fetch(`/api/admin/clients/${reportingSelectedClientId}/reporting/mappings/accounts`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ accountId, accountName, accountType, values }),
    });
    if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || res.statusText);
    const data = await res.json();
    mappingAccountValues[data.key] = data.account;
    input.dataset.original = input.value;
    input.style.background = '#e6ffed';
    setTimeout(() => { input.style.background = '#fff'; }, 600);
    setMappingStatus(`saved ${accountName} → ${dimensionList().find(d => String(d.id) === dimId)?.name || 'dim ' + dimId}`, 'success');
  } catch (e) {
    input.style.background = '#fff5f5';
    setMappingStatus('save failed: ' + e.message, 'error');
  }
}

document.getElementById('btn-refresh-mapping-accounts')?.addEventListener('click', async () => {
  if (!reportingSelectedClientId) return setMappingStatus('select a client first', 'error');
  await loadMappingsAndAccounts();
});

// ---- Configure Dimensions tab ----

function setDimensionsConfigStatus(msg, kind) {
  const el = document.getElementById('dimensions-config-status');
  if (!el) return;
  el.textContent = msg || '';
  el.style.color = kind === 'error' ? 'var(--red-600, #c62828)'
    : kind === 'success' ? 'var(--green-600)'
    : 'var(--gray-500)';
}

async function loadDimensionsConfig() {
  if (!reportingSelectedClientId) return;
  // Reuse the mapping loader which already fetches dimensions+accounts.
  // If we've never loaded, do it now.
  if (!mappingDimensions) {
    await loadMappingsAndAccounts();
  }
  renderDimensionsConfig();
}

function renderDimensionsConfig() {
  const container = document.getElementById('dimensions-config-container');
  if (!container) return;
  if (!reportingSelectedClientId) {
    container.innerHTML = '<div class="empty-state"><p>select a client above to configure dimensions</p></div>';
    return;
  }
  if (!mappingDimensions) {
    container.innerHTML = '<div class="empty-state"><p>loading dimensions…</p></div>';
    return;
  }
  const dims = dimensionList();
  container.innerHTML = dims.map(d => renderDimensionCard(d)).join('');
  wireDimensionCard();
}

// Each dimension is a card. Primary (dim 1) is expanded by default and
// shows a statement-type column on the options table; others are collapsed
// and use a simpler name-only options table.
function renderDimensionCard(d) {
  const opts = Array.isArray(d.options) ? d.options : [];
  const expanded = d.isPrimary || opts.length > 0;
  const headerLabel = d.name || (d.isPrimary ? 'Financial Statement (unnamed)' : `(unnamed)`);
  const primaryBadge = d.isPrimary
    ? '<span style="background:var(--blue-50);color:var(--blue-600);padding:1px 8px;border-radius:4px;font-size:0.7rem;font-weight:600;margin-left:8px;">PRIMARY</span>'
    : '';
  const countBadge = `<span style="font-size:0.78rem;color:var(--gray-500);margin-left:8px;">${opts.length} option${opts.length === 1 ? '' : 's'}</span>`;
  return `
    <details class="dim-card" data-dim="${d.id}" ${expanded ? 'open' : ''} style="border:1px solid var(--gray-200);border-radius:8px;margin-bottom:10px;background:#fff;">
      <summary style="padding:12px 16px;cursor:pointer;user-select:none;display:flex;align-items:center;gap:8px;font-size:0.92rem;font-weight:500;">
        <strong>${d.id}.</strong>
        <span class="dim-card-title">${escapeHtml(headerLabel)}</span>
        ${primaryBadge}
        ${countBadge}
      </summary>
      <div style="padding:0 16px 16px;">
        <div class="form-group" style="margin-bottom:8px;">
          <label style="font-size:0.78rem;color:var(--gray-600);">name${d.isPrimary ? ' (drives statements)' : ' (leave blank to hide from mapping table)'}</label>
          <input type="text" class="dim-name-input" data-dim="${d.id}" value="${escapeHtml(d.name || '')}" placeholder="${d.isPrimary ? 'Financial Statement' : 'e.g. Department, Cost Center'}" style="width:100%;padding:6px 8px;border:1px solid var(--gray-300);border-radius:4px;">
        </div>
        <div class="form-group" style="margin-bottom:12px;">
          <label style="font-size:0.78rem;color:var(--gray-600);">description (optional)</label>
          <input type="text" class="dim-desc-input" data-dim="${d.id}" value="${escapeHtml(d.description || '')}" placeholder="what does this dimension represent?" style="width:100%;padding:6px 8px;border:1px solid var(--gray-200);border-radius:4px;font-size:0.82rem;">
        </div>

        <h4 style="margin:12px 0 6px;font-size:0.82rem;font-weight:600;color:var(--gray-700);">options <span style="font-weight:400;color:var(--gray-500);">— rendered as a dropdown on the GL mapping table when at least one is configured</span></h4>
        ${renderOptionsTable(d)}
        <div style="display:flex;gap:8px;align-items:center;margin-top:8px;">
          <button class="btn-edit-client dim-add-option" data-dim="${d.id}">+ add option</button>
          <span style="font-size:0.78rem;color:var(--gray-500);">${d.isPrimary ? 'tag each option as Balance Sheet or Income Statement so we can split mapped accounts into the right statement.' : ''}</span>
        </div>
      </div>
    </details>
  `;
}

function renderOptionsTable(d) {
  const opts = Array.isArray(d.options) ? d.options : [];
  if (!opts.length) {
    return '<div class="empty-state" style="padding:12px;font-size:0.82rem;"><p style="margin:0;">no options yet</p></div>';
  }
  const showType = d.isPrimary;
  return `
    <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
      <thead>
        <tr style="border-bottom:1px solid var(--gray-200);text-align:left;">
          <th style="padding:6px 8px;font-weight:500;color:var(--gray-600);font-size:0.78rem;">option name</th>
          ${showType ? '<th style="padding:6px 8px;font-weight:500;color:var(--gray-600);font-size:0.78rem;width:200px;">statement type</th>' : ''}
          <th style="padding:6px 8px;width:60px;"></th>
        </tr>
      </thead>
      <tbody>
        ${opts.map((o, idx) => {
          const name = typeof o === 'string' ? o : o.name;
          const stmt = (typeof o === 'object' ? o.statementType : null) || '';
          return `
            <tr data-dim="${d.id}" data-opt-idx="${idx}" style="border-bottom:1px solid var(--gray-100);">
              <td style="padding:4px 8px;">
                <input type="text" class="dim-opt-name" data-dim="${d.id}" data-opt-idx="${idx}" value="${escapeHtml(name)}" style="width:100%;padding:4px 6px;border:1px solid var(--gray-200);border-radius:3px;font-size:0.82rem;">
              </td>
              ${showType ? `
                <td style="padding:4px 8px;">
                  <select class="dim-opt-stmt" data-dim="${d.id}" data-opt-idx="${idx}" style="width:100%;padding:4px 6px;border:1px solid var(--gray-200);border-radius:3px;font-size:0.82rem;background:#fff;">
                    <option value=""${stmt === '' ? ' selected' : ''}>— unclassified —</option>
                    <option value="balance_sheet"${stmt === 'balance_sheet' ? ' selected' : ''}>Balance Sheet</option>
                    <option value="income_statement"${stmt === 'income_statement' ? ' selected' : ''}>Income Statement</option>
                  </select>
                </td>
              ` : ''}
              <td style="padding:4px 8px;text-align:right;">
                <button class="btn-edit-client dim-opt-delete" data-dim="${d.id}" data-opt-idx="${idx}" title="remove option" style="color:var(--red-600,#c62828);padding:2px 8px;">×</button>
              </td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table>
  `;
}

function wireDimensionCard() {
  // Name + description: save on blur if changed.
  document.querySelectorAll('.dim-name-input').forEach(inp => {
    inp.dataset.original = inp.value;
    inp.addEventListener('blur', () => {
      if (inp.value === inp.dataset.original) return;
      saveDimensionMeta(inp.dataset.dim, { name: inp.value });
    });
  });
  document.querySelectorAll('.dim-desc-input').forEach(inp => {
    inp.dataset.original = inp.value;
    inp.addEventListener('blur', () => {
      if (inp.value === inp.dataset.original) return;
      saveDimensionMeta(inp.dataset.dim, { description: inp.value });
    });
  });
  // Option name: save the whole options array on blur.
  document.querySelectorAll('.dim-opt-name').forEach(inp => {
    inp.dataset.original = inp.value;
    inp.addEventListener('blur', () => {
      if (inp.value === inp.dataset.original) return;
      saveDimensionOptions(inp.dataset.dim);
    });
    inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') inp.blur(); });
  });
  // Statement type: save on change.
  document.querySelectorAll('.dim-opt-stmt').forEach(sel => {
    sel.dataset.original = sel.value;
    sel.addEventListener('change', () => {
      if (sel.value === sel.dataset.original) return;
      saveDimensionOptions(sel.dataset.dim);
    });
  });
  // Delete option.
  document.querySelectorAll('.dim-opt-delete').forEach(btn => {
    btn.addEventListener('click', () => {
      const dim = btn.dataset.dim;
      const idx = parseInt(btn.dataset.optIdx, 10);
      const d = mappingDimensions?.[String(dim)];
      if (!d || !Array.isArray(d.options)) return;
      d.options.splice(idx, 1);
      // Re-render this card and persist.
      renderDimensionsConfig();
      saveDimensionOptions(dim);
    });
  });
  // Add option.
  document.querySelectorAll('.dim-add-option').forEach(btn => {
    btn.addEventListener('click', () => {
      const dim = btn.dataset.dim;
      const d = mappingDimensions?.[String(dim)];
      if (!d) return;
      d.options = Array.isArray(d.options) ? d.options : [];
      d.options.push({ name: '', statementType: null });
      renderDimensionsConfig();
      // Focus the new row's name input so the user can start typing.
      const newIdx = d.options.length - 1;
      const focusEl = document.querySelector(`.dim-opt-name[data-dim="${dim}"][data-opt-idx="${newIdx}"]`);
      if (focusEl) {
        // Re-open the parent details if it collapsed during the rerender.
        focusEl.closest('.dim-card')?.setAttribute('open', '');
        focusEl.focus();
      }
    });
  });
}

// Read current options from the rendered DOM for one dimension and persist.
function collectDimensionOptionsFromDom(dimId) {
  const rows = document.querySelectorAll(`tr[data-dim="${dimId}"][data-opt-idx]`);
  const out = [];
  rows.forEach(row => {
    const nameEl = row.querySelector('.dim-opt-name');
    const stmtEl = row.querySelector('.dim-opt-stmt');
    if (!nameEl) return;
    const name = (nameEl.value || '').trim();
    if (!name) return; // server will drop blanks anyway, but skip locally to keep DOM index stable
    out.push({ name, statementType: stmtEl ? (stmtEl.value || null) : null });
  });
  return out;
}

async function saveDimensionMeta(dimId, patch) {
  const dimensions = { [dimId]: patch };
  await persistDimensions(dimensions);
}

async function saveDimensionOptions(dimId) {
  // Prefer the in-memory copy if a row was just spliced — that's the
  // canonical post-delete/post-add state. Otherwise read from the DOM.
  const inMem = mappingDimensions?.[String(dimId)]?.options;
  const fromDom = collectDimensionOptionsFromDom(dimId);
  const merged = (inMem && Array.isArray(inMem) && inMem.length === fromDom.length)
    ? fromDom
    : (inMem || fromDom);
  const dimensions = { [dimId]: { options: merged } };
  await persistDimensions(dimensions);
}

async function persistDimensions(dimensions) {
  setDimensionsConfigStatus('saving…');
  try {
    const res = await fetch(`/api/admin/clients/${reportingSelectedClientId}/reporting/mappings/dimensions`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({ dimensions }),
    });
    if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || res.statusText);
    const data = await res.json();
    mappingDimensions = data.dimensions;
    setDimensionsConfigStatus('saved ✓', 'success');
    renderDimensionsConfig();
    renderDimensionsSummary();
    // The mapping table needs to re-render too because options/select rendering depends on this.
    if (mappingQboAccounts && mappingQboAccounts.length) renderMappingTable();
  } catch (e) {
    setDimensionsConfigStatus('save failed: ' + e.message, 'error');
  }
}

document.getElementById('btn-configure-dimensions')?.addEventListener('click', () => {
  if (!reportingSelectedClientId) return setMappingStatus('select a client first', 'error');
  switchReportingTab('dimensions');
});

document.getElementById('btn-pull-tb')?.addEventListener('click', async () => {
  const clientId = document.getElementById('reporting-client-select').value;
  if (!clientId) return setReportingStatus('select a client first', 'error');

  reportingSelectedClientId = clientId;
  const btn = document.getElementById('btn-pull-tb');
  const originalLabel = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'pulling...';
  setReportingStatus('pulling monthly trial balances — this can take 15-30 seconds for 30+ months...');
  try {
    const res = await fetch(`/api/admin/clients/${clientId}/reporting/tb-snapshots`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
      body: JSON.stringify({}),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Pull failed');
    const snap = data.snapshot;
    const periodCount = (snap.fiscalYears || []).reduce((sum, fy) => sum + (fy.months?.length || 0), 0);
    setReportingStatus(`snapshot saved — ${snap.accounts.length} accounts × ${periodCount} months · close date ${snap.closeDate}`, 'success');
    await loadReportingSnapshots();
    renderReportingDetail(snap);
  } catch (e) {
    setReportingStatus('failed: ' + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = originalLabel;
  }
});

// ========================================
// INIT
// ========================================
async function init() {
  // Check for QBO callback
  const params = new URLSearchParams(window.location.search);
  if (params.get('qbo') === 'connected') {
    const connectedClientId = params.get('clientId');
    // If we came back from QBO OAuth for a specific client, navigate to that client
    if (connectedClientId) {
      // Switch to clients view and open the client detail after clients load
      setTimeout(async () => {
        document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
        document.querySelector('[data-admin-view="clients"]')?.classList.add('active');
        document.querySelectorAll('.admin-view').forEach(v => v.style.display = 'none');
        document.getElementById('view-clients').style.display = '';
        await loadClients();
        openClientDetail(connectedClientId);
        // Auto-jump to the fixed-assets tab and silently sync + import everything
        // from QBO. The whole point of connecting is to pull data — no reason to
        // make the user click through a modal just to confirm what they already
        // asked for by clicking "connect quickbooks".
        setTimeout(async () => {
          switchClientTab('close');
          await autoSyncAndImportFromQbo();
        }, 300);
      }, 100);
    }
    window.history.replaceState({}, '', window.location.pathname);
  } else if (params.get('qbo') === 'error') {
    alert('Failed to connect QuickBooks. Please try again.');
    window.history.replaceState({}, '', window.location.pathname);
  }

  // Load accounts (global, for document approval dropdowns)
  await loadAccounts();
  loadStats();
  loadAnalyses();
  loadQboTokens();
}

init();

// Auto-refresh every 30 seconds (but don't re-fetch accounts each time)
setInterval(() => {
  loadStats();
  loadAnalyses();
}, 30000);

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
      options += `<option value="${a.id}" data-name="${a.name.replace(/"/g, '&quot;')}"${selected}>${a.name}</option>`;
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
            ${analysis.documentType ? `<div class="analysis-summary-item"><span class="label">type</span><span class="value">${analysis.documentType}</span></div>` : ''}
            ${analysis.vendor ? `<div class="analysis-summary-item"><span class="label">vendor</span><span class="value">${analysis.vendor}</span></div>` : ''}
            ${analysis.customer ? `<div class="analysis-summary-item"><span class="label">customer</span><span class="value">${analysis.customer}</span></div>` : ''}
            ${analysis.date ? `<div class="analysis-summary-item"><span class="label">date</span><span class="value">${analysis.date}</span></div>` : ''}
            ${analysis.confidence ? `<div class="analysis-summary-item"><span class="label">confidence</span><span class="value">${analysis.confidence}</span></div>` : ''}
          </div>
          ${analysis.summary ? `<p>${analysis.summary}</p>` : ''}
        </div>

        <div id="entries-${a.id}">
        ${entries.map((entry, i) => `
          <div class="admin-entry-block" data-analysis-id="${a.id}" data-entry-index="${i}">
            <div class="admin-entry-header">
              <span class="admin-entry-type">${entry.type || 'journal entry'}</span>
              <span class="admin-entry-date">${a.status === 'pending' ? `<input type="date" class="edit-date" value="${entry.date || ''}" data-field="date" data-entry="${i}">` : (entry.date || '')}</span>
            </div>
            <div class="admin-entry-memo">${a.status === 'pending' ? `<input type="text" class="edit-memo" value="${(entry.memo || '').replace(/"/g, '&quot;')}" placeholder="memo / description" data-field="memo" data-entry="${i}">` : (entry.memo || '')}</div>
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
                    <td>${a.status === 'pending' ? renderAccountSelect(line.accountName, line.accountId, i, li) : line.accountName}</td>
                    <td>${a.status === 'pending' ? `<input type="text" class="edit-cell" value="${(line.description || '').replace(/"/g, '&quot;')}" data-field="description" data-entry="${i}" data-line="${li}">` : (line.description || '')}</td>
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
            <ul>${analysis.needsReview.map(n => `<li>${n}</li>`).join('')}</ul>
          </div>
        ` : ''}

        ${analysis.notes ? `<div class="admin-notes">${analysis.notes}</div>` : ''}

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
            ${a.status === 'approved' ? 'approved' : a.status === 'rejected' ? 'rejected' : a.status}
            ${a.reviewedAt ? ` on ${formatDate(a.reviewedAt)}` : ''}
            ${a.rejectReason ? ` — reason: ${a.rejectReason}` : ''}
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

function openClientDetail(clientId) {
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

  // Load accounts for this client's QBO connection
  loadAccounts(clientId);

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
  if (tab === 'fixed-assets' && selectedClientId) loadClientFixedAssets();
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
    renderFixedAssets();
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

// Per-asset list was removed from the UI in favor of class-level summaries
// (renderAssetClasses now shows count + total cost/accum/NBV per class). The
// detailed asset list still lives in the exported Excel workbook. This function
// just keeps the "last amortization run" footer line up to date.
function renderFixedAssets() {
  const lastRunEl = document.getElementById('last-amortization-run');
  if (!lastRunEl) return;
  if (fixedAssetRuns.length > 0) {
    const lastRun = fixedAssetRuns[fixedAssetRuns.length - 1];
    lastRunEl.textContent = `last run: ${lastRun.month} on ${new Date(lastRun.ranAt).toLocaleDateString()} — $${lastRun.totalAmount.toFixed(2)} (${lastRun.assetCount} assets)`;
  } else {
    lastRunEl.textContent = 'no amortization runs yet';
  }
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
    closeEl.textContent = data.closeDate ? `qbo books closed through: ${data.closeDate}` : 'no qbo close date set';
    reasonEl.textContent = data.reason ? `(${data.reason})` : '';

    if (data.alreadyRun) {
      previewEl.innerHTML =
        `<p style="color:var(--amber-600);">amortization was already run for ${data.month}${data.runDetails ? ` on ${new Date(data.runDetails.ranAt).toLocaleDateString()}.<br>total: $${data.runDetails.totalAmount.toFixed(2)} (${data.runDetails.assetCount} assets)` : ''}</p>`;
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
  await loadAmortizationPreview(null);
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
    if (data.reconciliation) showReconciliationResult(data.reconciliation);
    loadClientFixedAssets();
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

// Fixed asset event listeners
document.getElementById('asset-modal-cancel').addEventListener('click', closeAssetModal);
document.getElementById('asset-modal-save').addEventListener('click', saveAsset);
document.getElementById('asset-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeAssetModal(); });
document.getElementById('btn-run-amortization').addEventListener('click', openRunAmortization);
document.getElementById('btn-refresh-close-date').addEventListener('click', loadQBOCloseDate);
document.getElementById('amortization-cancel').addEventListener('click', closeAmortizationModal);
document.getElementById('amortization-confirm').addEventListener('click', confirmRunAmortization);
document.getElementById('amortization-month-input').addEventListener('change', (e) => {
  const val = e.target.value;
  if (val && /^\d{4}-\d{2}$/.test(val)) loadAmortizationPreview(val);
});
document.getElementById('amortization-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeAmortizationModal(); });
document.getElementById('btn-sync-qbo-assets').addEventListener('click', syncFromQBO);
document.getElementById('qbo-sync-cancel').addEventListener('click', () => { document.getElementById('qbo-sync-modal').style.display = 'none'; });
document.getElementById('qbo-sync-import').addEventListener('click', importSelectedAssets);
document.getElementById('qbo-sync-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) document.getElementById('qbo-sync-modal').style.display = 'none'; });
document.getElementById('btn-suggest-amort').addEventListener('click', suggestAmortization);

// Excel export — runs Claude review server-side and surfaces a traffic-light panel
// after the file download starts.
document.getElementById('btn-export-excel').addEventListener('click', async () => {
  if (!selectedClientId) return;
  const btn = document.getElementById('btn-export-excel');
  btn.textContent = 'exporting & reviewing...';
  btn.disabled = true;
  // Show a loading state in the review panel right away so the user knows
  // a review is running in parallel with the file generation.
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
    // Fetch the cached review and render the panel
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
  btn.textContent = 'export to Excel';
  btn.disabled = false;
});

// Render the Claude review traffic-light panel below the export button.
// Status can be: clean / warnings / errors / error / skipped / loading / null
function renderReviewPanel(review) {
  const panel = document.getElementById('claude-review-panel');
  if (!panel) return;
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
           <th style="padding:6px 10px;text-align:left;font-size:0.72rem;color:var(--gray-400);text-transform:uppercase;">Asset</th>
           <th style="padding:6px 10px;text-align:left;font-size:0.72rem;color:var(--gray-400);text-transform:uppercase;">Message</th>
         </tr></thead>
         <tbody>${findingRows}</tbody>
       </table>`
    : '<div style="font-size:0.82rem;color:var(--gray-400);margin-top:8px;font-style:italic;">No findings.</div>';

  panel.style.display = '';
  panel.innerHTML = `
    <div style="background:var(--gray-800);border-left:4px solid ${c.bar};padding:12px 16px;border-radius:6px;">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:6px;">
        <span style="background:${c.bar};color:#fff;font-weight:700;font-size:0.78rem;padding:3px 10px;border-radius:3px;letter-spacing:0.5px;">${c.label}</span>
        <span style="color:var(--gray-300);font-size:0.9rem;">Claude review — ${escapeHtml(review.asOfMonth || 'most recent period')}</span>
        <button id="btn-dismiss-review" style="margin-left:auto;background:transparent;border:1px solid var(--gray-600);color:var(--gray-400);font-size:0.75rem;padding:3px 10px;border-radius:3px;cursor:pointer;">dismiss</button>
      </div>
      <div style="color:var(--gray-200);font-size:0.88rem;line-height:1.4;">${escapeHtml(review.summary || '')}</div>
      ${findingsTable}
    </div>`;
  document.getElementById('btn-dismiss-review').addEventListener('click', () => renderReviewPanel(null));
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
          switchClientTab('fixed-assets');
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

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

  // Load accounts for this client's QBO connection
  loadAccounts(clientId);

  // Render client info
  document.getElementById('client-info-content').innerHTML = `
    <div style="display:grid;gap:12px;max-width:400px;padding:16px 0;">
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">email</span><br><strong>${client.email}</strong></div>
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">billing</span><br><span class="client-billing-badge ${client.billingFrequency}">${client.billingFrequency}</span></div>
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">client id</span><br><code style="font-size:0.82rem;">${clientId}</code></div>
    </div>`;
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
  document.getElementById('client-modal-error').style.display = 'none';
  document.getElementById('client-modal').style.display = '';
}

function closeClientModal() { document.getElementById('client-modal').style.display = 'none'; }

async function saveClient() {
  const name = document.getElementById('client-name').value.trim();
  const email = document.getElementById('client-email').value.trim();
  const password = document.getElementById('client-password').value;
  const billingFrequency = document.getElementById('client-billing').value;
  const errorEl = document.getElementById('client-modal-error');

  if (!name || !email) { errorEl.textContent = 'name and email are required'; errorEl.style.display = ''; return; }
  if (!editingClientId && !password) { errorEl.textContent = 'password is required for new clients'; errorEl.style.display = ''; return; }

  const body = { name, email, billingFrequency };
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
      if (!asset.active || !asset.assetAccountId) continue;
      const hasClass = assetClasses.some(c => c.glAccountId === asset.assetAccountId || c.glAccountName === asset.glAccountName);
      if (!hasClass && !missingGLAccounts.has(asset.assetAccountId)) {
        missingGLAccounts.set(asset.assetAccountId, {
          glAccountId: asset.assetAccountId,
          glAccountName: asset.glAccountName || asset.assetAccountName || asset.name,
        });
      }
    }

    if (missingGLAccounts.size > 0) {
      for (const [glId, info] of missingGLAccounts) {
        try {
          const createRes = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/classes`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
            body: JSON.stringify({
              glAccountId: info.glAccountId,
              glAccountName: info.glAccountName,
              method: 'straight-line',
              usefulLifeMonths: 60,
              salvageValue: 0,
            }),
          });
          const createData = await createRes.json();
          if (createData.assetClass) assetClasses.push(createData.assetClass);
        } catch (e) { console.error('Failed to auto-create class:', e); }
      }
      // Run AI suggestions for newly created classes
      runAISuggestionsForClasses();
    }

    renderAssetClasses();
    renderFixedAssets();
  } catch (e) { console.error('Failed to load fixed assets:', e); }
}

function getAssetClass(asset) {
  return assetClasses.find(c => c.glAccountId === asset.assetAccountId || c.glAccountName === asset.glAccountName) || null;
}

function renderAssetClasses() {
  const list = document.getElementById('asset-classes-list');
  if (!list) return;

  if (assetClasses.length === 0) {
    list.innerHTML = '<p style="color:var(--gray-400);text-align:center;padding:20px;font-size:0.85rem;">no GL accounts yet — sync from QBO to import assets and their GL accounts</p>';
    return;
  }

  list.innerHTML = assetClasses.map(c => {
    const assetCount = allFixedAssets.filter(a => a.active && (a.assetAccountId === c.glAccountId || a.glAccountName === c.glAccountName)).length;
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
            <span>salvage: $${(c.salvageValue || 0).toLocaleString()}</span>
            ${c.expenseAccountName ? `<span>expense acct: ${c.expenseAccountName}</span>` : ''}
            ${c.accumAccountName ? `<span>accum acct: ${c.accumAccountName}</span>` : ''}
          </div>
          ${ai ? `<div class="asset-class-ai">${ai.ccaClass ? `CCA Class ${ai.ccaClass} (${ai.ccaRate})` : ''} ${ai.reasoning ? `— ${ai.reasoning}` : ''}</div>` : ''}
        </div>
        <div class="asset-card-actions">
          <button class="btn-edit-client" onclick="openEditClass('${c.id}')">edit policy</button>
          <button class="btn-delete-asset" onclick="deleteClass('${c.id}')">delete</button>
        </div>
      </div>`;
  }).join('');
}

function renderFixedAssets() {
  const list = document.getElementById('fixed-assets-list');
  const activeAssets = allFixedAssets.filter(a => a.active);

  if (activeAssets.length === 0) {
    list.innerHTML = '<p style="color:var(--gray-400);text-align:center;padding:40px;">no fixed assets yet. sync from QBO to import.</p>';
  } else {
    list.innerHTML = activeAssets.map(a => {
      const cls = getAssetClass(a);
      const className = cls?.glAccountName || a.glAccountName || '—';
      const acqDate = a.acquisitionDate || '';

      return `
        <div class="asset-card">
          <div class="asset-card-main">
            <div class="asset-card-header">
              <div class="asset-card-name">${a.name}</div>
            </div>
            <div class="asset-card-desc">${className}${a.vendorName ? ` — ${a.vendorName}` : ''}</div>
            <div class="asset-card-amounts">
              <span class="asset-card-amount">cost: <strong>$${(a.originalCost || 0).toLocaleString()}</strong></span>
              <span class="asset-card-amount">acquired: <strong>${acqDate}</strong></span>
            </div>
          </div>
          <div class="asset-card-actions">
            <button class="btn-edit-client" onclick="openEditAsset('${a.id}')">edit</button>
            <button class="btn-delete-asset" onclick="deleteAsset('${a.id}')">delete</button>
          </div>
        </div>`;
    }).join('');
  }

  const lastRunEl = document.getElementById('last-amortization-run');
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
  document.getElementById('class-salvage').value = cls.salvageValue || 0;
  document.getElementById('class-modal-error').style.display = 'none';

  // Show AI suggestion if available
  const aiEl = document.getElementById('class-ai-suggestion');
  if (cls.aiSuggestion) {
    const ai = cls.aiSuggestion;
    let html = `<strong>AI: ${ai.method === 'declining-balance' ? 'declining balance' : 'straight-line'}</strong>`;
    if (ai.ccaClass) html += ` — CCA Class ${ai.ccaClass} (${ai.ccaRate})`;
    if (ai.reasoning) html += `<br>${ai.reasoning}`;
    aiEl.innerHTML = html;
    aiEl.style.display = '';
  } else {
    aiEl.style.display = 'none';
  }

  populateClassDropdowns(cls);
  toggleClassMethod();
  document.getElementById('asset-class-modal').style.display = 'flex';
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
  const salvageValue = document.getElementById('class-salvage').value || '0';
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
    salvageValue: parseFloat(salvageValue) || 0,
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
      document.getElementById('class-salvage').value = data.salvageValue || 0;

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
            $${(a.originalCost || 0).toLocaleString()}
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

  const importedAssets = [];

  for (const cb of checkboxes) {
    const account = JSON.parse(cb.dataset.account);
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

  // Auto-create asset classes for each unique GL account
  const uniqueGLAccounts = new Map();
  for (const asset of importedAssets) {
    if (asset.assetAccountId && !uniqueGLAccounts.has(asset.assetAccountId)) {
      uniqueGLAccounts.set(asset.assetAccountId, {
        glAccountId: asset.assetAccountId,
        glAccountName: asset.glAccountName || asset.assetAccountName || asset.name,
      });
    }
  }

  // Create classes for any GL accounts that don't already have one
  for (const [glId, info] of uniqueGLAccounts) {
    const exists = assetClasses.some(c => c.glAccountId === glId);
    if (!exists) {
      try {
        await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/classes`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
          body: JSON.stringify({
            glAccountId: info.glAccountId,
            glAccountName: info.glAccountName,
            method: 'straight-line',
            usefulLifeMonths: 60,
            salvageValue: 0,
          }),
        });
      } catch (e) { console.error('Failed to create asset class:', e); }
    }
  }

  // Close modal and show assets immediately
  document.getElementById('qbo-sync-modal').style.display = 'none';
  importBtn.textContent = 'import selected';
  importBtn.disabled = false;
  await loadClientFixedAssets();

  // Now run AI suggestions for each new class in the background
  if (uniqueGLAccounts.size > 0) {
    runAISuggestionsForClasses();
  }
}

/**
 * Run AI amortization suggestions for all asset classes that don't have one yet
 */
async function runAISuggestionsForClasses() {
  const statusEl = document.getElementById('amortization-status');
  const classesNeedingSuggestion = assetClasses.filter(c => !c.aiSuggestion);
  if (classesNeedingSuggestion.length === 0) return;

  statusEl.textContent = `getting AI amortization suggestions for ${classesNeedingSuggestion.length} class${classesNeedingSuggestion.length !== 1 ? 'es' : ''}...`;
  statusEl.style.display = '';

  let updated = 0;
  for (const cls of classesNeedingSuggestion) {
    try {
      const sugRes = await fetch('/api/admin/fixed-assets/suggest-amortization', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
        body: JSON.stringify({ assetName: cls.glAccountName, originalCost: 10000 }),
      });
      const suggestion = await sugRes.json();

      if (suggestion.usefulLifeMonths) {
        const suggestedMethod = suggestion.method || 'straight-line';
        const updateBody = {
          method: suggestedMethod,
          usefulLifeMonths: suggestion.usefulLifeMonths,
          decliningRate: suggestion.decliningRate || null,
          salvageValue: suggestion.salvageValue || 0,
          aiSuggestion: {
            ccaClass: suggestion.ccaClass || null,
            ccaRate: suggestion.ccaRate || null,
            reasoning: suggestion.reasoning || '',
            method: suggestedMethod,
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
    statusEl.textContent = `AI suggestions: ${updated} of ${classesNeedingSuggestion.length} classes updated...`;
  }

  statusEl.textContent = `AI suggestions complete — ${updated} class${updated !== 1 ? 'es' : ''} updated`;
  setTimeout(() => { statusEl.style.display = 'none'; }, 5000);

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

async function openRunAmortization() {
  const errorEl = document.getElementById('amortization-modal-error');
  errorEl.style.display = 'none';
  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/preview-amortization`, {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();

    if (data.alreadyRun) {
      document.getElementById('amortization-month').textContent = `month: ${data.month}`;
      document.getElementById('amortization-preview').innerHTML =
        `<p style="color:var(--amber-600);">amortization was already run for ${data.month} on ${new Date(data.runDetails.ranAt).toLocaleDateString()}.<br>total: $${data.runDetails.totalAmount.toFixed(2)} (${data.runDetails.assetCount} assets)</p>`;
      document.getElementById('amortization-confirm').style.display = 'none';
      document.getElementById('amortization-modal').style.display = 'flex';
      return;
    }
    if (data.eligibleCount === 0) {
      document.getElementById('amortization-month').textContent = `month: ${data.month}`;
      document.getElementById('amortization-preview').innerHTML = '<p style="color:var(--gray-500);">no eligible assets for amortization this month.</p>';
      document.getElementById('amortization-confirm').style.display = 'none';
      document.getElementById('amortization-modal').style.display = 'flex';
      return;
    }

    document.getElementById('amortization-month').textContent = `month: ${data.month}`;
    document.getElementById('amortization-preview').innerHTML = `
      <table class="amortization-preview-table">
        <thead><tr><th>asset</th><th>dr expense</th><th>cr accum</th><th style="text-align:right;">amount</th></tr></thead>
        <tbody>${data.lines.map(l => `<tr><td>${l.assetName}</td><td>${l.expenseAccountName}</td><td>${l.accumAccountName}</td><td style="text-align:right;">$${l.amount.toFixed(2)}</td></tr>`).join('')}</tbody>
        <tfoot><tr><td colspan="3">total</td><td style="text-align:right;">$${data.totalAmount.toFixed(2)}</td></tr></tfoot>
      </table>`;
    document.getElementById('amortization-confirm').style.display = '';
    document.getElementById('amortization-modal').style.display = 'flex';
  } catch (e) { alert('failed to preview amortization'); }
}

function closeAmortizationModal() { document.getElementById('amortization-modal').style.display = 'none'; }

async function confirmRunAmortization() {
  const errorEl = document.getElementById('amortization-modal-error');
  const btn = document.getElementById('amortization-confirm');
  errorEl.style.display = 'none'; btn.textContent = 'posting...'; btn.disabled = true;

  try {
    const res = await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets/run-amortization`, {
      method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
    });
    const data = await res.json();
    if (data.error) { errorEl.textContent = data.error; errorEl.style.display = ''; btn.textContent = 'post to quickbooks'; btn.disabled = false; return; }

    closeAmortizationModal();
    const statusEl = document.getElementById('amortization-status');
    statusEl.textContent = `amortization posted to quickbooks — $${data.run.totalAmount.toFixed(2)} for ${data.run.assetCount} assets`;
    statusEl.style.display = '';
    setTimeout(() => { statusEl.style.display = 'none'; }, 5000);
    loadClientFixedAssets();
  } catch (e) { errorEl.textContent = 'failed to post amortization'; errorEl.style.display = ''; }
  btn.textContent = 'post to quickbooks'; btn.disabled = false;
}

// Fixed asset event listeners
document.getElementById('asset-modal-cancel').addEventListener('click', closeAssetModal);
document.getElementById('asset-modal-save').addEventListener('click', saveAsset);
document.getElementById('asset-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeAssetModal(); });
document.getElementById('btn-run-amortization').addEventListener('click', openRunAmortization);
document.getElementById('amortization-cancel').addEventListener('click', closeAmortizationModal);
document.getElementById('amortization-confirm').addEventListener('click', confirmRunAmortization);
document.getElementById('amortization-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) closeAmortizationModal(); });
document.getElementById('btn-sync-qbo-assets').addEventListener('click', syncFromQBO);
document.getElementById('qbo-sync-cancel').addEventListener('click', () => { document.getElementById('qbo-sync-modal').style.display = 'none'; });
document.getElementById('qbo-sync-import').addEventListener('click', importSelectedAssets);
document.getElementById('qbo-sync-modal').addEventListener('click', (e) => { if (e.target === e.currentTarget) document.getElementById('qbo-sync-modal').style.display = 'none'; });
document.getElementById('btn-suggest-amort').addEventListener('click', suggestAmortization);


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

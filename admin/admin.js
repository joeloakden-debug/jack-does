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
async function loadAccounts() {
  try {
    const res = await fetch('/api/admin/accounts', {
      headers: { 'Authorization': getAuth() },
    });
    const data = await res.json();
    qboAccounts = data.accounts || [];
    console.log(`Loaded ${qboAccounts.length} QBO accounts`);
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

    const qboEl = document.getElementById('qbo-status');
    const dot = qboEl.querySelector('.status-dot');
    if (data.qboConnected) {
      dot.classList.remove('disconnected');
      dot.classList.add('connected');
      qboEl.lastChild.textContent = ' qbo connected';
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
let editingAssetId = null;

async function loadClients() {
  try {
    const res = await fetch('/api/admin/clients', { headers: { 'Authorization': getAuth() } });
    const data = await res.json();
    allClients = data.clients || [];
    renderClients();
  } catch (e) { console.error('Failed to load clients:', e); }
}

function renderClients() {
  const container = document.getElementById('clients-list');
  if (allClients.length === 0) {
    container.innerHTML = '<div class="empty-state"><p>no clients yet</p><span>click "+ add client" to create one</span></div>';
    return;
  }
  container.innerHTML = allClients.map(c => `
    <div class="client-card client-card-clickable" data-client-id="${c.id}" onclick="openClientDetail('${c.id}')">
      <div class="client-card-info">
        <span class="client-card-name">${c.name}</span>
        <span class="client-card-email">${c.email}</span>
      </div>
      <div class="client-card-meta">
        <span class="client-billing-badge ${c.billingFrequency}">${c.billingFrequency}</span>
        <button class="btn-edit-client" onclick="event.stopPropagation(); openEditClient('${c.id}')">edit</button>
      </div>
    </div>
  `).join('');
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

  // Render client info
  document.getElementById('client-info-content').innerHTML = `
    <div style="display:grid;gap:12px;max-width:400px;padding:16px 0;">
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">email</span><br><strong>${client.email}</strong></div>
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">billing</span><br><span class="client-billing-badge ${client.billingFrequency}">${client.billingFrequency}</span></div>
      <div><span style="font-size:0.78rem;color:var(--gray-400);text-transform:uppercase;">client id</span><br><code style="font-size:0.82rem;">${clientId}</code></div>
    </div>`;
}

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
    renderFixedAssets();
  } catch (e) { console.error('Failed to load fixed assets:', e); }
}

function renderFixedAssets() {
  const list = document.getElementById('fixed-assets-list');
  const activeAssets = allFixedAssets.filter(a => a.active);

  if (activeAssets.length === 0) {
    list.innerHTML = '<p style="color:var(--gray-400);text-align:center;padding:40px;">no fixed assets yet. sync from QBO or add manually.</p>';
  } else {
    list.innerHTML = activeAssets.map(a => {
      const monthly = ((a.originalCost - a.salvageValue) / a.usefulLifeMonths).toFixed(2);
      const acqDate = new Date(a.acquisitionDate);
      const now = new Date();
      const monthsElapsed = (now.getFullYear() * 12 + now.getMonth()) - (acqDate.getFullYear() * 12 + acqDate.getMonth());
      const remaining = Math.max(0, a.usefulLifeMonths - monthsElapsed);
      return `
        <div class="client-card">
          <div class="client-info">
            <div class="client-name">${a.name}</div>
            <div class="asset-card-amounts">
              <span class="asset-card-amount">cost: <strong>$${a.originalCost.toLocaleString()}</strong></span>
              <span class="asset-card-amount">monthly: <strong>$${monthly}</strong></span>
              <span class="asset-card-amount">remaining: <strong>${remaining} mo</strong></span>
            </div>
            <div style="margin-top:4px;font-size:0.75rem;color:var(--gray-400);">${a.expenseAccountName || '—'} → ${a.accumAccountName || '—'}</div>
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
      <label class="qbo-sync-item" style="display:flex;align-items:center;gap:10px;padding:10px;border-bottom:1px solid var(--gray-100);cursor:pointer;">
        <input type="checkbox" class="qbo-sync-check" data-account='${JSON.stringify(a).replace(/'/g, "&#39;")}' checked>
        <div style="flex:1;">
          <div style="font-weight:500;">${a.name}</div>
          <div style="font-size:0.78rem;color:var(--gray-400);">${a.accountType} — balance: $${a.currentBalance.toLocaleString()}</div>
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
  errorEl.style.display = 'none';

  for (const cb of checkboxes) {
    const account = JSON.parse(cb.dataset.account);
    // Import each as a fixed asset — user will need to set amortization details later
    const body = {
      name: account.name,
      originalCost: account.currentBalance || 0,
      usefulLifeMonths: 60, // default, user should edit
      salvageValue: 0,
      acquisitionDate: new Date().toISOString().split('T')[0],
      assetAccountId: account.qboAccountId,
      assetAccountName: account.name,
      expenseAccountId: '',
      expenseAccountName: '',
      accumAccountId: '',
      accumAccountName: '',
      qboAccountId: account.qboAccountId,
    };

    try {
      await fetch(`/api/admin/clients/${selectedClientId}/fixed-assets`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': getAuth() },
        body: JSON.stringify(body),
      });
    } catch (e) { console.error('Failed to import asset:', e); }
  }

  document.getElementById('qbo-sync-modal').style.display = 'none';
  loadClientFixedAssets();
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
    qboAccounts.filter(a => cfg.types.includes(a.AccountType)).forEach(a => {
      const opt = document.createElement('option');
      opt.value = a.Id;
      opt.textContent = `${a.Name} (${a.AccountType})`;
      opt.dataset.name = a.Name;
      if (a.Id === cfg.selected) opt.selected = true;
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
  document.getElementById('asset-modal-title').textContent = 'edit fixed asset';
  document.getElementById('asset-name').value = asset.name;
  document.getElementById('asset-cost').value = asset.originalCost;
  document.getElementById('asset-useful-life').value = asset.usefulLifeMonths;
  document.getElementById('asset-salvage').value = asset.salvageValue;
  document.getElementById('asset-acq-date').value = asset.acquisitionDate;
  document.getElementById('asset-modal-error').style.display = 'none';
  document.getElementById('ai-suggestion').style.display = 'none';
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
document.getElementById('btn-add-asset').addEventListener('click', openAddAsset);
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
      statusEl.textContent = 'qbo not connected — connect via portal first';
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
  // Load accounts first so dropdowns are ready before rendering analyses
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

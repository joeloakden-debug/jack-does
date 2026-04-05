const OAuthClient = require('intuit-oauth');
const QuickBooks = require('node-quickbooks');
const fs = require('fs');
const path = require('path');

// ========================================
// OAUTH CLIENT SETUP
// ========================================
const oauthClient = new OAuthClient({
  clientId: process.env.QBO_CLIENT_ID,
  clientSecret: process.env.QBO_CLIENT_SECRET,
  environment: process.env.QBO_ENVIRONMENT || 'sandbox', // 'sandbox' or 'production'
  redirectUri: process.env.QBO_REDIRECT_URI,
});

// Token store keyed by clientId (e.g. "demo-client", "acme-corp")
const tokenStore = new Map();
const TOKEN_FILE = path.join(__dirname, '.qbo-tokens.json');

// Load tokens on startup (env var first, then disk file)
function loadTokensFromDisk() {
  console.log('[QBO] Loading tokens... QBO_TOKENS env var exists:', !!process.env.QBO_TOKENS, 'length:', (process.env.QBO_TOKENS || '').length);
  try {
    // Prefer QBO_TOKENS env var (survives Railway redeploys)
    if (process.env.QBO_TOKENS) {
      const raw = process.env.QBO_TOKENS;
      console.log('[QBO] Parsing QBO_TOKENS, first 50 chars:', raw.substring(0, 50));
      const data = JSON.parse(raw);
      const keys = Object.keys(data);
      console.log('[QBO] Parsed successfully, keys:', keys.join(', '));

      // Migration: if the old format has a "default" key, keep it as-is
      // New format is keyed by clientId
      Object.entries(data).forEach(([key, val]) => tokenStore.set(key, val));
      console.log('QBO tokens loaded from env var — connection restored');
      return;
    }
    if (fs.existsSync(TOKEN_FILE)) {
      const data = JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf-8'));
      Object.entries(data).forEach(([key, val]) => tokenStore.set(key, val));
      console.log('QBO tokens loaded from disk — connection restored');
    }
  } catch (e) {
    console.error('Failed to load QBO tokens:', e.message);
  }
}

// Save tokens to disk and log for env var persistence
function saveTokensToDisk() {
  try {
    const data = {};
    tokenStore.forEach((val, key) => { data[key] = val; });
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(data, null, 2));
    // Log tokens as JSON so they can be saved to QBO_TOKENS env var on Railway
    console.log('[QBO_TOKENS] Copy this to Railway env var to persist across deploys:');
    console.log('[QBO_TOKENS_VALUE]', JSON.stringify(data));
  } catch (e) {
    console.error('Failed to save QBO tokens to disk:', e.message);
  }
}

// Restore tokens on startup
loadTokensFromDisk();

// ========================================
// OAUTH FLOW (per-client)
// ========================================

/**
 * Generate the QuickBooks OAuth authorization URL
 * @param {string} clientId - The client ID to associate this QBO connection with
 */
function getAuthUri(clientId) {
  // Encode clientId in the state parameter so we can associate it after callback
  const state = clientId ? `jackdoes-qbo:${clientId}` : 'jackdoes-qbo:default';
  return oauthClient.authorizeUri({
    scope: [OAuthClient.scopes.Accounting, OAuthClient.scopes.OpenId],
    state,
  });
}

/**
 * Handle the OAuth callback and store tokens keyed by clientId
 * @param {string} url - The callback URL
 * @param {string} clientId - The client ID to store tokens under
 */
async function handleCallback(url, clientId) {
  const authResponse = await oauthClient.createToken(url);
  const token = authResponse.getJson();

  const tokenData = {
    accessToken: token.access_token,
    refreshToken: token.refresh_token,
    realmId: token.realmId || oauthClient.getToken().realmId,
    expiresAt: Date.now() + (token.expires_in * 1000),
    refreshExpiresAt: Date.now() + (token.x_refresh_token_expires_in * 1000),
    createdAt: new Date().toISOString(),
    clientId: clientId || 'default',
  };

  // Store by the clientId
  const storeKey = clientId || 'default';
  tokenStore.set(storeKey, tokenData);

  // Persist to disk so tokens survive server restarts
  saveTokensToDisk();

  return tokenData;
}

/**
 * Refresh the access token if expired
 * @param {string} clientId - The client whose token to refresh
 */
async function refreshTokenIfNeeded(clientId = 'default') {
  const tokenData = tokenStore.get(clientId);
  if (!tokenData) {
    throw new Error('No QuickBooks connection found for this client. Please connect QuickBooks first.');
  }

  // Check if token is expired (with 5 min buffer)
  if (Date.now() > tokenData.expiresAt - 300000) {
    try {
      oauthClient.setToken({
        access_token: tokenData.accessToken,
        refresh_token: tokenData.refreshToken,
        token_type: 'bearer',
        realmId: tokenData.realmId,
      });

      const authResponse = await oauthClient.refresh();
      const newToken = authResponse.getJson();

      tokenData.accessToken = newToken.access_token;
      tokenData.refreshToken = newToken.refresh_token;
      tokenData.expiresAt = Date.now() + (newToken.expires_in * 1000);

      tokenStore.set(clientId, tokenData);
      saveTokensToDisk();
    } catch (refreshError) {
      // If refresh fails, the tokens are invalid — clear them so user can reconnect
      console.error(`[QBO] Refresh token invalid for client "${clientId}", clearing connection. User must reconnect.`);
      disconnect(clientId);
      throw new Error('The Refresh token is invalid, please Authorize again.');
    }
  }

  return tokenData;
}

/**
 * Disconnect QuickBooks for a specific client
 * @param {string} clientId - The client to disconnect (or 'all' to clear everything)
 */
function disconnect(clientId) {
  if (clientId === 'all') {
    tokenStore.clear();
  } else if (clientId) {
    tokenStore.delete(clientId);
  } else {
    // Legacy: clear everything
    tokenStore.clear();
  }
  try {
    // Rewrite token file with remaining connections
    const data = {};
    tokenStore.forEach((val, key) => { data[key] = val; });
    if (Object.keys(data).length > 0) {
      fs.writeFileSync(TOKEN_FILE, JSON.stringify(data, null, 2));
    } else {
      if (fs.existsSync(TOKEN_FILE)) fs.unlinkSync(TOKEN_FILE);
    }
  } catch (e) { /* ignore */ }
  console.log(`[QBO] Disconnected client "${clientId}" — tokens cleared`);
}

/**
 * Check if QuickBooks is connected for a specific client
 * @param {string} clientId
 */
function isConnected(clientId = 'default') {
  return tokenStore.has(clientId);
}

/**
 * Check if ANY client has a QBO connection (useful for global status)
 */
function isAnyConnected() {
  return tokenStore.size > 0;
}

/**
 * Get all connected client IDs
 */
function getAllConnections() {
  const connections = {};
  tokenStore.forEach((val, key) => {
    connections[key] = {
      realmId: val.realmId,
      createdAt: val.createdAt,
      clientId: val.clientId || key,
    };
  });
  return connections;
}

/**
 * Get the full token store as a JSON string (for env var export)
 */
function getTokensJson() {
  const data = {};
  tokenStore.forEach((val, key) => { data[key] = val; });
  return JSON.stringify(data);
}

// ========================================
// QUICKBOOKS DATA QUERIES
// ========================================

/**
 * Get a configured QuickBooks client
 * @param {string} clientId - The client whose QBO connection to use
 */
async function getQBClient(clientId = 'default') {
  const tokenData = await refreshTokenIfNeeded(clientId);

  const useSandbox = (process.env.QBO_ENVIRONMENT || 'sandbox') === 'sandbox';

  return new QuickBooks(
    process.env.QBO_CLIENT_ID,
    process.env.QBO_CLIENT_SECRET,
    tokenData.accessToken,
    false, // no token secret for OAuth2
    tokenData.realmId,
    useSandbox,
    true, // debug
    null, // minor version
    '2.0', // OAuth version
    tokenData.refreshToken
  );
}

/**
 * Promisify QuickBooks callback-style methods
 */
function qbPromise(qb, method, ...args) {
  return new Promise((resolve, reject) => {
    qb[method](...args, (err, data) => {
      if (err) reject(err);
      else resolve(data);
    });
  });
}

/**
 * Get Profit & Loss report
 */
async function getProfitAndLoss(startDate, endDate, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportProfitAndLoss', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get Balance Sheet report
 */
async function getBalanceSheet(asOfDate, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportBalanceSheet', {
    date: asOfDate,
  });
}

/**
 * Get Cash Flow report
 */
async function getCashFlow(startDate, endDate, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportCashFlow', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get Aged Receivables report
 */
async function getAgedReceivables(clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportAgedReceivables', {});
}

/**
 * Get Aged Payables report
 */
async function getAgedPayables(clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportAgedPayables', {});
}

/**
 * Get Trial Balance
 */
async function getTrialBalance(startDate, endDate, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportTrialBalance', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get General Ledger
 */
async function getGeneralLedger(startDate, endDate, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportGeneralLedgerDetail', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get General Ledger Detail for specific account(s)
 * @param {string} accountIds - Comma-separated account IDs
 * @param {string} clientId
 */
async function getGeneralLedgerForAccount(accountIds, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'reportGeneralLedgerDetail', {
    account: accountIds,
    date_macro: 'All',
  });
}

/**
 * Get Company Info
 */
async function getCompanyInfo(clientId = 'default') {
  const qb = await getQBClient(clientId);
  const tokenData = tokenStore.get(clientId);
  return qbPromise(qb, 'getCompanyInfo', tokenData.realmId);
}

/**
 * Helper to build date filter criteria for find* methods
 */
function dateFilter(startDate, endDate) {
  const criteria = [];
  if (startDate) criteria.push({ field: 'TxnDate', value: startDate, operator: '>=' });
  if (endDate) criteria.push({ field: 'TxnDate', value: endDate, operator: '<=' });
  return criteria;
}

/**
 * Get recent invoices
 */
async function getRecentInvoices(limit = 10, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findInvoices', { fetchAll: true });
}

/**
 * Get recent expenses/purchases
 */
async function getRecentExpenses(limit = 10, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findPurchases', { fetchAll: true });
}

/**
 * Get customer list
 */
async function getCustomers(clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findCustomers', { fetchAll: true });
}

/**
 * Get vendor list
 */
async function getVendors(clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findVendors', { fetchAll: true });
}

/**
 * Get account list
 */
async function getAccounts(clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findAccounts', { fetchAll: true });
}

/**
 * Get all purchases/expenses in a date range
 */
async function getExpenseTransactions(startDate, endDate, limit = 100, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findPurchases', dateFilter(startDate, endDate));
}

/**
 * Get all bills in a date range
 */
async function getBills(startDate, endDate, limit = 100, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findBills', dateFilter(startDate, endDate));
}

/**
 * Get invoices in a date range
 */
async function getInvoicesByDate(startDate, endDate, limit = 100, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findInvoices', dateFilter(startDate, endDate));
}

/**
 * Get all journal entries in a date range
 */
async function getJournalEntries(startDate, endDate, limit = 100, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findJournalEntries', dateFilter(startDate, endDate));
}

/**
 * Get payments received in a date range
 */
async function getPayments(startDate, endDate, limit = 100, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findPayments', dateFilter(startDate, endDate));
}

/**
 * Get deposits in a date range
 */
async function getDeposits(startDate, endDate, limit = 100, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findDeposits', dateFilter(startDate, endDate));
}

/**
 * Get transfers in a date range
 */
async function getTransfers(startDate, endDate, limit = 100, clientId = 'default') {
  const qb = await getQBClient(clientId);
  return qbPromise(qb, 'findTransfers', dateFilter(startDate, endDate));
}

// ========================================
// SMART DATA FETCHER
// Determines what data to pull based on the user's question
// ========================================

/**
 * Analyze the question and fetch relevant QuickBooks data
 * @param {string} question - The user's question
 * @param {string} clientId - The client whose QBO data to fetch
 */
async function fetchRelevantData(question, clientId = 'default') {
  const q = question.toLowerCase();
  const now = new Date();
  const results = {};

  // Date helpers
  const firstOfThisMonth = new Date(now.getFullYear(), now.getMonth(), 1)
    .toISOString().split('T')[0];
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const firstOfLastMonth = lastMonth.toISOString().split('T')[0];
  const lastOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0)
    .toISOString().split('T')[0];
  const today = now.toISOString().split('T')[0];
  const startOfYear = `${now.getFullYear()}-01-01`;

  // Determine the relevant date range for transaction queries
  let txnStart = firstOfLastMonth;
  let txnEnd = today;
  if (q.includes('last month') || q.includes('previous month')) {
    txnStart = firstOfLastMonth;
    txnEnd = lastOfLastMonth;
  } else if (q.includes('this month')) {
    txnStart = firstOfThisMonth;
    txnEnd = today;
  } else if (q.includes('this year') || q.includes('ytd') || q.includes('year to date')) {
    txnStart = startOfYear;
    txnEnd = today;
  } else if (q.includes('last year') || q.includes('previous year')) {
    txnStart = `${now.getFullYear() - 1}-01-01`;
    txnEnd = `${now.getFullYear() - 1}-12-31`;
  }

  // Detect if user is asking about specific line items, transactions, or details
  const wantsDetail = q.match(/detail|breakdown|transaction|what.*made up|what.*in|drill.*down|line.*item|specific|individual|tell me more|more info|what.*include|what.*consist|explain.*expense|explain.*cost|break.*down|itemize|list.*transaction|show.*transaction|what.*charge|what.*pay.*for|each|every|subscript|licen/);

  // Helper to safely fetch data without one failure killing everything
  async function safeFetch(label, fn) {
    try {
      return await fn();
    } catch (e) {
      console.error(`[QBO] Failed to fetch ${label}:`, e.message);
      return null;
    }
  }

  // Always try to get company info for context
  results.companyInfo = await safeFetch('companyInfo', () => getCompanyInfo(clientId));

  // Profit / revenue / income questions
  if (q.match(/profit|revenue|income|earn|loss|p&l|p\+l|how.*doing|performance/)) {
    if (q.includes('last month') || q.includes('previous month')) {
      results.profitAndLoss = await safeFetch('P&L last month', () => getProfitAndLoss(firstOfLastMonth, lastOfLastMonth, clientId));
    } else if (q.includes('this month')) {
      results.profitAndLoss = await safeFetch('P&L this month', () => getProfitAndLoss(firstOfThisMonth, today, clientId));
    } else if (q.includes('this year') || q.includes('ytd') || q.includes('year to date')) {
      results.profitAndLoss = await safeFetch('P&L YTD', () => getProfitAndLoss(startOfYear, today, clientId));
    } else {
      results.profitAndLossLastMonth = await safeFetch('P&L last month', () => getProfitAndLoss(firstOfLastMonth, lastOfLastMonth, clientId));
      results.profitAndLossThisMonth = await safeFetch('P&L this month', () => getProfitAndLoss(firstOfThisMonth, today, clientId));
    }
  }

  // Balance sheet / assets / liabilities / equity
  if (q.match(/balance sheet|assets|liabilities|equity|net worth|what.*own|what.*owe/)) {
    results.balanceSheet = await safeFetch('balance sheet', () => getBalanceSheet(today, clientId));
  }

  // Cash flow
  if (q.match(/cash flow|cash.*forecast|liquidity|cash.*position|cash.*need/)) {
    results.cashFlow = await safeFetch('cash flow', () => getCashFlow(firstOfLastMonth, today, clientId));
    if (!results.balanceSheet) results.balanceSheet = await safeFetch('balance sheet', () => getBalanceSheet(today, clientId));
  }

  // Receivables / who owes / overdue / outstanding
  if (q.match(/receivable|owed|overdue|outstanding|unpaid|collect|aging|past due/)) {
    results.agedReceivables = await safeFetch('aged receivables', () => getAgedReceivables(clientId));
  }

  // Payables / bills / what we owe
  if (q.match(/payable|bills|what.*we owe|what.*i owe|vendor.*balance|accounts payable/)) {
    results.agedPayables = await safeFetch('aged payables', () => getAgedPayables(clientId));
    if (wantsDetail) {
      results.billTransactions = await safeFetch('bills', () => getBills(txnStart, txnEnd, 100, clientId));
    }
  }

  // Invoices
  if (q.match(/invoice|billed|billing/)) {
    results.recentInvoices = await safeFetch('invoices', () => getInvoicesByDate(txnStart, txnEnd, 50, clientId));
  }

  // Expenses / spending — always fetch transactions for detail
  if (q.match(/expense|spend|cost|purchase|bought|subscript|software|rent|office|utilit|insurance|advertising|marketing|meal|travel|phone|internet|licen/) || wantsDetail) {
    results.expenseTransactions = await safeFetch('expense transactions', () => getExpenseTransactions(txnStart, txnEnd, 100, clientId));
    if (!results.billTransactions) results.billTransactions = await safeFetch('bills', () => getBills(txnStart, txnEnd, 100, clientId));
    if (!results.profitAndLoss) {
      if (q.includes('last month') || q.includes('previous month')) {
        results.profitAndLoss = await safeFetch('P&L', () => getProfitAndLoss(firstOfLastMonth, lastOfLastMonth, clientId));
      } else {
        results.profitAndLoss = await safeFetch('P&L', () => getProfitAndLoss(txnStart, txnEnd, clientId));
      }
    }
  }

  // Customers
  if (q.match(/customer|client list|who.*buy/)) {
    results.customers = await safeFetch('customers', () => getCustomers(clientId));
  }

  // Vendors / suppliers
  if (q.match(/vendor|supplier|who.*pay/)) {
    results.vendors = await safeFetch('vendors', () => getVendors(clientId));
    if (wantsDetail && !results.billTransactions) {
      results.billTransactions = await safeFetch('bills', () => getBills(txnStart, txnEnd, 100, clientId));
    }
  }

  // Deposits / payments received
  if (q.match(/deposit|payment.*received|money.*in|received/)) {
    results.deposits = await safeFetch('deposits', () => getDeposits(txnStart, txnEnd, 50, clientId));
    results.payments = await safeFetch('payments', () => getPayments(txnStart, txnEnd, 50, clientId));
  }

  // Transfers between accounts
  if (q.match(/transfer|moved.*money|between.*account/)) {
    results.transfers = await safeFetch('transfers', () => getTransfers(txnStart, txnEnd, 50, clientId));
  }

  // Journal entries
  if (q.match(/journal entr|adjustment|accrual/)) {
    results.journalEntries = await safeFetch('journal entries', () => getJournalEntries(txnStart, txnEnd, 50, clientId));
  }

  // Tax related
  if (q.match(/tax|deduct|write.*off|1099|w-2/)) {
    if (!results.profitAndLoss) results.profitAndLoss = await safeFetch('P&L YTD', () => getProfitAndLoss(startOfYear, today, clientId));
    if (!results.balanceSheet) results.balanceSheet = await safeFetch('balance sheet', () => getBalanceSheet(today, clientId));
    if (!results.expenseTransactions) results.expenseTransactions = await safeFetch('expenses YTD', () => getExpenseTransactions(startOfYear, today, 100, clientId));
  }

  // General / overview / summary
  if (q.match(/overview|summary|how.*business|status|snapshot|dashboard/)) {
    if (!results.profitAndLoss) results.profitAndLoss = await safeFetch('P&L', () => getProfitAndLoss(firstOfThisMonth, today, clientId));
    if (!results.balanceSheet) results.balanceSheet = await safeFetch('balance sheet', () => getBalanceSheet(today, clientId));
    if (!results.agedReceivables) results.agedReceivables = await safeFetch('aged receivables', () => getAgedReceivables(clientId));
  }

  // If nothing matched, get P&L + transactions for a general overview
  const dataKeys = Object.keys(results).filter(k => results[k] != null);
  if (dataKeys.length <= 1) {
    results.profitAndLoss = await safeFetch('P&L default', () => getProfitAndLoss(txnStart, txnEnd, clientId));
    results.balanceSheet = await safeFetch('balance sheet default', () => getBalanceSheet(today, clientId));
    results.expenseTransactions = await safeFetch('expenses default', () => getExpenseTransactions(txnStart, txnEnd, 50, clientId));
  }

  // Clean out null results from failed fetches
  Object.keys(results).forEach(k => { if (results[k] == null) delete results[k]; });

  return results;
}

// ========================================
// CREATE JOURNAL ENTRY
// ========================================

/**
 * Create a journal entry in QuickBooks
 * @param {Object} entry - { date, memo, lines: [{ accountId, accountName, description, amount, type: 'debit'|'credit' }] }
 * @param {string} clientId - The client whose QBO to post to
 */
async function createJournalEntry(entry, clientId = 'default') {
  const qb = await getQBClient(clientId);

  const lines = entry.lines.map(line => ({
    JournalEntryLineDetail: {
      PostingType: line.type === 'debit' ? 'Debit' : 'Credit',
      AccountRef: {
        value: line.accountId,
        name: line.accountName,
      },
    },
    DetailType: 'JournalEntryLineDetail',
    Amount: Math.abs(line.amount),
    Description: line.description || '',
  }));

  const journalEntry = {
    TxnDate: entry.date,
    PrivateNote: entry.memo || 'Auto-generated by jack does',
    Line: lines,
  };

  return qbPromise(qb, 'createJournalEntry', journalEntry);
}

/**
 * Create a bill (vendor expense) in QuickBooks
 */
async function createBill(bill, clientId = 'default') {
  const qb = await getQBClient(clientId);

  const lines = bill.lines.map(line => ({
    DetailType: 'AccountBasedExpenseLineDetail',
    Amount: line.amount,
    AccountBasedExpenseLineDetail: {
      AccountRef: {
        value: line.accountId,
        name: line.accountName,
      },
    },
    Description: line.description || '',
  }));

  const billObj = {
    VendorRef: { value: bill.vendorId },
    TxnDate: bill.date,
    DueDate: bill.dueDate || bill.date,
    PrivateNote: bill.memo || 'Auto-generated by jack does',
    Line: lines,
  };

  return qbPromise(qb, 'createBill', billObj);
}

/**
 * Create an invoice in QuickBooks
 */
async function createInvoice(invoice, clientId = 'default') {
  const qb = await getQBClient(clientId);

  const lines = invoice.lines.map(line => ({
    DetailType: 'SalesItemLineDetail',
    Amount: line.amount,
    SalesItemLineDetail: {
      ItemRef: { value: line.itemId || '1', name: line.itemName || 'Services' },
    },
    Description: line.description || '',
  }));

  const invoiceObj = {
    CustomerRef: { value: invoice.customerId },
    TxnDate: invoice.date,
    DueDate: invoice.dueDate || invoice.date,
    PrivateNote: invoice.memo || 'Auto-generated by jack does',
    Line: lines,
  };

  return qbPromise(qb, 'createInvoice', invoiceObj);
}

module.exports = {
  getAuthUri,
  handleCallback,
  isConnected,
  isAnyConnected,
  disconnect,
  getAllConnections,
  getTokensJson,
  refreshTokenIfNeeded,
  fetchRelevantData,
  getProfitAndLoss,
  getBalanceSheet,
  getCashFlow,
  getAgedReceivables,
  getAgedPayables,
  getTrialBalance,
  getGeneralLedger,
  getGeneralLedgerForAccount,
  getCompanyInfo,
  getRecentInvoices,
  getRecentExpenses,
  getExpenseTransactions,
  getBills,
  getInvoicesByDate,
  getJournalEntries,
  getPayments,
  getDeposits,
  getTransfers,
  getCustomers,
  getVendors,
  getAccounts,
  createJournalEntry,
  createBill,
  createInvoice,
};

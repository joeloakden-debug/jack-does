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

// Token store with file persistence (survives server restarts)
const tokenStore = new Map();
const TOKEN_FILE = path.join(__dirname, '.qbo-tokens.json');

// Load tokens on startup (env var first, then disk file)
function loadTokensFromDisk() {
  try {
    // Prefer QBO_TOKENS env var (survives Railway redeploys)
    if (process.env.QBO_TOKENS) {
      const data = JSON.parse(process.env.QBO_TOKENS);
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
// OAUTH FLOW
// ========================================

/**
 * Generate the QuickBooks OAuth authorization URL
 */
function getAuthUri() {
  return oauthClient.authorizeUri({
    scope: [OAuthClient.scopes.Accounting, OAuthClient.scopes.OpenId],
    state: 'jackdoes-qbo',
  });
}

/**
 * Handle the OAuth callback and store tokens
 */
async function handleCallback(url) {
  const authResponse = await oauthClient.createToken(url);
  const token = authResponse.getJson();

  const tokenData = {
    accessToken: token.access_token,
    refreshToken: token.refresh_token,
    realmId: token.realmId || oauthClient.getToken().realmId,
    expiresAt: Date.now() + (token.expires_in * 1000),
    refreshExpiresAt: Date.now() + (token.x_refresh_token_expires_in * 1000),
    createdAt: new Date().toISOString(),
  };

  // Store by realmId (company ID)
  tokenStore.set(tokenData.realmId, tokenData);

  // Also store as "default" for easy access
  tokenStore.set('default', tokenData);

  // Persist to disk so tokens survive server restarts
  saveTokensToDisk();

  return tokenData;
}

/**
 * Refresh the access token if expired
 */
async function refreshTokenIfNeeded(realmId = 'default') {
  const tokenData = tokenStore.get(realmId);
  if (!tokenData) {
    throw new Error('No QuickBooks connection found. Please connect QuickBooks first.');
  }

  // Check if token is expired (with 5 min buffer)
  if (Date.now() > tokenData.expiresAt - 300000) {
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

    tokenStore.set(realmId, tokenData);
    if (realmId !== 'default') tokenStore.set('default', tokenData);
    saveTokensToDisk();
  }

  return tokenData;
}

/**
 * Check if QuickBooks is connected
 */
function isConnected(realmId = 'default') {
  return tokenStore.has(realmId);
}

// ========================================
// QUICKBOOKS DATA QUERIES
// ========================================

/**
 * Get a configured QuickBooks client
 */
async function getQBClient(realmId = 'default') {
  const tokenData = await refreshTokenIfNeeded(realmId);

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
async function getProfitAndLoss(startDate, endDate, realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'reportProfitAndLoss', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get Balance Sheet report
 */
async function getBalanceSheet(asOfDate, realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'reportBalanceSheet', {
    date: asOfDate,
  });
}

/**
 * Get Cash Flow report
 */
async function getCashFlow(startDate, endDate, realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'reportCashFlow', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get Aged Receivables report
 */
async function getAgedReceivables(realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'reportAgedReceivables', {});
}

/**
 * Get Aged Payables report
 */
async function getAgedPayables(realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'reportAgedPayables', {});
}

/**
 * Get Trial Balance
 */
async function getTrialBalance(startDate, endDate, realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'reportTrialBalance', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get General Ledger
 */
async function getGeneralLedger(startDate, endDate, realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'reportGeneralLedgerDetail', {
    start_date: startDate,
    end_date: endDate,
  });
}

/**
 * Get Company Info
 */
async function getCompanyInfo(realmId = 'default') {
  const qb = await getQBClient(realmId);
  const tokenData = tokenStore.get(realmId);
  return qbPromise(qb, 'getCompanyInfo', tokenData.realmId);
}

/**
 * Query any entity (invoices, expenses, customers, etc.)
 */
async function query(sql, realmId = 'default') {
  const qb = await getQBClient(realmId);
  return new Promise((resolve, reject) => {
    qb.query(sql, (err, data) => {
      if (err) reject(err);
      else resolve(data);
    });
  });
}

/**
 * Get recent invoices
 */
async function getRecentInvoices(limit = 10, realmId = 'default') {
  return query(
    `SELECT * FROM Invoice ORDER BY MetaData.CreateTime DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get recent expenses/purchases
 */
async function getRecentExpenses(limit = 10, realmId = 'default') {
  return query(
    `SELECT * FROM Purchase ORDER BY MetaData.CreateTime DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get customer list
 */
async function getCustomers(realmId = 'default') {
  return query('SELECT * FROM Customer MAXRESULTS 100', realmId);
}

/**
 * Get vendor list
 */
async function getVendors(realmId = 'default') {
  return query('SELECT * FROM Vendor MAXRESULTS 100', realmId);
}

/**
 * Get account list
 */
async function getAccounts(realmId = 'default') {
  const qb = await getQBClient(realmId);
  return qbPromise(qb, 'findAccounts', { fetchAll: true });
}

/**
 * Get all purchases/expenses in a date range
 */
async function getExpenseTransactions(startDate, endDate, limit = 100, realmId = 'default') {
  return query(
    `SELECT * FROM Purchase WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' ORDER BY TxnDate DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get all bills in a date range
 */
async function getBills(startDate, endDate, limit = 100, realmId = 'default') {
  return query(
    `SELECT * FROM Bill WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' ORDER BY TxnDate DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get invoices in a date range
 */
async function getInvoicesByDate(startDate, endDate, limit = 100, realmId = 'default') {
  return query(
    `SELECT * FROM Invoice WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' ORDER BY TxnDate DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get all journal entries in a date range
 */
async function getJournalEntries(startDate, endDate, limit = 100, realmId = 'default') {
  return query(
    `SELECT * FROM JournalEntry WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' ORDER BY TxnDate DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get payments received in a date range
 */
async function getPayments(startDate, endDate, limit = 100, realmId = 'default') {
  return query(
    `SELECT * FROM Payment WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' ORDER BY TxnDate DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get deposits in a date range
 */
async function getDeposits(startDate, endDate, limit = 100, realmId = 'default') {
  return query(
    `SELECT * FROM Deposit WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' ORDER BY TxnDate DESC MAXRESULTS ${limit}`,
    realmId
  );
}

/**
 * Get transfers in a date range
 */
async function getTransfers(startDate, endDate, limit = 100, realmId = 'default') {
  return query(
    `SELECT * FROM Transfer WHERE TxnDate >= '${startDate}' AND TxnDate <= '${endDate}' ORDER BY TxnDate DESC MAXRESULTS ${limit}`,
    realmId
  );
}

// ========================================
// SMART DATA FETCHER
// Determines what data to pull based on the user's question
// ========================================

/**
 * Analyze the question and fetch relevant QuickBooks data
 */
async function fetchRelevantData(question, realmId = 'default') {
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

  try {
    // Always try to get company info for context
    results.companyInfo = await getCompanyInfo(realmId);
  } catch (e) {
    // Not critical
  }

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
  const wantsDetail = q.match(/detail|breakdown|transaction|what.*made up|what.*in|drill.*down|line.*item|specific|individual|tell me more|more info|what.*include|what.*consist|explain.*expense|explain.*cost|break.*down|itemize|list.*transaction|show.*transaction|what.*charge|what.*pay.*for|each|every|subscript/);

  try {
    // Profit / revenue / income questions
    if (q.match(/profit|revenue|income|earn|loss|p&l|p\+l|how.*doing|performance/)) {
      if (q.includes('last month') || q.includes('previous month')) {
        results.profitAndLoss = await getProfitAndLoss(firstOfLastMonth, lastOfLastMonth, realmId);
      } else if (q.includes('this month')) {
        results.profitAndLoss = await getProfitAndLoss(firstOfThisMonth, today, realmId);
      } else if (q.includes('this year') || q.includes('ytd') || q.includes('year to date')) {
        results.profitAndLoss = await getProfitAndLoss(startOfYear, today, realmId);
      } else {
        results.profitAndLossLastMonth = await getProfitAndLoss(firstOfLastMonth, lastOfLastMonth, realmId);
        results.profitAndLossThisMonth = await getProfitAndLoss(firstOfThisMonth, today, realmId);
      }
    }

    // Balance sheet / assets / liabilities / equity
    if (q.match(/balance sheet|assets|liabilities|equity|net worth|what.*own|what.*owe/)) {
      results.balanceSheet = await getBalanceSheet(today, realmId);
    }

    // Cash flow
    if (q.match(/cash flow|cash.*forecast|liquidity|cash.*position|cash.*need/)) {
      results.cashFlow = await getCashFlow(firstOfLastMonth, today, realmId);
      results.balanceSheet = results.balanceSheet || await getBalanceSheet(today, realmId);
    }

    // Receivables / who owes / overdue / outstanding
    if (q.match(/receivable|owed|overdue|outstanding|unpaid|collect|aging|past due/)) {
      results.agedReceivables = await getAgedReceivables(realmId);
    }

    // Payables / bills / what we owe
    if (q.match(/payable|bills|what.*we owe|what.*i owe|vendor.*balance|accounts payable/)) {
      results.agedPayables = await getAgedPayables(realmId);
      if (wantsDetail) {
        results.billTransactions = await getBills(txnStart, txnEnd, 100, realmId);
      }
    }

    // Invoices
    if (q.match(/invoice|billed|billing/)) {
      results.recentInvoices = await getInvoicesByDate(txnStart, txnEnd, 50, realmId);
    }

    // Expenses / spending — always fetch transactions for detail
    if (q.match(/expense|spend|cost|purchase|bought|subscript|software|rent|office|utilit|insurance|advertising|marketing|meal|travel|phone|internet/) || wantsDetail) {
      results.expenseTransactions = await getExpenseTransactions(txnStart, txnEnd, 100, realmId);
      results.billTransactions = results.billTransactions || await getBills(txnStart, txnEnd, 100, realmId);
      if (!results.profitAndLoss) {
        if (q.includes('last month') || q.includes('previous month')) {
          results.profitAndLoss = await getProfitAndLoss(firstOfLastMonth, lastOfLastMonth, realmId);
        } else {
          results.profitAndLoss = await getProfitAndLoss(txnStart, txnEnd, realmId);
        }
      }
    }

    // Customers
    if (q.match(/customer|client list|who.*buy/)) {
      results.customers = await getCustomers(realmId);
    }

    // Vendors / suppliers
    if (q.match(/vendor|supplier|who.*pay/)) {
      results.vendors = await getVendors(realmId);
      if (wantsDetail) {
        results.billTransactions = results.billTransactions || await getBills(txnStart, txnEnd, 100, realmId);
      }
    }

    // Deposits / payments received
    if (q.match(/deposit|payment.*received|money.*in|received/)) {
      results.deposits = await getDeposits(txnStart, txnEnd, 50, realmId);
      results.payments = await getPayments(txnStart, txnEnd, 50, realmId);
    }

    // Transfers between accounts
    if (q.match(/transfer|moved.*money|between.*account/)) {
      results.transfers = await getTransfers(txnStart, txnEnd, 50, realmId);
    }

    // Journal entries
    if (q.match(/journal entr|adjustment|accrual/)) {
      results.journalEntries = await getJournalEntries(txnStart, txnEnd, 50, realmId);
    }

    // Tax related
    if (q.match(/tax|deduct|write.*off|1099|w-2/)) {
      results.profitAndLoss = results.profitAndLoss ||
        await getProfitAndLoss(startOfYear, today, realmId);
      results.balanceSheet = results.balanceSheet ||
        await getBalanceSheet(today, realmId);
      results.expenseTransactions = results.expenseTransactions ||
        await getExpenseTransactions(startOfYear, today, 100, realmId);
    }

    // General / overview / summary
    if (q.match(/overview|summary|how.*business|status|snapshot|dashboard/)) {
      results.profitAndLoss = results.profitAndLoss ||
        await getProfitAndLoss(firstOfThisMonth, today, realmId);
      results.balanceSheet = results.balanceSheet ||
        await getBalanceSheet(today, realmId);
      results.agedReceivables = results.agedReceivables ||
        await getAgedReceivables(realmId);
    }

    // If nothing matched, get P&L + transactions for a general overview
    if (Object.keys(results).length <= 1) {
      results.profitAndLoss = await getProfitAndLoss(txnStart, txnEnd, realmId);
      results.balanceSheet = await getBalanceSheet(today, realmId);
      results.expenseTransactions = await getExpenseTransactions(txnStart, txnEnd, 50, realmId);
    }

  } catch (error) {
    results.error = `Error fetching data: ${error.message}`;
  }

  return results;
}

// ========================================
// CREATE JOURNAL ENTRY
// ========================================

/**
 * Create a journal entry in QuickBooks
 * @param {Object} entry - { date, memo, lines: [{ accountId, accountName, description, amount, type: 'debit'|'credit' }] }
 */
async function createJournalEntry(entry, realmId = 'default') {
  const qb = await getQBClient(realmId);

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
async function createBill(bill, realmId = 'default') {
  const qb = await getQBClient(realmId);

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
async function createInvoice(invoice, realmId = 'default') {
  const qb = await getQBClient(realmId);

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
  refreshTokenIfNeeded,
  fetchRelevantData,
  getProfitAndLoss,
  getBalanceSheet,
  getCashFlow,
  getAgedReceivables,
  getAgedPayables,
  getTrialBalance,
  getGeneralLedger,
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
  query,
  createJournalEntry,
  createBill,
  createInvoice,
};

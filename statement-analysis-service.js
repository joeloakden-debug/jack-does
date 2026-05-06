// Bank / credit-card statement ingestion + analysis service.
//
// Used by both the admin portal and the client portal so that any future
// business-logic enhancements (statement parsing rules, account-mapping
// heuristics, period reconciliation, etc.) live in exactly one place. The
// service is intentionally dependency-injected — callers pass the
// anthropic SDK client and the qbo module so the service has no global
// dependencies and is straightforward to swap in tests.
//
// Public surface:
//   init({ dataDir, anthropic, qbo, uploadsRoot })
//   listAnalyses(clientId)              -> [summary]
//   getAnalysis(clientId, analysisId)   -> full record or null
//   analyzeFile({ clientId, filePath, fileName, category, closeMonth })
//                                       -> { analysisId, analysis, record }
//   setStatus(clientId, analysisId, status)  -> updated record or null
//   removeAnalysis(clientId, analysisId) -> bool
//
// Persistence shape (data/bank-statement-analyses.json):
//   { "<client-id>": { analyses: [
//       { id, fileName, filePath, category, closeMonth, createdAt, status,
//         analysis: { documentType, summary, vendor, customer, date,
//                     totalAmount, currency, entries, notes, confidence,
//                     needsReview } }
//   ] } }

const fs = require('fs');
const path = require('path');

let _state = {
  dataFile: null,
  uploadsRoot: null,
  anthropic: null,
  qbo: null,
  data: null, // in-memory cache, mirrors dataFile
};

function init({ dataDir, anthropic, qbo, uploadsRoot }) {
  _state.dataFile = path.join(dataDir, 'bank-statement-analyses.json');
  _state.uploadsRoot = uploadsRoot;
  _state.anthropic = anthropic;
  _state.qbo = qbo;
  _state.data = _load();
}

function _load() {
  try {
    if (fs.existsSync(_state.dataFile)) {
      return JSON.parse(fs.readFileSync(_state.dataFile, 'utf-8'));
    }
  } catch (e) {
    console.error('[statement-analysis] failed to load:', e.message);
  }
  return {};
}

function _save() {
  if (!_state.dataFile) throw new Error('statement-analysis-service not initialized');
  // Atomic write — same pattern the rest of the app uses for state files.
  const tmp = `${_state.dataFile}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(_state.data, null, 2), 'utf-8');
  fs.renameSync(tmp, _state.dataFile);
}

function _clientBucket(clientId) {
  if (!_state.data[clientId]) _state.data[clientId] = { analyses: [] };
  if (!Array.isArray(_state.data[clientId].analyses)) _state.data[clientId].analyses = [];
  return _state.data[clientId];
}

function _summary(record) {
  const a = record.analysis || {};
  return {
    id: record.id,
    fileName: record.fileName,
    category: record.category,
    closeMonth: record.closeMonth || null,
    status: record.status,
    createdAt: record.createdAt,
    documentType: a.documentType || null,
    vendor: a.vendor || null,
    summary: a.summary || null,
    totalAmount: a.totalAmount || null,
    currency: a.currency || null,
    entryCount: Array.isArray(a.entries) ? a.entries.length : 0,
    confidence: a.confidence || null,
    needsReview: Array.isArray(a.needsReview) ? a.needsReview.length : 0,
  };
}

function listAnalyses(clientId) {
  const bucket = _clientBucket(clientId);
  return bucket.analyses
    .map(_summary)
    .sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
}

function getAnalysis(clientId, analysisId) {
  const bucket = _clientBucket(clientId);
  return bucket.analyses.find(r => r.id === analysisId) || null;
}

function setStatus(clientId, analysisId, status) {
  if (!['pending', 'approved', 'rejected', 'dismissed'].includes(status)) {
    throw new Error(`Invalid status: ${status}`);
  }
  const bucket = _clientBucket(clientId);
  const rec = bucket.analyses.find(r => r.id === analysisId);
  if (!rec) return null;
  rec.status = status;
  rec.updatedAt = new Date().toISOString();
  _save();
  return rec;
}

function removeAnalysis(clientId, analysisId) {
  const bucket = _clientBucket(clientId);
  const idx = bucket.analyses.findIndex(r => r.id === analysisId);
  if (idx === -1) return false;
  bucket.analyses.splice(idx, 1);
  _save();
  return true;
}

// Same prompt the legacy /api/process-document route uses. Kept verbatim
// here so any future business-logic tuning lands in this one place.
const STATEMENT_PROMPT = `You are "jack", an expert AI accountant. The user has uploaded a bank or credit-card statement (or another accounting document) for analysis.

Your job is to:
1. Identify each transaction or relevant entry in the document.
2. For each transaction, propose the appropriate journal entry (debits and credits) using the client's actual chart of accounts when provided.
3. Return a structured JSON response.

Respond with ONLY valid JSON in this exact format (no markdown, no code fences, no explanation outside the JSON):
{
  "documentType": "bank_statement|credit_card_statement|invoice|receipt|expense_report|payroll|other",
  "summary": "Brief human-readable summary of what this document is",
  "vendor": "Counterparty / institution name if applicable",
  "customer": "Customer name if applicable",
  "date": "YYYY-MM-DD primary date (statement period end for statements)",
  "periodStart": "YYYY-MM-DD if a statement period start is identifiable",
  "periodEnd": "YYYY-MM-DD if a statement period end is identifiable",
  "totalAmount": 123.45,
  "currency": "USD",
  "entries": [
    {
      "type": "journal_entry",
      "date": "YYYY-MM-DD",
      "memo": "Description of the transaction",
      "lines": [
        {
          "accountName": "Exact account name from chart of accounts when provided",
          "accountId": "QBO account ID if chart of accounts was supplied",
          "accountCategory": "Expense|Revenue|Asset|Liability|Equity",
          "description": "Line item description",
          "amount": 123.45,
          "type": "debit|credit"
        }
      ]
    }
  ],
  "notes": "Any additional notes or things the user should be aware of",
  "confidence": "high|medium|low",
  "needsReview": ["List of items that need human review or clarification"]
}

Rules:
- For BANK STATEMENTS: one entry per transaction. Outflows debit the appropriate expense / asset / payable account and credit the bank (cash) account. Inflows debit the bank account and credit the appropriate revenue / receivable / equity account.
- For CREDIT-CARD STATEMENTS: one entry per transaction. Charges debit the appropriate expense and credit the credit-card liability account. Payments to the card debit the credit-card liability and credit the bank account. Refunds debit the credit-card liability and credit the original expense account (negative expense effect).
- For INVOICES received (bills): debit the appropriate expense account, credit Accounts Payable.
- For INVOICES issued: debit Accounts Receivable, credit the appropriate revenue account.
- For RECEIPTS: debit the appropriate expense account, credit Cash/Bank.
- For PAYROLL: debit Salary/Wage Expense, credit Payroll Liabilities.
- Always ensure debits = credits for each entry.
- When a chart of accounts is provided, you MUST use exact account names from the list and include the corresponding accountId on every line.
- If you can't determine something with confidence, add it to needsReview.`;

async function _buildMessageContent({ filePath, fileName, category }) {
  const ext = path.extname(fileName).toLowerCase();
  if (['.jpg', '.jpeg', '.png'].includes(ext)) {
    const fileBuffer = fs.readFileSync(filePath);
    const base64 = fileBuffer.toString('base64');
    const mediaType = ext === '.png' ? 'image/png' : 'image/jpeg';
    return [
      { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64 } },
      { type: 'text', text: `This is an uploaded ${category || 'accounting'} document named "${fileName}". Please analyze it and determine the accounting entries needed. Respond with JSON only.` },
    ];
  }
  if (ext === '.pdf') {
    const fileBuffer = fs.readFileSync(filePath);
    const base64 = fileBuffer.toString('base64');
    return [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } },
      { type: 'text', text: `This is an uploaded PDF document named "${fileName}" in the "${category || 'general'}" category. Please analyze it and determine the accounting entries needed. Respond with JSON only.` },
    ];
  }
  if (ext === '.csv') {
    const content = fs.readFileSync(filePath, 'utf-8');
    return [
      { type: 'text', text: `This is a CSV file named "${fileName}" in the "${category || 'general'}" category. Here is the content:\n\n${content}\n\nPlease analyze it and determine the accounting entries needed. Respond with JSON only.` },
    ];
  }
  // Other text-readable files
  let content;
  try {
    content = fs.readFileSync(filePath, 'utf-8');
  } catch {
    throw Object.assign(new Error(`Cannot read file type ${ext}. Supported: PDF, CSV, JPG, PNG`), { statusCode: 400 });
  }
  return [
    { type: 'text', text: `This is a document named "${fileName}" in the "${category || 'general'}" category. Here is the content:\n\n${content}\n\nPlease analyze it and determine the accounting entries needed. Respond with JSON only.` },
  ];
}

async function _accountContextFor(clientId) {
  if (!_state.qbo || !_state.qbo.isConnected || !_state.qbo.isConnected(clientId)) return '';
  try {
    const accountsData = await _state.qbo.getAccounts(clientId);
    const list = accountsData?.QueryResponse?.Account || accountsData || [];
    const accounts = Array.isArray(list) ? list : [];
    if (accounts.length === 0) return '';
    const acctList = accounts
      .filter(a => a.Active !== false)
      .map(a => `- "${a.Name}" (ID: ${a.Id}, Type: ${a.AccountType})`)
      .join('\n');
    return `\n\nIMPORTANT — CHART OF ACCOUNTS:\nBelow is the client's actual QuickBooks chart of accounts. You MUST use these exact account names and include the "accountId" field for each journal entry line. Pick the most appropriate account for each line item based on the document content.\n\n${acctList}\n\nFor every line in your entries, set "accountName" to the exact account name from this list and "accountId" to the corresponding ID.`;
  } catch (e) {
    console.error('[statement-analysis] could not fetch QBO accounts:', e.message);
    return '';
  }
}

/**
 * Analyze a previously-uploaded file and persist the analysis.
 *
 * Path safety: filePath must resolve inside uploadsRoot (set at init).
 * That guards against `..` traversal regardless of which route called us.
 */
async function analyzeFile({ clientId, filePath, fileName, category, closeMonth }) {
  if (!clientId) throw Object.assign(new Error('clientId is required'), { statusCode: 400 });
  if (!filePath || !fileName) throw Object.assign(new Error('filePath and fileName are required'), { statusCode: 400 });
  if (!_state.uploadsRoot) throw new Error('statement-analysis-service not initialized');

  const uploadsRootReal = fs.realpathSync(_state.uploadsRoot);
  let fullPath;
  try {
    fullPath = fs.realpathSync(path.resolve(filePath));
  } catch (_) {
    throw Object.assign(new Error('File not found'), { statusCode: 404 });
  }
  const rel = path.relative(uploadsRootReal, fullPath);
  if (rel.startsWith('..') || path.isAbsolute(rel)) {
    throw Object.assign(new Error('Path outside of uploads directory'), { statusCode: 403 });
  }
  if (!fs.existsSync(fullPath)) {
    throw Object.assign(new Error('File not found'), { statusCode: 404 });
  }

  const messageContent = await _buildMessageContent({ filePath: fullPath, fileName, category });
  const accountContext = await _accountContextFor(clientId);

  const response = await _state.anthropic.messages.create({
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4096,
    system: STATEMENT_PROMPT + accountContext,
    messages: [{ role: 'user', content: messageContent }],
  });

  const rawText = response.content[0].text;
  let analysis;
  try {
    const jsonMatch = rawText.match(/\{[\s\S]*\}/);
    analysis = JSON.parse(jsonMatch ? jsonMatch[0] : rawText);
  } catch (parseErr) {
    const err = new Error('Could not parse document analysis');
    err.statusCode = 502;
    err.rawResponse = rawText;
    throw err;
  }

  const id = 'stmt_' + Date.now() + '_' + Math.random().toString(36).substring(2, 8);
  const record = {
    id,
    clientId,
    fileName,
    filePath: fullPath,
    category: category || 'bank_statement',
    closeMonth: closeMonth || null,
    createdAt: new Date().toISOString(),
    status: 'pending',
    analysis,
  };
  const bucket = _clientBucket(clientId);
  bucket.analyses.push(record);
  _save();

  return { analysisId: id, analysis, record };
}

module.exports = {
  init,
  listAnalyses,
  getAnalysis,
  analyzeFile,
  setStatus,
  removeAnalysis,
  // exposed for tests
  STATEMENT_PROMPT,
};

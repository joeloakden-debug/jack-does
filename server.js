require('dotenv').config();
const express = require('express');
const Anthropic = require('@anthropic-ai/sdk').default;
const multer = require('multer');
const { marked } = require('marked');

// Configure marked for clean HTML output
marked.setOptions({
  breaks: true,
  gfm: true,
});
const path = require('path');
const fs = require('fs');
const qbo = require('./qbo-service');
const excelService = require('./excel-service');

// Persistent data directory — set DATA_DIR env var to a Railway volume mount
// (e.g. /data) so files survive deploys. Falls back to __dirname for local dev.
const DATA_DIR = process.env.DATA_DIR || __dirname;
if (DATA_DIR !== __dirname) {
  console.log(`[data] using persistent data directory: ${DATA_DIR}`);
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

  // Seed: copy any committed data files into the volume if they don't exist there yet
  const seedFiles = ['clients.json', 'fixed-assets.json', 'prepaid-expenses.json', 'accrued-liabilities.json', 'shareholder-invoices.json'];
  for (const f of seedFiles) {
    const dest = path.join(DATA_DIR, f);
    const src = path.join(__dirname, f);
    if (!fs.existsSync(dest) && fs.existsSync(src)) {
      fs.copyFileSync(src, dest);
      console.log(`[data] seeded ${f} from repo into ${DATA_DIR}`);
    }
  }
}
function dataPath(filename) { return path.join(DATA_DIR, filename); }
const excelReviewService = require('./excel-review-service');
const prepaidReviewService = require('./prepaid-review-service');

const app = express();
const PORT = parseInt(process.env.PORT, 10) || 3000;
console.log(`[startup] PORT env var = "${process.env.PORT}", using port ${PORT}`);

// ========================================
// MIDDLEWARE
// ========================================
app.use(express.json());
app.use(express.static(__dirname)); // Serve the main site
app.use('/portal', express.static(path.join(__dirname, 'portal')));

// ========================================
// CLIENT REGISTRY (file-backed persistence)
// ========================================
const CLIENTS_FILE = dataPath('clients.json');

function loadClients() {
  try {
    if (fs.existsSync(CLIENTS_FILE)) {
      return JSON.parse(fs.readFileSync(CLIENTS_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Failed to load clients.json:', e.message);
  }
  return {
    'demo-client': { name: 'Demo Client', email: 'demo@company.com', billingFrequency: 'monthly', createdAt: new Date().toISOString() }
  };
}

function saveClients(clients) {
  fs.writeFileSync(CLIENTS_FILE, JSON.stringify(clients, null, 2), 'utf-8');
}

let CLIENTS = loadClients();

/**
 * Middleware: resolve which client is making the request.
 * Checks header, query param, or body. Falls back to 'demo-client'.
 * In production, this will extract clientId from a JWT or session token.
 */
function resolveClient(req, res, next) {
  // Check portal cookie first, then header/query/body
  const cookie = req.headers.cookie?.split(';')
    .find(c => c.trim().startsWith('portal_client='));
  const cookieClientId = cookie?.split('=')[1];
  const clientId = cookieClientId || req.headers['x-client-id'] || req.query.clientId || req.body?.clientId;
  if (!clientId || !CLIENTS[clientId]) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  req.clientId = clientId;
  req.clientName = CLIENTS[clientId].name;
  next();
}

// Admin auth — file-backed, falls back to env variable
const ADMIN_SETTINGS_FILE = dataPath('.admin-settings.json');

function getAdminPassword() {
  try {
    if (fs.existsSync(ADMIN_SETTINGS_FILE)) {
      const settings = JSON.parse(fs.readFileSync(ADMIN_SETTINGS_FILE, 'utf-8'));
      if (settings.password) return settings.password;
    }
  } catch (e) { /* fall through */ }
  return process.env.ADMIN_PASSWORD || 'jackdoes2026';
}

function saveAdminPassword(password) {
  fs.writeFileSync(ADMIN_SETTINGS_FILE, JSON.stringify({ password }, null, 2), 'utf-8');
}

function requireAdmin(req, res, next) {
  // Check for auth header (for API calls) or session cookie
  const authHeader = req.headers.authorization;
  const cookieAuth = req.headers.cookie?.split(';')
    .find(c => c.trim().startsWith('admin_auth='))
    ?.split('=')[1];

  const adminPw = getAdminPassword();
  if (authHeader === adminPw || cookieAuth === adminPw) {
    return next();
  }

  // If it's an API call, return 401
  if (req.path.startsWith('/api/admin')) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  // Otherwise serve login page
  return res.sendFile(path.join(__dirname, 'admin', 'login.html'));
}

// Admin static files (CSS/JS served without auth)
app.use('/admin/admin.css', express.static(path.join(__dirname, 'admin', 'admin.css')));
app.use('/admin/admin.js', express.static(path.join(__dirname, 'admin', 'admin.js')));

// Admin login endpoint
app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  if (password === getAdminPassword()) {
    res.json({ success: true });
  } else {
    res.status(401).json({ error: 'Invalid password' });
  }
});

// Admin: Change password
app.post('/api/admin/change-password', requireAdmin, (req, res) => {
  const { currentPassword, newPassword } = req.body;
  if (currentPassword !== getAdminPassword()) {
    return res.status(401).json({ error: 'Current password is incorrect' });
  }
  if (!newPassword || newPassword.length < 6) {
    return res.status(400).json({ error: 'New password must be at least 6 characters' });
  }
  saveAdminPassword(newPassword);
  res.json({ success: true });
});

// Admin pages require auth
app.get('/admin', requireAdmin, (req, res) => {
  res.sendFile(path.join(__dirname, 'admin', 'dashboard.html'));
});
app.get('/admin/dashboard.html', requireAdmin, (req, res) => {
  res.sendFile(path.join(__dirname, 'admin', 'dashboard.html'));
});

// ========================================
// CLIENT PORTAL AUTH
// ========================================

// Portal login — authenticates by email + password, returns clientId
app.post('/api/portal/login', (req, res) => {
  const { email, password } = req.body;
  const entry = Object.entries(CLIENTS).find(
    ([, c]) => c.email === email && c.password === password
  );
  if (entry) {
    const [clientId, client] = entry;
    res.json({ success: true, clientId, name: client.name });
  } else {
    res.status(401).json({ error: 'Invalid email or password' });
  }
});

// Middleware: require authenticated client (checks cookie)
function requireClient(req, res, next) {
  const cookie = req.headers.cookie?.split(';')
    .find(c => c.trim().startsWith('portal_client='));
  const clientId = cookie?.split('=')[1];
  if (clientId && CLIENTS[clientId]) {
    req.clientId = clientId;
    req.clientName = CLIENTS[clientId].name;
    return next();
  }
  // For API calls return 401, otherwise redirect to login
  if (req.path.startsWith('/api/')) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  return res.redirect('/portal/');
}

// Protect portal dashboard
app.get('/portal/dashboard.html', requireClient, (req, res) => {
  res.sendFile(path.join(__dirname, 'portal', 'dashboard.html'));
});

// ========================================
// CLAUDE API CLIENT
// ========================================
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

/** Call anthropic.messages.create with automatic retry on 429/529 errors */
async function claudeWithRetry(params, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return await anthropic.messages.create(params);
    } catch (e) {
      const status = e?.status || e?.statusCode || 0;
      const isRetryable = status === 429 || status === 529 || (e.message && e.message.includes('Overloaded'));
      if (isRetryable && attempt < maxRetries) {
        const delay = attempt * 5000; // 5s, 10s, 15s
        console.log(`[claude] ${status} on attempt ${attempt}/${maxRetries}, retrying in ${delay / 1000}s...`);
        await new Promise(r => setTimeout(r, delay));
        continue;
      }
      throw e;
    }
  }
}

const JACK_SYSTEM_PROMPT = `You are "jack", the AI accountant for "jack does" — a modern accounting firm.

Your personality:
- Friendly, approachable, and casual (use lowercase, conversational tone)
- Knowledgeable about accounting, bookkeeping, taxes, and finance
- You speak in first person as "jack" — never break character
- You're confident but always remind clients to verify critical financial decisions with a qualified professional
- You're enthusiastic about helping and making accounting feel less intimidating

Your capabilities:
- Answering questions about accounting principles, bookkeeping, taxes, and financial reporting
- Explaining financial concepts in simple terms
- Analyzing real QuickBooks financial data when provided
- Providing general tax guidance (not specific tax advice)
- Suggesting best practices for financial management
- Categorizing and understanding uploaded financial documents

CRITICAL — Response formatting rules (you MUST follow these):
- Your responses will be rendered as HTML in a chat interface
- For financial statements, reports, or any tabular data, ALWAYS use proper HTML tables:
  <table>
    <thead><tr><th>Account</th><th>Amount</th></tr></thead>
    <tbody><tr><td>Revenue</td><td>$50,000</td></tr></tbody>
    <tfoot><tr><td><strong>Total</strong></td><td><strong>$50,000</strong></td></tr></tfoot>
  </table>
- Use <strong> for bold, <em> for emphasis
- Use <br> for line breaks between paragraphs (NOT <p> tags — the chat bubble already wraps in <p>)
- For section headers within a response, use <strong> with a <br> before it
- Right-align all dollar amounts in tables using style="text-align:right"
- For income statements / P&L: group into Revenue, Cost of Goods Sold, Gross Profit, Operating Expenses, Net Income with subtotal rows
- For balance sheets: group into Assets, Liabilities, Equity with subtotal rows
- Use indentation for sub-accounts: &nbsp;&nbsp;&nbsp; before the account name
- Add a separator row between major sections using a row with a bottom border
- Keep commentary brief — lead with the formatted data, then add a short summary below
- For non-tabular responses, keep text concise and use <strong> for key figures

When QuickBooks data is provided:
- Analyze the actual numbers and give specific, data-driven answers
- Format currency values clearly (e.g., $12,345.67) and right-align them
- Compare periods when data is available
- Highlight important trends, concerns, or opportunities
- Be specific — reference actual account names, amounts, and dates from the data
- You have access to transaction-level detail (individual purchases, bills, invoices, payments). When a client asks about a specific expense category or line item, look through expenseTransactions and billTransactions to find and list the individual transactions
- Never say you don't have access to transaction data — you do. If the data is in the QuickBooks context, use it

Important rules:
- Never provide specific tax advice — always note that specific situations should be reviewed by a qualified tax professional
- If QuickBooks data is provided, use it to give real answers with real numbers
- If no QuickBooks data is available, let the client know they should connect QuickBooks for real-time insights
- Always be honest about what you can and can't do
- If you don't know something, say so
- When referencing the company, always use lowercase "jack does"`;

// Store conversation history per session (in-memory for now)
const conversations = new Map();

// ========================================
// QUICKBOOKS OAUTH ENDPOINTS
// ========================================

// Start QuickBooks connection — requires clientId to associate the QBO connection
app.get('/api/qbo/connect', (req, res) => {
  const clientId = req.query.clientId || 'default';
  const authUri = qbo.getAuthUri(clientId);
  // Store where to redirect after OAuth (admin or portal) and which client
  const from = req.query.from || 'portal';
  res.cookie('qbo_connect_from', from, { maxAge: 300000, httpOnly: true });
  res.cookie('qbo_connect_client', clientId, { maxAge: 300000, httpOnly: true });
  res.redirect(authUri);
});

// OAuth callback
app.get('/api/qbo/callback', async (req, res) => {
  try {
    // Extract clientId from cookie (set during connect) or from OAuth state param
    const clientCookie = req.headers.cookie?.split(';').find(c => c.trim().startsWith('qbo_connect_client='));
    const clientId = clientCookie?.split('=')[1]?.trim() || 'default';

    const tokenData = await qbo.handleCallback(req.url, clientId);
    console.log(`QuickBooks connected for client "${clientId}"! Realm ID:`, tokenData.realmId);

    // Redirect back to the originating page
    const fromCookie = req.headers.cookie?.split(';').find(c => c.trim().startsWith('qbo_connect_from='));
    const from = fromCookie?.split('=')[1]?.trim() || 'portal';

    if (from === 'admin') {
      res.redirect(`/admin/dashboard.html?qbo=connected&clientId=${clientId}`);
    } else {
      res.redirect('/portal/dashboard.html?qbo=connected');
    }
  } catch (error) {
    console.error('QuickBooks OAuth error:', error.message);
    const fromCookie = req.headers.cookie?.split(';').find(c => c.trim().startsWith('qbo_connect_from='));
    const from = fromCookie?.split('=')[1]?.trim() || 'portal';
    if (from === 'admin') {
      res.redirect('/admin/dashboard.html?qbo=error');
    } else {
      res.redirect('/portal/dashboard.html?qbo=error');
    }
  }
});

// Check connection status (per-client or global)
app.get('/api/qbo/status', (req, res) => {
  const clientId = req.query.clientId;
  if (clientId) {
    res.json({ connected: qbo.isConnected(clientId), clientId });
  } else {
    res.json({ connected: qbo.isAnyConnected(), connections: qbo.getAllConnections() });
  }
});

// Admin endpoint to get current QBO tokens (for saving to Railway env var)
app.get('/api/admin/qbo-tokens', requireAdmin, (req, res) => {
  if (!qbo.isAnyConnected()) {
    return res.json({ connected: false, tokens: null });
  }
  // Read tokens from memory (works even when loaded from env var with no disk file)
  const tokens = qbo.getTokensJson();
  res.json({ connected: true, tokens });
});

// Disconnect QuickBooks for a specific client
app.post('/api/qbo/disconnect', (req, res) => {
  const clientId = req.body.clientId || 'default';
  qbo.disconnect(clientId);
  res.json({ success: true, message: `QuickBooks disconnected for client "${clientId}"` });
});

// ========================================
// CHAT API ENDPOINT (with QuickBooks integration)
// ========================================
app.post('/api/chat', resolveClient, async (req, res) => {
  try {
    const { message, sessionId = 'default' } = req.body;

    if (!message) {
      return res.status(400).json({ error: 'Message is required' });
    }

    // Get or create conversation history (scoped by client)
    const convoKey = `${req.clientId}:${sessionId}`;
    if (!conversations.has(convoKey)) {
      conversations.set(convoKey, []);
    }
    const history = conversations.get(convoKey);

    // If QuickBooks is connected for this client, fetch relevant financial data
    let qboContext = '';
    if (qbo.isConnected(req.clientId)) {
      try {
        // Include recent USER messages (not assistant responses) for follow-up context
        const recentUserMessages = history
          .filter(m => m.role === 'user')
          .slice(-3)
          .map(m => m.content)
          .join(' ');
        const combinedQuery = `${message} ${recentUserMessages}`;
        console.log(`[QBO] Fetching data for client "${req.clientId}", query:`, message.substring(0, 100));
        const data = await qbo.fetchRelevantData(combinedQuery, req.clientId);
        const dataKeys = Object.keys(data).filter(k => k !== 'error');
        console.log('[QBO] Data fetched:', dataKeys.join(', '));
        if (data && dataKeys.length > 0 && !data.error) {
          qboContext = `\n\n--- QUICKBOOKS DATA (from the client's actual books) ---\n${JSON.stringify(data, null, 2)}\n--- END QUICKBOOKS DATA ---\n\nUse this real data to answer the client's question. Reference actual numbers and account names. When the client asks for detail on a specific line item, look through the transaction data (expenseTransactions, billTransactions) to find the individual transactions that make up that category.`;
        } else if (data.error) {
          console.error('[QBO] Data fetch returned error:', data.error);
          qboContext = `\n\n[Note: Tried to fetch QuickBooks data but got an error: ${data.error}. Answer the question generally and let the client know there was a temporary issue accessing their books.]`;
        }
      } catch (qboError) {
        console.error('[QBO] Data fetch exception:', qboError.message);
        qboContext = '\n\n[Note: QuickBooks is connected but there was an error fetching data. Answer generally.]';
      }
    }

    // Build the user message with QBO context if available
    const enrichedMessage = qboContext
      ? `${message}${qboContext}`
      : message;

    // Add user message to history (store original, not enriched)
    history.push({ role: 'user', content: enrichedMessage });

    // Keep conversation history manageable (last 20 messages)
    const recentHistory = history.slice(-20);

    // Build system prompt with connection status
    let systemPrompt = JACK_SYSTEM_PROMPT;
    if (!qbo.isConnected(req.clientId)) {
      systemPrompt += '\n\n[QuickBooks is NOT connected. When the client asks financial questions that need real data, let them know they can connect QuickBooks by clicking the "connect quickbooks" button in the portal for real-time insights from their actual books.]';
    }

    // Call Claude API
    const response = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4096,
      system: systemPrompt,
      messages: recentHistory,
    });

    const assistantMessage = response.content[0].text;

    // Convert markdown to HTML for rendering in chat
    const htmlMessage = marked(assistantMessage);

    // Add assistant response to history (store raw for Claude context)
    history.push({ role: 'assistant', content: assistantMessage });

    // Store original user message in history (replace enriched version)
    const userMsgIndex = history.length - 2;
    history[userMsgIndex] = { role: 'user', content: message };

    res.json({
      response: htmlMessage,
      qboConnected: qbo.isConnected(req.clientId),
    });

  } catch (error) {
    console.error('Chat API error:', error.message);

    if (error.status === 401) {
      return res.status(500).json({
        error: 'API key not configured. Please set ANTHROPIC_API_KEY in your .env file.'
      });
    }

    res.status(500).json({
      error: 'something went wrong. please try again.'
    });
  }
});

// ========================================
// FILE UPLOAD ENDPOINT
// ========================================
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const clientId = req.clientId || 'demo-client';
    const category = req.body.category || 'general';
    const clientDir = path.join(uploadsDir, clientId, category);
    if (!fs.existsSync(clientDir)) {
      fs.mkdirSync(clientDir, { recursive: true });
    }
    cb(null, clientDir);
  },
  filename: (req, file, cb) => {
    const timestamp = Date.now();
    const ext = path.extname(file.originalname);
    const name = path.basename(file.originalname, ext)
      .replace(/[^a-zA-Z0-9-_]/g, '_');
    cb(null, `${name}_${timestamp}${ext}`);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 25 * 1024 * 1024 }, // 25MB limit
  fileFilter: (req, file, cb) => {
    const allowed = ['.pdf', '.csv', '.xlsx', '.xls', '.jpg', '.jpeg', '.png', '.doc', '.docx'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error(`File type ${ext} not supported`));
    }
  }
});

app.post('/api/upload', resolveClient, upload.array('files', 10), (req, res) => {
  try {
    const files = req.files.map(f => ({
      name: f.originalname,
      size: f.size,
      category: req.body.category || 'general',
      path: f.path,
      storedName: f.filename,
      uploadedAt: new Date().toISOString(),
      clientId: req.clientId,
    }));
    res.json({ success: true, files });
  } catch (error) {
    console.error('Upload error:', error.message);
    res.status(500).json({ error: 'Upload failed' });
  }
});

// List uploaded files (scoped by client)
app.get('/api/files', resolveClient, (req, res) => {
  try {
    const clientDir = path.join(uploadsDir, req.clientId);
    const files = [];

    if (!fs.existsSync(clientDir)) {
      return res.json({ files: [] });
    }

    const categories = fs.readdirSync(clientDir).filter(f =>
      fs.statSync(path.join(clientDir, f)).isDirectory()
    );

    categories.forEach(category => {
      const categoryDir = path.join(clientDir, category);
      const categoryFiles = fs.readdirSync(categoryDir);
      categoryFiles.forEach(filename => {
        const filePath = path.join(categoryDir, filename);
        const stats = fs.statSync(filePath);
        files.push({
          name: filename,
          category,
          size: stats.size,
          uploadedAt: stats.mtime.toISOString(),
        });
      });
    });

    // Sort by most recent
    files.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));
    res.json({ files });
  } catch (error) {
    res.json({ files: [] });
  }
});

// ========================================
// DOCUMENT PROCESSING (Upload → Claude → QuickBooks)
// ========================================

const DOCUMENT_ANALYSIS_PROMPT = `You are "jack", an expert AI accountant. A client has uploaded a document for you to process.

Your job is to:
1. Analyze the document content
2. Determine what accounting entries are needed
3. Return a structured JSON response with the proposed entries

You MUST respond with ONLY valid JSON in this exact format (no markdown, no code fences, no explanation outside the JSON):
{
  "documentType": "invoice|receipt|bank_statement|expense_report|payroll|other",
  "summary": "Brief human-readable summary of what this document is",
  "vendor": "Vendor/supplier name if applicable",
  "customer": "Customer name if applicable",
  "date": "YYYY-MM-DD transaction date",
  "totalAmount": 123.45,
  "currency": "USD",
  "entries": [
    {
      "type": "journal_entry|bill|invoice",
      "date": "YYYY-MM-DD",
      "memo": "Description of this entry",
      "lines": [
        {
          "accountName": "Name of the QBO account (e.g., Office Supplies, Accounts Payable, Revenue)",
          "accountCategory": "Expense|Revenue|Asset|Liability|Equity",
          "description": "Line item description",
          "amount": 123.45,
          "type": "debit|credit"
        }
      ]
    }
  ],
  "notes": "Any additional notes or things the client should be aware of",
  "confidence": "high|medium|low",
  "needsReview": ["List of items that need human review or clarification"]
}

Rules for determining entries:
- For INVOICES received (bills): Debit the appropriate expense account, Credit Accounts Payable
- For INVOICES issued: Debit Accounts Receivable, Credit the appropriate revenue account
- For RECEIPTS: Debit the appropriate expense account, Credit Cash/Bank account
- For BANK STATEMENTS: Create entries for each transaction — debits to expense accounts, credits to bank for outflows; debits to bank, credits to revenue for inflows
- For PAYROLL docs: Debit Salary/Wage Expense, Credit Payroll Liabilities
- Always ensure debits = credits for each journal entry
- Use standard account names that would exist in a typical QuickBooks chart of accounts
- If you can't determine something with confidence, add it to needsReview`;

/**
 * Process an uploaded document through Claude to determine accounting entries
 */
app.post('/api/process-document', resolveClient, async (req, res) => {
  try {
    const { filePath, fileName, category } = req.body;

    if (!filePath || !fileName) {
      return res.status(400).json({ error: 'File path and name are required' });
    }

    // Read the file
    const fullPath = path.resolve(filePath);
    if (!fs.existsSync(fullPath)) {
      return res.status(404).json({ error: 'File not found' });
    }

    const ext = path.extname(fileName).toLowerCase();
    let messageContent = [];

    if (['.jpg', '.jpeg', '.png'].includes(ext)) {
      // Image files — send as base64 image to Claude
      const fileBuffer = fs.readFileSync(fullPath);
      const base64 = fileBuffer.toString('base64');
      const mediaType = ext === '.png' ? 'image/png' : 'image/jpeg';

      messageContent = [
        {
          type: 'image',
          source: { type: 'base64', media_type: mediaType, data: base64 },
        },
        {
          type: 'text',
          text: `This is an uploaded ${category || 'accounting'} document named "${fileName}". Please analyze it and determine the accounting entries needed. Respond with JSON only.`,
        },
      ];
    } else if (['.csv'].includes(ext)) {
      // CSV — read as text
      const content = fs.readFileSync(fullPath, 'utf-8');
      messageContent = [
        {
          type: 'text',
          text: `This is a CSV file named "${fileName}" in the "${category || 'general'}" category. Here is the content:\n\n${content}\n\nPlease analyze it and determine the accounting entries needed. Respond with JSON only.`,
        },
      ];
    } else if (['.pdf'].includes(ext)) {
      // PDF — send as base64 document to Claude
      const fileBuffer = fs.readFileSync(fullPath);
      const base64 = fileBuffer.toString('base64');

      messageContent = [
        {
          type: 'document',
          source: { type: 'base64', media_type: 'application/pdf', data: base64 },
        },
        {
          type: 'text',
          text: `This is an uploaded PDF document named "${fileName}" in the "${category || 'general'}" category. Please analyze it and determine the accounting entries needed. Respond with JSON only.`,
        },
      ];
    } else {
      // Other file types — try reading as text
      let content;
      try {
        content = fs.readFileSync(fullPath, 'utf-8');
      } catch {
        return res.status(400).json({ error: `Cannot read file type ${ext}. Supported: PDF, CSV, JPG, PNG` });
      }

      messageContent = [
        {
          type: 'text',
          text: `This is a document named "${fileName}" in the "${category || 'general'}" category. Here is the content:\n\n${content}\n\nPlease analyze it and determine the accounting entries needed. Respond with JSON only.`,
        },
      ];
    }

    // If QuickBooks is connected for this client, fetch account list so Claude can map to real accounts
    let accountContext = '';
    if (qbo.isConnected(req.clientId)) {
      try {
        const accountsData = await qbo.getAccounts(req.clientId);
        const accountsList = accountsData?.QueryResponse?.Account || accountsData || [];
        const accounts = Array.isArray(accountsList) ? accountsList : [];
        if (accounts.length > 0) {
          const acctList = accounts
            .filter(a => a.Active !== false)
            .map(a => `- "${a.Name}" (ID: ${a.Id}, Type: ${a.AccountType})`)
            .join('\n');
          accountContext = `\n\nIMPORTANT — CHART OF ACCOUNTS:\nBelow is the client's actual QuickBooks chart of accounts. You MUST use these exact account names and include the "accountId" field for each journal entry line. Pick the most appropriate account for each line item based on the document content.\n\n${acctList}\n\nFor every line in your entries, set "accountName" to the exact account name from this list and "accountId" to the corresponding ID.`;
        }
      } catch (e) {
        console.error('Could not fetch QBO accounts for context:', e.message);
      }
    }

    // Call Claude to analyze the document
    const response = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4096,
      system: DOCUMENT_ANALYSIS_PROMPT + accountContext,
      messages: [{ role: 'user', content: messageContent }],
    });

    const rawText = response.content[0].text;

    // Parse the JSON response
    let analysis;
    try {
      // Try to extract JSON from the response (in case Claude wraps it)
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      analysis = JSON.parse(jsonMatch ? jsonMatch[0] : rawText);
    } catch (parseErr) {
      console.error('Failed to parse Claude response:', rawText);
      return res.status(500).json({
        error: 'Could not parse document analysis',
        rawResponse: rawText,
      });
    }

    // Store the analysis for later approval
    const analysisId = 'analysis_' + Date.now() + '_' + Math.random().toString(36).substring(2, 8);
    documentAnalyses.set(analysisId, {
      analysis,
      filePath: fullPath,
      fileName,
      category,
      clientId: req.clientId,
      clientName: req.clientName,
      createdAt: new Date().toISOString(),
      status: 'pending', // pending, approved, rejected
    });

    res.json({
      success: true,
      analysisId,
      analysis,
    });

  } catch (error) {
    console.error('Document processing error:', error.message);
    res.status(500).json({ error: 'Failed to process document: ' + error.message });
  }
});

// In-memory store for document analyses (replace with DB in production)
const documentAnalyses = new Map();

/**
 * Approve proposed entries and push to QuickBooks
 */
app.post('/api/approve-entries', async (req, res) => {
  try {
    const { analysisId, entries } = req.body;

    if (!analysisId) {
      return res.status(400).json({ error: 'Analysis ID is required' });
    }

    const stored = documentAnalyses.get(analysisId);
    if (!stored) {
      return res.status(404).json({ error: 'Analysis not found' });
    }

    const postClientId = stored.clientId || 'default';
    if (!qbo.isConnected(postClientId)) {
      return res.status(400).json({ error: `QuickBooks is not connected for client "${postClientId}". Please connect QuickBooks first.` });
    }

    // Use provided entries (may have been edited by user) or fall back to original
    const entriesToPost = entries || stored.analysis.entries;
    const results = [];

    for (const entry of entriesToPost) {
      try {
        let result;

        if (entry.type === 'journal_entry') {
          result = await qbo.createJournalEntry({
            date: entry.date,
            memo: entry.memo,
            lines: entry.lines,
          }, postClientId);
        } else if (entry.type === 'bill') {
          result = await qbo.createBill({
            date: entry.date,
            memo: entry.memo,
            vendorId: entry.vendorId,
            dueDate: entry.dueDate,
            lines: entry.lines,
          }, postClientId);
        } else if (entry.type === 'invoice') {
          result = await qbo.createInvoice({
            date: entry.date,
            memo: entry.memo,
            customerId: entry.customerId,
            dueDate: entry.dueDate,
            lines: entry.lines,
          }, postClientId);
        } else {
          // Default to journal entry
          result = await qbo.createJournalEntry({
            date: entry.date,
            memo: entry.memo,
            lines: entry.lines,
          }, postClientId);
        }

        results.push({ success: true, entry: entry.memo, result });
      } catch (entryError) {
        results.push({ success: false, entry: entry.memo, error: entryError.message });
      }
    }

    // Update status
    stored.status = results.every(r => r.success) ? 'approved' : 'partial';
    documentAnalyses.set(analysisId, stored);

    res.json({
      success: true,
      results,
      message: `${results.filter(r => r.success).length} of ${results.length} entries posted to QuickBooks`,
    });

  } catch (error) {
    console.error('Approve entries error:', error.message);
    res.status(500).json({ error: 'Failed to post entries: ' + error.message });
  }
});

/**
 * Reject / discard proposed entries
 */
app.post('/api/reject-entries', (req, res) => {
  const { analysisId } = req.body;

  if (analysisId && documentAnalyses.has(analysisId)) {
    const stored = documentAnalyses.get(analysisId);
    stored.status = 'rejected';
    documentAnalyses.set(analysisId, stored);
  }

  res.json({ success: true });
});

/**
 * Admin: Get all document analyses (pending first)
 */
app.get('/api/admin/analyses', requireAdmin, (req, res) => {
  const analyses = [];
  documentAnalyses.forEach((value, key) => {
    analyses.push({ id: key, ...value });
  });
  // Pending first, then by date
  analyses.sort((a, b) => {
    if (a.status === 'pending' && b.status !== 'pending') return -1;
    if (a.status !== 'pending' && b.status === 'pending') return 1;
    return new Date(b.createdAt) - new Date(a.createdAt);
  });
  res.json({ analyses });
});

/**
 * Admin: Approve entries and post to QuickBooks
 */
app.post('/api/admin/approve', requireAdmin, async (req, res) => {
  try {
    const { analysisId, entries } = req.body;

    const stored = documentAnalyses.get(analysisId);
    if (!stored) {
      return res.status(404).json({ error: 'Analysis not found' });
    }

    const approveClientId = stored.clientId || 'default';
    if (!qbo.isConnected(approveClientId)) {
      return res.status(400).json({ error: `QuickBooks is not connected for client "${approveClientId}".` });
    }

    const entriesToPost = entries || stored.analysis.entries;
    const results = [];

    for (const entry of entriesToPost) {
      try {
        // Validate that all lines have accountId
        const missingAccount = entry.lines.find(l => !l.accountId);
        if (missingAccount) {
          console.error('Line missing accountId:', JSON.stringify(missingAccount));
          results.push({ success: false, entry: entry.memo, error: `Account "${missingAccount.accountName || 'unknown'}" has no QBO account ID. Please select an account from the dropdown.` });
          continue;
        }

        let result;
        console.log('Posting entry to QBO:', JSON.stringify(entry, null, 2));

        if (entry.type === 'bill') {
          result = await qbo.createBill({ date: entry.date, memo: entry.memo, vendorId: entry.vendorId, dueDate: entry.dueDate, lines: entry.lines }, approveClientId);
        } else if (entry.type === 'invoice') {
          result = await qbo.createInvoice({ date: entry.date, memo: entry.memo, customerId: entry.customerId, dueDate: entry.dueDate, lines: entry.lines }, approveClientId);
        } else {
          result = await qbo.createJournalEntry({ date: entry.date, memo: entry.memo, lines: entry.lines }, approveClientId);
        }
        results.push({ success: true, entry: entry.memo, result });
      } catch (entryError) {
        console.error('QBO posting error:', entryError);
        const errMsg = entryError?.Fault?.Error?.[0]?.Detail || entryError?.Fault?.Error?.[0]?.Message || entryError.message || JSON.stringify(entryError);
        results.push({ success: false, entry: entry.memo, error: errMsg });
      }
    }

    stored.status = results.every(r => r.success) ? 'approved' : 'partial';
    stored.reviewedAt = new Date().toISOString();
    documentAnalyses.set(analysisId, stored);

    res.json({
      success: true,
      results,
      message: `${results.filter(r => r.success).length} of ${results.length} entries posted`,
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/**
 * Admin: Reject entries
 */
app.post('/api/admin/reject', requireAdmin, (req, res) => {
  const { analysisId, reason } = req.body;
  const stored = documentAnalyses.get(analysisId);
  if (stored) {
    stored.status = 'rejected';
    stored.rejectReason = reason || '';
    stored.reviewedAt = new Date().toISOString();
    stored.clientRead = false; // Mark as unread notification for client
    documentAnalyses.set(analysisId, stored);
  }
  res.json({ success: true });
});

/**
 * Client: Get notifications (rejected documents, scoped by client)
 */
app.get('/api/notifications', resolveClient, (req, res) => {
  const notifications = [];
  documentAnalyses.forEach((stored, id) => {
    if (stored.status === 'rejected' && stored.clientId === req.clientId) {
      notifications.push({
        id,
        fileName: stored.fileName,
        category: stored.category || 'general',
        rejectReason: stored.rejectReason || '',
        reviewedAt: stored.reviewedAt,
        read: stored.clientRead === true,
        summary: stored.analysis?.summary || '',
      });
    }
  });

  // Sort newest first
  notifications.sort((a, b) => new Date(b.reviewedAt) - new Date(a.reviewedAt));

  const unreadCount = notifications.filter(n => !n.read).length;
  res.json({ notifications, unreadCount });
});

/**
 * Client: Mark notifications as read
 */
app.post('/api/notifications/read', resolveClient, (req, res) => {
  const { ids } = req.body;
  if (Array.isArray(ids)) {
    ids.forEach(id => {
      const stored = documentAnalyses.get(id);
      if (stored) {
        stored.clientRead = true;
        documentAnalyses.set(id, stored);
      }
    });
  }
  res.json({ success: true });
});

/**
 * Admin: Update/edit entries before approving
 */
app.put('/api/admin/analysis/:id', requireAdmin, (req, res) => {
  const stored = documentAnalyses.get(req.params.id);
  if (!stored) {
    return res.status(404).json({ error: 'Analysis not found' });
  }
  // Allow editing the analysis entries
  if (req.body.entries) {
    stored.analysis.entries = req.body.entries;
  }
  documentAnalyses.set(req.params.id, stored);
  res.json({ success: true, analysis: stored.analysis });
});

/**
 * Admin: Get QuickBooks chart of accounts
 */
app.get('/api/admin/accounts', requireAdmin, async (req, res) => {
  const acctClientId = req.query.clientId || 'default';
  if (!qbo.isConnected(acctClientId)) {
    return res.json({ accounts: [], clientId: acctClientId });
  }

  try {
    const data = await qbo.getAccounts(acctClientId);
    const rawAccounts = data?.QueryResponse?.Account || data || [];
    const accountsArray = Array.isArray(rawAccounts) ? rawAccounts : [];
    if (accountsArray.length > 0) {
      const accounts = accountsArray.map(a => ({
        id: a.Id,
        name: a.Name,
        acctNum: a.AcctNum || '',
        type: a.AccountType,
        subType: a.AccountSubType || '',
        fullyQualifiedName: a.FullyQualifiedName || a.Name,
        active: a.Active,
      })).filter(a => a.active !== false)
        .sort((a, b) => a.type.localeCompare(b.type) || a.name.localeCompare(b.name));

      return res.json({ accounts });
    }
    res.json({ accounts: [] });
  } catch (error) {
    console.error('Failed to fetch accounts:', error.message);
    res.json({ accounts: [], error: error.message });
  }
});

/**
 * Admin: Get dashboard stats
 */
app.get('/api/admin/stats', requireAdmin, (req, res) => {
  let pending = 0, approved = 0, rejected = 0;
  documentAnalyses.forEach(v => {
    if (v.status === 'pending') pending++;
    else if (v.status === 'approved') approved++;
    else if (v.status === 'rejected') rejected++;
  });
  res.json({ pending, approved, rejected, total: documentAnalyses.size, qboConnected: qbo.isAnyConnected(), qboConnections: qbo.getAllConnections(), clientCount: Object.keys(CLIENTS).length });
});


// Admin: Check QBO connection for a specific client
app.get('/api/admin/clients/:clientId/qbo-status', requireAdmin, (req, res) => {
  const clientId = req.params.clientId;
  const connected = qbo.isConnected(clientId);
  const connections = qbo.getAllConnections();
  const clientConnection = connections[clientId] || null;
  res.json({ connected, clientId, connection: clientConnection });
});

// ========================================
// CLIENT CONFIG (portal fetches this)
// ========================================
app.get('/api/client/config', resolveClient, (req, res) => {
  const client = CLIENTS[req.clientId];
  if (!client) {
    return res.json({ billingFrequency: 'monthly' });
  }
  res.json({
    billingFrequency: client.billingFrequency || 'monthly',
    name: client.name,
    email: client.email,
  });
});

// ========================================
// CLIENT MANAGEMENT (admin)
// ========================================
app.get('/api/admin/clients', requireAdmin, (req, res) => {
  const clientList = Object.entries(CLIENTS).map(([id, data]) => {
    const { password, ...rest } = data;
    return { id, ...rest };
  });
  clientList.sort((a, b) => a.name.localeCompare(b.name));
  res.json({ clients: clientList });
});

app.post('/api/admin/clients', requireAdmin, (req, res) => {
  const { name, email, password, billingFrequency, fiscalYearEnd } = req.body;
  if (!name || !email) {
    return res.status(400).json({ error: 'Name and email are required' });
  }
  if (!password) {
    return res.status(400).json({ error: 'Password is required' });
  }
  if (fiscalYearEnd && !/^\d{2}-\d{2}$/.test(fiscalYearEnd)) {
    return res.status(400).json({ error: 'fiscalYearEnd must be MM-DD' });
  }
  const id = name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
  if (CLIENTS[id]) {
    return res.status(409).json({ error: 'A client with that ID already exists' });
  }
  const validFreqs = ['monthly', 'quarterly', 'annual'];
  CLIENTS[id] = {
    name,
    email,
    password,
    billingFrequency: validFreqs.includes(billingFrequency) ? billingFrequency : 'monthly',
    fiscalYearEnd: fiscalYearEnd || null,
    createdAt: new Date().toISOString(),
  };
  saveClients(CLIENTS);
  res.json({ success: true, id, client: { ...CLIENTS[id], password: undefined } });
});

app.put('/api/admin/clients/:id', requireAdmin, (req, res) => {
  const client = CLIENTS[req.params.id];
  if (!client) {
    return res.status(404).json({ error: 'Client not found' });
  }
  const { name, email, password, billingFrequency, fiscalYearEnd } = req.body;
  if (name) client.name = name;
  if (email) client.email = email;
  if (password) client.password = password;
  const validFreqs = ['monthly', 'quarterly', 'annual'];
  if (billingFrequency && validFreqs.includes(billingFrequency)) {
    client.billingFrequency = billingFrequency;
  }
  if (fiscalYearEnd !== undefined) {
    if (fiscalYearEnd && !/^\d{2}-\d{2}$/.test(fiscalYearEnd)) {
      return res.status(400).json({ error: 'fiscalYearEnd must be MM-DD' });
    }
    client.fiscalYearEnd = fiscalYearEnd || null;
  }
  CLIENTS[req.params.id] = client;
  saveClients(CLIENTS);
  res.json({ success: true, client: { ...client, password: undefined } });
});

// ========================================
// FIXED ASSETS — per-client, file-backed
// ========================================
const FIXED_ASSETS_FILE = dataPath('fixed-assets.json');

function loadFixedAssets() {
  try {
    if (fs.existsSync(FIXED_ASSETS_FILE)) {
      return JSON.parse(fs.readFileSync(FIXED_ASSETS_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Failed to load fixed-assets.json:', e.message);
  }
  return {};
  // Structure: { "client-id": { assets: [...], amortizationRuns: [...] } }
}

function saveFixedAssets(data) {
  fs.writeFileSync(FIXED_ASSETS_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

function getClientAssets(clientId) {
  if (!fixedAssetsData[clientId]) {
    fixedAssetsData[clientId] = { assetClasses: [], assets: [], amortizationRuns: [] };
  }
  // Migration: ensure assetClasses exists for older data
  if (!fixedAssetsData[clientId].assetClasses) {
    fixedAssetsData[clientId].assetClasses = [];
  }
  return fixedAssetsData[clientId];
}

let fixedAssetsData = loadFixedAssets();

// Get fixed assets for a specific client
app.get('/api/admin/clients/:clientId/fixed-assets', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  res.json(clientData);
});

// ---- Asset Class CRUD ----

// Create asset class
app.post('/api/admin/clients/:clientId/fixed-assets/classes', requireAdmin, (req, res) => {
  const { glAccountId, glAccountName, method, usefulLifeMonths, decliningRate,
          salvageValue, expenseAccountId, expenseAccountName,
          accumAccountId, accumAccountName, aiSuggestion } = req.body;

  if (!glAccountName) return res.status(400).json({ error: 'GL account name is required' });

  const clientData = getClientAssets(req.params.clientId);

  // Prevent duplicate classes for same GL account
  const exists = clientData.assetClasses.find(c =>
    (glAccountId && c.glAccountId === glAccountId) || c.glAccountName === glAccountName
  );
  if (exists) return res.status(409).json({ error: 'A class for this GL account already exists', existingClass: exists });

  const assetClass = {
    id: 'class_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
    glAccountId: glAccountId || '',
    glAccountName,
    method: method || 'straight-line',
    usefulLifeMonths: parseInt(usefulLifeMonths, 10) || 60,
    decliningRate: decliningRate ? parseFloat(decliningRate) : null,
    salvageValue: parseFloat(salvageValue || 0),
    expenseAccountId: expenseAccountId || '',
    expenseAccountName: expenseAccountName || '',
    accumAccountId: accumAccountId || '',
    accumAccountName: accumAccountName || '',
    aiSuggestion: aiSuggestion || null,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  clientData.assetClasses.push(assetClass);
  saveFixedAssets(fixedAssetsData);
  res.json({ success: true, assetClass });
});

// Update asset class
app.put('/api/admin/clients/:clientId/fixed-assets/classes/:classId', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  const idx = clientData.assetClasses.findIndex(c => c.id === req.params.classId);
  if (idx === -1) return res.status(404).json({ error: 'Asset class not found' });

  const fields = ['glAccountName', 'method', 'usefulLifeMonths', 'decliningRate',
    'salvageValue', 'expenseAccountId', 'expenseAccountName',
    'accumAccountId', 'accumAccountName', 'aiSuggestion'];

  fields.forEach(f => {
    if (req.body[f] !== undefined) {
      if (f === 'usefulLifeMonths') {
        clientData.assetClasses[idx][f] = parseInt(req.body[f], 10);
      } else if (f === 'salvageValue' || f === 'decliningRate') {
        clientData.assetClasses[idx][f] = req.body[f] !== null ? parseFloat(req.body[f]) : null;
      } else {
        clientData.assetClasses[idx][f] = req.body[f];
      }
    }
  });
  clientData.assetClasses[idx].updatedAt = new Date().toISOString();

  saveFixedAssets(fixedAssetsData);
  res.json({ success: true, assetClass: clientData.assetClasses[idx] });
});

// Delete asset class
app.delete('/api/admin/clients/:clientId/fixed-assets/classes/:classId', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  const idx = clientData.assetClasses.findIndex(c => c.id === req.params.classId);
  if (idx === -1) return res.status(404).json({ error: 'Asset class not found' });
  clientData.assetClasses.splice(idx, 1);
  saveFixedAssets(fixedAssetsData);
  res.json({ success: true });
});

// ========================================
// EXCEL EXPORT / IMPORT
// ========================================
const excelUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// Export fixed assets to Excel workbook
app.get('/api/admin/clients/:clientId/fixed-assets/export-excel', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    const client = CLIENTS[clientId];
    const clientName = client?.name || clientId;
    const clientData = getClientAssets(clientId);
    // Amortization runs are kept in sync with QBO via the "Sync from QBO" flow, so the
    // saved JSON is the source of truth here. No live QBO call on export.

    // Build (or recompute) the continuity schedule for the most recent run month
    const runs = clientData.amortizationRuns || [];
    const sortedRuns = runs.slice().sort((a, b) => (b.month || '').localeCompare(a.month || ''));
    const asOfMonth = sortedRuns[0]?.month || formatMonthStr(new Date().getFullYear(), new Date().getMonth());
    const fye = CLIENTS[clientId]?.fiscalYearEnd || null;
    const continuitySchedule = buildContinuitySchedule(clientData, asOfMonth, fye);

    const workbook = await excelService.generateWorkbook(clientId, clientName, clientData, continuitySchedule);

    // Run Claude review on the generated workbook (advisory — never blocks the download).
    // Cache the result on clientData so the UI can fetch it via /last-review after the
    // file download starts.
    let review = null;
    try {
      review = await excelReviewService.reviewWorkbook(clientId, clientData, continuitySchedule, getAssetPolicy);
      excelService.addReviewSheet(workbook, review);
      clientData.lastReview = review;
      saveFixedAssets(fixedAssetsData);
    } catch (e) {
      console.error('[export] review failed:', e.message);
      review = {
        status: 'error',
        summary: `Review could not be completed: ${e.message}`,
        findings: [],
        generatedAt: new Date().toISOString(),
        asOfMonth,
        reviewError: e.message,
      };
      clientData.lastReview = review;
      saveFixedAssets(fixedAssetsData);
    }

    const fileName = `${clientName} - Fixed Assets.xlsx`.replace(/[<>:"/\\|?*]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('X-Review-Status', review?.status || 'unknown');
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Excel export error:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
});

// Get the most recent Claude review result for a client (populated by /export-excel)
app.get('/api/admin/clients/:clientId/fixed-assets/last-review', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  res.json({ review: clientData.lastReview || null });
});

// Import fixed assets from Excel workbook
app.post('/api/admin/clients/:clientId/fixed-assets/import-excel', requireAdmin, excelUpload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const workbook = new (require('exceljs').Workbook)();
    await workbook.xlsx.load(req.file.buffer);

    const parsed = excelService.parseWorkbook(workbook);

    // Set clientId on all assets
    const clientId = req.params.clientId;
    parsed.assets.forEach(a => { a.clientId = clientId; });

    // Replace in-memory data for this client
    fixedAssetsData[clientId] = {
      assetClasses: parsed.assetClasses,
      assets: parsed.assets,
      amortizationRuns: parsed.amortizationRuns,
    };
    saveFixedAssets(fixedAssetsData);

    res.json({
      success: true,
      assetCount: parsed.assets.length,
      classCount: parsed.assetClasses.length,
      runCount: parsed.amortizationRuns.length,
      warnings: parsed.warnings,
    });
  } catch (error) {
    console.error('Excel import error:', error);
    res.status(500).json({ error: 'Failed to import Excel file: ' + error.message });
  }
});

// Sync fixed assets from QBO — queries GL transactions for each Fixed Asset cost account
app.post('/api/admin/clients/:clientId/fixed-assets/sync-qbo', requireAdmin, async (req, res) => {
  try {
    const syncClientId = req.params.clientId;
    if (!qbo.isConnected(syncClientId)) {
      return res.status(400).json({ error: `QuickBooks is not connected for this client. Please connect QBO in the client detail view first.` });
    }

    // Get all accounts
    const accountsResult = await qbo.getAccounts(syncClientId);
    const accounts = accountsResult.QueryResponse?.Account || accountsResult || [];
    const allAccounts = Array.isArray(accounts) ? accounts : [];
    const fixedAssetAccounts = allAccounts.filter(a => a.AccountType === 'Fixed Asset' && a.Active);

    // Separate cost accounts from accumulated amortization/depreciation accounts.
    // QBO doesn't always populate AccountSubType, so also recognize accum accounts by
    // a name match against "accumulated amortization|depreciation".
    const accumSubTypes = ['AccumulatedDepreciation', 'AccumulatedAmortization'];
    const accumNameRe = /accumulated\s*(amortization|depreciation|amort|depr)/i;
    const isAccum = (a) => (a.AccountSubType && accumSubTypes.includes(a.AccountSubType))
      || accumNameRe.test(a.Name || '')
      || accumNameRe.test(a.FullyQualifiedName || '');
    const costAccounts = fixedAssetAccounts.filter(a => !isAccum(a));
    const accumAccounts = fixedAssetAccounts.filter(isAccum);
    console.log('[sync] cost accounts:', costAccounts.map(a => `${a.Id} ${a.Name}`));
    console.log('[sync] accum accounts:', accumAccounts.map(a => `${a.Id} ${a.Name}`));

    // Get expense accounts for auto-suggesting amortization expense account
    const expenseAccounts = allAccounts
      .filter(a => ['Expense', 'Other Expense'].includes(a.AccountType) && a.Active);

    // Check which assets are already imported
    const clientData = getClientAssets(req.params.clientId);
    const existingTxnKeys = new Set(clientData.assets.filter(a => a.active).map(a => a.txnKey).filter(Boolean));

    // Read fixed-asset acquisitions directly from JE / Bill / Purchase records.
    // Each entry corresponds to a single line that debits a cost account, with the
    // line description as the asset name and the line amount as the cost.
    const costAccountIdSet = new Set(costAccounts.map(a => String(a.Id)));
    const costAccountById = new Map(costAccounts.map(a => [String(a.Id), a]));
    let acquisitions = [];
    try {
      acquisitions = await qbo.getFixedAssetAcquisitions(syncClientId, costAccountIdSet);
    } catch (e) {
      console.error('Could not fetch fixed-asset acquisitions from QBO:', e.message);
    }
    console.log('[sync] acquisitions:', acquisitions.length);

    const importable = [];
    for (const acq of acquisitions) {
      if (!(acq.amount > 0)) continue;
      const costAcct = costAccountById.get(String(acq.accountId));
      if (!costAcct) continue;

      // Build a stable txnKey so re-syncing doesn't duplicate
      const txnKey = `${acq.txnType}-${acq.txnId}-${acq.accountId}-${acq.amount}`;
      if (existingTxnKeys.has(txnKey)) continue;

      // Auto-match accumulated amortization account by name proximity
      const nameBase = (costAcct.Name || '').toLowerCase().replace(/[-–—]/g, ' ').trim();
      const matchedAccum = accumAccounts.find(acc => {
        const accumName = (acc.Name || '').toLowerCase().replace(/[-–—]/g, ' ').trim();
        return accumName.includes(nameBase)
          || nameBase.includes(accumName.replace(/accumulated\s*(amortization|depreciation)\s*/i, '').trim());
      });
      const matchedExpense = expenseAccounts.find(exp => {
        const expName = (exp.Name || '').toLowerCase();
        return expName.includes('amortization') || expName.includes('depreciation');
      });

      // Asset name = the line description on the source transaction (what the user typed
      // into the JE/Bill/Expense). Fall back gracefully only if the line had no description.
      const assetName = acq.description || acq.vendorName || costAcct.Name;
      const assetDescription = acq.description || '';

      importable.push({
        qboAccountId: String(costAcct.Id),
        glAccountName: costAcct.Name,
        name: assetName,
        description: assetDescription,
        accountType: 'Fixed Asset',
        originalCost: acq.amount,
        txnDate: acq.txnDate,
        txnType: acq.txnType,
        docNum: acq.docNum,
        vendorName: acq.vendorName || '',
        memo: acq.description || '',
        txnKey,
        suggestedAccumAccountId: matchedAccum?.Id || '',
        suggestedAccumAccountName: matchedAccum?.Name || '',
        suggestedExpenseAccountId: matchedExpense?.Id || '',
        suggestedExpenseAccountName: matchedExpense?.Name || '',
      });
    }

    // Rehydrate posted amortization runs from QBO so the saved JSON reflects any
    // edits/deletions/backdated adjustments made directly in QBO. This is a full
    // replace, keyed by month, so the UI and Excel export share the same cached data.
    try {
      await hydrateAmortizationRunsFromQBO(syncClientId, clientData);
    } catch (e) {
      console.error('[sync] amortization rehydrate failed:', e.message);
    }

    // Record when this client last completed a full QBO sync
    clientData.lastSyncedAt = new Date().toISOString();
    saveFixedAssets(fixedAssetsData);

    res.json({ accounts: importable, lastSyncedAt: clientData.lastSyncedAt });
  } catch (error) {
    console.error('QBO sync error:', error.message);
    res.status(500).json({ error: `Failed to fetch from QBO: ${error.message}` });
  }
});

// AI suggest amortization policy for an asset class
app.post('/api/admin/fixed-assets/suggest-amortization', requireAdmin, async (req, res) => {
  try {
    const { assetName, assetType, originalCost } = req.body;

    const response = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 400,
      system: `You are a Canadian accounting expert. Given a fixed asset class/category, suggest the appropriate amortization method and parameters based on CRA Capital Cost Allowance (CCA) classes and standard accounting practices.

Respond in JSON only with this exact format:
{"method": "straight-line" or "declining-balance", "usefulLifeMonths": number or null, "decliningRate": number (annual rate as decimal e.g. 0.30 for 30%) or null, "salvageValue": number, "ccaClass": number or null, "ccaRate": "percentage string or null", "reasoning": "brief explanation"}

Rules:
- CCA uses declining balance method. If the asset clearly falls under a CCA class, recommend declining-balance with the CCA rate.
- If no CCA class applies or the asset is better suited for accounting purposes with straight-line, recommend straight-line with a useful life in months.
- For declining-balance, set usefulLifeMonths to null. For straight-line, set decliningRate to null.`,
      messages: [{ role: 'user', content: `Asset class: "${assetName}". Type: "${assetType || 'unknown'}". Typical cost: $${originalCost || 'unknown'}. What amortization method and parameters would you recommend?` }],
    });

    const text = response.content[0].text;
    // Parse JSON from response (handle markdown code blocks)
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const suggestion = JSON.parse(jsonMatch[0]);
      res.json(suggestion);
    } else {
      res.json({ error: 'Could not parse suggestion', raw: text });
    }
  } catch (error) {
    console.error('AI suggestion error:', error.message);
    res.status(500).json({ error: 'Failed to get AI suggestion' });
  }
});

// Create a new fixed asset for a client
app.post('/api/admin/clients/:clientId/fixed-assets', requireAdmin, (req, res) => {
  const { name, description, glAccountName, originalCost, usefulLifeMonths, salvageValue, acquisitionDate,
          assetAccountId, assetAccountName, expenseAccountId, expenseAccountName,
          accumAccountId, accumAccountName, qboAccountId, txnKey, vendorName, fromSync } = req.body;

  if (!name) {
    return res.status(400).json({ error: 'Name is required' });
  }
  // Strict validation only for manual adds (not syncs from QBO)
  if (!fromSync) {
    if (!originalCost || !usefulLifeMonths || !acquisitionDate) {
      return res.status(400).json({ error: 'Name, cost, useful life, and acquisition date are required' });
    }
    if (!assetAccountId || !expenseAccountId || !accumAccountId) {
      return res.status(400).json({ error: 'All three QBO accounts must be selected' });
    }
  }

  const clientData = getClientAssets(req.params.clientId);

  const asset = {
    id: 'asset_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
    clientId: req.params.clientId,
    name,
    description: description || '',
    glAccountName: glAccountName || '',
    vendorName: vendorName || '',
    originalCost: parseFloat(originalCost) || 0,
    usefulLifeMonths: parseInt(usefulLifeMonths, 10) || 60,
    salvageValue: parseFloat(salvageValue || 0),
    acquisitionDate: acquisitionDate || new Date().toISOString().split('T')[0],
    assetAccountId, assetAccountName,
    expenseAccountId, expenseAccountName,
    accumAccountId, accumAccountName,
    qboAccountId: qboAccountId || null,
    txnKey: txnKey || null,
    active: true,
    createdAt: new Date().toISOString(),
  };

  clientData.assets.push(asset);
  saveFixedAssets(fixedAssetsData);
  res.json({ success: true, asset });
});

// Update a fixed asset
app.put('/api/admin/clients/:clientId/fixed-assets/:id', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  const idx = clientData.assets.findIndex(a => a.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Asset not found' });

  const fields = ['name', 'description', 'originalCost', 'usefulLifeMonths', 'salvageValue', 'acquisitionDate',
    'assetAccountId', 'assetAccountName', 'expenseAccountId', 'expenseAccountName',
    'accumAccountId', 'accumAccountName', 'active', 'aiSuggestion'];

  fields.forEach(f => {
    if (req.body[f] !== undefined) {
      if (f === 'originalCost' || f === 'salvageValue') {
        clientData.assets[idx][f] = parseFloat(req.body[f]);
      } else if (f === 'usefulLifeMonths') {
        clientData.assets[idx][f] = parseInt(req.body[f], 10);
      } else {
        clientData.assets[idx][f] = req.body[f];
      }
    }
  });

  saveFixedAssets(fixedAssetsData);
  res.json({ success: true, asset: clientData.assets[idx] });
});

// Delete a fixed asset (soft delete)
app.delete('/api/admin/clients/:clientId/fixed-assets/:id', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  const idx = clientData.assets.findIndex(a => a.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Asset not found' });

  clientData.assets[idx].active = false;
  saveFixedAssets(fixedAssetsData);
  res.json({ success: true });
});

// Helper: resolve amortization policy for an asset (class-level takes precedence)
function getAssetPolicy(asset, assetClasses) {
  const cls = assetClasses.find(c => c.glAccountId === asset.assetAccountId || c.glAccountName === asset.glAccountName);
  if (cls) {
    return {
      method: cls.method || 'straight-line',
      usefulLifeMonths: cls.usefulLifeMonths,
      decliningRate: cls.decliningRate,
      salvageValue: cls.salvageValue || 0,
      expenseAccountId: cls.expenseAccountId,
      expenseAccountName: cls.expenseAccountName,
      accumAccountId: cls.accumAccountId,
      accumAccountName: cls.accumAccountName,
      className: cls.glAccountName,
    };
  }
  // Fallback to asset-level fields
  return {
    method: 'straight-line',
    usefulLifeMonths: asset.usefulLifeMonths || 60,
    decliningRate: null,
    salvageValue: asset.salvageValue || 0,
    expenseAccountId: asset.expenseAccountId,
    expenseAccountName: asset.expenseAccountName,
    accumAccountId: asset.accumAccountId,
    accumAccountName: asset.accumAccountName,
    className: null,
  };
}

// Helper: calculate monthly amortization amount
function calcMonthlyAmort(asset, policy, monthsElapsed, priorAmortTotal) {
  if (policy.method === 'declining-balance' && policy.decliningRate) {
    const bookValue = asset.originalCost - priorAmortTotal;
    if (bookValue <= policy.salvageValue) return 0;
    const monthlyRate = policy.decliningRate / 12;
    const monthlyAmount = bookValue * monthlyRate;
    const maxAmort = Math.max(0, bookValue - policy.salvageValue);
    return Math.round(Math.min(monthlyAmount, maxAmort) * 100) / 100;
  }
  // Straight-line
  const usefulLife = policy.usefulLifeMonths || 60;
  if (monthsElapsed >= usefulLife) return 0;
  return Math.round(((asset.originalCost - policy.salvageValue) / usefulLife) * 100) / 100;
}

// Helper: check if asset is eligible for amortization
function isEligibleForAmort(asset, policy, now, priorAmortTotal) {
  if (!asset.active) return false;
  const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const acqDate = new Date(asset.acquisitionDate);
  if (acqDate > lastDayOfMonth) return false;
  if (!policy.expenseAccountId || !policy.accumAccountId) return false;

  if (policy.method === 'declining-balance') {
    const bookValue = asset.originalCost - priorAmortTotal;
    return bookValue > (policy.salvageValue || 0) + 0.01;
  }
  // Straight-line: check months remaining
  const acqMonth = acqDate.getFullYear() * 12 + acqDate.getMonth();
  const curMonth = now.getFullYear() * 12 + now.getMonth();
  return (curMonth - acqMonth) < (policy.usefulLifeMonths || 60);
}

/**
 * Build a robust per-asset, per-month amortization map from run history.
 * Tolerates mismatched IDs (after re-syncs), missing IDs (QBO-hydrated runs),
 * and falls back to GL/accum account matching when only one asset is in that group.
 * Returns Map<assetId, Map<'YYYY-MM', amount>>
 */
function buildAmortizationMap(clientData) {
  const map = new Map();
  const assets = clientData.assets || [];
  for (const a of assets) map.set(a.id, new Map());

  const norm = (s) => String(s || '').toLowerCase().trim();

  // Pre-build accum/gl groupings for fallback matching
  const byAccum = new Map();
  const byGl = new Map();
  for (const a of assets) {
    const policy = getAssetPolicy(a, clientData.assetClasses) || {};
    const accum = a.accumAccountName || policy.accumAccountName || '';
    const gl = a.glAccountName || a.assetAccountName || '';
    if (accum) {
      if (!byAccum.has(accum)) byAccum.set(accum, []);
      byAccum.get(accum).push(a);
    }
    if (gl) {
      if (!byGl.has(gl)) byGl.set(gl, []);
      byGl.get(gl).push(a);
    }
  }

  function findAsset(ra) {
    // 1. Strict id match (new or legacy field name)
    const raId = ra.assetId || ra.id;
    if (raId) {
      const hit = assets.find(a => a.id === raId);
      if (hit) return hit;
    }
    // 2. Exact name match
    const raName = ra.assetName || ra.name;
    if (raName) {
      let hit = assets.find(a => a.name === raName);
      if (hit) return hit;
      // 3. Case-insensitive name match
      const target = norm(raName);
      hit = assets.find(a => norm(a.name) === target);
      if (hit) return hit;
    }
    // 4. Single-asset accum-account fallback
    if (ra.accumAccountName) {
      const grp = byAccum.get(ra.accumAccountName);
      if (grp && grp.length === 1) return grp[0];
    }
    // 5. Single-asset GL-account fallback
    if (ra.glAccountName) {
      const grp = byGl.get(ra.glAccountName);
      if (grp && grp.length === 1) return grp[0];
    }
    return null;
  }

  for (const run of clientData.amortizationRuns || []) {
    const month = run.month;
    if (!month) continue;
    const runAssets = run.assets || [];

    // Track per-asset matches in this run, and remember unmatched amount per accum group
    const matchedByAccum = new Map(); // accumName -> matched total
    const unmatchedByAccum = new Map(); // accumName -> unmatched total
    const allMatchedAssetIds = new Set();

    for (const ra of runAssets) {
      const asset = findAsset(ra);
      const amount = parseFloat(ra.amount) || 0;
      const accum = ra.accumAccountName || '';
      if (asset) {
        const inner = map.get(asset.id);
        inner.set(month, (inner.get(month) || 0) + amount);
        allMatchedAssetIds.add(asset.id);
        if (accum) matchedByAccum.set(accum, (matchedByAccum.get(accum) || 0) + amount);
      } else if (accum) {
        unmatchedByAccum.set(accum, (unmatchedByAccum.get(accum) || 0) + amount);
      }
    }

    // Fallback: if a run has totalAmount but no usable per-asset breakdown, distribute by cost
    const matchedTotal = Array.from(allMatchedAssetIds).reduce((s, id) => {
      return s + (map.get(id)?.get(month) || 0);
    }, 0);
    const runTotal = parseFloat(run.totalAmount) || 0;
    if (runTotal > 0 && matchedTotal === 0 && runAssets.length === 0) {
      // Pure totalAmount-only run — distribute across active assets by cost (single GL only)
      const active = assets.filter(a => a.active);
      const totalCost = active.reduce((s, a) => s + (parseFloat(a.originalCost) || 0), 0);
      if (totalCost > 0 && active.length > 0) {
        for (const a of active) {
          const share = (parseFloat(a.originalCost) || 0) / totalCost * runTotal;
          const inner = map.get(a.id);
          inner.set(month, (inner.get(month) || 0) + share);
        }
      }
    }

    // Distribute unmatched-per-accum totals across assets in that accum group, weighted by cost
    for (const [accum, unmatchedAmt] of unmatchedByAccum) {
      if (unmatchedAmt <= 0) continue;
      const grp = byAccum.get(accum) || [];
      if (grp.length === 0) continue;
      const totalCost = grp.reduce((s, a) => s + (parseFloat(a.originalCost) || 0), 0);
      if (totalCost <= 0) continue;
      for (const a of grp) {
        const share = (parseFloat(a.originalCost) || 0) / totalCost * unmatchedAmt;
        const inner = map.get(a.id);
        inner.set(month, (inner.get(month) || 0) + share);
      }
    }
  }

  return map;
}

// Helper: get total prior amortization for an asset (uses robust matcher)
function getPriorAmortTotal(assetIdOrAsset, amortizationRunsOrClientData) {
  // Backward compatible: if a string id + array were passed, use legacy strict match
  if (typeof assetIdOrAsset === 'string' && Array.isArray(amortizationRunsOrClientData)) {
    let total = 0;
    amortizationRunsOrClientData.forEach(run => {
      const assetEntry = run.assets?.find(a => a.id === assetIdOrAsset || a.assetId === assetIdOrAsset);
      if (assetEntry) total += parseFloat(assetEntry.amount) || 0;
    });
    return total;
  }
  // New form: pass the asset object and the full clientData
  const asset = assetIdOrAsset;
  const clientData = amortizationRunsOrClientData;
  const map = buildAmortizationMap(clientData);
  const inner = map.get(asset.id);
  if (!inner) return 0;
  let total = 0;
  for (const v of inner.values()) total += v;
  return total;
}

// ----- Continuity schedule helpers -----

/**
 * Given a target month "YYYY-MM" and a fiscal year-end "MM-DD",
 * return the YYYY-MM of the start of the fiscal year that contains targetMonth.
 */
function getFiscalYearStart(targetMonth, fiscalYearEnd) {
  const fye = (fiscalYearEnd && /^\d{2}-\d{2}$/.test(fiscalYearEnd)) ? fiscalYearEnd : '12-31';
  const fyeMonth = parseInt(fye.split('-')[0], 10); // 1-12
  const fyStartMonth = (fyeMonth % 12) + 1;          // 1-12
  const { year: ty, monthIndex: tmIdx } = parseMonthStr(targetMonth);
  const tm = tmIdx + 1; // 1-indexed
  const fyStartYear = tm >= fyStartMonth ? ty : ty - 1;
  return formatMonthStr(fyStartYear, fyStartMonth - 1);
}

/**
 * Build a fiscal-year-to-date continuity schedule snapshot from current clientData.
 *  - Rows: per asset, grouped by GL account
 *  - Columns: opening accum + one column per fiscal-year month up to asOfMonth
 *  - Subtotals per GL account, plus a grand total
 */
function buildContinuitySchedule(clientData, asOfMonth, fiscalYearEnd) {
  const fyStart = getFiscalYearStart(asOfMonth, fiscalYearEnd);
  const months = [];
  let { year: y, monthIndex: m } = parseMonthStr(fyStart);
  const { year: endY, monthIndex: endM } = parseMonthStr(asOfMonth);
  while (y * 12 + m <= endY * 12 + endM) {
    months.push(formatMonthStr(y, m));
    ({ year: y, monthIndex: m } = addMonth(y, m, 1));
  }

  const groups = new Map();
  const round2 = (n) => Math.round((n || 0) * 100) / 100;

  // Use robust amortization map (handles id/name/accum/cost-ratio matching)
  const amortMap = buildAmortizationMap(clientData);

  for (const a of clientData.assets || []) {
    if (!a.active) continue;
    const policy = getAssetPolicy(a, clientData.assetClasses) || {};
    const glName = a.glAccountName || a.assetAccountName || 'Unclassified';

    let openingAccum = 0;
    const monthlyAmort = {};
    for (const mo of months) monthlyAmort[mo] = 0;

    const inner = amortMap.get(a.id) || new Map();
    for (const [runMonth, amount] of inner) {
      if (runMonth < fyStart) {
        openingAccum += amount;
      } else if (Object.prototype.hasOwnProperty.call(monthlyAmort, runMonth)) {
        monthlyAmort[runMonth] += amount;
      }
    }

    const fyAmort = months.reduce((s, mo) => s + monthlyAmort[mo], 0);
    const closingAccum = openingAccum + fyAmort;
    const netBookValue = (parseFloat(a.originalCost) || 0) - closingAccum;

    if (!groups.has(glName)) groups.set(glName, { glAccountName: glName, assets: [] });

    groups.get(glName).assets.push({
      assetId: a.id,
      name: a.name,
      description: a.description || '',
      glAccountName: glName,
      vendorName: a.vendorName || '',
      acquisitionDate: a.acquisitionDate || '',
      originalCost: round2(a.originalCost),
      salvageValue: round2(a.salvageValue),
      method: policy.method || '',
      usefulLifeMonths: policy.usefulLifeMonths || null,
      expenseAccountName: policy.expenseAccountName || a.expenseAccountName || '',
      accumAccountName: policy.accumAccountName || a.accumAccountName || '',
      openingAccum: round2(openingAccum),
      monthlyAmort: Object.fromEntries(Object.entries(monthlyAmort).map(([k, v]) => [k, round2(v)])),
      closingAccum: round2(closingAccum),
      netBookValue: round2(netBookValue),
    });
  }

  // Subtotals + grand total
  const glAccounts = [];
  let totalCost = 0, totalOpening = 0, totalClosing = 0, totalNBV = 0;
  const totalMonthly = {};
  for (const mo of months) totalMonthly[mo] = 0;

  for (const group of groups.values()) {
    let cost = 0, opening = 0, closing = 0, nbv = 0;
    const monthly = {};
    for (const mo of months) monthly[mo] = 0;
    for (const asset of group.assets) {
      cost += asset.originalCost;
      opening += asset.openingAccum;
      closing += asset.closingAccum;
      nbv += asset.netBookValue;
      for (const mo of months) monthly[mo] += asset.monthlyAmort[mo];
    }
    group.subtotal = {
      cost: round2(cost),
      openingAccum: round2(opening),
      monthlyAmort: Object.fromEntries(Object.entries(monthly).map(([k, v]) => [k, round2(v)])),
      closingAccum: round2(closing),
      netBookValue: round2(nbv),
    };
    glAccounts.push(group);
    totalCost += cost;
    totalOpening += opening;
    totalClosing += closing;
    totalNBV += nbv;
    for (const mo of months) totalMonthly[mo] += monthly[mo];
  }

  return {
    asOfMonth,
    fiscalYearStart: fyStart,
    fiscalYearEnd: fiscalYearEnd || '12-31',
    months,
    glAccounts,
    total: {
      cost: round2(totalCost),
      openingAccum: round2(totalOpening),
      monthlyAmort: Object.fromEntries(Object.entries(totalMonthly).map(([k, v]) => [k, round2(v)])),
      closingAccum: round2(totalClosing),
      netBookValue: round2(totalNBV),
    },
    generatedAt: new Date().toISOString(),
  };
}

/**
 * Reconcile a continuity schedule against QBO trial balance.
 * Compares (a) cost account balances and (b) accumulated amortization balances per GL account.
 * Returns { checks: [...], allPassed: boolean }
 */
async function reconcileScheduleToQBO(clientId, schedule) {
  const checks = [];
  const TOLERANCE = 0.01;

  if (!qbo.isConnected(clientId)) {
    return { checks: [], allPassed: false, error: 'QBO not connected' };
  }

  // Pull TB as-of last day of asOfMonth
  const { year, monthIndex } = parseMonthStr(schedule.asOfMonth);
  const asOf = lastDayOfMonthDate(year, monthIndex).toISOString().split('T')[0];

  let tb;
  try {
    tb = await qbo.getTrialBalance(asOf, asOf, clientId);
  } catch (e) {
    return { checks: [], allPassed: false, error: `Failed to fetch trial balance: ${e.message}` };
  }

  // Flatten TB rows into a name -> balance map
  const tbBalances = new Map();
  function walkRows(rows) {
    if (!rows) return;
    for (const row of rows) {
      if (row.ColData) {
        const acctName = row.ColData[0]?.value || '';
        const debit = parseFloat(String(row.ColData[1]?.value || '').replace(/[$,\s]/g, '')) || 0;
        const credit = parseFloat(String(row.ColData[2]?.value || '').replace(/[$,\s]/g, '')) || 0;
        if (acctName) tbBalances.set(acctName, debit - credit);
      }
      if (row.Rows?.Row) walkRows(row.Rows.Row);
      if (row.Rows && Array.isArray(row.Rows)) walkRows(row.Rows);
    }
  }
  walkRows(tb?.Rows?.Row);
  console.log(`[fixed-assets] TB accounts found (${tbBalances.size}):`, Array.from(tbBalances.keys()).join(', '));

  // Fuzzy account name lookup — handles account number prefixes (e.g. "1500 Computer hardware"
  // vs "Computer hardware"), case differences, and leading/trailing whitespace.
  function findTBBalance(targetName) {
    if (!targetName) return undefined;
    // 1. Exact match
    if (tbBalances.has(targetName)) return tbBalances.get(targetName);
    // 2. Case-insensitive match
    const targetLower = targetName.toLowerCase().trim();
    for (const [key, val] of tbBalances) {
      if (key.toLowerCase().trim() === targetLower) return val;
    }
    // 3. Strip leading account number prefix (e.g. "1500 " or "1500-") from both sides
    const stripNum = s => s.replace(/^\d{3,5}[\s\-.:]+/, '').trim().toLowerCase();
    const targetStripped = stripNum(targetName);
    for (const [key, val] of tbBalances) {
      if (stripNum(key) === targetStripped) return val;
    }
    // 4. Check if one contains the other (for partial matches like sub-accounts)
    for (const [key, val] of tbBalances) {
      const keyLower = key.toLowerCase().trim();
      if (keyLower.includes(targetLower) || targetLower.includes(keyLower)) return val;
    }
    return undefined;
  }

  for (const group of schedule.glAccounts) {
    const glName = group.glAccountName;
    // Cost check: compare schedule cost subtotal to TB balance for this account
    const tbCost = findTBBalance(glName);
    const schedCost = group.subtotal.cost;
    if (tbCost !== undefined) {
      const diff = Math.round((tbCost - schedCost) * 100) / 100;
      checks.push({
        type: 'cost',
        glAccount: glName,
        schedule: schedCost,
        qbo: Math.round(tbCost * 100) / 100,
        difference: diff,
        status: Math.abs(diff) < TOLERANCE ? 'pass' : 'fail',
      });
    } else {
      checks.push({
        type: 'cost',
        glAccount: glName,
        schedule: schedCost,
        qbo: null,
        difference: null,
        status: 'missing',
        note: `account "${glName}" not found on trial balance`,
      });
    }

    // Accum check: try to find an accum account that matches this GL account
    const accumName = group.assets[0]?.accumAccountName || '';
    if (accumName) {
      const tbAccum = findTBBalance(accumName);
      const schedAccum = group.subtotal.closingAccum;
      if (tbAccum !== undefined) {
        // Accum is a contra asset — TB credit balance shows as negative
        const tbAccumAbs = Math.abs(tbAccum);
        const diff = Math.round((tbAccumAbs - schedAccum) * 100) / 100;
        checks.push({
          type: 'accum',
          glAccount: glName,
          accumAccount: accumName,
          schedule: schedAccum,
          qbo: Math.round(tbAccumAbs * 100) / 100,
          difference: diff,
          status: Math.abs(diff) < TOLERANCE ? 'pass' : 'fail',
        });
      } else {
        checks.push({
          type: 'accum',
          glAccount: glName,
          accumAccount: accumName,
          schedule: schedAccum,
          qbo: null,
          difference: null,
          status: 'missing',
          note: `accum account "${accumName}" not found on trial balance`,
        });
      }
    }
  }

  const allPassed = checks.length > 0 && checks.every(c => c.status === 'pass');
  return { checks, allPassed, asOf };
}

// ----- Helpers for determining the target amortization period -----

// Parse "YYYY-MM" string into { year, monthIndex (0-11) }
function parseMonthStr(s) {
  const [y, m] = s.split('-').map(Number);
  return { year: y, monthIndex: m - 1 };
}
function formatMonthStr(year, monthIndex) {
  return `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
}
function addMonth(year, monthIndex, delta) {
  const total = year * 12 + monthIndex + delta;
  return { year: Math.floor(total / 12), monthIndex: total % 12 };
}
// First day of a month as a Date for comparisons
function firstDayOfMonth(year, monthIndex) {
  return new Date(year, monthIndex, 1);
}
// Last day of a month as Date
function lastDayOfMonthDate(year, monthIndex) {
  return new Date(year, monthIndex + 1, 0);
}

/**
 * Determine the target amortization month.
 * Rules:
 *  1. Start with previous calendar month (e.g. if today is April, start with March).
 *  2. If QBO's book close date covers that month (last day <= close date), advance forward.
 *  3. If amortization has already been run for that month, advance forward.
 *  4. Never go beyond the current calendar month.
 * Returns { month: 'YYYY-MM', reason: string, closeDate: string|null, alreadyRun: bool }
 */
/**
 * Hydrate clientData.amortizationRuns from QBO journal entries.
 *
 * QBO is the source of truth — this does a FULL REPLACE of the amortization run history
 * on the client, so any edit, deletion, or backdated adjustment made directly in QBO
 * flows through on the next sync. Local-only metadata (snapshot, reconciliation) is
 * carried forward for months that still exist in QBO, keyed by month.
 *
 * Safe no-op if QBO is not connected. If the QBO fetch fails we leave the existing
 * local runs untouched rather than blowing them away.
 */
async function hydrateAmortizationRunsFromQBO(clientId, clientData) {
  if (!qbo.isConnected(clientId)) return clientData;
  try {
    const qboRuns = await qbo.getAmortizationRunsFromQBO(clientId);
    if (!Array.isArray(qboRuns)) return clientData;

    // Preserve local-only metadata (snapshot, reconciliation) keyed by month so a
    // fresh rehydrate doesn't wipe out cached snapshots for months QBO still has.
    const localByMonth = new Map((clientData.amortizationRuns || []).map(r => [r.month, r]));
    const replaced = qboRuns.map(qr => {
      const local = localByMonth.get(qr.month);
      if (local) {
        return {
          ...qr,
          snapshot: local.snapshot || null,
          reconciliation: local.reconciliation || null,
          // Keep the original local ranAt if QBO didn't give us one
          ranAt: qr.ranAt || local.ranAt || null,
        };
      }
      return qr;
    });
    replaced.sort((a, b) => (a.month || '').localeCompare(b.month || ''));
    clientData.amortizationRuns = replaced;
    clientData.amortizationRunsSyncedAt = new Date().toISOString();
    saveFixedAssets(fixedAssetsData);
  } catch (e) {
    console.error('Failed to hydrate amortization runs from QBO:', e.message);
  }
  return clientData;
}

function determineAmortizationMonth(clientData, closeDate) {
  const today = new Date();
  // Start with previous month
  let { year, monthIndex } = addMonth(today.getFullYear(), today.getMonth(), -1);
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();

  const closeDateObj = closeDate ? new Date(closeDate) : null;
  const runMonths = new Set((clientData.amortizationRuns || []).map(r => r.month));

  let safety = 0;
  while (safety++ < 120) {
    const monthStr = formatMonthStr(year, monthIndex);
    const lastDay = lastDayOfMonthDate(year, monthIndex);

    // Advance past closed periods
    if (closeDateObj && lastDay <= closeDateObj) {
      ({ year, monthIndex } = addMonth(year, monthIndex, 1));
      continue;
    }
    // Advance past months already run
    if (runMonths.has(monthStr)) {
      ({ year, monthIndex } = addMonth(year, monthIndex, 1));
      continue;
    }
    // Cap at current calendar month (don't amortize a future month)
    if (year * 12 + monthIndex > currentYear * 12 + currentMonth) {
      // All eligible months are already covered — fall back to current month
      return {
        month: formatMonthStr(currentYear, currentMonth),
        closeDate: closeDate || null,
        alreadyRun: runMonths.has(formatMonthStr(currentYear, currentMonth)),
        reason: 'all prior months already amortized',
      };
    }
    return {
      month: monthStr,
      closeDate: closeDate || null,
      alreadyRun: false,
      reason: closeDateObj ? 'first open month after close date' : 'first unposted prior month',
    };
  }
  // Fallback (should never hit)
  return { month: formatMonthStr(currentYear, currentMonth - 1), closeDate: closeDate || null, alreadyRun: false, reason: 'fallback' };
}

/**
 * Compute the full preview for a given month.
 * Returns { lines, totalAmount }
 */
function computeAmortizationPreview(clientData, targetMonth) {
  const { year, monthIndex } = parseMonthStr(targetMonth);
  // "now" for eligibility checks = last day of target month
  const asOf = lastDayOfMonthDate(year, monthIndex);

  const lines = [];
  let totalAmount = 0;

  clientData.assets.forEach(a => {
    const policy = getAssetPolicy(a, clientData.assetClasses);
    const priorAmort = getPriorAmortTotal(a, clientData);
    const acqDate = new Date(a.acquisitionDate);
    const monthsElapsed = (asOf.getFullYear() * 12 + asOf.getMonth()) - (acqDate.getFullYear() * 12 + acqDate.getMonth());

    if (!isEligibleForAmort(a, policy, asOf, priorAmort)) return;

    const monthly = calcMonthlyAmort(a, policy, monthsElapsed, priorAmort);
    if (monthly <= 0) return;

    totalAmount += monthly;
    lines.push({
      assetId: a.id,
      assetName: a.name,
      description: a.description || '',
      amount: monthly,
      method: policy.method,
      className: policy.glAccountName || a.glAccountName || a.assetAccountName || '',
      expenseAccountId: policy.expenseAccountId,
      expenseAccountName: policy.expenseAccountName,
      accumAccountId: policy.accumAccountId,
      accumAccountName: policy.accumAccountName,
    });
  });

  return { lines, totalAmount };
}

// Determine the current closing period based purely on the QBO close date.
// The closing period is always the month immediately after the close date.
// If no close date is set, defaults to last month.
app.get('/api/admin/clients/:clientId/close-period', requireAdmin, async (req, res) => {
  try {
    let closeDate = null;
    if (qbo.isConnected(req.params.clientId)) {
      try { closeDate = await qbo.getBookCloseDate(req.params.clientId); } catch (_) {}
    }
    const today = new Date();
    let closingMonth;
    let reason;
    if (closeDate) {
      const cd = new Date(closeDate + 'T00:00:00');
      // Closing period = month after close date
      const nextMonth = new Date(cd.getFullYear(), cd.getMonth() + 1, 1);
      closingMonth = formatMonthStr(nextMonth.getFullYear(), nextMonth.getMonth());
      reason = `first open month after close date (${closeDate})`;
    } else {
      // No close date — default to previous month
      const prev = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      closingMonth = formatMonthStr(prev.getFullYear(), prev.getMonth());
      reason = 'no close date set — defaulting to last month';
    }
    res.json({ month: closingMonth, closeDate, reason });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Fiscal year close calendar — returns per-month status for each module
app.get('/api/admin/clients/:clientId/close-calendar', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    const client = CLIENTS[clientId];
    if (!client) return res.status(404).json({ error: 'client not found' });

    // Determine fiscal year months
    // fiscalYearEnd is "MM-DD" (e.g. "02-28" for Feb 28 year-end)
    const fye = client.fiscalYearEnd || '12-31';
    const [fyeMonth, fyeDay] = fye.split('-').map(Number);
    const today = new Date();

    // Fiscal year that contains "today": starts month after FYE of previous calendar year
    // e.g. FYE = 02-28 → FY starts March. If today is Jan 2027, current FY started Mar 2026.
    let fyStartYear, fyStartMonth;
    if (fyeMonth === 12) {
      // Calendar year-end: FY = Jan to Dec of current year
      fyStartYear = today.getFullYear();
      fyStartMonth = 0; // January
    } else {
      // FY starts the month after FYE month
      fyStartMonth = fyeMonth; // fyeMonth is 0-based after subtracting... no, it's 1-based from MM
      // fyeMonth is 1-based (02 = February). FY starts month fyeMonth (0-based = fyeMonth)
      // e.g. FYE = Feb (02) → FY starts March (month index 2)
      const fyStartMonthIdx = fyeMonth; // fyeMonth=2 → March = index 2 ✓
      if (today.getMonth() >= fyStartMonthIdx) {
        fyStartYear = today.getFullYear();
      } else {
        fyStartYear = today.getFullYear() - 1;
      }
      fyStartMonth = fyStartMonthIdx;
    }

    // Build 12 months
    const months = [];
    for (let i = 0; i < 12; i++) {
      const m = (fyStartMonth + i) % 12;
      const y = fyStartYear + Math.floor((fyStartMonth + i) / 12);
      months.push(formatMonthStr(y, m));
    }

    // Gather run data from each module
    const faData = getClientAssets(clientId);
    const faRuns = new Set((faData.amortizationRuns || []).map(r => r.month));

    const ppData = getClientPrepaid(clientId);
    const ppConfigured = !!ppData.prepaidAccount;
    const ppRuns = new Set((ppData.amortizationRuns || []).map(r => r.month));
    const ppScanned = new Set(ppData.scannedMonths || []);

    const alData = getClientAccruedLiab(clientId);
    const alConfigured = !!alData.accruedLiabilitiesAccount;
    // A month is "complete" if JE was posted OR analysis found $0 to accrue
    const alComplete = new Set(
      (alData.analysisRuns || []).filter(r => {
        if (r.accrualJE) return true; // JE posted
        // Check if analysis found nothing to accrue
        const aAccts = (r.partA?.accounts || []).filter(a => a.status !== 'dismissed');
        const bTxns = (r.partB?.transactions || []).filter(t => t.status !== 'dismissed');
        const totalA = aAccts.reduce((s, a) => s + (Number(a.accrualAmount) || 0), 0);
        const totalB = bTxns.filter(t => !t.overlapWithPartA || !aAccts.some(a => a.accountId === t.accountId)).reduce((s, t) => s + (Number(t.accrualAmount) || 0), 0);
        return Math.round((totalA + totalB) * 100) / 100 <= 0;
      }).map(r => r.month)
    );

    const shiData = getClientShareholderInvoices(clientId);
    const shiConfigured = !!shiData.shareholderLoanAccount;
    // For shareholder invoices, a month is "complete" if there are any posted invoices for that month
    const shiPostedMonths = new Set(
      (shiData.invoices || []).filter(i => i.status === 'posted').map(i => i.closeMonth)
    );

    // Determine closing month
    let closeDate = null;
    if (qbo.isConnected(clientId)) {
      try { closeDate = await qbo.getBookCloseDate(clientId); } catch (_) {}
    }
    let closingMonth = null;
    if (closeDate) {
      const cd = new Date(closeDate + 'T00:00:00');
      const nm = new Date(cd.getFullYear(), cd.getMonth() + 1, 1);
      closingMonth = formatMonthStr(nm.getFullYear(), nm.getMonth());
    } else {
      // Fallback: default to last month (same logic as close-period endpoint)
      const now = new Date();
      const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      closingMonth = formatMonthStr(lastMonth.getFullYear(), lastMonth.getMonth());
    }

    const calendar = months.map(month => {
      const isFuture = closingMonth ? month > closingMonth : false;
      const isCurrent = month === closingMonth;
      const isClosed = closeDate ? (() => {
        const { year, monthIndex } = parseMonthStr(month);
        const lastDay = lastDayOfMonthDate(year, monthIndex);
        return lastDay <= new Date(closeDate + 'T00:00:00');
      })() : false;

      return {
        month,
        isCurrent,
        isClosed,
        isFuture: !isClosed && !isCurrent,
        modules: {
          shareholderInvoices: !shiConfigured ? 'skipped' : (shiPostedMonths.has(month) ? 'complete' : (isClosed ? 'closed' : 'pending')),
          fixedAssets: faRuns.has(month) ? 'complete' : (isClosed ? 'closed' : 'pending'),
          prepaidExpenses: !ppConfigured ? 'skipped' : ((ppRuns.has(month) || ppScanned.has(month)) ? 'complete' : (isClosed ? 'closed' : 'pending')),
          accruedLiabilities: !alConfigured ? 'skipped' : (alComplete.has(month) ? 'complete' : (isClosed ? 'closed' : 'pending')),
        },
      };
    });

    res.json({ fiscalYear: `${months[0]} to ${months[11]}`, months: calendar, closingMonth, closeDate });
  } catch (e) {
    console.error('[close-calendar] error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Get QBO book close date
app.get('/api/admin/clients/:clientId/book-close-date', requireAdmin, async (req, res) => {
  try {
    if (!qbo.isConnected(req.params.clientId)) {
      return res.json({ connected: false, closeDate: null });
    }
    const closeDate = await qbo.getBookCloseDate(req.params.clientId);
    res.json({ connected: true, closeDate });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// NOTE: there is no POST endpoint for book-close-date. QBO's Preferences API
// silently rejects third-party writes to BookCloseDate — it returns 200 OK but
// echoes the OLD value back instead of applying the change. Confirmed via live
// test on a company with the "Close the books" feature enabled. The UI shows
// the current value read-only and links out to QBO for edits.

// List attachments for a transaction
// GET /api/admin/clients/:clientId/qbo/attachments?txnId=123&txnType=Purchase
app.get('/api/admin/clients/:clientId/qbo/attachments', requireAdmin, async (req, res) => {
  try {
    if (!qbo.isConnected(req.params.clientId)) {
      return res.status(400).json({ error: 'QBO not connected' });
    }
    const { txnId, txnType } = req.query;
    if (!txnId || !txnType) {
      return res.status(400).json({ error: 'txnId and txnType required' });
    }
    const attachments = await qbo.getTransactionAttachments(txnId, txnType, req.params.clientId);
    res.json({ attachments });
  } catch (e) {
    console.error('[attachments] list error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Download a specific attachment (proxied through the server so the client
// never sees QBO credentials). Returns the raw file bytes.
// GET /api/admin/clients/:clientId/qbo/attachments/:attachableId/download
app.get('/api/admin/clients/:clientId/qbo/attachments/:attachableId/download', requireAdmin, async (req, res) => {
  try {
    if (!qbo.isConnected(req.params.clientId)) {
      return res.status(400).json({ error: 'QBO not connected' });
    }
    const file = await qbo.downloadAttachment(req.params.attachableId, req.params.clientId);
    res.setHeader('Content-Type', file.contentType);
    res.setHeader('Content-Disposition', `inline; filename="${file.fileName}"`);
    res.setHeader('Content-Length', file.size);
    res.send(file.buffer);
  } catch (e) {
    console.error('[attachments] download error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Smoke test: pick the most recent bill/expense for a client that has an
// attachment, download it, and return a summary. Lets us verify the whole
// attachment pipeline end-to-end before building on top of it.
// GET /api/admin/clients/:clientId/qbo/attachments/test
app.get('/api/admin/clients/:clientId/qbo/attachments/test', requireAdmin, async (req, res) => {
  try {
    if (!qbo.isConnected(req.params.clientId)) {
      return res.status(400).json({ error: 'QBO not connected' });
    }
    const clientId = req.params.clientId;

    // Pull a handful of recent bills and purchases and find the first one with an attachment.
    // QBO helpers return { QueryResponse: { Bill|Purchase: [...] } } so we unwrap here.
    const today = new Date().toISOString().slice(0, 10);
    const [billsRes, purchasesRes] = await Promise.all([
      qbo.getBills('2020-01-01', today, 50, clientId).catch(() => null),
      qbo.getExpenseTransactions('2020-01-01', today, 50, clientId).catch(() => null),
    ]);

    const bills = billsRes?.QueryResponse?.Bill || [];
    const purchases = purchasesRes?.QueryResponse?.Purchase || [];

    const candidates = [
      ...bills.map((b) => ({ id: b.Id, type: 'Bill', date: b.TxnDate, amount: b.TotalAmt, vendor: b.VendorRef?.name })),
      ...purchases.map((p) => ({ id: p.Id, type: 'Purchase', date: p.TxnDate, amount: p.TotalAmt, vendor: p.EntityRef?.name })),
    ].sort((a, b) => (b.date || '').localeCompare(a.date || ''));

    const scanned = [];
    let hit = null;
    for (const txn of candidates.slice(0, 30)) {
      const attachments = await qbo.getTransactionAttachments(txn.id, txn.type, clientId).catch(() => []);
      scanned.push({ ...txn, attachmentCount: attachments.length });
      if (attachments.length > 0) {
        hit = { txn, attachments };
        break;
      }
    }

    if (!hit) {
      return res.json({
        ok: false,
        message: 'No attachments found on the 30 most recent bills/purchases',
        scanned,
      });
    }

    // Try downloading the first attachment to confirm end-to-end works
    const first = hit.attachments[0];
    let downloadResult = null;
    try {
      const file = await qbo.downloadAttachment(first.id, clientId);
      downloadResult = {
        ok: true,
        fileName: file.fileName,
        contentType: file.contentType,
        sizeBytes: file.size,
      };
    } catch (e) {
      downloadResult = { ok: false, error: e.message };
    }

    res.json({
      ok: true,
      transaction: hit.txn,
      attachments: hit.attachments,
      download: downloadResult,
      scannedCount: scanned.length,
    });
  } catch (e) {
    console.error('[attachments] test error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Preview amortization for a client
app.get('/api/admin/clients/:clientId/fixed-assets/preview-amortization', requireAdmin, async (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  const clientId = req.params.clientId;

  // Optional manual month override via query string (YYYY-MM)
  const overrideMonth = req.query.month && /^\d{4}-\d{2}$/.test(req.query.month) ? req.query.month : null;

  // Fetch QBO book close date and hydrate amortization run history from QBO (best-effort)
  let closeDate = null;
  if (qbo.isConnected(clientId)) {
    try { closeDate = await qbo.getBookCloseDate(clientId); } catch (e) { /* ignore */ }
    await hydrateAmortizationRunsFromQBO(clientId, clientData);
  }

  let targetMonth, reason;
  if (overrideMonth) {
    targetMonth = overrideMonth;
    reason = 'manual override';
  } else {
    const determined = determineAmortizationMonth(clientData, closeDate);
    targetMonth = determined.month;
    reason = determined.reason;
  }

  const alreadyRun = clientData.amortizationRuns.find(r => r.month === targetMonth);
  if (alreadyRun) {
    return res.json({ alreadyRun: true, runDetails: alreadyRun, month: targetMonth, closeDate, reason });
  }

  const { lines, totalAmount } = computeAmortizationPreview(clientData, targetMonth);

  res.json({
    alreadyRun: false,
    month: targetMonth,
    closeDate,
    reason,
    eligibleCount: lines.length,
    totalAmount: Math.round(totalAmount * 100) / 100,
    lines,
  });
});

// Run amortization for a client and post JE to QBO
app.post('/api/admin/clients/:clientId/fixed-assets/run-amortization', requireAdmin, async (req, res) => {
  try {
    const clientData = getClientAssets(req.params.clientId);
    const amortClientId = req.params.clientId;

    if (!qbo.isConnected(amortClientId)) {
      return res.status(400).json({ error: 'QuickBooks is not connected for this client. Please connect QBO in the client detail view first.' });
    }

    // Determine target month: accept override from body, else auto-determine
    const overrideMonth = req.body?.month && /^\d{4}-\d{2}$/.test(req.body.month) ? req.body.month : null;

    let closeDate = null;
    try { closeDate = await qbo.getBookCloseDate(amortClientId); } catch (e) { /* ignore */ }
    // Hydrate from QBO so we never re-post a month that already exists there
    await hydrateAmortizationRunsFromQBO(amortClientId, clientData);

    const targetMonth = overrideMonth || determineAmortizationMonth(clientData, closeDate).month;

    // Guard: don't post to a closed period
    if (closeDate) {
      const { year, monthIndex } = parseMonthStr(targetMonth);
      const lastDay = lastDayOfMonthDate(year, monthIndex);
      if (lastDay <= new Date(closeDate)) {
        return res.status(400).json({ error: `Target month ${targetMonth} falls within closed period (close date ${closeDate})` });
      }
    }

    const alreadyRun = clientData.amortizationRuns.find(r => r.month === targetMonth);
    if (alreadyRun) {
      return res.status(400).json({ error: `Amortization already run for ${targetMonth}` });
    }

    const { year, monthIndex } = parseMonthStr(targetMonth);
    const asOf = lastDayOfMonthDate(year, monthIndex);
    const lastDayStr = asOf.toISOString().split('T')[0];

    const { lines: previewLines, totalAmount } = computeAmortizationPreview(clientData, targetMonth);

    if (previewLines.length === 0) {
      return res.status(400).json({ error: `No eligible assets for amortization in ${targetMonth}` });
    }

    const jeLines = [];
    const assetAmounts = [];
    for (const line of previewLines) {
      assetAmounts.push({
        assetId: line.assetId,
        id: line.assetId, // legacy alias
        assetName: line.assetName,
        name: line.assetName, // legacy alias
        description: line.description || '',
        glAccountName: line.className || '',
        expenseAccountId: line.expenseAccountId,
        expenseAccountName: line.expenseAccountName,
        accumAccountId: line.accumAccountId,
        accumAccountName: line.accumAccountName,
        method: line.method || '',
        amount: line.amount,
      });
      jeLines.push({ accountId: line.expenseAccountId, accountName: line.expenseAccountName, description: `Amortization - ${line.assetName}`, amount: line.amount, type: 'debit' });
      jeLines.push({ accountId: line.accumAccountId, accountName: line.accumAccountName, description: `Amortization - ${line.assetName}`, amount: line.amount, type: 'credit' });
    }

    const clientName = CLIENTS[amortClientId]?.name || amortClientId;
    const result = await qbo.createJournalEntry({
      date: lastDayStr,
      memo: `Fixed asset amortization - ${targetMonth} - ${clientName}`,
      lines: jeLines,
    }, amortClientId);

    const runRecord = {
      month: targetMonth,
      ranAt: new Date().toISOString(),
      journalEntryId: result.Id || null,
      totalAmount: Math.round(totalAmount * 100) / 100,
      assetCount: assetAmounts.length,
      assets: assetAmounts,
    };

    clientData.amortizationRuns.push(runRecord);

    // Build a continuity schedule snapshot reflecting state immediately after this run
    const fye = CLIENTS[amortClientId]?.fiscalYearEnd || null;
    const snapshot = buildContinuitySchedule(clientData, targetMonth, fye);
    runRecord.snapshot = snapshot;
    clientData.latestSchedule = snapshot;

    saveFixedAssets(fixedAssetsData);

    // Run reconciliation against QBO trial balance (best-effort, non-blocking on failure)
    let reconciliation = null;
    try {
      reconciliation = await reconcileScheduleToQBO(amortClientId, snapshot);
      runRecord.reconciliation = reconciliation;
      saveFixedAssets(fixedAssetsData);
    } catch (e) {
      console.error('Reconciliation error:', e.message);
      reconciliation = { error: e.message };
    }

    res.json({ success: true, run: runRecord, snapshot, reconciliation });
  } catch (error) {
    console.error('Amortization run error:', error.message);
    res.status(500).json({ error: `Failed to post journal entry: ${error.message}` });
  }
});

// Get the latest continuity schedule snapshot for a client
app.get('/api/admin/clients/:clientId/fixed-assets/continuity-schedule', requireAdmin, async (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  // Amortization runs are rehydrated during "Sync from QBO", not on every read.
  // Determine the most recent run month, or fall back to current month
  const runs = clientData.amortizationRuns || [];
  const sorted = runs.slice().sort((a, b) => (b.month || '').localeCompare(a.month || ''));
  const asOfMonth = sorted[0]?.month || formatMonthStr(new Date().getFullYear(), new Date().getMonth());
  const fye = CLIENTS[req.params.clientId]?.fiscalYearEnd || null;
  // Prefer stored snapshot if it matches the latest run; else recompute
  let schedule = clientData.latestSchedule;
  if (!schedule || schedule.asOfMonth !== asOfMonth) {
    schedule = buildContinuitySchedule(clientData, asOfMonth, fye);
  }
  res.json({ schedule });
});

// Re-run reconciliation on demand
app.post('/api/admin/clients/:clientId/fixed-assets/reconcile', requireAdmin, async (req, res) => {
  try {
    const clientData = getClientAssets(req.params.clientId);
    // Reconciliation reads cached run history; use "Sync from QBO" to refresh it.
    const runs = clientData.amortizationRuns || [];
    const sorted = runs.slice().sort((a, b) => (b.month || '').localeCompare(a.month || ''));
    const asOfMonth = sorted[0]?.month;
    if (!asOfMonth) return res.status(400).json({ error: 'No amortization runs to reconcile' });
    const fye = CLIENTS[req.params.clientId]?.fiscalYearEnd || null;
    const schedule = buildContinuitySchedule(clientData, asOfMonth, fye);
    const reconciliation = await reconcileScheduleToQBO(req.params.clientId, schedule);
    res.json({ schedule, reconciliation });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ========================================
// PREPAID EXPENSES — per-client, file-backed
// ========================================
// Schema per client:
// {
//   prepaidAccount: { id, name },       // single balance-sheet Prepaid Expenses GL account
//   items: [
//     {
//       id, vendor, description,
//       totalAmount,                    // original invoice total being amortized
//       startDate, endDate,             // inclusive service window (YYYY-MM-DD)
//       expenseAccountId, expenseAccountName,
//       openingBalance,                 // un-amortized amount at time of entry (for mid-life imports)
//       sourceTxnId, sourceTxnType,     // link to QBO txn if created by Part B scanner
//       createdAt, completedAt,
//     }
//   ],
//   amortizationRuns: [
//     { month, ranAt, journalEntryId, totalAmount, itemCount, items: [...] }
//   ]
// }

const PREPAID_FILE = dataPath('prepaid-expenses.json');

function loadPrepaidExpenses() {
  try {
    if (fs.existsSync(PREPAID_FILE)) {
      return JSON.parse(fs.readFileSync(PREPAID_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Failed to load prepaid-expenses.json:', e.message);
  }
  return {};
}

function savePrepaidExpenses(data) {
  fs.writeFileSync(PREPAID_FILE, JSON.stringify(data, null, 2), 'utf-8');
}

let prepaidData = loadPrepaidExpenses();

function getClientPrepaid(clientId) {
  if (!prepaidData[clientId]) {
    prepaidData[clientId] = { prepaidAccount: null, items: [], amortizationRuns: [], scanThreshold: 500 };
  }
  const c = prepaidData[clientId];
  if (!c.items) c.items = [];
  if (!c.amortizationRuns) c.amortizationRuns = [];
  if (!c.scannedMonths) c.scannedMonths = [];
  if (c.prepaidAccount === undefined) c.prepaidAccount = null;
  if (c.scanThreshold === undefined) c.scanThreshold = 500;
  return c;
}

// Utility: count inclusive calendar months between two YYYY-MM-DD dates
// (e.g. 2026-01-15 → 2026-06-20 = 6 months). Full-month proration per spec.
function monthsBetweenInclusive(startDate, endDate) {
  const s = new Date(startDate + 'T00:00:00');
  const e = new Date(endDate + 'T00:00:00');
  return (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth()) + 1;
}

// Months already recognized for this item as of end of targetMonth (exclusive of target)
// e.g. item starts 2026-01-01, target is 2026-04 → 3 months recognized prior (Jan/Feb/Mar)
function monthsRecognizedBefore(item, targetMonth) {
  const { year: ty, monthIndex: tm } = parseMonthStr(targetMonth);
  const s = new Date(item.startDate + 'T00:00:00');
  const startYear = s.getFullYear();
  const startMonth = s.getMonth();
  const elapsed = (ty - startYear) * 12 + (tm - startMonth);
  return Math.max(0, elapsed);
}

// Given a prepaid item and a target YYYY-MM, return how much to recognize this month.
// Returns 0 if item is not active in the target month.
function computePrepaidMonthAmount(item, targetMonth) {
  const totalMonths = monthsBetweenInclusive(item.startDate, item.endDate);
  if (totalMonths <= 0) return 0;

  // Use openingBalance if provided (mid-life import), else totalAmount
  const amortizableAmount = Number(item.openingBalance ?? item.totalAmount) || 0;
  if (amortizableAmount <= 0) return 0;

  // Is target month within the amortization window?
  const { year: ty, monthIndex: tm } = parseMonthStr(targetMonth);
  const s = new Date(item.startDate + 'T00:00:00');
  const e = new Date(item.endDate + 'T00:00:00');
  const targetFirst = new Date(ty, tm, 1);
  const targetLast = new Date(ty, tm + 1, 0);
  if (targetLast < new Date(s.getFullYear(), s.getMonth(), 1)) return 0;
  if (targetFirst > new Date(e.getFullYear(), e.getMonth() + 1, 0)) return 0;

  // How many months from opening balance are still to be recognized?
  // If openingBalance is provided, we treat startDate as the "reference start" for the
  // remaining balance — caller should set startDate appropriately.
  const recognizedBefore = monthsRecognizedBefore(item, targetMonth);
  const remainingMonths = totalMonths - recognizedBefore;
  if (remainingMonths <= 0) return 0;

  // Straight-line: amortizableAmount / totalMonths, but last month absorbs rounding
  const perMonth = Math.round((amortizableAmount / totalMonths) * 100) / 100;

  if (remainingMonths === 1) {
    // Last month — use the exact remainder
    const alreadyRecognized = perMonth * recognizedBefore;
    const remainder = Math.round((amortizableAmount - alreadyRecognized) * 100) / 100;
    return remainder;
  }
  return perMonth;
}

function computePrepaidPreview(clientData, targetMonth) {
  const lines = [];
  let total = 0;
  for (const item of clientData.items || []) {
    const amount = computePrepaidMonthAmount(item, targetMonth);
    if (amount > 0) {
      lines.push({
        itemId: item.id,
        vendor: item.vendor,
        description: item.description || '',
        expenseAccountId: item.expenseAccountId,
        expenseAccountName: item.expenseAccountName,
        startDate: item.startDate,
        endDate: item.endDate,
        amount,
      });
      total += amount;
    }
  }
  return { lines, totalAmount: Math.round(total * 100) / 100 };
}

// ---- CRUD ----

// Get prepaid expenses state for a client
app.get('/api/admin/clients/:clientId/prepaid-expenses', requireAdmin, (req, res) => {
  res.json(getClientPrepaid(req.params.clientId));
});

// Save prepaid module settings (threshold, etc.)
app.put('/api/admin/clients/:clientId/prepaid-expenses/settings', requireAdmin, (req, res) => {
  const c = getClientPrepaid(req.params.clientId);
  if (req.body.scanThreshold !== undefined) {
    c.scanThreshold = Math.max(0, Number(req.body.scanThreshold) || 500);
  }
  savePrepaidExpenses(prepaidData);
  res.json(c);
});

// Set the Prepaid Expenses GL account
app.put('/api/admin/clients/:clientId/prepaid-expenses/account', requireAdmin, (req, res) => {
  const { id, name } = req.body || {};
  if (!id || !name) return res.status(400).json({ error: 'id and name required' });
  const c = getClientPrepaid(req.params.clientId);
  c.prepaidAccount = { id: String(id), name: String(name) };
  savePrepaidExpenses(prepaidData);
  res.json(c);
});

// Create a prepaid item
app.post('/api/admin/clients/:clientId/prepaid-expenses/items', requireAdmin, (req, res) => {
  const {
    vendor, description, totalAmount, startDate, endDate,
    expenseAccountId, expenseAccountName, openingBalance,
    sourceTxnId, sourceTxnType,
  } = req.body || {};

  if (!vendor || !totalAmount || !startDate || !endDate || !expenseAccountId) {
    return res.status(400).json({ error: 'vendor, totalAmount, startDate, endDate, expenseAccountId required' });
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate) || !/^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
    return res.status(400).json({ error: 'startDate and endDate must be YYYY-MM-DD' });
  }
  if (new Date(endDate) < new Date(startDate)) {
    return res.status(400).json({ error: 'endDate must be on or after startDate' });
  }

  const c = getClientPrepaid(req.params.clientId);
  const item = {
    id: Date.now().toString() + Math.random().toString(36).slice(2, 7),
    vendor: String(vendor),
    description: description || '',
    totalAmount: Number(totalAmount),
    openingBalance: openingBalance != null ? Number(openingBalance) : Number(totalAmount),
    startDate,
    endDate,
    expenseAccountId: String(expenseAccountId),
    expenseAccountName: expenseAccountName || '',
    sourceTxnId: sourceTxnId || null,
    sourceTxnType: sourceTxnType || null,
    createdAt: new Date().toISOString(),
    completedAt: null,
  };
  c.items.push(item);
  savePrepaidExpenses(prepaidData);
  res.json(item);
});

// Update a prepaid item
app.put('/api/admin/clients/:clientId/prepaid-expenses/items/:id', requireAdmin, (req, res) => {
  const c = getClientPrepaid(req.params.clientId);
  const item = c.items.find(i => i.id === req.params.id);
  if (!item) return res.status(404).json({ error: 'item not found' });

  const allowed = ['vendor', 'description', 'totalAmount', 'openingBalance', 'startDate', 'endDate', 'expenseAccountId', 'expenseAccountName'];
  for (const key of allowed) {
    if (req.body[key] !== undefined) {
      item[key] = ['totalAmount', 'openingBalance'].includes(key) ? Number(req.body[key]) : req.body[key];
    }
  }
  savePrepaidExpenses(prepaidData);
  res.json(item);
});

// Delete a prepaid item
app.delete('/api/admin/clients/:clientId/prepaid-expenses/items/:id', requireAdmin, (req, res) => {
  const c = getClientPrepaid(req.params.clientId);
  const before = c.items.length;
  c.items = c.items.filter(i => i.id !== req.params.id);
  if (c.items.length === before) return res.status(404).json({ error: 'item not found' });
  savePrepaidExpenses(prepaidData);
  res.json({ success: true });
});

// Preview amortization for the current close period
app.get('/api/admin/clients/:clientId/prepaid-expenses/preview-amortization', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    const clientData = getClientPrepaid(clientId);
    const fixedAssetClient = getClientAssets(clientId);

    // Reuse the same target-month determination as fixed assets so the two
    // modules advance in lockstep through the close workflow.
    const overrideMonth = req.query.month && /^\d{4}-\d{2}$/.test(req.query.month) ? req.query.month : null;
    let closeDate = null;
    if (qbo.isConnected(clientId)) {
      try { closeDate = await qbo.getBookCloseDate(clientId); } catch (e) { /* ignore */ }
    }
    const targetMonth = overrideMonth || determineAmortizationMonth(fixedAssetClient, closeDate).month;

    const alreadyRun = clientData.amortizationRuns.find(r => r.month === targetMonth);
    if (alreadyRun) {
      return res.json({ alreadyRun: true, runDetails: alreadyRun, month: targetMonth, closeDate });
    }

    if (!clientData.prepaidAccount) {
      return res.json({
        alreadyRun: false,
        month: targetMonth,
        closeDate,
        notConfigured: true,
        reason: 'Prepaid Expenses GL account not set',
        eligibleCount: 0,
        totalAmount: 0,
        lines: [],
      });
    }

    const { lines, totalAmount } = computePrepaidPreview(clientData, targetMonth);
    res.json({
      alreadyRun: false,
      month: targetMonth,
      closeDate,
      prepaidAccount: clientData.prepaidAccount,
      eligibleCount: lines.length,
      totalAmount,
      lines,
    });
  } catch (e) {
    console.error('prepaid preview error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Run amortization and post JE to QBO
app.post('/api/admin/clients/:clientId/prepaid-expenses/run-amortization', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    const clientData = getClientPrepaid(clientId);
    const fixedAssetClient = getClientAssets(clientId);

    if (!qbo.isConnected(clientId)) {
      return res.status(400).json({ error: 'QuickBooks is not connected for this client.' });
    }
    if (!clientData.prepaidAccount) {
      return res.status(400).json({ error: 'Prepaid Expenses GL account not set. Configure it in settings first.' });
    }

    const overrideMonth = req.body?.month && /^\d{4}-\d{2}$/.test(req.body.month) ? req.body.month : null;

    let closeDate = null;
    try { closeDate = await qbo.getBookCloseDate(clientId); } catch (e) { /* ignore */ }
    const targetMonth = overrideMonth || determineAmortizationMonth(fixedAssetClient, closeDate).month;

    // Guard: closed period
    if (closeDate) {
      const { year, monthIndex } = parseMonthStr(targetMonth);
      const lastDay = lastDayOfMonthDate(year, monthIndex);
      if (lastDay <= new Date(closeDate)) {
        return res.status(400).json({ error: `Target month ${targetMonth} falls within closed period (close date ${closeDate})` });
      }
    }

    const alreadyRun = clientData.amortizationRuns.find(r => r.month === targetMonth);
    if (alreadyRun) return res.status(400).json({ error: `Prepaid amortization already run for ${targetMonth}` });

    const { lines, totalAmount } = computePrepaidPreview(clientData, targetMonth);
    if (lines.length === 0) return res.status(400).json({ error: `No prepaid items to amortize in ${targetMonth}` });

    const { year, monthIndex } = parseMonthStr(targetMonth);
    const lastDayStr = lastDayOfMonthDate(year, monthIndex).toISOString().split('T')[0];

    // Build the JE: Dr each expense line, Cr Prepaid Expenses for the total
    const jeLines = [];
    for (const line of lines) {
      jeLines.push({
        accountId: line.expenseAccountId,
        accountName: line.expenseAccountName,
        description: `Prepaid amortization - ${line.vendor}${line.description ? ' - ' + line.description : ''}`,
        amount: line.amount,
        type: 'debit',
      });
    }
    jeLines.push({
      accountId: clientData.prepaidAccount.id,
      accountName: clientData.prepaidAccount.name,
      description: `Prepaid amortization - ${targetMonth}`,
      amount: totalAmount,
      type: 'credit',
    });

    const clientName = CLIENTS[clientId]?.name || clientId;
    const result = await qbo.createJournalEntry({
      date: lastDayStr,
      memo: `Prepaid expense amortization - ${targetMonth} - ${clientName}`,
      lines: jeLines,
    }, clientId);

    // Mark items complete if fully amortized after this run
    for (const line of lines) {
      const item = clientData.items.find(i => i.id === line.itemId);
      if (!item) continue;
      // If the target month is the last month in the schedule, mark complete
      const totalMonths = monthsBetweenInclusive(item.startDate, item.endDate);
      const recognizedBefore = monthsRecognizedBefore(item, targetMonth);
      if (recognizedBefore + 1 >= totalMonths) {
        item.completedAt = new Date().toISOString();
      }
    }

    const runRecord = {
      month: targetMonth,
      ranAt: new Date().toISOString(),
      journalEntryId: result.Id || null,
      totalAmount,
      itemCount: lines.length,
      items: lines,
    };
    clientData.amortizationRuns.push(runRecord);
    savePrepaidExpenses(prepaidData);

    res.json({ success: true, run: runRecord });
  } catch (e) {
    console.error('prepaid run-amortization error:', e);
    res.status(500).json({ error: `Failed to post journal entry: ${e.message}` });
  }
});

// ---- Excel import/export for existing prepaid balances ----

// Download a blank template
app.get('/api/admin/clients/:clientId/prepaid-expenses/export-template', requireAdmin, async (req, res) => {
  try {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Prepaid Expenses');

    ws.columns = [
      { header: 'Vendor',              key: 'vendor',             width: 25 },
      { header: 'Description',         key: 'description',        width: 35 },
      { header: 'Total Amount',        key: 'totalAmount',        width: 14 },
      { header: 'Opening Balance',     key: 'openingBalance',     width: 16 },
      { header: 'Start Date',          key: 'startDate',          width: 14 },
      { header: 'End Date',            key: 'endDate',            width: 14 },
      { header: 'Expense Account ID',  key: 'expenseAccountId',   width: 20 },
      { header: 'Expense Account Name',key: 'expenseAccountName', width: 30 },
    ];
    ws.getRow(1).font = { bold: true };

    // Example row to show format (comment explains it's safe to delete)
    ws.addRow({
      vendor: 'Acme Insurance',
      description: '12-month policy',
      totalAmount: 1200,
      openingBalance: 800,
      startDate: '2026-01-01',
      endDate: '2026-12-31',
      expenseAccountId: '',
      expenseAccountName: 'Insurance Expense',
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="prepaid-expenses-template.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Export current prepaid schedule + run Claude review for the current close period
app.get('/api/admin/clients/:clientId/prepaid-expenses/export-excel', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    const client = CLIENTS[clientId];
    const clientName = client?.name || clientId;
    const clientData = getClientPrepaid(clientId);

    // Determine as-of month: query param > latest run month > current month
    let asOfMonth = req.query.month;
    if (!asOfMonth) {
      const runs = (clientData.amortizationRuns || []).slice().sort((a, b) => (b.month || '').localeCompare(a.month || ''));
      asOfMonth = runs[0]?.month || formatMonthStr(new Date().getFullYear(), new Date().getMonth());
    }

    const { snapshots, totals } = prepaidReviewService.buildItemSnapshots(clientData, asOfMonth);

    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    wb.creator = 'jack does';
    wb.created = new Date();

    // ---- Sheet 1: Schedule ----
    const ws = wb.addWorksheet('Prepaid Schedule', { properties: { tabColor: { argb: 'FF3b82f6' } } });
    ws.mergeCells(1, 1, 1, 11);
    ws.getCell(1, 1).value = `${clientName} — Prepaid Expenses Schedule (as of ${asOfMonth})`;
    ws.getCell(1, 1).font = { name: 'Calibri', size: 14, bold: true };
    ws.getRow(1).height = 22;

    const headers = [
      'Vendor', 'Description', 'Expense Account', 'Start Date', 'End Date',
      'Total Months', 'Opening Balance', 'Monthly Amount',
      'Months Through', 'Recognized To Date', 'Closing Balance',
    ];
    const headerRow = 3;
    headers.forEach((h, i) => {
      const cell = ws.getCell(headerRow, i + 1);
      cell.value = h;
      cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1a1a2e' } };
      cell.alignment = { vertical: 'middle' };
    });
    ws.getRow(headerRow).height = 20;

    const widths = [22, 28, 28, 12, 12, 12, 16, 14, 14, 18, 16];
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    const currencyFormat = '"$"#,##0.00;[Red]-"$"#,##0.00';
    snapshots.forEach((s, idx) => {
      const r = headerRow + 1 + idx;
      ws.getCell(r, 1).value = s.vendor;
      ws.getCell(r, 2).value = s.description;
      ws.getCell(r, 3).value = s.expenseAccountName;
      ws.getCell(r, 4).value = s.startDate;
      ws.getCell(r, 5).value = s.endDate;
      ws.getCell(r, 6).value = s.totalMonths;
      ws.getCell(r, 7).value = s.openingBalance;
      ws.getCell(r, 8).value = s.monthlyAmount;
      ws.getCell(r, 9).value = s.monthsThrough;
      ws.getCell(r, 10).value = s.actualRecognized;
      ws.getCell(r, 11).value = s.closingBalance;
      [7, 8, 10, 11].forEach(col => { ws.getCell(r, col).numFmt = currencyFormat; });
    });

    // Totals row
    const totalsRow = headerRow + 1 + snapshots.length;
    ws.getCell(totalsRow, 1).value = 'TOTALS';
    ws.getCell(totalsRow, 1).font = { bold: true };
    ws.getCell(totalsRow, 7).value = totals.opening;
    ws.getCell(totalsRow, 10).value = totals.recognized;
    ws.getCell(totalsRow, 11).value = totals.closing;
    [7, 10, 11].forEach(col => {
      ws.getCell(totalsRow, col).numFmt = currencyFormat;
      ws.getCell(totalsRow, col).font = { bold: true };
      ws.getCell(totalsRow, col).border = { top: { style: 'thin' } };
    });

    // ---- Run Claude review (advisory — never blocks download) ----
    let review = null;
    try {
      review = await prepaidReviewService.reviewPrepaid(clientId, clientData, asOfMonth);
      excelService.addReviewSheet(wb, review, 'Prepaid Expenses Schedule', 'Item');
      clientData.lastReview = review;
      savePrepaidExpenses(prepaidData);
    } catch (e) {
      console.error('[prepaid-export] review failed:', e.message);
      review = {
        status: 'error',
        summary: `Review could not be completed: ${e.message}`,
        findings: [],
        generatedAt: new Date().toISOString(),
        asOfMonth,
        reviewError: e.message,
      };
      clientData.lastReview = review;
      savePrepaidExpenses(prepaidData);
    }

    const fileName = `${clientName} - Prepaid Expenses.xlsx`.replace(/[<>:"/\\|?*]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('X-Review-Status', review?.status || 'unknown');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error('[prepaid-export] error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Get the most recent Claude review result for prepaid
app.get('/api/admin/clients/:clientId/prepaid-expenses/last-review', requireAdmin, (req, res) => {
  const c = getClientPrepaid(req.params.clientId);
  res.json({ review: c.lastReview || null });
});

// Import existing prepaid balances from Excel (replaces the item list)
app.post('/api/admin/clients/:clientId/prepaid-expenses/import-excel', requireAdmin, excelUpload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const ws = wb.worksheets[0];
    if (!ws) return res.status(400).json({ error: 'Empty workbook' });

    const c = getClientPrepaid(req.params.clientId);
    const imported = [];
    const warnings = [];

    // Header row at row 1; data from row 2
    // Build a header-to-column-index map tolerant to case/whitespace
    const headerMap = {};
    ws.getRow(1).eachCell((cell, col) => {
      headerMap[String(cell.value || '').trim().toLowerCase()] = col;
    });

    const idx = {
      vendor: headerMap['vendor'],
      description: headerMap['description'],
      totalAmount: headerMap['total amount'],
      openingBalance: headerMap['opening balance'],
      startDate: headerMap['start date'],
      endDate: headerMap['end date'],
      expenseAccountId: headerMap['expense account id'],
      expenseAccountName: headerMap['expense account name'],
    };

    function cellValue(row, col) {
      if (!col) return null;
      const v = row.getCell(col).value;
      if (v && typeof v === 'object' && 'result' in v) return v.result;
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      return v;
    }
    function toDateStr(v) {
      if (!v) return null;
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      const s = String(v).trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      const d = new Date(s);
      if (!isNaN(d)) return d.toISOString().slice(0, 10);
      return null;
    }

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const vendor = cellValue(row, idx.vendor);
      if (!vendor) continue; // skip blank rows
      const totalAmount = Number(cellValue(row, idx.totalAmount));
      const openingBalance = idx.openingBalance ? Number(cellValue(row, idx.openingBalance)) : totalAmount;
      const startDate = toDateStr(cellValue(row, idx.startDate));
      const endDate = toDateStr(cellValue(row, idx.endDate));
      const expenseAccountId = cellValue(row, idx.expenseAccountId);
      const expenseAccountName = cellValue(row, idx.expenseAccountName);

      if (!totalAmount || !startDate || !endDate) {
        warnings.push(`Row ${r}: missing required fields, skipped`);
        continue;
      }
      imported.push({
        id: Date.now().toString() + Math.random().toString(36).slice(2, 7),
        vendor: String(vendor),
        description: String(cellValue(row, idx.description) || ''),
        totalAmount,
        openingBalance: isNaN(openingBalance) ? totalAmount : openingBalance,
        startDate,
        endDate,
        expenseAccountId: expenseAccountId ? String(expenseAccountId) : '',
        expenseAccountName: expenseAccountName ? String(expenseAccountName) : '',
        sourceTxnId: null,
        sourceTxnType: null,
        createdAt: new Date().toISOString(),
        completedAt: null,
      });
    }

    c.items = imported;
    savePrepaidExpenses(prepaidData);

    res.json({ success: true, itemCount: imported.length, warnings });
  } catch (e) {
    console.error('prepaid import error:', e);
    res.status(500).json({ error: 'Failed to import Excel file: ' + e.message });
  }
});

// ========================================
// PREPAID EXPENSE SCANNER (Part B)
// ========================================
// Scans recent expense transactions in QBO, downloads attached invoices,
// and uses Claude to detect prepaid components (service periods that extend
// beyond the close month). Results are returned for user review — nothing
// is posted automatically.

const PREPAID_SCAN_PROMPT = `You are a senior accountant reviewing an invoice/receipt to determine if any portion represents a PREPAID EXPENSE — meaning the payment covers a service period that extends beyond the current accounting period.

Current close month: {CLOSE_MONTH}
Period end date: {PERIOD_END}

Transaction details:
- Vendor: {VENDOR}
- Date: {TXN_DATE}
- Total amount: ${'{AMOUNT}'}
- Description/memo: {MEMO}

Review the attached document and determine:
1. Does this invoice cover a specific service period? If so, what are the start and end dates?
2. Does any portion of this invoice cover services AFTER {PERIOD_END}?
3. If yes, how much should be prepaid (deferred) vs. expensed in the current period?

IMPORTANT GUIDELINES:
- Monthly subscriptions where the service period = the billing period are NOT prepaids (e.g. a March invoice for March service)
- Only flag items where the service period clearly extends beyond {PERIOD_END}
- Insurance policies, annual licenses, multi-month contracts, and retainers are common prepaids
- If the document is unclear or you cannot determine the service period, say so — do NOT guess

Respond with ONLY a JSON object (no markdown, no explanation):
{
  "isPrepaid": true/false,
  "confidence": "high" | "medium" | "low",
  "reasoning": "brief explanation",
  "servicePeriodStart": "YYYY-MM-DD or null if unknown",
  "servicePeriodEnd": "YYYY-MM-DD or null if unknown",
  "totalAmount": number,
  "prepaidAmount": number or null,
  "currentPeriodAmount": number or null,
  "description": "what this invoice is for"
}`;

// Map QBO transaction types (from findPurchases/findBills/etc.) to the
// Attachable EntityRef.type values the QBO API expects.
const TXN_TYPE_MAP = {
  Purchase: 'Purchase',
  Bill: 'Bill',
  JournalEntry: 'JournalEntry',
  Invoice: 'Invoice',
  Expense: 'Purchase',  // QBO UI calls them "Expenses" but API type is Purchase
};

// POST /api/admin/clients/:clientId/prepaid-expenses/scan
// Body: { month?, threshold?, maxTransactions?, expenseAccountIds? }
app.post('/api/admin/clients/:clientId/prepaid-expenses/scan', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    if (!qbo.isConnected(clientId)) {
      return res.status(400).json({ error: 'QBO not connected' });
    }

    // Determine the close period
    const fixedAssetClient = getClientAssets(clientId);
    let closeDate = null;
    try { closeDate = await qbo.getBookCloseDate(clientId); } catch (_) {}
    const targetMonth = (req.body?.month && /^\d{4}-\d{2}$/.test(req.body.month))
      ? req.body.month
      : determineAmortizationMonth(fixedAssetClient, closeDate).month;

    const { year, monthIndex } = parseMonthStr(targetMonth);
    const periodStart = `${year}-${String(monthIndex + 1).padStart(2, '0')}-01`;
    const periodEnd = lastDayOfMonthDate(year, monthIndex).toISOString().split('T')[0];

    const threshold = Number(req.body?.threshold) || 500;
    const maxTxns = Math.min(Number(req.body?.maxTransactions) || 30, 50);

    console.log(`[prepaid-scan] scanning ${targetMonth} for client ${clientId}, threshold $${threshold}, max ${maxTxns}`);

    // Pull expense transactions for the period
    const [billsRes, purchasesRes] = await Promise.all([
      qbo.getBills(periodStart, periodEnd, 200, clientId).catch(() => null),
      qbo.getExpenseTransactions(periodStart, periodEnd, 200, clientId).catch(() => null),
    ]);

    const bills = (billsRes?.QueryResponse?.Bill || []).map(b => ({
      id: b.Id, type: 'Bill', date: b.TxnDate, amount: Number(b.TotalAmt || 0),
      vendor: b.VendorRef?.name || 'Unknown', memo: b.PrivateNote || b.Memo || '',
      accountName: b.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.name || '',
      accountId: b.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.value || '',
    }));
    const purchases = (purchasesRes?.QueryResponse?.Purchase || []).map(p => ({
      id: p.Id, type: 'Purchase', date: p.TxnDate, amount: Number(p.TotalAmt || 0),
      vendor: p.EntityRef?.name || 'Unknown', memo: p.PrivateNote || p.Memo || '',
      accountName: p.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.name || '',
      accountId: p.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.value || '',
    }));

    let candidates = [...bills, ...purchases]
      .filter(t => t.amount >= threshold)
      .sort((a, b) => b.amount - a.amount)
      .slice(0, maxTxns);

    // Optional: filter to specific expense accounts
    if (req.body?.expenseAccountIds?.length) {
      const allowedIds = new Set(req.body.expenseAccountIds.map(String));
      candidates = candidates.filter(t => allowedIds.has(t.accountId));
    }

    console.log(`[prepaid-scan] found ${candidates.length} candidate transactions above $${threshold}`);

    if (candidates.length === 0) {
      // Record that this month was scanned (even with no results) so calendar shows complete
      const ppClient = getClientPrepaid(clientId);
      if (!ppClient.scannedMonths.includes(targetMonth)) {
        ppClient.scannedMonths.push(targetMonth);
        savePrepaidExpenses(prepaidData);
      }
      return res.json({
        month: targetMonth,
        periodStart,
        periodEnd,
        threshold,
        candidates: [],
        results: [],
        summary: 'No expense transactions above the threshold for this period.',
      });
    }

    // For each candidate, check for attachments and review with Claude
    const results = [];
    for (const txn of candidates) {
      const result = {
        txn,
        hasAttachment: false,
        attachment: null,
        claudeReview: null,
        error: null,
      };

      try {
        // Check for attachments
        const attachments = await qbo.getTransactionAttachments(
          txn.id, TXN_TYPE_MAP[txn.type] || txn.type, clientId
        ).catch(() => []);

        if (attachments.length > 0) {
          result.hasAttachment = true;
          result.attachment = { id: attachments[0].id, fileName: attachments[0].fileName, contentType: attachments[0].contentType };

          // Download the attachment
          const file = await qbo.downloadAttachment(attachments[0].id, clientId);

          // Determine if Claude can read this file type
          const supportedTypes = ['application/pdf', 'image/png', 'image/jpeg', 'image/gif', 'image/webp'];
          const canReview = supportedTypes.some(t => file.contentType.includes(t));

          if (canReview) {
            // Build the prompt
            const prompt = PREPAID_SCAN_PROMPT
              .replace('{CLOSE_MONTH}', targetMonth)
              .replace(/{PERIOD_END}/g, periodEnd)
              .replace('{VENDOR}', txn.vendor)
              .replace('{TXN_DATE}', txn.date)
              .replace('{AMOUNT}', txn.amount.toFixed(2))
              .replace('{MEMO}', txn.memo || '(none)');

            // Determine the media type for Claude
            let mediaType = file.contentType;
            if (mediaType.includes('pdf')) mediaType = 'application/pdf';
            else if (mediaType.includes('png')) mediaType = 'image/png';
            else if (mediaType.includes('jpeg') || mediaType.includes('jpg')) mediaType = 'image/jpeg';
            else if (mediaType.includes('gif')) mediaType = 'image/gif';
            else if (mediaType.includes('webp')) mediaType = 'image/webp';

            const response = await anthropic.messages.create({
              model: 'claude-sonnet-4-20250514',
              max_tokens: 1024,
              messages: [{
                role: 'user',
                content: [
                  {
                    type: mediaType === 'application/pdf' ? 'document' : 'image',
                    source: {
                      type: 'base64',
                      media_type: mediaType,
                      data: file.buffer.toString('base64'),
                    },
                  },
                  { type: 'text', text: prompt },
                ],
              }],
            });

            const rawText = response.content[0]?.text || '';
            try {
              const jsonMatch = rawText.match(/\{[\s\S]*\}/);
              result.claudeReview = JSON.parse(jsonMatch ? jsonMatch[0] : rawText);
            } catch (_) {
              result.claudeReview = { isPrepaid: false, confidence: 'low', reasoning: rawText.slice(0, 500), error: 'parse_failed' };
            }
          } else {
            result.claudeReview = { isPrepaid: false, confidence: 'low', reasoning: `Unsupported file type: ${file.contentType}. Manual review needed.` };
          }
        } else {
          // No attachment — do a text-only review based on description + amount
          const prompt = PREPAID_SCAN_PROMPT
            .replace('{CLOSE_MONTH}', targetMonth)
            .replace(/{PERIOD_END}/g, periodEnd)
            .replace('{VENDOR}', txn.vendor)
            .replace('{TXN_DATE}', txn.date)
            .replace('{AMOUNT}', txn.amount.toFixed(2))
            .replace('{MEMO}', txn.memo || '(none)');

          const response = await anthropic.messages.create({
            model: 'claude-sonnet-4-20250514',
            max_tokens: 1024,
            messages: [{
              role: 'user',
              content: prompt + '\n\nNOTE: No invoice/receipt document is available. Base your assessment only on the vendor name, amount, date, and memo above. If you cannot determine the service period without a document, set isPrepaid to false with reasoning explaining that manual review is recommended.',
            }],
          });

          const rawText = response.content[0]?.text || '';
          try {
            const jsonMatch = rawText.match(/\{[\s\S]*\}/);
            result.claudeReview = JSON.parse(jsonMatch ? jsonMatch[0] : rawText);
          } catch (_) {
            result.claudeReview = { isPrepaid: false, confidence: 'low', reasoning: rawText.slice(0, 500), error: 'parse_failed' };
          }
        }
      } catch (e) {
        console.error(`[prepaid-scan] error scanning txn ${txn.id}:`, e.message);
        result.error = e.message;
      }

      results.push(result);
    }

    const flagged = results.filter(r => r.claudeReview?.isPrepaid === true);

    // Record that this month was scanned so calendar shows complete
    const ppClient = getClientPrepaid(clientId);
    if (!ppClient.scannedMonths.includes(targetMonth)) {
      ppClient.scannedMonths.push(targetMonth);
      savePrepaidExpenses(prepaidData);
    }

    res.json({
      month: targetMonth,
      periodStart,
      periodEnd,
      threshold,
      candidateCount: candidates.length,
      results,
      flaggedCount: flagged.length,
      summary: flagged.length > 0
        ? `Found ${flagged.length} potential prepaid expense${flagged.length !== 1 ? 's' : ''} out of ${candidates.length} transactions scanned.`
        : `No prepaid expenses detected among ${candidates.length} transactions scanned.`,
    });
  } catch (e) {
    console.error('[prepaid-scan] error:', e);
    res.status(500).json({ error: e.message });
  }
});

// POST /api/admin/clients/:clientId/prepaid-expenses/scan/accept
// Takes a scan result and adds it as a prepaid item to the schedule
app.post('/api/admin/clients/:clientId/prepaid-expenses/scan/accept', requireAdmin, (req, res) => {
  try {
    const { vendor, description, totalAmount, prepaidAmount, startDate, endDate,
            expenseAccountId, expenseAccountName, sourceTxnId, sourceTxnType } = req.body || {};

    if (!vendor || !totalAmount || !startDate || !endDate || !expenseAccountId) {
      return res.status(400).json({ error: 'vendor, totalAmount, startDate, endDate, expenseAccountId required' });
    }

    const c = getClientPrepaid(req.params.clientId);
    const item = {
      id: Date.now().toString() + Math.random().toString(36).slice(2, 7),
      vendor: String(vendor),
      description: description || '',
      totalAmount: Number(totalAmount),
      openingBalance: Number(prepaidAmount || totalAmount),
      startDate,
      endDate,
      expenseAccountId: String(expenseAccountId),
      expenseAccountName: expenseAccountName || '',
      sourceTxnId: sourceTxnId || null,
      sourceTxnType: sourceTxnType || null,
      createdAt: new Date().toISOString(),
      completedAt: null,
    };
    c.items.push(item);
    savePrepaidExpenses(prepaidData);

    res.json({ success: true, item });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ========================================
// ACCRUED LIABILITIES — per-client, file-backed
// ========================================
const ACCRUED_LIABILITIES_FILE = dataPath('accrued-liabilities.json');

function loadAccruedLiabilities() {
  try {
    if (fs.existsSync(ACCRUED_LIABILITIES_FILE)) {
      return JSON.parse(fs.readFileSync(ACCRUED_LIABILITIES_FILE, 'utf-8'));
    }
  } catch (e) { console.error('Failed to load accrued-liabilities.json:', e.message); }
  return {};
}
function saveAccruedLiabilities(data) {
  fs.writeFileSync(ACCRUED_LIABILITIES_FILE, JSON.stringify(data, null, 2), 'utf-8');
}
let accruedLiabData = loadAccruedLiabilities();

function getClientAccruedLiab(clientId) {
  if (!accruedLiabData[clientId]) {
    accruedLiabData[clientId] = {
      accruedLiabilitiesAccount: null,
      materialityThreshold: 10,
      excludedAccountIds: [],
      analysisRuns: [],
    };
  }
  const c = accruedLiabData[clientId];
  if (!c.analysisRuns) c.analysisRuns = [];
  if (c.materialityThreshold === undefined) c.materialityThreshold = 10;
  if (!c.excludedAccountIds) c.excludedAccountIds = [];
  return c;
}

// ---- CRUD / settings ----

app.get('/api/admin/clients/:clientId/accrued-liabilities', requireAdmin, (req, res) => {
  res.json(getClientAccruedLiab(req.params.clientId));
});

app.put('/api/admin/clients/:clientId/accrued-liabilities/settings', requireAdmin, (req, res) => {
  const c = getClientAccruedLiab(req.params.clientId);
  if (req.body.materialityThreshold !== undefined) c.materialityThreshold = Math.max(0, Number(req.body.materialityThreshold) || 10);
  if (Array.isArray(req.body.excludedAccountIds)) c.excludedAccountIds = req.body.excludedAccountIds.map(String);
  saveAccruedLiabilities(accruedLiabData);
  res.json(c);
});

app.put('/api/admin/clients/:clientId/accrued-liabilities/account', requireAdmin, (req, res) => {
  const { id, name } = req.body || {};
  if (!id || !name) return res.status(400).json({ error: 'id and name required' });
  const c = getClientAccruedLiab(req.params.clientId);
  c.accruedLiabilitiesAccount = { id: String(id), name: String(name) };
  saveAccruedLiabilities(accruedLiabData);
  res.json(c);
});

// ---- P&L monthly report parser ----

/**
 * Parse the QBO P&L monthly report into per-account monthly totals.
 * Returns { months: ['2024-10', ...], accounts: [{ accountId, accountName, monthlyTotals: { 'YYYY-MM': amt } }] }
 */
function parsePnlMonthlyReport(report) {
  // The Columns header gives us the month labels
  const columns = report?.Columns?.Column || [];
  // First column is the account name, rest are months + total
  const monthLabels = [];
  for (let i = 1; i < columns.length; i++) {
    const colTitle = columns[i].ColTitle || '';
    // Month columns look like "Jan 2025", "Feb 2025", etc. or sometimes "YYYY-MM"
    // The last column is often "Total" — skip it
    if (colTitle.toLowerCase() === 'total') continue;
    // Parse month name to YYYY-MM
    const d = new Date(colTitle + ' 1');
    if (!isNaN(d)) {
      monthLabels.push({
        idx: i,
        month: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
      });
    }
  }

  const accounts = [];

  function walkRows(rows) {
    if (!rows) return;
    for (const row of rows) {
      if (row.type === 'Section' && row.Rows?.Row) {
        walkRows(row.Rows.Row);
      } else if (row.type === 'Data' && row.ColData) {
        // ColData[0] = account name, rest = monthly values
        const nameCol = row.ColData[0] || {};
        const accountId = nameCol.id || '';
        const accountName = nameCol.value || '';
        if (!accountId) continue;
        const monthlyTotals = {};
        for (const ml of monthLabels) {
          const val = parseFloat(row.ColData[ml.idx]?.value) || 0;
          if (val !== 0) monthlyTotals[ml.month] = val;
        }
        accounts.push({ accountId, accountName, monthlyTotals });
      }
    }
  }

  walkRows(report?.Rows?.Row || []);
  return { months: monthLabels.map(m => m.month), accounts };
}

// ---- Analysis endpoint ----

app.post('/api/admin/clients/:clientId/accrued-liabilities/analyze', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    if (!qbo.isConnected(clientId)) {
      return res.status(400).json({ error: 'QBO not connected' });
    }
    const c = getClientAccruedLiab(clientId);
    if (!c.accruedLiabilitiesAccount) {
      return res.status(400).json({ error: 'Accrued Liabilities GL account not set. Configure it in settings.' });
    }

    // Determine close month
    const fixedAssetClient = getClientAssets(clientId);
    let closeDate = null;
    try { closeDate = await qbo.getBookCloseDate(clientId); } catch (_) {}
    const targetMonth = (req.body?.month && /^\d{4}-\d{2}$/.test(req.body.month))
      ? req.body.month
      : determineAmortizationMonth(fixedAssetClient, closeDate).month;

    const { year, monthIndex } = parseMonthStr(targetMonth);
    const periodEnd = lastDayOfMonthDate(year, monthIndex).toISOString().split('T')[0];

    // ---- Part A: historical pattern analysis ----
    // 18 months back from start of close month
    const lookbackStart = new Date(year, monthIndex - 18, 1);
    const lookbackStartStr = lookbackStart.toISOString().split('T')[0];

    console.log(`[accrued-liab] Part A: P&L monthly ${lookbackStartStr} → ${periodEnd} for ${clientId}`);
    const pnlReport = await qbo.getProfitAndLossMonthly(lookbackStartStr, periodEnd, clientId);
    const parsed = parsePnlMonthlyReport(pnlReport);

    const excludeSet = new Set(c.excludedAccountIds || []);
    const partAAccounts = [];
    const priorMonths = parsed.months.filter(m => m < targetMonth);
    // Prior month (immediately before targetMonth) — used for "recent recurring" signal
    const priorMonthKey = (() => {
      const d = new Date(year, monthIndex - 1, 1);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    })();

    // Diagnostics: account for every P&L row we considered, so the UI can explain why nothing was flagged
    const diagnostics = {
      totalAccounts: parsed.accounts.length,
      excluded: 0,
      negligibleHistory: 0,
      zeroCurrentBelowFrequency: 0, // current=0 but frequency too low to flag
      withinTolerance: 0, // has history + current entry but inside materiality band
      evaluated: 0,
      nearMiss: [], // accounts with recent prior-month activity that didn't quite flag (for user visibility)
    };

    for (const acct of parsed.accounts) {
      if (excludeSet.has(acct.accountId)) { diagnostics.excluded++; continue; }
      let sumPrior = 0, countPrior = 0;
      for (const m of priorMonths) {
        const val = acct.monthlyTotals[m] || 0;
        if (val !== 0) { sumPrior += val; countPrior++; }
      }
      const average = priorMonths.length > 0 ? sumPrior / priorMonths.length : 0;
      const frequency = countPrior; // how many of the prior months had activity
      const currentMonth = acct.monthlyTotals[targetMonth] || 0;
      const appearedLastMonth = !!acct.monthlyTotals[priorMonthKey];
      const priorMonthAmount = acct.monthlyTotals[priorMonthKey] || 0;

      // Skip accounts with negligible history
      if (average < 1 && currentMonth < 1 && !appearedLastMonth) {
        diagnostics.negligibleHistory++;
        continue;
      }
      diagnostics.evaluated++;

      const gap = Math.round((average - currentMonth) * 100) / 100;
      const thresholdPct = c.materialityThreshold / 100;
      let flagged = false;
      let reason = '';

      // MISSING: current month has nothing, but the account is clearly recurring
      // Tiered detection so we catch both established patterns (3+ months)
      // AND newer subscriptions (appeared last month + at least once before, or just last month with material amount)
      if (currentMonth === 0) {
        if (average >= 1 && frequency >= 3) {
          flagged = true;
          reason = 'missing';
        } else if (appearedLastMonth && frequency >= 2) {
          flagged = true;
          reason = 'missing_recurring';
        } else if (appearedLastMonth && priorMonthAmount >= 1) {
          // New subscription: showed up last month with material spend, but not this month
          flagged = true;
          reason = 'missing_new_recurring';
        } else if (average >= 1 && frequency >= 2) {
          flagged = true;
          reason = 'missing';
        } else {
          diagnostics.zeroCurrentBelowFrequency++;
          if (appearedLastMonth) {
            diagnostics.nearMiss.push({
              accountId: acct.accountId,
              accountName: acct.accountName,
              priorMonthAmount: Math.round(priorMonthAmount * 100) / 100,
              frequency,
              note: 'appeared last month but frequency too low',
            });
          }
        }
      } else if (average > 0 && currentMonth < average * (1 - thresholdPct) && frequency >= 3) {
        flagged = true;
        reason = 'below_average';
      } else {
        diagnostics.withinTolerance++;
      }

      if (flagged) {
        // For new-recurring where 18-mo average is misleadingly low, use prior-month amount as suggested accrual
        let suggestedAccrual = gap;
        if (reason === 'missing_new_recurring' || reason === 'missing_recurring') {
          suggestedAccrual = Math.round(priorMonthAmount * 100) / 100;
        }
        partAAccounts.push({
          accountId: acct.accountId,
          accountName: acct.accountName,
          monthlyTotals: acct.monthlyTotals,
          average: Math.round(average * 100) / 100,
          frequency,
          currentMonth: Math.round(currentMonth * 100) / 100,
          priorMonthAmount: Math.round(priorMonthAmount * 100) / 100,
          appearedLastMonth,
          gap,
          flagged: true,
          reason,
          accrualAmount: suggestedAccrual,
          status: 'pending',
        });
      }
    }

    // Sort by accrual amount (absolute) descending so biggest items bubble up regardless of reason
    partAAccounts.sort((a, b) => Math.abs(b.accrualAmount) - Math.abs(a.accrualAmount));

    // ---- Part B: subsequent events ----
    const nextMonthFirst = new Date(year, monthIndex + 1, 1);
    const windowStart = nextMonthFirst.toISOString().split('T')[0];
    const todayStr = new Date().toISOString().split('T')[0];
    const windowEnd = todayStr;
    // If we're analyzing the current in-progress month, the "subsequent events" window is invalid
    // (windowStart is in the future). Skip the QBO fetch and note it in the diagnostics.
    const partBSkipped = windowStart > windowEnd;
    let partBNote = null;

    let subsequentTxns = [];
    if (partBSkipped) {
      partBNote = `Skipped — analyzing ${targetMonth} before it has closed (today is ${todayStr}). Subsequent-events window starts ${windowStart}.`;
      console.log(`[accrued-liab] Part B: ${partBNote}`);
    } else {
      console.log(`[accrued-liab] Part B: subsequent events ${windowStart} → ${windowEnd}`);

      const [billsRes, purchasesRes] = await Promise.all([
        qbo.getBills(windowStart, windowEnd, 500, clientId).catch(() => null),
        qbo.getExpenseTransactions(windowStart, windowEnd, 500, clientId).catch(() => null),
      ]);

      const bills = (billsRes?.QueryResponse?.Bill || []).map(b => ({
        txnId: b.Id, txnType: 'Bill', date: b.TxnDate,
        amount: Number(b.TotalAmt || 0),
        vendor: b.VendorRef?.name || 'Unknown',
        memo: b.PrivateNote || b.Memo || '',
        accountId: b.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.value || '',
        accountName: b.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.name || '',
      }));
      const purchases = (purchasesRes?.QueryResponse?.Purchase || []).map(p => ({
        txnId: p.Id, txnType: 'Purchase', date: p.TxnDate,
        amount: Number(p.TotalAmt || 0),
        vendor: p.EntityRef?.name || 'Unknown',
        memo: p.PrivateNote || p.Memo || '',
        accountId: p.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.value || '',
        accountName: p.Line?.[0]?.AccountBasedExpenseLineDetail?.AccountRef?.name || '',
      }));

      subsequentTxns = [...bills, ...purchases]
        .filter(t => !excludeSet.has(t.accountId) && t.amount > 0)
        .sort((a, b) => b.amount - a.amount)
        .map(t => ({
          ...t,
          flagged: true,
          accrualAmount: t.amount,
          status: 'pending',
          // Note overlap with Part A
          overlapWithPartA: partAAccounts.some(a => a.accountId === t.accountId),
        }));
    }

    // Store the run
    const run = {
      month: targetMonth,
      ranAt: new Date().toISOString(),
      partA: {
        lookbackMonths: priorMonths.length,
        priorMonth: priorMonthKey,
        accounts: partAAccounts,
        diagnostics,
      },
      partB: {
        windowStart,
        windowEnd,
        skipped: partBSkipped,
        note: partBNote,
        transactions: subsequentTxns,
      },
      accrualJE: null,
      reversalJE: null,
    };

    // Replace existing run for this month or push new one
    const existingIdx = c.analysisRuns.findIndex(r => r.month === targetMonth);
    if (existingIdx >= 0) {
      // Preserve JE data if already posted
      run.accrualJE = c.analysisRuns[existingIdx].accrualJE;
      run.reversalJE = c.analysisRuns[existingIdx].reversalJE;
      c.analysisRuns[existingIdx] = run;
    } else {
      c.analysisRuns.push(run);
    }
    saveAccruedLiabilities(accruedLiabData);

    res.json({
      month: targetMonth,
      periodEnd,
      partA: run.partA,
      partB: run.partB,
      alreadyPosted: !!run.accrualJE,
    });
  } catch (e) {
    console.error('[accrued-liab] analyze error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Update a flagged item (dismiss, change amount, etc.)
app.put('/api/admin/clients/:clientId/accrued-liabilities/runs/:month/items', requireAdmin, (req, res) => {
  const c = getClientAccruedLiab(req.params.clientId);
  const run = c.analysisRuns.find(r => r.month === req.params.month);
  if (!run) return res.status(404).json({ error: 'no analysis run for this month' });

  const { part, index, status, accrualAmount } = req.body;
  let item;
  if (part === 'A') {
    item = run.partA.accounts[index];
  } else if (part === 'B') {
    item = run.partB.transactions[index];
  }
  if (!item) return res.status(404).json({ error: 'item not found' });

  if (status !== undefined) item.status = status; // 'pending' | 'accrued' | 'dismissed'
  if (accrualAmount !== undefined) item.accrualAmount = Number(accrualAmount);

  saveAccruedLiabilities(accruedLiabData);
  res.json({ success: true, item });
});

// ---- Post accrual JE + auto-reversal ----

app.post('/api/admin/clients/:clientId/accrued-liabilities/post', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    if (!qbo.isConnected(clientId)) {
      return res.status(400).json({ error: 'QBO not connected' });
    }
    const c = getClientAccruedLiab(clientId);
    if (!c.accruedLiabilitiesAccount) {
      return res.status(400).json({ error: 'Accrued Liabilities GL account not set' });
    }

    const month = req.body?.month;
    if (!month) return res.status(400).json({ error: 'month required' });
    const run = c.analysisRuns.find(r => r.month === month);
    if (!run) return res.status(400).json({ error: `no analysis run for ${month}` });
    if (run.accrualJE) return res.status(400).json({ error: `accrual JE already posted for ${month}` });

    // Collect all non-dismissed items
    const lines = [];
    for (const a of (run.partA?.accounts || [])) {
      if (a.status === 'dismissed' || !a.accrualAmount || a.accrualAmount <= 0) continue;
      lines.push({
        source: 'pattern',
        accountId: a.accountId,
        accountName: a.accountName,
        description: `Accrued liabilities - ${a.reason === 'missing' ? 'missing' : 'below average'} - ${a.accountName}`,
        amount: Math.round(a.accrualAmount * 100) / 100,
      });
    }
    for (const t of (run.partB?.transactions || [])) {
      if (t.status === 'dismissed' || !t.accrualAmount || t.accrualAmount <= 0) continue;
      // Skip if Part A already covers this account (avoid double-count)
      if (t.overlapWithPartA && lines.some(l => l.accountId === t.accountId)) continue;
      lines.push({
        source: 'subsequent',
        accountId: t.accountId,
        accountName: t.accountName,
        description: `Accrued liabilities - ${t.vendor} (${t.date})`,
        amount: Math.round(t.accrualAmount * 100) / 100,
      });
    }

    if (lines.length === 0) {
      return res.status(400).json({ error: 'No accrual items to post (all dismissed or zero)' });
    }

    const totalAmount = Math.round(lines.reduce((s, l) => s + l.amount, 0) * 100) / 100;
    const { year, monthIndex } = parseMonthStr(month);
    const accrualDate = lastDayOfMonthDate(year, monthIndex).toISOString().split('T')[0];
    const reversalDate = new Date(year, monthIndex + 1, 1).toISOString().split('T')[0];
    const clientName = CLIENTS[clientId]?.name || clientId;

    // Build accrual JE: Dr expenses, Cr accrued liabilities
    const accrualJELines = [];
    for (const l of lines) {
      accrualJELines.push({ accountId: l.accountId, accountName: l.accountName, description: l.description, amount: l.amount, type: 'debit' });
    }
    accrualJELines.push({
      accountId: c.accruedLiabilitiesAccount.id,
      accountName: c.accruedLiabilitiesAccount.name,
      description: `Accrued liabilities - ${month}`,
      amount: totalAmount,
      type: 'credit',
    });

    const accrualResult = await qbo.createJournalEntry({
      date: accrualDate,
      memo: `Accrued liabilities - ${month} - ${clientName}`,
      lines: accrualJELines,
    }, clientId);

    // Build reversal JE: Cr expenses, Dr accrued liabilities (exact mirror)
    const reversalJELines = [];
    for (const l of lines) {
      reversalJELines.push({ accountId: l.accountId, accountName: l.accountName, description: `Reversal - ${l.description}`, amount: l.amount, type: 'credit' });
    }
    reversalJELines.push({
      accountId: c.accruedLiabilitiesAccount.id,
      accountName: c.accruedLiabilitiesAccount.name,
      description: `Reversal - accrued liabilities - ${month}`,
      amount: totalAmount,
      type: 'debit',
    });

    const reversalResult = await qbo.createJournalEntry({
      date: reversalDate,
      memo: `Reversal of accrued liabilities - ${month} - ${clientName}`,
      lines: reversalJELines,
    }, clientId);

    // Store results
    run.accrualJE = {
      journalEntryId: accrualResult.Id || null,
      date: accrualDate,
      totalAmount,
      lineCount: lines.length,
      lines,
    };
    run.reversalJE = {
      journalEntryId: reversalResult.Id || null,
      date: reversalDate,
      totalAmount,
    };

    // Mark all posted items as 'accrued'
    for (const a of (run.partA?.accounts || [])) { if (a.status === 'pending') a.status = 'accrued'; }
    for (const t of (run.partB?.transactions || [])) { if (t.status === 'pending') t.status = 'accrued'; }

    saveAccruedLiabilities(accruedLiabData);

    res.json({
      success: true,
      accrualJE: run.accrualJE,
      reversalJE: run.reversalJE,
    });
  } catch (e) {
    console.error('[accrued-liab] post error:', e);
    res.status(500).json({ error: `Failed to post: ${e.message}` });
  }
});

// DELETE an accrued liabilities analysis run for a specific month
app.delete('/api/admin/clients/:clientId/accrued-liabilities/runs/:month', requireAdmin, (req, res) => {
  const c = getClientAccruedLiab(req.params.clientId);
  const month = req.params.month;
  const before = c.analysisRuns.length;
  c.analysisRuns = c.analysisRuns.filter(r => r.month !== month);
  if (c.analysisRuns.length === before) return res.status(404).json({ error: 'No analysis run found for ' + month });
  saveAccruedLiabilities(accruedLiabData);
  res.json({ success: true, deleted: month });
});

// ========================================
// SHAREHOLDER-PAID INVOICES MODULE
// ========================================
const SHAREHOLDER_INVOICES_FILE = dataPath('shareholder-invoices.json');

function loadShareholderInvoices() {
  try {
    if (fs.existsSync(SHAREHOLDER_INVOICES_FILE)) {
      return JSON.parse(fs.readFileSync(SHAREHOLDER_INVOICES_FILE, 'utf-8'));
    }
  } catch (e) { console.error('Failed to load shareholder-invoices.json:', e.message); }
  return {};
}
function saveShareholderInvoices(data) {
  fs.writeFileSync(SHAREHOLDER_INVOICES_FILE, JSON.stringify(data, null, 2), 'utf-8');
}
let shareholderInvoiceData = loadShareholderInvoices();

function getClientShareholderInvoices(clientId) {
  if (!shareholderInvoiceData[clientId]) {
    shareholderInvoiceData[clientId] = {
      shareholderLoanAccount: null,
      invoices: [],
    };
  }
  const c = shareholderInvoiceData[clientId];
  if (!c.invoices) c.invoices = [];
  return c;
}

const SHAREHOLDER_INVOICE_PROMPT = `You are a senior accountant reviewing an invoice/receipt that was paid personally by a shareholder and needs to be recorded in the company's books.

The journal entry pattern is:
  Dr: Appropriate expense or asset account
  Cr: Shareholder Loan (liability)

Current close month: {CLOSE_MONTH}

The company's chart of accounts includes these expense/asset accounts:
{ACCOUNT_LIST}

Review the attached document carefully and determine:
1. What is the vendor name?
2. What is the invoice/receipt number? (look for "Invoice #", "Inv No.", "Receipt #", "Order #", "Reference #", etc.)
3. What is the invoice date?
4. What is the total amount (including taxes)?
5. What is this invoice for? (description)
6. Which GL account from the list above best fits each line item?
7. If there are multiple line items that should go to different accounts, list each separately.

CRITICAL RULE — BALANCED LINES:
The sum of all line amounts MUST equal totalAmount EXACTLY, down to the cent. Before responding, verify: add up every line amount and confirm it equals totalAmount. If it doesn't, you MUST add or adjust lines until it balances.

Common sources of imbalance:
- Tax (GST/HST/PST/VAT) not included as a line → add a "GST" or "HST" line
- Rounding → adjust the last line by the difference
- Shipping/handling fees omitted → add a line

Example: if totalAmount is 775.04 and subtotal items sum to 692.00, the remaining 83.04 is likely tax — add a line like {"description": "GST/HST", "amount": 83.04, "suggestedCategory": "taxes", ...}

IMPORTANT GUIDELINES:
- ALWAYS include tax as a separate line item so lines sum to the total
- If there are clearly distinct line items for different expense categories, break them out
- If the document is unclear, describe what you can see and flag for manual review
- Capital items (equipment, furniture, vehicles over $500) should be categorized as fixed assets
- For suggestedAccount, pick the best match from the chart of accounts list above. Use the exact account name.
- Be specific about the expense category — don't just say "expense"

Respond with ONLY a JSON object (no markdown, no explanation):
{
  "vendor": "vendor name",
  "invoiceNumber": "the invoice/receipt/order number from the document, or empty string if not found",
  "invoiceDate": "YYYY-MM-DD",
  "totalAmount": number,
  "description": "what this invoice is for",
  "currency": "CAD" or "USD",
  "confidence": "high" | "medium" | "low",
  "reasoning": "brief explanation of your categorization",
  "lines": [
    {
      "description": "line item description",
      "amount": number,
      "suggestedCategory": "expense category",
      "suggestedAccount": "exact account name from chart of accounts",
      "isCapital": false
    }
  ]
}`;

// Multer for shareholder invoice uploads (memory storage for AI processing)
const shareholderUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.pdf', '.jpg', '.jpeg', '.png', '.gif', '.webp'].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error(`File type ${ext} not supported. Use PDF or image files.`));
    }
  },
});

// GET settings + invoices
app.get('/api/admin/clients/:clientId/shareholder-invoices', requireAdmin, (req, res) => {
  const c = getClientShareholderInvoices(req.params.clientId);
  res.json(c);
});

// PUT settings (shareholder loan account)
app.put('/api/admin/clients/:clientId/shareholder-invoices/settings', requireAdmin, (req, res) => {
  const c = getClientShareholderInvoices(req.params.clientId);
  if (req.body.shareholderLoanAccount !== undefined) c.shareholderLoanAccount = req.body.shareholderLoanAccount;
  saveShareholderInvoices(shareholderInvoiceData);
  res.json({ ok: true });
});

// POST upload + AI analyze a single invoice
app.post('/api/admin/clients/:clientId/shareholder-invoices/upload', requireAdmin, shareholderUpload.single('invoice'), async (req, res) => {
  try {
    const clientId = req.params.clientId;
    const c = getClientShareholderInvoices(clientId);
    const file = req.file;
    if (!file) return res.status(400).json({ error: 'No file uploaded' });

    // Determine close month
    let closeDate = null;
    if (qbo.isConnected(clientId)) {
      try { closeDate = await qbo.getBookCloseDate(clientId); } catch (_) {}
    }
    let closeMonth;
    if (closeDate) {
      const cd = new Date(closeDate + 'T00:00:00');
      const nm = new Date(cd.getFullYear(), cd.getMonth() + 1, 1);
      closeMonth = formatMonthStr(nm.getFullYear(), nm.getMonth());
    } else {
      const now = new Date();
      const last = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      closeMonth = formatMonthStr(last.getFullYear(), last.getMonth());
    }

    console.log(`[shareholder-invoice] analyzing ${file.originalname} for client ${clientId}`);

    // Determine media type
    let mediaType = file.mimetype;
    if (mediaType.includes('pdf')) mediaType = 'application/pdf';
    else if (mediaType.includes('png')) mediaType = 'image/png';
    else if (mediaType.includes('jpeg') || mediaType.includes('jpg')) mediaType = 'image/jpeg';
    else if (mediaType.includes('gif')) mediaType = 'image/gif';
    else if (mediaType.includes('webp')) mediaType = 'image/webp';

    // Build chart of accounts list for the prompt
    let accountList = '(chart of accounts not available — suggest generic categories)';
    if (qbo.isConnected(clientId)) {
      try {
        const accounts = await qbo.getAccounts(clientId);
        const relevant = accounts.filter(a =>
          ['Expense', 'Other Expense', 'Cost of Goods Sold', 'Fixed Asset', 'Other Current Asset'].includes(a.AccountType)
        );
        accountList = relevant.map(a => `- ${a.AcctNum ? a.AcctNum + ' ' : ''}${a.Name} (${a.AccountType})`).join('\n');
      } catch (_) {}
    }

    const prompt = SHAREHOLDER_INVOICE_PROMPT
      .replace('{CLOSE_MONTH}', closeMonth)
      .replace('{ACCOUNT_LIST}', accountList);

    const response = await claudeWithRetry({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1500,
      messages: [{
        role: 'user',
        content: [
          {
            type: mediaType === 'application/pdf' ? 'document' : 'image',
            source: { type: 'base64', media_type: mediaType, data: file.buffer.toString('base64') },
          },
          { type: 'text', text: prompt },
        ],
      }],
    });

    const rawText = response.content[0]?.text || '';
    let analysis;
    try {
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      analysis = JSON.parse(jsonMatch ? jsonMatch[0] : rawText);
    } catch (_) {
      analysis = { vendor: 'Unknown', totalAmount: 0, description: rawText.slice(0, 300), confidence: 'low', reasoning: 'Could not parse AI response', lines: [] };
    }

    // Validate: lines must sum to totalAmount. If not, add a balancing line.
    if (analysis.lines && analysis.lines.length > 0 && analysis.totalAmount > 0) {
      const lineSum = Math.round(analysis.lines.reduce((s, l) => s + Number(l.amount || 0), 0) * 100) / 100;
      const total = Math.round(Number(analysis.totalAmount) * 100) / 100;
      const diff = Math.round((total - lineSum) * 100) / 100;
      if (Math.abs(diff) >= 0.01) {
        console.log(`[shareholder-invoice] lines sum to ${lineSum}, total is ${total}, adding balancing line of ${diff}`);
        analysis.lines.push({
          description: diff > 0 ? 'Tax (GST/HST/PST)' : 'Adjustment',
          amount: diff,
          suggestedCategory: diff > 0 ? 'taxes' : 'other',
          suggestedAccount: '',
          isCapital: false,
        });
      }
    }

    // Save the invoice record
    const invoice = {
      id: `shi-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      fileName: file.originalname,
      fileSize: file.size,
      fileType: file.mimetype,
      fileData: file.buffer.toString('base64'),
      uploadedAt: new Date().toISOString(),
      closeMonth,
      analysis,
      status: 'pending', // pending | posted | dismissed
      journalEntry: null,
    };

    c.invoices.push(invoice);
    saveShareholderInvoices(shareholderInvoiceData);

    res.json({ ok: true, invoice: { ...invoice, fileData: undefined } });
  } catch (e) {
    console.error('[shareholder-invoice] upload error:', e);
    const isOverloaded = e?.status === 529 || (e.message && e.message.includes('Overloaded'));
    const msg = isOverloaded
      ? 'AI service is temporarily busy. Please try again in a minute.'
      : e.message;
    res.status(isOverloaded ? 503 : 500).json({ error: msg });
  }
});

// POST /post — post a single invoice as a JE to QBO
app.post('/api/admin/clients/:clientId/shareholder-invoices/:invoiceId/post', requireAdmin, async (req, res) => {
  try {
    const clientId = req.params.clientId;
    const c = getClientShareholderInvoices(clientId);
    if (!c.shareholderLoanAccount) {
      return res.status(400).json({ error: 'Shareholder Loan GL account not configured' });
    }

    const invoice = c.invoices.find(i => i.id === req.params.invoiceId);
    if (!invoice) return res.status(404).json({ error: 'Invoice not found' });
    if (invoice.status === 'posted') return res.status(400).json({ error: 'Already posted' });

    // Ensure shareholder loan account has type info (may be missing if saved before this field existed)
    if (!c.shareholderLoanAccount.type && qbo.isConnected(clientId)) {
      try {
        const accts = await qbo.getAccounts(clientId);
        const allAccounts = accts?.QueryResponse?.Account || [];
        const match = allAccounts.find(a => String(a.Id) === String(c.shareholderLoanAccount.id));
        if (match) {
          c.shareholderLoanAccount.type = match.AccountType || '';
          saveShareholderInvoices(shareholderInvoiceData);
        }
      } catch (e) { /* non-critical */ }
    }

    // Build JE lines from request body (user may have edited)
    const jeLines = req.body.lines;
    if (!jeLines || jeLines.length === 0) return res.status(400).json({ error: 'No JE lines provided' });

    const jeDate = req.body.date || invoice.analysis.invoiceDate;
    const totalAmount = Math.round(jeLines.reduce((s, l) => s + Number(l.amount), 0) * 100) / 100;

    const clientName = CLIENTS[clientId]?.name || clientId;
    const memo = `Shareholder-paid invoice: ${invoice.analysis.vendor || invoice.fileName} — ${clientName}`;

    if (qbo.isConnected(clientId)) {
      // Look up full account list to detect tax liability accounts
      let allAccounts = [];
      try {
        const acctData = await qbo.getAccounts(clientId);
        allAccounts = acctData?.QueryResponse?.Account || [];
      } catch (e) { /* non-critical */ }

      const taxLiabilityIds = new Set(
        allAccounts
          .filter(a => {
            const sub = (a.AccountSubType || '').toLowerCase();
            const name = (a.Name || '').toLowerCase();
            return sub.includes('tax') || name.includes('gst') || name.includes('hst')
              || name.includes('pst') || name.includes('tax payable');
          })
          .map(a => String(a.Id))
      );

      // Detect tax lines by BOTH account ID and line description/name.
      // AI may map tax to a regular expense account but describe it as "GST", "PST (7%)", etc.
      const TAX_DESC_RE = /\b(GST|HST|PST|QST|RST|sales\s*tax|tax\s*\(\d)/i;

      const expenseLines = [];
      const taxLines = []; // { type: 'gst'|'hst'|'pst', amount }
      for (const line of jeLines) {
        const amt = Math.round(Number(line.amount) * 100) / 100;
        const desc = (line.description || line.accountName || '').toUpperCase();
        const isTaxByAccount = taxLiabilityIds.has(String(line.accountId));
        const isTaxByDesc = TAX_DESC_RE.test(line.description || '') || TAX_DESC_RE.test(line.accountName || '');

        if (isTaxByAccount || isTaxByDesc) {
          // Classify the tax type
          let taxType = 'gst'; // default
          if (/PST|QST|RST/i.test(desc)) taxType = 'pst';
          else if (/HST/i.test(desc)) taxType = 'hst';
          taxLines.push({ type: taxType, amount: amt, description: line.description });
          console.log(`[shareholder-invoice] tax line detected (${taxType}): "${line.description}" $${amt}`);
        } else {
          expenseLines.push({
            accountId: line.accountId,
            accountName: line.accountName,
            description: line.description || invoice.analysis.description,
            amount: amt,
          });
        }
      }

      const totalTaxAmount = taxLines.reduce((s, t) => s + t.amount, 0);
      const hasGST = taxLines.some(t => t.type === 'gst');
      const hasHST = taxLines.some(t => t.type === 'hst');
      const hasPST = taxLines.some(t => t.type === 'pst');
      console.log(`[shareholder-invoice] tax summary: GST=${hasGST} HST=${hasHST} PST=${hasPST} total=$${totalTaxAmount}`);

      // Find the appropriate QBO TaxCode based on which taxes are present
      let taxCodeId = null;
      if (totalTaxAmount > 0 && expenseLines.length > 0) {
        try {
          const tcData = await qbo.getTaxCodes(clientId);
          const taxCodes = (tcData?.QueryResponse?.TaxCode || []).filter(tc => tc.Active !== false);
          console.log('[shareholder-invoice] available tax codes:', taxCodes.map(tc => `${tc.Id}:${tc.Name}`).join(', '));

          // Match tax code based on what taxes are on the invoice:
          //   GST + PST → look for "GST/PST" combined code (e.g. "GST/PST BC")
          //   HST       → look for "HST" code
          //   GST only  → look for "GST" code
          let matched = null;
          const tcNames = taxCodes.map(tc => ({ ...tc, upper: (tc.Name || '').toUpperCase() }));

          if (hasGST && hasPST) {
            // Prefer combined GST/PST code
            matched = tcNames.find(tc => tc.upper.includes('GST') && tc.upper.includes('PST'));
            if (!matched) matched = tcNames.find(tc => tc.upper.includes('GST/PST'));
          }
          if (!matched && hasHST) {
            matched = tcNames.find(tc => tc.upper.includes('HST') && !tc.upper.includes('PST'));
          }
          if (!matched && hasGST && !hasPST) {
            // GST only — find a code that has GST but NOT PST
            matched = tcNames.find(tc => tc.upper.includes('GST') && !tc.upper.includes('PST') && tc.upper !== 'NON');
          }
          // Fallback: any active non-exempt tax code
          if (!matched) {
            matched = tcNames.find(tc => tc.upper !== 'NON' && tc.upper !== 'EXEMPT');
          }

          if (matched) {
            taxCodeId = String(matched.Id);
            console.log(`[shareholder-invoice] selected tax code: ${matched.Name} (${taxCodeId})`);
          }
        } catch (e) {
          console.error('[shareholder-invoice] failed to fetch tax codes:', e.message);
        }

        // Apply tax code to each expense line
        if (taxCodeId) {
          for (const line of expenseLines) {
            line.taxCodeId = taxCodeId;
          }
        } else {
          // No tax code found — absorb tax into expense lines proportionally
          console.log('[shareholder-invoice] no tax code found, absorbing tax into expense lines');
          const expTotal = expenseLines.reduce((s, l) => s + l.amount, 0);
          for (const line of expenseLines) {
            const share = expTotal > 0 ? line.amount / expTotal : 1 / expenseLines.length;
            line.amount = Math.round((line.amount + totalTaxAmount * share) * 100) / 100;
          }
        }
      }

      // If no expense lines at all (shouldn't happen), include all lines
      if (expenseLines.length === 0) {
        for (const line of jeLines) {
          expenseLines.push({
            accountId: line.accountId,
            accountName: line.accountName,
            description: line.description || invoice.analysis.description,
            amount: Math.round(Number(line.amount) * 100) / 100,
          });
        }
      }

      // Consolidate lines with the same GL account into a single line
      // to reduce noise in the GL (e.g. 4 lines all to "6100 Legal" → 1 line)
      const consolidated = new Map();
      for (const line of expenseLines) {
        const key = String(line.accountId);
        if (consolidated.has(key)) {
          const existing = consolidated.get(key);
          existing.amount = Math.round((existing.amount + line.amount) * 100) / 100;
          // Combine descriptions — use the overall invoice description if merging multiple
          if (line.description && !existing.descriptions.includes(line.description)) {
            existing.descriptions.push(line.description);
          }
        } else {
          consolidated.set(key, {
            ...line,
            descriptions: [line.description || ''],
          });
        }
      }
      // Replace expenseLines with consolidated version
      expenseLines.length = 0;
      for (const entry of consolidated.values()) {
        // Use invoice-level description when multiple lines are merged
        const desc = entry.descriptions.length > 1
          ? (invoice.analysis.description || entry.descriptions[0])
          : entry.descriptions[0];
        expenseLines.push({
          accountId: entry.accountId,
          accountName: entry.accountName,
          description: desc,
          amount: entry.amount,
          taxCodeId: entry.taxCodeId || undefined,
        });
      }
      console.log(`[shareholder-invoice] consolidated to ${expenseLines.length} line(s)`);

      // Find or create the vendor in QBO
      let vendorRef = null;
      const vendorName = invoice.analysis.vendor;
      if (vendorName) {
        try {
          vendorRef = await qbo.findOrCreateVendor(vendorName, clientId);
        } catch (vendorErr) {
          console.error(`[shareholder-invoice] vendor lookup/create failed for "${vendorName}":`, vendorErr.message);
        }
      }

      // Create a Purchase (Expense) transaction — paid from shareholder loan account
      const purchaseResult = await qbo.createPurchase({
        date: jeDate,
        memo,
        docNumber: invoice.analysis.invoiceNumber || '',
        accountId: c.shareholderLoanAccount.id,
        accountName: c.shareholderLoanAccount.name,
        accountType: c.shareholderLoanAccount.type || '',
        vendorId: vendorRef?.Id || null,
        vendorName: vendorRef?.DisplayName || vendorName || '',
        lines: expenseLines,
        globalTaxCalc: taxCodeId ? 'TaxExcluded' : 'NotApplicable',
      }, clientId);
      const txnId = purchaseResult.Id;

      // Attach the original invoice PDF/image to the expense in QBO
      let attachmentId = null;
      if (invoice.fileData) {
        try {
          const fileBuffer = Buffer.from(invoice.fileData, 'base64');
          const attachResult = await qbo.uploadAttachment({
            entityId: txnId,
            entityType: 'Purchase',
            fileName: invoice.fileName,
            contentType: invoice.fileType || 'application/pdf',
            buffer: fileBuffer,
          }, clientId);
          attachmentId = attachResult?.Id || null;
          console.log(`[shareholder-invoice] attached ${invoice.fileName} to Expense #${txnId} (attachment ${attachmentId})`);
        } catch (attachErr) {
          console.error(`[shareholder-invoice] failed to attach file to Expense #${txnId}:`, attachErr.message);
          // Don't fail the whole post — expense was created successfully
        }
      }

      invoice.journalEntry = {
        jeId: txnId,
        txnType: 'Purchase',
        date: jeDate,
        totalAmount,
        lineCount: jeLines.length,
        postedAt: new Date().toISOString(),
        attachmentId,
      };
    } else {
      // Offline mode — record intent without posting
      invoice.journalEntry = {
        jeId: null,
        date: jeDate,
        totalAmount,
        lineCount: jeLines.length,
        postedAt: new Date().toISOString(),
        offline: true,
      };
    }

    invoice.status = 'posted';
    saveShareholderInvoices(shareholderInvoiceData);

    res.json({ ok: true, journalEntry: invoice.journalEntry });
  } catch (e) {
    console.error('[shareholder-invoice] post error:', e.message);
    if (e.qboResponse) console.error('[shareholder-invoice] QBO response:', JSON.stringify(e.qboResponse, null, 2));
    if (e.response?.data) console.error('[shareholder-invoice] axios response data:', JSON.stringify(e.response.data, null, 2));
    // Extract detailed QBO error from all possible shapes
    const qboFault = e?.qboResponse?.Fault?.Error?.[0]?.Detail
      || e?.qboResponse?.Fault?.Error?.[0]?.Message
      || e?.Fault?.Error?.[0]?.Detail
      || e?.Fault?.Error?.[0]?.Message
      || e?.response?.data?.Fault?.Error?.[0]?.Detail
      || '';
    const detail = qboFault || e.message;
    console.error('[shareholder-invoice] detail:', detail);
    res.status(e.statusCode || 500).json({ error: detail });
  }
});

// PUT dismiss/undismiss an invoice
app.put('/api/admin/clients/:clientId/shareholder-invoices/:invoiceId/dismiss', requireAdmin, (req, res) => {
  const c = getClientShareholderInvoices(req.params.clientId);
  const invoice = c.invoices.find(i => i.id === req.params.invoiceId);
  if (!invoice) return res.status(404).json({ error: 'Invoice not found' });
  invoice.status = invoice.status === 'dismissed' ? 'pending' : 'dismissed';
  saveShareholderInvoices(shareholderInvoiceData);
  res.json({ ok: true, status: invoice.status });
});

// DELETE an invoice
app.delete('/api/admin/clients/:clientId/shareholder-invoices/:invoiceId', requireAdmin, (req, res) => {
  const c = getClientShareholderInvoices(req.params.clientId);
  const idx = c.invoices.findIndex(i => i.id === req.params.invoiceId);
  if (idx === -1) return res.status(404).json({ error: 'Invoice not found' });
  c.invoices.splice(idx, 1);
  saveShareholderInvoices(shareholderInvoiceData);
  res.json({ ok: true });
});

// ========================================
// START SERVER
// ========================================
app.listen(PORT, '0.0.0.0', () => {
  console.log(`jack does server running at http://localhost:${PORT}`);
  console.log(`Portal at http://localhost:${PORT}/portal/`);
  console.log(`Admin at http://localhost:${PORT}/admin`);

  if (!process.env.ANTHROPIC_API_KEY) {
    console.warn('\n⚠️  ANTHROPIC_API_KEY not set! Chat will not work.');
    console.warn('   Create a .env file with: ANTHROPIC_API_KEY=your-key-here\n');
  }

  if (!process.env.QBO_CLIENT_ID) {
    console.warn('⚠️  QuickBooks credentials not set! QBO integration will not work.');
  } else {
    console.log(`QuickBooks OAuth callback: http://localhost:${PORT}/api/qbo/callback`);
  }
});

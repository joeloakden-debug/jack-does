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
const CLIENTS_FILE = path.join(__dirname, 'clients.json');

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
const ADMIN_SETTINGS_FILE = path.join(__dirname, '.admin-settings.json');

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
const FIXED_ASSETS_FILE = path.join(__dirname, 'fixed-assets.json');

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
    // Pull any amortization JEs from QBO so the schedule reflects what's actually posted
    await hydrateAmortizationRunsFromQBO(clientId, clientData);

    // Build (or recompute) the continuity schedule for the most recent run month
    const runs = clientData.amortizationRuns || [];
    const sortedRuns = runs.slice().sort((a, b) => (b.month || '').localeCompare(a.month || ''));
    const asOfMonth = sortedRuns[0]?.month || formatMonthStr(new Date().getFullYear(), new Date().getMonth());
    const fye = CLIENTS[clientId]?.fiscalYearEnd || null;
    const continuitySchedule = buildContinuitySchedule(clientData, asOfMonth, fye);

    const workbook = await excelService.generateWorkbook(clientId, clientName, clientData, continuitySchedule);

    const fileName = `${clientName} - Fixed Assets.xlsx`.replace(/[<>:"/\\|?*]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Excel export error:', error);
    res.status(500).json({ error: 'Failed to generate Excel file' });
  }
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

    // Separate cost accounts from accumulated amortization/depreciation accounts
    const accumSubTypes = ['AccumulatedDepreciation', 'AccumulatedAmortization'];
    const costAccounts = fixedAssetAccounts
      .filter(a => !a.AccountSubType || !accumSubTypes.includes(a.AccountSubType));
    const accumAccounts = fixedAssetAccounts
      .filter(a => a.AccountSubType && accumSubTypes.includes(a.AccountSubType));

    // Get expense accounts for auto-suggesting amortization expense account
    const expenseAccounts = allAccounts
      .filter(a => ['Expense', 'Other Expense'].includes(a.AccountType) && a.Active);

    // Check which assets are already imported
    const clientData = getClientAssets(req.params.clientId);
    const existingTxnKeys = new Set(clientData.assets.filter(a => a.active).map(a => a.txnKey).filter(Boolean));

    // Query the General Ledger Detail for all cost accounts to find individual transactions
    const costAccountIds = costAccounts.map(a => a.Id).join(',');
    let glReport = null;
    if (costAccountIds) {
      try {
        glReport = await qbo.getGeneralLedgerForAccount(costAccountIds, syncClientId);
      } catch (e) {
        console.error('Could not fetch GL detail for fixed assets:', e.message);
      }
    }

    // Parse GL report into individual asset transactions
    const importable = [];

    if (glReport?.Rows?.Row) {
      // Build column index map from report metadata. We start with -1 (not present) and
      // resolve each field by walking glReport.Columns.Column. QBO can return very few
      // columns (e.g. just Memo/Description, Split, Amount, Balance for some accounts), so
      // hardcoded positions are unsafe.
      const reportCols = Array.isArray(glReport.Columns?.Column)
        ? glReport.Columns.Column
        : (glReport.Columns?.Column ? [glReport.Columns.Column] : []);
      console.log('[sync] raw report columns:', JSON.stringify(reportCols));

      const colIdx = { date: -1, txnType: -1, docNum: -1, name: -1, memo: -1, amount: -1, debit: -1, credit: -1, split: -1 };
      reportCols.forEach((c, i) => {
        const title = String(c.ColTitle || '').toLowerCase().trim();
        const type = String(c.ColType || '').toLowerCase().trim();
        if (title.includes('memo') || title.includes('description') || type === 'memo') colIdx.memo = i;
        else if (title === 'date' || type === 'tx_date' || type === 'date') colIdx.date = i;
        else if (title.includes('transaction type') || type === 'txn_type') colIdx.txnType = i;
        else if (title === 'num' || title.includes('doc num') || type === 'doc_num') colIdx.docNum = i;
        else if (title === 'name' || type === 'name' || title === 'vendor') colIdx.name = i;
        else if (title === 'split' || type === 'split_acc' || type === 'split') colIdx.split = i;
        else if (title === 'amount' || type === 'subt_amount' || type === 'amount') colIdx.amount = i;
        else if (title === 'debit' || type === 'debt_amt' || type === 'debit') colIdx.debit = i;
        else if (title === 'credit' || type === 'credt_amt' || type === 'credit') colIdx.credit = i;
      });
      console.log('[sync] GL Detail column map:', JSON.stringify(colIdx), 'titles:', reportCols.map(c => c.ColTitle));

      // GL Detail report structure: rows grouped by account, then individual transactions
      function parseGLRows(rows, parentAccountId, parentAccountName) {
        for (const row of rows) {
          if (row.Header?.ColData) {
            // This is an account header — get the account name
            const acctName = row.Header.ColData[0]?.value || '';
            // Find the matching QBO account
            const matchedAcct = costAccounts.find(a => a.Name === acctName || a.FullyQualifiedName === acctName);
            const acctId = matchedAcct?.Id || parentAccountId;
            const acctNameResolved = matchedAcct?.Name || acctName || parentAccountName;

            // Parse child rows (the actual transactions)
            if (row.Rows?.Row) {
              for (const txnRow of row.Rows.Row) {
                if (txnRow.ColData) {
                  const cols = txnRow.ColData;
                  console.log('[sync] row in', acctNameResolved, ':', JSON.stringify(cols));
                  const get = (idx) => (idx >= 0 && cols[idx]) ? (cols[idx].value || '') : '';
                  const txnDate = get(colIdx.date);
                  const txnType = get(colIdx.txnType);
                  const docNum = get(colIdx.docNum);
                  const vendorName = get(colIdx.name);
                  const memo = get(colIdx.memo);
                  // parseAmount: strip currency symbols, commas, parens, and parse with full precision
                  // Important: never round — preserve exact GL value
                  const parseAmount = (v) => {
                    if (v === null || v === undefined || v === '') return 0;
                    const s = String(v).replace(/[$,\s]/g, '').replace(/^\((.*)\)$/, '-$1');
                    const n = parseFloat(s);
                    return isNaN(n) ? 0 : n;
                  };
                  // Try amount column first; fall back to debit/credit columns
                  let amount = parseAmount(get(colIdx.amount));
                  if (amount === 0) amount = parseAmount(get(colIdx.debit));
                  if (amount === 0) amount = parseAmount(get(colIdx.credit));

                  // Skip zero amounts and closing/total rows
                  if (amount <= 0) continue;

                  // Build a unique key for this transaction to avoid re-importing
                  const txnKey = `${acctId}-${txnDate}-${amount}-${memo || docNum}`;

                  if (!existingTxnKeys.has(txnKey)) {
                    // Auto-match accumulated amortization account
                    const nameBase = acctNameResolved.toLowerCase().replace(/[-–—]/g, ' ').trim();
                    const matchedAccum = accumAccounts.find(acc => {
                      const accumName = acc.Name.toLowerCase().replace(/[-–—]/g, ' ').trim();
                      return accumName.includes(nameBase) || nameBase.includes(accumName.replace(/accumulated\s*(amortization|depreciation)\s*/i, '').trim());
                    });

                    // Auto-match expense account
                    const matchedExpense = expenseAccounts.find(exp => {
                      const expName = exp.Name.toLowerCase();
                      return expName.includes('amortization') || expName.includes('depreciation');
                    }) || expenseAccounts.find(exp => {
                      const words = nameBase.split(/\s+/);
                      const expName = exp.Name.toLowerCase();
                      return words.some(w => w.length > 3 && expName.includes(w));
                    });

                    // Use memo/vendor as the asset description (e.g. "Asus Vivobook")
                    const assetDescription = memo || vendorName || '';
                    const assetName = assetDescription || acctNameResolved;

                    importable.push({
                      qboAccountId: acctId,
                      glAccountName: acctNameResolved,
                      name: assetName,
                      description: assetDescription || '',
                      accountType: 'Fixed Asset',
                      originalCost: amount,
                      txnDate,
                      txnType,
                      docNum,
                      vendorName,
                      memo,
                      txnKey,
                      suggestedAccumAccountId: matchedAccum?.Id || '',
                      suggestedAccumAccountName: matchedAccum?.Name || '',
                      suggestedExpenseAccountId: matchedExpense?.Id || '',
                      suggestedExpenseAccountName: matchedExpense?.Name || '',
                    });
                  }
                }
                // Recurse into sub-rows if present
                if (txnRow.Rows?.Row) {
                  parseGLRows([txnRow], acctId, acctNameResolved);
                }
              }
            }
          }
          // Also recurse for nested structures
          if (row.Rows?.Row && !row.Header) {
            parseGLRows(row.Rows.Row, parentAccountId, parentAccountName);
          }
        }
      }

      parseGLRows(glReport.Rows.Row, null, null);
    }

    // Fallback: if GL report returned nothing, fall back to account-level import
    if (importable.length === 0 && costAccounts.length > 0) {
      const existingQboIds = new Set(clientData.assets.map(a => a.qboAccountId).filter(Boolean));
      for (const a of costAccounts) {
        if (existingQboIds.has(a.Id)) continue;
        const nameBase = a.Name.toLowerCase().replace(/[-–—]/g, ' ').trim();
        const matchedAccum = accumAccounts.find(acc => {
          const accumName = acc.Name.toLowerCase().replace(/[-–—]/g, ' ').trim();
          return accumName.includes(nameBase) || nameBase.includes(accumName.replace(/accumulated\s*(amortization|depreciation)\s*/i, '').trim());
        });
        const matchedExpense = expenseAccounts.find(exp => exp.Name.toLowerCase().includes('amortization') || exp.Name.toLowerCase().includes('depreciation'));

        importable.push({
          qboAccountId: a.Id,
          glAccountName: a.Name,
          name: a.Name || a.FullyQualifiedName,
          description: a.Description || '',
          accountType: a.AccountSubType || a.AccountType,
          originalCost: a.CurrentBalance ?? 0,
          txnDate: a.MetaData?.CreateTime ? a.MetaData.CreateTime.split('T')[0] : null,
          txnKey: null,
          suggestedAccumAccountId: matchedAccum?.Id || '',
          suggestedAccumAccountName: matchedAccum?.Name || '',
          suggestedExpenseAccountId: matchedExpense?.Id || '',
          suggestedExpenseAccountName: matchedExpense?.Name || '',
        });
      }
    }

    res.json({ accounts: importable });
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

  for (const group of schedule.glAccounts) {
    const glName = group.glAccountName;
    // Cost check: compare schedule cost subtotal to TB balance for this account
    const tbCost = tbBalances.get(glName);
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
      const tbAccum = tbBalances.get(accumName);
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
 * QBO is the source of truth — any amortization JE found there is merged into local
 * runs (keyed by month). New entries are persisted so subsequent operations are fast.
 * Safe no-op if QBO is not connected.
 */
async function hydrateAmortizationRunsFromQBO(clientId, clientData) {
  if (!qbo.isConnected(clientId)) return clientData;
  try {
    const qboRuns = await qbo.getAmortizationRunsFromQBO(clientId);
    if (!qboRuns || qboRuns.length === 0) return clientData;

    const localByMonth = new Map((clientData.amortizationRuns || []).map(r => [r.month, r]));
    let changed = false;
    for (const qr of qboRuns) {
      const existing = localByMonth.get(qr.month);
      if (!existing) {
        // Brand new — add it
        clientData.amortizationRuns.push(qr);
        localByMonth.set(qr.month, qr);
        changed = true;
      } else if (!existing.journalEntryId && qr.journalEntryId) {
        // Local record missing JE id — patch it
        existing.journalEntryId = qr.journalEntryId;
        existing.sourceQBO = true;
        changed = true;
      }
    }
    if (changed) {
      // Sort by month for sanity
      clientData.amortizationRuns.sort((a, b) => (a.month || '').localeCompare(b.month || ''));
      saveFixedAssets(fixedAssetsData);
    }
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

// Update QBO book close date (push back to QBO)
app.post('/api/admin/clients/:clientId/book-close-date', requireAdmin, async (req, res) => {
  try {
    if (!qbo.isConnected(req.params.clientId)) {
      return res.status(400).json({ error: 'QuickBooks is not connected for this client' });
    }
    const { closeDate } = req.body || {};
    if (closeDate && !/^\d{4}-\d{2}-\d{2}$/.test(closeDate)) {
      return res.status(400).json({ error: 'closeDate must be YYYY-MM-DD or null' });
    }
    const saved = await qbo.updateBookCloseDate(req.params.clientId, closeDate || null);
    res.json({ success: true, closeDate: saved });
  } catch (e) {
    console.error('Update book close date error:', e.message);
    res.status(500).json({ error: `Failed to update QBO close date: ${e.message}` });
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
  if (qbo.isConnected(req.params.clientId)) {
    await hydrateAmortizationRunsFromQBO(req.params.clientId, clientData);
  }
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
    await hydrateAmortizationRunsFromQBO(req.params.clientId, clientData);
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

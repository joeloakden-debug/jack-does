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

// Start QuickBooks connection
app.get('/api/qbo/connect', (req, res) => {
  const authUri = qbo.getAuthUri();
  res.redirect(authUri);
});

// OAuth callback
app.get('/api/qbo/callback', async (req, res) => {
  try {
    const tokenData = await qbo.handleCallback(req.url);
    console.log('QuickBooks connected! Realm ID:', tokenData.realmId);

    // Redirect back to portal with success
    res.redirect('/portal/dashboard.html?qbo=connected');
  } catch (error) {
    console.error('QuickBooks OAuth error:', error.message);
    res.redirect('/portal/dashboard.html?qbo=error');
  }
});

// Check connection status
app.get('/api/qbo/status', (req, res) => {
  res.json({ connected: qbo.isConnected() });
});

// Admin endpoint to get current QBO tokens (for saving to Railway env var)
app.get('/api/admin/qbo-tokens', requireAdmin, (req, res) => {
  if (!qbo.isConnected()) {
    return res.json({ connected: false, tokens: null });
  }
  // Read the token file
  const tokenFile = require('path').join(__dirname, '.qbo-tokens.json');
  try {
    const data = fs.readFileSync(tokenFile, 'utf-8');
    res.json({ connected: true, tokens: data });
  } catch (e) {
    res.json({ connected: true, tokens: null, error: 'Token file not found' });
  }
});

// Disconnect QuickBooks
app.post('/api/qbo/disconnect', (req, res) => {
  qbo.disconnect();
  res.json({ success: true, message: 'QuickBooks disconnected' });
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

    // If QuickBooks is connected, fetch relevant financial data
    let qboContext = '';
    if (qbo.isConnected()) {
      try {
        // Include recent USER messages (not assistant responses) for follow-up context
        const recentUserMessages = history
          .filter(m => m.role === 'user')
          .slice(-3)
          .map(m => m.content)
          .join(' ');
        const combinedQuery = `${message} ${recentUserMessages}`;
        console.log('[QBO] Fetching data for query keywords:', message.substring(0, 100));
        const data = await qbo.fetchRelevantData(combinedQuery);
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
    if (!qbo.isConnected()) {
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
      qboConnected: qbo.isConnected(),
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

    // If QuickBooks is connected, fetch account list so Claude can map to real accounts
    let accountContext = '';
    if (qbo.isConnected()) {
      try {
        const accountsData = await qbo.getAccounts();
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

    if (!qbo.isConnected()) {
      return res.status(400).json({ error: 'QuickBooks is not connected. Please connect QuickBooks first.' });
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
          });
        } else if (entry.type === 'bill') {
          result = await qbo.createBill({
            date: entry.date,
            memo: entry.memo,
            vendorId: entry.vendorId,
            dueDate: entry.dueDate,
            lines: entry.lines,
          });
        } else if (entry.type === 'invoice') {
          result = await qbo.createInvoice({
            date: entry.date,
            memo: entry.memo,
            customerId: entry.customerId,
            dueDate: entry.dueDate,
            lines: entry.lines,
          });
        } else {
          // Default to journal entry
          result = await qbo.createJournalEntry({
            date: entry.date,
            memo: entry.memo,
            lines: entry.lines,
          });
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

    if (!qbo.isConnected()) {
      return res.status(400).json({ error: 'QuickBooks is not connected.' });
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
          result = await qbo.createBill({ date: entry.date, memo: entry.memo, vendorId: entry.vendorId, dueDate: entry.dueDate, lines: entry.lines });
        } else if (entry.type === 'invoice') {
          result = await qbo.createInvoice({ date: entry.date, memo: entry.memo, customerId: entry.customerId, dueDate: entry.dueDate, lines: entry.lines });
        } else {
          result = await qbo.createJournalEntry({ date: entry.date, memo: entry.memo, lines: entry.lines });
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
  if (!qbo.isConnected()) {
    return res.json({ accounts: [] });
  }

  try {
    const data = await qbo.getAccounts();
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
  res.json({ pending, approved, rejected, total: documentAnalyses.size, qboConnected: qbo.isConnected(), clientCount: Object.keys(CLIENTS).length });
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
  const { name, email, password, billingFrequency } = req.body;
  if (!name || !email) {
    return res.status(400).json({ error: 'Name and email are required' });
  }
  if (!password) {
    return res.status(400).json({ error: 'Password is required' });
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
  const { name, email, password, billingFrequency } = req.body;
  if (name) client.name = name;
  if (email) client.email = email;
  if (password) client.password = password;
  const validFreqs = ['monthly', 'quarterly', 'annual'];
  if (billingFrequency && validFreqs.includes(billingFrequency)) {
    client.billingFrequency = billingFrequency;
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
    fixedAssetsData[clientId] = { assets: [], amortizationRuns: [] };
  }
  return fixedAssetsData[clientId];
}

let fixedAssetsData = loadFixedAssets();

// Get fixed assets for a specific client
app.get('/api/admin/clients/:clientId/fixed-assets', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  res.json(clientData);
});

// Sync fixed assets from QBO — pulls Fixed Asset accounts and their balances
app.post('/api/admin/clients/:clientId/fixed-assets/sync-qbo', requireAdmin, async (req, res) => {
  try {
    if (!qbo.isConnected()) {
      return res.status(400).json({ error: 'QuickBooks is not connected' });
    }

    // Get all accounts of type Fixed Asset
    const accountsResult = await qbo.getAccounts();
    const accounts = accountsResult.QueryResponse?.Account || accountsResult || [];
    const fixedAssetAccounts = (Array.isArray(accounts) ? accounts : [])
      .filter(a => a.AccountType === 'Fixed Asset' && a.Active);

    // Get balance sheet for current balances
    const today = new Date().toISOString().split('T')[0];
    let balanceSheet = null;
    try {
      balanceSheet = await qbo.getBalanceSheet(today);
    } catch (e) {
      console.error('Could not fetch balance sheet for asset balances:', e.message);
    }

    // Build balance lookup from balance sheet
    const balanceLookup = {};
    if (balanceSheet?.Rows?.Row) {
      function extractBalances(rows) {
        for (const row of rows) {
          if (row.ColData) {
            const name = row.ColData[0]?.value;
            const val = parseFloat(row.ColData[1]?.value);
            if (name && !isNaN(val)) balanceLookup[name] = val;
          }
          if (row.Rows?.Row) extractBalances(row.Rows.Row);
        }
      }
      extractBalances(balanceSheet.Rows.Row);
    }

    // Map to importable assets — only parent/detail accounts (not sub-totals)
    const clientData = getClientAssets(req.params.clientId);
    const existingQboIds = new Set(clientData.assets.map(a => a.qboAccountId).filter(Boolean));

    const importable = fixedAssetAccounts
      .filter(a => !a.AccountSubType || a.AccountSubType !== 'AccumulatedDepreciation')
      .filter(a => !existingQboIds.has(a.Id))
      .map(a => ({
        qboAccountId: a.Id,
        name: a.Name || a.FullyQualifiedName,
        accountType: a.AccountSubType || a.AccountType,
        currentBalance: balanceLookup[a.Name] ?? a.CurrentBalance ?? 0,
        description: a.Description || '',
      }));

    res.json({ accounts: importable, allFixedAssetAccounts: fixedAssetAccounts });
  } catch (error) {
    console.error('QBO sync error:', error.message);
    res.status(500).json({ error: `Failed to fetch from QBO: ${error.message}` });
  }
});

// AI suggest amortization period based on asset name/type
app.post('/api/admin/fixed-assets/suggest-amortization', requireAdmin, async (req, res) => {
  try {
    const { assetName, assetType, originalCost } = req.body;

    const response = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 300,
      system: `You are a Canadian accounting expert. Given a fixed asset, suggest the appropriate amortization period in months and salvage value based on CRA Capital Cost Allowance (CCA) classes and standard accounting practices. Respond in JSON only with this exact format: {"usefulLifeMonths": number, "salvageValue": number, "ccaClass": number or null, "ccaRate": "percentage or null", "reasoning": "brief explanation"}`,
      messages: [{ role: 'user', content: `Asset: "${assetName}". Type: "${assetType || 'unknown'}". Original cost: $${originalCost || 'unknown'}. What amortization period and salvage value would you recommend?` }],
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
  const { name, originalCost, usefulLifeMonths, salvageValue, acquisitionDate,
          assetAccountId, assetAccountName, expenseAccountId, expenseAccountName,
          accumAccountId, accumAccountName, qboAccountId } = req.body;

  if (!name || !originalCost || !usefulLifeMonths || !acquisitionDate) {
    return res.status(400).json({ error: 'Name, cost, useful life, and acquisition date are required' });
  }
  if (!assetAccountId || !expenseAccountId || !accumAccountId) {
    return res.status(400).json({ error: 'All three QBO accounts must be selected' });
  }

  const clientData = getClientAssets(req.params.clientId);

  const asset = {
    id: 'asset_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
    clientId: req.params.clientId,
    name,
    originalCost: parseFloat(originalCost),
    usefulLifeMonths: parseInt(usefulLifeMonths, 10),
    salvageValue: parseFloat(salvageValue || 0),
    acquisitionDate,
    assetAccountId, assetAccountName,
    expenseAccountId, expenseAccountName,
    accumAccountId, accumAccountName,
    qboAccountId: qboAccountId || null,
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

  const fields = ['name', 'originalCost', 'usefulLifeMonths', 'salvageValue', 'acquisitionDate',
    'assetAccountId', 'assetAccountName', 'expenseAccountId', 'expenseAccountName',
    'accumAccountId', 'accumAccountName', 'active'];

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

// Preview amortization for a client
app.get('/api/admin/clients/:clientId/fixed-assets/preview-amortization', requireAdmin, (req, res) => {
  const clientData = getClientAssets(req.params.clientId);
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

  const alreadyRun = clientData.amortizationRuns.find(r => r.month === currentMonth);
  if (alreadyRun) {
    return res.json({ alreadyRun: true, runDetails: alreadyRun, month: currentMonth });
  }

  const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  const eligible = clientData.assets.filter(a => {
    if (!a.active) return false;
    const acqDate = new Date(a.acquisitionDate);
    if (acqDate > lastDayOfMonth) return false;
    const acqMonth = acqDate.getFullYear() * 12 + acqDate.getMonth();
    const curMonth = now.getFullYear() * 12 + now.getMonth();
    return (curMonth - acqMonth) < a.usefulLifeMonths;
  });

  const lines = [];
  let totalAmount = 0;
  eligible.forEach(a => {
    const monthly = Math.round(((a.originalCost - a.salvageValue) / a.usefulLifeMonths) * 100) / 100;
    totalAmount += monthly;
    lines.push({ assetName: a.name, amount: monthly, expenseAccountName: a.expenseAccountName, accumAccountName: a.accumAccountName });
  });

  res.json({ alreadyRun: false, month: currentMonth, eligibleCount: eligible.length, totalAmount: Math.round(totalAmount * 100) / 100, lines });
});

// Run amortization for a client and post JE to QBO
app.post('/api/admin/clients/:clientId/fixed-assets/run-amortization', requireAdmin, async (req, res) => {
  try {
    const clientData = getClientAssets(req.params.clientId);
    const now = new Date();
    const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

    const alreadyRun = clientData.amortizationRuns.find(r => r.month === currentMonth);
    if (alreadyRun) {
      return res.status(400).json({ error: `Amortization already run for ${currentMonth}` });
    }

    if (!qbo.isConnected()) {
      return res.status(400).json({ error: 'QuickBooks is not connected. Please connect QBO first.' });
    }

    const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    const lastDayStr = lastDayOfMonth.toISOString().split('T')[0];

    const eligible = clientData.assets.filter(a => {
      if (!a.active) return false;
      const acqDate = new Date(a.acquisitionDate);
      if (acqDate > lastDayOfMonth) return false;
      const acqMonth = acqDate.getFullYear() * 12 + acqDate.getMonth();
      const curMonth = now.getFullYear() * 12 + now.getMonth();
      return (curMonth - acqMonth) < a.usefulLifeMonths;
    });

    if (eligible.length === 0) {
      return res.status(400).json({ error: 'No eligible assets for amortization this month' });
    }

    const jeLines = [];
    let totalAmount = 0;

    eligible.forEach(a => {
      const monthly = Math.round(((a.originalCost - a.salvageValue) / a.usefulLifeMonths) * 100) / 100;
      totalAmount += monthly;
      jeLines.push({ accountId: a.expenseAccountId, accountName: a.expenseAccountName, description: `Amortization - ${a.name}`, amount: monthly, type: 'debit' });
      jeLines.push({ accountId: a.accumAccountId, accountName: a.accumAccountName, description: `Amortization - ${a.name}`, amount: monthly, type: 'credit' });
    });

    const clientName = CLIENTS[req.params.clientId]?.name || req.params.clientId;
    const result = await qbo.createJournalEntry({
      date: lastDayStr,
      memo: `Fixed asset amortization - ${currentMonth} - ${clientName}`,
      lines: jeLines,
    });

    const runRecord = {
      month: currentMonth,
      ranAt: new Date().toISOString(),
      journalEntryId: result.Id || null,
      totalAmount: Math.round(totalAmount * 100) / 100,
      assetCount: eligible.length,
      assets: eligible.map(a => ({ id: a.id, name: a.name, amount: Math.round(((a.originalCost - a.salvageValue) / a.usefulLifeMonths) * 100) / 100 })),
    };

    clientData.amortizationRuns.push(runRecord);
    saveFixedAssets(fixedAssetsData);
    res.json({ success: true, run: runRecord });
  } catch (error) {
    console.error('Amortization run error:', error.message);
    res.status(500).json({ error: `Failed to post journal entry: ${error.message}` });
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

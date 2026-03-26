// ========================================
// CLIENT IDENTITY (swap with real auth in production)
// ========================================
const CLIENT_ID = 'demo-client';

// Helper: add clientId header to fetch requests
function clientHeaders(extra = {}) {
  return { 'X-Client-Id': CLIENT_ID, ...extra };
}

// ========================================
// QUICKBOOKS CONNECTION STATUS
// ========================================
async function checkQBOStatus() {
  const banner = document.getElementById('qbo-banner');
  if (!banner) return;

  try {
    const res = await fetch('/api/qbo/status');
    const data = await res.json();

    if (data.connected) {
      banner.classList.add('connected');
      banner.querySelector('.qbo-banner-text').innerHTML =
        '<strong>quickbooks connected</strong> — jack has access to your financial data';
      const btn = banner.querySelector('.btn-qbo');
      btn.textContent = 'connected';
      btn.classList.add('connected');
    }
  } catch (e) {
    // Server not running, hide banner
  }

  // Check URL params for connection result
  const params = new URLSearchParams(window.location.search);
  if (params.get('qbo') === 'connected') {
    banner.classList.add('connected');
    banner.querySelector('.qbo-banner-text').innerHTML =
      '<strong>quickbooks connected</strong> — jack has access to your financial data';
    const btn = banner.querySelector('.btn-qbo');
    btn.textContent = 'connected';
    btn.classList.add('connected');
    // Clean URL
    window.history.replaceState({}, '', window.location.pathname);
  }
}

checkQBOStatus();


// ========================================
// VIEW SWITCHING
// ========================================
const sidebarLinks = document.querySelectorAll('.sidebar-link[data-view]');
const views = document.querySelectorAll('.view');

sidebarLinks.forEach(link => {
  link.addEventListener('click', (e) => {
    e.preventDefault();
    const viewId = link.dataset.view;

    // Update active states
    sidebarLinks.forEach(l => l.classList.remove('active'));
    link.classList.add('active');

    views.forEach(v => v.classList.remove('active'));
    document.getElementById(`view-${viewId}`).classList.add('active');
  });
});


// ========================================
// CHAT FUNCTIONALITY
// ========================================
const chatMessages = document.getElementById('chat-messages');
const chatInput = document.getElementById('chat-input');
const chatSend = document.getElementById('chat-send');

// Auto-resize textarea
chatInput.addEventListener('input', () => {
  chatInput.style.height = 'auto';
  chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
});

// Send message on enter (shift+enter for new line)
chatInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    sendMessage();
  }
});

chatSend.addEventListener('click', sendMessage);

// Generate a simple session ID for conversation continuity
const sessionId = 'session_' + Math.random().toString(36).substring(2, 15);

async function sendMessage() {
  const message = chatInput.value.trim();
  if (!message) return;

  // Add user message
  appendMessage('user', message);
  chatInput.value = '';
  chatInput.style.height = 'auto';

  // Show typing indicator
  const typingEl = appendTyping();

  try {
    const response = await fetch('/api/chat', {
      method: 'POST',
      headers: clientHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ message, sessionId }),
    });

    const data = await response.json();
    typingEl.remove();

    if (data.error) {
      appendMessage('jack', `sorry, something went wrong: ${data.error}`);
    } else {
      appendMessage('jack', data.response);
    }
  } catch (error) {
    typingEl.remove();
    appendMessage('jack', "sorry, i'm having trouble connecting right now. please try again in a moment.");
  }
}

function escapeHtml(text) {
  const el = document.createElement('span');
  el.textContent = text;
  return el.innerHTML;
}

function appendMessage(sender, text) {
  const div = document.createElement('div');
  div.className = `chat-message ${sender}`;

  if (sender === 'jack') {
    // Jack's messages render as HTML (tables, bold, etc.)
    div.innerHTML = `
      <img src="../jack-avatar-clean.png" alt="jack" class="chat-avatar">
      <div class="chat-bubble">${text}</div>
    `;
  } else {
    // User messages are escaped plain text
    div.innerHTML = `
      <div class="chat-bubble"><p>${escapeHtml(text)}</p></div>
    `;
  }

  chatMessages.appendChild(div);
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

function appendTyping() {
  const div = document.createElement('div');
  div.className = 'chat-message jack';
  div.innerHTML = `
    <img src="../jack-avatar-clean.png" alt="jack" class="chat-avatar">
    <div class="chat-bubble">
      <div class="chat-typing">
        <span></span><span></span><span></span>
      </div>
    </div>
  `;
  chatMessages.appendChild(div);
  chatMessages.scrollTop = chatMessages.scrollHeight;
  return div;
}



// ========================================
// FILE UPLOAD
// ========================================
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const uploadQueue = document.getElementById('upload-queue');

// Click to upload
dropzone.addEventListener('click', () => fileInput.click());

// Drag and drop
dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropzone.classList.add('dragover');
});

dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('dragover');
});

dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  handleFiles(e.dataTransfer.files);
});

fileInput.addEventListener('change', () => {
  handleFiles(fileInput.files);
  fileInput.value = '';
});

function handleFiles(files) {
  Array.from(files).forEach(file => {
    addFileToQueue(file);
  });
}

async function addFileToQueue(file) {
  const size = formatFileSize(file.size);
  const div = document.createElement('div');
  div.className = 'upload-file';
  div.innerHTML = `
    <div class="upload-file-icon">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/></svg>
    </div>
    <div class="upload-file-info">
      <div class="upload-file-name">${file.name}</div>
      <div class="upload-file-size">${size}</div>
    </div>
    <span class="upload-file-status uploading">uploading...</span>
  `;
  uploadQueue.appendChild(div);

  // Get selected category
  const activeCategory = document.querySelector('.category-tag.active');
  const category = activeCategory ? activeCategory.textContent.trim() : 'general';

  try {
    const formData = new FormData();
    formData.append('files', file);
    formData.append('category', category);

    const response = await fetch('/api/upload', {
      method: 'POST',
      headers: { 'X-Client-Id': CLIENT_ID },
      body: formData,
    });

    const data = await response.json();
    const status = div.querySelector('.upload-file-status');

    if (data.success) {
      status.textContent = 'analyzing...';
      status.style.background = '#fff7ed';
      status.style.color = '#ea580c';

      // Auto-process the uploaded file through Claude
      const uploadedFile = data.files[0];
      processDocument(div, uploadedFile, category);
    } else {
      status.textContent = 'failed';
      status.classList.remove('uploading');
      status.style.background = '#fef2f2';
      status.style.color = '#ef4444';
    }
  } catch (error) {
    const status = div.querySelector('.upload-file-status');
    status.textContent = 'failed';
    status.classList.remove('uploading');
    status.style.background = '#fef2f2';
    status.style.color = '#ef4444';
  }
}

/**
 * Send uploaded document to Claude for analysis, queue for admin review
 */
async function processDocument(fileDiv, uploadedFile, category) {
  const status = fileDiv.querySelector('.upload-file-status');

  try {
    const response = await fetch('/api/process-document', {
      method: 'POST',
      headers: clientHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({
        filePath: uploadedFile.path,
        fileName: uploadedFile.name,
        category,
      }),
    });

    const data = await response.json();

    if (data.success && data.analysis) {
      status.textContent = 'jack is reviewing';
      status.classList.remove('uploading');
      status.style.background = '#eff6ff';
      status.style.color = '#2563eb';

      // Show a friendly note below the file
      const note = document.createElement('div');
      note.className = 'entry-note client-review-note';
      note.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" width="14" height="14"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> jack is analyzing your document. you'll be notified once it's been processed.`;
      fileDiv.after(note);
    } else {
      status.textContent = 'uploaded';
      status.classList.remove('uploading');
      status.classList.add('success');
    }
  } catch (error) {
    status.textContent = 'uploaded';
    status.classList.remove('uploading');
    status.classList.add('success');
  }
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}


// ========================================
// MY UPLOADED FILES — load and display
// ========================================
async function loadUploadedFiles() {
  const filesList = document.getElementById('files-list');
  if (!filesList) return;

  try {
    const res = await fetch('/api/files', { headers: clientHeaders() });
    const data = await res.json();

    if (!data.files || data.files.length === 0) {
      filesList.innerHTML = `
        <div class="files-empty">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" class="files-empty-icon"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/></svg>
          <p>no files uploaded yet</p>
          <button class="btn btn-primary" onclick="document.querySelector('[data-view=dashboard]').click()">upload your first document</button>
        </div>`;
      return;
    }

    filesList.innerHTML = `
      <table class="files-table">
        <thead>
          <tr>
            <th></th>
            <th>file name</th>
            <th>category</th>
            <th>size</th>
            <th>uploaded</th>
          </tr>
        </thead>
        <tbody>
          ${data.files.map(f => `
            <tr>
              <td>
                <div class="file-icon-sm">
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/></svg>
                </div>
              </td>
              <td class="file-name-cell">${f.name}</td>
              <td><span class="file-category-tag">${f.category}</span></td>
              <td class="file-size-cell">${formatFileSize(f.size)}</td>
              <td class="file-date-cell">${formatFileDate(f.uploadedAt)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>`;
  } catch (e) {
    filesList.innerHTML = `
      <div class="files-empty">
        <p>could not load files</p>
      </div>`;
  }
}

function formatFileDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: 'numeric', minute: '2-digit',
  });
}

// Load data when switching views
document.querySelectorAll('.sidebar-link[data-view]').forEach(link => {
  link.addEventListener('click', () => {
    if (link.dataset.view === 'files') {
      loadUploadedFiles();
    }
    if (link.dataset.view === 'notifications') {
      loadNotifications();
      markNotificationsRead();
    }
  });
});


// ========================================
// CATEGORY TAGS
// ========================================
document.querySelectorAll('.category-tag').forEach(tag => {
  tag.addEventListener('click', () => {
    document.querySelectorAll('.category-tag').forEach(t => t.classList.remove('active'));
    tag.classList.add('active');
  });
});


// ========================================
// LOGIN FORM (placeholder)
// ========================================
const loginForm = document.getElementById('login-form');
if (loginForm) {
  loginForm.addEventListener('submit', (e) => {
    e.preventDefault();
    // In production, this would authenticate against your backend
    window.location.href = 'dashboard.html';
  });
}


// ========================================
// NOTIFICATIONS
// ========================================
let cachedNotifications = [];

async function loadNotificationBadge() {
  try {
    const res = await fetch('/api/notifications', { headers: clientHeaders() });
    const data = await res.json();
    cachedNotifications = data.notifications || [];
    updateBadge(data.unreadCount || 0);
  } catch (e) {
    // Server not running
  }
}

function updateBadge(count) {
  const badge = document.getElementById('notification-badge');
  if (!badge) return;
  if (count > 0) {
    badge.textContent = count;
    badge.classList.add('visible');
  } else {
    badge.classList.remove('visible');
  }
}

async function loadNotifications() {
  const list = document.getElementById('notifications-list');
  if (!list) return;

  try {
    const res = await fetch('/api/notifications', { headers: clientHeaders() });
    const data = await res.json();
    cachedNotifications = data.notifications || [];
    updateBadge(data.unreadCount || 0);

    if (cachedNotifications.length === 0) {
      list.innerHTML = `
        <div class="notifications-empty">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>
          <p>no notifications</p>
          <span>when jack needs your attention on a document, it'll show up here</span>
        </div>`;
      return;
    }

    list.innerHTML = cachedNotifications.map(n => `
      <div class="notification-card ${n.read ? '' : 'unread'}" data-id="${n.id}">
        <div class="notification-card-header">
          <div class="notification-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          </div>
          <div class="notification-info">
            <div class="notification-filename">${n.fileName || 'unknown file'}</div>
            <div class="notification-meta">
              <span class="status-rejected">rejected</span>
              <span>${formatFileDate(n.reviewedAt)}</span>
            </div>
          </div>
        </div>
        ${n.rejectReason ? `
          <div class="notification-reason">
            <strong>reason:</strong> ${n.rejectReason}
          </div>
        ` : `
          <div class="notification-reason">
            <strong>reason:</strong> no reason provided
          </div>
        `}
        ${n.summary ? `<div class="notification-summary">${n.summary}</div>` : ''}
      </div>
    `).join('');
  } catch (e) {
    list.innerHTML = `
      <div class="notifications-empty">
        <p>could not load notifications</p>
      </div>`;
  }
}

async function markNotificationsRead() {
  const unread = cachedNotifications.filter(n => !n.read);
  if (unread.length === 0) return;

  try {
    await fetch('/api/notifications/read', {
      method: 'POST',
      headers: clientHeaders({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ ids: unread.map(n => n.id) }),
    });
    // Update badge immediately
    updateBadge(0);
    // Mark local cache as read
    cachedNotifications.forEach(n => n.read = true);
    // Remove unread styling
    document.querySelectorAll('.notification-card.unread').forEach(el => {
      el.classList.remove('unread');
    });
  } catch (e) {
    // Silently fail
  }
}

// Load badge on page load and poll every 30 seconds
loadNotificationBadge();
setInterval(loadNotificationBadge, 30000);

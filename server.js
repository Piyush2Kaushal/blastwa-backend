const express = require('express');
const cors = require('cors');
const http = require('http');
const { Server } = require('socket.io');
const multer = require('multer');
const XLSX = require('xlsx');
const qrcode = require('qrcode');
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const path = require('path');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
  cors: { origin: '*', methods: ['GET', 'POST'] }
});

app.use(cors());
app.use(express.json());

// ─── Storage ─────────────────────────────────────────────────────────────────
const upload = multer({ dest: 'uploads/' });
const mediaUpload = multer({ dest: 'media/' });

// ─── State ────────────────────────────────────────────────────────────────────
let whatsappClient = null;
let connectionStatus = 'disconnected'; // disconnected | qr | connecting | connected
let currentQR = null;
let profileInfo = null;
let campaignState = {
  isRunning: false,
  isPaused: false,
  contacts: [],
  currentIndex: 0,
  sent: 0,
  failed: 0,
  logs: []
};
let campaignAbortController = null;

// ─── WhatsApp Client ──────────────────────────────────────────────────────────
function createClient() {
  const client = new Client({
    authStrategy: new LocalAuth({ clientId: 'whatsapp-saas' }),
    puppeteer: {
      headless: true,
      executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || '/usr/bin/chromium',
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--no-first-run',
        '--no-zygote',
        '--single-process',
        '--disable-extensions'
      ]
    }
  });

  client.on('qr', async (qr) => {
    connectionStatus = 'qr';
    currentQR = await qrcode.toDataURL(qr);
    io.emit('status', { status: 'qr', qr: currentQR });
  });

  client.on('loading_screen', (percent) => {
    connectionStatus = 'connecting';
    io.emit('status', { status: 'connecting', percent });
  });

  client.on('authenticated', () => {
    connectionStatus = 'connecting';
    io.emit('status', { status: 'connecting' });
  });

  client.on('ready', async () => {
    connectionStatus = 'connected';
    currentQR = null;
    try {
      const info = client.info;
      profileInfo = {
        name: info.pushname,
        number: info.wid.user,
        platform: info.platform
      };
    } catch (e) { profileInfo = null; }
    io.emit('status', { status: 'connected', profile: profileInfo });
  });

  client.on('disconnected', (reason) => {
    connectionStatus = 'disconnected';
    profileInfo = null;
    io.emit('status', { status: 'disconnected', reason });
  });

  client.on('auth_failure', () => {
    connectionStatus = 'disconnected';
    io.emit('status', { status: 'disconnected', reason: 'auth_failure' });
  });

  return client;
}

// ─── Routes ───────────────────────────────────────────────────────────────────

// Status
app.get('/api/status', (req, res) => {
  res.json({ status: connectionStatus, qr: currentQR, profile: profileInfo });
});

// Connect
app.post('/api/connect', async (req, res) => {
  if (whatsappClient && connectionStatus === 'connected') {
    return res.json({ success: true, message: 'Already connected' });
  }
  try {
    if (whatsappClient) {
      try { await whatsappClient.destroy(); } catch (e) {}
    }
    connectionStatus = 'connecting';
    io.emit('status', { status: 'connecting' });
    whatsappClient = createClient();
    whatsappClient.initialize();
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Disconnect
app.post('/api/disconnect', async (req, res) => {
  try {
    if (whatsappClient) {
      await whatsappClient.logout();
      await whatsappClient.destroy();
      whatsappClient = null;
    }
    connectionStatus = 'disconnected';
    profileInfo = null;
    io.emit('status', { status: 'disconnected' });
    res.json({ success: true });
  } catch (err) {
    connectionStatus = 'disconnected';
    whatsappClient = null;
    res.json({ success: true });
  }
});

// Upload Excel
app.post('/api/upload-excel', upload.single('file'), (req, res) => {
  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const columns = data.length > 0 ? Object.keys(data[0]) : [];
    fs.unlinkSync(req.file.path);
    res.json({ success: true, columns, rows: data.length, preview: data.slice(0, 3), data });
  } catch (err) {
    res.status(400).json({ success: false, error: 'Failed to parse file: ' + err.message });
  }
});

// Upload Media
app.post('/api/upload-media', mediaUpload.single('file'), (req, res) => {
  try {
    const filePath = req.file.path;
    const ext = path.extname(req.file.originalname).toLowerCase();
    const newPath = filePath + ext;
    fs.renameSync(filePath, newPath);
    res.json({ success: true, path: newPath, filename: req.file.originalname });
  } catch (err) {
    res.status(400).json({ success: false, error: err.message });
  }
});

// Campaign state
app.get('/api/campaign/state', (req, res) => {
  res.json(campaignState);
});

// Start campaign
app.post('/api/campaign/start', async (req, res) => {
  if (connectionStatus !== 'connected') {
    return res.status(400).json({ success: false, error: 'WhatsApp not connected' });
  }
  if (campaignState.isRunning) {
    return res.status(400).json({ success: false, error: 'Campaign already running' });
  }

  const { contacts, messageTemplate, fieldMapping, mediaPath, delay = 3000 } = req.body;
  
  campaignState = {
    isRunning: true,
    isPaused: false,
    contacts,
    currentIndex: 0,
    sent: 0,
    failed: 0,
    logs: [],
    total: contacts.length,
    startTime: Date.now()
  };

  campaignAbortController = { aborted: false };
  const controller = campaignAbortController;

  res.json({ success: true });

  // Run campaign async
  (async () => {
    for (let i = 0; i < contacts.length; i++) {
      if (controller.aborted) break;

      while (campaignState.isPaused && !controller.aborted) {
        await new Promise(r => setTimeout(r, 500));
      }
      if (controller.aborted) break;

      const contact = contacts[i];
      campaignState.currentIndex = i;

      const numberField = fieldMapping?.number || 'number';
      let rawNumber = String(contact[numberField] || '').replace(/\D/g, '');
      if (!rawNumber.startsWith('91') && rawNumber.length === 10) rawNumber = '91' + rawNumber;
      const number = rawNumber + '@c.us';

      let message = messageTemplate;
      Object.keys(contact).forEach(key => {
        message = message.replace(new RegExp(`{{${key}}}`, 'g'), contact[key] || '');
      });

      const logEntry = {
        id: Date.now() + i,
        number: rawNumber,
        name: contact[fieldMapping?.name || 'name'] || rawNumber,
        message: message.substring(0, 60) + '...',
        status: 'sending',
        timestamp: new Date().toISOString()
      };

      campaignState.logs.unshift(logEntry);
      io.emit('campaign:progress', { ...campaignState, log: logEntry });

      try {
        if (mediaPath && fs.existsSync(mediaPath)) {
          const media = MessageMedia.fromFilePath(mediaPath);
          await whatsappClient.sendMessage(number, media, { caption: message });
        } else {
          await whatsappClient.sendMessage(number, message);
        }
        logEntry.status = 'sent';
        campaignState.sent++;
      } catch (err) {
        logEntry.status = 'failed';
        logEntry.error = err.message;
        campaignState.failed++;
      }

      io.emit('campaign:progress', { ...campaignState, log: logEntry });
      
      if (i < contacts.length - 1 && !controller.aborted) {
        await new Promise(r => setTimeout(r, delay));
      }
    }

    campaignState.isRunning = false;
    campaignState.isPaused = false;
    io.emit('campaign:complete', campaignState);
  })();
});

// Pause
app.post('/api/campaign/pause', (req, res) => {
  campaignState.isPaused = true;
  io.emit('campaign:paused', campaignState);
  res.json({ success: true });
});

// Resume
app.post('/api/campaign/resume', (req, res) => {
  campaignState.isPaused = false;
  io.emit('campaign:resumed', campaignState);
  res.json({ success: true });
});

// Stop
app.post('/api/campaign/stop', (req, res) => {
  if (campaignAbortController) campaignAbortController.aborted = true;
  campaignState.isRunning = false;
  campaignState.isPaused = false;
  io.emit('campaign:stopped', campaignState);
  res.json({ success: true });
});

// ─── Socket.IO ────────────────────────────────────────────────────────────────
io.on('connection', (socket) => {
  socket.emit('status', { status: connectionStatus, qr: currentQR, profile: profileInfo });
  socket.emit('campaign:progress', campaignState);
});

// ─── Start ────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8080;
server.listen(PORT, () => {
  console.log(`🚀 WhatsApp SaaS Backend running on http://localhost:${PORT}`);
  // Auto-initialize client to restore session
  whatsappClient = createClient();
  whatsappClient.initialize();
});

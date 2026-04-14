const path = require('path');
const fs = require('fs');
const https = require('https');
const { app, BrowserWindow, Menu, shell, ipcMain } = require('electron');

const CURRENT_VERSION = app.getVersion();
const UPDATE_CHECK_URL = 'https://api.github.com/repos/leo10m2010/foto-carnet/releases/latest';

function checkForUpdates(win) {
  const req = https.request(UPDATE_CHECK_URL, {
    headers: { 'User-Agent': 'FotoCarnet-Updater', 'Accept': 'application/vnd.github+json' }
  }, (res) => {
    let data = '';
    res.on('data', chunk => { data += chunk; });
    res.on('end', () => {
      try {
        const release = JSON.parse(data);
        const latest = (release.tag_name || '').replace(/^v/, '');
        if (latest && latest !== CURRENT_VERSION && isNewer(latest, CURRENT_VERSION)) {
          win.webContents.send('update-available', { version: latest, url: release.html_url });
        }
      } catch (_) {}
    });
  });
  req.on('error', () => {});
  req.setTimeout(8000, () => req.destroy());
  req.end();
}

function isNewer(a, b) {
  const pa = a.split('.').map(Number);
  const pb = b.split('.').map(Number);
  for (let i = 0; i < 3; i++) {
    if ((pa[i] || 0) > (pb[i] || 0)) return true;
    if ((pa[i] || 0) < (pb[i] || 0)) return false;
  }
  return false;
}

function createMainWindow() {
  const win = new BrowserWindow({
    width: 1500,
    height: 980,
    minWidth: 1100,
    minHeight: 760,
    autoHideMenuBar: true,
    backgroundColor: '#0a0e1a',
    icon: path.join(__dirname, '..', 'electron', 'resources', 'icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  });

  // Allow print popup window created by window.open in the current app.
  win.webContents.setWindowOpenHandler(() => ({ action: 'allow' }));

  // Keep external links in the system browser if any are clicked.
  win.webContents.on('will-navigate', (event, url) => {
    if (url.startsWith('http://') || url.startsWith('https://')) {
      event.preventDefault();
      shell.openExternal(url).catch(() => {});
    }
  });

  win.loadFile(path.join(__dirname, '..', 'index.html'));

  // Check for updates 5 seconds after launch (non-blocking)
  win.webContents.once('did-finish-load', () => {
    setTimeout(() => checkForUpdates(win), 5000);
  });
}

// RENIEC/apisperu.com API query — runs in main process to avoid renderer CORS restrictions
ipcMain.handle('reniec-query', (_event, { dni, token }) => {
  return new Promise((resolve) => {
    const options = {
      hostname: 'dniruc.apisperu.com',
      path: `/api/v1/dni/${encodeURIComponent(dni)}?token=${encodeURIComponent(token)}`,
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        try { resolve({ ok: true, body: JSON.parse(data) }); }
        catch (_) { resolve({ ok: false, error: 'JSON inválido' }); }
      });
    });

    req.on('error', (err) => resolve({ ok: false, error: err.message }));
    req.setTimeout(8000, () => { req.destroy(); resolve({ ok: false, error: 'Timeout' }); });
    req.end();
  });
});

// Read a local file and return it as a base64 data URL (used for session restore)
ipcMain.handle('read-file-as-dataurl', (_event, filePath) => {
  try {
    const buffer = fs.readFileSync(filePath);
    const ext = path.extname(filePath).slice(1).toLowerCase();
    const mimeMap = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png',
                      gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' };
    const mime = mimeMap[ext] || 'image/jpeg';
    return { ok: true, dataUrl: `data:${mime};base64,${buffer.toString('base64')}` };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

ipcMain.on('check-for-updates', (event) => {
  const win = BrowserWindow.fromWebContents(event.sender);
  if (win) checkForUpdates(win);
});

app.whenReady().then(() => {
  Menu.setApplicationMenu(null);
  createMainWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

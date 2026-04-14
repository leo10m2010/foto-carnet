const path = require('path');
const fs = require('fs');
const https = require('https');
const { app, BrowserWindow, Menu, shell, ipcMain } = require('electron');

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

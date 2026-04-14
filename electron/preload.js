const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('desktopMeta', {
  platform: process.platform,
  isElectron: true
});

contextBridge.exposeInMainWorld('electronAPI', {
  queryRENIEC:       (dni, token) => ipcRenderer.invoke('reniec-query', { dni, token }),
  readFileAsDataURL: (filePath)   => ipcRenderer.invoke('read-file-as-dataurl', filePath),
  getPathForFile:    (file)       => { try { return webUtils.getPathForFile(file) || ''; } catch (_) { return ''; } },
  onUpdateAvailable: (cb)        => ipcRenderer.on('update-available', (_e, info) => cb(info)),
  checkForUpdates:   ()          => ipcRenderer.send('check-for-updates')
});

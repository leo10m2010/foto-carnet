// ===================== INIT =====================

document.addEventListener('DOMContentLoaded', async () => {
    setupFileHandlers();
    setupLivePreview();
    setupCanvasDrag();
    initializeEditorState();
    setupHistoryControls();
    setupKeyboardShortcuts();
    initFilmstrip();
    const savedMode = localStorage.getItem('carnet-ui-mode') || 'simple';
    setUIMode(savedMode);
    await restoreSession();
    setupUpdateBanner();
});

function manualCheckForUpdates() {
    if (!window.electronAPI?.checkForUpdates) return;
    const btn = document.getElementById('btn-check-updates');
    const banner = document.getElementById('update-banner');
    const bannerAlreadyVisible = banner?.style.display !== 'none';

    function resetBtn() {
        if (!btn) return;
        btn.disabled = false;
        btn.innerHTML = '<i data-lucide="refresh-cw" style="width:11px;height:11px;vertical-align:-1px;margin-right:3px;"></i>Buscar actualizaciones';
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i data-lucide="refresh-cw" style="width:11px;height:11px;vertical-align:-1px;margin-right:3px;"></i>Buscando…';
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    const result = window.electronAPI.checkForUpdates();

    if (result && typeof result.then === 'function') {
        // Tauri: Promise-based, resolves with true/false
        result.then(found => {
            resetBtn();
            if (!found && !bannerAlreadyVisible) showToast('Ya tienes la última versión instalada ✓', 'success');
        }).catch(resetBtn);
    } else {
        // Electron: fire-and-forget via IPC; check banner state after timeout
        setTimeout(() => {
            resetBtn();
            if (banner?.style.display === 'none' && !bannerAlreadyVisible) {
                showToast('Ya tienes la última versión instalada ✓', 'success');
            }
        }, 5000);
    }
}

function setupUpdateBanner() {
    if (!window.electronAPI?.onUpdateAvailable) return;
    window.electronAPI.onUpdateAvailable(({ version, url }) => {
        const banner = document.getElementById('update-banner');
        const text   = document.getElementById('update-banner-text');
        const link   = document.getElementById('update-banner-link');
        if (!banner || !text || !link) return;
        text.textContent = `Nueva versión v${version} disponible`;
        link.onclick = (e) => { e.preventDefault(); window.open(url, '_blank'); };
        banner.style.display = 'flex';
        if (typeof lucide !== 'undefined') lucide.createIcons();
    });
}

window.addEventListener('beforeunload', () => {
    revokePhotoObjectUrls();
});

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
});

window.addEventListener('beforeunload', () => {
    revokePhotoObjectUrls();
});

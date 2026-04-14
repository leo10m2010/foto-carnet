// ===================== NAVIGATION =====================

function navigateRecord(delta) {
    if (state.records.length === 0) return;
    if (state.inlineEditor.active) closeInlineEditor({ commit: true });
    if (state.photoColorPicker.active) stopPhotoColorPickMode();
    if (state.photoCropMode.active) setPhotoCropMode(false);

    state.currentIndex += delta;
    if (state.currentIndex < 0) state.currentIndex = 0;
    if (state.currentIndex >= state.records.length) state.currentIndex = state.records.length - 1;

    updatePhotoInputsForCurrentRecord(); // Sync DOM for this record's photo config
    showSidebarNameEditor();
    tryRender();
    updateNavigation();
}

function clearAll() {
    // Abort any running RENIEC query
    state.reniecGeneration++;

    // Close any active overlays before resetting state
    if (state.inlineEditor.active) closeInlineEditor({ commit: false });
    if (state.photoColorPicker.active) stopPhotoColorPickMode();
    if (state.photoCropMode.active) setPhotoCropMode(false);

    // Revoke all photo object URLs to free browser memory
    revokePhotoObjectUrls();

    // Reset all state
    state.templateImage       = null;
    state.templateFileName    = '';
    state.templatePath        = null;
    state.templateDataUrl     = null;
    state.photoPaths          = {};
    clearTimeout(_saveSessionTimer); // Cancel any pending debounced save before clearing
    _saveSessionTimer = null;
    localStorage.removeItem(SESSION_KEY);
    state.records             = [];
    state.photosMap           = {};
    state.photoImageCache.clear();
    state.photoFaceBoxes      = {};
    state.photosCount         = 0;
    state.csvData             = null;
    state.csvRows             = [];
    state.photoOverrides      = {};
    state.globalPhotoConfig   = null;
    state.currentIndex        = 0;
    state.preflightReport     = null;
    state.history.undoStack   = [];
    state.history.redoStack   = [];
    state.history.lastSignature = '';
    state.drag.selectedId       = null;
    state.drag.active           = false;
    state.drag.photoPanActive   = false;
    state.drag.resizeHandle     = null;
    state.drag.elementId        = null;
    state.drag.snapGuides       = null;
    state.drag.hoveredId        = null;
    state.drag.historyCaptured  = false;
    state.hitboxes              = [];
    resetZoom();
    invalidatePreflightReport();

    // Reset file inputs so the same files can be re-selected
    ['input-template', 'input-photos-files', 'input-photos-folder', 'input-data'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });

    // Reset upload zone visuals
    const zoneTemplate = document.getElementById('zone-template');
    if (zoneTemplate) zoneTemplate.classList.remove('has-file');
    document.getElementById('template-file-name').textContent = '';
    document.getElementById('badge-template').classList.remove('completed');
    document.getElementById('badge-template').textContent = '1';

    const zonePhotos = document.getElementById('zone-photos');
    if (zonePhotos) zonePhotos.classList.remove('has-file');
    document.getElementById('photos-file-name').textContent = '';
    document.getElementById('badge-photos').classList.remove('completed');
    document.getElementById('badge-photos').textContent = '2';

    const zoneData = document.getElementById('zone-data');
    if (zoneData) zoneData.classList.remove('has-file');
    document.getElementById('data-file-name').textContent = '';

    // Hide data preview, RENIEC chip and column mapping
    document.getElementById('data-preview').style.display = 'none';
    document.getElementById('stat-records').textContent = '0';
    document.getElementById('stat-photos').textContent = '0';
    const chipReniec = document.getElementById('chip-reniec');
    if (chipReniec) chipReniec.style.display = 'none';
    document.getElementById('column-mapping').style.display = 'none';
    document.getElementById('preflight-report').style.display = 'none';

    // Hide canvas, show placeholder
    const canvas = document.getElementById('carnet-canvas');
    if (canvas) { canvas.style.display = 'none'; canvas.width = 0; canvas.height = 0; }
    const placeholder = document.getElementById('preview-placeholder');
    if (placeholder) placeholder.style.display = '';

    // Reset status bar
    document.getElementById('status-template').textContent  = 'Plantilla: —';
    document.getElementById('status-dimensions').textContent = '—';
    document.getElementById('status-text').textContent       = 'Sin datos cargados';
    const dot = document.getElementById('status-dot');
    if (dot) dot.className = 'status-dot';

    // Reset HUD and history buttons
    updateEditorHud();
    updateHistoryButtons();
    updateNavigation();
    renderFilmstrip();

    // Restore field defaults (positions, sizes, colors)
    initializeEditorState();

    showToast('Sesión limpiada. Puedes empezar de nuevo.', 'info');
}

function updateNavigation() {
    const total = state.records.length;
    const current = total > 0 ? state.currentIndex + 1 : 0;

    document.getElementById('current-index').textContent = current;
    document.getElementById('total-records').textContent = total;
    document.getElementById('btn-prev').disabled = state.currentIndex <= 0;
    document.getElementById('btn-next').disabled = state.currentIndex >= total - 1;

    const hasData = total > 0 && state.templateImage;
    document.getElementById('btn-export-png').disabled = !hasData;
    document.getElementById('btn-export-zip').disabled = !hasData;
    document.getElementById('btn-export-pdf').disabled = !hasData;
    document.getElementById('btn-print').disabled = !hasData;
    if (!hasData) {
        renderPreflightReport(null);
    }
    updateHistoryButtons();
    updateEditorHud();
    updateFilmstripActive();
    scrollFilmstripToActive();
}

function updateStatusBar() {
    const dot = document.getElementById('status-dot');
    const text = document.getElementById('status-text');

    if (state.records.length > 0) {
        dot.classList.add('active');
        text.textContent = `${state.records.length} registros listos`;
    } else {
        dot.classList.remove('active');
        text.textContent = 'Sin datos cargados';
    }
}

// ===================== ZOOM =====================

function changeZoom(delta) {
    state.zoom = Math.max(0.2, Math.min(3, state.zoom + delta));
    document.getElementById('zoom-level').textContent = Math.round(state.zoom * 100) + '%';
    const canvas = document.getElementById('carnet-canvas');
    canvas.style.transform = `scale(${state.zoom})`;
}

function resetZoom() {
    state.zoom = 1;
    document.getElementById('zoom-level').textContent = '100%';
    const canvas = document.getElementById('carnet-canvas');
    canvas.style.transform = 'scale(1)';
}

// ===================== SECTION TOGGLE =====================

function toggleSection(sectionId) {
    const section = document.getElementById(sectionId);
    section.classList.toggle('collapsed');
}


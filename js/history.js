function initializeEditorState() {
    state.globalPhotoConfig = readPhotoConfigFromInputs();

    const editableFields = document.querySelectorAll('input[id^="field-"], select[id^="field-"]');
    editableFields.forEach(field => {
        if (field.type === 'checkbox') {
            state.defaultFieldValues[field.id] = { type: 'checkbox', checked: !!field.checked };
        } else {
            state.defaultFieldValues[field.id] = { type: 'value', value: field.value };
        }
    });

    updateEditorHud();
}

function setUIMode(mode = 'simple') {
    const normalized = mode === 'advanced' ? 'advanced' : 'simple';
    state.uiMode = normalized;

    document.body.classList.toggle('ui-mode-simple', normalized === 'simple');
    document.body.classList.toggle('ui-mode-advanced', normalized === 'advanced');

    const simpleBtn = document.getElementById('btn-mode-simple');
    const advancedBtn = document.getElementById('btn-mode-advanced');
    if (simpleBtn) simpleBtn.classList.toggle('is-active', normalized === 'simple');
    if (advancedBtn) advancedBtn.classList.toggle('is-active', normalized === 'advanced');

    try {
        localStorage.setItem('carnet-ui-mode', normalized);
    } catch (_) {}
}

function setupRangeHistoryListener(elemId, snapshotName, sessionProp) {
    const el = document.getElementById(elemId);
    if (!el) return;
    el.addEventListener('pointerdown', () => {
        pushUndoSnapshot(snapshotName);
        state.history[sessionProp] = Date.now() + 700;
    });
    el.addEventListener('change', () => {
        state.history[sessionProp] = 0;
    });
}

function setupHistoryControls() {
    setupRangeHistoryListener('hud-photo-zoom',     'photo-zoom-range',   'zoomSessionUntil');
    setupRangeHistoryListener('hud-photo-rotation', 'photo-rotate-range', 'rotationSessionUntil');
    updateHistoryButtons();
}

function readTrackedInputState() {
    const tracked = {};
    const ids = new Set();

    document.querySelectorAll('input[id^="field-"], select[id^="field-"]').forEach(el => ids.add(el.id));
    [
        'photo-individual-mode',
        'pdf-width-cm', 'pdf-height-cm', 'pdf-orientation', 'pdf-page-size',
        'pdf-margin', 'pdf-gap', 'pdf-cut-length', 'pdf-cut-guides', 'export-dpi',
        'map-dni', 'map-extra'
    ].forEach(id => ids.add(id));

    ids.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        if (el.type === 'checkbox') {
            tracked[id] = { t: 'c', v: !!el.checked };
        } else {
            tracked[id] = { t: 'v', v: String(el.value ?? '') };
        }
    });
    return tracked;
}

function applyTrackedInputState(values = {}) {
    Object.entries(values).forEach(([id, cfg]) => {
        const el = document.getElementById(id);
        if (!el || !cfg) return;
        if (cfg.t === 'c') {
            el.checked = !!cfg.v;
        } else {
            el.value = cfg.v;
        }
    });
}

function createHistorySnapshot() {
    return {
        records: state.records.map(r => ({ ...r })),
        photoOverrides: JSON.parse(JSON.stringify(state.photoOverrides || {})),
        globalPhotoConfig: state.globalPhotoConfig ? { ...state.globalPhotoConfig } : null,
        currentIndex: state.currentIndex,
        selectedId: state.drag.selectedId || null,
        inputValues: readTrackedInputState(),
        uiMode: state.uiMode
    };
}

function getSnapshotSignature(snapshot) {
    try {
        return JSON.stringify(snapshot);
    } catch (_) {
        return `${Date.now()}-${Math.random()}`;
    }
}

function pushUndoSnapshot(reason = 'edit') {
    if (state.history.suspend) return;

    const snap = createHistorySnapshot();
    const sig = getSnapshotSignature(snap);
    if (sig === state.history.lastSignature) return;

    if (state.history.undoStack.length >= state.history.maxSize) {
        state.history.undoStack.shift();
    }
    state.history.undoStack.push(snap);
    state.history.redoStack = [];
    state.history.lastSignature = sig;
    updateHistoryButtons();
}

function applyHistorySnapshot(snapshot) {
    if (!snapshot) return;
    state.history.suspend = true;
    try {
        state.records = Array.isArray(snapshot.records) ? snapshot.records.map(r => ({ ...r })) : [];
        state.photoOverrides = snapshot.photoOverrides ? JSON.parse(JSON.stringify(snapshot.photoOverrides)) : {};
        state.globalPhotoConfig = snapshot.globalPhotoConfig ? { ...snapshot.globalPhotoConfig } : null;
        applyTrackedInputState(snapshot.inputValues || {});
        state.currentIndex = clamp(toInt(snapshot.currentIndex, 0), 0, Math.max(0, state.records.length - 1));
        state.drag.selectedId = snapshot.selectedId || null;
        setUIMode(snapshot.uiMode || state.uiMode || 'simple');

        showDataPreview();
        updatePhotoInputsForCurrentRecord();
        showSidebarNameEditor();
        updateNavigation();
        updateStatusBar();
    } finally {
        state.history.suspend = false;
    }

    tryRender();
}

function undoEdit() {
    if (state.history.undoStack.length === 0) {
        showToast('No hay más acciones para deshacer', 'info');
        return;
    }

    const current = createHistorySnapshot();
    const previous = state.history.undoStack.pop();
    state.history.redoStack.push(current);
    applyHistorySnapshot(previous);
    state.history.lastSignature = getSnapshotSignature(previous);
    updateHistoryButtons();
}

function redoEdit() {
    if (state.history.redoStack.length === 0) {
        showToast('No hay acciones para rehacer', 'info');
        return;
    }

    const current = createHistorySnapshot();
    const next = state.history.redoStack.pop();
    state.history.undoStack.push(current);
    applyHistorySnapshot(next);
    state.history.lastSignature = getSnapshotSignature(next);
    updateHistoryButtons();
}

function updateHistoryButtons() {
    const undoBtn = document.getElementById('btn-undo');
    const redoBtn = document.getElementById('btn-redo');
    if (undoBtn) undoBtn.disabled = state.history.undoStack.length === 0;
    if (redoBtn) redoBtn.disabled = state.history.redoStack.length === 0;
}

function setupKeyboardShortcuts() {
    document.addEventListener('keydown', (e) => {
        const tag = e.target?.tagName;
        const isTypingTarget = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || e.target?.isContentEditable;
        if (isTypingTarget) return;

        if ((e.ctrlKey || e.metaKey) && e.key === '0') {
            e.preventDefault();
            resetZoom();
            return;
        }

        if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'z') {
            e.preventDefault();
            undoEdit();
            return;
        }

        if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) {
            e.preventDefault();
            redoEdit();
            return;
        }

        if (e.key === 'Escape' && state.photoColorPicker.active) {
            e.preventDefault();
            stopPhotoColorPickMode();
            showToast('Gotero cancelado', 'info');
            updateEditorHud();
            return;
        }

        if (e.key === 'Escape' && state.photoCropMode.active) {
            e.preventDefault();
            setPhotoCropMode(false);
            showToast('Modo reencuadre desactivado', 'info');
            return;
        }

        const step = e.shiftKey ? 10 : 1;
        if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
            // No element selected: ← → navigate between records
            if (!state.drag.selectedId) {
                if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
                    e.preventDefault();
                    navigateRecord(e.key === 'ArrowRight' ? 1 : -1);
                }
                return;
            }
            e.preventDefault();
            const isPhotoPan = state.drag.selectedId === 'photo' && (state.photoCropMode.active || e.altKey || e.ctrlKey || e.metaKey);
            const panStep = e.shiftKey ? 20 : 6;

            if (isPhotoPan) {
                if (e.key === 'ArrowUp') panSelectedPhoto(0, -panStep);
                if (e.key === 'ArrowDown') panSelectedPhoto(0, panStep);
                if (e.key === 'ArrowLeft') panSelectedPhoto(-panStep, 0);
                if (e.key === 'ArrowRight') panSelectedPhoto(panStep, 0);
                return;
            }

            if (e.key === 'ArrowUp') nudgeSelectedElement(0, -step);
            if (e.key === 'ArrowDown') nudgeSelectedElement(0, step);
            if (e.key === 'ArrowLeft') nudgeSelectedElement(-step, 0);
            if (e.key === 'ArrowRight') nudgeSelectedElement(step, 0);
            return;
        }

        if (e.key.toLowerCase() === 'x') {
            e.preventDefault();
            alignSelectedElement('x');
            return;
        }

        if (e.key.toLowerCase() === 'y') {
            e.preventDefault();
            alignSelectedElement('y');
            return;
        }

        if (e.key.toLowerCase() === 'r') {
            e.preventDefault();
            resetSelectedElement();
            return;
        }

        if (e.key === 'Enter' && state.drag.selectedId) {
            e.preventDefault();
            startInlineTextEditFromSelection();
            return;
        }

        // Photo-only shortcuts — require photo to be selected
        if (state.drag.selectedId !== 'photo') return;

        if (e.key.toLowerCase() === 'i') {
            e.preventDefault();
            const checkbox = document.getElementById('hud-photo-individual');
            if (checkbox && !checkbox.disabled) togglePhotoIndividualFromHud(!checkbox.checked);
            return;
        }

        if (e.key.toLowerCase() === 'g') {
            e.preventDefault();
            if (state.photoColorPicker.active) {
                stopPhotoColorPickMode();
            } else {
                startPhotoColorPick();
            }
            return;
        }

        if (e.key.toLowerCase() === 'c') {
            e.preventDefault();
            togglePhotoCropMode();
            return;
        }

    });
}

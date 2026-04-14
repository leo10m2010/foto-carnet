// ===================== LIVE PREVIEW & CONFIG STATE =====================

function setupLivePreview() {
    const allInputs = document.querySelectorAll('.section-body input, .section-body select');
    allInputs.forEach(input => {
        input.addEventListener('input', (e) => handleInputChange(e));
        input.addEventListener('change', (e) => handleInputChange(e));
    });
}

function handleInputChange(e) {
    if (!state.records.length) return;
    invalidatePreflightReport();
    saveSessionDebounced();
    const isPhotoInput = e.target.id.startsWith('field-photo-');
    const isIndividualCheckbox = e.target.id === 'photo-individual-mode' || e.target.id === 'hud-photo-individual';
    const shouldTrack = !state.history.suspend && !state.drag.active &&
        (e.type === 'change' || isIndividualCheckbox);
    if (shouldTrack) {
        pushUndoSnapshot(`input:${e.target.id}`);
    }
    const record = state.records[state.currentIndex];
    const recordKey = getRecordKey(record);

    // If a coordinate/size input changed directly via typing
    if (isPhotoInput && !isIndividualCheckbox) {
        savePhotoConfigFromDOM();
        syncHudPhotoControls(getPhotoConfig());
        updatePhotoSwatches();
    }

    // If they checked or unchecked the box
    if (isIndividualCheckbox) {
        const isIndividual = !!e.target.checked;
        setPhotoIndividualModeControlValue(isIndividual);
        if (isIndividual) {
            if (!state.photoOverrides[recordKey]) {
                if (!state.globalPhotoConfig) {
                    state.globalPhotoConfig = getPhotoConfig();
                }
                state.photoOverrides[recordKey] = { ...state.globalPhotoConfig };
            }
        } else {
            // Reverting to global for this record
            delete state.photoOverrides[recordKey];
        }
        updatePhotoInputsForCurrentRecord(); // Sync DOM
    }

    tryRender();
}

function setPhotoIndividualModeControlValue(enabled) {
    const sidebar = document.getElementById('photo-individual-mode');
    const hud = document.getElementById('hud-photo-individual');
    if (sidebar) sidebar.checked = !!enabled;
    if (hud) hud.checked = !!enabled;
}

function syncHudPhotoControls(config) {
    const normalized = normalizePhotoConfig(config);
    const hudIndividual = document.getElementById('hud-photo-individual');
    const hudBgEnable = document.getElementById('hud-photo-bg-enable');
    const hudBgColor = document.getElementById('hud-photo-bg-color');
    const hudZoom = document.getElementById('hud-photo-zoom');
    const hudZoomValue = document.getElementById('hud-photo-zoom-value');
    const fitCover = document.getElementById('hud-fit-cover');
    const fitContain = document.getElementById('hud-fit-contain');
    const cropBtn = document.getElementById('hud-crop-mode');
    const photoIndividual = document.getElementById('photo-individual-mode');
    const goteroBtn = document.getElementById('hud-gotero-btn');
    const rotationVal = document.getElementById('hud-rotation-value');

    if (hudIndividual && photoIndividual) hudIndividual.checked = !!photoIndividual.checked;
    if (hudBgEnable) hudBgEnable.checked = !!normalized.bgEnabled;
    if (hudBgColor) hudBgColor.value = normalized.bgColor;
    if (hudZoom) hudZoom.value = normalized.scale.toFixed(2);
    if (hudZoomValue) hudZoomValue.textContent = `${normalized.scale.toFixed(2)}x`;
    if (fitCover) fitCover.classList.toggle('is-active', normalized.fit === 'cover');
    if (fitContain) fitContain.classList.toggle('is-active', normalized.fit === 'contain');
    if (cropBtn) cropBtn.classList.toggle('is-active', state.photoCropMode.active);
    if (goteroBtn) goteroBtn.classList.toggle('is-active', state.photoColorPicker.active);
    const rot = normalized.rotation || 0;
    if (rotationVal) rotationVal.textContent = `${rot}°`;
    const hudRotSlider = document.getElementById('hud-photo-rotation');
    if (hudRotSlider) hudRotSlider.value = rot;
    const hud = document.getElementById('editor-hud');
    if (hud) {
        hud.classList.toggle('crop-mode', state.photoCropMode.active);
        hud.classList.toggle('individual-mode', !!(hudIndividual && hudIndividual.checked));
    }
}

function updatePhotoInputsForCurrentRecord() {
    if (!state.records.length) return;
    const record = state.records[state.currentIndex];
    const recordKey = getRecordKey(record);
    const hasOverride = !!state.photoOverrides[recordKey];
    
    setPhotoIndividualModeControlValue(hasOverride);
    
    if (!state.globalPhotoConfig) {
        state.globalPhotoConfig = getPhotoConfig();
    }

    const baseConfig = state.globalPhotoConfig || readPhotoConfigFromInputs();
    const mergedConfig = hasOverride ? { ...baseConfig, ...state.photoOverrides[recordKey] } : baseConfig;
    const config = normalizePhotoConfig(mergedConfig);

    document.getElementById('field-photo-x').value = config.x;
    document.getElementById('field-photo-y').value = config.y;
    document.getElementById('field-photo-w').value = config.w;
    document.getElementById('field-photo-h').value = config.h;
    document.getElementById('field-photo-fit').value = config.fit;
    document.getElementById('field-photo-scale').value = config.scale.toFixed(2);
    document.getElementById('field-photo-offset-x').value = config.offsetX;
    document.getElementById('field-photo-offset-y').value = config.offsetY;
    document.getElementById('field-photo-bg-enable').checked = !!config.bgEnabled;
    document.getElementById('field-photo-bg-color').value = config.bgColor;
    const rotInput = document.getElementById('field-photo-rotation');
    if (rotInput) rotInput.value = config.rotation || 0;

    syncHudPhotoControls(config);

    updatePhotoSwatches();
    updateEditorHud();
}

function savePhotoConfigFromDOM() {
    if (!state.records.length) return;
    const isIndividual = !!document.getElementById('photo-individual-mode')?.checked;
    const record = state.records[state.currentIndex];
    const recordKey = getRecordKey(record);
    const config = normalizePhotoConfig(readPhotoConfigFromInputs());
    if (isIndividual) {
        state.photoOverrides[recordKey] = config;
    } else {
        state.globalPhotoConfig = config;
    }
}


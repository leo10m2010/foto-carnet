// ===================== SESSION PERSISTENCE =====================

const SESSION_KEY = 'fotocarnet_session_v2';
let _saveSessionTimer = null;

function saveSessionDebounced() {
    clearTimeout(_saveSessionTimer);
    _saveSessionTimer = setTimeout(saveSession, 2000);
}

function buildSessionData(includeDataUrl) {
    return {
        v: 2,
        savedAt: Date.now(),
        templatePath: state.templatePath || null,
        templateDataUrl: includeDataUrl ? (state.templateDataUrl || null) : null,
        templateFileName: state.templateFileName || '',
        templateW: state.templateImage?.width || 0,
        templateH: state.templateImage?.height || 0,
        photoPaths: state.photoPaths || {},
        records: state.records,
        photoOverrides: state.photoOverrides || {},
        globalPhotoConfig: state.globalPhotoConfig || null,
        currentIndex: state.currentIndex || 0,
        inputValues: readTrackedInputState(),
    };
}

async function saveSession() {
    if (!state.records.length) return;
    try {
        localStorage.setItem(SESSION_KEY, JSON.stringify(buildSessionData(true)));
    } catch (err) {
        if (err.name === 'QuotaExceededError' || err.code === 22) {
            try {
                localStorage.setItem(SESSION_KEY, JSON.stringify(buildSessionData(false)));
                console.warn('[Sesión] templateDataUrl omitida por cuota; se usará ruta de archivo.');
            } catch (_) {
                console.warn('[Sesión] Error al guardar incluso sin templateDataUrl:', _);
                showToast('No se pudo guardar la sesión: almacenamiento lleno. Exporta antes de cerrar.', 'error');
            }
        } else {
            console.warn('[Sesión] Error al guardar:', err);
            showToast('Error al guardar sesión automática.', 'error');
        }
    }
}

async function restoreSession() {
    try {
        const raw = localStorage.getItem(SESSION_KEY);
        if (!raw) return false;
        const data = JSON.parse(raw);
        if (!data.v || data.v < 2 || !Array.isArray(data.records) || !data.records.length) return false;

        // Discard sessions older than 7 days
        if (Date.now() - data.savedAt > 7 * 24 * 60 * 60 * 1000) {
            localStorage.removeItem(SESSION_KEY);
            return false;
        }

        // 1. Restore text data immediately (no file I/O)
        state.records        = data.records;
        state.photoOverrides = data.photoOverrides || {};
        state.globalPhotoConfig = data.globalPhotoConfig || null;
        state.currentIndex   = Math.min(data.currentIndex || 0, data.records.length - 1);
        state.templateFileName = data.templateFileName || '';
        state.photoPaths     = data.photoPaths || {};
        state.photosCount    = Object.keys(state.photoPaths).length;

        // Mark photos as "available via path" in photosMap so rendering works lazily
        state.photosMap = {};
        for (const [dniKey, filePath] of Object.entries(state.photoPaths)) {
            state.photosMap[dniKey] = filePath; // lazy: getPhotoImageByKey reads it on demand
        }

        // 2. Restore field values (positions, sizes, fonts…)
        if (data.inputValues) applyTrackedInputState(data.inputValues);

        // 3. Reload template image
        let templateOk = false;

        // 3a. Try dataURL saved in session (most reliable — no filesystem access needed)
        if (!templateOk && data.templateDataUrl) {
            await new Promise(resolve => {
                const img = new Image();
                img.onload = () => {
                    state.templateImage  = img;
                    state.templatePath   = data.templatePath || null;
                    state.templateDataUrl = data.templateDataUrl;
                    templateOk = true;
                    resolve();
                };
                img.onerror = resolve;
                img.src = data.templateDataUrl;
            });
        }

        // 3b. Fallback: re-read from disk via IPC (only if dataURL wasn't saved)
        if (!templateOk && data.templatePath && window.electronAPI?.readFileAsDataURL) {
            const result = await window.electronAPI.readFileAsDataURL(data.templatePath);
            if (result.ok) {
                await new Promise(resolve => {
                    const img = new Image();
                    img.onload = () => {
                        state.templateImage  = img;
                        state.templatePath   = data.templatePath;
                        state.templateDataUrl = result.dataUrl;
                        templateOk = true;
                        resolve();
                    };
                    img.onerror = resolve;
                    img.src = result.dataUrl;
                });
            }
        }

        // 4. Update UI badges / zones
        if (templateOk && state.templateImage) {
            const w = state.templateImage.width, h = state.templateImage.height;
            document.getElementById('zone-template')?.classList.add('has-file');
            document.getElementById('template-file-name').textContent = `✅ ${state.templateFileName} (${w}×${h})`;
            document.getElementById('badge-template')?.classList.add('completed');
            document.getElementById('badge-template').textContent = '✓';
            document.getElementById('status-template').textContent  = `Plantilla: ${state.templateFileName}`;
            document.getElementById('status-dimensions').textContent = `${w}×${h}px`;
        }

        if (state.records.length > 0) {
            const photoCount = Object.keys(state.photoPaths).length;
            document.getElementById('zone-photos')?.classList.add('has-file');
            document.getElementById('photos-file-name').textContent =
                `✅ ${photoCount} foto${photoCount !== 1 ? 's' : ''} (sesión restaurada)`;
            document.getElementById('badge-photos')?.classList.add('completed');
            document.getElementById('badge-photos').textContent = '✓';
        }

        // 5. Refresh all UI
        showDataPreview();
        document.getElementById('data-preview').style.display = 'block';
        updatePhotoInputsForCurrentRecord();
        updateNavigation();

        showSidebarNameEditor();
        if (templateOk) tryRender();
        renderFilmstrip();

        const mins = Math.round((Date.now() - data.savedAt) / 60000);
        const ageText = mins < 60 ? `${mins} min` : `${Math.round(mins / 60)}h`;
        showToast(
            `Sesión restaurada (${ageText} atrás) — ${state.records.length} registros, ${state.photosCount} fotos`,
            'info'
        );
        return true;
    } catch (err) {
        console.warn('[Sesión] Error al restaurar:', err);
        return false;
    }
}

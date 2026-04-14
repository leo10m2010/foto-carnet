/* =====================================================
   Generador Masivo de Carnets — App Logic
   ===================================================== */

// ===================== LRU IMAGE CACHE =====================
// Keeps at most MAX_SIZE decoded HTMLImageElements in memory.
// Evicts the least-recently-used entry when full.
class LRUImageCache {
    constructor(maxSize = 40) {
        this._map  = new Map();
        this._max  = maxSize;
    }
    get(key) {
        if (!this._map.has(key)) return undefined;
        // Promote to most-recently-used
        const val = this._map.get(key);
        this._map.delete(key);
        this._map.set(key, val);
        return val;
    }
    set(key, value) {
        if (this._map.has(key)) this._map.delete(key);
        this._map.set(key, value);
        if (this._map.size > this._max) {
            // Evict oldest entry (first key in insertion order)
            this._map.delete(this._map.keys().next().value);
        }
    }
    has(key)  { return this._map.has(key); }
    clear()   { this._map.clear(); }
}

// ===================== STATE =====================

const state = {
    templateImage: null,      // HTMLImageElement
    templateFileName: '',
    templatePath: null,       // Filesystem path (Electron only, for session restore)
    records: [],              // Array of { dni, nombres, apellidos, extra, hasPhoto }
    photosMap: {},            // { "07971267": objectURL/dataURL/filePath, ... }
    photoPaths: {},           // { "07971267": filePath } — for session restore
    photoObjectUrls: [],      // Temporary object URLs to revoke on reload
    photoImageCache: new LRUImageCache(40), // LRU — keeps last 40 decoded images in memory
    photosCount: 0,
    csvData: null,            // Optional CSV data keyed by DNI
    csvRows: [],              // Raw CSV rows for remapping
    photoOverrides: {},       // { [dni]: { x, y, w, h } }
    globalPhotoConfig: null,  // Default photo position/size for all records
    defaultFieldValues: {},   // Snapshot of original field values (for quick reset)
    currentIndex: 0,
    zoom: 1,
    renderTimer: null,        // Debounce timer for hover renders
    // Drag-and-drop state
    drag: {
        active: false,
        elementId: null,
        selectedId: null,      // Persistent selection
        resizeHandle: null,    // nw, ne, sw, se
        photoPanActive: false,
        startMouseX: 0,
        startMouseY: 0,
        startElemX: 0,
        startElemY: 0,
        startElemW: 0,
        startElemH: 0,
        startPhotoOffsetX: 0,
        startPhotoOffsetY: 0,
        startInputX: 0,
        startInputY: 0,
        snapGuides: null,
        hoveredId: null,
        historyCaptured: false
    },
    inlineEditor: {
        active: false,
        fieldId: null
    },
    photoColorPicker: {
        active: false
    },
    photoCropMode: {
        active: false
    },
    hitboxes: [],
    uiMode: 'simple',
    preflightReport: null,
    photoFaceBoxes: {},
    history: {
        undoStack: [],
        redoStack: [],
        maxSize: 60,
        suspend: false,
        lastSignature: '',
        zoomSessionUntil: 0,
        panSessionUntil: 0
    },
    job: {
        active: false,
        cancelRequested: false,
        label: ''
    },
    reniecGeneration: 0   // Incremented on every photo reload; aborts stale RENIEC queries
};

// ===================== SESSION PERSISTENCE =====================

const SESSION_KEY = 'fotocarnet_session_v2';
let _saveSessionTimer = null;

function saveSessionDebounced() {
    clearTimeout(_saveSessionTimer);
    _saveSessionTimer = setTimeout(saveSession, 2000);
}

async function saveSession() {
    if (!state.records.length) return;
    try {
        const data = {
            v: 2,
            savedAt: Date.now(),
            templatePath: state.templatePath || null,
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
        localStorage.setItem(SESSION_KEY, JSON.stringify(data));
    } catch (err) {
        console.warn('[Sesión] Error al guardar:', err);
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

        // 3. Reload template from disk via IPC
        let templateOk = false;
        if (data.templatePath && window.electronAPI?.readFileAsDataURL) {
            const result = await window.electronAPI.readFileAsDataURL(data.templatePath);
            if (result.ok) {
                await new Promise(resolve => {
                    const img = new Image();
                    img.onload = () => {
                        state.templateImage  = img;
                        state.templatePath   = data.templatePath;
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

        if (templateOk) tryRender();

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

// ===================== INIT =====================

document.addEventListener('DOMContentLoaded', async () => {
    setupFileHandlers();
    setupLivePreview();
    setupCanvasDrag();
    initializeEditorState();
    setupHistoryControls();
    setupKeyboardShortcuts();
    const savedMode = localStorage.getItem('carnet-ui-mode') || 'simple';
    setUIMode(savedMode);
    await restoreSession();
});

window.addEventListener('beforeunload', () => {
    revokePhotoObjectUrls();
});

// ===================== FILE HANDLERS =====================

function setupFileHandlers() {
    document.getElementById('input-template').addEventListener('change', handleTemplateUpload);
    document.getElementById('input-photos-files').addEventListener('change', handlePhotosUpload);
    document.getElementById('input-photos-folder').addEventListener('change', handlePhotosUpload);
    document.getElementById('input-data').addEventListener('change', handleDataUpload);

    // Drag-and-drop for upload zones
    ['zone-template', 'zone-photos', 'zone-data'].forEach(id => {
        const zone = document.getElementById(id);
        if (!zone) return;
        zone.addEventListener('dragover', e => {
            e.preventDefault();
            zone.style.borderColor = 'var(--accent-1)';
        });
        zone.addEventListener('dragleave', () => {
            zone.style.borderColor = '';
        });
        zone.addEventListener('drop', e => {
            e.preventDefault();
            zone.style.borderColor = '';
            const input = zone.querySelector('input[type="file"]');
            if (e.dataTransfer.files.length > 0) {
                input.files = e.dataTransfer.files;
                input.dispatchEvent(new Event('change'));
            }
        });
    });
}

const _scriptLoadCache = {};

function loadScriptOnce(key, src) {
    if (_scriptLoadCache[key]) return _scriptLoadCache[key];

    _scriptLoadCache[key] = new Promise((resolve, reject) => {
        const existing = document.querySelector(`script[data-lib="${key}"]`);
        if (existing) {
            if (existing.dataset.loaded === '1') {
                resolve();
                return;
            }
            existing.addEventListener('load', () => resolve(), { once: true });
            existing.addEventListener('error', () => reject(new Error(`No se pudo cargar ${key}`)), { once: true });
            return;
        }

        const script = document.createElement('script');
        script.src = src;
        script.async = true;
        script.dataset.lib = key;
        script.addEventListener('load', () => {
            script.dataset.loaded = '1';
            resolve();
        }, { once: true });
        script.addEventListener('error', () => reject(new Error(`No se pudo cargar ${key} desde ${src}`)), { once: true });
        document.head.appendChild(script);
    });

    return _scriptLoadCache[key];
}

async function ensureXLSX() {
    if (window.XLSX) return;
    await loadScriptOnce('xlsx', 'vendor/xlsx/xlsx.full.min.js');
    if (!window.XLSX) throw new Error('XLSX no disponible');
}

async function ensureJsPDF() {
    if (window.jspdf && window.jspdf.jsPDF) return;
    await loadScriptOnce('jspdf', 'vendor/jspdf/jspdf.umd.min.js');
    if (!(window.jspdf && window.jspdf.jsPDF)) throw new Error('jsPDF no disponible');
}

async function ensureJSZip() {
    if (window.JSZip) return;
    await loadScriptOnce('jszip', 'vendor/jszip/jszip.min.js');
    if (!window.JSZip) throw new Error('JSZip no disponible');
}

function toInt(value, fallback = 0) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : fallback;
}

function revokePhotoObjectUrls() {
    if (!Array.isArray(state.photoObjectUrls)) return;
    state.photoObjectUrls.forEach(url => {
        try { URL.revokeObjectURL(url); } catch (_) {}
    });
    state.photoObjectUrls = [];
}

function toFloat(value, fallback = 0) {
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : fallback;
}

function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
}

function normalizeHexColor(value, fallback = '#d9dee8') {
    const raw = String(value || '').trim();
    if (/^#[0-9A-F]{6}$/i.test(raw)) return raw.toLowerCase();
    if (/^#[0-9A-F]{3}$/i.test(raw)) {
        const short = raw.slice(1);
        return `#${short[0]}${short[0]}${short[1]}${short[1]}${short[2]}${short[2]}`.toLowerCase();
    }
    return fallback;
}

function rgbToHex(r, g, b) {
    const toHex = (n) => clamp(Math.round(n), 0, 255).toString(16).padStart(2, '0');
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function readPhotoConfigFromInputs() {
    return {
        x: Math.max(0, toInt(document.getElementById('field-photo-x')?.value, 0)),
        y: Math.max(0, toInt(document.getElementById('field-photo-y')?.value, 0)),
        w: Math.max(20, toInt(document.getElementById('field-photo-w')?.value, 20)),
        h: Math.max(20, toInt(document.getElementById('field-photo-h')?.value, 20)),
        fit: document.getElementById('field-photo-fit')?.value || 'cover',
        scale: clamp(toFloat(document.getElementById('field-photo-scale')?.value, 1), 0.2, 5),
        offsetX: toInt(document.getElementById('field-photo-offset-x')?.value, 0),
        offsetY: toInt(document.getElementById('field-photo-offset-y')?.value, 0),
        bgEnabled: !!document.getElementById('field-photo-bg-enable')?.checked,
        bgColor: normalizeHexColor(document.getElementById('field-photo-bg-color')?.value, '#d9dee8')
    };
}

function normalizePhotoConfig(config = {}) {
    const fitValue = config.fit === 'contain' ? 'contain' : 'cover';
    return {
        x: Math.max(0, toInt(config.x, 0)),
        y: Math.max(0, toInt(config.y, 0)),
        w: Math.max(20, toInt(config.w, 20)),
        h: Math.max(20, toInt(config.h, 20)),
        fit: fitValue,
        scale: clamp(toFloat(config.scale, 1), 0.2, 5),
        offsetX: toInt(config.offsetX, 0),
        offsetY: toInt(config.offsetY, 0),
        bgEnabled: !!config.bgEnabled,
        bgColor: normalizeHexColor(config.bgColor, '#d9dee8')
    };
}

function normalizeDNI(value) {
    const raw = String(value ?? '').trim();
    if (!raw) return '';

    const digits = raw.replace(/\D/g, '');
    if (!digits) return raw.toUpperCase();
    if (digits.length <= 8) return digits.padStart(8, '0');
    return digits;
}

function getRecordKey(record) {
    if (!record) return '';
    return record.dniKey || normalizeDNI(record.dni);
}

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

function setupHistoryControls() {
    const zoomRange = document.getElementById('hud-photo-zoom');
    if (zoomRange) {
        zoomRange.addEventListener('pointerdown', () => {
            pushUndoSnapshot('photo-zoom-range');
            state.history.zoomSessionUntil = Date.now() + 700;
        });
        zoomRange.addEventListener('change', () => {
            state.history.zoomSessionUntil = 0;
        });
    }
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

    state.history.undoStack.push(snap);
    if (state.history.undoStack.length > state.history.maxSize) {
        state.history.undoStack.shift();
    }
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
        }
    });
}

// ---- Template ----

function handleTemplateUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (ev) => {
        const img = new Image();
        img.onload = () => {
            state.templateImage = img;
            state.templateFileName = file.name;
            state.templatePath = file.path || null; // Electron: absolute path for session restore

            document.getElementById('zone-template').classList.add('has-file');
            document.getElementById('template-file-name').textContent = `✅ ${file.name} (${img.width}×${img.height})`;
            document.getElementById('badge-template').classList.add('completed');
            document.getElementById('badge-template').textContent = '✓';

            document.getElementById('status-template').textContent = `Plantilla: ${file.name}`;
            document.getElementById('status-dimensions').textContent = `${img.width}×${img.height}px`;

            showToast('Plantilla cargada correctamente', 'success');
            saveSessionDebounced();
            tryRender();
        };
        img.src = ev.target.result;
    };
    reader.readAsDataURL(file);
}

// ---- Photos (PRIMARY data source) ----

function parsePhotoFilename(filename) {
    // Strip ALL trailing image extensions (handles doubles like ".jpg.jpg")
    const baseName = filename.replace(/(\.(jpg|jpeg|png|gif|bmp|webp))+$/i, '').trim()
                             .replace(/_/g, ' '); // normalize underscores to spaces

    // Helper: split a "APELLIDOS NOMBRES" text block by Peruvian convention
    // (2 apellidos + 1-2 nombres)
    function splitApellidosNombres(text) {
        const words = text.trim().split(/\s+/);
        if (words.length >= 4) return { apellidos: words.slice(0, 2).join(' '), nombres: words.slice(2).join(' ') };
        if (words.length === 3) return { apellidos: words.slice(0, 2).join(' '), nombres: words[2] };
        if (words.length === 2) return { apellidos: words[0], nombres: words[1] };
        return { apellidos: text.trim(), nombres: '' };
    }

    let match;

    // 1. Standard: "12345678 - APELLIDOS NOMBRES"  (most common)
    match = baseName.match(/^(\d+)\s*[-–]\s*(.+)$/);
    if (match) {
        const dni = match[1].trim();
        return { dni, dniKey: normalizeDNI(dni), ...splitApellidosNombres(match[2]) };
    }

    // 2. Reversed: "APELLIDOS NOMBRES - 12345678"
    match = baseName.match(/^(.+)\s*[-–]\s*(\d+)$/);
    if (match) {
        const dni = match[2].trim();
        return { dni, dniKey: normalizeDNI(dni), ...splitApellidosNombres(match[1]) };
    }

    // 3. DNI as first token, space separator: "12345678 APELLIDOS NOMBRES"
    match = baseName.match(/^(\d{6,12})\s+(.+)$/);
    if (match) {
        const dni = match[1].trim();
        return { dni, dniKey: normalizeDNI(dni), ...splitApellidosNombres(match[2]) };
    }

    // 4. DNI as last token, space separator: "APELLIDOS NOMBRES 12345678"
    match = baseName.match(/^(.+)\s+(\d{6,12})$/);
    if (match) {
        const dni = match[2].trim();
        return { dni, dniKey: normalizeDNI(dni), ...splitApellidosNombres(match[1]) };
    }

    // 5. Only digits — bare DNI with no name
    if (/^\d{6,12}$/.test(baseName)) {
        return { dni: baseName, dniKey: normalizeDNI(baseName), nombres: '', apellidos: '' };
    }

    // 6. Fallback: no DNI found — treat as pure name file, parse nombres/apellidos
    //    The baseName becomes the DNI key so it still appears in the table,
    //    but at least nombres/apellidos are populated correctly.
    return { dni: baseName, dniKey: normalizeDNI(baseName), ...splitApellidosNombres(baseName) };
}

function handlePhotosUpload(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    state.globalPhotoConfig = readPhotoConfigFromInputs();

    revokePhotoObjectUrls();
    state.reniecGeneration++;          // Invalidates any in-progress RENIEC query
    state.photosMap = {};
    state.photoPaths = {};
    state.photoImageCache.clear();
    state.photoFaceBoxes = {};
    state.photoOverrides = {};
    state.photosCount = 0;
    state.records = [];
    invalidatePreflightReport();

    // Filter images: check extension OR MIME type
    const imageFiles = files.filter(f => {
        // Skip hidden / system files
        if (f.name.startsWith('.') || f.name === 'Thumbs.db' || f.name === 'desktop.ini') return false;
        // Check extension
        if (/\.(jpg|jpeg|png|gif|bmp|webp)$/i.test(f.name)) return true;
        // Fallback: check MIME type
        if (f.type && f.type.startsWith('image/')) return true;
        return false;
    });

    console.log(`[Fotos] Total archivos en carpeta: ${files.length}, Imágenes detectadas: ${imageFiles.length}`);
    if (imageFiles.length === 0) {
        // Log sample filenames for debugging
        const sampleNames = files.slice(0, 5).map(f => `"${f.name}" (type: ${f.type || 'N/A'})`);
        console.log('[Fotos] Archivos encontrados:', sampleNames);
        showToast(`No se encontraron imágenes. ${files.length} archivos en la carpeta. Revisa la consola (F12) para más detalles.`, 'error');
        return;
    }

    const parsedRecords = [];

    imageFiles.forEach(file => {
        // Parse filename for data
        const parsed = parsePhotoFilename(file.name);
        const dniKey = parsed.dniKey || normalizeDNI(parsed.dni);
        const objectUrl = URL.createObjectURL(file);

        // If duplicated DNI appears, keep the latest file and release old URL immediately
        if (state.photosMap[dniKey]) {
            try { URL.revokeObjectURL(state.photosMap[dniKey]); } catch (_) {}
        }

        state.photosMap[dniKey] = objectUrl;
        state.photoObjectUrls.push(objectUrl);
        if (file.path) state.photoPaths[dniKey] = file.path; // Electron path for session restore
        state.photosCount++;

        parsedRecords.push({
            dni: parsed.dni,
            dniKey,
            nombres: parsed.nombres,
            apellidos: parsed.apellidos,
            extra: '',    // Will be filled from CSV if available
            hasPhoto: true
        });
    });

    // Sort by DNI for consistent ordering
    parsedRecords.sort((a, b) => (a.dniKey || '').localeCompare(b.dniKey || '') || a.dni.localeCompare(b.dni));
    state.records = parsedRecords;
    state.currentIndex = 0;

    // If CSV data exists, merge it
    if (Array.isArray(state.csvRows) && state.csvRows.length > 0) {
        mergeCSVData();
    }

    document.getElementById('zone-photos').classList.add('has-file');
    document.getElementById('photos-file-name').textContent = `✅ ${imageFiles.length} fotos cargadas (datos extraídos)`;
    document.getElementById('badge-photos').classList.add('completed');
    document.getElementById('badge-photos').textContent = '✓';

    showDataPreview();
    document.getElementById('data-preview').style.display = 'block';
    updatePhotoInputsForCurrentRecord();

    updateNavigation();
    updateStatusBar();
    state.history.undoStack = [];
    state.history.redoStack = [];
    state.history.lastSignature = getSnapshotSignature(createHistorySnapshot());
    updateHistoryButtons();
    showToast(`${imageFiles.length} registros extraídos de las fotos`, 'success');
    saveSessionDebounced();
    tryRender();

    // Auto-query RENIEC in background (no UI controls shown)
    enrichWithRENIEC();
}

// ---- RENIEC API enrichment (runs automatically, no UI controls) ----

const RENIEC_TOKEN = '43861d5404ff0ab128d5ea8f0aefe52b2f5edba382cdf05b1f061316fbf8';

async function enrichWithRENIEC() {
    // Capture the generation token at the moment this query starts.
    // If the user reloads photos, state.reniecGeneration increments and
    // every check below will abort — preventing stale data from being written.
    const myGeneration = state.reniecGeneration;
    const isStale = () => state.reniecGeneration !== myGeneration;

    const toEnrich = state.records.filter(r => /^\d{8}$/.test(r.dniKey || r.dni) && !r.reniecOk);
    if (toEnrich.length === 0) return;

    // Build a fast lookup: dniKey -> index in state.records (avoids O(n²) findIndex)
    const dniIndexMap = new Map(state.records.map((r, i) => [r.dniKey || r.dni, i]));

    showToast(`Verificando ${toEnrich.length} DNI${toEnrich.length > 1 ? 's' : ''} en RENIEC…`, 'info');
    updateReniecStatChip('…');

    let ok = 0, notFound = 0, errors = 0;

    for (let i = 0; i < toEnrich.length; i++) {
        // Abort immediately if the user loaded a new set of photos
        if (isStale()) return;

        const record = toEnrich[i];
        const dni = record.dniKey || record.dni;
        // Resolve index outside try/catch so it's accessible in both blocks
        const idx = dniIndexMap.get(dni);
        if (idx === undefined) continue;

        try {
            let json;

            if (window.electronAPI?.queryRENIEC) {
                const result = await window.electronAPI.queryRENIEC(dni, RENIEC_TOKEN);
                if (!result.ok) throw new Error(result.error);
                json = result.body;
            } else {
                const resp = await fetch('https://api.json.pe/api/dni', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${RENIEC_TOKEN}`
                    },
                    body: JSON.stringify({ dni })
                });
                json = await resp.json();
            }

            // Check again after the await — a reload could have happened during the request
            if (isStale()) return;

            if (json.success && json.data) {
                const d = json.data;
                const nombres   = (d.nombres || '').trim();
                const apellidos = `${(d.apellido_paterno || '')} ${(d.apellido_materno || '')}`.trim();
                if (!state.records[idx].filenameNombres)   state.records[idx].filenameNombres   = state.records[idx].nombres;
                if (!state.records[idx].filenameApellidos) state.records[idx].filenameApellidos = state.records[idx].apellidos;
                if (nombres)   state.records[idx].nombres   = nombres;
                if (apellidos) state.records[idx].apellidos = apellidos;
                state.records[idx].reniecNombres   = nombres;
                state.records[idx].reniecApellidos = apellidos;
                state.records[idx].reniecOk = true;
                ok++;
            } else {
                state.records[idx].reniecOk = false;
                notFound++;
            }
        } catch (_) {
            if (isStale()) return;
            state.records[idx].reniecOk = false;
            errors++;
        }

        // Refresh table every 5 records to reduce repaints
        if (i % 5 === 4 || i === toEnrich.length - 1) {
            if (isStale()) return;
            showDataPreview();
            updateReniecStatChip(`${ok}/${toEnrich.length}`);
        }

        await new Promise(r => setTimeout(r, 200)); // ≈ 5 req/s
    }

    if (isStale()) return;

    showDataPreview();
    tryRender();
    updateReniecStatChip(`${ok}/${toEnrich.length}`);

    const corrected = state.records.filter(r => r.reniecOk && r.filenameNombres &&
        (r.filenameNombres.toUpperCase() !== r.reniecNombres?.toUpperCase() ||
         r.filenameApellidos?.toUpperCase() !== r.reniecApellidos?.toUpperCase())).length;

    let msg = `RENIEC: ${ok} verificados`;
    if (corrected > 0) msg += `, ${corrected} nombre${corrected > 1 ? 's' : ''} corregido${corrected > 1 ? 's' : ''}`;
    if (notFound > 0)  msg += `, ${notFound} no encontrado${notFound > 1 ? 's' : ''}`;
    if (errors > 0)    msg += `, ${errors} con error`;

    showToast(msg, ok > 0 ? 'success' : 'warning');
    saveSession(); // Persist RENIEC-enriched names
}

function updateReniecStatChip(text) {
    const chip = document.getElementById('chip-reniec');
    const el   = document.getElementById('stat-reniec');
    if (chip) chip.style.display = '';
    if (el)   el.textContent = text;
}

// ---- CSV / Excel (OPTIONAL — for extra fields like cargo) ----

function handleDataUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const isCSV = /\.csv$/i.test(file.name);

    const reader = new FileReader();
    reader.onload = async (ev) => {
        try {
            await ensureXLSX();
            let data;
            if (isCSV) {
                const workbook = XLSX.read(ev.target.result, { type: 'binary', codepage: 65001 });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            } else {
                const workbook = XLSX.read(ev.target.result, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            }

            if (data.length === 0) {
                showToast('El archivo no contiene datos', 'error');
                return;
            }

            // Store CSV data keyed by DNI
            const columns = Object.keys(data[0]);
            const dniCol = autoDetectDNIColumn(columns);
            const extraCol = autoDetectExtraColumn(columns);

            state.csvRows = data;
            state.csvData = dniCol ? buildCSVIndex(dniCol) : {};
            invalidatePreflightReport();

            // Show column mapping
            populateCSVMapping(columns, dniCol, extraCol);
            document.getElementById('column-mapping').style.display = 'block';

            document.getElementById('zone-data').classList.add('has-file');
            document.getElementById('data-file-name').textContent = `✅ ${file.name} (${data.length} registros)`;
            document.getElementById('badge-data').textContent = '✓';

            // If photos already loaded, merge
            if (state.records.length > 0) {
                mergeCSVData();
                showDataPreview();
                tryRender();
            }

            showToast(`CSV cargado: ${data.length} registros. Se vincularán por DNI.`, 'success');
        } catch (err) {
            showToast('Error al leer el archivo: ' + err.message, 'error');
            console.error(err);
        }
    };

    if (isCSV) {
        reader.readAsBinaryString(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

function autoDetectDNIColumn(columns) {
    const cols = columns.map(c => c.toLowerCase().trim());
    const keywords = ['dni', 'documento', 'cedula', 'ci', 'id', 'doc', 'num_doc', 'numero_documento', 'rut'];
    for (const kw of keywords) {
        const idx = cols.findIndex(c => c.includes(kw));
        if (idx !== -1) return columns[idx];
    }
    return columns[0]; // fallback to first column
}

function autoDetectExtraColumn(columns) {
    const cols = columns.map(c => c.toLowerCase().trim());
    const keywords = ['cargo', 'puesto', 'area', 'departamento', 'facultad', 'carrera', 'tipo', 'categoria', 'extra', 'condicion'];
    for (const kw of keywords) {
        const idx = cols.findIndex(c => c.includes(kw));
        if (idx !== -1) return columns[idx];
    }
    return '';
}

function populateCSVMapping(columns, defaultDni, defaultExtra) {
    const mapDni = document.getElementById('map-dni');
    const mapExtra = document.getElementById('map-extra');

    mapDni.innerHTML = '';
    mapExtra.innerHTML = '<option value="">— Ninguno —</option>';

    columns.forEach(col => {
        const opt1 = document.createElement('option');
        opt1.value = col;
        opt1.textContent = col;
        if (col === defaultDni) opt1.selected = true;
        mapDni.appendChild(opt1);

        const opt2 = document.createElement('option');
        opt2.value = col;
        opt2.textContent = col;
        if (col === defaultExtra) opt2.selected = true;
        mapExtra.appendChild(opt2);
    });

    mapDni.onchange = () => { remergeCSV(); };
    mapExtra.onchange = () => { remergeCSV(); };
}

function buildCSVIndex(dniColumn) {
    const index = {};
    if (!dniColumn || !Array.isArray(state.csvRows)) return index;

    state.csvRows.forEach(row => {
        const key = normalizeDNI(row[dniColumn]);
        if (!key) return;
        index[key] = row;
    });

    return index;
}

function remergeCSV() {
    if (!Array.isArray(state.csvRows) || state.csvRows.length === 0 || state.records.length === 0) return;
    mergeCSVData();
    showDataPreview();
    tryRender();
}

function mergeCSVData() {
    if (!Array.isArray(state.csvRows) || state.csvRows.length === 0) return;

    const dniCol = document.getElementById('map-dni')?.value || '';
    const extraCol = document.getElementById('map-extra')?.value || '';
    state.csvData = buildCSVIndex(dniCol);

    state.records.forEach(record => {
        record.extra = '';

        // Find matching CSV row by DNI
        const key = getRecordKey(record);
        const csvRow = state.csvData[key];
        if (csvRow && extraCol) {
            record.extra = String(csvRow[extraCol] || '').trim();
        }
    });
}

function showDataPreview() {
    const thead = document.querySelector('#data-table thead');
    const tbody = document.querySelector('#data-table tbody');

    thead.innerHTML = '<tr><th>DNI</th><th>Nombres</th><th>Apellidos</th><th>Extra</th><th>Foto</th><th title="Verificado en RENIEC">✓</th></tr>';

    const preview = state.records.slice(0, 20);
    tbody.innerHTML = preview.map(r => {
        let verif = '<td class="reniec-pending">…</td>';
        if (r.reniecOk === true) {
            // Did the filename match what RENIEC has?
            const fnNom = (r.filenameNombres   || '').toUpperCase().trim();
            const fnAp  = (r.filenameApellidos || '').toUpperCase().trim();
            const rnNom = (r.reniecNombres     || '').toUpperCase().trim();
            const rnAp  = (r.reniecApellidos   || '').toUpperCase().trim();
            const matched = fnNom === rnNom && fnAp === rnAp;
            const tip = matched
                ? 'Nombre del archivo coincide con RENIEC'
                : `Archivo: ${escapeHtml(r.filenameApellidos)} ${escapeHtml(r.filenameNombres)} → corregido con RENIEC`;
            verif = `<td class="reniec-ok" title="${tip}">${matched ? '✓' : '✓*'}</td>`;
        } else if (r.reniecOk === false) {
            verif = '<td class="reniec-err" title="No encontrado en RENIEC">✗</td>';
        }

        return `<tr>
            <td>${escapeHtml(r.dni)}</td>
            <td>${escapeHtml(r.nombres)}</td>
            <td>${escapeHtml(r.apellidos)}</td>
            <td>${escapeHtml(r.extra || '—')}</td>
            <td>${r.hasPhoto ? '✅' : '❌'}</td>
            ${verif}
        </tr>`;
    }).join('');

    document.getElementById('stat-records').textContent = state.records.length;
    document.getElementById('stat-photos').textContent = state.records.filter(r => r.hasPhoto).length + '/' + state.records.length;
}

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

    if (hudIndividual && photoIndividual) hudIndividual.checked = !!photoIndividual.checked;
    if (hudBgEnable) hudBgEnable.checked = !!normalized.bgEnabled;
    if (hudBgColor) hudBgColor.value = normalized.bgColor;
    if (hudZoom) hudZoom.value = normalized.scale.toFixed(2);
    if (hudZoomValue) hudZoomValue.textContent = `${normalized.scale.toFixed(2)}x`;
    if (fitCover) fitCover.classList.toggle('is-active', normalized.fit === 'cover');
    if (fitContain) fitContain.classList.toggle('is-active', normalized.fit === 'contain');
    if (cropBtn) cropBtn.classList.toggle('is-active', state.photoCropMode.active);
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

// ===================== RENDERING ENGINE =====================

function getFieldConfig(fieldName) {
    const get = (suffix, fallback) => {
        const el = document.getElementById(`field-${fieldName}-${suffix}`);
        return el ? el.value : fallback;
    };
    return {
        x: Math.max(0, toInt(get('x', 0), 0)),
        y: Math.max(0, toInt(get('y', 0), 0)),
        size: Math.max(6, toInt(get('size', 16), 16)),
        color: get('color', '#000000'),
        font: get('font', 'Poppins'),
        align: get('align', 'center'),
        bold: get('bold', ''),
        maxWidth: Math.max(50, toInt(get('maxw', 300), 300))
    };
}

function getPhotoConfig() {
    if (state.records.length > 0 && state.currentIndex < state.records.length) {
        const record = state.records[state.currentIndex];
        const recordKey = getRecordKey(record);
        if (recordKey && state.photoOverrides[recordKey]) {
            return { ...state.photoOverrides[recordKey] };
        }
    }

    if (state.globalPhotoConfig) {
        return normalizePhotoConfig({ ...state.globalPhotoConfig });
    }

    return normalizePhotoConfig(readPhotoConfigFromInputs());
}

function getPhotoConfigForRecord(record) {
    const key = getRecordKey(record);
    const override = key ? state.photoOverrides[key] : null;

    if (override) return normalizePhotoConfig({ ...override });
    if (state.globalPhotoConfig) return normalizePhotoConfig({ ...state.globalPhotoConfig });
    return getPhotoConfig();
}

function getBarcodeConfig() {
    return {
        x: Math.max(0, toInt(document.getElementById('field-barcode-x').value, 0)),
        y: Math.max(0, toInt(document.getElementById('field-barcode-y').value, 0)),
        w: Math.max(40, toInt(document.getElementById('field-barcode-w').value, 40)),
        h: Math.max(20, toInt(document.getElementById('field-barcode-h').value, 20)),
        format: document.getElementById('field-barcode-format').value,
        showText: document.getElementById('field-barcode-showtext').value === 'true'
    };
}

function tryRender() {
    if (!state.templateImage || state.records.length === 0) return;
    renderCarnet(state.currentIndex).then(() => {
        if (!state.drag.active) drawSelectionOverlay();
        updateEditorHud();
    });
}

function getCurrentRecord() {
    if (!state.records.length) return null;
    return state.records[state.currentIndex] || null;
}

function getCurrentPhotoImage() {
    const record = getCurrentRecord();
    if (!record) return Promise.resolve(null);
    const key = getRecordKey(record);
    return getPhotoImageByKey(key);
}

async function getPhotoImageByKey(key) {
    if (!key) return null;
    const fromCache = state.photoImageCache.get(key);
    if (fromCache) return fromCache;

    let source = state.photosMap[key];

    // Session restore: if photosMap only has a file path (not a blob/data URL),
    // read from disk via Electron IPC and create a proper object URL.
    if (source && !source.startsWith('blob:') && !source.startsWith('data:')
            && window.electronAPI?.readFileAsDataURL) {
        const result = await window.electronAPI.readFileAsDataURL(source);
        if (result.ok) {
            // Create an object URL from the data URL so memory is managed normally
            const resp = await fetch(result.dataUrl);
            const blob = await resp.blob();
            const objUrl = URL.createObjectURL(blob);
            state.photosMap[key] = objUrl;
            state.photoObjectUrls.push(objUrl);
            source = objUrl;
        } else {
            return null;
        }
    }

    if (!source) return null;

    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            state.photoImageCache.set(key, img);
            resolve(img);
        };
        img.onerror = () => resolve(null);
        img.src = source;
    });
}

async function detectPrimaryFace(photoImg, cacheKey = '') {
    if (!photoImg) return null;
    if (cacheKey && Object.prototype.hasOwnProperty.call(state.photoFaceBoxes, cacheKey)) {
        return state.photoFaceBoxes[cacheKey];
    }

    if (typeof FaceDetector === 'undefined') {
        if (cacheKey) state.photoFaceBoxes[cacheKey] = null;
        return null;
    }

    try {
        const detector = new FaceDetector({ fastMode: true, maxDetectedFaces: 1 });
        const faces = await detector.detect(photoImg);
        if (!faces || !faces.length || !faces[0].boundingBox) {
            if (cacheKey) state.photoFaceBoxes[cacheKey] = null;
            return null;
        }

        const box = faces[0].boundingBox;
        const normalized = {
            x: toFloat(box.x, 0),
            y: toFloat(box.y, 0),
            width: Math.max(1, toFloat(box.width, 1)),
            height: Math.max(1, toFloat(box.height, 1))
        };
        if (cacheKey) state.photoFaceBoxes[cacheKey] = normalized;
        return normalized;
    } catch (_) {
        if (cacheKey) state.photoFaceBoxes[cacheKey] = null;
        return null;
    }
}

function getPhotoDrawRect(photoImg, photoConfig) {
    const px = photoConfig.x;
    const py = photoConfig.y;
    const pw = photoConfig.w;
    const ph = photoConfig.h;

    const sourceW = photoImg.naturalWidth || photoImg.width;
    const sourceH = photoImg.naturalHeight || photoImg.height;
    if (!sourceW || !sourceH) return;

    const scaleX = pw / sourceW;
    const scaleY = ph / sourceH;
    const baseScale = photoConfig.fit === 'contain' ? Math.min(scaleX, scaleY) : Math.max(scaleX, scaleY);
    const finalScale = baseScale * photoConfig.scale;

    const drawW = sourceW * finalScale;
    const drawH = sourceH * finalScale;
    const drawX = px + (pw - drawW) / 2 + photoConfig.offsetX;
    const drawY = py + (ph - drawH) / 2 + photoConfig.offsetY;

    return {
        frameX: px,
        frameY: py,
        frameW: pw,
        frameH: ph,
        drawX,
        drawY,
        drawW,
        drawH,
        sourceW,
        sourceH
    };
}

function drawPhotoInFrame(ctx, photoImg, photoConfig) {
    const rect = getPhotoDrawRect(photoImg, photoConfig);
    if (!rect) return;

    ctx.save();
    ctx.beginPath();
    ctx.rect(rect.frameX, rect.frameY, rect.frameW, rect.frameH);
    ctx.clip();
    ctx.drawImage(photoImg, rect.drawX, rect.drawY, rect.drawW, rect.drawH);
    ctx.restore();
}

function samplePhotoPixel(photoImg, x, y) {
    const sourceW = photoImg.naturalWidth || photoImg.width;
    const sourceH = photoImg.naturalHeight || photoImg.height;
    if (!sourceW || !sourceH) return null;

    const off = document.createElement('canvas');
    off.width = sourceW;
    off.height = sourceH;
    const octx = off.getContext('2d', { willReadFrequently: true });
    octx.drawImage(photoImg, 0, 0, sourceW, sourceH);

    const sx = clamp(Math.floor(x), 0, sourceW - 1);
    const sy = clamp(Math.floor(y), 0, sourceH - 1);
    const rgba = octx.getImageData(sx, sy, 1, 1).data;
    return rgbToHex(rgba[0], rgba[1], rgba[2]);
}

function getPhotoColorFromCanvasPoint(mx, my, photoImg, photoConfig) {
    const rect = getPhotoDrawRect(photoImg, photoConfig);
    if (!rect) return null;

    const insideFrame = mx >= rect.frameX && mx <= rect.frameX + rect.frameW &&
        my >= rect.frameY && my <= rect.frameY + rect.frameH;
    if (!insideFrame) return null;

    const sourceX = ((mx - rect.drawX) / rect.drawW) * rect.sourceW;
    const sourceY = ((my - rect.drawY) / rect.drawH) * rect.sourceH;
    return samplePhotoPixel(photoImg, sourceX, sourceY);
}

function setPhotoBgColor(color) {
    pushUndoSnapshot('photo-bg-color');
    const normalized = normalizeHexColor(color, '#d9dee8');
    const colorInput = document.getElementById('field-photo-bg-color');
    const enabledInput = document.getElementById('field-photo-bg-enable');
    const hudColor = document.getElementById('hud-photo-bg-color');
    const hudEnabled = document.getElementById('hud-photo-bg-enable');

    if (colorInput) colorInput.value = normalized;
    if (hudColor) hudColor.value = normalized;
    if (enabledInput) enabledInput.checked = true;
    if (hudEnabled) hudEnabled.checked = true;

    invalidatePreflightReport();
    savePhotoConfigFromDOM();
    syncHudPhotoControls(getPhotoConfig());
    tryRender();
}

function togglePhotoBgFromHud(enabled) {
    pushUndoSnapshot('photo-bg-toggle');
    const input = document.getElementById('field-photo-bg-enable');
    if (!input) return;
    input.checked = !!enabled;
    invalidatePreflightReport();
    savePhotoConfigFromDOM();
    syncHudPhotoControls(getPhotoConfig());
    tryRender();
}

function stopPhotoColorPickMode() {
    state.photoColorPicker.active = false;
    const canvas = document.getElementById('carnet-canvas');
    if (canvas) canvas.style.cursor = 'default';
    updateEditorHud();
}

function startPhotoColorPick() {
    if (state.drag.selectedId !== 'photo') {
        state.drag.selectedId = 'photo';
        tryRender();
    }
    state.photoColorPicker.active = true;
    const canvas = document.getElementById('carnet-canvas');
    if (canvas) canvas.style.cursor = 'crosshair';
    showToast('Haz clic dentro de la foto para tomar un color', 'info');
    updateEditorHud();
}

async function autoPickPhotoBgColor() {
    const photoImg = await getCurrentPhotoImage();
    if (!photoImg) {
        showToast('No se pudo leer la foto actual para muestrear color', 'error');
        return;
    }

    const sourceW = photoImg.naturalWidth || photoImg.width;
    const sourceH = photoImg.naturalHeight || photoImg.height;
    if (!sourceW || !sourceH) return;

    const samplePoints = [
        [0.12, 0.10], [0.5, 0.08], [0.88, 0.10],
        [0.18, 0.22], [0.82, 0.22], [0.5, 0.18]
    ];

    let r = 0;
    let g = 0;
    let b = 0;
    samplePoints.forEach(([rx, ry]) => {
        const hex = samplePhotoPixel(photoImg, sourceW * rx, sourceH * ry);
        const c = normalizeHexColor(hex, '#d9dee8');
        r += Number.parseInt(c.slice(1, 3), 16);
        g += Number.parseInt(c.slice(3, 5), 16);
        b += Number.parseInt(c.slice(5, 7), 16);
    });

    const count = samplePoints.length;
    const picked = rgbToHex(r / count, g / count, b / count);
    setPhotoBgColor(picked);
    showToast('Color sugerido aplicado desde la foto', 'success');
}

async function autoFrameCurrentPhoto() {
    if (state.drag.selectedId !== 'photo') {
        state.drag.selectedId = 'photo';
    }

    const record = getCurrentRecord();
    if (!record) return;
    const key = getRecordKey(record);
    const photoImg = await getPhotoImageByKey(key);
    if (!photoImg) {
        showToast('No se pudo abrir la foto para auto-encuadre', 'error');
        return;
    }

    pushUndoSnapshot('photo-auto-frame');

    const cfg = getPhotoConfig();
    const sourceW = photoImg.naturalWidth || photoImg.width;
    const sourceH = photoImg.naturalHeight || photoImg.height;
    if (!sourceW || !sourceH) {
        showToast('La foto actual no tiene dimensiones válidas', 'error');
        return;
    }

    const fitInput = document.getElementById('field-photo-fit');
    const scaleInput = document.getElementById('field-photo-scale');
    const offsetXInput = document.getElementById('field-photo-offset-x');
    const offsetYInput = document.getElementById('field-photo-offset-y');
    const bgEnableInput = document.getElementById('field-photo-bg-enable');
    if (!fitInput || !scaleInput || !offsetXInput || !offsetYInput) return;

    fitInput.value = 'cover';

    const face = await detectPrimaryFace(photoImg, key);
    if (face) {
        const baseScale = Math.max(cfg.w / sourceW, cfg.h / sourceH);
        const targetFaceWidth = cfg.w * 0.38;
        const desiredFinalScale = clamp(targetFaceWidth / face.width, baseScale * 0.75, baseScale * 5);
        const scaleValue = clamp(desiredFinalScale / baseScale, 0.2, 5);

        const drawW = sourceW * baseScale * scaleValue;
        const drawH = sourceH * baseScale * scaleValue;
        const baseX = (cfg.w - drawW) / 2;
        const baseY = (cfg.h - drawH) / 2;
        const faceCenterX = (face.x + face.width / 2) * baseScale * scaleValue;
        const faceCenterY = (face.y + face.height / 2) * baseScale * scaleValue;

        const targetCenterX = cfg.w / 2;
        // Target: face center at 42% from top of the photo slot (lower = face more centered)
        const targetCenterY = cfg.h * 0.42;
        const offsetX = Math.round(targetCenterX - (baseX + faceCenterX));
        const offsetY = Math.round(targetCenterY - (baseY + faceCenterY));

        scaleInput.value = scaleValue.toFixed(2);
        offsetXInput.value = offsetX;
        offsetYInput.value = offsetY;
        showToast('Auto-encuadre de rostro aplicado', 'success');
    } else {
        // Fallback if face detector is unavailable or no face was detected.
        const currentScale = toFloat(scaleInput.value, 1);
        scaleInput.value = clamp(Math.max(currentScale, 1.12), 0.2, 5).toFixed(2);
        offsetXInput.value = '0';
        offsetYInput.value = '0';
        showToast('Auto-encuadre aplicado (modo estándar)', 'info');
    }

    if (bgEnableInput && !bgEnableInput.checked) {
        bgEnableInput.checked = true;
        await autoPickPhotoBgColor();
    }

    savePhotoConfigFromDOM();
    syncHudPhotoControls(getPhotoConfig());
    invalidatePreflightReport();
    tryRender();
}

async function updatePhotoSwatches() {
    const container = document.getElementById('editor-hud-swatches');
    if (!container) return;

    const photoImg = await getCurrentPhotoImage();
    if (!photoImg) {
        container.innerHTML = '';
        return;
    }

    const sourceW = photoImg.naturalWidth || photoImg.width;
    const sourceH = photoImg.naturalHeight || photoImg.height;
    if (!sourceW || !sourceH) {
        container.innerHTML = '';
        return;
    }

    const points = [
        [0.1, 0.1], [0.5, 0.08], [0.9, 0.1], [0.25, 0.2], [0.75, 0.2], [0.5, 0.3]
    ];
    const unique = [];

    points.forEach(([rx, ry]) => {
        const color = normalizeHexColor(samplePhotoPixel(photoImg, sourceW * rx, sourceH * ry), '#d9dee8');
        if (!unique.includes(color)) unique.push(color);
    });

    container.innerHTML = unique.slice(0, 6).map(color =>
        `<button type="button" class="swatch-btn" style="background:${color}" onclick="setPhotoBgColor('${color}')" title="${color}"></button>`
    ).join('');
}

function renderCarnet(index, targetCanvas, exportScale = 1) {
    return new Promise((resolve) => {
        const record = state.records[index];
        if (!record) { resolve(null); return; }

        const template = state.templateImage;
        const canvas = targetCanvas || document.getElementById('carnet-canvas');
        const ctx = canvas.getContext('2d');

        canvas.width = template.width * exportScale;
        canvas.height = template.height * exportScale;

        if (exportScale !== 1) {
            ctx.scale(exportScale, exportScale);
            ctx.imageSmoothingEnabled = true;
            ctx.imageSmoothingQuality = 'high';
        }

        // NOTE: Do NOT draw template here — photo goes FIRST (behind template)
        // Template is drawn AFTER photo so transparent areas reveal the photo underneath

        const photoConfig = getPhotoConfigForRecord(record);
        const photoKey = getRecordKey(record);
        const photoDataUrl = state.photosMap[photoKey];

        const drawTextsAndBarcode = () => {
            // --- PHOTO hitbox (for drag-and-drop) ---
            if (!targetCanvas) {
                state.hitboxes.push({
                    id: 'photo',
                    x: photoConfig.x, y: photoConfig.y,
                    w: photoConfig.w, h: photoConfig.h
                });
            }

            // --- NOMBRES (first names — top row) ---
            if (record.nombres) {
                const cfg = getFieldConfig('nombres');
                const fontStr = `${cfg.bold} ${cfg.size}px ${cfg.font}`.trim();
                ctx.font = fontStr;
                ctx.fillStyle = cfg.color;
                ctx.textAlign = cfg.align;
                ctx.textBaseline = 'top';

                let text = record.nombres;
                let metrics = ctx.measureText(text);
                if (metrics.width > cfg.maxWidth) {
                    while (ctx.measureText(text + '…').width > cfg.maxWidth && text.length > 0) {
                        text = text.slice(0, -1);
                    }
                    text += '…';
                }

                const finalMetrics = ctx.measureText(text);
                const textW = finalMetrics.width;
                const textH = cfg.size;

                let hitX = cfg.x;
                if (cfg.align === 'center') hitX = cfg.x - textW / 2;
                else if (cfg.align === 'right') hitX = cfg.x - textW;

                ctx.fillText(text, cfg.x, cfg.y);

                if (!targetCanvas) {
                    state.hitboxes.push({ id: 'nombres', x: hitX, y: cfg.y, w: textW, h: textH });
                }
            }

            // --- APELLIDOS (last names — second row) ---
            if (record.apellidos) {
                const cfg = getFieldConfig('apellidos');
                const fontStr = `${cfg.bold} ${cfg.size}px ${cfg.font}`.trim();
                ctx.font = fontStr;
                ctx.fillStyle = cfg.color;
                ctx.textAlign = cfg.align;
                ctx.textBaseline = 'top';

                let text = record.apellidos;
                let metrics = ctx.measureText(text);
                if (metrics.width > cfg.maxWidth) {
                    while (ctx.measureText(text + '…').width > cfg.maxWidth && text.length > 0) {
                        text = text.slice(0, -1);
                    }
                    text += '…';
                }

                const finalMetrics = ctx.measureText(text);
                const textW = finalMetrics.width;
                const textH = cfg.size;

                let hitX = cfg.x;
                if (cfg.align === 'center') hitX = cfg.x - textW / 2;
                else if (cfg.align === 'right') hitX = cfg.x - textW;

                ctx.fillText(text, cfg.x, cfg.y);

                if (!targetCanvas) {
                    state.hitboxes.push({ id: 'apellidos', x: hitX, y: cfg.y, w: textW, h: textH });
                }
            }

            // --- DNI (with prefix) ---
            if (record.dni) {
                const cfg = getFieldConfig('dni');
                const prefix = document.getElementById('field-dni-prefix')?.value || '';
                const fontStr = `${cfg.bold} ${cfg.size}px ${cfg.font}`.trim();
                ctx.font = fontStr;
                ctx.fillStyle = cfg.color;
                ctx.textAlign = cfg.align;
                ctx.textBaseline = 'top';

                const text = prefix + record.dni;
                const finalMetrics = ctx.measureText(text);
                const textW = finalMetrics.width;
                const textH = cfg.size;

                let hitX = cfg.x;
                if (cfg.align === 'center') hitX = cfg.x - textW / 2;
                else if (cfg.align === 'right') hitX = cfg.x - textW;

                ctx.fillText(text, cfg.x, cfg.y);

                if (!targetCanvas) {
                    state.hitboxes.push({ id: 'dni', x: hitX, y: cfg.y, w: textW, h: textH });
                }
            }

            // --- EXTRA / CARGO ---
            if (record.extra) {
                const cfg = getFieldConfig('extra');
                const fontStr = `${cfg.bold} ${cfg.size}px ${cfg.font}`.trim();
                ctx.font = fontStr;
                ctx.fillStyle = cfg.color;
                ctx.textAlign = cfg.align;
                ctx.textBaseline = 'top';

                let text = record.extra;
                let metrics = ctx.measureText(text);
                if (metrics.width > cfg.maxWidth) {
                    while (ctx.measureText(text + '…').width > cfg.maxWidth && text.length > 0) {
                        text = text.slice(0, -1);
                    }
                    text += '…';
                }

                const finalMetrics = ctx.measureText(text);
                const textW = finalMetrics.width;
                const textH = cfg.size;

                let hitX = cfg.x;
                if (cfg.align === 'center') hitX = cfg.x - textW / 2;
                else if (cfg.align === 'right') hitX = cfg.x - textW;

                ctx.fillText(text, cfg.x, cfg.y);

                if (!targetCanvas) {
                    state.hitboxes.push({ id: 'extra', x: hitX, y: cfg.y, w: textW, h: textH });
                }
            }

            // --- BARCODE ---
            if (record.dni) {
                drawBarcode(ctx, record.dni);
                if (!targetCanvas) {
                    const bcfg = getBarcodeConfig();
                    const bcCenteredX = Math.round((ctx.canvas.width / (ctx.getTransform().a || 1) - bcfg.w) / 2);
                    state.hitboxes.push({
                        id: 'barcode',
                        x: bcCenteredX, y: bcfg.y,
                        w: bcfg.w, h: bcfg.h
                    });
                }
            }

            // Show canvas
            canvas.style.display = 'block';
            if (document.getElementById('preview-placeholder')) {
                document.getElementById('preview-placeholder').style.display = 'none';
            }

            // Apply zoom
            if (!targetCanvas) {
                canvas.style.transform = `scale(${state.zoom})`;
                canvas.style.transformOrigin = 'center center';
            }

            updateNavigation();
            resolve(canvas);
        };

        // Reset hitboxes before rendering
        if (!targetCanvas) state.hitboxes = [];


        // Helper: draw photo then template on top, then texts
        const drawPhotoThenTemplate = (photoImg) => {
            if (photoConfig.bgEnabled) {
                ctx.save();
                ctx.fillStyle = photoConfig.bgColor;
                ctx.fillRect(photoConfig.x, photoConfig.y, photoConfig.w, photoConfig.h);
                ctx.restore();
            }

            // 1) Draw PHOTO first (behind everything)
            if (photoImg) {
                drawPhotoInFrame(ctx, photoImg, photoConfig);
            } else {
                // Placeholder for missing photo
                ctx.save();
                ctx.fillStyle = 'rgba(200, 200, 200, 0.3)';
                ctx.fillRect(photoConfig.x, photoConfig.y, photoConfig.w, photoConfig.h);
                ctx.fillStyle = '#999';
                ctx.font = '14px Poppins, Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                ctx.fillText('Sin foto', photoConfig.x + photoConfig.w / 2, photoConfig.y + photoConfig.h / 2);
                ctx.restore();
            }

            // 2) Draw TEMPLATE on top (transparent areas will show photo)
            ctx.drawImage(template, 0, 0);

            // 3) Draw texts and barcode on top of everything
            drawTextsAndBarcode();
        };

        if (photoDataUrl) {
            // Use LRU cache to avoid re-decoding on every render
            const cachedImg = state.photoImageCache.get(photoKey);
            if (cachedImg) {
                drawPhotoThenTemplate(cachedImg);
            } else {
                const photoImg = new Image();
                photoImg.onload = () => {
                    state.photoImageCache.set(photoKey, photoImg);
                    drawPhotoThenTemplate(photoImg);
                };
                photoImg.onerror = () => drawPhotoThenTemplate(null);
                photoImg.src = photoDataUrl;
            }
        } else {
            drawPhotoThenTemplate(null);
        }
    });
}

function drawBarcode(ctx, dniValue) {
    const cfg = getBarcodeConfig();
    // ctx.canvas.width is the physical pixel size; divide by the context's scale
    // factor to get the logical width (same coordinate space as drawing commands).
    const ctxScaleX = ctx.getTransform().a || 1;
    const logicalWidth = ctx.canvas.width / ctxScaleX;
    const centeredX = Math.round((logicalWidth - cfg.w) / 2);

    try {
        const barcodeCanvas = document.createElement('canvas');
        JsBarcode(barcodeCanvas, dniValue, {
            format: cfg.format,
            width: 2,
            height: cfg.h - (cfg.showText ? 18 : 0),
            displayValue: cfg.showText,
            fontSize: 12,
            margin: 0,
            background: 'transparent',
            lineColor: '#000000'
        });
        ctx.drawImage(barcodeCanvas, centeredX, cfg.y, cfg.w, cfg.h);
    } catch (err) {
        ctx.save();
        ctx.fillStyle = '#cc0000';
        ctx.font = '10px Arial';
        ctx.textAlign = 'center';
        ctx.fillText('Error código barras', centeredX + cfg.w / 2, cfg.y + cfg.h / 2);
        ctx.restore();
    }
}

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
    state.zoom                  = 1;
    state.inlineEditor.active   = false;
    state.inlineEditor.fieldId  = null;
    state.photoCropMode.active  = false;
    state.photoColorPicker.active = false;
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

// ===================== CANVAS INTERACTION (Canva-style) =====================

function setupCanvasDrag() {
    const canvas = document.getElementById('carnet-canvas');

    canvas.addEventListener('mousedown', onCanvasMouseDown);
    canvas.addEventListener('mousemove', onCanvasMouseMove);
    canvas.addEventListener('mouseup', onCanvasMouseUp);
    canvas.addEventListener('mouseleave', onCanvasMouseUp);
    canvas.addEventListener('dblclick', onCanvasDoubleClick);
    canvas.addEventListener('wheel', onCanvasWheel, { passive: false });

    // Touch support
    canvas.addEventListener('touchstart', e => {
        e.preventDefault();
        const t = e.touches[0];
        onCanvasMouseDown({ clientX: t.clientX, clientY: t.clientY, preventDefault: () => {} });
    }, { passive: false });
    canvas.addEventListener('touchmove', e => {
        e.preventDefault();
        const t = e.touches[0];
        onCanvasMouseMove({ clientX: t.clientX, clientY: t.clientY });
    }, { passive: false });
    canvas.addEventListener('touchend', () => onCanvasMouseUp());
}

function getCanvasCoords(e) {
    const canvas = document.getElementById('carnet-canvas');
    const rect = canvas.getBoundingClientRect();
    if (!rect.width || !rect.height) {
        return { x: 0, y: 0 };
    }
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return {
        x: (e.clientX - rect.left) * scaleX,
        y: (e.clientY - rect.top) * scaleY
    };
}

function hitTestResizeHandle(mx, my) {
    const selId = state.drag.selectedId;
    if (!selId || (selId !== 'photo' && selId !== 'barcode')) return null;

    const hb = state.hitboxes.find(h => h.id === selId);
    if (!hb) return null;

    const hs = 16;
    const corners = [
        { name: 'nw', cx: hb.x, cy: hb.y },
        { name: 'ne', cx: hb.x + hb.w, cy: hb.y },
        { name: 'sw', cx: hb.x, cy: hb.y + hb.h },
        { name: 'se', cx: hb.x + hb.w, cy: hb.y + hb.h }
    ];

    for (const c of corners) {
        if (mx >= c.cx - hs && mx <= c.cx + hs && my >= c.cy - hs && my <= c.cy + hs) {
            return c.name;
        }
    }
    return null;
}

function hitTestElement(mx, my) {
    for (let i = state.hitboxes.length - 1; i >= 0; i--) {
        const hb = state.hitboxes[i];
        const pad = 10;
        if (mx >= hb.x - pad && mx <= hb.x + hb.w + pad &&
            my >= hb.y - pad && my <= hb.y + hb.h + pad) {
            return hb.id;
        }
    }
    return null;
}

function onCanvasWheel(e) {
    if (!state.templateImage || state.records.length === 0) return;
    const coords = getCanvasCoords(e);
    const hitId = hitTestElement(coords.x, coords.y);
    const selectedIsPhoto = state.drag.selectedId === 'photo';
    const affectsPhoto = hitId === 'photo' || selectedIsPhoto;
    if (!affectsPhoto) return;

    e.preventDefault();
    if (!selectedIsPhoto) state.drag.selectedId = 'photo';

    const step = e.shiftKey ? 0.09 : 0.05;
    const delta = e.deltaY < 0 ? step : -step;
    const current = toFloat(document.getElementById('field-photo-scale')?.value, 1);
    setSelectedPhotoZoom((current + delta).toFixed(2), { trackHistory: true, updateHud: true });
}

function onCanvasMouseDown(e) {
    if (!state.templateImage || state.records.length === 0) return;
    e.preventDefault();
    if (state.inlineEditor.active) closeInlineEditor({ commit: true });

    const coords = getCanvasCoords(e);

    if (state.photoColorPicker.active) {
        const photoCfg = getPhotoConfig();
        const record = getCurrentRecord();
        const key = record ? getRecordKey(record) : '';
        const cached = key ? state.photoImageCache.get(key) : null;

        if (!cached) {
            stopPhotoColorPickMode();
            showToast('Espera un instante y vuelve a intentar el gotero', 'warning');
            return;
        }

        const picked = getPhotoColorFromCanvasPoint(coords.x, coords.y, cached, photoCfg);
        stopPhotoColorPickMode();
        if (!picked) {
            showToast('Haz clic dentro del área de la foto para tomar color', 'info');
            return;
        }
        setPhotoBgColor(picked);
        showToast(`Color aplicado: ${picked}`, 'success');
        return;
    }

    const hitHandle = hitTestResizeHandle(coords.x, coords.y);
    const hitId = hitTestElement(coords.x, coords.y);

    // Drag on photo = always pan photo. Alt+drag = move the frame instead.
    const quickPanMode = hitId === 'photo' && !hitHandle && !e.altKey;
    if (quickPanMode) {
        state.drag.selectedId = 'photo';
        state.drag.active = true;
        state.drag.historyCaptured = false;
        state.drag.photoPanActive = true;
        state.drag.resizeHandle = null;
        state.drag.elementId = 'photo';
        state.drag.startMouseX = coords.x;
        state.drag.startMouseY = coords.y;
        state.drag.startPhotoOffsetX = toInt(document.getElementById('field-photo-offset-x')?.value, 0);
        state.drag.startPhotoOffsetY = toInt(document.getElementById('field-photo-offset-y')?.value, 0);
        document.getElementById('carnet-canvas').style.cursor = 'grabbing';
        renderCarnet(state.currentIndex).then(() => {
            drawSelectionOverlay();
            updateEditorHud();
        });
        return;
    }

    if (hitHandle) {
        const id = state.drag.selectedId;
        const hb = state.hitboxes.find(h => h.id === id);
        if (!hb) return; // safety: hitbox desynced
        state.drag.active = true;
        state.drag.historyCaptured = false;
        state.drag.photoPanActive = false;
        state.drag.resizeHandle = hitHandle;
        state.drag.elementId = id;
        state.drag.startMouseX = coords.x;
        state.drag.startMouseY = coords.y;
        state.drag.startInputX = toInt(document.getElementById(`field-${id}-x`)?.value, 0);
        state.drag.startInputY = toInt(document.getElementById(`field-${id}-y`)?.value, 0);
        state.drag.startElemX = hb.x;
        state.drag.startElemY = hb.y;
        state.drag.startElemW = hb.w;
        state.drag.startElemH = hb.h;
        document.getElementById('carnet-canvas').style.cursor = getCursorForHandle(hitHandle);
        updateEditorHud();
        return;
    }

    if (hitId) {
        const hb = state.hitboxes.find(h => h.id === hitId);
        if (!hb) return; // safety: hitbox desynced
        state.drag.selectedId = hitId;
        state.drag.active = true;
        state.drag.historyCaptured = false;
        state.drag.photoPanActive = false;
        state.drag.resizeHandle = null;
        state.drag.elementId = hitId;
        state.drag.startMouseX = coords.x;
        state.drag.startMouseY = coords.y;
        state.drag.startInputX = toInt(document.getElementById(`field-${hitId}-x`)?.value, 0);
        state.drag.startInputY = toInt(document.getElementById(`field-${hitId}-y`)?.value, 0);
        state.drag.startElemX = hb.x;
        state.drag.startElemY = hb.y;
        state.drag.startElemW = hb.w;
        state.drag.startElemH = hb.h;
        state.drag.snapGuides = { x: null, y: null };

        document.getElementById('carnet-canvas').style.cursor = 'grabbing';
        renderCarnet(state.currentIndex).then(() => {
            drawSelectionOverlay();
            updateEditorHud();
        });
        updateEditorHud();
    } else {
        // Click empty area
        if (state.drag.selectedId) {
            state.drag.selectedId = null;
            state.drag.hoveredId = null;
            renderCarnet(state.currentIndex);
            updateEditorHud();
        }
    }
}

function onCanvasMouseMove(e) {
    const canvas = document.getElementById('carnet-canvas');
    if (!state.templateImage || state.records.length === 0) return;

    const coords = getCanvasCoords(e);

    if (state.photoColorPicker.active) {
        canvas.style.cursor = 'crosshair';
        return;
    }

    if (state.drag.active) {
        let dx = coords.x - state.drag.startMouseX;
        let dy = coords.y - state.drag.startMouseY;
        const id = state.drag.elementId;
        const handle = state.drag.resizeHandle;
        invalidatePreflightReport();

        const movedEnough = Math.abs(dx) > 1 || Math.abs(dy) > 1;
        if (movedEnough && !state.drag.historyCaptured) {
            const reason = state.drag.photoPanActive ? 'photo-pan' : (handle ? 'resize' : 'move');
            pushUndoSnapshot(reason);
            state.drag.historyCaptured = true;
        }

        state.drag.snapGuides = { x: null, y: null };

        if (state.drag.photoPanActive && id === 'photo') {
            const offsetXInput = document.getElementById('field-photo-offset-x');
            const offsetYInput = document.getElementById('field-photo-offset-y');
            if (offsetXInput && offsetYInput) {
                offsetXInput.value = Math.round(state.drag.startPhotoOffsetX + dx);
                offsetYInput.value = Math.round(state.drag.startPhotoOffsetY + dy);
                savePhotoConfigFromDOM();
            }

            renderCarnet(state.currentIndex).then(() => {
                drawSelectionOverlay();
                updateEditorHud();
            });
            return;
        }

        if (handle) {
            // RESIZE
            let newX = state.drag.startElemX;
            let newY = state.drag.startElemY;
            let newW = state.drag.startElemW;
            let newH = state.drag.startElemH;

            const aspect = state.drag.startElemH / state.drag.startElemW;
            const keepRatio = e.shiftKey || e.altKey;

            if (handle === 'se') {
                newW = Math.max(20, newW + Math.round(dx));
                if (keepRatio) {
                    newH = Math.round(newW * aspect);
                } else {
                    newH = Math.max(20, newH + Math.round(dy));
                }
            } else if (handle === 'sw') {
                newW = Math.max(20, newW - Math.round(dx));
                if (keepRatio) {
                    newH = Math.round(newW * aspect);
                    newX = state.drag.startElemX + (state.drag.startElemW - newW);
                } else {
                    newX += Math.round(dx);
                    newH = Math.max(20, newH + Math.round(dy));
                }
            } else if (handle === 'ne') {
                newW = Math.max(20, newW + Math.round(dx));
                if (keepRatio) {
                    newH = Math.round(newW * aspect);
                    newY = state.drag.startElemY + (state.drag.startElemH - newH);
                } else {
                    newY += Math.round(dy);
                    newH = Math.max(20, newH - Math.round(dy));
                }
            } else if (handle === 'nw') {
                newW = Math.max(20, newW - Math.round(dx));
                if (keepRatio) {
                    newH = Math.round(newW * aspect);
                    newX = state.drag.startElemX + (state.drag.startElemW - newW);
                    newY = state.drag.startElemY + (state.drag.startElemH - newH);
                } else {
                    newX += Math.round(dx);
                    newY += Math.round(dy);
                    newH = Math.max(20, newH - Math.round(dy));
                }
            }

            // For resize, the input coords are the bounding box coords (only applies to photo/barcode)
            document.getElementById(`field-${id}-x`).value = newX;
            document.getElementById(`field-${id}-y`).value = newY;
            document.getElementById(`field-${id}-w`).value = newW;
            document.getElementById(`field-${id}-h`).value = newH;
            if (id === 'photo') savePhotoConfigFromDOM();
        } else {
            // MOVE with Magnetic Snapping
            // We snap based on the VISUAL bounding box (startElem)
            let newVisX = state.drag.startElemX + dx;
            let newVisY = state.drag.startElemY + dy;
            
            const centerX = state.templateImage.width / 2;
            const centerY = state.templateImage.height / 2;
            const elemCenterX = newVisX + state.drag.startElemW / 2;
            const elemCenterY = newVisY + state.drag.startElemH / 2;
            
            const snapThreshold = 12; // pixels

            if (Math.abs(elemCenterX - centerX) < snapThreshold) {
                dx = centerX - (state.drag.startElemX + state.drag.startElemW / 2);
                state.drag.snapGuides.x = centerX;
            }
            if (Math.abs(elemCenterY - centerY) < snapThreshold) {
                dy = centerY - (state.drag.startElemY + state.drag.startElemH / 2);
                state.drag.snapGuides.y = centerY;
            }

            // Also snap photo to edges of the canvas to make full cover easier
            if (id === 'photo') {
                if (Math.abs(state.drag.startElemX + dx) < snapThreshold) dx = -state.drag.startElemX;
                if (Math.abs(state.drag.startElemY + dy) < snapThreshold) dy = -state.drag.startElemY;
            }

            // Re-apply snapped dx, dy to the original INPUT anchors (to preserve text centering)
            document.getElementById(`field-${id}-x`).value = Math.round(state.drag.startInputX + dx);
            document.getElementById(`field-${id}-y`).value = Math.round(state.drag.startInputY + dy);
            
            if (id === 'photo') savePhotoConfigFromDOM();
        }

        renderCarnet(state.currentIndex).then(() => {
            drawSelectionOverlay();
            updateEditorHud();
        });
    } else {
        const handle = hitTestResizeHandle(coords.x, coords.y);
        if (handle) {
            canvas.style.cursor = getCursorForHandle(handle);
            return;
        }

        const hitId = hitTestElement(coords.x, coords.y);
        if (hitId) {
            if (hitId === 'photo' && !e.altKey) {
                canvas.style.cursor = 'move';
            } else {
                canvas.style.cursor = e.altKey && hitId === 'photo' ? 'grab' : 'grab';
            }
            if (state.drag.hoveredId !== hitId) {
                state.drag.hoveredId = hitId;
                renderCarnet(state.currentIndex).then(() => {
                    drawSelectionOverlay();
                    updateEditorHud();
                });
            }
        } else {
            canvas.style.cursor = 'default';
            if (state.drag.hoveredId) {
                state.drag.hoveredId = null;
                renderCarnet(state.currentIndex).then(() => {
                    drawSelectionOverlay();
                    updateEditorHud();
                });
            }
        }
    }
}

function onCanvasMouseUp() {
    if (state.drag.active) {
        state.drag.active = false;
        state.drag.historyCaptured = false;
        state.drag.photoPanActive = false;
        state.drag.resizeHandle = null;
        state.drag.snapGuides = null;
        const canvas = document.getElementById('carnet-canvas');
        canvas.style.cursor = state.photoCropMode.active && state.drag.selectedId === 'photo' ? 'move' : 'default';
        renderCarnet(state.currentIndex).then(() => {
            drawSelectionOverlay();
            updateEditorHud();
        });
        updateEditorHud();
    }
}

function getCursorForHandle(h) {
    return { nw: 'nw-resize', ne: 'ne-resize', sw: 'sw-resize', se: 'se-resize' }[h] || 'grab';
}

function drawSelectionOverlay() {
    const canvas = document.getElementById('carnet-canvas');
    const ctx = canvas.getContext('2d');

    // Draw Snap Guides (Pink Canva-style lines)
    if (state.drag.snapGuides) {
        ctx.save();
        ctx.strokeStyle = '#ec4899'; // Pink-500
        ctx.lineWidth = 1.5;
        ctx.setLineDash([6, 4]);

        if (state.drag.snapGuides.x !== null) {
            ctx.beginPath();
            ctx.moveTo(state.drag.snapGuides.x, 0);
            ctx.lineTo(state.drag.snapGuides.x, canvas.height);
            ctx.stroke();
        }
        if (state.drag.snapGuides.y !== null) {
            ctx.beginPath();
            ctx.moveTo(0, state.drag.snapGuides.y);
            ctx.lineTo(canvas.width, state.drag.snapGuides.y);
            ctx.stroke();
        }
        ctx.restore();
    }

    const labels = {
        nombres: 'Nombres', apellidos: 'Apellidos', dni: 'DNI',
        extra: 'Cargo', photo: 'Foto', barcode: 'Código de Barras'
    };

    state.hitboxes.forEach(hb => {
        const isSelected = hb.id === state.drag.selectedId;
        const isHovered = hb.id === state.drag.hoveredId && !isSelected;

        if (!isSelected && !isHovered) return;

        const color = isSelected ? '#6366f1' : 'rgba(99, 102, 241, 0.40)';
        const lw = isSelected ? 2.5 : 1.5;
        const pad = 5;

        ctx.save();

        // Border
        ctx.strokeStyle = color;
        ctx.lineWidth = lw;
        ctx.setLineDash(isSelected ? [] : [5, 5]);
        ctx.strokeRect(hb.x - pad, hb.y - pad, hb.w + pad * 2, hb.h + pad * 2);
        ctx.setLineDash([]);

        // Corner handles
        if (isSelected) {
            const hs = 9;
            const canResize = (hb.id === 'photo' || hb.id === 'barcode');

            const corners = [
                [hb.x - pad - 1, hb.y - pad - 1],
                [hb.x + hb.w + pad - hs + 1, hb.y - pad - 1],
                [hb.x - pad - 1, hb.y + hb.h + pad - hs + 1],
                [hb.x + hb.w + pad - hs + 1, hb.y + hb.h + pad - hs + 1]
            ];

            corners.forEach(([cx, cy]) => {
                ctx.fillStyle = canResize ? '#ffffff' : color;
                ctx.fillRect(cx, cy, hs, hs);
                ctx.strokeStyle = '#6366f1';
                ctx.lineWidth = 2;
                ctx.strokeRect(cx, cy, hs, hs);
            });
        }

        // Label badge
        const label = labels[hb.id] || hb.id;
        ctx.font = 'bold 13px Inter, Poppins, Arial';
        const tw = ctx.measureText(label).width;
        const lbW = tw + 14;
        const lbH = 22;
        const lbX = hb.x - pad;
        const lbY = hb.y - pad - lbH - 6;

        ctx.fillStyle = color;
        if (ctx.roundRect) {
            ctx.beginPath();
            ctx.roundRect(lbX, lbY, lbW, lbH, 5);
            ctx.fill();
        } else {
            ctx.fillRect(lbX, lbY, lbW, lbH);
        }

        ctx.fillStyle = '#ffffff';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        ctx.fillText(label, lbX + 7, lbY + lbH / 2);

        if (isSelected && hb.id === 'photo') {
            const hint = 'Arrastrar: encuadrar · Rueda: zoom · Alt+arrastrar: mover marco · Doble clic: auto-encuadre';
            ctx.font = '600 11px Inter, Arial';
            const hintW = ctx.measureText(hint).width + 12;
            const hintH = 18;
            const hintX = Math.max(6, hb.x - pad);
            const hintY = Math.min(canvas.height - hintH - 6, hb.y + hb.h + pad + 8);
            ctx.fillStyle = 'rgba(15, 23, 42, 0.92)';
            ctx.strokeStyle = 'rgba(99, 102, 241, 0.85)';
            ctx.lineWidth = 1;
            ctx.fillRect(hintX, hintY, hintW, hintH);
            ctx.strokeRect(hintX, hintY, hintW, hintH);
            ctx.fillStyle = '#dbeafe';
            ctx.fillText(hint, hintX + 6, hintY + hintH / 2);
        }

        ctx.restore();
    });
}

function onCanvasDoubleClick(e) {
    if (!state.templateImage || state.records.length === 0) return;
    const coords = getCanvasCoords(e);
    const hitId = hitTestElement(coords.x, coords.y);
    if (!hitId) return;

    state.drag.selectedId = hitId;
    const hb = state.hitboxes.find(h => h.id === hitId);

    renderCarnet(state.currentIndex).then(() => {
        drawSelectionOverlay();
        updateEditorHud();
        if (hitId === 'photo') {
            autoFrameCurrentPhoto();
        } else if (hb && ['nombres', 'apellidos', 'dni', 'extra'].includes(hitId)) {
            openInlineEditor(hitId, hb);
        }
    });
}

function startInlineTextEditFromSelection() {
    const selected = getSelectedHitbox();
    if (!selected) return;

    if (!['nombres', 'apellidos', 'dni', 'extra'].includes(selected.id)) {
        showToast('Solo puedes editar texto de Nombres, Apellidos, DNI o Cargo', 'info');
        return;
    }

    openInlineEditor(selected.id, selected.hb);
}

function closeInlineEditor(options = { commit: false }) {
    const input = document.getElementById('canvas-inline-editor');
    if (!input) {
        state.inlineEditor.active = false;
        state.inlineEditor.fieldId = null;
        return;
    }

    const shouldCommit = !!options.commit;
    const fieldId = state.inlineEditor.fieldId;

    if (shouldCommit && fieldId && state.records.length > 0) {
        pushUndoSnapshot(`inline-edit:${fieldId}`);
        const record = state.records[state.currentIndex];
        const value = input.value.trim();
        invalidatePreflightReport();

        if (fieldId === 'dni') {
            if (value) record.dni = value;
        } else if (fieldId === 'extra') {
            record.extra = value;
        } else if (fieldId === 'nombres' || fieldId === 'apellidos') {
            record[fieldId] = value;
        }

        showDataPreview();
        tryRender();
    }

    input.remove();
    state.inlineEditor.active = false;
    state.inlineEditor.fieldId = null;
}

function openInlineEditor(fieldId, hitbox) {
    if (!['nombres', 'apellidos', 'dni', 'extra'].includes(fieldId)) return;
    if (!hitbox || !state.records.length) return;

    closeInlineEditor({ commit: false });

    const record = state.records[state.currentIndex];
    const canvas = document.getElementById('carnet-canvas');
    const previewArea = document.getElementById('preview-area');
    const canvasRect = canvas.getBoundingClientRect();
    const previewRect = previewArea.getBoundingClientRect();

    if (!canvasRect.width || !canvasRect.height) return;

    const scaleX = canvasRect.width / canvas.width;
    const scaleY = canvasRect.height / canvas.height;

    const editor = document.createElement('input');
    editor.type = 'text';
    editor.id = 'canvas-inline-editor';
    editor.className = 'canvas-inline-editor';

    const initialValue = String(record[fieldId] ?? '');
    editor.value = initialValue;
    editor.setAttribute('aria-label', `Editar ${fieldId}`);

    const left = (canvasRect.left - previewRect.left) + hitbox.x * scaleX + previewArea.scrollLeft - 8;
    const top = (canvasRect.top - previewRect.top) + hitbox.y * scaleY + previewArea.scrollTop - 6;
    const width = Math.max(160, hitbox.w * scaleX + 18);

    editor.style.left = `${Math.max(8, left)}px`;
    editor.style.top = `${Math.max(8, top)}px`;
    editor.style.width = `${Math.min(width, previewArea.clientWidth - 20)}px`;

    previewArea.appendChild(editor);
    editor.focus();
    editor.select();

    state.inlineEditor.active = true;
    state.inlineEditor.fieldId = fieldId;

    editor.addEventListener('keydown', (ev) => {
        if (ev.key === 'Enter') {
            ev.preventDefault();
            closeInlineEditor({ commit: true });
        } else if (ev.key === 'Escape') {
            ev.preventDefault();
            closeInlineEditor({ commit: false });
        }
    });

    editor.addEventListener('blur', () => {
        closeInlineEditor({ commit: true });
    });
}

function getSelectedHitbox() {
    const selectedId = state.drag.selectedId;
    if (!selectedId) return null;

    const hb = state.hitboxes.find(h => h.id === selectedId);
    if (!hb) return null;

    return { id: selectedId, hb };
}

function visualXToAnchorX(id, visualX, width) {
    if (!['nombres', 'apellidos', 'dni', 'extra'].includes(id)) {
        return Math.round(visualX);
    }

    const align = document.getElementById(`field-${id}-align`)?.value || 'left';
    if (align === 'center') return Math.round(visualX + width / 2);
    if (align === 'right') return Math.round(visualX + width);
    return Math.round(visualX);
}

function applyVisualPositionToInputs(id, hitbox, visualX, visualY) {
    const xInput = document.getElementById(`field-${id}-x`);
    const yInput = document.getElementById(`field-${id}-y`);
    if (!xInput || !yInput) return;

    const clampedX = Math.max(0, Math.round(visualX));
    const clampedY = Math.max(0, Math.round(visualY));

    xInput.value = visualXToAnchorX(id, clampedX, hitbox.w);
    yInput.value = clampedY;

    if (id === 'photo') {
        savePhotoConfigFromDOM();
    }
}

function nudgeSelectedElement(dx, dy) {
    if (!state.templateImage || state.records.length === 0) return;

    const selected = getSelectedHitbox();
    if (!selected) return;

    const { id, hb } = selected;
    const maxX = Math.max(0, state.templateImage.width - hb.w);
    const maxY = Math.max(0, state.templateImage.height - hb.h);

    const nextX = Math.min(maxX, Math.max(0, hb.x + dx));
    const nextY = Math.min(maxY, Math.max(0, hb.y + dy));

    applyVisualPositionToInputs(id, hb, nextX, nextY);
    tryRender();
}

function alignSelectedElement(axis = 'x') {
    if (!state.templateImage || state.records.length === 0) return;

    const selected = getSelectedHitbox();
    if (!selected) {
        showToast('Selecciona un elemento en el canvas primero', 'info');
        return;
    }

    const { id, hb } = selected;
    pushUndoSnapshot(`align-${axis}`);
    const centerX = Math.round((state.templateImage.width - hb.w) / 2);
    const centerY = Math.round((state.templateImage.height - hb.h) / 2);

    const nextX = axis === 'x' ? centerX : hb.x;
    const nextY = axis === 'y' ? centerY : hb.y;

    applyVisualPositionToInputs(id, hb, nextX, nextY);
    tryRender();
    showToast(`Elemento centrado en eje ${axis.toUpperCase()}`, 'success');
}

function resetSelectedElement() {
    const selected = getSelectedHitbox();
    if (!selected) {
        showToast('Selecciona un elemento para restablecer', 'info');
        return;
    }

    const { id } = selected;
    pushUndoSnapshot('reset-element');
    const keys = id === 'photo'
        ? [
            'field-photo-x', 'field-photo-y', 'field-photo-w', 'field-photo-h',
            'field-photo-fit', 'field-photo-scale', 'field-photo-offset-x', 'field-photo-offset-y',
            'field-photo-bg-enable', 'field-photo-bg-color'
        ]
        : [`field-${id}-x`, `field-${id}-y`, `field-${id}-w`, `field-${id}-h`];

    keys.forEach(inputId => {
        const el = document.getElementById(inputId);
        if (!el) return;
        if (!Object.prototype.hasOwnProperty.call(state.defaultFieldValues, inputId)) return;

        const defaultCfg = state.defaultFieldValues[inputId];
        if (defaultCfg?.type === 'checkbox') {
            el.checked = !!defaultCfg.checked;
        } else if (defaultCfg?.type === 'value') {
            el.value = defaultCfg.value;
        }
    });

    if (id === 'photo') {
        savePhotoConfigFromDOM();
    }

    tryRender();
    showToast('Elemento restablecido a valores iniciales', 'success');
}

function adjustSelectedPhotoZoom(delta) {
    if (state.drag.selectedId !== 'photo') return;

    const input = document.getElementById('field-photo-scale');
    if (!input) return;

    const current = toFloat(input.value, 1);
    const next = clamp(current + delta, 0.2, 5).toFixed(2);
    setSelectedPhotoZoom(next, { trackHistory: true });
}

function setSelectedPhotoZoom(value, options = {}) {
    if (state.drag.selectedId !== 'photo') return;
    const input = document.getElementById('field-photo-scale');
    if (!input) return;

    const now = Date.now();
    const shouldTrack = !!options.trackHistory || now > state.history.zoomSessionUntil;
    if (shouldTrack) {
        pushUndoSnapshot('photo-zoom');
        state.history.zoomSessionUntil = now + 380;
    }

    const next = clamp(toFloat(value, 1), 0.2, 5);
    input.value = next.toFixed(2);
    savePhotoConfigFromDOM();
    if (options.updateHud !== false) {
        syncHudPhotoControls(getPhotoConfig());
    }
    invalidatePreflightReport();
    tryRender();
}

function panSelectedPhoto(dx, dy) {
    if (state.drag.selectedId !== 'photo') return;
    if (Date.now() > state.history.panSessionUntil) {
        pushUndoSnapshot('photo-pan-nudge');
        state.history.panSessionUntil = Date.now() + 350;
    }

    const inputX = document.getElementById('field-photo-offset-x');
    const inputY = document.getElementById('field-photo-offset-y');
    if (!inputX || !inputY) return;

    const nextX = toInt(inputX.value, 0) + dx;
    const nextY = toInt(inputY.value, 0) + dy;
    inputX.value = nextX;
    inputY.value = nextY;
    invalidatePreflightReport();
    savePhotoConfigFromDOM();
    tryRender();
}

function togglePhotoIndividualFromHud(enabled) {
    if (!state.records.length) return;
    const target = { id: 'photo-individual-mode', checked: !!enabled };
    handleInputChange({ target });
}

function setPhotoFitMode(mode) {
    if (state.drag.selectedId !== 'photo') return;
    if (!['cover', 'contain'].includes(mode)) return;
    pushUndoSnapshot('photo-fit');

    const input = document.getElementById('field-photo-fit');
    if (!input) return;

    input.value = mode;
    invalidatePreflightReport();
    savePhotoConfigFromDOM();
    syncHudPhotoControls(getPhotoConfig());
    tryRender();
}

function setPhotoCropMode(active) {
    state.photoCropMode.active = !!active;
    if (!state.photoCropMode.active) {
        state.drag.photoPanActive = false;
    }
    const canvas = document.getElementById('carnet-canvas');
    if (canvas) {
        canvas.style.cursor = state.photoCropMode.active && state.drag.selectedId === 'photo' ? 'move' : 'default';
    }
    syncHudPhotoControls(getPhotoConfig());
    updateEditorHud();
}

function togglePhotoCropMode() {
    if (state.drag.selectedId !== 'photo') {
        state.drag.selectedId = 'photo';
        tryRender();
    }
    setPhotoCropMode(!state.photoCropMode.active);
    const msg = state.photoCropMode.active
        ? 'Modo reencuadre activo: arrastra la foto dentro del marco'
        : 'Modo reencuadre desactivado';
    showToast(msg, 'info');
}

function resetSelectedPhotoCrop() {
    if (state.drag.selectedId !== 'photo') return;
    pushUndoSnapshot('photo-reset-crop');

    const fitInput = document.getElementById('field-photo-fit');
    const scaleInput = document.getElementById('field-photo-scale');
    const offsetXInput = document.getElementById('field-photo-offset-x');
    const offsetYInput = document.getElementById('field-photo-offset-y');

    if (!fitInput || !scaleInput || !offsetXInput || !offsetYInput) return;

    fitInput.value = 'cover';
    scaleInput.value = '1.00';
    offsetXInput.value = '0';
    offsetYInput.value = '0';
    invalidatePreflightReport();
    savePhotoConfigFromDOM();
    syncHudPhotoControls(getPhotoConfig());
    tryRender();
}

function updateEditorHud() {
    const hud = document.getElementById('editor-hud');
    if (!hud) return;

    const hasRenderableData = !!state.templateImage && state.records.length > 0;
    const actionButtons = hud.querySelectorAll('button');
    const photoControls = document.getElementById('editor-hud-photo');
    const swatches = document.getElementById('editor-hud-swatches');
    const hudIndividual = document.getElementById('hud-photo-individual');
    const hudBgEnable = document.getElementById('hud-photo-bg-enable');
    const hudBgColor = document.getElementById('hud-photo-bg-color');
    const hudZoom = document.getElementById('hud-photo-zoom');
    const fitCover = document.getElementById('hud-fit-cover');
    const fitContain = document.getElementById('hud-fit-contain');
    const cropBtn = document.getElementById('hud-crop-mode');
    const nameEl = document.getElementById('editor-hud-name');
    const detailsEl = document.getElementById('editor-hud-details');
    const setPhotoHudDisabled = (disabled) => {
        if (photoControls) {
            photoControls.querySelectorAll('button').forEach(btn => { btn.disabled = disabled; });
        }
        if (swatches) {
            swatches.querySelectorAll('button').forEach(btn => { btn.disabled = disabled; });
        }
        if (hudIndividual) hudIndividual.disabled = disabled;
        if (hudBgEnable) hudBgEnable.disabled = disabled;
        if (hudBgColor) hudBgColor.disabled = disabled;
        if (hudZoom) hudZoom.disabled = disabled;
        if (fitCover) fitCover.disabled = disabled;
        if (fitContain) fitContain.disabled = disabled;
        if (cropBtn) cropBtn.disabled = disabled;
    };

    if (!hasRenderableData) {
        hud.classList.remove('active');
        actionButtons.forEach(btn => { btn.disabled = true; });
        setPhotoHudDisabled(true);
        if (photoControls) photoControls.classList.remove('active');
        if (swatches) swatches.classList.remove('active');
        if (nameEl) nameEl.textContent = 'Sin selección';
        if (detailsEl) detailsEl.textContent = 'Carga plantilla y fotos para editar';
        return;
    }

    hud.classList.add('active');

    const labels = {
        nombres: 'Nombres',
        apellidos: 'Apellidos',
        dni: 'DNI',
        extra: 'Cargo / Extra',
        photo: 'Foto',
        barcode: 'Código de Barras'
    };

    const selected = getSelectedHitbox();
    if (!selected) {
        actionButtons.forEach(btn => { btn.disabled = true; });
        setPhotoHudDisabled(true);
        if (photoControls) photoControls.classList.remove('active');
        if (swatches) swatches.classList.remove('active');
        if (nameEl) nameEl.textContent = 'Sin selección';
        if (detailsEl) detailsEl.textContent = 'Haz clic sobre un elemento para editar';
        return;
    }

    const { id, hb } = selected;
    actionButtons.forEach(btn => { btn.disabled = false; });
    if (photoControls) {
        photoControls.classList.toggle('active', id === 'photo');
    }
    if (swatches) {
        swatches.classList.toggle('active', id === 'photo');
    }
    setPhotoHudDisabled(id !== 'photo');
    if (nameEl) nameEl.textContent = labels[id] || id;
    if (detailsEl) {
        if (id === 'photo') {
            const photoCfg = getPhotoConfig();
            syncHudPhotoControls(photoCfg);
            const pickerLabel = state.photoColorPicker.active ? ' · Gotero activo' : '';
            const cropLabel = state.photoCropMode.active ? ' · Reencuadre activo' : '';
            detailsEl.textContent = `X ${Math.round(hb.x)} · Y ${Math.round(hb.y)} · W ${Math.round(hb.w)} · H ${Math.round(hb.h)} · ${photoCfg.fit.toUpperCase()} · Zoom ${photoCfg.scale.toFixed(2)}${pickerLabel}${cropLabel}`;
        } else {
            if (state.photoCropMode.active) {
                state.photoCropMode.active = false;
                state.drag.photoPanActive = false;
                const canvas = document.getElementById('carnet-canvas');
                if (canvas) canvas.style.cursor = 'default';
            }
            detailsEl.textContent = `X ${Math.round(hb.x)} · Y ${Math.round(hb.y)} · W ${Math.round(hb.w)} · H ${Math.round(hb.h)}`;
        }
    }
}

// ===================== PRE-CHEQUEO =====================

function invalidatePreflightReport() {
    state.preflightReport = null;
    renderPreflightReport(null);
}

function getPhotoUpscaleFactor(photoCfg, sourceW, sourceH, exportScale = 1) {
    if (!sourceW || !sourceH) return 999;
    const scaleX = photoCfg.w / sourceW;
    const scaleY = photoCfg.h / sourceH;
    const baseScale = photoCfg.fit === 'contain' ? Math.min(scaleX, scaleY) : Math.max(scaleX, scaleY);
    return baseScale * photoCfg.scale * exportScale;
}

function renderPreflightReport(report) {
    const box = document.getElementById('preflight-report');
    if (!box) return;

    if (!report) {
        box.style.display = 'none';
        box.innerHTML = '';
        return;
    }

    const duplicateList = report.duplicates.slice(0, 8)
        .map(d => `• DNI ${escapeHtml(d.key)} (${d.count} veces)`)
        .join('<br>');
    const missingList = report.missingPhotos.slice(0, 8)
        .map(d => `• ${escapeHtml(d.dni || 'SIN_DNI')} - ${escapeHtml(d.name || 'Registro sin nombre')}`)
        .join('<br>');
    const lowQualityList = report.lowQuality.slice(0, 8)
        .map(d => `• ${escapeHtml(d.dni || 'SIN_DNI')} (${d.width}×${d.height}px, x${d.factor.toFixed(2)} de escalado)`)
        .join('<br>');

    box.innerHTML = `
        <div class="pf-summary ${report.ok ? 'pf-ok' : 'pf-error'}">
            ${report.ok ? 'Listo para exportar' : 'Se detectaron puntos críticos'}
        </div>
        <div class="pf-summary">
            Total: <strong>${report.total}</strong> ·
            Duplicados: <strong class="${report.duplicates.length ? 'pf-warn' : 'pf-ok'}">${report.duplicates.length}</strong> ·
            Sin foto: <strong class="${report.missingPhotos.length ? 'pf-error' : 'pf-ok'}">${report.missingPhotos.length}</strong> ·
            Baja calidad: <strong class="${report.lowQuality.length ? 'pf-warn' : 'pf-ok'}">${report.lowQuality.length}</strong>
        </div>
        ${duplicateList ? `<div class="pf-list"><strong class="pf-warn">DNI duplicados</strong><br>${duplicateList}</div>` : ''}
        ${missingList ? `<div class="pf-list"><strong class="pf-error">Registros sin foto</strong><br>${missingList}</div>` : ''}
        ${lowQualityList ? `<div class="pf-list"><strong class="pf-warn">Fotos con posible pixelado en el DPI actual</strong><br>${lowQualityList}</div>` : ''}
    `;
    box.style.display = 'block';
}

async function runPreflightCheck(options = {}) {
    const opts = {
        showToastOnPass: true,
        silent: false,
        ...options
    };

    if (!state.templateImage || state.records.length === 0) {
        const emptyReport = {
            ok: false,
            total: 0,
            duplicates: [],
            missingPhotos: [],
            lowQuality: []
        };
        state.preflightReport = emptyReport;
        renderPreflightReport(emptyReport);
        if (!opts.silent) showToast('No hay datos suficientes para validar', 'warning');
        return emptyReport;
    }

    const { widthCM, heightCM } = getConfiguredCarnetSizeCM();
    const dpi = getExportDPI();
    const targetW = cmToPx(widthCM, dpi);
    const targetH = cmToPx(heightCM, dpi);
    const exportScale = getRenderScaleForTargetPx(targetW, targetH);

    const counts = {};
    const duplicates = [];
    const missingPhotos = [];
    const lowQuality = [];
    const seenDuplicate = new Set();

    for (let i = 0; i < state.records.length; i++) {
        assertJobNotCancelled();
        const record = state.records[i];
        const key = getRecordKey(record);
        counts[key] = (counts[key] || 0) + 1;
        if (counts[key] > 1 && key && !seenDuplicate.has(key)) {
            seenDuplicate.add(key);
            duplicates.push({ key, count: counts[key] });
        } else if (counts[key] > 1 && key) {
            const idx = duplicates.findIndex(d => d.key === key);
            if (idx >= 0) duplicates[idx].count = counts[key];
        }

        const src = key ? state.photosMap[key] : null;
        if (!src) {
            missingPhotos.push({
                index: i,
                dni: record?.dni || '',
                name: `${record?.apellidos || ''} ${record?.nombres || ''}`.trim()
            });
            continue;
        }

        const img = await getPhotoImageByKey(key);
        if (!img) {
            missingPhotos.push({
                index: i,
                dni: record?.dni || '',
                name: `${record?.apellidos || ''} ${record?.nombres || ''}`.trim()
            });
            continue;
        }

        const sourceW = img.naturalWidth || img.width;
        const sourceH = img.naturalHeight || img.height;
        const photoCfg = getPhotoConfigForRecord(record);
        const factor = getPhotoUpscaleFactor(photoCfg, sourceW, sourceH, exportScale);
        if (factor > 1.12) {
            lowQuality.push({
                index: i,
                dni: record?.dni || '',
                factor,
                width: sourceW,
                height: sourceH
            });
        }

        if (i % 20 === 0) {
            await new Promise(r => setTimeout(r, 0));
        }
    }

    const report = {
        ok: missingPhotos.length === 0,
        total: state.records.length,
        duplicates,
        missingPhotos,
        lowQuality,
        dpi,
        widthCM,
        heightCM
    };

    state.preflightReport = report;
    renderPreflightReport(report);

    if (!opts.silent) {
        if (!report.ok) {
            showToast(`Pre-chequeo: ${missingPhotos.length} registro(s) sin foto`, 'error');
        } else if (duplicates.length || lowQuality.length) {
            showToast(`Pre-chequeo listo: ${duplicates.length} duplicados, ${lowQuality.length} con posible pixelado`, 'warning');
        } else if (opts.showToastOnPass) {
            showToast('Pre-chequeo OK: listo para exportar', 'success');
        }
    }

    return report;
}

// ===================== EXPORT PNG =====================

function getConfiguredCarnetSizeCM() {
    const widthCM = Math.max(1, Number.parseFloat(document.getElementById('pdf-width-cm')?.value) || 5.4);
    const heightCM = Math.max(1, Number.parseFloat(document.getElementById('pdf-height-cm')?.value) || 8.5);
    return { widthCM, heightCM };
}

function getExportDPI() {
    const dpiRaw = Number.parseInt(document.getElementById('export-dpi')?.value, 10);
    if (!Number.isFinite(dpiRaw)) return 300;
    return clamp(dpiRaw, 150, 1200);
}

function cmToPx(cm, dpi) {
    return Math.max(1, Math.round((cm / 2.54) * dpi));
}

// Max canvas dimension in pixels (Chrome/Electron limit is ~16 384 px per side,
// but we use 8 000 to stay well within safe memory on lower-end machines).
const MAX_CANVAS_SIDE = 8000;

function getRenderScaleForTargetPx(targetWidthPx, targetHeightPx) {
    if (!state.templateImage) return 1;
    const tw = state.templateImage.width  || 1;
    const th = state.templateImage.height || 1;
    const scaleByW = targetWidthPx  / tw;
    const scaleByH = targetHeightPx / th;
    const idealScale = Math.max(scaleByW, scaleByH);
    // Also clamp so neither canvas dimension exceeds MAX_CANVAS_SIDE
    const maxByW = MAX_CANVAS_SIDE / tw;
    const maxByH = MAX_CANVAS_SIDE / th;
    const safeMax = Math.min(maxByW, maxByH, 12);
    return clamp(idealScale, 1, safeMax);
}

async function renderCarnetAtPhysicalSize(index, widthCM, heightCM, dpi) {
    const targetW = cmToPx(widthCM, dpi);
    const targetH = cmToPx(heightCM, dpi);
    const renderScale = getRenderScaleForTargetPx(targetW, targetH);

    const renderCanvas = document.createElement('canvas');
    await renderCarnet(index, renderCanvas, renderScale);

    // Ensure exact output dimensions in pixels for the requested physical size.
    const finalCanvas = document.createElement('canvas');
    finalCanvas.width = targetW;
    finalCanvas.height = targetH;
    const fctx = finalCanvas.getContext('2d');
    fctx.imageSmoothingEnabled = true;
    fctx.imageSmoothingQuality = 'high';
    fctx.clearRect(0, 0, targetW, targetH);

    const scale = Math.min(targetW / renderCanvas.width, targetH / renderCanvas.height);
    const drawW = renderCanvas.width * scale;
    const drawH = renderCanvas.height * scale;
    const drawX = (targetW - drawW) / 2;
    const drawY = (targetH - drawH) / 2;
    fctx.drawImage(renderCanvas, drawX, drawY, drawW, drawH);

    // Free the intermediate render canvas; caller keeps only finalCanvas
    renderCanvas.width = 0;
    renderCanvas.height = 0;

    return finalCanvas;
}

function canvasToBlob(canvas, type = 'image/png', quality = 0.98) {
    return new Promise((resolve, reject) => {
        if (typeof canvas.toBlob === 'function') {
            canvas.toBlob((blob) => {
                if (!blob) {
                    reject(new Error('No se pudo generar blob del canvas'));
                    return;
                }
                resolve(blob);
            }, type, quality);
            return;
        }

        try {
            const dataUrl = canvas.toDataURL(type, quality);
            fetch(dataUrl)
                .then(r => r.blob())
                .then(resolve)
                .catch(reject);
        } catch (err) {
            reject(err);
        }
    });
}

function sanitizeFileComponent(value, fallback = 'archivo') {
    const base = String(value || fallback)
        .replace(/[<>:"/\\|?*\x00-\x1F]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    return (base || fallback).replace(/\s/g, '_').slice(0, 120);
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
}

const JOB_CANCELLED_ERROR = '__JOB_CANCELLED__';

function beginJob(label = 'job') {
    state.job.active = true;
    state.job.cancelRequested = false;
    state.job.label = label;
}

function endJob() {
    state.job.active = false;
    state.job.cancelRequested = false;
    state.job.label = '';
}

function cancelCurrentJob() {
    if (!state.job.active) return;
    state.job.cancelRequested = true;
    const textEl = document.getElementById('modal-text');
    if (textEl) textEl.textContent = 'Cancelando operación...';
    const cancelBtn = document.getElementById('modal-cancel-btn');
    if (cancelBtn) {
        cancelBtn.disabled = true;
        cancelBtn.textContent = 'Cancelando...';
    }
}

function assertJobNotCancelled() {
    if (state.job.active && state.job.cancelRequested) {
        const err = new Error(JOB_CANCELLED_ERROR);
        err.code = JOB_CANCELLED_ERROR;
        throw err;
    }
}

function isJobCancelledError(err) {
    return err && (err.code === JOB_CANCELLED_ERROR || String(err.message || '') === JOB_CANCELLED_ERROR);
}

async function exportPNG() {
    if (state.records.length === 0 || !state.templateImage) return;
    const { widthCM, heightCM } = getConfiguredCarnetSizeCM();
    const dpi = getExportDPI();

    beginJob('export-png');
    showModal('Exportando...', `Generando PNG ${widthCM.toFixed(1)}×${heightCM.toFixed(1)} cm @ ${dpi} DPI`, false);

    try {
        const check = await runPreflightCheck({ silent: true, showToastOnPass: false });
        assertJobNotCancelled();
        if (!check.ok) {
            showToast('Pre-chequeo bloqueó la exportación. Revisa registros sin foto.', 'error');
            return;
        }

        const offCanvas = await renderCarnetAtPhysicalSize(state.currentIndex, widthCM, heightCM, dpi);
        assertJobNotCancelled();
        const record = state.records[state.currentIndex];
        const dniValue = record?.dni || 'carnet';
        const pngBlob = await canvasToBlob(offCanvas, 'image/png');
        assertJobNotCancelled();

        downloadBlob(pngBlob, `carnet_${sanitizeFileComponent(dniValue)}_${dpi}dpi.png`);
        showToast('PNG descargado en alta calidad', 'success');
    } catch (err) {
        if (isJobCancelledError(err)) {
            showToast('Exportación cancelada por el usuario', 'warning');
        } else {
            showToast(`Error al exportar PNG: ${err.message || err}`, 'error');
            console.error(err);
        }
    } finally {
        hideModal();
        endJob();
    }
}

async function exportAllZIP() {
    if (state.records.length === 0 || !state.templateImage) return;
    const { widthCM, heightCM } = getConfiguredCarnetSizeCM();
    const dpi = getExportDPI();

    beginJob('export-zip');
    showModal('Generando ZIP...', `Renderizando 0 de ${state.records.length} en ${widthCM.toFixed(1)}×${heightCM.toFixed(1)} cm @ ${dpi} DPI`, true);

    try {
        await ensureJSZip();
        const check = await runPreflightCheck({ silent: true, showToastOnPass: false });
        assertJobNotCancelled();
        if (!check.ok) {
            showToast('Pre-chequeo bloqueó la exportación. Revisa registros sin foto.', 'error');
            return;
        }

        const zip = new window.JSZip();
        const folder = zip.folder('carnets');

        for (let i = 0; i < state.records.length; i++) {
            assertJobNotCancelled();
            const progress = ((i + 1) / state.records.length) * 85;
            updateModal(`Renderizando carnet ${i + 1} de ${state.records.length}`, progress);

            const canvas = await renderCarnetAtPhysicalSize(i, widthCM, heightCM, dpi);
            const record = state.records[i];
            updateModal(
                `Renderizando ${i + 1}/${state.records.length}: ${record?.apellidos || ''} ${record?.nombres || ''}`.trim(),
                ((i + 1) / state.records.length) * 85
            );

            const pngBlob = await canvasToBlob(canvas, 'image/png');
            // Free canvas GPU/CPU memory immediately after converting to blob
            canvas.width = 0;
            canvas.height = 0;

            const nameParts = [record?.dni, record?.apellidos, record?.nombres].filter(Boolean).join(' - ');
            const safeName = sanitizeFileComponent(nameParts || `registro_${i + 1}`, `registro_${i + 1}`);
            folder.file(`${String(i + 1).padStart(4, '0')}_${safeName}.png`, pngBlob);

            if (i % 3 === 0) {
                await new Promise(r => setTimeout(r, 0)); // Keep UI responsive
            }
        }

        const zipBlob = await zip.generateAsync(
            { type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } },
            (meta) => {
                assertJobNotCancelled();
                const zipProgress = 85 + (meta.percent || 0) * 0.15;
                updateModal(`Comprimiendo ZIP... ${Math.round(meta.percent || 0)}%`, zipProgress);
            }
        );

        assertJobNotCancelled();
        const fileName = `carnets_${widthCM.toFixed(1)}x${heightCM.toFixed(1)}cm_${dpi}dpi.zip`.replace(/\s/g, '');
        downloadBlob(zipBlob, fileName);
        showToast(`ZIP generado: ${state.records.length} carnets individuales`, 'success');
    } catch (err) {
        if (isJobCancelledError(err)) {
            showToast('Exportación ZIP cancelada por el usuario', 'warning');
        } else {
            showToast(`Error al generar ZIP: ${err.message || err}`, 'error');
            console.error(err);
        }
    } finally {
        hideModal();
        endJob();
    }
}

// ===================== EXPORT PDF =====================

function drawPDFCutGuides(pdf, x, y, w, h, markLengthMM = 3) {
    const mark = Math.max(1, Number.parseFloat(markLengthMM) || 3);
    pdf.setDrawColor(120, 120, 120);
    pdf.setLineWidth(0.2);

    // Main cut rectangle
    pdf.rect(x, y, w, h);

    // Top-left
    pdf.line(x - mark, y, x, y);
    pdf.line(x, y - mark, x, y);

    // Top-right
    pdf.line(x + w, y - mark, x + w, y);
    pdf.line(x + w, y, x + w + mark, y);

    // Bottom-left
    pdf.line(x - mark, y + h, x, y + h);
    pdf.line(x, y + h, x, y + h + mark);

    // Bottom-right
    pdf.line(x + w, y + h, x + w + mark, y + h);
    pdf.line(x + w, y + h, x + w, y + h + mark);
}

async function exportPDF() {
    if (state.records.length === 0 || !state.templateImage) return;
    beginJob('export-pdf');
    showModal('Generando PDF...', `Procesando carnet 0 de ${state.records.length}`, true);

    try {
        await ensureJsPDF();
        const check = await runPreflightCheck({ silent: true, showToastOnPass: false });
        assertJobNotCancelled();
        if (!check.ok) {
            showToast('Pre-chequeo bloqueó la exportación. Revisa registros sin foto.', 'error');
            return;
        }

        const { jsPDF } = window.jspdf;
        const orientation = document.getElementById('pdf-orientation').value;
        const pageSize = String(document.getElementById('pdf-page-size')?.value || 'a4').toLowerCase();
        const marginMM = Math.max(0, Number.parseFloat(document.getElementById('pdf-margin').value) || 10);
        const gapMM = Math.max(0, Number.parseFloat(document.getElementById('pdf-gap').value) || 5);
        const showCutGuides = !!document.getElementById('pdf-cut-guides')?.checked;
        const cutMarkLengthMM = Math.max(1, Number.parseFloat(document.getElementById('pdf-cut-length')?.value) || 3);
        const exportDPI = getExportDPI();

        const pdf = new jsPDF({ orientation, unit: 'mm', format: pageSize });

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const usableWidth = pageWidth - 2 * marginMM;
        const usableHeight = pageHeight - 2 * marginMM;

        // Read custom dimensions directly from the inputs (converting cm to mm)
        const carnetW = Math.max(10, (Number.parseFloat(document.getElementById('pdf-width-cm').value) || 5.4) * 10);
        const carnetH = Math.max(10, (Number.parseFloat(document.getElementById('pdf-height-cm').value) || 8.5) * 10);
        const targetCardPxW = cmToPx(carnetW / 10, exportDPI);
        const targetCardPxH = cmToPx(carnetH / 10, exportDPI);
        const pdfRenderScale = getRenderScaleForTargetPx(targetCardPxW, targetCardPxH);
        const usePngInPdf = exportDPI >= 450;
        const imageMimeType = usePngInPdf ? 'image/png' : 'image/jpeg';
        const imageFormat = usePngInPdf ? 'PNG' : 'JPEG';

        // Auto calculate how many fit per page
        const cols = Math.max(1, Math.floor((usableWidth + gapMM) / (carnetW + gapMM)));
        const rows = Math.max(1, Math.floor((usableHeight + gapMM) / (carnetH + gapMM)));
        const perPage = cols * rows;

        let slotIdx = 0;
        let isFirstPage = true;

        // Center the grid on the page
        const gridTotalW = cols * carnetW + (cols - 1) * gapMM;
        const gridTotalH = rows * carnetH + (rows - 1) * gapMM;
        const startX = marginMM + (usableWidth - gridTotalW) / 2;
        const startY = marginMM + (usableHeight - gridTotalH) / 2;

        for (let i = 0; i < state.records.length; i++) {
            assertJobNotCancelled();
            const rec = state.records[i];
            updateModal(
                `Procesando ${i + 1}/${state.records.length}: ${rec?.apellidos || ''} ${rec?.nombres || ''}`.trim(),
                ((i + 1) / state.records.length) * 100
            );

            // Render resolution based on selected DPI for sharper exports.
            const offCanvas = document.createElement('canvas');
            await renderCarnet(i, offCanvas, pdfRenderScale);

            const imgData = usePngInPdf
                ? offCanvas.toDataURL(imageMimeType)
                : offCanvas.toDataURL(imageMimeType, 0.98);

            // Free canvas memory immediately after extracting image data
            offCanvas.width = 0;
            offCanvas.height = 0;

            const col = slotIdx % cols;
            const row = Math.floor(slotIdx / cols);
            const x = startX + col * (carnetW + gapMM);
            const y = startY + row * (carnetH + gapMM);

            if (slotIdx === 0 && !isFirstPage) pdf.addPage();
            isFirstPage = false;

            pdf.addImage(imgData, imageFormat, x, y, carnetW, carnetH);
            if (showCutGuides) {
                drawPDFCutGuides(pdf, x, y, carnetW, carnetH, cutMarkLengthMM);
            }

            slotIdx++;
            if (slotIdx >= perPage) slotIdx = 0;

            if (i % 5 === 0) await new Promise(r => setTimeout(r, 10)); // Yield to UI thread
        }

        assertJobNotCancelled();
        pdf.save('carnets_masivos.pdf');
        showToast(`PDF ${pageSize.toUpperCase()} generado con ${state.records.length} carnets @ ${exportDPI} DPI`, 'success');
    } catch (err) {
        if (isJobCancelledError(err)) {
            showToast('Exportación PDF cancelada por el usuario', 'warning');
        } else {
            showToast(`Error al generar PDF: ${err.message || err}`, 'error');
            console.error(err);
        }
    } finally {
        hideModal();
        endJob();
    }
}

// ===================== PRINT =====================

async function printAll() {
    if (state.records.length === 0 || !state.templateImage) return;
    beginJob('print');
    showModal('Preparando impresión...', `Renderizando carnet 0 de ${state.records.length}`, true);

    let printWindow = null;
    try {
        const check = await runPreflightCheck({ silent: true, showToastOnPass: false });
        assertJobNotCancelled();
        if (!check.ok) {
            showToast('Pre-chequeo bloqueó la impresión. Revisa registros sin foto.', 'error');
            return;
        }

        // Adjust max-width based on user input for CM
        const customW = document.getElementById('pdf-width-cm');
        const maxWidthMM = Math.max(10, (Number.parseFloat(customW?.value) || 5.4) * 10);

        printWindow = window.open('', '_blank');
        if (!printWindow) {
            showToast('El navegador bloqueó la ventana de impresión. Permite ventanas emergentes e inténtalo otra vez.', 'error');
            return;
        }
        printWindow.document.write(`<html><head><title>Carnets — Impresión</title>
            <style>
                body { margin: 0; padding: 10mm; font-family: Arial; text-align: center; }
                .carnet-wrapper { display: inline-block; margin: 3mm; page-break-inside: avoid; }
                .carnet-img { max-width: ${maxWidthMM}mm; border: 1px dotted #ccc; }
                @media print { body { padding: 5mm; } .carnet-img { border: none; } }
            </style></head><body>`);

        for (let i = 0; i < state.records.length; i++) {
            assertJobNotCancelled();
            updateModal(`Renderizando carnet ${i + 1} de ${state.records.length}`, ((i + 1) / state.records.length) * 100);
            
            // Render in high-res (3x scale)
            const offCanvas = document.createElement('canvas');
            await renderCarnet(i, offCanvas, 3);
            
            // Use JPEG 0.95 to keep browser memory usage low
            printWindow.document.write(`
                <div class="carnet-wrapper">
                    <img src="${offCanvas.toDataURL('image/jpeg', 0.95)}" class="carnet-img">
                </div>
            `);
            
            if (i % 5 === 0) await new Promise(r => setTimeout(r, 10)); // Yield to UI
        }

        assertJobNotCancelled();
        printWindow.document.write('</body></html>');
        printWindow.document.close();

        // Wait for images to load before calling print()
        printWindow.onload = () => {
            setTimeout(() => printWindow.print(), 200);
        };
        showToast('Diálogo de impresión preparado', 'info');
    } catch (err) {
        if (isJobCancelledError(err)) {
            if (printWindow && !printWindow.closed) printWindow.close();
            showToast('Impresión cancelada por el usuario', 'warning');
        } else {
            showToast(`Error al preparar impresión: ${err.message || err}`, 'error');
            console.error(err);
        }
    } finally {
        hideModal();
        endJob();
    }
}

// ===================== MODAL =====================

function showModal(title, text, cancellable = false) {
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-text').textContent = text;
    document.getElementById('progress-fill').style.width = '0%';
    const cancelBtn = document.getElementById('modal-cancel-btn');
    if (cancelBtn) {
        cancelBtn.style.display = cancellable ? 'inline-flex' : 'none';
        cancelBtn.disabled = false;
        cancelBtn.textContent = 'Cancelar exportación';
    }
    document.getElementById('modal-loading').classList.add('active');
}

function updateModal(text, percent) {
    document.getElementById('modal-text').textContent = text;
    document.getElementById('progress-fill').style.width = `${clamp(toFloat(percent, 0), 0, 100)}%`;
}

function hideModal() {
    const cancelBtn = document.getElementById('modal-cancel-btn');
    if (cancelBtn) {
        cancelBtn.style.display = 'none';
        cancelBtn.disabled = false;
        cancelBtn.textContent = 'Cancelar exportación';
    }
    document.getElementById('modal-loading').classList.remove('active');
}

// ===================== TOAST =====================

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    const icons = { success: '✅', error: '❌', info: 'ℹ️', warning: '⚠️' };
    toast.innerHTML = `<span>${icons[type] || 'ℹ️'}</span> ${escapeHtml(message)}`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

// ===================== UTILS =====================

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}


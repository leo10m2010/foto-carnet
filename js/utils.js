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
        bgColor: normalizeHexColor(document.getElementById('field-photo-bg-color')?.value, '#d9dee8'),
        rotation: toFloat(document.getElementById('field-photo-rotation')?.value, 0)
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
        bgColor: normalizeHexColor(config.bgColor, '#d9dee8'),
        rotation: toFloat(config.rotation, 0)
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

// ===================== UTILS =====================

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}


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
    templateDataUrl: null,    // Base64 dataURL of template (saved in session for reliable restore)
    records: [],              // Array of { dni, nombres, apellidos, extra, hasPhoto }
    photosMap: {},            // { "07971267": objectURL/dataURL/filePath, ... }
    photoPaths: {},           // { "07971267": filePath } — for session restore
    photoObjectUrls: [],      // Temporary object URLs to revoke on reload
    photoImageCache: new LRUImageCache(80), // LRU — keeps last 80 decoded images in memory
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
        panSessionUntil: 0,
        rotationSessionUntil: 0
    },
    job: {
        active: false,
        cancelRequested: false,
        label: ''
    },
    reniecGeneration: 0,   // Incremented on every photo reload; aborts stale RENIEC queries
    filmstripVisible: true
};

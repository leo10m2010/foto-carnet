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
    return getPhotoConfigForRecord(getCurrentRecord() || {});
}

function getPhotoConfigForRecord(record) {
    const key = getRecordKey(record);
    const override = key ? state.photoOverrides[key] : null;
    if (override) return normalizePhotoConfig({ ...override });
    if (state.globalPhotoConfig) return normalizePhotoConfig({ ...state.globalPhotoConfig });
    return normalizePhotoConfig(readPhotoConfigFromInputs());
}

function getBarcodeConfig() {
    const v = (id, fb) => document.getElementById(id)?.value ?? fb;
    return {
        x: Math.max(0, toInt(v('field-barcode-x', 0), 0)),
        y: Math.max(0, toInt(v('field-barcode-y', 0), 0)),
        w: Math.max(40, toInt(v('field-barcode-w', 40), 40)),
        h: Math.max(20, toInt(v('field-barcode-h', 20), 20)),
        format: v('field-barcode-format', 'CODE128'),
        showText: v('field-barcode-showtext', 'false') === 'true'
    };
}

function tryRender() {
    if (!state.templateImage || state.records.length === 0) return;
    renderCarnet(state.currentIndex).then(() => {
        if (!state.drag.active) drawSelectionOverlay();
        updateEditorHud();
    });
    if (!state.job.active) refreshFilmstripDebounced();
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
            try {
                // Convert data URL → blob URL without going through fetch/HTTP stack
                const [header, b64] = result.dataUrl.split(',');
                const mime = header.match(/:(.*?);/)[1];
                const bytes = atob(b64);
                const arr = new Uint8Array(bytes.length);
                for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
                const objUrl = URL.createObjectURL(new Blob([arr], { type: mime }));
                state.photosMap[key] = objUrl;
                state.photoObjectUrls.push(objUrl);
                source = objUrl;
            } catch (_) {
                state.photosMap[key] = result.dataUrl;
                source = result.dataUrl;
            }
        } else {
            // IPC failed — try file:// URL as last resort (works for local files in Electron)
            source = 'file:///' + source.replace(/\\/g, '/').replace(/^\/+/, '');
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

    const angleDeg = photoConfig.rotation || 0;
    const cx = rect.frameX + rect.frameW / 2;
    const cy = rect.frameY + rect.frameH / 2;

    ctx.save();
    // Clip to frame first so rotation stays inside the frame boundary
    ctx.beginPath();
    ctx.rect(rect.frameX, rect.frameY, rect.frameW, rect.frameH);
    ctx.clip();

    if (angleDeg !== 0) {
        ctx.translate(cx, cy);
        ctx.rotate(angleDeg * Math.PI / 180);
        ctx.translate(-cx, -cy);
    }

    ctx.drawImage(photoImg, rect.drawX, rect.drawY, rect.drawW, rect.drawH);
    ctx.restore();
}

function samplePhotoPixel(photoImg, x, y) {
    const sourceW = photoImg.naturalWidth || photoImg.width;
    const sourceH = photoImg.naturalHeight || photoImg.height;
    if (!sourceW || !sourceH) return null;

    // Sample at reduced resolution — color accuracy is fine at 256px and avoids
    // allocating a full multi-megapixel canvas for a single pixel read.
    const sampleSize = 256;
    const scale = Math.min(1, sampleSize / Math.max(sourceW, sourceH));
    const sw = Math.max(1, Math.round(sourceW * scale));
    const sh = Math.max(1, Math.round(sourceH * scale));

    const off = document.createElement('canvas');
    off.width  = sw;
    off.height = sh;
    const octx = off.getContext('2d', { willReadFrequently: true });
    octx.drawImage(photoImg, 0, 0, sw, sh);

    const sx = clamp(Math.floor(x * scale), 0, sw - 1);
    const sy = clamp(Math.floor(y * scale), 0, sh - 1);
    const rgba = octx.getImageData(sx, sy, 1, 1).data;
    return rgbToHex(rgba[0], rgba[1], rgba[2]);
}

// Sample multiple relative points [rx, ry] from a photo in one canvas draw pass.
function samplePhotoColors(photoImg, points) {
    const sourceW = photoImg.naturalWidth || photoImg.width;
    const sourceH = photoImg.naturalHeight || photoImg.height;
    if (!sourceW || !sourceH) return points.map(() => '#d9dee8');

    const sampleSize = 256;
    const scale = Math.min(1, sampleSize / Math.max(sourceW, sourceH));
    const sw = Math.max(1, Math.round(sourceW * scale));
    const sh = Math.max(1, Math.round(sourceH * scale));

    const off = document.createElement('canvas');
    off.width = sw;
    off.height = sh;
    const octx = off.getContext('2d', { willReadFrequently: true });
    octx.drawImage(photoImg, 0, 0, sw, sh);

    return points.map(([rx, ry]) => {
        const sx = clamp(Math.floor(sw * rx), 0, sw - 1);
        const sy = clamp(Math.floor(sh * ry), 0, sh - 1);
        const rgba = octx.getImageData(sx, sy, 1, 1).data;
        return normalizeHexColor(rgbToHex(rgba[0], rgba[1], rgba[2]), '#d9dee8');
    });
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

    const samplePoints = [
        [0.12, 0.10], [0.5, 0.08], [0.88, 0.10],
        [0.18, 0.22], [0.82, 0.22], [0.5, 0.18]
    ];

    const colors = samplePhotoColors(photoImg, samplePoints);
    let r = 0, g = 0, b = 0;
    colors.forEach(c => {
        r += Number.parseInt(c.slice(1, 3), 16);
        g += Number.parseInt(c.slice(3, 5), 16);
        b += Number.parseInt(c.slice(5, 7), 16);
    });

    const picked = rgbToHex(r / colors.length, g / colors.length, b / colors.length);
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

    const points = [
        [0.1, 0.1], [0.5, 0.08], [0.9, 0.1], [0.25, 0.2], [0.75, 0.2], [0.5, 0.3]
    ];
    const colors = samplePhotoColors(photoImg, points);
    const unique = colors.filter((c, i, arr) => arr.indexOf(c) === i);

    container.innerHTML = '';
    unique.slice(0, 6).forEach(color => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'swatch-btn';
        btn.style.background = color;
        btn.title = color;
        btn.addEventListener('click', () => setPhotoBgColor(color));
        container.appendChild(btn);
    });
}

// Draws a text field on ctx, applying truncation and registering a hitbox.
function drawTextField(ctx, text, cfg, id, targetCanvas) {
    const fontStr = `${cfg.bold} ${cfg.size}px ${cfg.font}`.trim();
    ctx.font = fontStr;
    ctx.fillStyle = cfg.color;
    ctx.textAlign = cfg.align;
    ctx.textBaseline = 'top';

    let displayText = text;
    let textW = ctx.measureText(displayText).width;
    if (textW > cfg.maxWidth) {
        while (ctx.measureText(displayText + '…').width > cfg.maxWidth && displayText.length > 0) {
            displayText = displayText.slice(0, -1);
        }
        displayText += '…';
        textW = ctx.measureText(displayText).width;
    }
    let hitX = cfg.x;
    if (cfg.align === 'center') hitX = cfg.x - textW / 2;
    else if (cfg.align === 'right') hitX = cfg.x - textW;

    ctx.fillText(displayText, cfg.x, cfg.y);

    if (!targetCanvas) {
        state.hitboxes.push({ id, x: hitX, y: cfg.y, w: textW, h: cfg.size });
    }
}

function renderCarnet(index, targetCanvas, exportScale = 1) {
    return new Promise((resolve) => {
        const record = state.records[index];
        if (!record) { resolve(null); return; }

        const template = state.templateImage;
        const canvas = targetCanvas || document.getElementById('carnet-canvas');
        const ctx = canvas.getContext('2d');

        const newW = template.width * exportScale;
        const newH = template.height * exportScale;
        if (canvas.width !== newW || canvas.height !== newH) {
            canvas.width  = newW;  // assigning either dimension resets the context
            canvas.height = newH;
        } else {
            ctx.clearRect(0, 0, newW, newH);
        }

        ctx.resetTransform();
        if (exportScale !== 1) {
            ctx.scale(exportScale, exportScale);
            ctx.imageSmoothingEnabled = true;
            ctx.imageSmoothingQuality = 'high';
        }

        const photoConfig = getPhotoConfigForRecord(record);
        const photoKey = getRecordKey(record);

        const drawTextsAndBarcode = () => {
            // --- PHOTO hitbox (for drag-and-drop) ---
            if (!targetCanvas) {
                state.hitboxes.push({
                    id: 'photo',
                    x: photoConfig.x, y: photoConfig.y,
                    w: photoConfig.w, h: photoConfig.h
                });
            }

            if (record.nombres)  drawTextField(ctx, record.nombres,  getFieldConfig('nombres'),  'nombres',  targetCanvas);
            if (record.apellidos) drawTextField(ctx, record.apellidos, getFieldConfig('apellidos'), 'apellidos', targetCanvas);
            if (record.dni) {
                const prefix = document.getElementById('field-dni-prefix')?.value || '';
                drawTextField(ctx, prefix + record.dni, getFieldConfig('dni'), 'dni', targetCanvas);
            }
            if (record.extra)    drawTextField(ctx, record.extra,    getFieldConfig('extra'),    'extra',    targetCanvas);

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

            if (!targetCanvas) updateNavigation();
            resolve(canvas);
        };

        // Reset hitboxes before rendering
        if (!targetCanvas) state.hitboxes = [];


        // Photo renders behind the template so transparent areas show through.
        const drawPhotoThenTemplate = (photoImg) => {
            if (photoConfig.bgEnabled) {
                ctx.save();
                ctx.fillStyle = photoConfig.bgColor;
                ctx.fillRect(photoConfig.x, photoConfig.y, photoConfig.w, photoConfig.h);
                ctx.restore();
            }

            if (photoImg) {
                drawPhotoInFrame(ctx, photoImg, photoConfig);
            } else {
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

            ctx.drawImage(template, 0, 0);
            drawTextsAndBarcode();
        };

        // Use getPhotoImageByKey so session-restored file paths load via IPC automatically
        getPhotoImageByKey(photoKey).then(photoImg => {
            drawPhotoThenTemplate(photoImg);
        }).catch(() => {
            // On any unexpected error, render without photo (avoids hanging Promise)
            drawPhotoThenTemplate(null);
        });
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


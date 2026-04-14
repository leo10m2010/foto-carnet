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
        state.drag.snapGuides = { xLines: [], yLines: [] };

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

        state.drag.snapGuides = { xLines: [], yLines: [] };

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

            const canvasW = state.templateImage.width;
            const canvasH = state.templateImage.height;
            const centerX = canvasW / 2;
            const centerY = canvasH / 2;
            const elemCenterX = newVisX + state.drag.startElemW / 2;
            const elemCenterY = newVisY + state.drag.startElemH / 2;
            const elemR = newVisX + state.drag.startElemW;
            const elemB = newVisY + state.drag.startElemH;

            const snapThreshold = 12;

            // Canvas center snapping
            if (Math.abs(elemCenterX - centerX) < snapThreshold) {
                dx = centerX - (state.drag.startElemX + state.drag.startElemW / 2);
                state.drag.snapGuides.xLines.push(centerX);
            }
            if (Math.abs(elemCenterY - centerY) < snapThreshold) {
                dy = centerY - (state.drag.startElemY + state.drag.startElemH / 2);
                state.drag.snapGuides.yLines.push(centerY);
            }

            // Canvas edge snapping (all elements)
            if (Math.abs(newVisX) < snapThreshold) {
                dx = -state.drag.startElemX;
                state.drag.snapGuides.xLines.push(0);
            } else if (Math.abs(elemR - canvasW) < snapThreshold) {
                dx = canvasW - state.drag.startElemW - state.drag.startElemX;
                state.drag.snapGuides.xLines.push(canvasW);
            }
            if (Math.abs(newVisY) < snapThreshold) {
                dy = -state.drag.startElemY;
                state.drag.snapGuides.yLines.push(0);
            } else if (Math.abs(elemB - canvasH) < snapThreshold) {
                dy = canvasH - state.drag.startElemH - state.drag.startElemY;
                state.drag.snapGuides.yLines.push(canvasH);
            }

            // Cross-element snapping (align to other elements' edges & centers)
            state.hitboxes.filter(h => h.id !== id).forEach(other => {
                const newX = state.drag.startElemX + dx;
                const newCX = newX + state.drag.startElemW / 2;
                const otherCX = other.x + other.w / 2;
                if (Math.abs(newX - other.x) < snapThreshold) {
                    dx = other.x - state.drag.startElemX;
                    state.drag.snapGuides.xLines.push(other.x);
                } else if (Math.abs(newCX - otherCX) < snapThreshold) {
                    dx = otherCX - (state.drag.startElemX + state.drag.startElemW / 2);
                    state.drag.snapGuides.xLines.push(otherCX);
                }
                const newY = state.drag.startElemY + dy;
                const newCY = newY + state.drag.startElemH / 2;
                const otherCY = other.y + other.h / 2;
                if (Math.abs(newY - other.y) < snapThreshold) {
                    dy = other.y - state.drag.startElemY;
                    state.drag.snapGuides.yLines.push(other.y);
                } else if (Math.abs(newCY - otherCY) < snapThreshold) {
                    dy = otherCY - (state.drag.startElemY + state.drag.startElemH / 2);
                    state.drag.snapGuides.yLines.push(otherCY);
                }
            });

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
        ctx.strokeStyle = '#ec4899';
        ctx.lineWidth = 1.5;
        ctx.setLineDash([6, 4]);
        (state.drag.snapGuides.xLines || []).forEach(x => {
            ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, canvas.height); ctx.stroke();
        });
        (state.drag.snapGuides.yLines || []).forEach(y => {
            ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(canvas.width, y); ctx.stroke();
        });
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

        // Hover fill tint
        if (isHovered) {
            ctx.fillStyle = 'rgba(99, 102, 241, 0.07)';
            ctx.fillRect(hb.x, hb.y, hb.w, hb.h);
        }

        // Outer glow ring (selected only)
        if (isSelected) {
            ctx.strokeStyle = 'rgba(99, 102, 241, 0.22)';
            ctx.lineWidth = 7;
            ctx.setLineDash([]);
            ctx.strokeRect(hb.x - pad - 3, hb.y - pad - 3, hb.w + (pad + 3) * 2, hb.h + (pad + 3) * 2);
        }

        // Border
        ctx.strokeStyle = color;
        ctx.lineWidth = lw;
        ctx.setLineDash(isSelected ? [] : [5, 5]);
        ctx.strokeRect(hb.x - pad, hb.y - pad, hb.w + pad * 2, hb.h + pad * 2);
        ctx.setLineDash([]);

        // Corner handles — circular
        if (isSelected) {
            const r = 5;
            const canResize = (hb.id === 'photo' || hb.id === 'barcode');
            const corners = [
                [hb.x - pad, hb.y - pad],
                [hb.x + hb.w + pad, hb.y - pad],
                [hb.x - pad, hb.y + hb.h + pad],
                [hb.x + hb.w + pad, hb.y + hb.h + pad]
            ];
            corners.forEach(([cx, cy]) => {
                ctx.beginPath();
                ctx.arc(cx, cy, r, 0, Math.PI * 2);
                ctx.fillStyle = canResize ? '#ffffff' : color;
                ctx.fill();
                ctx.strokeStyle = '#6366f1';
                ctx.lineWidth = 2;
                ctx.stroke();
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
            'field-photo-bg-enable', 'field-photo-bg-color', 'field-photo-rotation'
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

function rotatePhoto(deltaDeg, reset = false) {
    if (state.drag.selectedId !== 'photo') return;
    const input = document.getElementById('field-photo-rotation');
    if (!input) return;
    pushUndoSnapshot('photo-rotate');
    const current = toFloat(input.value, 0);
    let next = reset ? 0 : current + deltaDeg;
    // Keep within -180..180 for the slider range
    next = ((next + 540) % 360) - 180;
    input.value = next;
    invalidatePreflightReport();
    savePhotoConfigFromDOM();
    syncHudPhotoControls(getPhotoConfig());
    tryRender();
}

function setSelectedPhotoRotation(value) {
    if (state.drag.selectedId !== 'photo') return;
    const input = document.getElementById('field-photo-rotation');
    if (!input) return;
    const now = Date.now();
    // Debounced snapshot — one undo entry per drag session (same as zoom)
    if (now > (state.history.rotationSessionUntil || 0)) {
        pushUndoSnapshot('photo-rotate');
        state.history.rotationSessionUntil = now + 500;
    }
    input.value = value;
    invalidatePreflightReport();
    savePhotoConfigFromDOM();
    syncHudPhotoControls(getPhotoConfig());
    tryRender();
}

function showSidebarNameEditor() {
    const panel = document.getElementById('record-name-editor');
    if (!panel) return;
    const record = state.records[state.currentIndex];
    if (!record) { panel.style.display = 'none'; return; }
    document.getElementById('sidebar-nombres').value   = record.nombres   || '';
    document.getElementById('sidebar-apellidos').value = record.apellidos || '';
    panel.style.display = 'block';
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function commitSidebarNameEdit() {
    if (!state.records.length) return;
    const nombres   = (document.getElementById('sidebar-nombres')?.value   || '').trim();
    const apellidos = (document.getElementById('sidebar-apellidos')?.value || '').trim();
    pushUndoSnapshot('sidebar-name-edit');
    const record = state.records[state.currentIndex];
    record.nombres   = nombres;
    record.apellidos = apellidos;
    invalidatePreflightReport();
    showDataPreview();
    saveSessionDebounced();
    tryRender();
    showToast('Nombre actualizado', 'success');
}

function cancelSidebarNameEdit() {
    showSidebarNameEditor(); // re-sync inputs to current record — no state change
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
    if (typeof lucide !== 'undefined') lucide.createIcons();
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
        ['hud-rotate-ccw', 'hud-rotate-cw', 'hud-rotate-reset', 'hud-photo-rotation'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.disabled = disabled;
        });
    };

    if (!hasRenderableData) {
        hud.classList.remove('active', 'crop-mode', 'individual-mode');
        actionButtons.forEach(btn => { btn.disabled = true; });
        setPhotoHudDisabled(true);
        if (photoControls) photoControls.classList.remove('active');
        if (swatches) swatches.classList.remove('active');
        if (nameEl) nameEl.textContent = 'Sin selección';
        if (detailsEl) detailsEl.textContent = 'Carga plantilla y fotos para editar';
        return;
    }

    hud.classList.add('active');
    hud.classList.toggle('crop-mode', state.photoCropMode.active && state.drag.selectedId === 'photo');

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
        hud.classList.remove('individual-mode');
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
    if (id !== 'photo') hud.classList.remove('individual-mode');
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


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

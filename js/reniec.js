const RENIEC_TOKEN = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJlbWFpbCI6InRoZXJlbmF0b2VtQGdtYWlsLmNvbSJ9.CGxlSecmJ6IQP7esT6LtQNwQ64u9q_rOYou0qyXdQt0';

async function enrichWithRENIEC() {
    // Capture the generation token at the moment this query starts.
    // If the user reloads photos, state.reniecGeneration increments and
    // every check below will abort — preventing stale data from being written.
    const myGeneration = state.reniecGeneration;
    const isStale = () => state.reniecGeneration !== myGeneration;

    const toEnrich = state.records.filter(r => /^\d{8}$/.test(getRecordKey(r)) && !r.reniecOk);
    if (toEnrich.length === 0) return;

    // Build a fast lookup: dniKey -> index in state.records (avoids O(n²) findIndex)
    const dniIndexMap = new Map(state.records.map((r, i) => [getRecordKey(r), i]));

    showToast(`Verificando ${toEnrich.length} DNI${toEnrich.length > 1 ? 's' : ''} en RENIEC…`, 'info');
    updateReniecStatChip('…');

    let ok = 0, notFound = 0, errors = 0;

    for (let i = 0; i < toEnrich.length; i++) {
        if (i > 0) await new Promise(r => setTimeout(r, 200)); // ≈ 5 req/s, skip on first
        if (isStale()) return;

        const record = toEnrich[i];
        const dni = getRecordKey(record);
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
                const resp = await fetch(
                    `https://dniruc.apisperu.com/api/v1/dni/${dni}?token=${RENIEC_TOKEN}`
                );
                json = await resp.json();
            }

            // Check again after the await — a reload could have happened during the request
            if (isStale()) return;
            // Guard against undo/redo replacing state.records between awaits:
            // if the record at idx no longer matches our DNI, skip to avoid a wrong write
            if (idx >= state.records.length || getRecordKey(state.records[idx]) !== dni) continue;

            // apisperu.com: success response has `nombres` field directly; failure has success:false
            if (json && json.nombres && json.success !== false) {
                const nombres   = (json.nombres || '').trim();
                const apellidos = `${(json.apellidoPaterno || '')} ${(json.apellidoMaterno || '')}`.trim();
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

        // Update chip counter every 5 records (cheap text update, no DOM rebuild)
        if (i % 5 === 4 || i === toEnrich.length - 1) {
            if (isStale()) return;
            updateReniecStatChip(`${ok}/${toEnrich.length}`);
        }
    }

    if (isStale()) return;

    showDataPreview();
    tryRender();
    updateReniecStatChip(`${ok}/${toEnrich.length}`);
    updateStatusBar();

    const corrected = state.records.filter(r => r.reniecOk && r.filenameNombres &&
        (r.filenameNombres.toUpperCase() !== r.reniecNombres?.toUpperCase() ||
         r.filenameApellidos?.toUpperCase() !== r.reniecApellidos?.toUpperCase())).length;

    let msg = `RENIEC: ${ok} verificados`;
    if (corrected > 0) msg += `, ${corrected} nombre${corrected > 1 ? 's' : ''} corregido${corrected > 1 ? 's' : ''}`;
    if (notFound > 0)  msg += `, ${notFound} no encontrado${notFound > 1 ? 's' : ''}`;
    if (errors > 0)    msg += `, ${errors} con error`;

    showToast(msg, ok > 0 ? 'success' : 'warning');

    // After enrichment: detect if filename parser had nombres/apellidos swapped
    // and auto-repair unverified records if the swap pattern is consistent.
    detectAndFixNameSwap();
    updateFilmstripTooltips(); // Update card titles without rebuilding the DOM

    saveSession(); // Persist RENIEC-enriched names
}

// ---- Detect & fix nombres↔apellidos swap ----
// Compares what the filename parser produced vs what RENIEC says is correct.
// If ≥50% of RENIEC-verified records had the fields inverted, fixes all unverified ones.
function detectAndFixNameSwap() {
    const verified = state.records.filter(r =>
        r.reniecOk === true &&
        r.filenameNombres !== undefined &&
        r.reniecNombres   !== undefined &&
        r.reniecApellidos !== undefined
    );
    if (verified.length < 2) return; // Not enough data to detect a pattern

    function norm(s) {
        return (s || '').toUpperCase().trim().replace(/\s+/g, ' ');
    }

    // Word-intersection score (0–1): how many words of 'a' appear in 'b'
    function matchScore(a, b) {
        const wa = norm(a).split(' ').filter(Boolean);
        const wb = norm(b).split(' ').filter(Boolean);
        if (!wa.length || !wb.length) return 0;
        const common = wa.filter(w => wb.includes(w)).length;
        return common / Math.max(wa.length, wb.length);
    }

    let swapCount = 0;
    for (const r of verified) {
        // A "swap" means: what was stored as nombres ≈ RENIEC's apellidos (and vice versa)
        const nombresWasApellidos = matchScore(r.filenameNombres,   r.reniecApellidos) >= 0.5;
        const apellidosWasNombres = matchScore(r.filenameApellidos, r.reniecNombres)   >= 0.5;
        if (nombresWasApellidos && apellidosWasNombres) swapCount++;
    }

    if (swapCount / verified.length < 0.5) return; // No consistent pattern — don't touch anything

    // Consistent swap detected: fix all records that RENIEC didn't verify
    let fixed = 0;
    state.records.forEach(r => {
        if (r.reniecOk !== true) {
            const tmp  = r.nombres;
            r.nombres  = r.apellidos;
            r.apellidos = tmp;
            fixed++;
        }
    });

    if (fixed > 0) {
        showDataPreview();
        tryRender();
        saveSessionDebounced();
        showToast(
            `Se detectó formato invertido: nombres↔apellidos corregidos en ${fixed} registro${fixed !== 1 ? 's' : ''}.`,
            'info'
        );
    }
}

function updateReniecStatChip(text) {
    const chip = document.getElementById('chip-reniec');
    const el   = document.getElementById('stat-reniec');
    if (chip) chip.style.display = '';
    if (el)   el.textContent = text;
}

// ---- CSV / Excel (OPTIONAL — for extra fields like cargo) ----

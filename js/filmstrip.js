// ===================== FILMSTRIP NAVIGATOR =====================

let _filmstripObserver = null;
let _filmstripRefreshTimer = null;

function getRecordDisplayName(record) {
    return [record.apellidos, record.nombres].filter(Boolean).join(' ').trim();
}

function initFilmstrip() {
    const saved = localStorage.getItem('filmstrip_visible');
    state.filmstripVisible = saved !== 'false'; // default: visible
    applyFilmstripVisibility(false);
}

function toggleFilmstrip() {
    state.filmstripVisible = !state.filmstripVisible;
    localStorage.setItem('filmstrip_visible', String(state.filmstripVisible));
    applyFilmstripVisibility(true);
    if (state.filmstripVisible) {
        setTimeout(scrollFilmstripToActive, 260);
    }
}

function applyFilmstripVisibility(animate) {
    const bar = document.getElementById('filmstrip-bar');
    if (!bar) return;
    if (!animate) bar.classList.add('no-transition');
    bar.classList.toggle('is-open', !!state.filmstripVisible);
    const arrow = document.getElementById('filmstrip-arrow');
    if (arrow) arrow.textContent = state.filmstripVisible ? '▼' : '▲';
    if (!animate) {
        bar.offsetHeight; // force reflow before re-enabling transitions
        bar.classList.remove('no-transition');
    }
}

function renderFilmstrip() {
    const track = document.getElementById('filmstrip-track');
    if (!track) return;

    if (_filmstripObserver) {
        _filmstripObserver.disconnect();
        _filmstripObserver = null;
    }

    track.innerHTML = '';

    const countEl = document.getElementById('filmstrip-count');
    if (countEl) countEl.textContent = state.records.length ? `${state.records.length} fotos` : '';

    if (!state.records.length) return;

    const fragment = document.createDocumentFragment();
    state.records.forEach((record, i) => {
        const card = document.createElement('div');
        card.className = 'filmstrip-thumb' + (i === state.currentIndex ? ' is-active' : '');
        card.dataset.index = i;
        card.title = `${i + 1}. ${getRecordDisplayName(record)}`;
        card.addEventListener('click', () => goToFilmstripIndex(i));

        card.innerHTML =
            `<div class="thumb-img-wrap">` +
                `<canvas class="thumb-canvas" data-index="${i}"></canvas>` +
                `<span class="thumb-num">${i + 1}</span>` +
            `</div>` +
            `<div class="thumb-name">${i + 1}</div>`;

        fragment.appendChild(card);
    });

    track.appendChild(fragment);

    const container = document.getElementById('filmstrip-container');
    _filmstripObserver = new IntersectionObserver(
        (entries) => {
            entries.forEach(entry => {
                if (!entry.isIntersecting) return;
                const canvas = entry.target;
                if (canvas.dataset.loaded) return;
                canvas.dataset.loaded = '1';
                _filmstripObserver.unobserve(canvas);
                _loadThumb(canvas);
            });
        },
        { root: container, rootMargin: '0px 320px 0px 320px', threshold: 0 }
    );

    track.querySelectorAll('.thumb-canvas').forEach(canvas => _filmstripObserver.observe(canvas));

    if (state.filmstripVisible) requestAnimationFrame(scrollFilmstripToActive);
}

function _loadThumb(canvas) {
    const i = Number(canvas.dataset.index);
    if (!state.templateImage || i < 0 || i >= state.records.length) return;

    // Render at ~96px wide (2× display density for sharpness) into the thumb canvas
    const scale = Math.max(0.04, 96 / (state.templateImage.width || 630));
    renderCarnet(i, canvas, scale).catch(() => {});
}

function updateFilmstripActive() {
    const track = document.getElementById('filmstrip-track');
    if (!track || !track.children.length) return;
    const prev = track.querySelector('.filmstrip-thumb.is-active');
    if (prev) prev.classList.remove('is-active');
    const next = track.children[state.currentIndex];
    if (next) next.classList.add('is-active');
}

function scrollFilmstripToActive() {
    const container = document.getElementById('filmstrip-container');
    const track = document.getElementById('filmstrip-track');
    if (!container || !track) return;
    const thumb = track.children[state.currentIndex];
    if (!thumb) return;
    container.scrollLeft = thumb.offsetLeft - container.clientWidth / 2 + thumb.offsetWidth / 2;
}

// Re-renders visible thumbnails after a debounce — called from tryRender on any edit.
function refreshFilmstripDebounced() {
    clearTimeout(_filmstripRefreshTimer);
    _filmstripRefreshTimer = setTimeout(_refreshVisibleThumbs, 600);
}

function _refreshVisibleThumbs() {
    const track = document.getElementById('filmstrip-track');
    const container = document.getElementById('filmstrip-container');
    if (!track || !container || !state.filmstripVisible || !state.templateImage) return;
    // Re-render only canvases currently in the scroll viewport (not just historically loaded).
    const { left: cLeft, right: cRight } = container.getBoundingClientRect();
    track.querySelectorAll('.thumb-canvas[data-loaded="1"]').forEach(canvas => {
        const { left, right } = canvas.getBoundingClientRect();
        if (right >= cLeft && left <= cRight) _loadThumb(canvas);
    });
}

function updateFilmstripTooltips() {
    const track = document.getElementById('filmstrip-track');
    if (!track) return;
    Array.from(track.children).forEach((card, i) => {
        const record = state.records[i];
        if (!record) return;
        card.title = `${i + 1}. ${getRecordDisplayName(record)}`;
    });
}

function goToFilmstripIndex(i) {
    if (i < 0 || i >= state.records.length || i === state.currentIndex) return;
    navigateRecord(i - state.currentIndex);
}

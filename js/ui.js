
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
    const duration = type === 'error' ? Math.max(4000, message.length * 55) : 3000;
    setTimeout(() => toast.remove(), duration);
}


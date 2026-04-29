export const $ = id => document.getElementById(id);

export function escHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

export function toast(message, type = 'info') {
    const area = $('toast-area');
    if (!area) return;
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.textContent = message;
    area.appendChild(el);
    setTimeout(() => el.remove(), 4000);
}

export function showError(message) {
    const banner = $('error-banner');
    if (!banner) return;
    banner.textContent = message;
    banner.style.display = 'block';
    setTimeout(() => { banner.style.display = 'none'; }, 8000);
}

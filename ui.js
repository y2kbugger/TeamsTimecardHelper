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

let errorBannerTimer = 0;

// The newest error shows its full time; an older one's timer does not cut it short.
export function showError(message) {
    const banner = $('error-banner');
    if (!banner) return;
    banner.textContent = message;
    banner.style.display = 'block';
    clearTimeout(errorBannerTimer);
    errorBannerTimer = setTimeout(() => { banner.style.display = 'none'; }, 8000);
}

// ── Dev server side channel (devserver.py): only on localhost, never in production ──

const isDevServer = location.hostname === 'localhost';

function postToDevServer(path, text) {
    if (!isDevServer) return;
    void fetch(path, { method: 'POST', body: text, keepalive: true }).catch(() => { });
}

// A condensed report written to .agent/<name>.txt, for reading from the shell.
export function reportToDevServer(name, text) {
    postToDevServer(`/__agent/report?name=${encodeURIComponent(name)}`, text);
}

// Uncaught errors, rejections and console.error, appended to .agent/browser.log.
if (isDevServer && !window.__devLogInstalled) {
    window.__devLogInstalled = true;
    const log = (kind, detail) => postToDevServer('/__agent/log', `${new Date().toISOString()} ${location.pathname} ${kind}: ${detail}`);
    window.addEventListener('error', event => log('error', `${event.message} at ${event.filename}:${event.lineno}`));
    window.addEventListener('unhandledrejection', event => log('rejection', event.reason?.stack || event.reason));
    const consoleError = console.error.bind(console);
    console.error = (...args) => {
        log('console.error', args.map(arg => arg?.stack || (typeof arg === 'string' ? arg : JSON.stringify(arg))).join(' '));
        consoleError(...args);
    };
}

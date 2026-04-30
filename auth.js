import { config } from './config.js';
import { $, escHtml, toast, showError } from './ui.js';

const DEFAULT_AUTHORITY_TENANT = 'organizations';

let msalInstance = null;
let currentAccount = null;
let popupTokenPromise = null;
let initializedConfigKey = null;
let initializeSessionPromise = null;

export function isPopupContext() {
    return window.opener !== null && window.opener !== window;
}

function renderPopupStatus(message) {
    document.title = message;
    document.body.innerHTML = `
    <div style="min-height:100vh;display:flex;align-items:center;justify-content:center;background:#0b1020;color:#f3f6fb;font:16px/1.5 system-ui,sans-serif;padding:24px;text-align:center;">
      <div>
        <div style="font-size:18px;font-weight:600;margin-bottom:8px;">${escHtml(message)}</div>
        <div style="color:#9fb0c6;">This window should close automatically.</div>
      </div>
    </div>
  `;
}

function closePopupWindow(message) {
    renderPopupStatus(message);
    setTimeout(() => window.close(), 150);
}

export function getRuntimeConfig() {
    const clientId = String(config.clientId || '').trim();
    if (!clientId) throw new Error('Missing clientId in config.js');

    const authorityTenant = String(config.authorityTenant || DEFAULT_AUTHORITY_TENANT).trim() || DEFAULT_AUTHORITY_TENANT;
    const scopes = Array.isArray(config.scopes) && config.scopes.length
        ? config.scopes.map(scope => String(scope).trim()).filter(Boolean)
        : ['Team.ReadBasic.All', 'Schedule.ReadWrite.All'];

    return {
        clientId,
        authority: `https://login.microsoftonline.com/${authorityTenant}`,
        scopes,
        redirectUri: window.location.origin + window.location.pathname,
        appPagePath: config.appPagePath || './index.html',
        loginPagePath: config.loginPagePath || './login.html',
    };
}

function getConfigKey(c) {
    return JSON.stringify({ clientId: c.clientId, authority: c.authority, redirectUri: c.redirectUri, scopes: c.scopes });
}

function clearMsalState() {
    currentAccount = null;
    if (msalInstance) {
        try { msalInstance.setActiveAccount(null); } catch { /* ignore */ }
    }
    for (const storage of [sessionStorage, localStorage]) {
        for (let i = storage.length - 1; i >= 0; i -= 1) {
            const key = storage.key(i);
            if (key && key.startsWith('msal.')) storage.removeItem(key);
        }
    }
}

function redirectTo(pathOption) {
    const c = getRuntimeConfig();
    const target = new URL(c[pathOption], window.location.href);
    if (target.pathname !== window.location.pathname) {
        window.location.replace(target.href);
    }
}

const redirectToApp = () => redirectTo('appPagePath');
const redirectToLogin = () => redirectTo('loginPagePath');

function showSigninScreen() {
    const signInScreen = $('signin-screen');
    if (signInScreen) signInScreen.style.display = 'flex';
}

export async function initializeSession() {
    if (!initializeSessionPromise) {
        initializeSessionPromise = initializeSessionInternal()
            .finally(() => { initializeSessionPromise = null; });
    }
    return initializeSessionPromise;
}

async function initializeSessionInternal() {
    const c = getRuntimeConfig();
    const configKey = getConfigKey(c);
    if (!msalInstance || initializedConfigKey !== configKey) {
        msalInstance = new msal.PublicClientApplication({
            auth: { clientId: c.clientId, authority: c.authority, redirectUri: c.redirectUri },
            cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
        });
        initializedConfigKey = configKey;
        await msalInstance.initialize();
    }

    let redirectResponse = null;
    try {
        redirectResponse = await msalInstance.handleRedirectPromise();
        if (redirectResponse?.account) {
            msalInstance.setActiveAccount(redirectResponse.account);
            currentAccount = redirectResponse.account;
        }
    } catch (error) {
        if (isPopupContext()) renderPopupStatus('Sign-in hit an error. You can close this window.');
        else showError('Auth redirect error: ' + error.message);
    }

    if (isPopupContext()) {
        const popupAccount = redirectResponse?.account || msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
        if (popupAccount) closePopupWindow('Sign-in complete.');
        return { ready: false, authenticated: Boolean(popupAccount), account: popupAccount };
    }

    const activeAccount = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
    if (activeAccount) {
        msalInstance.setActiveAccount(activeAccount);
        currentAccount = activeAccount;
    } else {
        currentAccount = null;
    }

    return { ready: true, authenticated: Boolean(currentAccount), account: currentAccount };
}

export async function ensureSession({ redirectIfMissing = false } = {}) {
    const session = await initializeSession();
    if (!session.ready) return session;
    if (!session.authenticated && redirectIfMissing) redirectToLogin();
    return session;
}

function getScopes() {
    return getRuntimeConfig().scopes;
}

export async function signIn({ navigateToApp = document.body?.dataset.page === 'login' } = {}) {
    await initializeSession();
    if (currentAccount) {
        if (navigateToApp) redirectToApp();
        return currentAccount;
    }

    const loginRequest = { scopes: getScopes() };
    try {
        const response = await msalInstance.loginPopup(loginRequest);
        currentAccount = response.account;
        msalInstance.setActiveAccount(response.account);
        if (navigateToApp) redirectToApp();
        return currentAccount;
    } catch (popupError) {
        if (popupError.errorCode === 'popup_window_error' || popupError.errorCode === 'empty_window_error') {
            await msalInstance.loginRedirect(loginRequest);
            return null;
        }
        if (popupError.message && popupError.message.includes('AADSTS700016')) {
            clearMsalState();
            showError('App not found in that directory (AADSTS700016). Check the hardcoded client ID or tenant in config.js.');
            throw popupError;
        }
        showError('Sign-in failed: ' + popupError.message);
        throw popupError;
    }
}

export async function signOut({ navigateToLogin = true } = {}) {
    if (msalInstance) {
        await msalInstance.logoutPopup({ account: currentAccount }).catch(() => { });
    }
    clearMsalState();
    if (navigateToLogin) redirectToLogin();
}

export async function requireAppSession() {
    const session = await ensureSession({ redirectIfMissing: true });
    return session.ready && session.authenticated;
}

export function getCurrentAccount() {
    return currentAccount;
}

async function acquireToken() {
    const request = { scopes: getScopes(), account: currentAccount };
    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch {
        if (isPopupContext()) throw new Error('Cannot refresh token inside a popup context');
        if (!popupTokenPromise) {
            popupTokenPromise = msalInstance.acquireTokenPopup(request)
                .finally(() => { popupTokenPromise = null; });
        }
        try {
            const response = await popupTokenPromise;
            return response.accessToken;
        } catch (popupError) {
            currentAccount = null;
            toast('Session expired — please sign in again.', 'error');
            throw popupError;
        }
    }
}

async function graphRequest(baseUrl, url, options) {
    const token = await acquireToken();
    const fullUrl = url.startsWith('http') ? url : `${baseUrl}${url}`;
    const headers = {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...options.headers,
    };
    const response = await fetch(fullUrl, { ...options, headers });
    if (!response.ok) {
        let message = `Graph API error ${response.status}`;
        try { message = (await response.json())?.error?.message || message; } catch { /* ignore */ }
        const error = new Error(message);
        error.status = response.status;
        throw error;
    }
    return response.status === 204 ? null : response.json();
}

export const graphFetch = (url, options = {}) => graphRequest('https://graph.microsoft.com/v1.0', url, options);
export const graphFetchBeta = (url, options = {}) => graphRequest('https://graph.microsoft.com/beta', url, options);

async function bootLoginPage() {
    let session;
    try { session = await ensureSession(); }
    catch (error) { showError(error.message); showSigninScreen(); return; }

    if (!session.ready) return;
    if (session.authenticated) { redirectToApp(); return; }
    showSigninScreen();
}

const signInButton = $('btn-signin');
if (signInButton) signInButton.addEventListener('click', () => void signIn());

if (isPopupContext()) {
    renderPopupStatus('Completing sign-in...');
} else if (document.body.dataset.page === 'login') {
    void bootLoginPage();
}

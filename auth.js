'use strict';

const DEFAULT_AUTHORITY_TENANT = 'organizations';

let msalInstance = null;
let currentAccount = null;
let popupTokenPromise = null;
let initializedConfigKey = null;
let initializeSessionPromise = null;

function $(id) {
    return document.getElementById(id);
}

function toast(message, type = 'info') {
    const area = $('toast-area');
    if (!area) return;
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.textContent = message;
    area.appendChild(el);
    window.setTimeout(() => el.remove(), 4000);
}

function showError(message) {
    const banner = $('error-banner');
    if (!banner) return;
    banner.textContent = message;
    banner.style.display = 'block';
    window.setTimeout(() => {
        banner.style.display = 'none';
    }, 8000);
}

function escHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function isPopupContext() {
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
    window.setTimeout(() => {
        window.close();
    }, 150);
}

function getRuntimeConfig() {
    const raw = window.timecardsConfig || {};
    const clientId = String(raw.clientId || '').trim();
    if (!clientId) {
        throw new Error('Missing timecardsConfig.clientId in config.js');
    }

    const authorityTenant = String(raw.authorityTenant || DEFAULT_AUTHORITY_TENANT).trim() || DEFAULT_AUTHORITY_TENANT;
    const scopes = Array.isArray(raw.scopes) && raw.scopes.length
        ? raw.scopes.map(scope => String(scope).trim()).filter(Boolean)
        : ['Team.ReadBasic.All', 'Schedule.ReadWrite.All'];
    const redirectUri = window.location.origin + window.location.pathname;

    return {
        clientId,
        authority: `https://login.microsoftonline.com/${authorityTenant}`,
        scopes,
        redirectUri,
        appPagePath: raw.appPagePath || './index.html',
        loginPagePath: raw.loginPagePath || './login.html',
    };
}

function getConfigKey(config) {
    return JSON.stringify({
        clientId: config.clientId,
        authority: config.authority,
        redirectUri: config.redirectUri,
        scopes: config.scopes,
    });
}

function clearMsalState() {
    currentAccount = null;
    if (msalInstance) {
        try {
            msalInstance.setActiveAccount(null);
        } catch {
            // ignore MSAL cleanup errors
        }
    }

    for (const storage of [sessionStorage, localStorage]) {
        for (let index = storage.length - 1; index >= 0; index -= 1) {
            const key = storage.key(index);
            if (key && key.startsWith('msal.')) {
                storage.removeItem(key);
            }
        }
    }
}

function redirectToApp() {
    const config = getRuntimeConfig();
    const target = new URL(config.appPagePath, window.location.href);
    if (target.pathname === window.location.pathname) {
        return;
    }
    window.location.replace(target.href);
}

function redirectToLogin() {
    const config = getRuntimeConfig();
    const target = new URL(config.loginPagePath, window.location.href);
    if (target.pathname === window.location.pathname) {
        return;
    }
    window.location.replace(target.href);
}

function showSigninScreen() {
    const signInScreen = $('signin-screen');
    if (signInScreen) {
        signInScreen.style.display = 'flex';
    }
}

async function initializeSession() {
    if (!initializeSessionPromise) {
        initializeSessionPromise = initializeSessionInternal()
            .finally(() => {
                initializeSessionPromise = null;
            });
    }
    return initializeSessionPromise;
}

async function initializeSessionInternal() {
    const config = getRuntimeConfig();
    const configKey = getConfigKey(config);
    if (!msalInstance || initializedConfigKey !== configKey) {
        msalInstance = new msal.PublicClientApplication({
            auth: {
                clientId: config.clientId,
                authority: config.authority,
                redirectUri: config.redirectUri,
            },
            cache: {
                cacheLocation: 'sessionStorage',
                storeAuthStateInCookie: false,
            },
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
        if (isPopupContext()) {
            renderPopupStatus('Sign-in hit an error. You can close this window.');
        } else {
            showError('Auth redirect error: ' + error.message);
        }
    }

    if (isPopupContext()) {
        const popupAccount = redirectResponse?.account || msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
        if (popupAccount) {
            closePopupWindow('Sign-in complete.');
        }
        return { ready: false, authenticated: Boolean(popupAccount), account: popupAccount };
    }

    const activeAccount = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || null;
    if (activeAccount) {
        msalInstance.setActiveAccount(activeAccount);
        currentAccount = activeAccount;
    } else {
        currentAccount = null;
    }

    return {
        ready: true,
        authenticated: Boolean(currentAccount),
        account: currentAccount,
    };
}

async function ensureSession(options = {}) {
    const { redirectIfMissing = false } = options;
    const session = await initializeSession();
    if (!session.ready) {
        return session;
    }
    if (!session.authenticated && redirectIfMissing) {
        redirectToLogin();
    }
    return session;
}

function getScopes() {
    return getRuntimeConfig().scopes;
}

async function signIn(options = {}) {
    const { navigateToApp = document.body?.dataset.page === 'login' } = options;

    await initializeSession();
    if (currentAccount) {
        if (navigateToApp) {
            redirectToApp();
        }
        return currentAccount;
    }

    const loginRequest = { scopes: getScopes() };
    try {
        const response = await msalInstance.loginPopup(loginRequest);
        currentAccount = response.account;
        msalInstance.setActiveAccount(response.account);
        if (navigateToApp) {
            redirectToApp();
        }
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

async function signOut(options = {}) {
    const { navigateToLogin = true } = options;
    if (msalInstance) {
        await msalInstance.logoutPopup({ account: currentAccount }).catch(() => { });
    }
    clearMsalState();
    if (navigateToLogin) {
        redirectToLogin();
    }
}

async function requireAppSession() {
    const session = await ensureSession({ redirectIfMissing: true });
    return session.ready && session.authenticated;
}

function getCurrentAccount() {
    return currentAccount;
}

function getAccountStorageKey() {
    return currentAccount?.homeAccountId || currentAccount?.username || 'default';
}

function getAuthStatus() {
    return {
        ready: Boolean(msalInstance),
        authenticated: Boolean(currentAccount),
        account: currentAccount,
        isPopupContext: isPopupContext(),
    };
}

async function acquireToken() {
    const request = { scopes: getScopes(), account: currentAccount };
    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch {
        if (isPopupContext()) {
            throw new Error('Cannot refresh token inside a popup context');
        }

        if (!popupTokenPromise) {
            popupTokenPromise = msalInstance.acquireTokenPopup(request)
                .finally(() => {
                    popupTokenPromise = null;
                });
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

async function graphFetch(url, options = {}) {
    const token = await acquireToken();
    const base = url.startsWith('http') ? url : `https://graph.microsoft.com/v1.0${url}`;
    const headers = {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...options.headers,
    };
    const response = await fetch(base, { ...options, headers });
    if (!response.ok) {
        let message = `Graph API error ${response.status}`;
        try {
            const body = await response.json();
            message = body?.error?.message || message;
        } catch {
            // ignore response parse errors
        }
        const error = new Error(message);
        error.status = response.status;
        throw error;
    }
    if (response.status === 204) {
        return null;
    }
    return response.json();
}

async function graphFetchBeta(url, options = {}) {
    const token = await acquireToken();
    const headers = {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...options.headers,
    };
    const response = await fetch(`https://graph.microsoft.com/beta${url}`, { ...options, headers });
    if (!response.ok) {
        let message = `Graph API (beta) error ${response.status}`;
        try {
            const body = await response.json();
            message = body?.error?.message || message;
        } catch {
            // ignore response parse errors
        }
        const error = new Error(message);
        error.status = response.status;
        throw error;
    }
    if (response.status === 204) {
        return null;
    }
    return response.json();
}

async function bootLoginPage() {
    let session;
    try {
        session = await ensureSession();
    } catch (error) {
        showError(error.message);
        showSigninScreen();
        return;
    }

    if (!session.ready) {
        return;
    }

    if (session.authenticated) {
        redirectToApp();
        return;
    }

    showSigninScreen();
}

async function bootAppPage() {
    try {
        await ensureSession({ redirectIfMissing: true });
    } catch (error) {
        showError(error.message);
    }
}

const signInButton = $('btn-signin');
if (signInButton) {
    signInButton.addEventListener('click', () => {
        void signIn();
    });
}

window.timecardsAuth = {
    ensureSession,
    requireAppSession,
    signIn,
    signOut,
    graphFetch,
    graphFetchBeta,
    getCurrentAccount,
    getAccountStorageKey,
    getAuthStatus,
    getRuntimeConfig,
    isPopupContext,
};

(async function bootAuthPage() {
    if (isPopupContext()) {
        renderPopupStatus('Completing sign-in...');
        return;
    }

    if (document.body.dataset.page === 'login') {
        await bootLoginPage();
        return;
    }

    if (document.body.dataset.page === 'app') {
        await bootAppPage();
    }
})();

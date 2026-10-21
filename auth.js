// MSAL sign-in and Graph requests.
//
// Tokens refresh silently while the refresh token lasts. On an unmanaged PC that lease is short
// (8 of 24 hours), so a page left open over sleep comes back unable to refresh. Then nothing fails:
// every Graph request waits, a "Session expired" bar offers Sign in, and its click opens the
// popup (browsers block popups that no click opened). Once signed in, the waiting requests carry
// on, so the click that found the session expired still does what it was for.
import { config } from './config.js';
import { $, escHtml, showError } from './ui.js';
import { MOCK_ACCOUNT, mockEnabled, mockGraphFetch } from './mockgraph.js';

// Mock mode only: set to make every token request need interaction, as an expired session does.
export const MOCK_EXPIRED_KEY = 'mock_tc_session_expired';

const AUTHORITY = `https://login.microsoftonline.com/${config.authorityTenant}`;
// Registered redirect URIs (AADSTS50011 for any other). Popups and the hidden renewal iframe come
// back to an empty page whose opener or parent reads the response; a full-page redirect (when a
// popup is blocked) comes back to the login page, which completes it.
const AUTH_DONE_URI = new URL(config.authDonePagePath, window.location.href).href;
const LOGIN_URI = new URL(config.loginPagePath, window.location.href).href;

let msalInstance = null;
let currentAccount = null;
let initializePromise = null;
let reauth = null; // { promise, resolve } while a sign-in is awaited

export function getRuntimeConfig() {
    return { ...config, authority: AUTHORITY, redirectUri: AUTH_DONE_URI };
}

export function getCurrentAccount() {
    return currentAccount;
}

function setCurrentAccount(account) {
    currentAccount = account || null;
    if (msalInstance) msalInstance.setActiveAccount(currentAccount);
}

// Errors that only an interactive sign-in can fix: expired or revoked refresh tokens, missing
// consent, and the hidden-iframe renewal timing out (third-party cookies blocked).
function needsInteraction(error) {
    if (error instanceof msal.InteractionRequiredAuthError) return true;
    return /interaction_required|login_required|consent_required|no_tokens_found|no_account_error|token_refresh_required|monitor_window_timeout|block_iframe_reload|invalid_grant/i
        .test(`${error?.errorCode || ''} ${error?.message || ''}`);
}

const popupBlocked = error => /popup_window_error|empty_window_error/.test(error?.errorCode || '');
const popupCancelled = error => /user_cancelled/.test(error?.errorCode || '');

async function initializeSessionInternal() {
    if (mockEnabled) {
        currentAccount = MOCK_ACCOUNT;
        return { ready: true, authenticated: true, account: currentAccount };
    }
    if (!msalInstance) {
        msalInstance = new msal.PublicClientApplication({
            auth: { clientId: config.clientId, authority: AUTHORITY, redirectUri: AUTH_DONE_URI },
            cache: { cacheLocation: 'sessionStorage' },
        });
        await msalInstance.initialize();
    }
    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response?.account) setCurrentAccount(response.account);
    } catch (error) {
        showError('Sign-in redirect failed: ' + error.message);
    }
    setCurrentAccount(msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0]);
    return { ready: true, authenticated: Boolean(currentAccount), account: currentAccount };
}

export function initializeSession() {
    initializePromise ||= initializeSessionInternal().catch(error => {
        initializePromise = null;
        throw error;
    });
    return initializePromise;
}

export async function ensureSession({ redirectIfMissing = false } = {}) {
    const session = await initializeSession();
    if (!session.authenticated && redirectIfMissing) redirectTo(config.loginPagePath);
    return session;
}

export async function requireAppSession() {
    return (await ensureSession({ redirectIfMissing: true })).authenticated;
}

function redirectTo(path) {
    const target = new URL(path, window.location.href);
    if (target.pathname !== window.location.pathname) window.location.replace(target.href);
}

// The login page's button. A click, so the popup is allowed; a blocked popup falls back to a redirect.
export async function signIn({ navigateToApp = document.body.dataset.page === 'login' } = {}) {
    await initializeSession();
    if (!currentAccount) {
        try {
            setCurrentAccount((await msalInstance.loginPopup({ scopes: config.scopes })).account);
        } catch (error) {
            if (popupBlocked(error)) return msalInstance.loginRedirect({ scopes: config.scopes, redirectUri: LOGIN_URI });
            if (!popupCancelled(error)) showError('Sign-in failed: ' + error.message);
            throw error;
        }
    }
    if (navigateToApp) redirectTo(config.appPagePath);
    return currentAccount;
}

// Local only: forgets this app's tokens and account, and leaves the Microsoft session alone.
export async function signOut({ navigateToLogin = true } = {}) {
    if (msalInstance) await msalInstance.clearCache();
    setCurrentAccount(null);
    if (navigateToLogin) redirectTo(config.loginPagePath);
}

// ── Expired session ─────────────────────────────────────────────────────────

// Resolves once the user has signed in again through the bar.
function waitForSignIn() {
    if (!reauth) {
        let resolve;
        reauth = { promise: new Promise(r => { resolve = r; }), resolve };
        renderSessionBar('Your session expired, so what is shown may be out of date. Sign in to refresh it; what you were doing will finish.');
        document.body.classList.add('session-expired');
    }
    return reauth.promise;
}

function renderSessionBar(message) {
    let bar = $('session-bar');
    if (!bar) {
        bar = document.createElement('div');
        bar.id = 'session-bar';
        bar.setAttribute('role', 'alert');
        document.body.prepend(bar);
    }
    bar.innerHTML = `<span>${escHtml(message)}</span><button type="button" class="btn btn-warn">Sign in</button>`;
    // Called straight from the click with no await before it, so the popup counts as user-opened.
    bar.querySelector('button').addEventListener('click', () => void signInAgain());
}

async function signInAgain() {
    try {
        if (mockEnabled) {
            sessionStorage.removeItem(MOCK_EXPIRED_KEY);
        } else {
            const request = { scopes: config.scopes, account: currentAccount || undefined, loginHint: currentAccount?.username };
            setCurrentAccount((await msalInstance.acquireTokenPopup(request)).account);
        }
    } catch (error) {
        if (popupBlocked(error)) {
            await msalInstance.acquireTokenRedirect({ scopes: config.scopes, loginHint: currentAccount?.username, redirectUri: LOGIN_URI });
            return;
        }
        if (!popupCancelled(error)) renderSessionBar(`Sign-in failed: ${error.message}. Try again.`);
        return;
    }
    $('session-bar')?.remove();
    document.body.classList.remove('session-expired');
    const { resolve } = reauth;
    reauth = null;
    resolve();
}

async function acquireToken({ forceRefresh = false } = {}) {
    await initializeSession();
    // Unbounded on purpose: each pass that finds no token waits for the user to sign in again,
    // for as long as that takes. It spins only as fast as they click Sign in.
    for (;;) {
        if (mockEnabled) {
            if (!sessionStorage.getItem(MOCK_EXPIRED_KEY)) return 'mock-token';
        } else if (currentAccount) {
            try {
                // Locally, `sessionStorage.dev_no_iframe_renewal` stops renewal at the refresh token, as
                // on the owner's unmanaged PC, where the hidden-iframe renewal cannot get a token either.
                const noIframe = location.hostname === 'localhost' && sessionStorage.getItem('dev_no_iframe_renewal');
                const response = await msalInstance.acquireTokenSilent({
                    scopes: config.scopes, account: currentAccount, forceRefresh,
                    ...(noIframe ? { cacheLookupPolicy: msal.CacheLookupPolicy.RefreshToken } : {}),
                });
                setCurrentAccount(response.account);
                return response.accessToken;
            } catch (error) {
                if (!needsInteraction(error)) throw error;
            }
        }
        await waitForSignIn();
        forceRefresh = false;
    }
}

// On returning to the tab (after sleep, say), find out now whether the session lasted, so the
// bar is already up before the user clicks anything.
function checkSessionOnReturn() {
    if (document.hidden || !currentAccount || reauth) return;
    acquireToken().catch(error => showError('Session check failed: ' + error.message));
}

// ── Graph ───────────────────────────────────────────────────────────────────

async function graphRequest(baseUrl, url, options, retryUnauthorized = true) {
    const token = await acquireToken({ forceRefresh: !retryUnauthorized });
    const fullUrl = url.startsWith('http') ? url : `${baseUrl}${url}`;
    if (mockEnabled) return mockGraphFetch(fullUrl, options);
    const response = await fetch(fullUrl, {
        ...options,
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...options.headers },
    });
    if (response.ok) return response.status === 204 ? null : response.json();

    const payload = await response.json().catch(() => null);
    // A 401 can mean the token was revoked early: retry once with a freshly fetched one.
    if (response.status === 401 && retryUnauthorized) return graphRequest(baseUrl, url, options, false);
    const error = new Error(payload?.error?.message || `Graph API error ${response.status}`);
    error.status = response.status;
    error.code = payload?.error?.code || '';
    throw error;
}

export const graphFetch = (url, options = {}) => graphRequest('https://graph.microsoft.com/v1.0', url, options);
export const graphFetchBeta = (url, options = {}) => graphRequest('https://graph.microsoft.com/beta', url, options);

// ── Page boot ───────────────────────────────────────────────────────────────

async function bootLoginPage() {
    try {
        if ((await ensureSession()).authenticated) { redirectTo(config.appPagePath); return; }
    } catch (error) {
        showError(error.message);
    }
    $('signin-screen').style.display = 'flex';
    $('btn-signin').addEventListener('click', () => void signIn().catch(() => { }));
}

document.addEventListener('visibilitychange', checkSessionOnReturn);
if (document.body.dataset.page === 'login') void bootLoginPage();

import * as auth from './auth.js';
import {
    fetchJoinedTeams,
    fetchWeekTimeCards,
    buildTimeCardsPageUrl,
    getTeamCache,
    getWeekRange,
    STORAGE_KEYS,
} from './app.js';
import { formatDateInputValue } from './core.js';
import { $, reportToDevServer } from './ui.js';
import { mockCards, mockEnabled, mockGraphFetch, MOCK_TEAM, seedMockDb } from './mockgraph.js';

const SESSION_KEY_SELECTED_TEAM_ID = STORAGE_KEYS.selectedTeamId;
const DAY_MS = 24 * 60 * 60 * 1000;

function parseParams() {
    const params = new URLSearchParams(window.location.search);
    return {
        test: (params.get('test') || 'all').trim(),
        teamId: (params.get('teamId') || '').trim(),
        week: (params.get('week') || '').trim(),
    };
}

function formatJson(value) {
    return typeof value === 'string' ? value : JSON.stringify(value, null, 2);
}

function formatAccount(account) {
    if (!account) return 'not signed in';
    return `${account.name || account.username || 'signed in'} (${account.username || account.homeAccountId || 'no identifier'})`;
}

function setAuthStatus(session) {
    const authStatus = $('test-auth-status');
    const signInButton = $('btn-test-signin');
    const signOutButton = $('btn-test-signout');
    if (!authStatus || !signInButton || !signOutButton) return;
    authStatus.textContent = session.authenticated
        ? `Authenticated: ${formatAccount(session.account)}`
        : 'Authenticated: no';
    signInButton.style.display = session.authenticated ? 'none' : 'inline-flex';
    signOutButton.style.display = session.authenticated ? 'inline-flex' : 'none';
}

function renderMeta(params) {
    const meta = $('test-meta');
    if (!meta) return;
    const queryBits = [`test=${params.test}`];
    if (params.teamId) queryBits.push(`teamId=${params.teamId}`);
    if (params.week) queryBits.push(`week=${params.week}`);
    const names = tests.map(t => t.name).join(', ');
    meta.textContent = `Run one test with ?test=<name>. Available: ${names}. Current query: ${queryBits.join('&')}`;
}

function renderSummary(message) {
    const summary = $('test-summary');
    if (summary) summary.textContent = message;
}

function createResultRow(name) {
    const results = $('test-results');
    const row = document.createElement('section');
    row.className = 'test-result test-result-running';
    row.innerHTML = `
    <div class="test-result-name">${name}</div>
    <div class="test-result-body">running…</div>
  `;
    results.appendChild(row);
    return row;
}

function updateResultRow(row, status, details) {
    row.className = `test-result test-result-${status}`;
    const body = row.querySelector('.test-result-body');
    if (body) body.textContent = formatJson(details);
}

// ─────────────────────────────────────────────
// Probes (Microsoft Graph contract verification)
// These are verification-only; production code never runs them.
// ─────────────────────────────────────────────
function buildTimeCardFilterUrl(teamId, filterExpr, pageSize = 1) {
    return `/teams/${teamId}/schedule/timeCards?$top=${pageSize}&$filter=${encodeURIComponent(filterExpr)}`;
}

function isRejectedFilterError(message) {
    return /filter|unsupported|not allowed|could not find a property|invalid|orderby|order by|sort/i.test(message || '');
}

function escapeODataLiteral(value) {
    return String(value || '').replace(/'/g, "''");
}

async function probeFilterExpression(teamId, expression) {
    try {
        await auth.graphFetch(buildTimeCardFilterUrl(teamId, expression));
        return { supported: true, message: 'Graph accepted the filter.' };
    } catch (error) {
        const message = error?.message || 'Unknown Graph error';
        if (isRejectedFilterError(message)) return { supported: false, message };
        return { supported: null, message };
    }
}

async function probeFilterAlternatives(teamId, label, expressions) {
    let lastMessage = '';
    for (const expression of expressions) {
        const r = await probeFilterExpression(teamId, expression);
        if (r.supported === true) {
            return { label, supported: true, acceptedExpression: expression, triedExpressions: expressions, message: r.message };
        }
        if (r.supported === null) {
            return { label, supported: null, acceptedExpression: null, triedExpressions: expressions, message: r.message };
        }
        lastMessage = r.message;
    }
    return { label, supported: false, acceptedExpression: null, triedExpressions: expressions, message: lastMessage || 'Rejected.' };
}

async function probeDocumentedFilterMatrix(teamId) {
    const account = auth.getCurrentAccount();
    const userId = account?.localAccountId || account?.homeAccountId || '00000000-0000-0000-0000-000000000000';
    const now = Date.now();
    const lowerBoundIso = new Date(now - 30 * DAY_MS).toISOString();
    const upperBoundIso = new Date(now + DAY_MS).toISOString();

    const definitions = [
        { key: 'clockInGe', label: 'clockInEvent/dateTime ge', expressions: [`clockInEvent/dateTime ge ${lowerBoundIso}`, `clockInEvent/datetime ge ${lowerBoundIso}`] },
        { key: 'clockInLe', label: 'clockInEvent/dateTime le', expressions: [`clockInEvent/dateTime le ${upperBoundIso}`, `clockInEvent/datetime le ${upperBoundIso}`] },
        { key: 'clockOutGe', label: 'clockOutEvent/dateTime ge', expressions: [`clockOutEvent/dateTime ge ${lowerBoundIso}`, `clockOutEvent/datetime ge ${lowerBoundIso}`] },
        { key: 'clockOutLe', label: 'clockOutEvent/dateTime le', expressions: [`clockOutEvent/dateTime le ${upperBoundIso}`, `clockOutEvent/datetime le ${upperBoundIso}`] },
        { key: 'stateEq', label: 'state eq', expressions: ["state eq 'clockedIn'"] },
        { key: 'userIdEq', label: 'userId eq', expressions: [`userId eq '${escapeODataLiteral(userId)}'`] },
    ];

    const entries = await Promise.all(definitions.map(async d => ({
        key: d.key,
        result: await probeFilterAlternatives(teamId, d.label, d.expressions),
    })));

    const matrix = { checkedAt: new Date().toISOString() };
    entries.forEach(e => { matrix[e.key] = e.result; });
    return matrix;
}

async function probeLastModifiedFilter(teamId) {
    const r = await probeFilterExpression(teamId, 'lastModifiedDateTime ge 1970-01-01T00:00:00Z');
    return r.supported;
}

async function probeLastModifiedOrderBy(teamId) {
    const url = `/teams/${teamId}/schedule/timeCards?$top=1&$orderby=${encodeURIComponent('lastModifiedDateTime desc')}`;
    try {
        await auth.graphFetch(url);
        return true;
    } catch (error) {
        const message = error?.message || '';
        if (/lastmodifieddatetime|orderby|order by|sort|not allowed|unsupported/i.test(message)) return false;
        return null;
    }
}

async function probeCombinedWeekQuery(teamId, weekStartInput) {
    const { weekStart, weekEndInclusive } = getWeekRange(weekStartInput);
    const filterExpr = `((clockInEvent/dateTime le ${weekEndInclusive.toISOString()} and clockOutEvent/dateTime ge ${weekStart.toISOString()}) or state eq 'clockedIn' or state eq 'onBreak')`;
    const r = await probeFilterExpression(teamId, filterExpr);
    return r.supported;
}

// ─────────────────────────────────────────────
// Test helpers
// ─────────────────────────────────────────────
function summarizeCard(card) {
    if (!card) return null;
    const breaks = card.breaks || [];
    return {
        id: card.id || null,
        state: card.state || null,
        clockIn: card.clockIn?.dateTime || null,
        clockOut: card.clockOut?.dateTime || null,
        breakCount: breaks.length,
        openBreakCount: breaks.filter(b => b.start && !b.end).length,
        lastModifiedDateTime: card.lastModifiedDateTime || null,
    };
}

function assertProbeRejected(supported, label) {
    if (supported !== false) throw new Error(`Expected ${label} to be rejected, got ${supported}`);
    return supported;
}

function assertDocumentedFilterSupport(matrix) {
    const required = ['clockInGe', 'clockInLe', 'clockOutGe', 'clockOutLe', 'stateEq', 'userIdEq'];
    const unsupported = required.filter(key => matrix?.[key]?.supported !== true).map(key => matrix?.[key]?.label || key);
    if (unsupported.length) {
        throw new Error(`Expected documented timecard filters to be supported: ${unsupported.join(', ')}`);
    }
    return matrix;
}

async function requireTeam(context) {
    if (context.team) return context.team;
    const teams = await fetchJoinedTeams();
    context.teams = teams;
    const savedId = context.params.teamId || localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID) || '';
    context.team = teams.find(t => t.id === savedId) || teams[0] || null;
    if (!context.team) throw new Error('No joined teams were available for the current account');
    return context.team;
}

// ─────────────────────────────────────────────
// Flows: the real app in an iframe, against the mock backend
// ─────────────────────────────────────────────
function check(condition, message) {
    if (!condition) throw new Error(message);
}

async function waitFor(describe, predicate, timeoutMs = 4000) {
    const deadline = performance.now() + timeoutMs;
    for (;;) {
        const value = predicate();
        if (value) return value;
        if (performance.now() > deadline) throw new Error(`Timed out waiting for ${describe}`);
        await new Promise(resolve => setTimeout(resolve, 20));
    }
}

// A fresh app on the seeded mock data, showing this week, already past its first render.
async function openApp({ beforeLoad = () => { } } = {}) {
    seedMockDb();
    sessionStorage.removeItem(auth.MOCK_EXPIRED_KEY);
    sessionStorage.removeItem(STORAGE_KEYS.selectedWeek);
    sessionStorage.removeItem(STORAGE_KEYS.lastActivity);
    sessionStorage.removeItem(STORAGE_KEYS.joinedTeams);
    localStorage.setItem(SESSION_KEY_SELECTED_TEAM_ID, MOCK_TEAM.id);
    await beforeLoad();
    const frame = document.createElement('iframe');
    frame.className = 'test-app-frame';
    frame.src = 'index.html';
    document.body.appendChild(frame);
    await new Promise(resolve => frame.addEventListener('load', resolve, { once: true }));
    const doc = frame.contentDocument;
    const app = {
        doc,
        $: selector => doc.querySelector(selector),
        $$: selector => [...doc.querySelectorAll(selector)],
        clockLabel: () => doc.getElementById('btn-clock').textContent,
        weekLabel: () => doc.getElementById('week-range-label').textContent,
        // Steps back a week and waits until that week's cards are on screen (the seed fills it).
        async previousWeek() {
            const label = this.weekLabel();
            this.click('#btn-week-prev');
            await waitFor('the previous week to load', () => this.weekLabel() !== label
                && doc.getElementById('timecards-loading').style.display === 'none' && this.$$('.tc-row').length);
        },
        click(selector) {
            const element = waitClickable(doc, selector);
            element.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
            element.click();
        },
        close: () => frame.remove(),
    };
    await waitFor('the first render', () => doc.getElementById('weeks-container')?.style.display === 'block');
    return app;
}

function waitClickable(doc, selector) {
    const element = doc.querySelector(selector);
    if (!element) throw new Error(`No element matches ${selector}`);
    if (element.disabled) throw new Error(`${selector} is disabled`);
    return element;
}

async function withApp(run, options) {
    const app = await openApp(options);
    try { return await run(app); }
    finally { app.close(); }
}

const openMockCard = () => mockCards().find(card => card.state !== 'clockedOut');

const flowTests = [
    {
        name: 'flow-clock-in-break-clock-out',
        mode: 'mock',
        run: () => withApp(async app => {
            check(app.clockLabel() === 'Clock In', `starts clocked out, got ${app.clockLabel()}`);
            app.click('#btn-clock');
            await waitFor('Clock Out', () => app.clockLabel() === 'Clock Out');
            check(openMockCard()?.state === 'clockedIn', 'the backend has an open card');

            app.click('.break-row-control button');
            await waitFor('On Break', () => app.clockLabel() === 'On Break');
            check(app.$('#btn-clock').disabled, 'cannot clock out on break');
            check(openMockCard()?.state === 'onBreak', 'the backend is on break');

            (await waitFor('the End Break button', () => app.$$('.breaks-list button').find(b => b.textContent === 'End Break'))).click();
            await waitFor('Clock Out again', () => app.clockLabel() === 'Clock Out');

            app.click('#btn-clock');
            await waitFor('Clock In again', () => app.clockLabel() === 'Clock In');
            const card = mockCards().at(-1);
            check(card.state === 'clockedOut' && card.breaks.length === 1 && card.breaks[0].end, `backend card ends clocked out with one closed break: ${JSON.stringify(card)}`);
            return { cards: mockCards().length };
        }),
    },
    {
        name: 'flow-past-week-still-shows-active-card',
        mode: 'mock',
        // The app starts on last week (a reload while there), so it never saw this week's open card.
        run: () => withApp(async app => {
            await waitFor('last week to load', () => app.$$('.tc-row').length);
            await waitFor('Clock Out on a past week', () => app.clockLabel() === 'Clock Out');
            return { week: app.weekLabel() };
        }, {
            beforeLoad: async () => {
                const lastWeek = new Date();
                lastWeek.setDate(lastWeek.getDate() - 7);
                sessionStorage.setItem(STORAGE_KEYS.selectedWeek, formatDateInputValue(lastWeek));
                sessionStorage.setItem(STORAGE_KEYS.lastActivity, String(Date.now()));
                await mockGraphFetch(`/teams/${MOCK_TEAM.id}/schedule/timeCards/clockIn`, { method: 'POST', body: '{}' });
            },
        }),
    },
    {
        name: 'flow-clock-in-from-past-week',
        mode: 'mock',
        run: () => withApp(async app => {
            await app.previousWeek();
            const rowsBefore = app.$$('.tc-row').length;
            app.click('#btn-clock');
            await waitFor('Clock Out', () => app.clockLabel() === 'Clock Out');
            check(app.$$('.tc-row').length === rowsBefore, 'the past week still shows its own cards');
            check(!app.$('#btn-clock').disabled, 'the clock button is usable');
            return { rows: rowsBefore };
        }),
    },
    {
        name: 'flow-edit-rejects-break-before-clock-in',
        mode: 'mock',
        run: () => withApp(async app => {
            const before = JSON.stringify(mockCards());
            app.click('.tc-row [data-action="edit"]');
            await waitFor('the edit dialog', () => !app.$('#edit-modal').classList.contains('hidden'));
            const clockIn = app.$('#edit-clock-in').value;
            const breakStart = app.$('.break-edit-row .break-start');
            breakStart.value = `${clockIn.slice(0, 11)}00:05`;
            app.click('#btn-edit-save');
            await waitFor('the error toast', () => app.$$('.toast.error').some(t => /Break start/.test(t.textContent)));
            check(JSON.stringify(mockCards()) === before, 'nothing was saved');
            return { clockIn, breakStart: breakStart.value };
        }),
    },
    {
        name: 'flow-edit-saves-notes',
        mode: 'mock',
        run: () => withApp(async app => {
            const cardId = app.$('.tc-row').dataset.cardId;
            app.click('.tc-row [data-action="edit"]');
            await waitFor('the edit dialog', () => !app.$('#edit-modal').classList.contains('hidden'));
            app.$('#edit-notes').value = 'flow note';
            app.click('#btn-edit-save');
            await waitFor('the dialog to close', () => app.$('#edit-modal').classList.contains('hidden'));
            const card = mockCards().find(c => c.id === cardId);
            check(card.notes?.content === 'flow note', `notes saved: ${JSON.stringify(card.notes)}`);
            check(card.breaks.length === 1 && card.breaks[0].breakId, 'the break kept its id');
            return { cardId };
        }),
    },
    {
        name: 'flow-delete-card',
        mode: 'mock',
        run: () => withApp(async app => {
            const cardId = app.$('.tc-row').dataset.cardId;
            app.click('.tc-row [data-action="delete"]');
            await waitFor('the confirm dialog', () => !app.$('#confirm-modal').classList.contains('hidden'));
            app.click('#btn-confirm-ok');
            await waitFor('the row to go', () => !app.$(`.tc-row[data-card-id="${cardId}"]`));
            check(!mockCards().some(c => c.id === cardId), 'the backend card is gone');
            return { cardId };
        }),
    },
    {
        name: 'flow-expired-session-finishes-the-click',
        mode: 'mock',
        run: () => withApp(async app => {
            sessionStorage.setItem(auth.MOCK_EXPIRED_KEY, '1');
            const cardsBefore = mockCards().length;
            app.click('#btn-clock');
            await waitFor('the session bar', () => app.$('#session-bar button'));
            check(mockCards().length === cardsBefore, 'nothing happens until signed in');
            check(app.$('#btn-clock').disabled, 'the clicked button waits');
            app.click('#session-bar button');
            await waitFor('Clock Out after signing in', () => app.clockLabel() === 'Clock Out');
            check(!app.$('#session-bar'), 'the bar is gone');
            check(openMockCard(), 'the clock-in reached the backend');
            return { cards: mockCards().length };
        }),
    },
    {
        name: 'flow-expired-session-shows-on-return',
        mode: 'mock',
        run: () => withApp(async app => {
            sessionStorage.setItem(auth.MOCK_EXPIRED_KEY, '1');
            app.doc.dispatchEvent(new Event('visibilitychange'));
            await waitFor('the session bar without any click', () => app.$('#session-bar'));
            check(app.doc.body.classList.contains('session-expired'), 'the stale data is marked');
            app.click('#session-bar button');
            await waitFor('the bar to go', () => !app.$('#session-bar'));
            check(!app.doc.body.classList.contains('session-expired'), 'the mark goes with the bar');
            return {};
        }),
    },
    {
        name: 'flow-idle-returns-to-current-week',
        mode: 'mock',
        run: () => withApp(async app => {
            const thisWeek = app.weekLabel();
            await app.previousWeek();
            sessionStorage.setItem(STORAGE_KEYS.lastActivity, String(Date.now() - 2 * 3600 * 1000));
            app.doc.dispatchEvent(new Event('visibilitychange'));
            await waitFor('this week again', () => app.weekLabel() === thisWeek);
            return { week: thisWeek };
        }),
    },
];

// ─────────────────────────────────────────────
// Test definitions. mode: 'live' runs against real Graph only, 'mock' with ?mock only, 'both' always.
// ─────────────────────────────────────────────
const tests = [
    {
        name: 'auth-status',
        requiresAuth: false,
        run: async ctx => ({
            authenticated: ctx.session.authenticated,
            account: ctx.session.account ? {
                name: ctx.session.account.name || '',
                username: ctx.session.account.username || '',
            } : null,
            redirectUri: auth.getRuntimeConfig().redirectUri,
        }),
    },
    ...flowTests,
    {
        name: 'confirm-supported-timecard-filters',
        requiresAuth: true,
        run: async ctx => {
            const team = await requireTeam(ctx);
            return {
                teamId: team.id,
                support: assertDocumentedFilterSupport(await probeDocumentedFilterMatrix(team.id)),
            };
        },
    },
    {
        name: 'confirm-no-combined-week-query',
        requiresAuth: true,
        run: async ctx => {
            const team = await requireTeam(ctx);
            return {
                teamId: team.id,
                supported: assertProbeRejected(
                    await probeCombinedWeekQuery(team.id, ctx.params.week),
                    'combined weekly timecard query support',
                ),
            };
        },
    },
    {
        name: 'joined-teams',
        requiresAuth: true,
        run: async ctx => {
            const teams = await fetchJoinedTeams();
            ctx.teams = teams;
            return {
                count: teams.length,
                firstTeam: teams[0] ? { id: teams[0].id, displayName: teams[0].displayName } : null,
            };
        },
    },
    {
        name: 'resolve-team',
        requiresAuth: true,
        run: async ctx => {
            const teams = await fetchJoinedTeams();
            ctx.teams = teams;
            const savedId = ctx.params.teamId || localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID) || '';
            ctx.team = teams.find(t => t.id === savedId) || teams[0] || null;
            return {
                requestedTeamId: ctx.params.teamId || null,
                resolvedTeam: ctx.team ? { id: ctx.team.id, displayName: ctx.team.displayName } : null,
                availableTeams: teams.length,
            };
        },
    },
    {
        name: 'timecards-top-1',
        requiresAuth: true,
        run: async ctx => {
            const team = await requireTeam(ctx);
            const result = await auth.graphFetch(buildTimeCardsPageUrl(team.id, { pageSize: 1 }));
            return {
                teamId: team.id,
                count: Array.isArray(result?.value) ? result.value.length : 0,
                hasNextLink: Boolean(result?.['@odata.nextLink']),
            };
        },
    },
    {
        name: 'selected-week-fetch-timing',
        requiresAuth: true,
        run: async ctx => {
            const team = await requireTeam(ctx);
            const startedAt = performance.now();
            const snapshot = await fetchWeekTimeCards(team.id, ctx.params.week);
            const durationMs = Math.round(performance.now() - startedAt);
            return {
                teamId: team.id,
                weekStart: formatDateInputValue(snapshot.weekStart),
                cardsCount: snapshot.cards.length,
                pageCount: snapshot.pageCount,
                durationMs,
                activeCard: summarizeCard(snapshot.activeCard),
                sample: snapshot.cards.slice(0, 10).map(summarizeCard),
            };
        },
    },
    {
        name: 'server-vs-cache',
        requiresAuth: true,
        run: async ctx => {
            const team = await requireTeam(ctx);
            const snapshot = await fetchWeekTimeCards(team.id, ctx.params.week);
            const cache = getTeamCache(team.id);
            return {
                teamId: team.id,
                selectedWeek: formatDateInputValue(snapshot.weekStart),
                serverCardsCount: snapshot.cards.length,
                cachedWeekKey: cache.currentWeekKey || null,
                cachedCardsCount: cache.currentWeekCards.length,
            };
        },
    },
    {
        name: 'confirm-no-lastmodified-filter',
        requiresAuth: true,
        run: async ctx => {
            const team = await requireTeam(ctx);
            return {
                teamId: team.id,
                supported: assertProbeRejected(
                    await probeLastModifiedFilter(team.id),
                    'lastModifiedDateTime filter support',
                ),
            };
        },
    },
    {
        name: 'confirm-no-lastmodified-orderby',
        requiresAuth: true,
        run: async ctx => {
            const team = await requireTeam(ctx);
            return {
                teamId: team.id,
                supported: assertProbeRejected(
                    await probeLastModifiedOrderBy(team.id),
                    'lastModifiedDateTime orderby support',
                ),
            };
        },
    },
];

function getRequestedTests(params) {
    const mode = test => test.mode || (test.requiresAuth ? 'live' : 'both');
    const runnable = tests.filter(test => mode(test) === 'both' || mode(test) === (mockEnabled ? 'mock' : 'live'));
    if (!params.test || params.test === 'all') return runnable;
    return runnable.filter(test => test.name === params.test);
}

async function runSuite() {
    const params = parseParams();
    renderMeta(params);
    $('test-results').innerHTML = '';

    let session;
    try {
        session = await auth.ensureSession({ redirectIfMissing: false });
    } catch (error) {
        renderSummary('Auth bootstrap failed before tests could run.');
        setAuthStatus({ authenticated: false, account: null });
        const row = createResultRow('auth-bootstrap');
        updateResultRow(row, 'fail', error.message || String(error));
        return;
    }

    setAuthStatus(session);

    const requested = getRequestedTests(params);
    if (!requested.length) {
        renderSummary(`Unknown test "${params.test}".`);
        const row = createResultRow('test-selection');
        updateResultRow(row, 'fail', `Unknown test name: ${params.test}`);
        return;
    }

    const context = { params, session, team: null, teams: null };
    const startedAt = performance.now();
    let passed = 0;
    const problems = [];

    for (const test of requested) {
        const row = createResultRow(test.name);
        if (test.requiresAuth && !session.authenticated) {
            problems.push(`SKIP ${test.name}: not signed in`);
            updateResultRow(row, 'skip', 'Skipped because the integrated browser is not signed in yet.');
            continue;
        }

        try {
            const result = await test.run(context);
            passed += 1;
            updateResultRow(row, 'pass', result);
        } catch (error) {
            const message = error?.message || String(error);
            problems.push(`FAIL ${test.name}: ${message}`);
            updateResultRow(row, 'fail', message);
            console.error(`[self-test] FAIL ${test.name}`, error);
        }
    }

    const failed = problems.filter(line => line.startsWith('FAIL')).length;
    const skipped = problems.length - failed;
    renderSummary(`Requested ${requested.length} test(s). Passed: ${passed}. Failed: ${failed}. Skipped: ${skipped}. Authenticated: ${session.authenticated ? 'yes' : 'no'}.`);
    const seconds = ((performance.now() - startedAt) / 1000).toFixed(1);
    reportToDevServer(mockEnabled ? 'mock-tests' : 'browser-tests', [
        `${failed ? 'FAIL' : 'ok'} ${mockEnabled ? 'mock' : 'live'} browser tests: ${passed}/${requested.length} passed, ${failed} failed, ${skipped} skipped, signed in: ${session.authenticated ? 'yes' : 'no'}, ${seconds}s`,
        ...problems,
    ].join('\n'));
}

function bindActions() {
    const signInButton = $('btn-test-signin');
    const signOutButton = $('btn-test-signout');

    if (signInButton) {
        signInButton.addEventListener('click', async () => {
            signInButton.disabled = true;
            try { await auth.signIn({ navigateToApp: false }); }
            finally { signInButton.disabled = false; await runSuite(); }
        });
    }

    if (signOutButton) {
        signOutButton.addEventListener('click', async () => {
            signOutButton.disabled = true;
            try { await auth.signOut({ navigateToLogin: false }); }
            finally { signOutButton.disabled = false; await runSuite(); }
        });
    }
}

bindActions();
void runSuite();

'use strict';

const auth = window.timecardsAuth;
const app = window.timecardsApp;

const SESSION_KEY_SELECTED_TEAM_ID = 'tc_selected_team_id';
const DAY_MS = 24 * 60 * 60 * 1000;

function $(id) {
    return document.getElementById(id);
}

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
    const { weekStart, weekEndInclusive } = app.getWeekRange(weekStartInput);
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
    const teams = await app.fetchJoinedTeams();
    context.teams = teams;
    const savedId = context.params.teamId || localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID) || '';
    context.team = teams.find(t => t.id === savedId) || teams[0] || null;
    if (!context.team) throw new Error('No joined teams were available for the current account');
    return context.team;
}

// ─────────────────────────────────────────────
// Test definitions
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
    {
        name: 'app-api',
        requiresAuth: false,
        run: async () => ({
            hasAuthApi: Boolean(auth),
            hasAppApi: Boolean(app),
            appMethods: Object.keys(app || {}),
        }),
    },
    {
        name: 'timeline-open-card-update',
        requiresAuth: false,
        run: async () => {
            const clockIn = new Date('2026-04-29T16:00:00.000Z');
            const attemptedClockOut = new Date('2026-04-29T18:15:00.000Z');
            const projection = app.buildProjectedTimelineUpdate({
                id: 'self-test-open-card',
                state: 'clockedIn',
                userId: 'self-test-user',
                createdDateTime: clockIn.toISOString(),
                lastModifiedDateTime: clockIn.toISOString(),
                clockIn: { dateTime: clockIn.toISOString(), atApprovedLocation: false },
                clockOut: null,
                breaks: [],
                notes: null,
            }, new Date('2026-04-29T15:30:00.000Z'), attemptedClockOut);

            return {
                effectiveEnd: projection.effectiveEnd,
                updatedState: projection.updatedCard.state,
                updatedClockOut: projection.updatedCard.clockOut,
                hasClockOutEvent: Boolean(projection.requestBody.clockOutEvent),
            };
        },
    },
    {
        name: 'cache-snapshot',
        requiresAuth: false,
        run: async ctx => {
            const teamId = ctx.params.teamId || localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID) || '';
            if (!teamId) {
                return { selectedTeamId: '', reason: 'No saved team is currently selected.' };
            }
            const cache = app.getTeamCache(teamId);
            const selectedWeek = app.getResolvedWeekStart();
            const selectedWeekKey = app.formatDateInputValue(selectedWeek);
            const hasCurrentWeek = cache.currentWeekKey === selectedWeekKey;
            const now = Date.now();
            return {
                selectedTeamId: teamId,
                selectedWeek: selectedWeekKey,
                cachedWeekKey: cache.currentWeekKey || null,
                currentWeek: hasCurrentWeek ? {
                    cardsCount: cache.currentWeekCards.length,
                    fetchedAt: cache.currentWeekFetchedAt || 0,
                    ageMs: cache.currentWeekFetchedAt ? (now - cache.currentWeekFetchedAt) : null,
                } : null,
                hasPendingRequest: Boolean(cache.pendingWeekPromise),
                activeCard: summarizeCard(cache.activeCard),
            };
        },
    },
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
            const teams = await app.fetchJoinedTeams();
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
            const teams = await app.fetchJoinedTeams();
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
            const result = await auth.graphFetch(app.buildTimeCardsPageUrl(team.id, { pageSize: 1 }));
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
            const snapshot = await app.fetchWeekTimeCards(team.id, ctx.params.week);
            const durationMs = Math.round(performance.now() - startedAt);
            return {
                teamId: team.id,
                weekStart: app.formatDateInputValue(snapshot.weekStart),
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
            const snapshot = await app.fetchWeekTimeCards(team.id, ctx.params.week);
            const cache = app.getTeamCache(team.id);
            return {
                teamId: team.id,
                selectedWeek: app.formatDateInputValue(snapshot.weekStart),
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
    if (!params.test || params.test === 'all') return tests;
    const found = tests.find(test => test.name === params.test);
    return found ? [found] : [];
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
    let passed = 0;
    let failed = 0;
    let skipped = 0;

    for (const test of requested) {
        const row = createResultRow(test.name);
        if (test.requiresAuth && !session.authenticated) {
            skipped += 1;
            updateResultRow(row, 'skip', 'Skipped because the integrated browser is not signed in yet.');
            continue;
        }

        try {
            const result = await test.run(context);
            passed += 1;
            updateResultRow(row, 'pass', result);
            console.log(`[self-test] PASS ${test.name}`, result);
        } catch (error) {
            failed += 1;
            updateResultRow(row, 'fail', error?.message || String(error));
            console.error(`[self-test] FAIL ${test.name}`, error);
        }
    }

    renderSummary(`Requested ${requested.length} test(s). Passed: ${passed}. Failed: ${failed}. Skipped: ${skipped}. Authenticated: ${session.authenticated ? 'yes' : 'no'}.`);
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

if (!auth.isPopupContext()) {
    bindActions();
    void runSuite();
}

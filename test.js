'use strict';

const auth = window.timecardsAuth;
const appApi = window.timecardsApp;

function $(id) {
    return document.getElementById(id);
}

function parseParams() {
    const params = new URLSearchParams(window.location.search);
    return {
        test: (params.get('test') || 'all').trim(),
        teamId: (params.get('teamId') || '').trim(),
        week: (params.get('week') || '').trim(),
        force: params.get('force') !== '0',
    };
}

function formatJson(value) {
    if (typeof value === 'string') {
        return value;
    }
    return JSON.stringify(value, null, 2);
}

function formatAccount(account) {
    if (!account) {
        return 'not signed in';
    }
    return `${account.name || account.username || 'signed in'} (${account.username || account.homeAccountId || 'no identifier'})`;
}

function setAuthStatus(session) {
    const authStatus = $('test-auth-status');
    const signInButton = $('btn-test-signin');
    const signOutButton = $('btn-test-signout');
    if (!authStatus || !signInButton || !signOutButton) {
        return;
    }

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
    if (params.teamId) {
        queryBits.push(`teamId=${params.teamId}`);
    }
    if (params.week) {
        queryBits.push(`week=${params.week}`);
    }
    queryBits.push(`force=${params.force ? '1' : '0'}`);
    meta.textContent = `Run one test with ?test=<name>. Available: auth-status, app-api, timeline-open-card-update, cache-snapshot, confirm-supported-timecard-filters, confirm-no-combined-week-query, joined-teams, resolve-team, timecards-top-1, selected-week-fetch-timing, server-vs-cache, confirm-no-lastmodified-filter, confirm-no-lastmodified-orderby. Current query: ${queryBits.join('&')}`;
}

function renderSummary(message) {
    const summary = $('test-summary');
    if (summary) {
        summary.textContent = message;
    }
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
    if (body) {
        body.textContent = formatJson(details);
    }
}

async function requireTeam(context) {
    if (context.team) {
        return context.team;
    }
    const resolved = await appApi.resolveTeamForSelfTest(context.params.teamId);
    context.team = resolved.team;
    context.teams = resolved.teams;
    if (!context.team) {
        throw new Error('No joined teams were available for the current account');
    }
    return context.team;
}

function assertDocumentedFilterSupport(support) {
    const expected = {
        clockInGe: 'clockInEvent/dateTime ge',
        clockInLe: 'clockInEvent/dateTime le',
        clockOutGe: 'clockOutEvent/dateTime ge',
        clockOutLe: 'clockOutEvent/dateTime le',
        stateEq: 'state eq',
        userIdEq: 'userId eq',
    };

    const unsupported = Object.entries(expected)
        .filter(([key]) => support?.[key]?.supported !== true)
        .map(([, label]) => label);

    if (unsupported.length) {
        throw new Error(`Expected documented timecard filters to be supported: ${unsupported.join(', ')}`);
    }

    return support;
}

function assertProbeRejected(result, label) {
    if (result !== false) {
        throw new Error(`Expected ${label} to be rejected, got ${String(result)}`);
    }
    return result;
}

const tests = [
    {
        name: 'auth-status',
        requiresAuth: false,
        run: async context => ({
            authenticated: context.session.authenticated,
            account: context.session.account
                ? {
                    name: context.session.account.name || '',
                    username: context.session.account.username || '',
                }
                : null,
            redirectUri: auth.getRuntimeConfig().redirectUri,
        }),
    },
    {
        name: 'app-api',
        requiresAuth: false,
        run: async () => ({
            hasAuthApi: Boolean(auth),
            hasAppApi: Boolean(appApi),
            appMethods: Object.keys(appApi || {}),
        }),
    },
    {
        name: 'timeline-open-card-update',
        requiresAuth: false,
        run: async () => {
            const clockIn = new Date('2026-04-29T16:00:00.000Z');
            const attemptedClockOut = new Date('2026-04-29T18:15:00.000Z');
            const projection = appApi.buildProjectedTimelineUpdate({
                id: 'self-test-open-card',
                state: 'clockedIn',
                userId: 'self-test-user',
                createdDateTime: clockIn.toISOString(),
                lastModifiedDateTime: clockIn.toISOString(),
                clockIn: {
                    dateTime: clockIn.toISOString(),
                    atApprovedLocation: false,
                },
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
        run: async context => appApi.getPersistedTimeCardCacheSnapshotForSelfTest(context.params.teamId),
    },
    {
        name: 'confirm-supported-timecard-filters',
        requiresAuth: true,
        run: async context => {
            const team = await requireTeam(context);
            return {
                teamId: team.id,
                support: assertDocumentedFilterSupport(
                    await appApi.probeTimeCardDocumentedFilterSupport(team.id, context.params.force),
                ),
            };
        },
    },
    {
        name: 'confirm-no-combined-week-query',
        requiresAuth: true,
        run: async context => {
            const team = await requireTeam(context);
            const result = assertProbeRejected(
                await appApi.probeTimeCardCombinedWeekQuerySupport(team.id, context.params.week, context.params.force),
                'combined weekly timecard query support',
            );
            return {
                teamId: team.id,
                supported: result,
            };
        },
    },
    {
        name: 'joined-teams',
        requiresAuth: true,
        run: async context => {
            const teams = await appApi.fetchJoinedTeamsForSelfTest();
            context.teams = teams;
            return {
                count: teams.length,
                firstTeam: teams[0] ? { id: teams[0].id, displayName: teams[0].displayName } : null,
            };
        },
    },
    {
        name: 'resolve-team',
        requiresAuth: true,
        run: async context => {
            const resolved = await appApi.resolveTeamForSelfTest(context.params.teamId);
            context.team = resolved.team;
            context.teams = resolved.teams;
            return {
                requestedTeamId: context.params.teamId || null,
                resolvedTeam: resolved.team ? { id: resolved.team.id, displayName: resolved.team.displayName } : null,
                availableTeams: resolved.teams.length,
            };
        },
    },
    {
        name: 'timecards-top-1',
        requiresAuth: true,
        run: async context => {
            const team = await requireTeam(context);
            const result = await appApi.fetchTimeCardsPageForSelfTest(team.id, { pageSize: 1 });
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
        run: async context => {
            const team = await requireTeam(context);
            return {
                teamId: team.id,
                timing: await appApi.fetchWeekTimeCardsForSelfTest(team.id, context.params.week, { includeTiming: true }),
            };
        },
    },
    {
        name: 'server-vs-cache',
        requiresAuth: true,
        run: async context => {
            const team = await requireTeam(context);
            const serverWeek = await appApi.fetchWeekTimeCardsForSelfTest(team.id, context.params.week, { includeTiming: false });

            return {
                teamId: team.id,
                selectedWeek: serverWeek.weekStart,
                cache: appApi.getPersistedTimeCardCacheSnapshotForSelfTest(team.id),
                serverWeek,
            };
        },
    },
    {
        name: 'confirm-no-lastmodified-filter',
        requiresAuth: true,
        run: async context => {
            const team = await requireTeam(context);
            const result = assertProbeRejected(
                await appApi.probeTimeCardLastModifiedFilterSupport(team.id, context.params.force),
                'lastModifiedDateTime filter support',
            );
            return {
                teamId: team.id,
                supported: result,
            };
        },
    },
    {
        name: 'confirm-no-lastmodified-orderby',
        requiresAuth: true,
        run: async context => {
            const team = await requireTeam(context);
            const result = assertProbeRejected(
                await appApi.probeTimeCardLastModifiedOrderBySupport(team.id, context.params.force),
                'lastModifiedDateTime orderby support',
            );
            return {
                teamId: team.id,
                supported: result,
            };
        },
    },
];

function getRequestedTests(params) {
    if (!params.test || params.test === 'all') {
        return tests;
    }
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

    const context = {
        params,
        session,
        team: null,
        teams: null,
    };

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
            try {
                await auth.signIn({ navigateToApp: false });
            } finally {
                signInButton.disabled = false;
                await runSuite();
            }
        });
    }

    if (signOutButton) {
        signOutButton.addEventListener('click', async () => {
            signOutButton.disabled = true;
            try {
                await auth.signOut({ navigateToLogin: false });
            } finally {
                signOutButton.disabled = false;
                await runSuite();
            }
        });
    }
}

if (!auth.isPopupContext()) {
    bindActions();
    void runSuite();
}

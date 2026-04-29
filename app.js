/**
 * Teams Timecards — app.js
 * Main app shell: team selection, timecards, editing, drag/resize.
 */

'use strict';

(function () {

    // ─────────────────────────────────────────────
    // App configuration
    // ─────────────────────────────────────────────
    const SESSION_KEY_SELECTED_TEAM_ID = 'tc_selected_team_id';
    const SESSION_KEY_SELECTED_WEEK = 'tc_selected_week';
    const SESSION_CACHE_KEY_JOINED_TEAMS = 'tc_joined_teams';
    const SNAP_MINUTES = 5;
    const DEFAULT_BREAK_DURATION_MS = 10 * 60 * 1000;
    const EXTEND_TO_NOW_LAG_MS = 30 * 1000;
    const DAY_MS = 24 * 60 * 60 * 1000;
    const CACHE_STALE_MS = 45 * 1000;
    const BACKGROUND_REFRESH_MS = 30 * 1000;
    const TIMECARD_PAGE_SIZE = 250;

    const auth = window.timecardsAuth;
    const graphFetch = (...args) => auth.graphFetch(...args);
    const graphFetchBeta = (...args) => auth.graphFetchBeta(...args);
    const defaultPageTitle = document.title || 'Teams Timecards';

    // App state
    let selectedTeam = null;
    let allTimeCards = [];       // flat array of normalized timecard objects
    let activeTimeCard = null;   // currently open (not clockedOut) timecard if any
    let highlightedCardId = null;
    let latestVisibleCardId = null;
    let selectedWeekStart = null;
    let liveClockTimerId = null;
    let backgroundRefreshTimerId = null;
    const timeCardCache = createTeamCacheState();
    // Edit modal state
    let editTarget = null;       // { card, weekStartTs }

    // Confirm callback
    let confirmCallback = null;

    // ─────────────────────────────────────────────
    // DOM refs
    // ─────────────────────────────────────────────
    const $ = id => document.getElementById(id);

    const appEl = $('app');
    const userInfo = $('user-info');
    const btnSignout = $('btn-signout');
    const toolbarTitle = $('toolbar-title');
    const toolbarActions = $('toolbar-actions');
    const btnChangeTeam = $('btn-change-team');
    const weekNav = $('week-nav');
    const btnWeekPrev = $('btn-week-prev');
    const btnWeekNext = $('btn-week-next');
    const weekPicker = $('week-picker');
    const weekRangeLabel = $('week-range-label');
    const currentWeekTotal = $('current-week-total');
    const stateBadge = $('current-state-badge');
    const btnClockIn = $('btn-clock-in');
    const btnClockOut = $('btn-clock-out');
    const btnStartBreak = $('btn-start-break');
    const btnEndBreak = $('btn-end-break');
    const teamPickerPanel = $('team-picker-panel');
    const teamPickerMsg = $('team-picker-msg');
    const teamPickerSelect = $('team-picker-select');
    const btnTeamSave = $('btn-team-save');
    const noTeamMsg = $('no-team-msg');
    const tcLoading = $('timecards-loading');
    const weeksContainer = $('weeks-container');
    const dragTooltip = $('drag-tooltip');
    const errorBanner = $('error-banner');

    function bindEvent(element, eventName, handler) {
        if (element) {
            element.addEventListener(eventName, handler);
        }
    }

    // Auth button
    bindEvent(btnSignout, 'click', handleSignOut);

    // Toolbar action buttons
    bindEvent(btnClockIn, 'click', () => doClockIn());
    bindEvent(btnClockOut, 'click', () => doClockOut());
    bindEvent(btnStartBreak, 'click', () => doStartBreak());
    bindEvent(btnEndBreak, 'click', () => doEndBreak());

    // Team picker
    bindEvent(btnChangeTeam, 'click', () => {
        showTeamPicker('Choose a different team. Your selection is saved in your browser.');
    });
    bindEvent(btnTeamSave, 'click', saveSelectedTeam);
    bindEvent(btnWeekPrev, 'click', () => shiftSelectedWeek(-7));
    bindEvent(btnWeekNext, 'click', () => shiftSelectedWeek(7));
    bindEvent(weekPicker, 'change', onWeekPickerChange);
    bindEvent(document, 'visibilitychange', handleVisibilityChange);

    // Edit modal buttons
    bindEvent($('btn-edit-cancel'), 'click', closeEditModal);
    bindEvent($('btn-edit-save'), 'click', saveEditModal);
    bindEvent($('btn-add-break-edit'), 'click', addBreakRow);

    // Confirm modal buttons
    bindEvent($('btn-confirm-cancel'), 'click', () => closeConfirm(false));
    bindEvent($('btn-confirm-ok'), 'click', () => closeConfirm(true));

    // ─────────────────────────────────────────────
    // Boot
    // ─────────────────────────────────────────────
    (async function boot() {
        if (document.body.dataset.page !== 'app' || auth.isPopupContext()) {
            return;
        }

        const authenticated = await auth.requireAppSession();
        if (!authenticated) {
            return;
        }

        const account = auth.getCurrentAccount();
        userInfo.textContent = account?.name || account?.username || '';
        btnSignout.style.display = 'inline-flex';
        appEl.style.display = 'flex';
        await loadTeams();
    })();

    async function handleSignOut() {
        stopLiveClockUpdates();
        stopBackgroundRefreshLoop();
        resetTeamCache(timeCardCache, '', { clearProbes: true });
        selectedTeam = null;
        activeTimeCard = null;
        highlightedCardId = null;
        latestVisibleCardId = null;
        allTimeCards = [];
        localStorage.removeItem(SESSION_KEY_SELECTED_TEAM_ID);
        localStorage.removeItem(SESSION_KEY_SELECTED_WEEK);
        sessionStorage.removeItem(SESSION_CACHE_KEY_JOINED_TEAMS);
        updateDocumentTitle();
        await auth.signOut();
    }

    // ─────────────────────────────────────────────
    // Teams
    // ─────────────────────────────────────────────
    let allTeams = [];

    function showStatusPanel(message, tone = 'default') {
        updateDocumentTitle();
        teamPickerPanel.style.display = 'none';
        noTeamMsg.innerHTML = message;
        noTeamMsg.style.display = 'block';
        noTeamMsg.style.color = tone === 'error' ? 'var(--danger)' : 'var(--text2)';
        weeksContainer.style.display = 'none';
        tcLoading.style.display = 'none';
        weekNav.style.display = 'none';
        currentWeekTotal.style.display = 'none';
        toolbarActions.style.display = 'none';
        stateBadge.style.display = 'none';
        btnChangeTeam.style.display = 'none';
    }

    function populateTeamPicker(preferredTeamId) {
        teamPickerSelect.innerHTML = allTeams.map(team =>
            `<option value="${team.id}">${escHtml(team.displayName)}</option>`
        ).join('');

        const fallbackTeamId = preferredTeamId || localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID) || allTeams[0]?.id || '';
        if (fallbackTeamId) {
            teamPickerSelect.value = fallbackTeamId;
        }

        btnTeamSave.disabled = !allTeams.length;
        teamPickerSelect.disabled = !allTeams.length;
    }

    function showTeamPicker(message) {
        stopBackgroundRefreshLoop();
        selectedTeam = null;
        activeTimeCard = null;
        highlightedCardId = null;
        latestVisibleCardId = null;
        updateDocumentTitle();
        teamPickerMsg.textContent = message;
        populateTeamPicker(localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID));

        toolbarTitle.textContent = 'Choose team';
        weekNav.style.display = 'none';
        currentWeekTotal.style.display = 'none';
        toolbarActions.style.display = 'none';
        stateBadge.style.display = 'none';
        btnChangeTeam.style.display = 'none';
        tcLoading.style.display = 'none';
        weeksContainer.style.display = 'none';
        noTeamMsg.style.display = 'none';
        teamPickerPanel.style.display = 'block';
    }

    async function loadTeams() {
        toolbarTitle.textContent = 'Loading team…';
        showStatusPanel('<span class="spinner"></span>Loading teams…');
        try {
            allTeams = await fetchJoinedTeamsForSelfTest({ preferCache: true });
            if (!allTeams.length) {
                localStorage.removeItem(SESSION_KEY_SELECTED_TEAM_ID);
                showStatusPanel('No teams found.');
                return;
            }

            const savedTeamId = localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID);
            const savedTeam = allTeams.find(team => team.id === savedTeamId);
            if (savedTeam) {
                await selectTeam(savedTeam);
                return;
            }

            if (savedTeamId) {
                localStorage.removeItem(SESSION_KEY_SELECTED_TEAM_ID);
            }

            if (allTeams.length === 1) {
                await selectTeam(allTeams[0]);
                return;
            }

            showTeamPicker(
                savedTeamId
                    ? 'Your saved team is no longer available. Choose a new team.'
                    : 'Choose the team to use for timecards. This is saved in your browser.'
            );
        } catch (e) {
            showStatusPanel(`Failed to load teams.<br>${escHtml(e.message)}`, 'error');
        }
    }

    function normalizeJoinedTeams(teams) {
        return (teams || [])
            .filter(team => team?.id && team?.displayName)
            .map(team => ({
                id: team.id,
                displayName: team.displayName,
                description: team.description || '',
            }))
            .sort((a, b) => a.displayName.localeCompare(b.displayName));
    }

    function readCachedJoinedTeams() {
        const raw = sessionStorage.getItem(SESSION_CACHE_KEY_JOINED_TEAMS);
        if (!raw) {
            return null;
        }

        try {
            const parsed = JSON.parse(raw);
            if (!Array.isArray(parsed)) {
                throw new Error('Joined team cache was not an array.');
            }
            return normalizeJoinedTeams(parsed);
        } catch {
            sessionStorage.removeItem(SESSION_CACHE_KEY_JOINED_TEAMS);
            return null;
        }
    }

    function writeCachedJoinedTeams(teams) {
        sessionStorage.setItem(SESSION_CACHE_KEY_JOINED_TEAMS, JSON.stringify(normalizeJoinedTeams(teams)));
    }

    async function fetchJoinedTeamsForSelfTest({ preferCache = false } = {}) {
        if (preferCache) {
            const cachedTeams = readCachedJoinedTeams();
            if (cachedTeams !== null) {
                return cachedTeams;
            }
        }

        const data = await graphFetch('/me/joinedTeams?$select=id,displayName,description');
        const teams = normalizeJoinedTeams(data?.value || []);
        writeCachedJoinedTeams(teams);
        return teams;
    }

    async function resolveTeamForSelfTest(preferredTeamId = '') {
        const teams = await fetchJoinedTeamsForSelfTest();
        const savedTeamId = preferredTeamId || localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID) || '';
        const team = teams.find(item => item.id === savedTeamId) || teams[0] || null;
        return { team, teams };
    }

    async function fetchTimeCardsPageForSelfTest(teamId, options = {}) {
        if (!teamId) {
            throw new Error('A teamId is required for timecard self tests');
        }
        return graphFetch(buildTimeCardsPageUrl(teamId, options));
    }

    async function saveSelectedTeam() {
        const team = allTeams.find(item => item.id === teamPickerSelect.value);
        if (!team) {
            toast('Choose a team first', 'error');
            return;
        }

        btnTeamSave.disabled = true;
        try {
            await selectTeam(team);
        } finally {
            btnTeamSave.disabled = false;
        }
    }

    async function selectTeam(team) {
        selectedTeam = team;
        localStorage.setItem(SESSION_KEY_SELECTED_TEAM_ID, team.id);
        toolbarTitle.textContent = team.displayName;
        toolbarActions.style.display = 'none';
        stateBadge.style.display = 'none';
        btnChangeTeam.style.display = 'inline-flex';
        weekNav.style.display = 'flex';
        teamPickerPanel.style.display = 'none';
        noTeamMsg.style.display = 'none';
        ensureSelectedWeek();
        syncWeekControls();
        await loadTimeCards();
        syncBackgroundRefreshLoop();
    }

    // ─────────────────────────────────────────────
    // TimeCards loading
    // ─────────────────────────────────────────────
    async function loadTimeCards() {
        if (!selectedTeam) return;
        teamPickerPanel.style.display = 'none';
        noTeamMsg.style.display = 'none';
        ensureSelectedWeek();
        syncWeekControls();

        const cache = getTeamCache(selectedTeam.id);
        const selectedWeekKey = getWeekCacheKey(selectedWeekStart);
        const hasCachedWeek = cache.currentWeekKey === selectedWeekKey;
        const cachedCards = hasCachedWeek ? cache.currentWeekCards : [];
        const cachedFetchedAt = hasCachedWeek ? cache.currentWeekFetchedAt : 0;

        if (cachedCards.length) {
            applyTimeCards(cachedCards);
            tcLoading.style.display = 'none';
        } else {
            weeksContainer.style.display = 'none';
            tcLoading.style.display = 'block';
        }

        const shouldRefresh = !cachedCards.length
            || (Date.now() - cachedFetchedAt) > CACHE_STALE_MS;
        if (!shouldRefresh) {
            return;
        }

        try {
            await refreshTimeCards({ forceSpinner: !cachedCards.length });
        } catch (e) {
            if (!cachedCards.length) {
                tcLoading.style.display = 'none';
                weeksContainer.style.display = 'none';
            }
            showError('Failed to load timecards: ' + e.message);
        }
    }

    async function refreshTimeCards({ forceSpinner = false, reset = false } = {}) {
        if (!selectedTeam) return;
        ensureSelectedWeek();
        const teamId = selectedTeam.id;
        const targetWeekStart = getResolvedWeekStart(selectedWeekStart);
        const targetWeekKey = getWeekCacheKey(targetWeekStart);
        const cache = getTeamCache(teamId);

        if (reset) {
            resetTeamCache(cache, teamId);
        }

        const hasCachedWeek = cache.currentWeekKey === targetWeekKey;

        if (cache.pendingWeekKey === targetWeekKey && cache.pendingWeekPromise) {
            return cache.pendingWeekPromise;
        }

        if (forceSpinner && !(hasCachedWeek && cache.currentWeekCards.length)) {
            weeksContainer.style.display = 'none';
            tcLoading.style.display = 'block';
        }

        const request = (async () => {
            const snapshot = await fetchWeekTimeCards(teamId, targetWeekStart);
            setCachedWeek(cache, targetWeekStart, snapshot.cards, snapshot.fetchedAt);
            maybeWarnAboutPagedWeek(cache, targetWeekKey, snapshot.pageCount);
            const knownActiveCardWasReturned = Boolean(
                cache.activeCard && snapshot.cards.some(card => card.id === cache.activeCard.id),
            );
            cache.activeCard = snapshot.activeCard
                || (knownActiveCardWasReturned ? null : cache.activeCard);

            if (selectedTeam && selectedTeam.id === teamId && getWeekCacheKey(selectedWeekStart) === targetWeekKey) {
                applyTimeCards(cache.currentWeekCards);
            }

            return cache.currentWeekCards;
        })().finally(() => {
            if (cache.pendingWeekKey === targetWeekKey) {
                cache.pendingWeekKey = '';
                cache.pendingWeekPromise = null;
            }
        });

        cache.pendingWeekKey = targetWeekKey;
        cache.pendingWeekPromise = request;
        return request;
    }

    function createTeamCacheState(teamId = '') {
        return {
            teamId,
            activeCard: null,
            currentWeekKey: '',
            currentWeekCards: [],
            currentWeekFetchedAt: 0,
            pendingWeekKey: '',
            pendingWeekPromise: null,
            warnedMultiPageWeekKey: '',
            lastModifiedFilterSupport: null,
            lastModifiedFilterProbePromise: null,
            lastModifiedOrderBySupport: null,
            lastModifiedOrderByProbePromise: null,
            filterSupportMatrix: null,
            filterSupportProbePromise: null,
            combinedWeekQuerySupport: null,
            combinedWeekQueryProbePromise: null,
        };
    }

    function getTeamCache(teamId) {
        if (timeCardCache.teamId !== teamId) {
            resetTeamCache(timeCardCache, teamId, { clearProbes: true });
        }
        return timeCardCache;
    }

    function buildTimeCardsPageUrl(teamId, { newestFirst = false, pageSize = TIMECARD_PAGE_SIZE } = {}) {
        let url = `/teams/${teamId}/schedule/timeCards?$top=${pageSize}`;
        if (newestFirst) {
            url += `&$orderby=${encodeURIComponent('lastModifiedDateTime desc')}`;
        }
        return url;
    }

    function getResolvedWeekStart(weekStartInput = null) {
        if (weekStartInput instanceof Date && !Number.isNaN(weekStartInput.getTime())) {
            return getWeekStart(weekStartInput);
        }

        if (typeof weekStartInput === 'string' && weekStartInput) {
            const parsed = /^\d{4}-\d{2}-\d{2}$/.test(weekStartInput)
                ? new Date(`${weekStartInput}T00:00:00`)
                : new Date(weekStartInput);
            if (!Number.isNaN(parsed.getTime())) {
                return getWeekStart(parsed);
            }
        }

        if (selectedWeekStart instanceof Date && !Number.isNaN(selectedWeekStart.getTime())) {
            return getWeekStart(selectedWeekStart);
        }

        const storedWeek = localStorage.getItem(SESSION_KEY_SELECTED_WEEK);
        if (storedWeek) {
            const parsed = new Date(`${storedWeek}T00:00:00`);
            if (!Number.isNaN(parsed.getTime())) {
                return getWeekStart(parsed);
            }
        }

        const fallback = activeTimeCard?.clockIn?.dateTime
            ? new Date(activeTimeCard.clockIn.dateTime)
            : new Date();
        return getWeekStart(fallback);
    }

    function getWeekCacheKey(weekStartInput) {
        return formatDateInputValue(getResolvedWeekStart(weekStartInput));
    }

    function setCachedWeek(cache, weekStartInput, cards, fetchedAt = Date.now()) {
        cache.currentWeekKey = getWeekCacheKey(weekStartInput);
        cache.currentWeekCards = dedupeCardsById(cards);
        cache.currentWeekFetchedAt = fetchedAt;
        return cache.currentWeekCards;
    }

    function resetTeamCache(cache, teamId = cache.teamId, { clearProbes = false } = {}) {
        cache.teamId = teamId;
        cache.activeCard = null;
        cache.currentWeekKey = '';
        cache.currentWeekCards = [];
        cache.currentWeekFetchedAt = 0;
        cache.pendingWeekKey = '';
        cache.pendingWeekPromise = null;
        cache.warnedMultiPageWeekKey = '';

        if (clearProbes) {
            cache.lastModifiedFilterSupport = null;
            cache.lastModifiedFilterProbePromise = null;
            cache.lastModifiedOrderBySupport = null;
            cache.lastModifiedOrderByProbePromise = null;
            cache.filterSupportMatrix = null;
            cache.filterSupportProbePromise = null;
            cache.combinedWeekQuerySupport = null;
            cache.combinedWeekQueryProbePromise = null;
        }
    }

    function maybeWarnAboutPagedWeek(cache, weekKey, pageCount) {
        if (pageCount > 1) {
            if (cache.warnedMultiPageWeekKey !== weekKey) {
                toast(`This week required ${pageCount} pages to load all timecards.`, 'warn');
                cache.warnedMultiPageWeekKey = weekKey;
            }
            return;
        }

        if (cache.warnedMultiPageWeekKey === weekKey) {
            cache.warnedMultiPageWeekKey = '';
        }
    }

    async function probeTimeCardLastModifiedFilterSupport(teamId, force = false) {
        if (!teamId) return null;
        const cache = getTeamCache(teamId);
        if (!force && cache.lastModifiedFilterSupport !== null) {
            return cache.lastModifiedFilterSupport;
        }
        if (cache.lastModifiedFilterProbePromise) {
            return cache.lastModifiedFilterProbePromise;
        }

        const filterExpr = encodeURIComponent('lastModifiedDateTime ge 1970-01-01T00:00:00Z');
        const probeUrl = `/teams/${teamId}/schedule/timeCards?$top=1&$filter=${filterExpr}`;

        const request = graphFetch(probeUrl)
            .then(() => {
                cache.lastModifiedFilterSupport = true;
                console.info('[timecards] Graph accepts lastModifiedDateTime filter.');
                return true;
            })
            .catch(error => {
                const message = error?.message || 'Unknown Graph error';
                const rejectedByGraph = /lastmodifieddatetime|filter|not allowed|unsupported/i.test(message);
                if (rejectedByGraph) {
                    cache.lastModifiedFilterSupport = false;
                    console.info('[timecards] Graph rejected lastModifiedDateTime filter.', message);
                    return false;
                }

                cache.lastModifiedFilterSupport = null;
                console.warn('[timecards] Could not probe lastModifiedDateTime filter.', message);
                return null;
            })
            .finally(() => {
                cache.lastModifiedFilterProbePromise = null;
            });

        cache.lastModifiedFilterProbePromise = request;
        return request;
    }

    async function probeTimeCardLastModifiedOrderBySupport(teamId, force = false) {
        if (!teamId) return null;
        const cache = getTeamCache(teamId);
        if (!force && cache.lastModifiedOrderBySupport !== null) {
            return cache.lastModifiedOrderBySupport;
        }
        if (cache.lastModifiedOrderByProbePromise) {
            return cache.lastModifiedOrderByProbePromise;
        }

        const probeUrl = buildTimeCardsPageUrl(teamId, { newestFirst: true, pageSize: 1 });
        const request = graphFetch(probeUrl)
            .then(() => {
                cache.lastModifiedOrderBySupport = true;
                console.info('[timecards] Graph accepts lastModifiedDateTime orderby.');
                return true;
            })
            .catch(error => {
                const message = error?.message || 'Unknown Graph error';
                const rejectedByGraph = /lastmodifieddatetime|orderby|order by|sort|not allowed|unsupported/i.test(message);
                if (rejectedByGraph) {
                    cache.lastModifiedOrderBySupport = false;
                    console.info('[timecards] Graph rejected lastModifiedDateTime orderby.', message);
                    return false;
                }

                cache.lastModifiedOrderBySupport = null;
                console.warn('[timecards] Could not probe lastModifiedDateTime orderby.', message);
                return null;
            })
            .finally(() => {
                cache.lastModifiedOrderByProbePromise = null;
            });

        cache.lastModifiedOrderByProbePromise = request;
        return request;
    }

    function isRejectedTimeCardFilterProbeError(message) {
        return /filter|unsupported|not allowed|could not find a property|invalid/i.test(message || '');
    }

    function escapeODataStringLiteral(value) {
        return String(value || '').replace(/'/g, "''");
    }

    function buildTimeCardFilterProbeUrl(teamId, filterExpr) {
        return `/teams/${teamId}/schedule/timeCards?$top=1&$filter=${encodeURIComponent(filterExpr)}`;
    }

    async function probeSingleTimeCardFilterSupport(teamId, definition) {
        let rejectedMessage = '';

        for (const expression of definition.expressions) {
            try {
                await graphFetch(buildTimeCardFilterProbeUrl(teamId, expression));
                return {
                    label: definition.label,
                    supported: true,
                    acceptedExpression: expression,
                    triedExpressions: definition.expressions,
                    message: 'Graph accepted the filter expression.',
                };
            } catch (error) {
                const message = error?.message || 'Unknown Graph error';
                if (isRejectedTimeCardFilterProbeError(message)) {
                    rejectedMessage = message;
                    continue;
                }

                return {
                    label: definition.label,
                    supported: null,
                    acceptedExpression: null,
                    triedExpressions: definition.expressions,
                    message,
                };
            }
        }

        return {
            label: definition.label,
            supported: false,
            acceptedExpression: null,
            triedExpressions: definition.expressions,
            message: rejectedMessage || 'Graph rejected the filter expression.',
        };
    }

    async function probeTimeCardDocumentedFilterSupport(teamId, force = false) {
        if (!teamId) return null;
        const cache = getTeamCache(teamId);
        if (!force && cache.filterSupportMatrix) {
            return cache.filterSupportMatrix;
        }
        if (cache.filterSupportProbePromise) {
            return cache.filterSupportProbePromise;
        }

        const now = Date.now();
        const lowerBoundIso = new Date(now - (30 * DAY_MS)).toISOString();
        const upperBoundIso = new Date(now + DAY_MS).toISOString();
        const currentUserId = auth.getCurrentAccount()?.localAccountId
            || auth.getCurrentAccount()?.homeAccountId
            || '00000000-0000-0000-0000-000000000000';

        const definitions = [
            {
                key: 'clockInGe',
                label: 'clockInEvent/dateTime ge',
                expressions: [
                    `clockInEvent/dateTime ge ${lowerBoundIso}`,
                    `clockInEvent/datetime ge ${lowerBoundIso}`,
                ],
            },
            {
                key: 'clockInLe',
                label: 'clockInEvent/dateTime le',
                expressions: [
                    `clockInEvent/dateTime le ${upperBoundIso}`,
                    `clockInEvent/datetime le ${upperBoundIso}`,
                ],
            },
            {
                key: 'clockOutGe',
                label: 'clockOutEvent/dateTime ge',
                expressions: [
                    `clockOutEvent/dateTime ge ${lowerBoundIso}`,
                    `clockOutEvent/datetime ge ${lowerBoundIso}`,
                ],
            },
            {
                key: 'clockOutLe',
                label: 'clockOutEvent/dateTime le',
                expressions: [
                    `clockOutEvent/dateTime le ${upperBoundIso}`,
                    `clockOutEvent/datetime le ${upperBoundIso}`,
                ],
            },
            {
                key: 'stateEq',
                label: 'state eq',
                expressions: ["state eq 'clockedIn'"],
            },
            {
                key: 'userIdEq',
                label: 'userId eq',
                expressions: [`userId eq '${escapeODataStringLiteral(currentUserId)}'`],
            },
        ];

        const request = Promise.all(definitions.map(async definition => ({
            key: definition.key,
            result: await probeSingleTimeCardFilterSupport(teamId, definition),
        })))
            .then(entries => {
                const matrix = {
                    checkedAt: new Date().toISOString(),
                };

                entries.forEach(entry => {
                    matrix[entry.key] = entry.result;
                });

                cache.filterSupportMatrix = matrix;
                return matrix;
            })
            .finally(() => {
                cache.filterSupportProbePromise = null;
            });

        cache.filterSupportProbePromise = request;
        return request;
    }

    function buildCombinedWeekTimeCardFilter(weekStartInput) {
        const { weekStart, weekEndInclusive } = getWeekRange(weekStartInput);
        const weekStartIso = weekStart.toISOString();
        const weekEndIso = weekEndInclusive.toISOString();
        return `((clockInEvent/dateTime le ${weekEndIso} and clockOutEvent/dateTime ge ${weekStartIso}) or state eq 'clockedIn' or state eq 'onBreak')`;
    }

    function buildSelectedWeekClockInWindowFilter(weekStartInput) {
        const { weekStart, weekEndInclusive } = getWeekRange(weekStartInput);
        const weekStartIso = weekStart.toISOString();
        const weekEndIso = weekEndInclusive.toISOString();
        return `clockInEvent/dateTime ge ${weekStartIso} and clockInEvent/dateTime le ${weekEndIso}`;
    }

    async function probeTimeCardCombinedWeekQuerySupport(teamId, weekStartInput, force = false) {
        if (!teamId) return null;
        const cache = getTeamCache(teamId);
        if (!force && cache.combinedWeekQuerySupport !== null) {
            return cache.combinedWeekQuerySupport;
        }
        if (cache.combinedWeekQueryProbePromise) {
            return cache.combinedWeekQueryProbePromise;
        }

        const filterExpr = buildCombinedWeekTimeCardFilter(weekStartInput);
        const request = graphFetch(buildTimeCardsFilterUrl(teamId, filterExpr, { pageSize: 1 }))
            .then(() => {
                cache.combinedWeekQuerySupport = true;
                return true;
            })
            .catch(error => {
                const message = error?.message || 'Unknown Graph error';
                const rejectedByGraph = /filter|unsupported|not allowed|could not find a property|invalid/i.test(message);
                if (rejectedByGraph) {
                    cache.combinedWeekQuerySupport = false;
                    return false;
                }

                cache.combinedWeekQuerySupport = null;
                console.warn('[timecards] Could not probe combined weekly timecard filter.', message);
                return null;
            })
            .finally(() => {
                cache.combinedWeekQueryProbePromise = null;
            });

        cache.combinedWeekQueryProbePromise = request;
        return request;
    }

    async function confirmWeekTimeCardQuerySupport(teamId, weekStartInput) {
        const support = await probeTimeCardDocumentedFilterSupport(teamId, false);
        const requiredEntries = {
            clockInGe: 'clockInEvent/dateTime ge',
            clockInLe: 'clockInEvent/dateTime le',
        };

        const unsupported = Object.entries(requiredEntries)
            .filter(([key]) => support?.[key]?.supported !== true)
            .map(([, label]) => label);

        if (unsupported.length) {
            throw new Error(`Required timecard filters are unavailable: ${unsupported.join(', ')}`);
        }

        return true;
    }

    function mergeCardsIntoCache(cache, pageCards) {
        if (!pageCards.length) {
            return;
        }

        const byId = new Map(cache.cards.map(card => [card.id, card]));
        pageCards.forEach(card => {
            byId.set(card.id, card);
        });
        cache.cards = Array.from(byId.values());
    }

    function isOpenCardState(card) {
        if (!card) {
            return false;
        }
        const state = deriveCardState(card);
        return state === 'clockedIn' || state === 'onBreak';
    }

    function getWeekRange(weekStartInput = null) {
        const weekStart = getResolvedWeekStart(weekStartInput);
        const weekEndExclusive = new Date(weekStart.getTime() + (7 * DAY_MS));
        const weekEndInclusive = new Date(weekEndExclusive.getTime() - 1);

        return {
            weekStart,
            weekEndExclusive,
            weekEndInclusive,
        };
    }

    function cardOverlapsRange(card, rangeStartMs, rangeEndExclusiveMs, referenceNowMs = Date.now()) {
        if (!card?.clockIn?.dateTime) {
            return false;
        }

        const startMs = new Date(card.clockIn.dateTime).getTime();
        const endMs = card.clockOut?.dateTime
            ? new Date(card.clockOut.dateTime).getTime()
            : referenceNowMs;

        return startMs < rangeEndExclusiveMs && endMs >= rangeStartMs;
    }

    function filterCardsForWeek(cards, weekStartInput = null, referenceNowMs = Date.now()) {
        const { weekStart, weekEndExclusive } = getWeekRange(weekStartInput);
        return dedupeCardsById(cards).filter(card => cardOverlapsRange(
            card,
            weekStart.getTime(),
            weekEndExclusive.getTime(),
            referenceNowMs,
        ));
    }

    function buildTimeCardsFilterUrl(teamId, filterExpr, { pageSize = TIMECARD_PAGE_SIZE } = {}) {
        return `/teams/${teamId}/schedule/timeCards?$top=${pageSize}&$filter=${encodeURIComponent(filterExpr)}`;
    }

    async function fetchTimeCardsByFilter(teamId, filterExpr, { pageSize = TIMECARD_PAGE_SIZE } = {}) {
        const snapshot = { cards: [] };
        let nextUrl = buildTimeCardsFilterUrl(teamId, filterExpr, { pageSize });
        let pageCount = 0;

        while (nextUrl) {
            const data = await graphFetch(nextUrl);
            const pageCards = (data?.value || []).map(normalizeCard);
            pageCount += 1;
            mergeCardsIntoCache(snapshot, pageCards);
            nextUrl = data?.['@odata.nextLink'] || null;
        }

        return {
            cards: snapshot.cards,
            pageCount,
        };
    }

    async function fetchWeekTimeCards(teamId, weekStartInput, { includeTiming = false } = {}) {
        const filterExpr = buildSelectedWeekClockInWindowFilter(weekStartInput);
        const startedAt = includeTiming ? performance.now() : 0;
        const result = await fetchTimeCardsByFilter(teamId, filterExpr, { pageSize: TIMECARD_PAGE_SIZE });
        const durationMs = includeTiming ? Math.round(performance.now() - startedAt) : null;
        const mergedCards = dedupeCardsById(result.cards);
        const filteredCards = filterCardsForWeek(mergedCards, weekStartInput);
        const activeCards = filteredCards
            .filter(card => isOpenCardState(card))
            .sort((a, b) => getCardSortMs(b) - getCardSortMs(a));
        const fetchedAt = Date.now();

        return {
            weekStart: getResolvedWeekStart(weekStartInput),
            weekKey: getWeekCacheKey(weekStartInput),
            fetchedAt,
            activeCard: activeCards.find(card => isOpenCardState(card)) || null,
            cards: filteredCards,
            timings: includeTiming ? {
                totalDurationMs: durationMs,
                queries: [{
                    key: 'selectedWeekClockInWindow',
                    label: 'Selected week cards by clock-in window',
                    filter: filterExpr,
                    cardsCount: mergedCards.length,
                    pageCount: result.pageCount,
                    durationMs,
                }],
            } : null,
        };
    }

    function applyTimeCards(cards) {
        allTimeCards = filterCardsForWeek(cards, selectedWeekStart);
        activeTimeCard = selectedTeam
            ? (getTeamCache(selectedTeam.id).activeCard || null)
            : null;
        if (highlightedCardId && !allTimeCards.some(card => card.id === highlightedCardId)) {
            highlightedCardId = null;
        }
        updateToolbarState();
        renderWeeks();
    }

    function deriveCardState(card) {
        if (card.clockOut) {
            return 'clockedOut';
        }
        if (card.breaks.some(item => item.start && !item.end)) {
            return 'onBreak';
        }
        return 'clockedIn';
    }

    function cloneCardBreaks(breaks) {
        return (breaks || []).map(item => ({
            breakId: item.breakId,
            start: item.start ? {
                dateTime: item.start.dateTime,
                notes: item.start.notes,
            } : null,
            end: item.end ? {
                dateTime: item.end.dateTime,
                notes: item.end.notes,
            } : null,
        }));
    }

    function cloneItemBody(itemBody) {
        return itemBody ? {
            contentType: itemBody.contentType,
            content: itemBody.content,
        } : null;
    }

    function getUpdatedCardForLocalState(card, { clockIn, clockOut }) {
        const updatedCard = {
            ...card,
            clockIn: clockIn ? {
                ...(card.clockIn || {}),
                dateTime: clockIn.toISOString(),
                atApprovedLocation: card.clockIn?.atApprovedLocation ?? false,
            } : null,
            clockOut: clockOut ? {
                ...(card.clockOut || {}),
                dateTime: clockOut.toISOString(),
                atApprovedLocation: card.clockOut?.atApprovedLocation ?? false,
            } : null,
            breaks: cloneCardBreaks(card.breaks),
            lastModifiedDateTime: new Date().toISOString(),
        };

        updatedCard.state = deriveCardState(updatedCard);
        return updatedCard;
    }

    function getUpdatedCardForEdit(card, { clockIn, clockOut, breaks, notes }) {
        const updatedCard = {
            ...card,
            clockIn: clockIn ? {
                ...(card.clockIn || {}),
                dateTime: clockIn.toISOString(),
                atApprovedLocation: card.clockIn?.atApprovedLocation ?? false,
            } : null,
            clockOut: clockOut ? {
                ...(card.clockOut || {}),
                dateTime: clockOut.toISOString(),
                atApprovedLocation: card.clockOut?.atApprovedLocation ?? false,
            } : null,
            breaks: cloneCardBreaks(breaks),
            notes: notes ? {
                contentType: 'text',
                content: notes,
            } : null,
            lastModifiedDateTime: new Date().toISOString(),
        };

        updatedCard.state = deriveCardState(updatedCard);
        return updatedCard;
    }

    function buildTimeCardUpdateBody(card) {
        const body = {
            clockInEvent: card.clockIn ? {
                dateTime: card.clockIn.dateTime,
                atApprovedLocation: card.clockIn.atApprovedLocation ?? false,
            } : undefined,
            clockOutEvent: card.clockOut ? {
                dateTime: card.clockOut.dateTime,
                atApprovedLocation: card.clockOut.atApprovedLocation ?? false,
            } : undefined,
            breaks: card.breaks.map(b => ({
                breakId: b.breakId,
                start: b.start ? { dateTime: b.start.dateTime } : undefined,
                end: b.end ? { dateTime: b.end.dateTime } : undefined,
                notes: { contentType: 'text', content: '' },
            })),
            notes: cloneItemBody(card.notes) || undefined,
        };

        if (!body.clockOutEvent) delete body.clockOutEvent;
        if (!body.notes) delete body.notes;
        return body;
    }

    function getEventDateTime(event) {
        return event?.dateTime || '';
    }

    function getItemBodyContent(itemBody) {
        return itemBody?.content || '';
    }

    function getSelectedWeekCards(cards) {
        return filterCardsForWeek(cards, selectedWeekStart);
    }

    function upsertCardInLoadedWeeks(cache, updatedCard, touchedAt = Date.now()) {
        if (!cache.currentWeekKey) {
            return;
        }

        const { weekStart, weekEndExclusive } = getWeekRange(cache.currentWeekKey);
        const overlapsWeek = cardOverlapsRange(
            updatedCard,
            weekStart.getTime(),
            weekEndExclusive.getTime(),
        );
        const cardIndex = cache.currentWeekCards.findIndex(card => card.id === updatedCard.id);

        if (!overlapsWeek) {
            if (cardIndex !== -1) {
                cache.currentWeekCards.splice(cardIndex, 1);
                cache.currentWeekFetchedAt = touchedAt;
            }
            return;
        }

        if (cardIndex === -1) {
            cache.currentWeekCards.push(updatedCard);
        } else {
            cache.currentWeekCards[cardIndex] = updatedCard;
        }

        cache.currentWeekCards = dedupeCardsById(cache.currentWeekCards);
        cache.currentWeekFetchedAt = touchedAt;
    }

    function removeCardFromLoadedWeeks(cache, cardId, touchedAt = Date.now()) {
        if (!cache.currentWeekCards.length) {
            return;
        }

        const cardIndex = cache.currentWeekCards.findIndex(card => card.id === cardId);
        if (cardIndex !== -1) {
            cache.currentWeekCards.splice(cardIndex, 1);
            cache.currentWeekFetchedAt = touchedAt;
        }
    }

    function renderSelectedWeekFromCache(cache) {
        const selectedWeekKey = getWeekCacheKey(selectedWeekStart);
        if (cache.currentWeekKey !== selectedWeekKey) {
            return;
        }

        applyTimeCards(cache.currentWeekCards);
    }

    function applyCardsLocally(nextCards, { activeCard = activeTimeCard, fetchedAt = Date.now() } = {}) {
        const visibleCards = getSelectedWeekCards(nextCards);
        if (selectedTeam) {
            const cache = getTeamCache(selectedTeam.id);
            setCachedWeek(cache, selectedWeekStart, visibleCards, fetchedAt);
            cache.activeCard = isOpenCardState(activeCard) ? activeCard : null;
        }

        applyTimeCards(visibleCards);
    }

    function applyUpdatedCardLocally(updatedCard) {
        if (!selectedTeam) {
            return;
        }

        const cache = getTeamCache(selectedTeam.id);
        upsertCardInLoadedWeeks(cache, updatedCard);

        if (isOpenCardState(updatedCard)) {
            cache.activeCard = updatedCard;
        } else if (activeTimeCard?.id === updatedCard.id) {
            cache.activeCard = null;
        }

        renderSelectedWeekFromCache(cache);
    }

    function removeCardLocally(cardId) {
        if (!selectedTeam || !cardId) {
            return;
        }

        const cache = getTeamCache(selectedTeam.id);
        removeCardFromLoadedWeeks(cache, cardId);
        if (cache.activeCard?.id === cardId) {
            cache.activeCard = null;
        }
        renderSelectedWeekFromCache(cache);
    }

    function requireTimeCardResponse(response, actionLabel) {
        if (!response?.id) {
            throw new Error(`${actionLabel} did not return an updated timecard.`);
        }

        return normalizeCard(response);
    }

    function resolveUpdatedTimeCard(response, fallbackCard, actionLabel) {
        if (response?.id) {
            return normalizeCard(response);
        }
        if (fallbackCard?.id) {
            return fallbackCard;
        }

        throw new Error(`${actionLabel} did not return an updated timecard.`);
    }

    async function fetchTimeCardById(teamId, cardId) {
        if (!teamId || !cardId) {
            throw new Error('A teamId and cardId are required to fetch a timecard.');
        }

        const response = await graphFetch(`/teams/${teamId}/schedule/timeCards/${cardId}`);
        return requireTimeCardResponse(response, 'Fetch timecard');
    }

    function storeFetchedWeekSnapshot(teamId, weekStartInput, snapshot) {
        const cache = getTeamCache(teamId);
        setCachedWeek(cache, weekStartInput, snapshot.cards, snapshot.fetchedAt);
        maybeWarnAboutPagedWeek(cache, snapshot.weekKey, snapshot.pageCount);
        cache.activeCard = snapshot.activeCard || null;
        return snapshot.activeCard || null;
    }

    async function resolveTimeCardActionResponse(response, {
        actionLabel,
        teamId,
        cardId = '',
        weekStartInput = null,
    } = {}) {
        if (response?.id) {
            return normalizeCard(response);
        }

        if (cardId) {
            return fetchTimeCardById(teamId, cardId);
        }

        if (weekStartInput !== null) {
            const snapshot = await fetchWeekTimeCards(teamId, weekStartInput);
            const activeCard = storeFetchedWeekSnapshot(teamId, weekStartInput, snapshot);
            if (activeCard?.id) {
                return activeCard;
            }
        }

        throw new Error(`${actionLabel} did not return an updated timecard.`);
    }

    function refreshTimeCardsInBackground(options = {}) {
        void refreshTimeCards(options).catch(error => {
            showError('Failed to refresh timecards: ' + error.message);
        });
    }

    function handleVisibilityChange() {
        if (document.hidden || !selectedTeam) {
            return;
        }
        refreshTimeCardsInBackground();
    }

    function stopBackgroundRefreshLoop() {
        if (backgroundRefreshTimerId !== null) {
            window.clearInterval(backgroundRefreshTimerId);
            backgroundRefreshTimerId = null;
        }
    }

    function syncBackgroundRefreshLoop() {
        if (document.body.dataset.page !== 'app' || !selectedTeam) {
            stopBackgroundRefreshLoop();
            return;
        }

        if (backgroundRefreshTimerId === null) {
            backgroundRefreshTimerId = window.setInterval(() => {
                if (document.hidden || !selectedTeam) {
                    return;
                }
                refreshTimeCardsInBackground();
            }, BACKGROUND_REFRESH_MS);
        }
    }

    function dedupeCardsById(cards) {
        const byId = new Map();
        cards.forEach(card => {
            byId.set(card.id, card);
        });
        return Array.from(byId.values());
    }

    function applyServerConfirmedCard(serverCard) {
        applyUpdatedCardLocally(serverCard);
        return true;
    }

    function normalizeCard(raw) {
        const card = {
            id: raw.id,
            state: raw.state,
            userId: raw.userId,
            createdDateTime: raw.createdDateTime,
            lastModifiedDateTime: raw.lastModifiedDateTime,
            clockIn: raw.clockInEvent ? {
                dateTime: raw.clockInEvent.dateTime,
                atApprovedLocation: raw.clockInEvent.atApprovedLocation,
                notes: raw.clockInEvent.notes,
            } : null,
            clockOut: raw.clockOutEvent ? {
                dateTime: raw.clockOutEvent.dateTime,
                atApprovedLocation: raw.clockOutEvent.atApprovedLocation,
                notes: raw.clockOutEvent.notes,
            } : null,
            breaks: (raw.breaks || []).map(b => ({
                breakId: b.breakId,
                start: b.start ? { dateTime: b.start.dateTime, notes: b.start.notes } : null,
                end: b.end ? { dateTime: b.end.dateTime, notes: b.end.notes } : null,
            })),
            notes: raw.notes || null,
        };

        card.state = deriveCardState(card);
        return card;
    }

    function getCardAnchorDateTime(card) {
        return card?.clockIn?.dateTime || card?.clockInEvent?.dateTime || card?.createdDateTime || null;
    }

    // ─────────────────────────────────────────────
    // Render weeks + cards
    // ─────────────────────────────────────────────
    function renderWeeks() {
        tcLoading.style.display = 'none';
        weeksContainer.innerHTML = '';
        weeksContainer.style.display = 'block';
        ensureSelectedWeek();
        syncWeekControls();

        const group = getSelectedWeekGroup(allTimeCards);
        latestVisibleCardId = getLatestVisibleCardId(group.cards);
        updateCurrentWeekTotal(group);
        group.cards.sort((a, b) => getCardSortMs(a) - getCardSortMs(b));
        renderDaySections(group, weeksContainer);

        // Attach interact.js dragging to all .tc-block elements
        attachTimelineInteractions();
        syncHighlightedCardVisuals();
        syncLiveClockUpdates();
    }

    function setHighlightedCardId(cardId) {
        highlightedCardId = cardId || null;
        syncHighlightedCardVisuals();
    }

    function syncHighlightedCardVisuals() {
        document.querySelectorAll('.tc-row, .tc-block, .break-overlay').forEach(element => {
            element.classList.toggle('link-highlighted', Boolean(highlightedCardId) && element.dataset.cardId === highlightedCardId);
        });
    }

    function getLatestVisibleCardId(cards) {
        let latestCard = null;
        cards.forEach(card => {
            if (!latestCard || getCardSortMs(card) > getCardSortMs(latestCard)) {
                latestCard = card;
            }
        });
        return latestCard?.id || null;
    }

    function getSelectedWeekGroup(cards) {
        const weekStart = getResolvedWeekStart(selectedWeekStart);
        return {
            weekStart,
            cards: filterCardsForWeek(cards, weekStart),
        };
    }

    function getWeekStart(date) {
        const d = new Date(date);
        d.setHours(0, 0, 0, 0);
        const day = d.getDay(); // 0=Sun
        d.setDate(d.getDate() - day);
        return d;
    }

    function getDayStart(date) {
        const d = new Date(date);
        d.setHours(0, 0, 0, 0);
        return d;
    }

    function groupByDay(cards) {
        const map = new Map();
        cards.forEach(card => {
            const anchorDateTime = getCardAnchorDateTime(card);
            if (!anchorDateTime) return;
            const dayStart = getDayStart(new Date(anchorDateTime));
            const key = dayStart.getTime();
            if (!map.has(key)) map.set(key, { dayStart, cards: [] });
            map.get(key).cards.push(card);
        });
        return Array.from(map.values());
    }

    function renderDaySections(group, containerEl) {
        const dayGroups = groupByDay(group.cards)
            .sort((a, b) => b.dayStart - a.dayStart);

        if (!dayGroups.length) {
            containerEl.innerHTML = '<div class="panel-empty">No timecards for this Sunday-Saturday week.</div>';
            return;
        }

        dayGroups.forEach(dayGroup => {
            containerEl.appendChild(buildDaySection(dayGroup));
        });
    }

    function buildDaySection(dayGroup) {
        const totalMs = dayGroup.cards.reduce((sum, card) => sum + workedMs(card), 0);
        const section = document.createElement('section');
        section.className = 'day-section';
        section.dataset.dayStart = dayGroup.dayStart.getTime();
        section.innerHTML = `
        <div class="day-section-header">
            <div class="day-section-label">
                <span>${fmtDayShort(dayGroup.dayStart)}</span>
                <strong>${fmtDate(dayGroup.dayStart)}</strong>
            </div>
            <span class="day-section-total">${formatDurationHms(totalMs)} worked</span>
        </div>
        <div class="day-timeline-row">
            <div class="day-timeline-wrapper">
                <div class="day-timeline-container">
                    <div class="timeline-axis day-timeline-axis"></div>
                </div>
            </div>
        </div>
        <div class="card-list day-card-list"></div>
    `;

        renderDayTimeline(dayGroup, section);
        renderDayCardList(dayGroup.cards, section.querySelector('.day-card-list'));
        return section;
    }

    // ─────────────────────────────────────────────
    // Timeline rendering
    // ─────────────────────────────────────────────
    function renderTimeline(group, blockEl) {
        const axisEl = blockEl.querySelector(`#axis-${group.weekStart.getTime()}`);
        if (!axisEl) return;
        axisEl.innerHTML = '';

        // 7 days, starting at midnight Monday
        const weekStartMs = group.weekStart.getTime();
        const weekEndMs = weekStartMs + 7 * 24 * 3600 * 1000;
        const totalSpanMs = weekEndMs - weekStartMs;

        // Tick every 2 hours, day dividers daily
        const tickInterval = 2 * 3600 * 1000;
        for (let t = weekStartMs; t <= weekEndMs; t += tickInterval) {
            const pct = ((t - weekStartMs) / totalSpanMs) * 100;
            const d = new Date(t);
            const isDayBoundary = d.getHours() === 0 && d.getMinutes() === 0;

            if (isDayBoundary) {
                const divider = document.createElement('div');
                divider.className = 'day-divider';
                divider.style.left = pct + '%';
                axisEl.appendChild(divider);

                const dl = document.createElement('div');
                dl.className = 'day-label';
                dl.style.left = pct + '%';
                dl.textContent = fmtDayShort(d);
                axisEl.appendChild(dl);
            } else {
                const tick = document.createElement('div');
                tick.className = 'hour-tick';
                tick.style.left = pct + '%';
                axisEl.appendChild(tick);

                const hl = document.createElement('div');
                hl.className = 'hour-label';
                hl.style.left = pct + '%';
                hl.textContent = fmtHour(d);
                axisEl.appendChild(hl);
            }
        }

        appendTimelineBlocks(axisEl, group.cards, weekStartMs, weekEndMs, totalSpanMs);
    }

    function renderDayTimeline(dayGroup, sectionEl) {
        const axisEl = sectionEl.querySelector('.day-timeline-axis');
        if (!axisEl) return;
        axisEl.innerHTML = '';

        const dayStartMs = dayGroup.dayStart.getTime();
        const dayEndMs = dayStartMs + 24 * 3600 * 1000;
        const totalSpanMs = dayEndMs - dayStartMs;
        const tickInterval = 3600 * 1000;

        for (let t = dayStartMs; t <= dayEndMs; t += tickInterval) {
            const pct = ((t - dayStartMs) / totalSpanMs) * 100;
            const d = new Date(t);

            const tick = document.createElement('div');
            tick.className = 'hour-tick';
            tick.style.left = pct + '%';
            axisEl.appendChild(tick);

            if (d.getHours() % 2 === 0) {
                const hl = document.createElement('div');
                hl.className = 'hour-label';
                hl.style.left = pct + '%';
                hl.textContent = fmtHour(d);
                axisEl.appendChild(hl);
            }
        }

        appendTimelineBlocks(axisEl, dayGroup.cards, dayStartMs, dayEndMs, totalSpanMs);
    }

    function appendTimelineBlocks(axisEl, cards, axisStartMs, axisEndMs, totalSpanMs) {
        cards.forEach(card => {
            if (!card.clockIn) return;

            const startMs = new Date(card.clockIn.dateTime).getTime();
            const endMs = card.clockOut ? new Date(card.clockOut.dateTime).getTime() : Date.now();
            if (endMs <= axisStartMs || startMs >= axisEndMs) return;

            const clippedStartMs = Math.max(startMs, axisStartMs);
            const clippedEndMs = Math.min(endMs, axisEndMs);
            const leftPct = ((clippedStartMs - axisStartMs) / totalSpanMs) * 100;
            const widthPct = ((clippedEndMs - clippedStartMs) / totalSpanMs) * 100;

            const tcBlock = document.createElement('div');
            tcBlock.className = 'tc-block' + (card.id === activeTimeCard?.id ? ' active-card-block' : '');
            tcBlock.style.left = `${Math.max(0, leftPct)}%`;
            tcBlock.style.width = `${Math.max(0.1, widthPct)}%`;
            tcBlock.dataset.cardId = card.id;
            tcBlock.dataset.startMs = startMs;
            tcBlock.dataset.endMs = card.clockOut ? new Date(card.clockOut.dateTime).getTime() : '';
            tcBlock.dataset.weekStartMs = axisStartMs;
            tcBlock.dataset.totalSpanMs = totalSpanMs;
            tcBlock.title = `${fmtDateTime(new Date(startMs))}${card.clockOut ? ' → ' + fmtDateTime(new Date(card.clockOut.dateTime)) : ' (active)'}`;

            const label = document.createElement('div');
            label.className = 'tc-label';
            label.textContent = fmtTime(new Date(clippedStartMs));
            tcBlock.appendChild(label);

            const leftHandle = document.createElement('div');
            leftHandle.className = 'resize-handle';
            leftHandle.dataset.edge = 'left';
            tcBlock.appendChild(leftHandle);

            const rightHandle = document.createElement('div');
            rightHandle.className = 'resize-handle';
            rightHandle.dataset.edge = 'right';
            tcBlock.appendChild(rightHandle);

            tcBlock.addEventListener('pointerdown', () => {
                setHighlightedCardId(card.id);
            });

            axisEl.appendChild(tcBlock);

            card.breaks.forEach((b, breakIndex) => {
                if (!b.start) return;
                const bStart = new Date(b.start.dateTime).getTime();
                const bEnd = b.end ? new Date(b.end.dateTime).getTime() : Date.now();
                if (bEnd <= axisStartMs || bStart >= axisEndMs) return;

                const clippedBreakStartMs = Math.max(bStart, axisStartMs);
                const clippedBreakEndMs = Math.min(bEnd, axisEndMs);
                const bLeft = ((clippedBreakStartMs - axisStartMs) / totalSpanMs) * 100;
                const bWidth = ((clippedBreakEndMs - clippedBreakStartMs) / totalSpanMs) * 100;

                const overlay = document.createElement('div');
                overlay.className = 'break-overlay';
                overlay.style.left = `${Math.max(0, bLeft)}%`;
                overlay.style.width = `${Math.max(0.05, bWidth)}%`;
                overlay.dataset.cardId = card.id;
                overlay.dataset.breakIndex = String(breakIndex);
                overlay.dataset.breakId = b.breakId || '';
                overlay.dataset.weekStartMs = axisStartMs;
                overlay.dataset.totalSpanMs = totalSpanMs;
                overlay.dataset.cardStartMs = startMs;
                overlay.dataset.cardEndMs = endMs;
                overlay.dataset.startMs = bStart;
                overlay.dataset.endMs = b.end ? String(bEnd) : '';
                overlay.title = `${fmtDateTime(new Date(bStart))}${b.end ? ' → ' + fmtDateTime(new Date(bEnd)) : ' (open break)'}`;

                const leftHandle = document.createElement('div');
                leftHandle.className = 'resize-handle';
                leftHandle.dataset.edge = 'left';
                overlay.appendChild(leftHandle);

                const rightHandle = document.createElement('div');
                rightHandle.className = 'resize-handle';
                rightHandle.dataset.edge = 'right';
                overlay.appendChild(rightHandle);

                overlay.addEventListener('pointerdown', () => {
                    setHighlightedCardId(card.id);
                });

                axisEl.appendChild(overlay);
            });
        });
    }

    // ─────────────────────────────────────────────
    // interact.js timeline drag / resize
    // ─────────────────────────────────────────────
    function attachTimelineInteractions() {
        if (typeof interact === 'undefined') return;

        interact('.tc-block').unset(); // clean up previous bindings
        interact('.break-overlay').unset();

        interact('.tc-block')
            .draggable({
                axis: 'x',
                ignoreFrom: '.resize-handle, [data-end-ms=""]',
                listeners: {
                    start: markInteractionActive,
                    move: onTimelineDragMove,
                    end: onTimelineDragEnd,
                },
                inertia: false,
            })
            .resizable({
                edges: {
                    left: '.resize-handle[data-edge="left"]',
                    right: '.resize-handle[data-edge="right"]',
                },
                axis: 'x',
                listeners: {
                    start: markInteractionActive,
                    move: onTimelineResizeMove,
                    end: onTimelineResizeEnd,
                },
                inertia: false,
            });

        interact('.break-overlay')
            .draggable({
                axis: 'x',
                listeners: {
                    start: markInteractionActive,
                    move: onBreakDragMove,
                    end: onBreakDragEnd,
                },
                inertia: false,
            })
            .resizable({
                edges: {
                    left: '.resize-handle[data-edge="left"]',
                    right: '.resize-handle[data-edge="right"]',
                },
                axis: 'x',
                listeners: {
                    start: markInteractionActive,
                    move: onBreakResizeMove,
                    end: onBreakResizeEnd,
                },
                inertia: false,
            });
    }

    function markInteractionActive(event) {
        event.target.dataset.isInteracting = 'true';
        if (event.edges?.left) {
            event.target.dataset.resizeEdge = 'left';
        } else if (event.edges?.right) {
            event.target.dataset.resizeEdge = 'right';
        } else {
            delete event.target.dataset.resizeEdge;
        }
    }

    function clearInteractionActive(el) {
        delete el.dataset.isInteracting;
        delete el.dataset.resizeEdge;
    }

    function isOpenTimelineBlock(el) {
        return el.classList.contains('tc-block') && !el.dataset.endMs;
    }

    function getActiveResizeEdge(el, event) {
        if (el.dataset.resizeEdge) {
            return el.dataset.resizeEdge;
        }
        if (event?.edges?.left) {
            return 'left';
        }
        if (event?.edges?.right) {
            return 'right';
        }
        return '';
    }

    function getAxisWidth(tcBlock) {
        const axis = tcBlock.closest('.timeline-axis');
        return axis ? axis.offsetWidth : 1;
    }

    function msFromPct(pct, totalSpanMs) {
        return (pct / 100) * totalSpanMs;
    }

    function snapMs(ms) {
        const snap = SNAP_MINUTES * 60 * 1000;
        return Math.round(ms / snap) * snap;
    }

    function pctFromMs(ms, axisStartMs, totalSpanMs) {
        return ((ms - axisStartMs) / totalSpanMs) * 100;
    }

    function getRangeBoundsPct(el) {
        const axisStartMs = Number(el.dataset.weekStartMs);
        const totalSpanMs = Number(el.dataset.totalSpanMs);
        const axisEndMs = axisStartMs + totalSpanMs;
        const rangeStartMs = Number(el.dataset.cardStartMs);
        const rangeEndMs = Number(el.dataset.cardEndMs);

        const minLeftPct = pctFromMs(Math.max(axisStartMs, rangeStartMs), axisStartMs, totalSpanMs);
        const maxRightPct = pctFromMs(Math.min(axisEndMs, rangeEndMs), axisStartMs, totalSpanMs);

        return {
            minLeftPct: Math.max(0, minLeftPct),
            maxRightPct: Math.min(100, maxRightPct),
        };
    }

    function getOverlayStartEndMs(el) {
        const weekStartMs = Number(el.dataset.weekStartMs);
        const totalSpanMs = Number(el.dataset.totalSpanMs);
        const leftPct = parseFloat(el.style.left);
        const widthPct = parseFloat(el.style.width);

        return {
            startMs: snapMs(weekStartMs + msFromPct(leftPct, totalSpanMs)),
            endMs: snapMs(weekStartMs + msFromPct(leftPct + widthPct, totalSpanMs)),
        };
    }

    function onTimelineDragMove(event) {
        const el = event.target;
        if (isOpenTimelineBlock(el)) {
            return;
        }

        const axisW = getAxisWidth(el);
        const totalSpanMs = parseInt(el.dataset.totalSpanMs, 10);
        const weekStartMs = parseInt(el.dataset.weekStartMs, 10);

        const deltaPct = (event.dx / axisW) * 100;
        const currentLeft = parseFloat(el.style.left);
        const newLeft = Math.max(0, Math.min(currentLeft + deltaPct, 100 - parseFloat(el.style.width)));
        el.style.left = newLeft + '%';

        // Compute preview times
        const newStartMs = weekStartMs + msFromPct(newLeft, totalSpanMs);
        const width = parseFloat(el.style.width);
        const newEndMs = newStartMs + msFromPct(width, totalSpanMs);

        showDragTooltip(event.clientX, event.clientY,
            `${fmtTime(new Date(newStartMs))} → ${fmtTime(new Date(newEndMs))}`);

        el.dataset.pendingStartMs = snapMs(newStartMs);
        el.dataset.pendingEndMs = snapMs(newEndMs);
    }

    async function onTimelineDragEnd(event) {
        hideDragTooltip();
        const el = event.target;
        try {
            if (isOpenTimelineBlock(el)) {
                return;
            }

            const cardId = el.dataset.cardId;
            const pendingStart = el.dataset.pendingStartMs;
            const pendingEnd = el.dataset.pendingEndMs;
            if (!pendingStart) return;

            const card = allTimeCards.find(c => c.id === cardId);
            if (!card) return;

            const newStart = new Date(parseInt(pendingStart, 10));
            const newEnd = pendingEnd ? new Date(parseInt(pendingEnd, 10)) : null;

            if (newEnd && newEnd <= newStart) {
                toast('End time must be after start time', 'error');
                reRenderWeeks();
                return;
            }

            await persistCardTimeUpdate(card, newStart, newEnd);
        } finally {
            clearInteractionActive(el);
        }
    }

    function onTimelineResizeMove(event) {
        const el = event.target;
        const isOpenCard = isOpenTimelineBlock(el);
        const activeResizeEdge = getActiveResizeEdge(el, event);
        if (isOpenCard && activeResizeEdge === 'right') {
            return;
        }

        const axisW = getAxisWidth(el);
        const totalSpanMs = parseInt(el.dataset.totalSpanMs, 10);
        const weekStartMs = parseInt(el.dataset.weekStartMs, 10);

        const deltaWPct = (event.deltaRect.width / axisW) * 100;
        const deltaLPct = (event.deltaRect.left / axisW) * 100;

        let newLeft = parseFloat(el.style.left) + deltaLPct;
        let newWidth = parseFloat(el.style.width) + deltaWPct;
        newLeft = Math.max(0, newLeft);
        newWidth = Math.max(0.1, newWidth);

        el.style.left = newLeft + '%';
        el.style.width = newWidth + '%';

        const newStartMs = weekStartMs + msFromPct(newLeft, totalSpanMs);
        const newEndMs = newStartMs + msFromPct(newWidth, totalSpanMs);

        el.dataset.pendingStartMs = snapMs(newStartMs);
        if (isOpenCard) {
            delete el.dataset.pendingEndMs;
            showDragTooltip(event.clientX, event.clientY,
                `${fmtTime(new Date(newStartMs))} → active`);
            return;
        }

        showDragTooltip(event.clientX, event.clientY,
            `${fmtTime(new Date(newStartMs))} → ${fmtTime(new Date(newEndMs))}`);

        el.dataset.pendingEndMs = snapMs(newEndMs);
    }

    async function onTimelineResizeEnd(event) {
        hideDragTooltip();
        const el = event.target;
        try {
            const isOpenCard = isOpenTimelineBlock(el);
            if (isOpenCard && getActiveResizeEdge(el, event) === 'right') {
                return;
            }

            const cardId = el.dataset.cardId;
            const pendingStart = el.dataset.pendingStartMs;
            const pendingEnd = el.dataset.pendingEndMs;
            if (!pendingStart) return;

            const card = allTimeCards.find(c => c.id === cardId);
            if (!card) return;

            const newStart = new Date(parseInt(pendingStart, 10));
            const newEnd = isOpenCard || !pendingEnd ? null : new Date(parseInt(pendingEnd, 10));

            if (newEnd && newEnd <= newStart) {
                toast('End time must be after start time', 'error');
                reRenderWeeks();
                return;
            }

            await persistCardTimeUpdate(card, newStart, newEnd);
        } finally {
            clearInteractionActive(el);
        }
    }

    function onBreakDragMove(event) {
        const el = event.target;
        const axisW = getAxisWidth(el);
        const deltaPct = (event.dx / axisW) * 100;
        const widthPct = parseFloat(el.style.width);
        const bounds = getRangeBoundsPct(el);
        const currentLeft = parseFloat(el.style.left);
        const newLeft = clamp(currentLeft + deltaPct, bounds.minLeftPct, bounds.maxRightPct - widthPct);
        const totalSpanMs = Number(el.dataset.totalSpanMs);
        const weekStartMs = Number(el.dataset.weekStartMs);
        const newStartMs = weekStartMs + msFromPct(newLeft, totalSpanMs);
        const newEndMs = newStartMs + msFromPct(widthPct, totalSpanMs);

        el.style.left = `${newLeft}%`;
        showDragTooltip(event.clientX, event.clientY,
            `${fmtTime(new Date(newStartMs))} → ${fmtTime(new Date(newEndMs))}`);

        const snapped = getOverlayStartEndMs(el);
        el.dataset.pendingStartMs = String(snapped.startMs);
        el.dataset.pendingEndMs = String(snapped.endMs);
    }

    async function onBreakDragEnd(event) {
        hideDragTooltip();
        const el = event.target;
        try {
            const card = allTimeCards.find(item => item.id === el.dataset.cardId);
            const breakIndex = Number(el.dataset.breakIndex);
            const pendingStart = Number(el.dataset.pendingStartMs);
            const pendingEnd = Number(el.dataset.pendingEndMs);

            if (!card || !Number.isInteger(breakIndex) || !Number.isFinite(pendingStart) || !Number.isFinite(pendingEnd)) {
                reRenderWeeks();
                return;
            }

            const newStart = new Date(pendingStart);
            const newEnd = new Date(pendingEnd);
            if (newEnd <= newStart) {
                toast('Break end must be after break start', 'error');
                reRenderWeeks();
                return;
            }

            await persistBreakTimeUpdate(card, breakIndex, newStart, newEnd);
        } finally {
            clearInteractionActive(el);
        }
    }

    function onBreakResizeMove(event) {
        const el = event.target;
        const axisW = getAxisWidth(el);
        const deltaWPct = (event.deltaRect.width / axisW) * 100;
        const deltaLPct = (event.deltaRect.left / axisW) * 100;
        const bounds = getRangeBoundsPct(el);

        let newLeft = parseFloat(el.style.left) + deltaLPct;
        let newWidth = parseFloat(el.style.width) + deltaWPct;
        newWidth = Math.max(0.05, newWidth);

        if (newLeft < bounds.minLeftPct) {
            newWidth -= (bounds.minLeftPct - newLeft);
            newLeft = bounds.minLeftPct;
        }

        if (newLeft + newWidth > bounds.maxRightPct) {
            newWidth = bounds.maxRightPct - newLeft;
        }

        newWidth = Math.max(0.05, newWidth);
        if (newLeft + newWidth > bounds.maxRightPct) {
            newLeft = Math.max(bounds.minLeftPct, bounds.maxRightPct - newWidth);
        }

        const totalSpanMs = Number(el.dataset.totalSpanMs);
        const weekStartMs = Number(el.dataset.weekStartMs);
        const newStartMs = weekStartMs + msFromPct(newLeft, totalSpanMs);
        const newEndMs = newStartMs + msFromPct(newWidth, totalSpanMs);

        el.style.left = `${newLeft}%`;
        el.style.width = `${newWidth}%`;
        showDragTooltip(event.clientX, event.clientY,
            `${fmtTime(new Date(newStartMs))} → ${fmtTime(new Date(newEndMs))}`);

        const snapped = getOverlayStartEndMs(el);
        el.dataset.pendingStartMs = String(snapped.startMs);
        el.dataset.pendingEndMs = String(snapped.endMs);
    }

    async function onBreakResizeEnd(event) {
        hideDragTooltip();
        const el = event.target;
        try {
            const card = allTimeCards.find(item => item.id === el.dataset.cardId);
            const breakIndex = Number(el.dataset.breakIndex);
            const pendingStart = Number(el.dataset.pendingStartMs);
            const pendingEnd = Number(el.dataset.pendingEndMs);

            if (!card || !Number.isInteger(breakIndex) || !Number.isFinite(pendingStart) || !Number.isFinite(pendingEnd)) {
                reRenderWeeks();
                return;
            }

            const newStart = new Date(pendingStart);
            const newEnd = new Date(pendingEnd);
            if (newEnd <= newStart) {
                toast('Break end must be after break start', 'error');
                reRenderWeeks();
                return;
            }

            await persistBreakTimeUpdate(card, breakIndex, newStart, newEnd);
        } finally {
            clearInteractionActive(el);
        }
    }

    function showDragTooltip(x, y, text) {
        dragTooltip.style.display = 'block';
        dragTooltip.style.left = (x + 12) + 'px';
        dragTooltip.style.top = (y - 28) + 'px';
        dragTooltip.textContent = text;
    }
    function hideDragTooltip() {
        dragTooltip.style.display = 'none';
    }

    // ─────────────────────────────────────────────
    // Card list rendering
    // ─────────────────────────────────────────────
    function renderDayCardList(cards, listEl) {
        if (!listEl) return;
        listEl.innerHTML = '';

        [...cards]
            .sort((a, b) => getCardSortMs(b) - getCardSortMs(a))
            .forEach(card => {
                const stack = document.createElement('div');
                stack.className = 'tc-card-stack';
                const row = buildCardRow(card);
                stack.appendChild(row);

                // breaks sub-list
                if (card.breaks.length) {
                    const breaksList = buildBreaksList(card);
                    stack.appendChild(breaksList);
                }
                listEl.appendChild(stack);
            });
    }

    function buildCardRow(card) {
        const row = document.createElement('div');
        row.className = 'tc-row'
            + (card.id === activeTimeCard?.id ? ' active-card' : '')
            + (card.id === highlightedCardId ? ' link-highlighted' : '');
        row.dataset.cardId = card.id;

        const inDt = card.clockIn ? fmtDateTime(new Date(card.clockIn.dateTime)) : '—';
        const outDt = card.clockOut ? fmtDateTime(new Date(card.clockOut.dateTime)) : '—';
        const workedHms = formatDurationHms(workedMs(card));
        const showExtendToNow = card.id === latestVisibleCardId && Boolean(card.clockIn && card.clockOut);

        row.innerHTML = `
    <div class="tc-row-date">${card.clockIn ? fmtDate(new Date(card.clockIn.dateTime)) : '—'}</div>
        <div class="tc-row-time" data-field="clock-in"><span class="tc-row-time-value">${inDt}</span></div>
            <div class="tc-row-time" data-field="clock-out"><span class="tc-row-time-value">${outDt}</span><span class="tc-row-total">(${workedHms})</span></div>
    <div class="tc-row-actions">
            ${showExtendToNow ? `<button class="btn btn-secondary" data-action="extend-now" data-card-id="${card.id}">Extend to now</button>` : ''}
      <button class="btn btn-secondary" data-action="edit" data-card-id="${card.id}">Edit</button>
      <button class="btn btn-danger" data-action="delete" data-card-id="${card.id}">Delete</button>
    </div>
  `;

        row.addEventListener('click', () => {
            setHighlightedCardId(card.id);
        });

        row.querySelectorAll('[data-action]').forEach(btn => {
            btn.addEventListener('click', e => {
                e.stopPropagation();
                const action = btn.dataset.action;
                const cid = btn.dataset.cardId;
                const c = allTimeCards.find(x => x.id === cid);
                if (!c) return;
                if (action === 'extend-now') extendCardToNow(c);
                if (action === 'edit') openEditModal(c);
                if (action === 'delete') confirmDeleteCard(c);
            });
        });

        return row;
    }

    function buildBreaksList(card) {
        const wrapper = document.createElement('div');
        wrapper.className = 'breaks-list';
        card.breaks.forEach((b, idx) => {
            const bStart = b.start ? fmtDateTime(new Date(b.start.dateTime)) : '—';
            const bEnd = b.end ? fmtDateTime(new Date(b.end.dateTime)) : 'ongoing';
            const dur = b.start && b.end
                ? `${((new Date(b.end.dateTime) - new Date(b.start.dateTime)) / 60000).toFixed(0)}m`
                : '';
            const brow = document.createElement('div');
            brow.className = 'break-row';
            brow.innerHTML = `
            <span class="break-label">↳ Break ${idx + 1}:</span>
      <span class="break-row-detail">${bStart} → ${bEnd} ${dur ? `(${dur})` : ''}</span>
      ${b.breakId ? `<button class="btn btn-danger btn-compact" data-action="delete-break" data-card-id="${card.id}" data-break-id="${b.breakId}">Delete</button>` : ''}
    `;

            const deleteBtn = brow.querySelector('[data-action="delete-break"]');
            if (deleteBtn) {
                deleteBtn.addEventListener('click', event => {
                    event.stopPropagation();
                    confirmDeleteBreak(card.id, b.breakId);
                });
            }

            wrapper.appendChild(brow);
        });
        return wrapper;
    }

    // ─────────────────────────────────────────────
    // Toolbar state
    // ─────────────────────────────────────────────
    function updateToolbarState() {
        const state = activeTimeCard?.state || 'clockedOut';
        const showClockIn = state === 'clockedOut';
        const showClockOut = state === 'clockedIn' || state === 'onBreak';
        const showStartBreak = state === 'clockedIn';
        const showEndBreak = state === 'onBreak';

        btnClockIn.disabled = false;
        btnClockOut.disabled = false;
        btnStartBreak.disabled = false;
        btnEndBreak.disabled = false;

        btnClockIn.style.display = showClockIn ? 'inline-flex' : 'none';
        btnClockOut.style.display = showClockOut ? 'inline-flex' : 'none';
        btnStartBreak.style.display = showStartBreak ? 'inline-flex' : 'none';
        btnEndBreak.style.display = showEndBreak ? 'inline-flex' : 'none';

        toolbarActions.style.display = selectedTeam && (showClockIn || showClockOut || showStartBreak || showEndBreak)
            ? 'flex'
            : 'none';
        stateBadge.style.display = selectedTeam ? 'inline-block' : 'none';

        stateBadge.textContent = state === 'clockedIn' ? '● Clocked In'
            : state === 'onBreak' ? '◐ On Break'
                : '○ Clocked Out';
        stateBadge.className = state === 'clockedIn' ? 'chip chip-in'
            : state === 'onBreak' ? 'chip chip-break'
                : 'chip chip-out';
    }

    // ─────────────────────────────────────────────
    // Timecard actions
    // ─────────────────────────────────────────────
    async function doClockIn() {
        if (!selectedTeam) return;
        btnClockIn.disabled = true;
        const teamId = selectedTeam.id;
        try {
            const response = await graphFetch(`/teams/${teamId}/schedule/timeCards/clockIn`, {
                method: 'POST',
                body: JSON.stringify({ atApprovedLocation: false, notes: { contentType: 'text', content: '' } }),
            });
            const updatedCard = await resolveTimeCardActionResponse(response, {
                actionLabel: 'Clock in',
                teamId,
                weekStartInput: getResolvedWeekStart(new Date()),
            });
            applyServerConfirmedCard(updatedCard);
            toast('Clocked in', 'success');
        } catch (e) {
            toast('Clock in failed: ' + e.message, 'error');
            btnClockIn.disabled = false;
        }
    }

    async function doClockOut() {
        if (!activeTimeCard) return;
        btnClockOut.disabled = true;
        const teamId = selectedTeam.id;
        const targetCard = activeTimeCard;
        try {
            const response = await graphFetch(`/teams/${teamId}/schedule/timeCards/${targetCard.id}/clockOut`, {
                method: 'POST',
                body: JSON.stringify({ atApprovedLocation: false, notes: { contentType: 'text', content: '' } }),
            });
            const updatedCard = await resolveTimeCardActionResponse(response, {
                actionLabel: 'Clock out',
                teamId,
                cardId: targetCard.id,
            });
            applyServerConfirmedCard(updatedCard);
            toast('Clocked out', 'success');
        } catch (e) {
            toast('Clock out failed: ' + e.message, 'error');
            btnClockOut.disabled = false;
        }
    }

    async function doStartBreak() {
        if (!activeTimeCard) return;
        btnStartBreak.disabled = true;
        const teamId = selectedTeam.id;
        const targetCard = activeTimeCard;
        try {
            const response = await graphFetch(`/teams/${teamId}/schedule/timeCards/${targetCard.id}/startBreak`, {
                method: 'POST',
                body: JSON.stringify({ atApprovedLocation: false, notes: { contentType: 'text', content: '' } }),
            });
            const updatedCard = await resolveTimeCardActionResponse(response, {
                actionLabel: 'Start break',
                teamId,
                cardId: targetCard.id,
            });
            applyServerConfirmedCard(updatedCard);
            toast('Break started', 'success');
        } catch (e) {
            toast('Start break failed: ' + e.message, 'error');
            btnStartBreak.disabled = false;
        }
    }

    async function doEndBreak() {
        if (!activeTimeCard) return;
        btnEndBreak.disabled = true;
        const teamId = selectedTeam.id;
        const targetCard = activeTimeCard;
        try {
            const response = await graphFetch(`/teams/${teamId}/schedule/timeCards/${targetCard.id}/endBreak`, {
                method: 'POST',
                body: JSON.stringify({ atApprovedLocation: false, notes: { contentType: 'text', content: '' } }),
            });
            const updatedCard = await resolveTimeCardActionResponse(response, {
                actionLabel: 'End break',
                teamId,
                cardId: targetCard.id,
            });
            applyServerConfirmedCard(updatedCard);
            toast('Break ended', 'success');
        } catch (e) {
            toast('End break failed: ' + e.message, 'error');
            btnEndBreak.disabled = false;
        }
    }

    // ─────────────────────────────────────────────
    // Delete timecard
    // ─────────────────────────────────────────────
    function confirmDeleteCard(card) {
        const inStr = card.clockIn ? fmtDateTime(new Date(card.clockIn.dateTime)) : 'unknown';
        $('confirm-title').textContent = 'Delete Timecard';
        $('confirm-msg').textContent = `Are you sure you want to permanently delete the timecard starting ${inStr}? This cannot be undone.`;
        $('confirm-beta-warn').style.display = 'block';
        $('confirm-modal').classList.remove('hidden');
        confirmCallback = async (ok) => {
            if (!ok) return;
            await deleteCard(card);
        };
    }

    async function deleteCard(card) {
        try {
            await graphFetchBeta(`/teams/${selectedTeam.id}/schedule/timeCards/${card.id}`, {
                method: 'DELETE',
            });
            removeCardLocally(card.id);
            toast('Timecard deleted', 'success');
        } catch (e) {
            toast('Delete failed: ' + e.message, 'error');
        }
    }

    function confirmDeleteBreak(cardId, breakId) {
        const card = allTimeCards.find(item => item.id === cardId);
        const currentBreak = card?.breaks.find(item => item.breakId === breakId);
        if (!card || !currentBreak || !breakId) {
            return;
        }

        const breakStart = currentBreak.start?.dateTime ? fmtDateTime(new Date(currentBreak.start.dateTime)) : 'this break';
        $('confirm-title').textContent = 'Delete Break';
        $('confirm-msg').textContent = `Delete the break starting ${breakStart}? This cannot be undone.`;
        $('confirm-beta-warn').style.display = 'none';
        $('confirm-modal').classList.remove('hidden');
        confirmCallback = async (ok) => {
            if (!ok) return;
            await deleteBreak(card, breakId);
        };
    }

    function getUpdatedCardWithoutBreak(card, breakId) {
        const updatedCard = {
            ...card,
            clockIn: card.clockIn ? {
                ...card.clockIn,
            } : null,
            clockOut: card.clockOut ? {
                ...card.clockOut,
            } : null,
            breaks: cloneCardBreaks(card.breaks).filter(item => item.breakId !== breakId),
            lastModifiedDateTime: new Date().toISOString(),
        };

        updatedCard.state = deriveCardState(updatedCard);
        return updatedCard;
    }

    async function deleteBreak(card, breakId) {
        if (!selectedTeam || !breakId) {
            return;
        }

        try {
            const updatedCard = getUpdatedCardWithoutBreak(card, breakId);
            const response = await graphFetch(`/teams/${selectedTeam.id}/schedule/timeCards/${card.id}`, {
                method: 'PUT',
                body: JSON.stringify(buildTimeCardUpdateBody(updatedCard)),
            });

            applyServerConfirmedCard(resolveUpdatedTimeCard(response, updatedCard, 'Delete break'));

            toast('Break deleted', 'success');
        } catch (e) {
            toast('Delete break failed: ' + e.message, 'error');
        }
    }

    // ─────────────────────────────────────────────
    // Edit timecard modal
    // ─────────────────────────────────────────────
    function openEditModal(card) {
        editTarget = card;
        $('edit-clock-in').value = card.clockIn ? toLocalDatetimeInput(new Date(card.clockIn.dateTime)) : '';
        $('edit-clock-out').value = card.clockOut ? toLocalDatetimeInput(new Date(card.clockOut.dateTime)) : '';
        $('edit-notes').value = card.notes?.content || '';

        // Populate breaks
        const container = $('breaks-edit-container');
        container.innerHTML = '';
        card.breaks.forEach(b => addBreakRow(null, b));

        $('edit-modal').classList.remove('hidden');
    }

    function closeEditModal() {
        editTarget = null;
        $('edit-modal').classList.add('hidden');
    }

    function addBreakRow(e, existingBreak = null) {
        const container = $('breaks-edit-container');
        const idx = container.children.length;
        const breakId = existingBreak?.breakId || '';
        const defaultBreak = existingBreak
            ? {
                startVal: existingBreak.start ? toLocalDatetimeInput(new Date(existingBreak.start.dateTime)) : '',
                endVal: existingBreak.end ? toLocalDatetimeInput(new Date(existingBreak.end.dateTime)) : '',
            }
            : getDefaultBreakInputValues();
        const startVal = defaultBreak.startVal;
        const endVal = defaultBreak.endVal;

        const row = document.createElement('div');
        row.className = 'break-edit-row';
        row.dataset.breakId = breakId;
        row.innerHTML = `
    <div>
      <label>Break ${idx + 1} Start</label>
      <input type="datetime-local" class="break-start" value="${startVal}" />
    </div>
    <div>
      <label>Break ${idx + 1} End</label>
      <input type="datetime-local" class="break-end" value="${endVal}" />
    </div>
        <button class="btn btn-danger btn-break-remove" title="Remove break">✕</button>
  `;
        row.querySelector('button').addEventListener('click', () => row.remove());
        container.appendChild(row);
    }

    function parseEditModalDateTimeValue(inputValue, originalIso = '') {
        if (!inputValue) {
            return null;
        }

        if (originalIso) {
            const originalDate = new Date(originalIso);
            if (!Number.isNaN(originalDate.getTime()) && toLocalDatetimeInput(originalDate) === inputValue) {
                return originalDate;
            }
        }

        return new Date(inputValue);
    }

    function getEditModalTimeRangeMs() {
        const inVal = $('edit-clock-in').value;
        const outVal = $('edit-clock-out').value;
        const startMs = inVal
            ? new Date(inVal).getTime()
            : (editTarget?.clockIn?.dateTime ? new Date(editTarget.clockIn.dateTime).getTime() : NaN);

        if (!Number.isFinite(startMs)) {
            return null;
        }

        let endMs = outVal
            ? new Date(outVal).getTime()
            : (editTarget?.clockOut?.dateTime ? new Date(editTarget.clockOut.dateTime).getTime() : Date.now());

        if (!Number.isFinite(endMs)) {
            endMs = startMs + DEFAULT_BREAK_DURATION_MS;
        }

        if (endMs < startMs) {
            endMs = startMs;
        }

        return { startMs, endMs };
    }

    function clampMs(value, min, max) {
        return Math.min(Math.max(value, min), max);
    }

    function clamp(value, min, max) {
        return Math.min(Math.max(value, min), max);
    }

    function getDefaultBreakInputValues() {
        const range = getEditModalTimeRangeMs();
        if (!range) {
            return { startVal: '', endVal: '' };
        }

        const spanMs = Math.max(0, range.endMs - range.startMs);
        const durationMs = spanMs > 0 ? Math.min(DEFAULT_BREAK_DURATION_MS, spanMs) : DEFAULT_BREAK_DURATION_MS;
        const midpointMs = range.startMs + Math.floor(spanMs / 2);

        let startMs = midpointMs - Math.floor(durationMs / 2);
        let endMs = startMs + durationMs;

        if (spanMs > 0) {
            if (startMs < range.startMs) {
                startMs = range.startMs;
                endMs = startMs + durationMs;
            }
            if (endMs > range.endMs) {
                endMs = range.endMs;
                startMs = endMs - durationMs;
            }

            startMs = clampMs(snapMs(startMs), range.startMs, range.endMs);
            endMs = clampMs(snapMs(endMs), range.startMs, range.endMs);
            if (endMs <= startMs) {
                endMs = Math.min(range.endMs, startMs + durationMs);
            }
        } else {
            startMs = Math.max(snapMs(startMs), range.startMs);
            endMs = startMs + durationMs;
        }

        return {
            startVal: toLocalDatetimeInput(new Date(startMs)),
            endVal: toLocalDatetimeInput(new Date(endMs)),
        };
    }

    async function saveEditModal() {
        if (!editTarget || !selectedTeam) return;

        const inVal = $('edit-clock-in').value;
        const outVal = $('edit-clock-out').value;
        const notes = $('edit-notes').value.trim();

        if (!inVal) {
            toast('Clock-in time is required', 'error');
            return;
        }
        const newStart = parseEditModalDateTimeValue(inVal, editTarget.clockIn?.dateTime || '');
        const newEnd = parseEditModalDateTimeValue(outVal, editTarget.clockOut?.dateTime || '');

        if (newEnd && newEnd <= newStart) {
            toast('Clock-out must be after clock-in', 'error');
            return;
        }

        // Collect breaks from modal
        const breakRows = $('breaks-edit-container').querySelectorAll('.break-edit-row');
        const breaks = [];
        for (const row of breakRows) {
            const bStart = row.querySelector('.break-start').value;
            const bEnd = row.querySelector('.break-end').value;
            if (!bStart) continue;
            const existingBreak = editTarget.breaks.find(item => item.breakId === (row.dataset.breakId || '')) || null;
            const bStartDt = parseEditModalDateTimeValue(bStart, existingBreak?.start?.dateTime || '');
            const bEndDt = parseEditModalDateTimeValue(bEnd, existingBreak?.end?.dateTime || '');

            if (newEnd && bStartDt < newStart) { toast('Break start must be within card time range', 'error'); return; }
            if (newEnd && bEndDt && bEndDt > newEnd) { toast('Break end must be within card time range', 'error'); return; }
            if (bEndDt && bEndDt <= bStartDt) { toast('Break end must be after break start', 'error'); return; }

            breaks.push({
                breakId: row.dataset.breakId || undefined,
                start: { dateTime: bStartDt.toISOString() },
                end: bEndDt ? { dateTime: bEndDt.toISOString() } : undefined,
            });
        }

        const updatedCard = getUpdatedCardForEdit(editTarget, {
            clockIn: newStart,
            clockOut: newEnd,
            breaks,
            notes,
        });
        const body = buildTimeCardUpdateBody(updatedCard);

        const btn = $('btn-edit-save');
        btn.disabled = true;
        btn.textContent = 'Saving…';
        try {
            const response = await graphFetch(`/teams/${selectedTeam.id}/schedule/timeCards/${editTarget.id}`, {
                method: 'PUT',
                body: JSON.stringify(body),
            });

            applyServerConfirmedCard(resolveUpdatedTimeCard(response, updatedCard, 'Save timecard'));

            toast('Timecard saved', 'success');
            closeEditModal();
        } catch (e) {
            toast('Save failed: ' + e.message, 'error');
        } finally {
            btn.disabled = false;
            btn.textContent = 'Save Changes';
        }
    }

    // ─────────────────────────────────────────────
    // Persist timeline drag/resize
    // ─────────────────────────────────────────────
    function buildProjectedTimelineUpdate(card, newStart, newEnd) {
        const effectiveEnd = card?.clockOut ? newEnd : null;
        const updatedCard = getUpdatedCardForLocalState(card, {
            clockIn: newStart,
            clockOut: effectiveEnd,
        });

        return {
            effectiveEnd,
            updatedCard,
            requestBody: buildTimeCardUpdateBody(updatedCard),
        };
    }

    async function persistCardTimeUpdate(card, newStart, newEnd) {
        const projection = buildProjectedTimelineUpdate(card, newStart, newEnd);

        try {
            const response = await graphFetch(`/teams/${selectedTeam.id}/schedule/timeCards/${card.id}`, {
                method: 'PUT',
                body: JSON.stringify(projection.requestBody),
            });

            applyServerConfirmedCard(resolveUpdatedTimeCard(response, projection.updatedCard, 'Update timecard'));

            toast(`Updated: ${fmtTime(newStart)} → ${projection.effectiveEnd ? fmtTime(projection.effectiveEnd) : 'active'}`, 'success');
        } catch (e) {
            toast('Update failed: ' + e.message, 'error');
            reRenderWeeks();
        }
    }

    async function extendCardToNow(card) {
        if (!card?.clockIn || !card?.clockOut) {
            return;
        }

        const currentEndMs = new Date(card.clockOut.dateTime).getTime();
        const targetEndMs = Date.now() - EXTEND_TO_NOW_LAG_MS;
        if (!Number.isFinite(targetEndMs) || targetEndMs <= currentEndMs) {
            toast('Card already reaches now', 'info');
            return;
        }

        setHighlightedCardId(card.id);
        await persistCardTimeUpdate(
            card,
            new Date(card.clockIn.dateTime),
            new Date(targetEndMs),
        );
    }

    function getUpdatedCardForBreakRange(card, breakIndex, { start, end }) {
        const updatedCard = {
            ...card,
            clockIn: card.clockIn ? {
                ...card.clockIn,
            } : null,
            clockOut: card.clockOut ? {
                ...card.clockOut,
            } : null,
            breaks: cloneCardBreaks(card.breaks),
            notes: cloneItemBody(card.notes),
            lastModifiedDateTime: new Date().toISOString(),
        };

        const targetBreak = updatedCard.breaks[breakIndex];
        if (!targetBreak) {
            return null;
        }

        targetBreak.start = start ? {
            ...(targetBreak.start || {}),
            dateTime: start.toISOString(),
        } : null;
        targetBreak.end = end ? {
            ...(targetBreak.end || {}),
            dateTime: end.toISOString(),
        } : null;
        updatedCard.state = deriveCardState(updatedCard);
        return updatedCard;
    }

    async function persistBreakTimeUpdate(card, breakIndex, newStart, newEnd) {
        if (!selectedTeam || !card?.clockIn) {
            reRenderWeeks();
            return;
        }

        const cardStartMs = new Date(card.clockIn.dateTime).getTime();
        const cardEndMs = card.clockOut ? new Date(card.clockOut.dateTime).getTime() : Date.now();
        if (newStart.getTime() < cardStartMs || newEnd.getTime() > cardEndMs) {
            toast('Break must stay within the timecard range', 'error');
            reRenderWeeks();
            return;
        }
        if (newEnd <= newStart) {
            toast('Break end must be after break start', 'error');
            reRenderWeeks();
            return;
        }

        const updatedCard = getUpdatedCardForBreakRange(card, breakIndex, {
            start: newStart,
            end: newEnd,
        });
        if (!updatedCard) {
            reRenderWeeks();
            return;
        }

        try {
            const response = await graphFetch(`/teams/${selectedTeam.id}/schedule/timeCards/${card.id}`, {
                method: 'PUT',
                body: JSON.stringify(buildTimeCardUpdateBody(updatedCard)),
            });

            applyServerConfirmedCard(resolveUpdatedTimeCard(response, updatedCard, 'Break update'));

            toast(`Break updated: ${fmtTime(newStart)} → ${fmtTime(newEnd)}`, 'success');
        } catch (e) {
            toast('Break update failed: ' + e.message, 'error');
            reRenderWeeks();
        }
    }

    // ─────────────────────────────────────────────
    // Confirm modal
    // ─────────────────────────────────────────────
    function closeConfirm(ok) {
        $('confirm-modal').classList.add('hidden');
        $('confirm-beta-warn').style.display = 'none';
        if (confirmCallback) {
            const cb = confirmCallback;
            confirmCallback = null;
            cb(ok);
        }
    }

    // ─────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────
    function reRenderWeeks() {
        renderWeeks();
    }

    function updateCurrentWeekTotal(group) {
        const totalMs = group.cards.reduce((sum, card) => sum + workedMs(card), 0);
        const titleTotal = `${formatDurationHms(totalMs)}`;
        currentWeekTotal.textContent = titleTotal;
        currentWeekTotal.style.display = selectedTeam ? 'inline-flex' : 'none';
        currentWeekTotal.classList.toggle('hidden-start', !selectedTeam);
        updateDocumentTitle(selectedTeam ? titleTotal : '');
    }

    function updateDocumentTitle(totalText = '') {
        document.title = totalText ? `${totalText} - ${defaultPageTitle}` : defaultPageTitle;
    }

    function stopLiveClockUpdates() {
        if (liveClockTimerId !== null) {
            window.clearInterval(liveClockTimerId);
            liveClockTimerId = null;
        }
    }

    function syncLiveClockUpdates() {
        const hasRunningCard = Boolean(selectedTeam && allTimeCards.some(card => card.clockIn && !card.clockOut));
        if (!hasRunningCard) {
            stopLiveClockUpdates();
            return;
        }

        updateLiveClockDisplays();
        if (liveClockTimerId === null) {
            liveClockTimerId = window.setInterval(updateLiveClockDisplays, 1000);
        }
    }

    function updateLiveClockDisplays() {
        if (!selectedTeam) {
            stopLiveClockUpdates();
            return;
        }

        const currentWeekGroup = getSelectedWeekGroup(allTimeCards);
        updateCurrentWeekTotal(currentWeekGroup);
        updateVisibleDayTotals(currentWeekGroup.cards);

        allTimeCards.forEach(card => {
            if (!card.clockIn || card.clockOut) {
                return;
            }

            document.querySelectorAll(`.tc-row[data-card-id="${card.id}"]`).forEach(row => {
                updateLiveCardRow(row, card);
            });
        });

        document.querySelectorAll('.tc-block[data-end-ms=""]').forEach(block => {
            updateLiveTimelineBlock(block);
        });
    }

    function updateVisibleDayTotals(cards) {
        const totalsByDay = new Map();
        cards.forEach(card => {
            const anchorDateTime = getCardAnchorDateTime(card);
            if (!anchorDateTime) {
                return;
            }
            const dayStartMs = getDayStart(new Date(anchorDateTime)).getTime();
            totalsByDay.set(dayStartMs, (totalsByDay.get(dayStartMs) || 0) + workedMs(card));
        });

        document.querySelectorAll('.day-section').forEach(section => {
            const totalEl = section.querySelector('.day-section-total');
            if (!totalEl) {
                return;
            }
            const dayStartMs = Number(section.dataset.dayStart);
            totalEl.textContent = `${formatDurationHms(totalsByDay.get(dayStartMs) || 0)} worked`;
        });
    }

    function updateLiveCardRow(row, card) {
        const clockInEl = row.querySelector('[data-field="clock-in"]');
        const clockOutEl = row.querySelector('[data-field="clock-out"]');

        if (clockInEl) {
            const inText = card.clockIn ? fmtDateTime(new Date(card.clockIn.dateTime)) : '—';
            clockInEl.innerHTML = `<span class="tc-row-time-value">${inText}</span>`;
        }

        if (clockOutEl) {
            const outText = card.clockOut ? fmtDateTime(new Date(card.clockOut.dateTime)) : '—';
            clockOutEl.innerHTML = `<span class="tc-row-time-value">${outText}</span><span class="tc-row-total">(${formatDurationHms(workedMs(card))})</span>`;
        }
    }

    function updateLiveTimelineBlock(block) {
        if (block.dataset.isInteracting === 'true') {
            return;
        }

        const cardId = block.dataset.cardId;
        const card = allTimeCards.find(item => item.id === cardId);
        if (!card || !card.clockIn || card.clockOut) {
            return;
        }

        const axisStartMs = Number(block.dataset.weekStartMs);
        const totalSpanMs = Number(block.dataset.totalSpanMs);
        if (!Number.isFinite(axisStartMs) || !Number.isFinite(totalSpanMs) || totalSpanMs <= 0) {
            return;
        }

        const startMs = new Date(card.clockIn.dateTime).getTime();
        const axisEndMs = axisStartMs + totalSpanMs;
        const clippedStartMs = Math.max(startMs, axisStartMs);
        const clippedEndMs = Math.min(Date.now(), axisEndMs);
        const leftPct = ((clippedStartMs - axisStartMs) / totalSpanMs) * 100;
        const widthPct = ((clippedEndMs - clippedStartMs) / totalSpanMs) * 100;

        block.style.left = `${Math.max(0, leftPct)}%`;
        block.style.width = `${Math.max(0.1, widthPct)}%`;
        block.title = `${fmtDateTime(new Date(startMs))} (active)`;
    }

    function summarizeCardForDebug(card) {
        if (!card) {
            return null;
        }

        const openBreakCount = (card.breaks || []).filter(item => item.start && !item.end).length;
        return {
            id: card.id || null,
            state: deriveCardState(card),
            clockIn: card.clockIn?.dateTime || null,
            clockOut: card.clockOut?.dateTime || null,
            breakCount: Array.isArray(card.breaks) ? card.breaks.length : 0,
            openBreakCount,
            lastModifiedDateTime: card.lastModifiedDateTime || null,
        };
    }

    function getPersistedTimeCardCacheSnapshotForSelfTest(teamId = '') {
        const resolvedTeamId = teamId || localStorage.getItem(SESSION_KEY_SELECTED_TEAM_ID) || '';
        if (!resolvedTeamId) {
            return {
                selectedTeamId: '',
                reason: 'No saved team is currently selected.',
            };
        }

        const cache = getTeamCache(resolvedTeamId);
        const now = Date.now();
        const selectedWeek = getResolvedWeekStart();
        const selectedWeekKey = getWeekCacheKey(selectedWeek);
        const hasCurrentWeek = cache.currentWeekKey === selectedWeekKey;

        return {
            selectedTeamId: resolvedTeamId,
            selectedWeek: selectedWeekKey,
            storage: 'memory-only',
            inMemory: {
                cachedWeekKey: cache.currentWeekKey || null,
                currentWeek: hasCurrentWeek ? {
                    cardsCount: cache.currentWeekCards.length,
                    fetchedAt: cache.currentWeekFetchedAt || 0,
                    ageMs: cache.currentWeekFetchedAt ? (now - cache.currentWeekFetchedAt) : null,
                    isStale: !cache.currentWeekFetchedAt || (now - cache.currentWeekFetchedAt) > CACHE_STALE_MS,
                } : null,
                hasPendingRequest: cache.pendingWeekKey === selectedWeekKey,
                warnedMultiPageWeekKey: cache.warnedMultiPageWeekKey || null,
                activeCard: summarizeCardForDebug(cache.activeCard),
            },
            persisted: null,
            pendingMutations: [],
            probes: {
                lastModifiedFilterSupport: cache.lastModifiedFilterSupport,
                lastModifiedOrderBySupport: cache.lastModifiedOrderBySupport,
                filterSupportMatrix: cache.filterSupportMatrix,
                combinedWeekQuerySupport: cache.combinedWeekQuerySupport,
            },
            timers: {
                cacheStaleMs: CACHE_STALE_MS,
                backgroundRefreshMs: BACKGROUND_REFRESH_MS,
                pageSize: TIMECARD_PAGE_SIZE,
            },
        };
    }

    async function fetchWeekTimeCardsForSelfTest(teamId, weekStartInput = '', options = {}) {
        await confirmWeekTimeCardQuerySupport(teamId);

        const weekStart = getResolvedWeekStart(weekStartInput);
        const snapshot = await fetchWeekTimeCards(teamId, weekStart, {
            includeTiming: options.includeTiming === true,
        });

        return {
            weekStart: formatDateInputValue(weekStart),
            cardsCount: snapshot.cards.length,
            activeCard: summarizeCardForDebug(snapshot.activeCard),
            timings: snapshot.timings,
            sample: snapshot.cards.slice(0, 10).map(summarizeCardForDebug),
        };
    }

    async function resolveTimeCardActionResponseForSelfTest(teamId, options = {}) {
        if (!teamId) {
            throw new Error('A teamId is required for timecard action self tests');
        }

        const resolvedCard = await resolveTimeCardActionResponse(null, {
            actionLabel: options.actionLabel || 'Timecard action self test',
            teamId,
            cardId: options.cardId || '',
            weekStartInput: Object.prototype.hasOwnProperty.call(options, 'weekStartInput')
                ? options.weekStartInput
                : getResolvedWeekStart(new Date()),
        });

        return summarizeCardForDebug(resolvedCard);
    }

    function getCardSortMs(card) {
        if (card.clockOut?.dateTime) return new Date(card.clockOut.dateTime).getTime();
        if (card.clockIn?.dateTime) return new Date(card.clockIn.dateTime).getTime();
        return card.createdDateTime ? new Date(card.createdDateTime).getTime() : 0;
    }

    function ensureSelectedWeek() {
        if (selectedWeekStart instanceof Date && !Number.isNaN(selectedWeekStart.getTime())) {
            selectedWeekStart = getWeekStart(selectedWeekStart);
            return;
        }

        const stored = localStorage.getItem(SESSION_KEY_SELECTED_WEEK);
        if (stored) {
            const parsed = new Date(`${stored}T00:00:00`);
            if (!Number.isNaN(parsed.getTime())) {
                selectedWeekStart = getWeekStart(parsed);
                return;
            }
        }

        const fallback = activeTimeCard?.clockIn?.dateTime
            ? new Date(activeTimeCard.clockIn.dateTime)
            : new Date();
        selectedWeekStart = getWeekStart(fallback);
    }

    function shiftSelectedWeek(days) {
        ensureSelectedWeek();
        selectedWeekStart = new Date(selectedWeekStart.getTime() + (days * DAY_MS));
        selectedWeekStart = getWeekStart(selectedWeekStart);
        persistSelectedWeek();
        renderWeeks();
        void loadTimeCards();
    }

    function onWeekPickerChange() {
        const value = weekPicker.value;
        if (!value) return;
        const picked = new Date(`${value}T00:00:00`);
        if (Number.isNaN(picked.getTime())) return;
        selectedWeekStart = getWeekStart(picked);
        persistSelectedWeek();
        renderWeeks();
        void loadTimeCards();
    }

    function persistSelectedWeek() {
        if (!selectedWeekStart) return;
        localStorage.setItem(SESSION_KEY_SELECTED_WEEK, formatDateInputValue(selectedWeekStart));
        syncWeekControls();
    }

    function syncWeekControls() {
        ensureSelectedWeek();
        const weekEnd = new Date(selectedWeekStart.getTime() + (6 * DAY_MS));
        weekPicker.value = formatDateInputValue(selectedWeekStart);
        weekRangeLabel.textContent = `${fmtDate(selectedWeekStart)} — ${fmtDate(weekEnd)}`;
    }

    function formatDateInputValue(date) {
        const year = date.getFullYear();
        const month = `${date.getMonth() + 1}`.padStart(2, '0');
        const day = `${date.getDate()}`.padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    function workedMs(card) {
        if (!card.clockIn) return 0;
        const start = new Date(card.clockIn.dateTime).getTime();
        const end = card.clockOut ? new Date(card.clockOut.dateTime).getTime() : Date.now();
        const worked = end - start;
        // Subtract breaks
        const breakMs = card.breaks.reduce((sum, b) => {
            if (!b.start) return sum;
            const bs = new Date(b.start.dateTime).getTime();
            const be = b.end ? new Date(b.end.dateTime).getTime() : Date.now();
            return sum + (be - bs);
        }, 0);
        return Math.max(0, worked - breakMs);
    }

    function formatDurationHms(durationMs) {
        const totalSeconds = Math.max(0, Math.floor(durationMs / 1000));
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    }

    const dateTimeFmt = new Intl.DateTimeFormat(undefined, {
        month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true,
    });
    const dateFmt = new Intl.DateTimeFormat(undefined, {
        month: 'short', day: 'numeric', weekday: 'short',
    });
    const timeFmt = new Intl.DateTimeFormat(undefined, {
        hour: '2-digit', minute: '2-digit', hour12: true,
    });
    const hourFmt = new Intl.DateTimeFormat(undefined, {
        hour: 'numeric', hour12: true,
    });
    const dayShortFmt = new Intl.DateTimeFormat(undefined, {
        weekday: 'short', month: 'numeric', day: 'numeric',
    });

    function fmtDateTime(d) { return dateTimeFmt.format(d); }
    function fmtDate(d) { return dateFmt.format(d); }
    function fmtTime(d) { return timeFmt.format(d); }
    function fmtHour(d) { return hourFmt.format(d); }
    function fmtDayShort(d) { return dayShortFmt.format(d); }

    function toLocalDatetimeInput(d) {
        // Format: YYYY-MM-DDTHH:MM  (no seconds, no Z)
        const pad = n => String(n).padStart(2, '0');
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }

    function escHtml(str) {
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    // ─────────────────────────────────────────────
    // Toast notifications
    // ─────────────────────────────────────────────
    function toast(msg, type = 'info') {
        const area = $('toast-area');
        const el = document.createElement('div');
        el.className = `toast ${type}`;
        el.textContent = msg;
        area.appendChild(el);
        setTimeout(() => el.remove(), 4000);
    }

    function showError(msg) {
        errorBanner.textContent = msg;
        errorBanner.style.display = 'block';
        setTimeout(() => { errorBanner.style.display = 'none'; }, 8000);
    }

    window.timecardsApp = {
        buildTimeCardsPageUrl,
        buildProjectedTimelineUpdate,
        fetchJoinedTeamsForSelfTest,
        resolveTeamForSelfTest,
        fetchTimeCardsPageForSelfTest,
        fetchWeekTimeCardsForSelfTest,
        resolveTimeCardActionResponseForSelfTest,
        getPersistedTimeCardCacheSnapshotForSelfTest,
        probeTimeCardCombinedWeekQuerySupport,
        probeTimeCardDocumentedFilterSupport,
        probeTimeCardLastModifiedFilterSupport,
        probeTimeCardLastModifiedOrderBySupport,
    };

})();

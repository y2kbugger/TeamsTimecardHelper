// The app page: the team, the week, its days and cards, the toolbar, and the edit dialog.
// After every change the week is read back from Graph, which answers most writes with 204 No
// Content: what is drawn is always what the server holds.
import * as auth from './auth.js';
import { $, escHtml, toast, showError } from './ui.js';
import {
    addDays, assert, buildNewTimeCardBody, buildTimeCardUpdateBody, buildUpdatedCard, buildWeekDays, filterCardsForWeek,
    dedupeCardsById, findActiveCard, formatDateInputValue, formatDurationHms, getCardSortMs,
    FULL_DAY_HOURS, getDayStart, getWeekRange, getWeekStart, isLongCard, isSameDay, normalizeCard, weekHourWindow, parseLocalDate,
    addBreak, readEditForm, toLocalDatetimeInput, workedMs,
} from './core.js';
import { attachTimeline, isDragging, renderTimeline, updateNowLines, updateRunningSpans } from './timeline.js';
import { mockEnabled } from './mockgraph.js';

// Mock mode keeps its own keys, so it never disturbs a real session in another tab.
const STORAGE_PREFIX = mockEnabled ? 'mock_' : '';
export const STORAGE_KEYS = Object.freeze({
    selectedTeamId: `${STORAGE_PREFIX}tc_selected_team_id`, // localStorage
    selectedWeek: `${STORAGE_PREFIX}tc_selected_week`, // sessionStorage, as below
    lastActivity: `${STORAGE_PREFIX}tc_last_activity_at`,
    joinedTeams: `${STORAGE_PREFIX}tc_joined_teams`,
});
// Left by earlier versions of the app.
const STALE_STORAGE_PREFIXES = ['tc_timecard_cache_v1', 'tc_lastmodified_', 'tc_selected_week'];

const WEEK_STALE_MS = 45 * 1000;
const BACKGROUND_REFRESH_MS = 30 * 1000;
const LIVE_TICK_MS = 1000;
const IDLE_RETURN_TO_THIS_WEEK_MS = 60 * 60 * 1000;
const EXTEND_TO_NOW_LAG_MS = 30 * 1000;
const TIMECARD_PAGE_SIZE = 250;
const TIMECARD_PAGES_MAX = 20;

const state = {
    teams: [],
    team: null,
    weekStart: null, // the shown week's Sunday, local midnight
    cards: [], // the shown week's cards
    cardsWeekKey: '',
    cardsFetchedAtMs: 0,
    activeCard: null, // the open card, from the current week whatever week is shown
    highlightedCardId: null,
    renderHeld: false, // a render was asked for during a drag
    renderedDayKey: '', // today, when the week was last drawn; a new day draws it again
    renderedHoursKey: '', // the hours the timelines span, when last drawn
    busyCount: 0,
    changesInFlight: 0, // writes not yet read back; until then the cards held are out of date
    weekLoad: null, // { key, promise } while a week loads
    loadCount: 0, // week loads started; only the newest one's answer is shown
    editCard: null,
    confirmResolve: null,
};

const el = Object.fromEntries([
    'app', 'user-info', 'btn-signout', 'btn-team', 'toolbar-title', 'busy', 'status-card',
    'btn-week-prev', 'btn-week-next', 'week-picker', 'week-range-label', 'week-total', 'btn-clock',
    'team-modal', 'team-picker-msg', 'team-picker-select', 'btn-team-save', 'btn-team-cancel', 'status-panel',
    'timecards-loading', 'weeks', 'drag-tooltip', 'edit-modal', 'edit-clock-in', 'edit-clock-out',
    'edit-notes', 'breaks-edit', 'btn-edit-save', 'btn-edit-cancel', 'confirm-modal',
    'confirm-title', 'confirm-msg', 'confirm-beta-warn', 'btn-confirm-ok', 'btn-confirm-cancel',
].map(id => [id.replace(/-(.)/g, (_, c) => c.toUpperCase()), $(id)]));

const show = (element, visible) => { element.hidden = !visible; };

let weeklyAverage = null;

if (document.body.dataset.page === 'app') void boot();

async function boot() {
    for (const key of Object.keys(localStorage)) {
        if (STALE_STORAGE_PREFIXES.some(prefix => key.startsWith(prefix))) localStorage.removeItem(key);
    }
    if (!await auth.requireAppSession()) return;
    const account = auth.getCurrentAccount();
    el.userInfo.textContent = account?.name || account?.username || '';
    show(el.btnSignout, true);
    show(el.app, true);
    bindEvents();
    weeklyAverage = await import('./weeklyavg.js');
    weeklyAverage.initWeeklyAverageFeature({ fetchTimeCardsForDateRange, getCurrentAppState, showWeek: selectWeek });
    await loadTeams();
    setInterval(onBackgroundTick, BACKGROUND_REFRESH_MS);
    setInterval(onLiveTick, LIVE_TICK_MS);
}

function bindEvents() {
    el.btnSignout.addEventListener('click', () => {
        sessionStorage.removeItem(STORAGE_KEYS.joinedTeams);
        void auth.signOut();
    });
    el.btnTeam.addEventListener('click', () => openTeamDialog('Your choice is saved in this browser.'));
    el.btnTeamCancel.addEventListener('click', closeTeamDialog);
    el.btnTeamSave.addEventListener('click', () => {
        const team = state.teams.find(t => t.id === el.teamPickerSelect.value);
        if (!team) return;
        show(el.teamModal, false);
        if (team.id !== state.team?.id) void selectTeam(team);
    });
    el.btnWeekPrev.addEventListener('click', () => selectWeek(addDays(state.weekStart, -7)));
    el.btnWeekNext.addEventListener('click', () => selectWeek(addDays(state.weekStart, 7)));
    el.weekPicker.addEventListener('change', () => {
        const picked = parseLocalDate(el.weekPicker.value);
        if (picked) selectWeek(picked);
    });
    el.btnClock.addEventListener('click', () => void clockAction('clockIn'));
    el.btnEditCancel.addEventListener('click', closeEditDialog);
    el.btnEditSave.addEventListener('click', () => void saveEditDialog());
    el.btnConfirmOk.addEventListener('click', () => closeConfirm(true));
    el.btnConfirmCancel.addEventListener('click', () => closeConfirm(false));
    document.addEventListener('keydown', event => {
        if (event.key !== 'Escape') return;
        if (!el.confirmModal.hidden) closeConfirm(false);
        else if (!el.editModal.hidden) closeEditDialog();
        else if (!el.teamModal.hidden) closeTeamDialog();
        else highlightCard(null);
    });
    // A click away from every card lets go of the selected one; one in a dialog is about it.
    document.addEventListener('pointerdown', event => {
        if (!event.target.closest('[data-card-id], .modal-overlay')) highlightCard(null);
    });
    document.addEventListener('pointerdown', noteActivity, true);
    document.addEventListener('keydown', noteActivity, true);
    // Back on the tab: the clock's part at once, then Graph's when the read comes back.
    // Leaving the tab lets go of the selected card, as a click away does.
    document.addEventListener('visibilitychange', () => {
        if (document.hidden) return highlightCard(null);
        onLiveTick();
        onBackgroundTick();
    });
    el.weeks.addEventListener('click', onWeeksClick);
    narrowScreen.addEventListener('change', () => { if (state.team && !el.weeks.hidden) render(); });
    attachTimeline(el.weeks, {
        // None while a change is on its way: a drag of the card as it was would undo the change.
        cardById: id => (state.changesInFlight ? undefined : state.cards.find(card => card.id === id)),
        cards: () => (state.changesInFlight ? null : state.cards),
        onSelect: highlightCard,
        onCommit: (card, movedMs) => void saveCard(card, `Moved to ${timeFormat.format(movedMs)}`),
        onAddBreak: (card, proposal) => void saveCard(addBreak(card, proposal),
            `Break added ${timeFormat.format(proposal.startMs)} → ${timeFormat.format(proposal.endMs)}`),
        onAddCard: proposal => void change(null,
            `Card added ${timeFormat.format(proposal.startMs)} → ${timeFormat.format(proposal.endMs)}`,
            () => auth.graphFetch(timeCardsPath(state.team.id), { method: 'POST', body: JSON.stringify(buildNewTimeCardBody(proposal)) })),
        // A saved drag is drawn by the read-back after the save; drawing now would snap it back.
        onEnd: committed => { if (!committed) render(); },
        tooltip: {
            show(x, y, text) {
                el.dragTooltip.textContent = text;
                el.dragTooltip.style.transform = `translate(${x + 12}px, ${y - 30}px)`;
                show(el.dragTooltip, true);
            },
            hide: () => show(el.dragTooltip, false),
        },
    });
}

// For the weekly average report.
export function getCurrentAppState() {
    return {
        team: state.team ? { ...state.team } : null,
        selectedWeekStart: state.weekStart ? formatDateInputValue(state.weekStart) : '',
        dataVersion: state.cardsFetchedAtMs,
    };
}

// ── Graph ───────────────────────────────────────────────────────────────────

// Every request goes through here, so the toolbar can show that something is on its way.
async function graph(request) {
    state.busyCount += 1;
    show(el.busy, true);
    try {
        return await request();
    } finally {
        state.busyCount -= 1;
        assert(state.busyCount >= 0, 'busy count never negative');
        show(el.busy, state.busyCount > 0);
    }
}

const timeCardsPath = teamId => `/teams/${teamId}/schedule/timeCards`;

export async function fetchJoinedTeams({ preferCache = false } = {}) {
    if (preferCache) {
        try {
            const cached = JSON.parse(sessionStorage.getItem(STORAGE_KEYS.joinedTeams));
            if (Array.isArray(cached)) return cached;
        } catch { /* fetch below */ }
    }
    const data = await graph(() => auth.graphFetch('/me/joinedTeams?$select=id,displayName'));
    const teams = (data?.value || [])
        .filter(team => team?.id && team?.displayName)
        .map(team => ({ id: team.id, displayName: team.displayName }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName));
    sessionStorage.setItem(STORAGE_KEYS.joinedTeams, JSON.stringify(teams));
    return teams;
}

// Cards that clocked in within [rangeStart, rangeEndInclusive], every page. The tenant rejects
// `or` in timeCard filters, so this clock-in window is the one query shape there is.
async function fetchClockedInBetween(teamId, rangeStart, rangeEndInclusive) {
    const filter = `clockInEvent/dateTime ge ${rangeStart.toISOString()} and clockInEvent/dateTime le ${rangeEndInclusive.toISOString()}`;
    let nextUrl = `${timeCardsPath(teamId)}?$top=${TIMECARD_PAGE_SIZE}&$filter=${encodeURIComponent(filter)}`;
    const cards = [];
    let pageCount = 0;
    for (; nextUrl; pageCount += 1) {
        if (pageCount === TIMECARD_PAGES_MAX) throw new Error(`More than ${TIMECARD_PAGES_MAX} pages of timecards.`);
        const data = await graph(() => auth.graphFetch(nextUrl));
        (data?.value || []).filter(raw => raw?.id).forEach(raw => cards.push(normalizeCard(raw)));
        nextUrl = data?.['@odata.nextLink'] || null;
    }
    return { cards: dedupeCardsById(cards), pageCount, fetchedAtMs: Date.now() };
}

// The week of `weekInput` (a date or 'YYYY-MM-DD'; this week when absent).
export async function fetchWeekTimeCards(teamId, weekInput = null) {
    const { weekStart, weekEndInclusive } = getWeekRange(parseLocalDate(weekInput) || new Date());
    const result = await fetchClockedInBetween(teamId, weekStart, weekEndInclusive);
    const cards = filterCardsForWeek(result.cards, weekStart);
    return { ...result, weekStart, cards, activeCard: findActiveCard(cards) };
}

// Whole local days, from rangeStart's through rangeEnd's (the weekly average report).
export function fetchTimeCardsForDateRange(teamId, rangeStart, rangeEnd) {
    const end = new Date(addDays(getDayStart(rangeEnd), 1).getTime() - 1);
    return fetchClockedInBetween(teamId, getDayStart(rangeStart), end);
}

// ── Teams ───────────────────────────────────────────────────────────────────

async function loadTeams() {
    el.toolbarTitle.textContent = 'Loading teams…';
    showStatus('<span class="spinner"></span>Loading teams…');
    try {
        state.teams = await fetchJoinedTeams({ preferCache: true });
    } catch (error) {
        showStatus(`Could not load your teams.<br>${escHtml(error.message)}`, true);
        return;
    }
    const savedId = localStorage.getItem(STORAGE_KEYS.selectedTeamId);
    const saved = state.teams.find(team => team.id === savedId);
    if (saved) return selectTeam(saved);
    if (state.teams.length === 1) return selectTeam(state.teams[0]);
    if (!state.teams.length) return showStatus('You are not in any team.');
    el.toolbarTitle.textContent = 'Choose a team';
    showStatus('Choose a team to see its timecards.');
    openTeamDialog(savedId
        ? 'Your saved team is no longer available. Choose another; it is saved in this browser.'
        : 'Choose the team to use for timecards. Your choice is saved in this browser.');
}

function showStatus(html, isError = false) {
    el.statusPanel.innerHTML = html;
    el.statusPanel.classList.toggle('is-error', isError);
    showOnly(el.statusPanel);
}

// The content area shows one of: the status panel, the loading line, the weeks.
function showOnly(panel) {
    for (const each of [el.statusPanel, el.timecardsLoading, el.weeks]) show(each, each === panel);
    const weekShown = panel === el.weeks || panel === el.timecardsLoading;
    show(el.statusCard, weekShown);
    // Clock In only while no card is open; an open card has its own Clock Out.
    show(el.btnClock, weekShown && !state.activeCard);
    // The team's name is the button that changes it, when there is another to change to.
    el.btnTeam.disabled = state.teams.length < 2;
    weeklyAverage?.setWeeklyAverageTriggerVisible(Boolean(state.team) && weekShown);
}

// With no team chosen yet it cannot be dismissed: there is nothing to show without one.
function openTeamDialog(message) {
    el.teamPickerMsg.textContent = message;
    el.teamPickerSelect.innerHTML = state.teams
        .map(team => `<option value="${escHtml(team.id)}">${escHtml(team.displayName)}</option>`).join('');
    el.teamPickerSelect.value = state.team?.id || state.teams[0]?.id || '';
    show(el.btnTeamCancel, Boolean(state.team));
    show(el.teamModal, true);
    el.teamPickerSelect.focus();
}

function closeTeamDialog() {
    if (state.team) show(el.teamModal, false);
}

async function selectTeam(team) {
    state.team = team;
    state.cards = [];
    state.cardsWeekKey = '';
    state.activeCard = null;
    localStorage.setItem(STORAGE_KEYS.selectedTeamId, team.id);
    el.toolbarTitle.textContent = team.displayName;
    state.weekStart = restoredWeekStart();
    await loadWeek();
}

// ── The week ────────────────────────────────────────────────────────────────

// The week chosen earlier in this tab, while the user has been active; otherwise this week.
function restoredWeekStart() {
    const stored = isIdle() ? null : parseLocalDate(sessionStorage.getItem(STORAGE_KEYS.selectedWeek));
    return getWeekStart(stored || new Date());
}

function noteActivity() {
    sessionStorage.setItem(STORAGE_KEYS.lastActivity, String(Date.now()));
}

function isIdle() {
    const lastActivityMs = Number(sessionStorage.getItem(STORAGE_KEYS.lastActivity)) || 0;
    return Date.now() - lastActivityMs > IDLE_RETURN_TO_THIS_WEEK_MS;
}

const weekKey = weekStart => formatDateInputValue(weekStart);

function selectWeek(date) {
    if (!state.team) return;
    state.weekStart = getWeekStart(date);
    sessionStorage.setItem(STORAGE_KEYS.selectedWeek, weekKey(state.weekStart));
    void loadWeek();
}

// Shows the selected week, reading it from Graph unless a fresh copy is at hand. With the shown
// week not the current one, the current week is read too: the open card lives there. `force`
// reads it whatever is at hand or on its way: after a change, an earlier read is out of date.
async function loadWeek({ force = false } = {}) {
    const team = state.team;
    const key = weekKey(state.weekStart);
    syncWeekControls();
    const fresh = state.cardsWeekKey === key && Date.now() - state.cardsFetchedAtMs < WEEK_STALE_MS;
    if (fresh && !force) return render();
    if (state.cardsWeekKey !== key) showOnly(el.timecardsLoading);
    if (!force && state.weekLoad?.key === `${team.id}:${key}`) return state.weekLoad.promise;

    state.loadCount += 1;
    const loadNumber = state.loadCount;
    const thisWeekKey = weekKey(getWeekStart(new Date()));
    const promise = Promise.all([
        fetchWeekTimeCards(team.id, state.weekStart),
        key === thisWeekKey ? null : fetchWeekTimeCards(team.id, new Date()),
    ]).then(([week, thisWeek]) => {
        if (loadNumber !== state.loadCount || state.team !== team || weekKey(state.weekStart) !== key) return;
        state.cards = week.cards;
        state.cardsWeekKey = key;
        state.cardsFetchedAtMs = week.fetchedAtMs;
        state.activeCard = (thisWeek || week).activeCard;
        render();
    }).catch(error => {
        if (state.cardsWeekKey !== key) showStatus(`Could not load timecards.<br>${escHtml(error.message)}`, true);
        else showError('Could not refresh timecards: ' + error.message);
    }).finally(() => {
        if (state.weekLoad?.promise === promise) state.weekLoad = null;
    });
    state.weekLoad = { key: `${team.id}:${key}`, promise };
    return promise;
}

function onBackgroundTick() {
    if (document.hidden || !state.team || !state.weekStart) return;
    const thisWeekStart = getWeekStart(new Date());
    if (isIdle() && state.weekStart.getTime() !== thisWeekStart.getTime()) return selectWeek(thisWeekStart);
    void loadWeek({ force: true });
}

const dateFormat = new Intl.DateTimeFormat(undefined, { weekday: 'short', month: 'short', day: 'numeric' });
const dateTimeFormat = new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
const timeFormat = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });

function syncWeekControls() {
    el.weekPicker.value = weekKey(state.weekStart);
    el.weekRangeLabel.textContent = `${dateFormat.format(state.weekStart)} – ${dateFormat.format(addDays(state.weekStart, 6))}`;
}

// ── Rendering ───────────────────────────────────────────────────────────────

function render() {
    if (isDragging()) {
        state.renderHeld = true;
        return;
    }
    state.renderHeld = false;
    assert(state.team && state.weekStart, 'a week of a team to render');
    // Another week's cards are never drawn under this week's dates; its load renders when done.
    if (state.cardsWeekKey !== weekKey(state.weekStart)) return;
    syncWeekControls();
    el.btnClock.disabled = false; // disabled again below while a change is on its way
    const nowMs = Date.now();
    state.renderedDayKey = weekKey(getDayStart(nowMs));
    const days = buildWeekDays(state.cards, state.weekStart, nowMs);
    const hours = timelineHours(days, nowMs);
    state.renderedHoursKey = hoursKey(hours);
    const latestId = state.cards.reduce((latest, card) => (!latest || getCardSortMs(card) > getCardSortMs(latest) ? card : latest), null)?.id;
    el.weeks.replaceChildren(...(days.length
        ? days.map(day => renderDay(day, nowMs, latestId, hours))
        : [Object.assign(document.createElement('p'), { className: 'panel-empty', textContent: 'No timecards this week.' })]));
    showOnly(el.weeks);
    renderTotals(days);
    highlightCard(state.highlightedCardId);
    if (state.changesInFlight) disableActions();
}

// A narrow screen (as timecards.css has it) spans the hours the week's cards use, the same on
// every day; a wide one the whole day.
const narrowScreen = window.matchMedia('(max-width: 760px)');

function timelineHours(days, nowMs) {
    return narrowScreen.matches ? weekHourWindow(days, nowMs) : FULL_DAY_HOURS;
}

const hoursKey = hours => `${hours.start}-${hours.end}`;

function disableActions() {
    for (const button of [el.btnClock, ...el.weeks.querySelectorAll('button')]) button.disabled = true;
}

function renderTotals(days) {
    const totalMs = days.reduce((sum, day) => sum + day.workedMs, 0);
    el.weekTotal.textContent = formatDurationHms(totalMs);
    document.title = `${formatDurationHms(totalMs)} · Teams Timecards`;
    days.forEach(day => {
        const total = el.weeks.querySelector(`[data-day-total="${day.dayStartMs}"]`);
        if (total) total.textContent = formatDurationHms(day.workedMs);
    });
}

function renderDay(day, nowMs, latestId, hours) {
    const section = document.createElement('section');
    section.className = 'day';
    section.classList.toggle('is-today', isSameDay(day.dayStart, new Date(nowMs)));
    section.classList.toggle('is-past', day.dayEndMs <= nowMs);
    section.innerHTML = `
        <header class="day-header">
            <h2>${dateFormat.format(day.dayStart)}</h2>
            <span class="day-total" data-day-total="${day.dayStartMs}"></span>
        </header>`;
    section.appendChild(renderTimeline(day, nowMs, hours));
    const list = document.createElement('div');
    list.className = 'card-list';
    day.listed.forEach(card => list.appendChild(renderCard(card, card.id === latestId, nowMs)));
    section.appendChild(list);
    return section;
}

function renderCard(card, isLatest, nowMs) {
    const item = document.createElement('div');
    item.className = 'card';
    item.dataset.cardId = card.id;
    const isLong = isLongCard(card, nowMs);
    item.classList.toggle('is-long', isLong);
    // The open card carries what is done to it: Clock Out here, and its break below.
    const isActive = state.activeCard?.id === card.id;
    const clockOut = card.clockOut ? new Date(card.clockOut.dateTime) : null;
    // Only today's last card: extending an older one to now would add days of time.
    const canExtend = isLatest && clockOut && isSameDay(clockOut, new Date(nowMs)) && !state.activeCard;
    item.innerHTML = `
        <div class="card-row">
            <span class="card-time">${dateTimeFormat.format(new Date(card.clockIn.dateTime))}</span>
            <span class="card-arrow">→</span>
            <span class="card-time">${clockOut ? timeFormat.format(clockOut) : '<span class="running-dot"></span>now'}</span>
            <span class="card-worked" data-worked="${escHtml(card.id)}">${formatDurationHms(workedMs(card, nowMs))}</span>
            ${isLong ? '<span class="card-long" title="Longer than any real shift: check its clock-out">over 16 h</span>' : ''}
            ${card.notes?.content ? `<span class="card-notes" title="${escHtml(card.notes.content)}">${escHtml(card.notes.content)}</span>` : ''}
            ${isActive ? `<button class="btn btn-danger btn-compact" data-action="clockOut"${card.state === 'onBreak' ? ' disabled title="End the break before clocking out"' : ''}>Clock Out</button>` : ''}
            <span class="card-actions">
                ${canExtend ? '<button class="btn btn-secondary btn-compact" data-action="extend">Extend to now</button>' : ''}
                <button class="btn btn-secondary btn-compact" data-action="edit">Edit</button>
                <button class="btn btn-quiet btn-compact btn-icon" data-action="delete" title="Delete timecard" aria-label="Delete timecard">✕</button>
            </span>
        </div>`;
    const breakAction = isActive ? { clockedIn: 'startBreak', onBreak: 'endBreak' }[card.state] : null;
    if (card.breaks.length || breakAction) item.appendChild(renderBreaks(card, breakAction));
    return item;
}

function renderBreaks(card, breakAction) {
    const list = document.createElement('div');
    list.className = 'breaks';
    card.breaks.forEach((b, index) => {
        if (!b.start) return;
        const startMs = new Date(b.start.dateTime).getTime();
        const row = document.createElement('div');
        row.className = 'break-row';
        row.dataset.breakIndex = String(index);
        row.innerHTML = `
            <span class="break-label">Break</span>
            <span class="break-times">${timeFormat.format(startMs)} → ${b.end ? timeFormat.format(new Date(b.end.dateTime)) : 'now'}</span>
            <span class="break-length">${b.end ? `${Math.round((new Date(b.end.dateTime) - startMs) / 60000)} min` : ''}</span>
            ${!b.end && breakAction === 'endBreak' ? '<button class="btn btn-success btn-compact" data-action="endBreak">End Break</button>' : ''}
            <button class="btn btn-quiet btn-compact btn-icon" data-action="delete-break" title="Delete break" aria-label="Delete break">✕</button>`;
        list.appendChild(row);
    });
    if (breakAction === 'startBreak') {
        const row = document.createElement('div');
        row.className = 'break-row';
        row.innerHTML = '<button class="btn btn-warn btn-compact" data-action="startBreak">Start Break</button>';
        list.appendChild(row);
    }
    return list;
}

function highlightCard(cardId) {
    state.highlightedCardId = cardId || null;
    el.weeks.querySelectorAll('[data-card-id]').forEach(element => {
        element.classList.toggle('is-highlighted', element.dataset.cardId === state.highlightedCardId);
    });
}

// Every second, and at once on returning to the tab: what depends on the clock and not on Graph.
// What a render decides once (which day is today, which days have begun, a card's "over 16 h")
// is drawn again when the clock changes it, without waiting for the next read.
function onLiveTick() {
    if (el.weeks.hidden || !state.weekStart) return;
    const nowMs = Date.now();
    const running = state.cards.filter(card => !card.clockOut);
    const turnedLong = running.some(card =>
        el.weeks.querySelector(`.card[data-card-id="${CSS.escape(card.id)}"]`)?.classList.contains('is-long') === false
        && isLongCard(card, nowMs));
    if (state.renderedDayKey !== weekKey(getDayStart(nowMs)) || turnedLong) return render();
    updateNowLines(el.weeks, nowMs);
    if (!running.length) return;
    const days = buildWeekDays(state.cards, state.weekStart, nowMs);
    // A running card past the end of a narrow screen's window widens it.
    if (hoursKey(timelineHours(days, nowMs)) !== state.renderedHoursKey) return render();
    renderTotals(days);
    running.forEach(card => {
        const worked = el.weeks.querySelector(`[data-worked="${CSS.escape(card.id)}"]`);
        if (worked) worked.textContent = formatDurationHms(workedMs(card, nowMs));
    });
    updateRunningSpans(el.weeks, new Map(running.map(card => [card.id, card])), nowMs);
}

// ── Actions ─────────────────────────────────────────────────────────────────

function onWeeksClick(event) {
    const cardElement = event.target.closest('[data-card-id]');
    if (!cardElement) return;
    highlightCard(cardElement.dataset.cardId);
    const button = event.target.closest('button[data-action]');
    const card = state.cards.find(c => c.id === cardElement.dataset.cardId);
    if (!button || !card) return;
    const breakIndex = Number(button.closest('[data-break-index]')?.dataset.breakIndex);
    const actions = {
        extend: () => extendToNow(card),
        edit: () => openEditDialog(card),
        delete: () => deleteCard(card),
        'delete-break': () => deleteBreak(card, breakIndex),
        clockOut: () => clockAction('clockOut', button),
        startBreak: () => clockAction('startBreak', button),
        endBreak: () => clockAction('endBreak', button),
    };
    void actions[button.dataset.action]?.();
}

// Runs a change, then reads the week back. Until both are done every action is disabled, so no
// second change is built from the cards as they were before this one.
async function change(button, successMessage, request) {
    state.changesInFlight += 1;
    document.body.classList.add('is-saving');
    if (button) button.disabled = true;
    disableActions();
    try {
        await graph(request);
        toast(successMessage, 'success');
        await loadWeek({ force: true });
        return true;
    } catch (error) {
        toast(error.message, 'error');
        return false;
    } finally {
        state.changesInFlight -= 1;
        assert(state.changesInFlight >= 0, 'changes in flight never negative');
        document.body.classList.toggle('is-saving', state.changesInFlight > 0);
        if (button) button.disabled = false;
        render();
    }
}

const CLOCK_ACTIONS = {
    clockIn: { path: () => '/clockIn', done: 'Clocked in' },
    clockOut: { path: id => `/${id}/clockOut`, done: 'Clocked out' },
    startBreak: { path: id => `/${id}/startBreak`, done: 'Break started' },
    endBreak: { path: id => `/${id}/endBreak`, done: 'Break ended' },
};

function clockAction(actionName, button = el.btnClock) {
    const action = CLOCK_ACTIONS[actionName];
    if (!action || !state.team) return;
    const cardId = actionName === 'clockIn' ? '' : state.activeCard?.id;
    if (actionName !== 'clockIn' && !cardId) return;
    return change(button, action.done, () => auth.graphFetch(`${timeCardsPath(state.team.id)}${action.path(cardId)}`, {
        method: 'POST',
        body: JSON.stringify({ atApprovedLocation: false, notes: { contentType: 'text', content: '' } }),
    }));
}

function saveCard(updatedCard, successMessage, button = null) {
    return change(button, successMessage, () => auth.graphFetch(`${timeCardsPath(state.team.id)}/${updatedCard.id}`, {
        method: 'PUT',
        body: JSON.stringify(buildTimeCardUpdateBody(updatedCard)),
    }));
}

function extendToNow(card) {
    const targetMs = Date.now() - EXTEND_TO_NOW_LAG_MS;
    if (targetMs <= new Date(card.clockOut.dateTime).getTime()) return toast('The card already reaches now.');
    return saveCard(buildUpdatedCard(card, { clockOut: new Date(targetMs) }), `Extended to ${timeFormat.format(targetMs)}`);
}

async function deleteCard(card) {
    const confirmed = await askConfirm({
        title: 'Delete timecard',
        message: `Delete the timecard from ${dateTimeFormat.format(new Date(card.clockIn.dateTime))}? This cannot be undone.`,
        beta: true,
    });
    if (confirmed) await change(null, 'Timecard deleted', () => auth.graphFetchBeta(`${timeCardsPath(state.team.id)}/${card.id}`, { method: 'DELETE' }));
}

async function deleteBreak(card, breakIndex) {
    const removed = card.breaks[breakIndex];
    if (!removed) return;
    const confirmed = await askConfirm({
        title: 'Delete break',
        message: `Delete the break from ${timeFormat.format(new Date(removed.start.dateTime))}? This cannot be undone.`,
    });
    if (confirmed) await saveCard(buildUpdatedCard(card, { breaks: card.breaks.filter((_, i) => i !== breakIndex) }), 'Break deleted');
}

function askConfirm({ title, message, beta = false }) {
    el.confirmTitle.textContent = title;
    el.confirmMsg.textContent = message;
    show(el.confirmBetaWarn, beta);
    show(el.confirmModal, true);
    el.btnConfirmOk.focus();
    return new Promise(resolve => { state.confirmResolve = resolve; });
}

function closeConfirm(confirmed) {
    show(el.confirmModal, false);
    state.confirmResolve?.(confirmed);
    state.confirmResolve = null;
}

// ── The edit dialog ─────────────────────────────────────────────────────────

function openEditDialog(card) {
    state.editCard = card;
    el.editClockIn.value = toLocalDatetimeInput(new Date(card.clockIn.dateTime));
    el.editClockOut.value = card.clockOut ? toLocalDatetimeInput(new Date(card.clockOut.dateTime)) : '';
    el.editNotes.value = card.notes?.content || '';
    el.breaksEdit.replaceChildren();
    card.breaks.forEach(b => addBreakRow({
        breakId: b.breakId,
        start: b.start ? toLocalDatetimeInput(new Date(b.start.dateTime)) : '',
        end: b.end ? toLocalDatetimeInput(new Date(b.end.dateTime)) : '',
    }));
    show(el.editModal, true);
    el.editClockIn.focus();
}

function closeEditDialog() {
    state.editCard = null;
    show(el.editModal, false);
}

// The dialog edits and removes a card's breaks; new ones are added on the timeline.
function addBreakRow({ breakId = '', start = '', end = '' }) {
    const row = document.createElement('div');
    row.className = 'break-edit-row';
    row.dataset.breakId = breakId;
    row.innerHTML = `
        <label>Start <input type="datetime-local" class="break-start" value="${escHtml(start)}" /></label>
        <label>End <input type="datetime-local" class="break-end" value="${escHtml(end)}" /></label>
        <button type="button" class="btn btn-secondary btn-compact btn-icon" title="Remove break" aria-label="Remove break">✕</button>`;
    row.querySelector('button').addEventListener('click', () => row.remove());
    el.breaksEdit.appendChild(row);
}

async function saveEditDialog() {
    const card = state.editCard;
    if (!card) return;
    const result = readEditForm(card, {
        clockIn: el.editClockIn.value,
        clockOut: el.editClockOut.value,
        notes: el.editNotes.value,
        breaks: [...el.breaksEdit.children].map(row => ({
            breakId: row.dataset.breakId,
            start: row.querySelector('.break-start').value,
            end: row.querySelector('.break-end').value,
        })),
    });
    if (result.error) return toast(result.error, 'error');
    el.btnEditSave.textContent = 'Saving…';
    const saved = await saveCard(buildUpdatedCard(card, result.changes), 'Timecard saved', el.btnEditSave);
    el.btnEditSave.textContent = 'Save';
    el.btnEditSave.disabled = false;
    if (saved) closeEditDialog();
}

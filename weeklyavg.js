import { showError } from './ui.js';

const STORAGE_KEY_START_DATE = 'tc_weekly_avg_start_date';
const STORAGE_KEY_WEEKLY_HOURS = 'tc_weekly_avg_weekly_hours';
const STORAGE_KEY_WORKDAYS = 'tc_weekly_avg_workdays';
const DEFAULT_WEEKLY_HOURS = 25;
const DAY_MS = 24 * 60 * 60 * 1000;
const HOUR_MS = 60 * 60 * 1000;
const DEFAULT_WORKDAYS = [1, 2, 3, 4, 5];
const WORKDAY_OPTIONS = [
    { value: 0, shortLabel: 'S', fullLabel: 'Sunday' },
    { value: 1, shortLabel: 'M', fullLabel: 'Monday' },
    { value: 2, shortLabel: 'T', fullLabel: 'Tuesday' },
    { value: 3, shortLabel: 'W', fullLabel: 'Wednesday' },
    { value: 4, shortLabel: 'T', fullLabel: 'Thursday' },
    { value: 5, shortLabel: 'F', fullLabel: 'Friday' },
    { value: 6, shortLabel: 'S', fullLabel: 'Saturday' },
];

let fetchTimeCardsForDateRange = null;
let formatDateInputValue = null;
let formatDurationHms = null;
let getCurrentAppState = null;
let getWeekRange = null;

let triggerButton = null;
let modalOverlay = null;
let startInput = null;
let hoursInput = null;
let workdaysFieldset = null;
let headlineEl = null;
let metaEl = null;
let loadingEl = null;
let emptyEl = null;
let tableWrapEl = null;
let tableBodyEl = null;
let refreshToken = 0;
let cardsCache = { key: '', cards: [] };
let pendingCardsKey = '';
let pendingCardsPromise = null;
let isOpen = false;
let eventsBound = false;

const shortDateFmt = new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric' });
const longDateFmt = new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
const hoursFmt = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 0 });

export function initWeeklyAverageFeature(appApi) {
    fetchTimeCardsForDateRange = appApi?.fetchTimeCardsForDateRange || fetchTimeCardsForDateRange;
    formatDateInputValue = appApi?.formatDateInputValue || formatDateInputValue;
    formatDurationHms = appApi?.formatDurationHms || formatDurationHms;
    getCurrentAppState = appApi?.getCurrentAppState || getCurrentAppState;
    getWeekRange = appApi?.getWeekRange || getWeekRange;
    if (!fetchTimeCardsForDateRange || !formatDateInputValue || !formatDurationHms || !getCurrentAppState || !getWeekRange) {
        throw new Error('Weekly average app API not initialized.');
    }
    if (!triggerButton || !modalOverlay) buildFeatureUi();
    syncInputsFromSettings();
    setWeeklyAverageTriggerVisible(Boolean(getCurrentAppState()?.team?.id));
    bindEventsOnce();
}

export function setWeeklyAverageTriggerVisible(visible) {
    if (!triggerButton) return;
    triggerButton.style.display = visible ? 'inline-flex' : 'none';
    triggerButton.disabled = !visible;
    if (!visible && isOpen) closeModal();
}

function buildFeatureUi() {
    const headerActions = document.querySelector('#header-actions');
    if (!headerActions || !document.body) return;

    document.querySelectorAll('.weekly-avg-trigger').forEach(element => element.remove());
    document.querySelectorAll('#weekly-avg-modal').forEach(element => element.remove());

    triggerButton = document.createElement('button');
    triggerButton.type = 'button';
    triggerButton.className = 'btn btn-secondary weekly-avg-trigger';
    triggerButton.textContent = 'Weekly Average';
    triggerButton.setAttribute('aria-haspopup', 'dialog');
    triggerButton.setAttribute('aria-expanded', 'false');
    triggerButton.style.display = 'none';
    headerActions.appendChild(triggerButton);

    modalOverlay = document.createElement('div');
    modalOverlay.className = 'modal-overlay hidden';
    modalOverlay.id = 'weekly-avg-modal';
    modalOverlay.innerHTML = `
        <div class="modal weekly-avg-modal-shell" role="dialog" aria-modal="true" aria-labelledby="weekly-avg-title">
            <h2 id="weekly-avg-title">Weekly average</h2>
            <div class="weekly-avg-header">
                <div>
                    <p class="weekly-avg-headline">Weekly average</p>
                    <p class="weekly-avg-meta"></p>
                </div>
                <p class="weekly-avg-note">Expected time is split evenly across the selected workdays from the chosen start date. Days left off the schedule count as 00:00:00. Difference columns are actual minus expected, so negative means behind.</p>
            </div>
            <div class="weekly-avg-toolbar">
                <label class="weekly-avg-field">
                    <span>Avg from</span>
                    <input type="date" id="weekly-avg-start-date" aria-label="Weekly average start date" />
                </label>
                <label class="weekly-avg-field">
                    <span>Hours/week</span>
                    <input type="number" id="weekly-avg-hours" min="0" step="0.25" inputmode="decimal" aria-label="Weekly target hours" />
                </label>
                <fieldset class="weekly-avg-workdays" aria-label="Weekly average workdays">
                    <legend>Workdays</legend>
                    <div class="weekly-avg-workdays-grid">
                        ${WORKDAY_OPTIONS.map(option => `
                            <label class="weekly-avg-day-chip" title="${option.fullLabel}">
                                <input type="checkbox" value="${option.value}" data-day-label="${option.fullLabel}" />
                                <span aria-hidden="true">${option.shortLabel}</span>
                                <span class="sr-only">${option.fullLabel}</span>
                            </label>
                        `).join('')}
                    </div>
                </fieldset>
            </div>
            <div class="weekly-avg-loading hidden-start"><span class="spinner"></span>Loading weekly average...</div>
            <div class="weekly-avg-empty hidden-start"></div>
            <div class="weekly-avg-table-wrap hidden-start">
                <table class="weekly-avg-table">
                    <thead>
                        <tr>
                            <th>Week</th>
                            <th>Expected/day</th>
                            <th>Expected week</th>
                            <th>Actual week</th>
                            <th>Weekly</th>
                            <th>Cum.</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
            <div class="modal-actions">
                <button type="button" class="btn btn-secondary" id="btn-weekly-avg-close">Close</button>
            </div>
        </div>
    `;
    document.body.appendChild(modalOverlay);

    startInput = modalOverlay.querySelector('#weekly-avg-start-date');
    hoursInput = modalOverlay.querySelector('#weekly-avg-hours');
    workdaysFieldset = modalOverlay.querySelector('.weekly-avg-workdays');
    headlineEl = modalOverlay.querySelector('.weekly-avg-headline');
    metaEl = modalOverlay.querySelector('.weekly-avg-meta');
    loadingEl = modalOverlay.querySelector('.weekly-avg-loading');
    emptyEl = modalOverlay.querySelector('.weekly-avg-empty');
    tableWrapEl = modalOverlay.querySelector('.weekly-avg-table-wrap');
    tableBodyEl = modalOverlay.querySelector('tbody');
}

function bindEventsOnce() {
    if (eventsBound) return;
    eventsBound = true;

    triggerButton.addEventListener('click', openModal);
    modalOverlay.addEventListener('click', event => {
        if (event.target === modalOverlay) closeModal();
    });
    modalOverlay.querySelector('#btn-weekly-avg-close').addEventListener('click', closeModal);
    startInput.addEventListener('change', onSettingsChange);
    hoursInput.addEventListener('change', onSettingsChange);
    workdaysFieldset.addEventListener('change', onSettingsChange);
    document.addEventListener('keydown', onDocumentKeyDown);
}

function onDocumentKeyDown(event) {
    if (event.key === 'Escape' && isOpen) closeModal();
}

function openModal() {
    const appState = getCurrentAppState();
    if (!appState?.team?.id) return;
    isOpen = true;
    triggerButton.setAttribute('aria-expanded', 'true');
    syncInputsFromSettings();
    modalOverlay.classList.remove('hidden');
    startInput.focus();
    void refreshFromAppState(appState);
}

function closeModal() {
    if (!modalOverlay || !isOpen) return;
    isOpen = false;
    triggerButton?.setAttribute('aria-expanded', 'false');
    modalOverlay.classList.add('hidden');
    triggerButton?.focus();
}

function onSettingsChange() {
    const startDate = parseDateInput(startInput.value);
    const weeklyHours = parseWeeklyHours(hoursInput.value);
    const workdays = readSelectedWorkdaysFromInputs();
    if (!startDate || weeklyHours === null) {
        syncInputsFromSettings();
        return;
    }
    persistSettings(startDate, weeklyHours, workdays);
    syncInputsFromSettings();
    if (isOpen) void refreshFromAppState(getCurrentAppState());
}

function getDefaultStartDate() {
    const now = new Date();
    return new Date(now.getFullYear(), 0, 1);
}

function startOfDay(date) {
    const value = new Date(date);
    value.setHours(0, 0, 0, 0);
    return value;
}

function addCalendarDays(date, days) {
    const value = startOfDay(date);
    value.setDate(value.getDate() + days);
    return value;
}

function parseDateInput(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return null;
    const parsed = new Date(`${value}T00:00:00`);
    return Number.isNaN(parsed.getTime()) ? null : startOfDay(parsed);
}

function parseWeeklyHours(value) {
    const parsed = Number.parseFloat(value);
    if (!Number.isFinite(parsed) || parsed < 0) return null;
    return Number(parsed.toFixed(2));
}

function parseStoredWorkdays(value) {
    if (typeof value !== 'string' || !value.trim()) return null;
    const values = value.split(',')
        .map(item => Number.parseInt(item, 10))
        .filter(day => Number.isInteger(day) && day >= 0 && day <= 6);
    return normalizeWorkdays(values);
}

function readSelectedWorkdaysFromInputs() {
    if (!workdaysFieldset) return DEFAULT_WORKDAYS.slice();
    const values = Array.from(workdaysFieldset.querySelectorAll('input[type="checkbox"]:checked'))
        .map(input => Number.parseInt(input.value, 10));
    return normalizeWorkdays(values);
}

function normalizeWorkdays(values) {
    const uniqueValues = Array.from(new Set((values || []).filter(day => Number.isInteger(day) && day >= 0 && day <= 6)));
    const orderedValues = WORKDAY_OPTIONS
        .map(option => option.value)
        .filter(day => uniqueValues.includes(day));
    return orderedValues.length ? orderedValues : DEFAULT_WORKDAYS.slice();
}

function countWorkdaysPerWeek(workdays) {
    return Math.max(1, workdays.length);
}

function isWorkday(day, workdaysSet) {
    return workdaysSet.has(day.getDay());
}

function describeWorkdays(workdays) {
    const labels = WORKDAY_OPTIONS
        .filter(option => workdays.includes(option.value))
        .map(option => option.fullLabel);
    return labels.length ? labels.join(', ') : 'Monday, Tuesday, Wednesday, Thursday, Friday';
}

function formatWeeklyHours(value) {
    return hoursFmt.format(value);
}

function readSettings() {
    const startDate = parseDateInput(localStorage.getItem(STORAGE_KEY_START_DATE) || '') || getDefaultStartDate();
    const weeklyHours = parseWeeklyHours(localStorage.getItem(STORAGE_KEY_WEEKLY_HOURS) || '') ?? DEFAULT_WEEKLY_HOURS;
    const workdays = parseStoredWorkdays(localStorage.getItem(STORAGE_KEY_WORKDAYS)) || DEFAULT_WORKDAYS.slice();
    return {
        startDate,
        startKey: formatDateInputValue(startDate),
        weeklyHours,
        workdays,
        workdaysSet: new Set(workdays),
        workdaysDescription: describeWorkdays(workdays),
        dailyTargetMs: (weeklyHours * HOUR_MS) / countWorkdaysPerWeek(workdays),
    };
}

function persistSettings(startDate, weeklyHours, workdays) {
    localStorage.setItem(STORAGE_KEY_START_DATE, formatDateInputValue(startDate));
    localStorage.setItem(STORAGE_KEY_WEEKLY_HOURS, String(weeklyHours));
    localStorage.setItem(STORAGE_KEY_WORKDAYS, normalizeWorkdays(workdays).join(','));
}

function syncInputsFromSettings() {
    const settings = readSettings();
    if (startInput) startInput.value = settings.startKey;
    if (hoursInput) hoursInput.value = String(settings.weeklyHours);
    if (workdaysFieldset) {
        const selected = new Set(settings.workdays);
        workdaysFieldset.querySelectorAll('input[type="checkbox"]').forEach(input => {
            const isChecked = selected.has(Number.parseInt(input.value, 10));
            input.checked = isChecked;
            input.parentElement?.classList.toggle('is-selected', isChecked);
        });
    }
}

function setModalState({ headline = 'Weekly average', meta = '', empty = '', loading = false, showTable = false } = {}) {
    headlineEl.textContent = headline;
    metaEl.textContent = meta;
    loadingEl.style.display = loading ? 'block' : 'none';
    emptyEl.textContent = empty;
    emptyEl.style.display = empty ? 'block' : 'none';
    tableWrapEl.style.display = showTable ? 'block' : 'none';
}

async function refreshFromAppState(appState) {
    if (!modalOverlay || !isOpen) return;

    const settings = readSettings();
    const today = new Date();
    const todayStart = startOfDay(today);
    const currentWeekStart = getWeekRange(today).weekStart;

    if (settings.startDate.getTime() > todayStart.getTime()) {
        tableBodyEl.innerHTML = '';
        setModalState({
            headline: 'Weekly average',
            meta: `Target ${formatWeeklyHours(settings.weeklyHours)} h/week across ${settings.workdaysDescription} from ${longDateFmt.format(settings.startDate)}.`,
            empty: 'Choose a start date on or before today to calculate the running weekly average.',
        });
        return;
    }

    const team = appState?.team || null;
    if (!team?.id) {
        tableBodyEl.innerHTML = '';
        setModalState({
            headline: 'Weekly average',
            meta: `Target ${formatWeeklyHours(settings.weeklyHours)} h/week across ${settings.workdaysDescription} from ${longDateFmt.format(settings.startDate)}.`,
            empty: 'Choose a team to compare actual hours against the weekly target.',
        });
        return;
    }

    const refreshId = ++refreshToken;
    const requestStart = getWeekRange(settings.startDate).weekStart;
    setModalState({
        headline: 'Weekly average',
        meta: `Target ${formatWeeklyHours(settings.weeklyHours)} h/week across ${settings.workdaysDescription} from ${longDateFmt.format(settings.startDate)} for ${team.displayName}.`,
        loading: true,
    });

    try {
        const cards = await getCardsForRange(team.id, requestStart, today, appState?.dataVersion || 0);
        if (refreshId !== refreshToken || !isOpen) return;
        const model = buildWeeklyModel(cards, settings, currentWeekStart, today);
        renderWeeklyModel(team.displayName, settings, model);
    } catch (error) {
        if (refreshId !== refreshToken || !isOpen) return;
        tableBodyEl.innerHTML = '';
        showError('Failed to load weekly average: ' + error.message);
        setModalState({
            headline: 'Weekly average',
            meta: `Target ${formatWeeklyHours(settings.weeklyHours)} h/week across ${settings.workdaysDescription} from ${longDateFmt.format(settings.startDate)} for ${team.displayName}.`,
            empty: 'Could not load weekly average data right now.',
        });
    }
}

async function getCardsForRange(teamId, requestStart, requestEnd, dataVersion) {
    const requestKey = [
        teamId,
        formatDateInputValue(requestStart),
        formatDateInputValue(startOfDay(requestEnd)),
        String(dataVersion || 0),
    ].join(':');

    if (cardsCache.key === requestKey) return cardsCache.cards;
    if (pendingCardsKey === requestKey && pendingCardsPromise) return pendingCardsPromise;

    pendingCardsKey = requestKey;
    pendingCardsPromise = fetchTimeCardsForDateRange(teamId, requestStart, requestEnd)
        .then(result => {
            cardsCache = { key: requestKey, cards: result.cards || [] };
            return cardsCache.cards;
        })
        .finally(() => {
            if (pendingCardsKey === requestKey) {
                pendingCardsKey = '';
                pendingCardsPromise = null;
            }
        });
    return pendingCardsPromise;
}

function buildWeeklyModel(cards, settings, currentWeekStart, now) {
    const rangeStartMs = settings.startDate.getTime();
    const rangeEndMs = now.getTime();
    const dailyActuals = buildDailyActuals(cards, rangeStartMs, rangeEndMs);
    const currentWeekKey = formatDateInputValue(currentWeekStart);
    const todayStart = startOfDay(now).getTime();
    const rows = [];
    let expectedCumulativeMs = 0;
    let actualCumulativeMs = 0;
    const firstWeekStart = getWeekRange(settings.startDate).weekStart;

    for (let weekStart = new Date(firstWeekStart); weekStart.getTime() <= currentWeekStart.getTime(); weekStart = addCalendarDays(weekStart, 7)) {
        const weekStartMs = weekStart.getTime();
        const weekEnd = addCalendarDays(weekStart, 6);
        let expectedWeekMs = 0;
        let actualWeekMs = 0;

        for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
            const day = addCalendarDays(weekStart, dayIndex);
            const dayMs = day.getTime();
            const dayKey = formatDateInputValue(day);
            const inScope = dayMs >= rangeStartMs && dayMs <= todayStart;
            const expectedDayMs = inScope ? getExpectedDayMs(day, settings) : 0;
            expectedWeekMs += expectedDayMs;
            actualWeekMs += dailyActuals.get(dayKey) || 0;
        }

        expectedCumulativeMs += expectedWeekMs;
        actualCumulativeMs += actualWeekMs;
        rows.push({
            weekStart,
            weekEnd,
            expectedDailyMs: settings.dailyTargetMs,
            expectedWeekMs,
            actualWeekMs,
            weeklyDifferenceMs: actualWeekMs - expectedWeekMs,
            expectedCumulativeMs,
            actualCumulativeMs,
            differenceMs: actualCumulativeMs - expectedCumulativeMs,
            isCurrentWeek: formatDateInputValue(weekStart) === currentWeekKey,
            isPartialWeek: expectedWeekMs < (settings.dailyTargetMs * settings.workdays.length),
        });
    }

    return {
        rows: rows.reverse(),
        expectedCumulativeMs,
        actualCumulativeMs,
        differenceMs: actualCumulativeMs - expectedCumulativeMs,
    };
}

function getExpectedDayMs(day, settings) {
    return isWorkday(day, settings.workdaysSet) ? settings.dailyTargetMs : 0;
}

function buildDailyActuals(cards, rangeStartMs, rangeEndMs) {
    const totals = new Map();
    cards.forEach(card => {
        getWorkSegments(card, rangeEndMs).forEach(([segmentStartMs, segmentEndMs]) => {
            const clippedStartMs = Math.max(segmentStartMs, rangeStartMs);
            const clippedEndMs = Math.min(segmentEndMs, rangeEndMs);
            if (clippedEndMs <= clippedStartMs) return;

            let cursorMs = clippedStartMs;
            while (cursorMs < clippedEndMs) {
                const dayStart = startOfDay(cursorMs);
                const nextDayMs = dayStart.getTime() + DAY_MS;
                const sliceEndMs = Math.min(clippedEndMs, nextDayMs);
                const dayKey = formatDateInputValue(dayStart);
                totals.set(dayKey, (totals.get(dayKey) || 0) + (sliceEndMs - cursorMs));
                cursorMs = sliceEndMs;
            }
        });
    });
    return totals;
}

function getWorkSegments(card, referenceEndMs) {
    if (!card?.clockIn?.dateTime) return [];
    const startMs = new Date(card.clockIn.dateTime).getTime();
    const endMs = card.clockOut?.dateTime ? new Date(card.clockOut.dateTime).getTime() : referenceEndMs;
    if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs <= startMs) return [];

    const breaks = (card.breaks || [])
        .filter(item => item?.start?.dateTime)
        .map(item => ({
            startMs: new Date(item.start.dateTime).getTime(),
            endMs: item.end?.dateTime ? new Date(item.end.dateTime).getTime() : referenceEndMs,
        }))
        .filter(item => Number.isFinite(item.startMs))
        .sort((a, b) => a.startMs - b.startMs);

    const segments = [];
    let cursorMs = startMs;
    breaks.forEach(item => {
        const breakStartMs = Math.min(Math.max(item.startMs, startMs), endMs);
        const breakEndMs = Math.min(Math.max(item.endMs, breakStartMs), endMs);
        if (breakStartMs > cursorMs) segments.push([cursorMs, breakStartMs]);
        cursorMs = Math.max(cursorMs, breakEndMs);
    });
    if (cursorMs < endMs) segments.push([cursorMs, endMs]);
    return segments.filter(([segmentStartMs, segmentEndMs]) => segmentEndMs > segmentStartMs);
}

function renderWeeklyModel(teamName, settings, model) {
    tableBodyEl.innerHTML = '';
    model.rows.forEach(row => tableBodyEl.appendChild(buildRow(row)));
    setModalState({
        headline: describeDifference(model.differenceMs),
        meta: `Actual ${formatDurationHms(model.actualCumulativeMs)} vs expected ${formatDurationHms(model.expectedCumulativeMs)} for ${teamName}, starting ${longDateFmt.format(settings.startDate)} at ${formatWeeklyHours(settings.weeklyHours)} h/week across ${settings.workdaysDescription}.`,
        showTable: true,
    });
}

function describeDifference(differenceMs) {
    if (Math.abs(differenceMs) < 1000) return 'On target';
    const direction = differenceMs < 0 ? 'behind target' : 'ahead of target';
    return `${formatDurationHms(Math.abs(differenceMs))} ${direction}`;
}

function buildRow(row) {
    const tr = document.createElement('tr');
    const weeklyDiffClass = getDifferenceClass(row.weeklyDifferenceMs);
    const cumulativeDiffClass = getDifferenceClass(row.differenceMs);
    const weekState = row.isCurrentWeek ? 'Current week' : row.isPartialWeek ? 'Partial week' : '';
    tr.className = row.isCurrentWeek ? 'weekly-avg-row-current' : '';
    tr.innerHTML = `
        <td>
            <div class="weekly-avg-week-label">${shortDateFmt.format(row.weekStart)} - ${shortDateFmt.format(row.weekEnd)}</div>
            ${weekState ? `<div class="weekly-avg-week-state">${weekState}</div>` : ''}
        </td>
        <td class="weekly-avg-mono">${formatDurationHms(row.expectedDailyMs)}</td>
        <td class="weekly-avg-mono">${formatDurationHms(row.expectedWeekMs)}</td>
        <td class="weekly-avg-mono">${formatDurationHms(row.actualWeekMs)}</td>
        <td class="weekly-avg-mono ${weeklyDiffClass}">${formatSignedDuration(row.weeklyDifferenceMs)}</td>
        <td class="weekly-avg-mono ${cumulativeDiffClass}">${formatSignedDuration(row.differenceMs)}</td>
    `;
    return tr;
}

function getDifferenceClass(durationMs) {
    return durationMs < -999
        ? 'weekly-avg-diff-negative'
        : durationMs > 999
            ? 'weekly-avg-diff-positive'
            : 'weekly-avg-diff-neutral';
}

function formatSignedDuration(durationMs) {
    const prefix = durationMs < 0 ? '-' : durationMs > 0 ? '+' : '';
    return `${prefix}${formatDurationHms(Math.abs(durationMs))}`;
}

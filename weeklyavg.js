import { showError } from './ui.js';
import {
    buildWeeklyModel, formatDateInputValue, formatDurationHms, formatSignedDuration, getDayStart, getWeekStart, parseLocalDate,
} from './core.js';

const STORAGE_KEY_START_DATE = 'tc_weekly_avg_start_date';
const STORAGE_KEY_WEEKLY_HOURS = 'tc_weekly_avg_weekly_hours';
const STORAGE_KEY_WORKDAYS = 'tc_weekly_avg_workdays';
const DEFAULT_WEEKLY_HOURS = 25;
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
let getCurrentAppState = null;

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
let previousBodyOverflow = '';
let previousDocumentOverflow = '';

const shortDateFmt = new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric' });
const weekdayFmt = new Intl.DateTimeFormat(undefined, { weekday: 'short' });
const longDateFmt = new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
const hoursFmt = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2, minimumFractionDigits: 0 });

export function initWeeklyAverageFeature(appApi) {
    ({ fetchTimeCardsForDateRange, getCurrentAppState } = appApi);
    if (!fetchTimeCardsForDateRange || !getCurrentAppState) throw new Error('Weekly average app API not initialized.');
    if (!triggerButton) {
        buildFeatureUi();
        bindEvents();
    }
    syncInputsFromSettings();
    setWeeklyAverageTriggerVisible(Boolean(getCurrentAppState()?.team?.id));
}

export function setWeeklyAverageTriggerVisible(visible) {
    if (!triggerButton) return;
    triggerButton.hidden = !visible;
    triggerButton.disabled = !visible;
    if (!visible && isOpen) closeModal();
}

function buildFeatureUi() {
    triggerButton = document.createElement('button');
    triggerButton.type = 'button';
    triggerButton.className = 'btn btn-secondary weekly-avg-trigger';
    triggerButton.textContent = 'Weekly Average';
    triggerButton.setAttribute('aria-haspopup', 'dialog');
    triggerButton.setAttribute('aria-expanded', 'false');
    triggerButton.hidden = true;
    document.querySelector('#header-actions').appendChild(triggerButton);

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
                <p class="weekly-avg-note">Expected time is split evenly across the selected workdays from the chosen start date, and counts only through today. Days left off the schedule count as 00:00:00. Difference columns are actual minus expected, so negative means behind.</p>
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
            <div class="weekly-avg-loading" hidden><span class="spinner"></span>Loading weekly average…</div>
            <div class="weekly-avg-empty" hidden></div>
            <div class="weekly-avg-table-wrap" hidden>
                <table class="weekly-avg-table">
                    <thead>
                        <tr>
                            <th></th>
                            <th>Expected</th>
                            <th>Actual</th>
                            <th title="Difference this week or day">±Δ</th>
                            <th title="Running difference">Σ</th>
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

function bindEvents() {
    triggerButton.addEventListener('click', openModal);
    modalOverlay.addEventListener('click', event => {
        if (event.target === modalOverlay) closeModal();
    });
    modalOverlay.querySelector('#btn-weekly-avg-close').addEventListener('click', closeModal);
    startInput.addEventListener('change', onSettingsChange);
    hoursInput.addEventListener('change', onSettingsChange);
    workdaysFieldset.addEventListener('change', onSettingsChange);
    document.addEventListener('keydown', event => {
        if (event.key === 'Escape' && isOpen) closeModal();
    });
}

function openModal() {
    const appState = getCurrentAppState();
    if (!appState?.team?.id) return;
    isOpen = true;
    triggerButton.setAttribute('aria-expanded', 'true');
    syncInputsFromSettings();
    previousBodyOverflow = document.body.style.overflow;
    previousDocumentOverflow = document.documentElement.style.overflow;
    document.body.style.overflow = 'hidden';
    document.documentElement.style.overflow = 'hidden';
    modalOverlay.classList.remove('hidden');
    startInput.focus();
    void refreshFromAppState(appState);
}

function closeModal() {
    if (!isOpen) return;
    isOpen = false;
    triggerButton.setAttribute('aria-expanded', 'false');
    modalOverlay.classList.add('hidden');
    document.body.style.overflow = previousBodyOverflow;
    document.documentElement.style.overflow = previousDocumentOverflow;
    triggerButton.focus();
}

function onSettingsChange() {
    const startDate = parseDateInput(startInput.value);
    const weeklyHours = parseWeeklyHours(hoursInput.value);
    if (startDate && weeklyHours !== null) {
        localStorage.setItem(STORAGE_KEY_START_DATE, formatDateInputValue(startDate));
        localStorage.setItem(STORAGE_KEY_WEEKLY_HOURS, String(weeklyHours));
        localStorage.setItem(STORAGE_KEY_WORKDAYS, readSelectedWorkdaysFromInputs().join(','));
    }
    syncInputsFromSettings();
    if (isOpen) void refreshFromAppState(getCurrentAppState());
}

function parseDateInput(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return null;
    return parseLocalDate(value);
}

function parseWeeklyHours(value) {
    const parsed = Number.parseFloat(value);
    if (!Number.isFinite(parsed) || parsed < 0) return null;
    return Number(parsed.toFixed(2));
}

function normalizeWorkdays(values) {
    const ordered = WORKDAY_OPTIONS.map(option => option.value).filter(day => values.includes(day));
    return ordered.length ? ordered : DEFAULT_WORKDAYS.slice();
}

function readSelectedWorkdaysFromInputs() {
    return normalizeWorkdays(Array.from(workdaysFieldset.querySelectorAll('input[type="checkbox"]:checked'))
        .map(input => Number.parseInt(input.value, 10)));
}

function readSettings() {
    const startDate = parseDateInput(localStorage.getItem(STORAGE_KEY_START_DATE) || '') || new Date(new Date().getFullYear(), 0, 1);
    const weeklyHours = parseWeeklyHours(localStorage.getItem(STORAGE_KEY_WEEKLY_HOURS) || '') ?? DEFAULT_WEEKLY_HOURS;
    const storedWorkdays = (localStorage.getItem(STORAGE_KEY_WORKDAYS) || '').split(',').map(item => Number.parseInt(item, 10));
    const workdays = normalizeWorkdays(storedWorkdays);
    const workdaysDescription = WORKDAY_OPTIONS.filter(option => workdays.includes(option.value)).map(option => option.fullLabel).join(', ');
    return {
        startDate,
        weeklyHours,
        workdays,
        description: `${hoursFmt.format(weeklyHours)} h/week across ${workdaysDescription} from ${longDateFmt.format(startDate)}`,
    };
}

function syncInputsFromSettings() {
    const settings = readSettings();
    startInput.value = formatDateInputValue(settings.startDate);
    hoursInput.value = String(settings.weeklyHours);
    workdaysFieldset.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.checked = settings.workdays.includes(Number.parseInt(input.value, 10));
        input.parentElement.classList.toggle('is-selected', input.checked);
    });
}

function setModalState({ headline = 'Weekly average', meta = '', empty = '', loading = false, showTable = false } = {}) {
    headlineEl.textContent = headline;
    metaEl.textContent = meta;
    loadingEl.hidden = !loading;
    emptyEl.textContent = empty;
    emptyEl.hidden = !empty;
    tableWrapEl.hidden = !showTable;
    if (!showTable) tableBodyEl.innerHTML = '';
}

async function refreshFromAppState(appState) {
    if (!isOpen) return;
    const settings = readSettings();
    const now = new Date();
    const team = appState?.team;

    if (settings.startDate > getDayStart(now)) {
        setModalState({ meta: `Target ${settings.description}.`, empty: 'Choose a start date on or before today to calculate the running weekly average.' });
        return;
    }
    if (!team?.id) {
        setModalState({ meta: `Target ${settings.description}.`, empty: 'Choose a team to compare actual hours against the weekly target.' });
        return;
    }

    const refreshId = ++refreshToken;
    const meta = `Target ${settings.description} for ${team.displayName}.`;
    setModalState({ meta, loading: true });
    try {
        const cards = await getCardsForRange(team.id, getWeekStart(settings.startDate), now, appState.dataVersion || 0);
        if (refreshId !== refreshToken || !isOpen) return;
        renderWeeklyModel(team.displayName, settings, buildWeeklyModel(cards, settings, now));
    } catch (error) {
        if (refreshId !== refreshToken || !isOpen) return;
        showError('Failed to load weekly average: ' + error.message);
        setModalState({ meta, empty: 'Could not load weekly average data right now.' });
    }
}

async function getCardsForRange(teamId, requestStart, requestEnd, dataVersion) {
    const requestKey = [teamId, formatDateInputValue(requestStart), formatDateInputValue(requestEnd), dataVersion].join(':');
    if (cardsCache.key === requestKey) return cardsCache.cards;
    if (pendingCardsKey === requestKey && pendingCardsPromise) return pendingCardsPromise;

    pendingCardsKey = requestKey;
    pendingCardsPromise = fetchTimeCardsForDateRange(teamId, requestStart, requestEnd)
        .then(result => {
            cardsCache = { key: requestKey, cards: result.cards };
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

function renderWeeklyModel(teamName, settings, model) {
    tableBodyEl.innerHTML = '';
    model.rows.forEach((row, index) => tableBodyEl.appendChild(
        row.rowType === 'day' ? buildDayRow(row, model.rows[index - 1]?.rowType === 'week') : buildWeekRow(row),
    ));
    setModalState({
        headline: describeDifference(model.differenceMs),
        meta: `Actual ${formatDurationHms(model.actualCumulativeMs)} vs expected ${formatDurationHms(model.expectedCumulativeMs)} through today for ${teamName}, at ${settings.description}.`,
        showTable: true,
    });
}

function describeDifference(differenceMs) {
    if (Math.abs(differenceMs) < 1000) return 'On target';
    return `${formatDurationHms(Math.abs(differenceMs))} ${differenceMs < 0 ? 'behind target' : 'ahead of target'}`;
}

function differenceCells(row) {
    const cell = ms => {
        const tone = ms < -999 ? 'negative' : ms > 999 ? 'positive' : 'neutral';
        return `<td class="weekly-avg-mono weekly-avg-mono-signed weekly-avg-diff-${tone}">${formatSignedDuration(ms)}</td>`;
    };
    return cell(row.differenceMs) + cell(row.cumulativeDifferenceMs);
}

function buildWeekRow(row) {
    const tr = document.createElement('tr');
    const weekState = row.isCurrentWeek
        ? `Current week, ${formatDurationHms(row.targetMs)} target`
        : row.isPartialWeek ? 'Partial week' : '';
    tr.className = row.isCurrentWeek ? 'weekly-avg-row-current' : '';
    tr.innerHTML = `
        <td class="weekly-avg-cell-label">
            <div class="weekly-avg-week-label">${shortDateFmt.format(row.weekStart)} - ${shortDateFmt.format(row.weekEnd)}</div>
            ${weekState ? `<div class="weekly-avg-week-state">${weekState}</div>` : ''}
        </td>
        <td class="weekly-avg-mono">${formatDurationHms(row.expectedMs)}</td>
        <td class="weekly-avg-mono">${formatDurationHms(row.actualMs)}</td>
        ${differenceCells(row)}
    `;
    return tr;
}

function buildDayRow(row, isFirst) {
    const tr = document.createElement('tr');
    tr.className = [
        'weekly-avg-row-detail',
        isFirst ? 'weekly-avg-row-detail-first' : '',
        row.isFuture ? 'weekly-avg-row-detail-future' : '',
        row.isWorkday ? '' : 'weekly-avg-row-detail-off',
    ].filter(Boolean).join(' ');
    tr.innerHTML = `
        <td class="weekly-avg-cell-label">
            <div class="weekly-avg-week-label weekly-avg-day-label">${weekdayFmt.format(row.day)}</div>
        </td>
        <td class="weekly-avg-mono">${row.isWorkday ? formatDurationHms(row.targetMs) : '—'}</td>
        <td class="weekly-avg-mono">${formatDurationHms(row.actualMs)}</td>
        ${row.isFuture ? '<td></td><td></td>' : differenceCells(row)}
    `;
    return tr;
}

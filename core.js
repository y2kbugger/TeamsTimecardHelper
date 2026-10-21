// Pure logic: no DOM, no storage, no network. Unit-tested by `node --test` (tests/core.test.js).

export const SNAP_MINUTES = 5;

// ── Calendar ────────────────────────────────────────────────────────────────
// Calendar days, never 24h steps: daylight saving days are 23 or 25 hours long.

export function getDayStart(date) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    return d;
}

export function addDays(date, days) {
    const d = new Date(date);
    d.setDate(d.getDate() + days);
    return d;
}

// Weeks run Sunday to Saturday.
export function getWeekStart(date) {
    const d = getDayStart(date);
    return addDays(d, -d.getDay());
}

export function getWeekRange(date) {
    const weekStart = getWeekStart(date);
    const weekEndExclusive = addDays(weekStart, 7);
    return { weekStart, weekEndExclusive, weekEndInclusive: new Date(weekEndExclusive.getTime() - 1) };
}

export function isSameDay(a, b) {
    return getDayStart(a).getTime() === getDayStart(b).getTime();
}

// 'YYYY-MM-DD' in local time, as <input type="date"> wants it.
export function formatDateInputValue(date) {
    const pad = n => String(n).padStart(2, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

// 'YYYY-MM-DDTHH:MM' in local time, as <input type="datetime-local"> wants it.
export function toLocalDatetimeInput(date) {
    const pad = n => String(n).padStart(2, '0');
    return `${formatDateInputValue(date)}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

// Parses 'YYYY-MM-DD' as local midnight; anything else as Date does. Returns null when invalid.
export function parseLocalDate(value) {
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : new Date(value);
    if (typeof value !== 'string' || !value) return null;
    const parsed = /^\d{4}-\d{2}-\d{2}$/.test(value) ? new Date(`${value}T00:00:00`) : new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function snapMs(ms) {
    const snap = SNAP_MINUTES * 60 * 1000;
    return Math.round(ms / snap) * snap;
}

export function formatDurationHms(durationMs) {
    const totalSeconds = Math.max(0, Math.floor(durationMs / 1000));
    const pad = n => String(n).padStart(2, '0');
    return `${pad(Math.floor(totalSeconds / 3600))}:${pad(Math.floor((totalSeconds % 3600) / 60))}:${pad(totalSeconds % 60)}`;
}

export function formatSignedDuration(durationMs) {
    const prefix = durationMs < 0 ? '-' : durationMs > 0 ? '+' : ' ';
    return `${prefix}${formatDurationHms(Math.abs(durationMs))}`;
}

// ── Timecards ───────────────────────────────────────────────────────────────
// A card is Graph's timeCard reshaped: clockIn/clockOut/breaks[].start/end each hold { dateTime }.

const ms = event => (event?.dateTime ? new Date(event.dateTime).getTime() : NaN);

export function deriveCardState(card) {
    if (card.clockOut) return 'clockedOut';
    if (card.breaks.some(item => item.start && !item.end)) return 'onBreak';
    return 'clockedIn';
}

export function isOpenCard(card) {
    return Boolean(card) && deriveCardState(card) !== 'clockedOut';
}

export function normalizeCard(raw) {
    const event = e => (e ? { dateTime: e.dateTime, atApprovedLocation: e.atApprovedLocation, notes: e.notes } : null);
    const card = {
        id: raw.id,
        userId: raw.userId,
        createdDateTime: raw.createdDateTime,
        lastModifiedDateTime: raw.lastModifiedDateTime,
        clockIn: event(raw.clockInEvent),
        clockOut: event(raw.clockOutEvent),
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

export function cardStartMs(card) {
    return ms(card.clockIn);
}

export function cardEndMs(card, nowMs = Date.now()) {
    return card.clockOut ? ms(card.clockOut) : nowMs;
}

export function cardOverlapsRange(card, rangeStartMs, rangeEndExclusiveMs, nowMs = Date.now()) {
    if (!card?.clockIn?.dateTime) return false;
    return cardStartMs(card) < rangeEndExclusiveMs && cardEndMs(card, nowMs) >= rangeStartMs;
}

// The newest card sorts last: by clock-out, else clock-in, else creation.
export function getCardSortMs(card) {
    if (card.clockOut?.dateTime) return ms(card.clockOut);
    if (card.clockIn?.dateTime) return ms(card.clockIn);
    return card.createdDateTime ? new Date(card.createdDateTime).getTime() : 0;
}

export function dedupeCardsById(cards) {
    return Array.from(new Map(cards.map(card => [card.id, card])).values());
}

export function filterCardsForWeek(cards, weekDate, nowMs = Date.now()) {
    const { weekStart, weekEndExclusive } = getWeekRange(weekDate);
    return dedupeCardsById(cards).filter(card => cardOverlapsRange(card, weekStart.getTime(), weekEndExclusive.getTime(), nowMs));
}

// The open card that started last, if any.
export function findActiveCard(cards) {
    return cards.filter(isOpenCard).sort((a, b) => getCardSortMs(b) - getCardSortMs(a))[0] || null;
}

export function workedMs(card, nowMs = Date.now()) {
    if (!card.clockIn) return 0;
    const breakMs = card.breaks.reduce((sum, b) => (b.start ? sum + ((b.end ? ms(b.end) : nowMs) - ms(b.start)) : sum), 0);
    return Math.max(0, cardEndMs(card, nowMs) - cardStartMs(card) - breakMs);
}

// Cards grouped by the local day they clocked in, newest day first.
export function groupByDay(cards) {
    const days = new Map();
    cards.forEach(card => {
        const anchor = card.clockIn?.dateTime || card.createdDateTime;
        if (!anchor) return;
        const dayStart = getDayStart(new Date(anchor));
        const key = dayStart.getTime();
        if (!days.has(key)) days.set(key, { dayStart, cards: [] });
        days.get(key).cards.push(card);
    });
    return Array.from(days.values()).sort((a, b) => b.dayStart - a.dayStart);
}

export function cloneCardBreaks(breaks) {
    return (breaks || []).map(b => ({
        breakId: b.breakId,
        start: b.start ? { ...b.start } : null,
        end: b.end ? { ...b.end } : null,
    }));
}

function buildEventObj(prior, date) {
    return date ? { ...(prior || {}), dateTime: date.toISOString(), atApprovedLocation: prior?.atApprovedLocation ?? false } : null;
}

// changes: { clockIn?: Date, clockOut?: Date|null, breaks?: [], notes?: string }
export function buildUpdatedCard(card, changes = {}, now = new Date()) {
    const updated = {
        ...card,
        clockIn: changes.clockIn !== undefined ? buildEventObj(card.clockIn, changes.clockIn) : (card.clockIn ? { ...card.clockIn } : null),
        clockOut: changes.clockOut !== undefined ? buildEventObj(card.clockOut, changes.clockOut) : (card.clockOut ? { ...card.clockOut } : null),
        breaks: cloneCardBreaks(changes.breaks ?? card.breaks),
        notes: changes.notes !== undefined
            ? (changes.notes ? { contentType: 'text', content: changes.notes } : null)
            : (card.notes ? { contentType: card.notes.contentType, content: card.notes.content } : null),
        lastModifiedDateTime: now.toISOString(),
    };
    updated.state = deriveCardState(updated);
    return updated;
}

// The body of PUT /teams/{id}/schedule/timeCards/{id}.
export function buildTimeCardUpdateBody(card) {
    const event = e => (e ? { dateTime: e.dateTime, atApprovedLocation: e.atApprovedLocation ?? false } : undefined);
    const body = {
        clockInEvent: event(card.clockIn),
        breaks: card.breaks.map(b => ({
            breakId: b.breakId,
            start: b.start ? { dateTime: b.start.dateTime } : undefined,
            end: b.end ? { dateTime: b.end.dateTime } : undefined,
            notes: { contentType: 'text', content: '' },
        })),
    };
    if (card.clockOut) body.clockOutEvent = event(card.clockOut);
    if (card.notes) body.notes = { ...card.notes };
    return body;
}

// A timeline drag of a card's edges. An open card keeps running: its end is never set.
export function buildProjectedTimelineUpdate(card, newStart, newEnd) {
    const effectiveEnd = card?.clockOut ? newEnd : null;
    const updatedCard = buildUpdatedCard(card, { clockIn: newStart, clockOut: effectiveEnd });
    return { effectiveEnd, updatedCard, requestBody: buildTimeCardUpdateBody(updatedCard) };
}

// ── Weekly average ──────────────────────────────────────────────────────────

export const WEEKDAYS = [0, 1, 2, 3, 4, 5, 6];

// [startMs, endMs) spans of work in a card: clock-in to clock-out minus breaks.
export function getWorkSegments(card, nowMs) {
    if (!card?.clockIn?.dateTime) return [];
    const startMs = cardStartMs(card);
    const endMs = cardEndMs(card, nowMs);
    if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs <= startMs) return [];

    const breaks = card.breaks
        .filter(b => b?.start?.dateTime)
        .map(b => ({ startMs: ms(b.start), endMs: b.end?.dateTime ? ms(b.end) : nowMs }))
        .filter(b => Number.isFinite(b.startMs))
        .sort((a, b) => a.startMs - b.startMs);

    const segments = [];
    let cursorMs = startMs;
    breaks.forEach(b => {
        const breakStartMs = Math.min(Math.max(b.startMs, startMs), endMs);
        const breakEndMs = Math.min(Math.max(b.endMs, breakStartMs), endMs);
        if (breakStartMs > cursorMs) segments.push([cursorMs, breakStartMs]);
        cursorMs = Math.max(cursorMs, breakEndMs);
    });
    if (cursorMs < endMs) segments.push([cursorMs, endMs]);
    return segments;
}

// Worked ms per local day ('YYYY-MM-DD'), split at midnights, within [rangeStartMs, rangeEndMs).
export function buildDailyActuals(cards, rangeStartMs, rangeEndMs) {
    const totals = new Map();
    cards.forEach(card => getWorkSegments(card, rangeEndMs).forEach(([segmentStartMs, segmentEndMs]) => {
        let cursorMs = Math.max(segmentStartMs, rangeStartMs);
        const endMs = Math.min(segmentEndMs, rangeEndMs);
        while (cursorMs < endMs) {
            const dayStart = getDayStart(cursorMs);
            const sliceEndMs = Math.min(endMs, addDays(dayStart, 1).getTime());
            const key = formatDateInputValue(dayStart);
            totals.set(key, (totals.get(key) || 0) + (sliceEndMs - cursorMs));
            cursorMs = sliceEndMs;
        }
    }));
    return totals;
}

// settings: { startDate: Date (local midnight), weeklyHours: number, workdays: number[] (0 = Sunday) }
// Expected time counts only through today: a day that has not come yet is not owed yet.
// Week and day rows are newest first; the current week's days follow its week row.
export function buildWeeklyModel(cards, settings, now = new Date()) {
    const workdays = new Set(settings.workdays);
    const dailyTargetMs = (settings.weeklyHours * 3600 * 1000) / Math.max(1, workdays.size);
    const rangeStartMs = settings.startDate.getTime();
    const todayStartMs = getDayStart(now).getTime();
    const currentWeekStartMs = getWeekStart(now).getTime();
    const dailyActuals = buildDailyActuals(cards, rangeStartMs, now.getTime());

    const weeks = [];
    let expectedCumulativeMs = 0;
    let actualCumulativeMs = 0;
    for (let weekStart = getWeekStart(settings.startDate); weekStart.getTime() <= currentWeekStartMs; weekStart = addDays(weekStart, 7)) {
        const isCurrentWeek = weekStart.getTime() === currentWeekStartMs;
        let expectedWeekMs = 0;
        let actualWeekMs = 0;
        let fullWeekTargetMs = 0;
        const days = [];
        for (let dayIndex = 0; dayIndex < 7; dayIndex += 1) {
            const day = addDays(weekStart, dayIndex);
            const dayMs = day.getTime();
            const isBeforeStart = dayMs < rangeStartMs;
            const isFuture = dayMs > todayStartMs;
            const isWorkday = !isBeforeStart && workdays.has(day.getDay());
            const targetMs = isWorkday ? dailyTargetMs : 0;
            const expectedDayMs = isFuture ? 0 : targetMs;
            const actualDayMs = dailyActuals.get(formatDateInputValue(day)) || 0;
            fullWeekTargetMs += targetMs;
            expectedWeekMs += expectedDayMs;
            actualWeekMs += actualDayMs;
            days.push({
                rowType: 'day',
                day,
                targetMs,
                actualMs: actualDayMs,
                differenceMs: actualDayMs - expectedDayMs,
                cumulativeDifferenceMs: (actualCumulativeMs + actualWeekMs) - (expectedCumulativeMs + expectedWeekMs),
                isFuture,
                isWorkday,
            });
        }
        expectedCumulativeMs += expectedWeekMs;
        actualCumulativeMs += actualWeekMs;
        weeks.push({
            rowType: 'week',
            weekStart,
            weekEnd: addDays(weekStart, 6),
            targetMs: fullWeekTargetMs,
            expectedMs: expectedWeekMs,
            actualMs: actualWeekMs,
            differenceMs: actualWeekMs - expectedWeekMs,
            cumulativeDifferenceMs: actualCumulativeMs - expectedCumulativeMs,
            isCurrentWeek,
            isPartialWeek: fullWeekTargetMs < dailyTargetMs * workdays.size,
            days: isCurrentWeek ? days.reverse() : [],
        });
    }

    return {
        rows: weeks.reverse().flatMap(week => [week, ...week.days]),
        expectedCumulativeMs,
        actualCumulativeMs,
        differenceMs: actualCumulativeMs - expectedCumulativeMs,
    };
}

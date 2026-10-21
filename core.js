// Pure logic: no DOM, no storage, no network. Unit-tested by `node --test` (tests/core.test.js).

export const SNAP_MS = 5 * 60 * 1000;

// For our own bugs: what this code derived itself. Input from Graph or a user is validated instead.
export function assert(condition, message) {
    if (!condition) throw new Error(`Assertion failed: ${message}`);
}

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
    return Math.round(ms / SNAP_MS) * SNAP_MS;
}

export function clamp(value, min, max) {
    assert(min <= max, `clamp bounds ${min} <= ${max}`);
    return Math.min(Math.max(value, min), max);
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

// Longer than any real shift: most likely a clock-out forgotten overnight, to be fixed by hand.
export const CARD_LONG_MS = 16 * 60 * 60 * 1000;

export function isLongCard(card, nowMs = Date.now()) {
    return cardEndMs(card, nowMs) - cardStartMs(card) > CARD_LONG_MS;
}

export function workedMs(card, nowMs = Date.now()) {
    if (!card.clockIn) return 0;
    const breakMs = card.breaks.reduce((sum, b) => (b.start ? sum + ((b.end ? ms(b.end) : nowMs) - ms(b.start)) : sum), 0);
    return Math.max(0, cardEndMs(card, nowMs) - cardStartMs(card) - breakMs);
}

// The days of a week that have anything to show, newest first. A card is listed on the day it
// clocked in and drawn on every day it spans (one past midnight shows on both). Worked time is
// split at midnight, so the day totals add up to the week's.
export function buildWeekDays(cards, weekStart, nowMs = Date.now()) {
    const weekStartMs = weekStart.getTime();
    const weekEndMs = addDays(weekStart, 7).getTime();
    const actuals = buildDailyActuals(cards, weekStartMs, Math.min(weekEndMs, nowMs), nowMs);
    const days = [];
    for (let dayIndex = 6; dayIndex >= 0; dayIndex -= 1) {
        const dayStart = addDays(weekStart, dayIndex);
        const dayStartMs = dayStart.getTime();
        const dayEndMs = addDays(dayStart, 1).getTime();
        const listed = cards
            .filter(card => card.clockIn && isSameDay(new Date(card.clockIn.dateTime), dayStart))
            .sort((a, b) => getCardSortMs(b) - getCardSortMs(a));
        const drawn = cards.filter(card => card.clockIn
            && cardStartMs(card) < dayEndMs && cardEndMs(card, nowMs) > dayStartMs);
        if (!listed.length && !drawn.length) continue;
        const workedMs = actuals.get(formatDateInputValue(dayStart)) || 0;
        days.push({ dayStart, dayStartMs, dayEndMs, listed, drawn, workedMs });
    }
    return days;
}

// ── Dragging an edge ────────────────────────────────────────────────────────
// target: { breakIndex: null } for the card itself, or the index of one of its breaks.
// edge: 'start' or 'end'. A running card or an ongoing break has no end to drag.

function breakSpans(card, nowMs) {
    return card.breaks
        .map((b, index) => ({ index, startMs: ms(b.start), endMs: b.end ? ms(b.end) : nowMs }))
        .filter(span => Number.isFinite(span.startMs))
        .sort((a, b) => a.startMs - b.startMs);
}

// How far an edge may move, keeping the card valid: breaks inside the card, apart from each
// other, nothing shorter than one snap, nothing ending in the future. Null when it cannot move.
export function dragBounds(card, breakIndex, edge, nowMs = Date.now()) {
    const startMs = cardStartMs(card);
    const endMs = cardEndMs(card, nowMs);
    const spans = breakSpans(card, nowMs);
    let bounds;
    if (breakIndex === null) {
        if (edge === 'end' && !card.clockOut) return null;
        bounds = edge === 'start'
            ? { minMs: -Infinity, maxMs: Math.min(spans[0]?.startMs ?? Infinity, endMs - SNAP_MS) }
            : { minMs: Math.max(startMs + SNAP_MS, spans.at(-1)?.endMs ?? -Infinity), maxMs: nowMs };
    } else {
        const position = spans.findIndex(span => span.index === breakIndex);
        assert(position !== -1, `break ${breakIndex} exists`);
        const span = spans[position];
        if (edge === 'end' && !card.breaks[breakIndex].end) return null;
        bounds = edge === 'start'
            ? { minMs: Math.max(startMs, spans[position - 1]?.endMs ?? -Infinity), maxMs: span.endMs - SNAP_MS }
            : { minMs: span.startMs + SNAP_MS, maxMs: Math.min(spans[position + 1]?.startMs ?? Infinity, endMs) };
    }
    return bounds.minMs <= bounds.maxMs ? bounds : null;
}

// The edge moved to `targetMs`, snapped and held within its bounds; null when it cannot move.
export function applyDrag(card, breakIndex, edge, targetMs, nowMs = Date.now()) {
    const bounds = dragBounds(card, breakIndex, edge, nowMs);
    if (!bounds) return null;
    const movedMs = clamp(snapMs(targetMs), bounds.minMs, bounds.maxMs);
    const moved = new Date(movedMs);
    if (breakIndex === null) {
        return { movedMs, card: buildUpdatedCard(card, edge === 'start' ? { clockIn: moved } : { clockOut: moved }) };
    }
    const breaks = cloneCardBreaks(card.breaks);
    breaks[breakIndex][edge] = { ...(breaks[breakIndex][edge] || {}), dateTime: moved.toISOString() };
    return { movedMs, card: buildUpdatedCard(card, { breaks }) };
}

// ── Adding a break ──────────────────────────────────────────────────────────

export const BREAK_DEFAULT_MS = 10 * 60 * 1000;

// A new break around atMs: BREAK_DEFAULT_MS long (shorter in a smaller gap), snapped, inside the
// card and inside the gap between its breaks that holds atMs. Null on a break, outside the card,
// or in a gap shorter than one snap.
export function proposeBreak(card, atMs, nowMs = Date.now()) {
    const startMs = cardStartMs(card);
    const endMs = cardEndMs(card, nowMs);
    if (!(atMs > startMs && atMs < endMs)) return null;
    const spans = breakSpans(card, nowMs);
    if (spans.some(span => atMs >= span.startMs && atMs <= span.endMs)) return null;
    const gapStartMs = Math.max(startMs, ...spans.filter(span => span.endMs < atMs).map(span => span.endMs));
    const gapEndMs = Math.min(endMs, ...spans.filter(span => span.startMs > atMs).map(span => span.startMs));
    const lengthMs = Math.floor(Math.min(BREAK_DEFAULT_MS, gapEndMs - gapStartMs) / SNAP_MS) * SNAP_MS;
    if (lengthMs < SNAP_MS) return null;
    const breakStartMs = clamp(snapMs(atMs - lengthMs / 2), gapStartMs, gapEndMs - lengthMs);
    return { startMs: breakStartMs, endMs: breakStartMs + lengthMs };
}

export function addBreak(card, { startMs, endMs }) {
    assert(startMs < endMs, 'a break ends after it starts');
    const added = { breakId: undefined, start: { dateTime: new Date(startMs).toISOString() }, end: { dateTime: new Date(endMs).toISOString() } };
    return buildUpdatedCard(card, { breaks: [...card.breaks, added] });
}

// ── The edit dialog ─────────────────────────────────────────────────────────

// A datetime-local value, or the original timestamp when the value is unchanged: the input has
// minute precision, and an untouched field must not lose its seconds.
function readDatetimeInput(value, originalIso) {
    if (!value) return null;
    if (originalIso && toLocalDatetimeInput(new Date(originalIso)) === value) return new Date(originalIso);
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? undefined : parsed;
}

// form: { clockIn, clockOut, notes, breaks: [{ breakId, start, end }] } as the inputs hold them.
// Returns { changes } for buildUpdatedCard, or { error } saying what to fix. Nothing may be in the
// future, and a clocked-out card stays clocked out: clearing its clock-out would set it running
// again from its old clock-in, beside whatever card is open now.
export function readEditForm(card, form, nowMs = Date.now()) {
    const clockIn = readDatetimeInput(form.clockIn, card.clockIn?.dateTime);
    const clockOut = readDatetimeInput(form.clockOut, card.clockOut?.dateTime);
    if (!clockIn) return { error: 'Clock-in time is required.' };
    if (clockOut === undefined) return { error: 'Clock-out is not a valid time.' };
    if (!clockOut && card.clockOut) return { error: 'A clocked-out card needs its clock-out.' };
    if (clockOut && clockOut <= clockIn) return { error: 'Clock-out must be after clock-in.' };
    if ((clockOut || clockIn).getTime() > nowMs) return { error: 'Times cannot be in the future.' };

    const breaks = [];
    for (const row of form.breaks) {
        if (!row.start && !row.end) continue;
        const original = card.breaks.find(b => b.breakId && b.breakId === row.breakId);
        const start = readDatetimeInput(row.start, original?.start?.dateTime);
        const end = readDatetimeInput(row.end, original?.end?.dateTime);
        if (!start) return { error: 'Every break needs a start.' };
        if (end === undefined) return { error: 'A break end is not a valid time.' };
        if (start < clockIn || (clockOut && start >= clockOut)) return { error: 'Break start must be within the card.' };
        if (end && end <= start) return { error: 'Break end must be after break start.' };
        if (end && clockOut && end > clockOut) return { error: 'Break end must be within the card.' };
        if (!end && clockOut) return { error: 'Only a running card can have a break without an end.' };
        if ((end || start).getTime() > nowMs) return { error: 'Times cannot be in the future.' };
        breaks.push({ breakId: row.breakId || undefined, start: { dateTime: start.toISOString() }, end: end ? { dateTime: end.toISOString() } : null });
    }
    const sorted = [...breaks].sort((a, b) => ms(a.start) - ms(b.start));
    for (let i = 1; i < sorted.length; i += 1) {
        if (!sorted[i - 1].end || ms(sorted[i - 1].end) > ms(sorted[i].start)) return { error: 'Breaks must not overlap.' };
    }
    return { changes: { clockIn, clockOut, breaks, notes: form.notes.trim() } };
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

// ── Weekly average ──────────────────────────────────────────────────────────

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
// A running card or an ongoing break runs until nowMs.
export function buildDailyActuals(cards, rangeStartMs, rangeEndMs, nowMs = rangeEndMs) {
    const totals = new Map();
    cards.forEach(card => getWorkSegments(card, nowMs).forEach(([segmentStartMs, segmentEndMs]) => {
        let cursorMs = Math.max(segmentStartMs, rangeStartMs);
        const endMs = Math.min(segmentEndMs, rangeEndMs);
        while (cursorMs < endMs) {
            const dayStart = getDayStart(cursorMs);
            const sliceEndMs = Math.min(endMs, addDays(dayStart, 1).getTime());
            assert(sliceEndMs > cursorMs, 'each day slice moves forward');
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

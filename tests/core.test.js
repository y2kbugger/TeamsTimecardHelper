// Run with TZ=America/New_York so the daylight saving cases bite (agent-test.sh sets it).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
    addDays, getWeekStart, getWeekRange, formatDateInputValue, toLocalDatetimeInput, parseLocalDate,
    snapMs, formatDurationHms, formatSignedDuration, normalizeCard, deriveCardState, workedMs,
    filterCardsForWeek, findActiveCard, buildWeekDays, buildUpdatedCard, buildTimeCardUpdateBody,
    getWorkSegments, isLongCard, proposeCard, buildNewTimeCardBody, buildDailyActuals, buildWeeklyModel,
    dragBounds, applyDrag, readEditForm, proposeBreak, addBreak,
} from '../core.js';

const local = s => new Date(s); // 'YYYY-MM-DDTHH:MM' parses as local time
const day = d => formatDateInputValue(d);

function card({ id = 'c', clockIn, clockOut = null, breaks = [] }) {
    return normalizeCard({
        id,
        clockInEvent: clockIn ? { dateTime: local(clockIn).toISOString() } : null,
        clockOutEvent: clockOut ? { dateTime: local(clockOut).toISOString() } : null,
        breaks: breaks.map(([start, end], i) => ({
            breakId: `b${i}`,
            start: { dateTime: local(start).toISOString() },
            end: end ? { dateTime: local(end).toISOString() } : null,
        })),
    });
}

test('timezone is one with daylight saving', () => {
    assert.notEqual(new Date('2026-01-15').getTimezoneOffset(), new Date('2026-07-15').getTimezoneOffset());
});

test('weeks start on Sunday at local midnight', () => {
    assert.equal(day(getWeekStart(local('2026-10-21T15:00'))), '2026-10-18');
    assert.equal(day(getWeekStart(local('2026-10-18T00:00'))), '2026-10-18');
    assert.equal(getWeekStart(local('2026-10-24T23:59')).getHours(), 0);
});

test('week stepping crosses daylight saving changes', () => {
    assert.equal(day(getWeekStart(addDays(local('2026-11-01T00:00'), 7))), '2026-11-08');
    assert.equal(day(getWeekStart(addDays(local('2026-03-15T00:00'), -7))), '2026-03-08');
    for (const week of ['2026-03-08', '2026-11-01']) {
        const { weekEndExclusive } = getWeekRange(parseLocalDate(week));
        assert.equal(weekEndExclusive.getHours(), 0, week);
        assert.equal(day(weekEndExclusive), day(addDays(parseLocalDate(week), 7)), week);
    }
});

test('input formatting and parsing', () => {
    assert.equal(toLocalDatetimeInput(local('2026-01-02T03:04')), '2026-01-02T03:04');
    assert.equal(parseLocalDate('2026-01-02').getHours(), 0);
    assert.equal(parseLocalDate('nope'), null);
    assert.equal(parseLocalDate(''), null);
});

test('durations', () => {
    assert.equal(formatDurationHms(3_723_999), '01:02:03');
    assert.equal(formatDurationHms(-5), '00:00:00');
    assert.equal(formatDurationHms(100 * 3600e3), '100:00:00');
    assert.equal(formatSignedDuration(-60e3), '-00:01:00');
    assert.equal(formatSignedDuration(60e3), '+00:01:00');
    assert.equal(snapMs(local('2026-01-01T10:02:31').getTime()), local('2026-01-01T10:05').getTime());
});

test('card state', () => {
    assert.equal(deriveCardState(card({ clockIn: '2026-01-01T09:00', clockOut: '2026-01-01T10:00' })), 'clockedOut');
    assert.equal(deriveCardState(card({ clockIn: '2026-01-01T09:00' })), 'clockedIn');
    assert.equal(deriveCardState(card({ clockIn: '2026-01-01T09:00', breaks: [['2026-01-01T09:30', null]] })), 'onBreak');
});

test('worked time subtracts breaks, open ones up to now', () => {
    const closed = card({ clockIn: '2026-01-01T09:00', clockOut: '2026-01-01T12:00', breaks: [['2026-01-01T10:00', '2026-01-01T10:15']] });
    assert.equal(workedMs(closed), 2.75 * 3600e3);
    const open = card({ clockIn: '2026-01-01T09:00', breaks: [['2026-01-01T10:00', null]] });
    assert.equal(workedMs(open, local('2026-01-01T11:00').getTime()), 3600e3);
});

test('week filter keeps overlapping cards, active card is the newest open one', () => {
    const cards = [
        card({ id: 'before', clockIn: '2026-10-17T09:00', clockOut: '2026-10-17T10:00' }),
        card({ id: 'spans', clockIn: '2026-10-17T23:00', clockOut: '2026-10-18T01:00' }),
        card({ id: 'in', clockIn: '2026-10-19T09:00', clockOut: '2026-10-19T10:00' }),
        card({ id: 'open', clockIn: '2026-10-21T09:00' }),
    ];
    const now = local('2026-10-21T12:00').getTime();
    assert.deepEqual(filterCardsForWeek(cards, local('2026-10-20T00:00'), now).map(c => c.id), ['spans', 'in', 'open']);
    assert.equal(findActiveCard(cards).id, 'open');
    assert.equal(findActiveCard(cards.slice(0, 3)), null);
});

test('week days: every day begun, newest first, cards listed newest first on their clock-in day', () => {
    const days = buildWeekDays([
        card({ id: 'a', clockIn: '2026-10-19T09:00', clockOut: '2026-10-19T10:00' }),
        card({ id: 'b', clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T10:00' }),
        card({ id: 'c', clockIn: '2026-10-19T13:00', clockOut: '2026-10-19T14:00' }),
    ], parseLocalDate('2026-10-18'), local('2026-10-21T23:00').getTime());
    assert.deepEqual(days.map(d => [day(d.dayStart), d.listed.map(c => c.id).join(''), d.workedMs / 3600e3]),
        [['2026-10-21', 'b', 1], ['2026-10-20', '', 0], ['2026-10-19', 'ca', 2], ['2026-10-18', '', 0]]);
});

test('week days: a card past midnight is drawn on both days and its time split between them', () => {
    const days = buildWeekDays([card({ id: 'late', clockIn: '2026-10-19T22:00', clockOut: '2026-10-20T01:00' })],
        parseLocalDate('2026-10-18'), local('2026-10-20T12:00').getTime());
    assert.deepEqual(days.map(d => [day(d.dayStart), d.listed.length, d.drawn.length, d.workedMs / 3600e3]),
        [['2026-10-20', 0, 1, 1], ['2026-10-19', 1, 1, 2], ['2026-10-18', 0, 0, 0]]);
});

test('new card: 30 minutes centred and snapped, between cards, never past now', () => {
    const at = s => local(s).getTime();
    const day = { dayStartMs: at('2026-10-21T00:00'), dayEndMs: at('2026-10-22T00:00') };
    const now = at('2026-10-21T18:00');
    const cards = [
        card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T12:00' }),
        card({ clockIn: '2026-10-21T12:20', clockOut: '2026-10-21T15:00' }),
    ];
    const hm = ({ startMs, endMs }) => `${local(startMs).toTimeString().slice(0, 5)}-${local(endMs).toTimeString().slice(0, 5)}`;
    assert.equal(hm(proposeCard(cards, day, at('2026-10-21T16:02'), now)), '15:45-16:15');
    assert.equal(hm(proposeCard(cards, day, at('2026-10-21T15:01'), now)), '15:00-15:30', 'held after the card before');
    assert.equal(hm(proposeCard(cards, day, at('2026-10-21T12:10'), now)), '12:00-12:20', 'shorter in a 20-minute gap');
    assert.equal(hm(proposeCard(cards, day, at('2026-10-21T17:55'), now)), '17:30-18:00', 'over by now');
    assert.equal(hm(proposeCard([], day, at('2026-10-21T00:05'), now)), '00:00-00:30', 'inside the day');
    assert.equal(proposeCard(cards, day, at('2026-10-21T10:00'), now), null, 'not on a card');
    assert.equal(proposeCard(cards, day, at('2026-10-21T18:30'), now), null, 'not after now');
    const ragged = [card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T12:03' }), card({ clockIn: '2026-10-21T12:07', clockOut: '2026-10-21T13:00' })];
    assert.equal(proposeCard(ragged, day, at('2026-10-21T12:05'), now), null, 'no snap fits between 12:03 and 12:07');
    const body = buildNewTimeCardBody({ startMs: at('2026-10-21T16:00'), endMs: at('2026-10-21T16:30') });
    assert.deepEqual(Object.keys(body), ['clockInEvent', 'clockOutEvent', 'breaks']);
    assert.equal(body.clockOutEvent.dateTime, local('2026-10-21T16:30').toISOString());
});

test('week days: a running card counts up to now', () => {
    const [today] = buildWeekDays([card({ clockIn: '2026-10-21T09:00' })], parseLocalDate('2026-10-18'), local('2026-10-21T10:30').getTime());
    assert.equal(today.workedMs, 1.5 * 3600e3);
});

test('drag: snaps to 5 minutes and stays inside its neighbours', () => {
    const now = local('2026-10-21T18:00').getTime();
    const c = card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T12:00', breaks: [['2026-10-21T10:00', '2026-10-21T10:15']] });
    const at = s => local(s).getTime();
    assert.equal(applyDrag(c, null, 'start', at('2026-10-21T08:41'), now).movedMs, at('2026-10-21T08:40'));
    assert.equal(applyDrag(c, null, 'start', at('2026-10-21T10:30'), now).movedMs, at('2026-10-21T10:00'), 'not past the break');
    assert.equal(applyDrag(c, null, 'end', at('2026-10-21T10:05'), now).movedMs, at('2026-10-21T10:15'), 'not before the break ends');
    assert.equal(applyDrag(c, null, 'end', at('2026-10-21T19:00'), now).movedMs, now, 'not into the future');
    assert.equal(applyDrag(c, 0, 'start', at('2026-10-21T08:00'), now).movedMs, at('2026-10-21T09:00'), 'not before clock-in');
    assert.equal(applyDrag(c, 0, 'end', at('2026-10-21T10:01'), now).movedMs, at('2026-10-21T10:05'), 'at least one snap long');
    const moved = applyDrag(c, 0, 'end', at('2026-10-21T10:30'), now).card;
    assert.equal(moved.breaks[0].breakId, 'b0');
    assert.equal(moved.breaks[0].end.dateTime, local('2026-10-21T10:30').toISOString());
    assert.equal(c.breaks[0].end.dateTime, local('2026-10-21T10:15').toISOString(), 'the original is untouched');
});

test('drag: a running card and an ongoing break have no end to drag', () => {
    const now = local('2026-10-21T12:00').getTime();
    const c = card({ clockIn: '2026-10-21T09:00', breaks: [['2026-10-21T10:00', '2026-10-21T10:15'], ['2026-10-21T11:00', null]] });
    assert.equal(dragBounds(c, null, 'end', now), null);
    assert.equal(dragBounds(c, 1, 'end', now), null);
    assert.equal(dragBounds(c, 1, 'start', now).minMs, local('2026-10-21T10:15').getTime(), 'after the break before it');
    const moved = applyDrag(c, 1, 'start', local('2026-10-21T10:50').getTime(), now).card;
    assert.equal(moved.breaks[1].end, null, 'the ongoing break keeps going');
    assert.equal(moved.state, 'onBreak');
});

test('drag: an edge with no room to move refuses', () => {
    const now = local('2026-10-21T12:00').getTime();
    const c = card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T09:03', breaks: [['2026-10-21T09:01', '2026-10-21T09:02']] });
    assert.equal(dragBounds(c, 0, 'start', now), null);
    assert.equal(applyDrag(c, 0, 'start', now, now), null);
    assert.ok(dragBounds(c, null, 'end', now), 'the card itself can still grow');
});

test('new break: ten minutes around the pointer, snapped', () => {
    const now = local('2026-10-21T18:00').getTime();
    const c = card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T12:00', breaks: [['2026-10-21T10:30', '2026-10-21T10:45']] });
    const at = s => local(s).getTime();
    const hhmm = ms => toLocalDatetimeInput(new Date(ms)).slice(11);
    const proposed = s => { const p = proposeBreak(c, at(s), now); return p && `${hhmm(p.startMs)}-${hhmm(p.endMs)}`; };
    assert.equal(proposed('2026-10-21T11:22'), '11:15-11:25');
    assert.equal(proposed('2026-10-21T10:47'), '10:45-10:55', 'pushed out of the break before it');
    assert.equal(proposed('2026-10-21T10:28'), '10:20-10:30', 'pushed out of the break after it');
    assert.equal(proposed('2026-10-21T09:01'), '09:00-09:10', 'held inside the card');
    assert.equal(proposed('2026-10-21T10:35'), null, 'not on a break');
    assert.equal(proposed('2026-10-21T08:00'), null, 'not outside the card');
    const added = addBreak(c, proposeBreak(c, at('2026-10-21T11:22'), now));
    assert.equal(added.breaks.length, 2);
    assert.equal(added.breaks[0].breakId, 'b0', 'the old break keeps its id');
    assert.equal(c.breaks.length, 1, 'the original is untouched');
});

test('new break: shorter in a small gap, none in a gap under five minutes', () => {
    const now = local('2026-10-21T18:00').getTime();
    const c = card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T12:00', breaks: [['2026-10-21T10:00', '2026-10-21T10:15'], ['2026-10-21T10:22', '2026-10-21T10:30'], ['2026-10-21T10:33', '2026-10-21T10:40']] });
    const gap7 = proposeBreak(c, local('2026-10-21T10:18').getTime(), now);
    assert.equal(gap7.endMs - gap7.startMs, 5 * 60e3);
    assert.equal(proposeBreak(c, local('2026-10-21T10:31').getTime(), now), null);
});

test('new break: on a running card, not past now', () => {
    const now = local('2026-10-21T10:04').getTime();
    const c = card({ clockIn: '2026-10-21T09:00' });
    const p = proposeBreak(c, local('2026-10-21T10:02').getTime(), now);
    assert.ok(p.endMs <= now, 'ends by now');
});

const EDIT_NOW = local('2026-10-21T18:00').getTime();

test('edit form: unchanged fields keep their seconds', () => {
    const c = normalizeCard({ id: 'c', clockInEvent: { dateTime: '2026-10-21T13:00:42.000Z' }, clockOutEvent: { dateTime: '2026-10-21T15:00:17.000Z' }, breaks: [] });
    const form = { clockIn: toLocalDatetimeInput(new Date(c.clockIn.dateTime)), clockOut: toLocalDatetimeInput(new Date(c.clockOut.dateTime)), notes: ' hi ', breaks: [] };
    const { changes } = readEditForm(c, form, EDIT_NOW);
    assert.equal(changes.clockIn.toISOString(), '2026-10-21T13:00:42.000Z');
    assert.equal(changes.clockOut.toISOString(), '2026-10-21T15:00:17.000Z');
    assert.equal(changes.notes, 'hi');
});

test('edit form: refuses what would make a bad card', () => {
    const c = card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T12:00', breaks: [['2026-10-21T10:00', '2026-10-21T10:15']] });
    const form = (over = {}) => ({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T12:00', notes: '', breaks: [{ breakId: 'b0', start: '2026-10-21T10:00', end: '2026-10-21T10:15' }], ...over });
    const error = over => readEditForm(c, form(over), EDIT_NOW).error;
    assert.equal(error({}), undefined);
    assert.match(error({ clockIn: '' }), /Clock-in time is required/);
    assert.match(error({ clockOut: '' }), /needs its clock-out/);
    assert.match(error({ clockOut: '2026-10-21T18:05' }), /in the future/);
    assert.match(error({ breaks: [{ start: '2026-10-21T10:00', end: '' }], clockOut: '' }), /needs its clock-out/);
    assert.match(error({ clockOut: '2026-10-21T08:00' }), /Clock-out must be after/);
    assert.match(error({ breaks: [{ start: '2026-10-21T08:30', end: '2026-10-21T09:30' }] }), /Break start must be within/);
    assert.match(error({ breaks: [{ start: '2026-10-21T11:00', end: '2026-10-21T12:30' }] }), /Break end must be within/);
    assert.match(error({ breaks: [{ start: '2026-10-21T11:00', end: '2026-10-21T10:30' }] }), /Break end must be after/);
    assert.match(error({ breaks: [{ start: '2026-10-21T11:00', end: '' }] }), /without an end/);
    assert.match(error({ breaks: [{ start: '', end: '2026-10-21T11:00' }] }), /needs a start/);
    assert.match(error({ breaks: [{ start: '2026-10-21T10:00', end: '2026-10-21T10:30' }, { start: '2026-10-21T10:20', end: '2026-10-21T10:40' }] }), /must not overlap/);
});

test('edit form: a running card may keep an ongoing break; empty rows are dropped', () => {
    const c = card({ clockIn: '2026-10-21T09:00', breaks: [['2026-10-21T10:00', null]] });
    const { changes } = readEditForm(c, { clockIn: '2026-10-21T09:00', clockOut: '', notes: '', breaks: [{ breakId: 'b0', start: '2026-10-21T10:00', end: '' }, { start: '', end: '' }] }, EDIT_NOW);
    assert.match(readEditForm(c, { clockIn: '2026-10-21T09:00', clockOut: '', notes: '', breaks: [{ start: '2026-10-21T18:30', end: '' }] }, EDIT_NOW).error, /in the future/);
    assert.equal(changes.clockOut, null);
    assert.equal(changes.breaks.length, 1);
    assert.equal(changes.breaks[0].breakId, 'b0');
    assert.equal(changes.breaks[0].end, null);
});

test('update body omits clock-out on an open card and keeps break ids', () => {
    const open = card({ clockIn: '2026-01-01T09:00', breaks: [['2026-01-01T09:30', null]] });
    const body = buildTimeCardUpdateBody(buildUpdatedCard(open, { notes: 'hi' }));
    assert.equal('clockOutEvent' in body, false);
    assert.equal(body.breaks[0].breakId, 'b0');
    assert.equal(body.breaks[0].end, undefined);
    assert.deepEqual(body.notes, { contentType: 'text', content: 'hi' });
});

test('a card over 16 hours is long, a running one too', () => {
    const now = local('2026-10-22T09:00').getTime();
    assert.equal(isLongCard(card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-22T01:00' }), now), false);
    assert.equal(isLongCard(card({ clockIn: '2026-10-21T09:00', clockOut: '2026-10-22T01:01' }), now), true);
    assert.equal(isLongCard(card({ clockIn: '2026-10-21T17:00' }), now), false);
    assert.equal(isLongCard(card({ clockIn: '2026-10-21T16:59' }), local('2026-10-22T09:00').getTime() + 3600e3), true);
});

test('dragging an open card never clocks it out', () => {
    const open = card({ clockIn: '2026-01-01T09:00' });
    const now = local('2026-01-01T11:00').getTime();
    assert.equal(dragBounds(open, null, 'end', now), null);
    const { card: moved } = applyDrag(open, null, 'start', local('2026-01-01T08:30').getTime(), now);
    assert.equal(moved.state, 'clockedIn');
    assert.equal('clockOutEvent' in buildTimeCardUpdateBody(moved), false);
});

test('work segments skip breaks and clip them to the card', () => {
    const c = card({ clockIn: '2026-01-01T09:00', clockOut: '2026-01-01T12:00', breaks: [['2026-01-01T08:00', '2026-01-01T09:30'], ['2026-01-01T11:00', '2026-01-01T11:30']] });
    const hours = getWorkSegments(c, 0).map(([a, b]) => [new Date(a).getHours() + new Date(a).getMinutes() / 60, new Date(b).getHours() + new Date(b).getMinutes() / 60]);
    assert.deepEqual(hours, [[9.5, 11], [11.5, 12]]);
});

test('daily actuals split at midnight, also on daylight saving days', () => {
    const c = card({ clockIn: '2026-03-07T23:00', clockOut: '2026-03-08T04:00' });
    const totals = buildDailyActuals([c], 0, local('2026-03-09T00:00').getTime());
    assert.equal(totals.get('2026-03-07'), 3600e3);
    assert.equal(totals.get('2026-03-08'), 3 * 3600e3); // 00:00 to 04:00 is 3h: clocks skip 02:00
});

test('weekly model expects time only through today', () => {
    // 25 h/week over Mon-Fri is 5 h/day. Today is Wednesday; worked 5 h on Mon and Tue, 2 h today.
    const now = local('2026-10-21T15:00');
    const cards = [
        card({ id: 'm', clockIn: '2026-10-19T09:00', clockOut: '2026-10-19T14:00' }),
        card({ id: 't', clockIn: '2026-10-20T09:00', clockOut: '2026-10-20T14:00' }),
        card({ id: 'w', clockIn: '2026-10-21T09:00', clockOut: '2026-10-21T11:00' }),
    ];
    const model = buildWeeklyModel(cards, { startDate: parseLocalDate('2026-10-18'), weeklyHours: 25, workdays: [1, 2, 3, 4, 5] }, now);
    assert.equal(model.expectedCumulativeMs, 15 * 3600e3);
    assert.equal(model.differenceMs, -3 * 3600e3);
    const [week, ...days] = model.rows;
    assert.equal(week.targetMs, 25 * 3600e3);
    assert.equal(week.expectedMs, 15 * 3600e3);
    assert.deepEqual(days.map(d => day(d.day)), ['2026-10-24', '2026-10-23', '2026-10-22', '2026-10-21', '2026-10-20', '2026-10-19', '2026-10-18']);
    const friday = days[1];
    assert.equal(friday.isFuture, true);
    assert.equal(friday.differenceMs, 0);
    assert.equal(friday.cumulativeDifferenceMs, -3 * 3600e3);
});

test('weekly model: the first week is partial when the start date is mid-week', () => {
    const now = local('2026-10-30T12:00');
    const model = buildWeeklyModel([], { startDate: parseLocalDate('2026-10-21'), weeklyHours: 25, workdays: [1, 2, 3, 4, 5] }, now);
    const weeks = model.rows.filter(r => r.rowType === 'week');
    assert.equal(weeks.length, 2);
    assert.equal(weeks[1].isPartialWeek, true);
    assert.equal(weeks[1].targetMs, 15 * 3600e3);
});

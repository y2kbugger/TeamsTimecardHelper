// One day's timeline: cards and their breaks drawn across the day (24 hours, or 23 or 25 on
// daylight saving days), dragging their edges, and adding breaks. A drag previews the snapped,
// clamped time and saves only on release; Escape cancels. While a drag runs the app holds back
// re-renders, so a background refresh cannot pull the block from under the pointer. Hovering a
// card shows a faint ghost of the break a Ctrl+click would add there, hovering free time one of a
// card; holding Ctrl arms it. A plain click only selects, so a click meant to select never adds
// anything to a real timesheet. The timeline has its own tip rather than the browser's titles: it
// shows at once, above the axis, and a handle pointed at or dragged gets a marker above it.
import {
    applyDrag, assert, cardEndMs, cardStartMs, clamp, dragBounds, formatDurationShort, FULL_DAY_HOURS, isLongCard,
    proposeBreak, proposeCard,
} from './core.js';

const HOUR_LABEL_EVERY_HOURS = 2;
const BLOCK_WIDTH_MIN_PERCENT = 0.15;
const TIP_GAP_PX = 20; // between the tip and the top of its axis, clear of the hour labels
const TIP_MARGIN_PX = 6; // the least room between the tip and the edge of the window

const hourFormat = new Intl.DateTimeFormat(undefined, { hour: 'numeric' });
const timeFormat = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });

let drag = null; // the drag in progress, see startDrag
// { element, kind: 'break' | 'card', cardId (a break's card), proposal } while something new is
// offered. Never during a drag, so the click that follows a drag's release finds none and adds
// nothing.
let ghost = null;
// { axis, x, span, edge } under the pointer, for the tip: span is a card's block or a break's
// overlay, edge a handle of it. Null during a drag, which keeps the tip on the dragged edge.
let hover = null;
let tip = null; // the tip's element, made on first use
let marker = null; // the marker above the handle pointed at or dragged

export function isDragging() {
    return drag !== null;
}

function percentOfDay(day, timeMs) {
    return ((timeMs - day.dayStartMs) / (day.dayEndMs - day.dayStartMs)) * 100;
}

// Positions an element over [startMs, endMs) of the day, clipped to the day.
function place(element, day, startMs, endMs) {
    const left = percentOfDay(day, Math.max(startMs, day.dayStartMs));
    const right = percentOfDay(day, Math.min(endMs, day.dayEndMs));
    element.style.left = `${left}%`;
    element.style.width = `${Math.max(BLOCK_WIDTH_MIN_PERCENT, right - left)}%`;
}

// A card is worked on in the day it clocked in. On a day it only runs into past midnight it is an
// untouchable ghost; at home, an edge past midnight gets its handle at the end of the axis.
// endSet: whether its end is a set time; a running span has none and ends at now.
function makeSpan(className, day, { cardId, breakIndex, startMs, endMs, endSet, card, nowMs, overflow }) {
    const element = document.createElement('div');
    element.className = className;
    element.dataset.cardId = cardId;
    if (breakIndex !== null) element.dataset.breakIndex = String(breakIndex);
    // Its times, for the tip.
    element.dataset.startMs = String(startMs);
    if (endSet) element.dataset.endMs = String(endMs);
    place(element, day, startMs, endMs);
    if (overflow) {
        element.classList.add('is-overflow');
        return element;
    }
    for (const edge of ['start', 'end']) {
        if (!dragBounds(card, breakIndex, edge, nowMs)) continue;
        const handle = document.createElement('div');
        handle.className = 'resize-handle';
        handle.dataset.edge = edge;
        element.appendChild(handle);
    }
    return element;
}

// day: one entry of core's buildWeekDays. hours: the part of the day the axis spans, { start, end }
// in whole hours since midnight (core's weekHourWindow on a narrow screen), the whole day by
// default. The axis's data-day-start-ms and data-day-end-ms are that span, which every position
// on it is a fraction of.
export function renderTimeline(day, nowMs, hours = FULL_DAY_HOURS) {
    const atHour = hour => {
        const at = new Date(day.dayStartMs);
        at.setHours(hour);
        return at.getTime();
    };
    const view = { dayStartMs: atHour(hours.start), dayEndMs: hours.end === 24 ? day.dayEndMs : atHour(hours.end) };
    assert(view.dayStartMs < view.dayEndMs, 'an axis spans some time');
    const axis = document.createElement('div');
    axis.className = 'timeline';
    axis.dataset.dayStartMs = String(view.dayStartMs);
    axis.dataset.dayEndMs = String(view.dayEndMs);
    axis.style.setProperty('--hours', String(hours.end - hours.start));
    // The whole day labels every other hour, and a narrow screen every sixth (the rest are
    // minor); a window labels every hour, and a narrow screen every other one past 8 hours.
    const fullDay = hours.start === 0 && hours.end === 24;
    const step = fullDay ? HOUR_LABEL_EVERY_HOURS : 1;
    for (let hour = hours.start + step; hour < hours.end; hour += step) {
        const minor = fullDay ? hour % 6 : hours.end - hours.start > 8 && hour % 2;
        const label = document.createElement('div');
        label.className = minor ? 'hour-label hour-minor' : 'hour-label';
        label.style.left = `${percentOfDay(view, atHour(hour))}%`;
        label.textContent = hourFormat.format(atHour(hour));
        axis.appendChild(label);
    }
    placeNowLine(axis, view, nowMs);
    day.drawn.forEach(card => appendCard(axis, day, view, card, nowMs));
    return axis;
}

// The now line, on today's axis only. Moved every second with the running spans, not only when
// the week is drawn, so it never lags them (on returning to the tab, say).
function placeNowLine(axis, day, nowMs) {
    let line = axis.querySelector('.now-line');
    if (!(nowMs > day.dayStartMs && nowMs < day.dayEndMs)) return line?.remove();
    if (!line) {
        line = document.createElement('div');
        line.className = 'now-line';
        axis.prepend(line);
    }
    line.style.left = `${percentOfDay(day, nowMs)}%`;
}

export function updateNowLines(root, nowMs) {
    root.querySelectorAll('.timeline').forEach(axis => placeNowLine(axis, axisDay(axis), nowMs));
}

function axisDay(axis) {
    return { dayStartMs: Number(axis.dataset.dayStartMs), dayEndMs: Number(axis.dataset.dayEndMs) };
}

// day: the calendar day (which card is at home here); view: the part of it the axis spans.
function appendCard(axis, day, view, card, nowMs) {
    const startMs = cardStartMs(card);
    const endMs = cardEndMs(card, nowMs);
    // Outside the window (only a ghost can be): nothing, rather than a sliver at its edge.
    if (endMs <= view.dayStartMs || startMs >= view.dayEndMs) return;
    const overflow = startMs < day.dayStartMs;
    const block = makeSpan('tc-block', view, {
        cardId: card.id, breakIndex: null, startMs, endMs, endSet: Boolean(card.clockOut), card, nowMs, overflow,
    });
    if (!card.clockOut) block.classList.add('is-running');
    if (isLongCard(card, nowMs)) block.classList.add('is-long');
    if (startMs >= day.dayStartMs) {
        const label = document.createElement('span');
        label.className = 'tc-label';
        label.textContent = timeFormat.format(startMs);
        block.appendChild(label);
    }
    axis.appendChild(block);
    card.breaks.forEach((b, breakIndex) => {
        if (!b.start) return;
        const breakStartMs = new Date(b.start.dateTime).getTime();
        const breakEndMs = b.end ? new Date(b.end.dateTime).getTime() : nowMs;
        if (breakEndMs <= view.dayStartMs || breakStartMs >= view.dayEndMs) return;
        const overlay = makeSpan('break-overlay', view, {
            cardId: card.id, breakIndex, startMs: breakStartMs, endMs: breakEndMs, endSet: Boolean(b.end), card, nowMs,
            overflow,
        });
        if (!b.end) overlay.classList.add('is-running');
        axis.appendChild(overlay);
    });
}

// Running cards and ongoing breaks grow with the clock.
export function updateRunningSpans(root, cardsById, nowMs) {
    root.querySelectorAll('.timeline .is-running').forEach(element => {
        if (drag?.element === element) return;
        const card = cardsById.get(element.dataset.cardId);
        if (!card) return;
        const axis = element.parentElement;
        const day = axisDay(axis);
        const breakIndex = element.dataset.breakIndex;
        if (breakIndex === undefined) element.classList.toggle('is-long', isLongCard(card, nowMs));
        const startMs = breakIndex === undefined
            ? cardStartMs(card)
            : new Date(card.breaks[Number(breakIndex)].start.dateTime).getTime();
        place(element, day, startMs, nowMs);
    });
    drawHover(); // a running span's length in the tip grows too
}

// ── Pointing at a timeline ──────────────────────────────────────────────────

// handlers: { cardById(id), cards(), onCommit(updatedCard, movedMs), onEnd(committed),
//             onSelect(cardId), onAddBreak(card, { startMs, endMs }), onAddCard({ startMs, endMs }) }.
//             cardById and cards answer nothing while no change may be made.
export function attachTimeline(root, handlers) {
    root.addEventListener('pointerdown', event => {
        const span = event.target.closest('.tc-block, .break-overlay');
        if (!span) return;
        handlers.onSelect(span.dataset.cardId);
        const handle = event.target.closest('.resize-handle');
        if (handle && event.button === 0) {
            hideGhost();
            startDrag(event, span, handle, handlers);
        }
    });
    root.addEventListener('pointermove', event => {
        updateHover(event);
        updateGhost(event, handlers);
        drawHover();
    });
    root.addEventListener('pointerleave', () => {
        hideGhost();
        clearHover();
    });
    // The tip is placed on the window, so it would stay behind on a scroll.
    window.addEventListener('scroll', clearHover, { capture: true, passive: true });
    root.addEventListener('click', event => {
        // Only the ghost on show, and only where it was offered: a re-render since leaves it
        // detached and stale.
        if (!event.ctrlKey || !ghost?.element.isConnected) return;
        const target = ghostTarget(event);
        if (target?.kind !== ghost.kind || target.cardId !== ghost.cardId || target.axis !== ghost.element.parentElement) return;
        const { kind, cardId, proposal } = ghost;
        hideGhost();
        if (kind === 'card') return handlers.onAddCard(proposal);
        const card = handlers.cardById(cardId);
        if (card) handlers.onAddBreak(card, proposal);
    });
    window.addEventListener('keydown', event => {
        if (event.key === 'Control') armGhost(true);
        if (event.key !== 'Escape') return;
        if (drag) endDrag(false);
        hideGhost();
        clearHover();
    });
    window.addEventListener('keyup', event => {
        if (event.key === 'Control') armGhost(false);
    });
    window.addEventListener('blur', () => {
        armGhost(false);
        clearHover();
    });
}

// Ctrl held: the ghost is what a click adds now, and the tip says so.
function armGhost(armed) {
    setArmed(armed);
    drawHover();
}

function setArmed(armed) {
    document.body.classList.toggle('ctrl-held', armed);
    ghost?.element.classList.toggle('is-armed', armed);
}

// What a Ctrl+click under the pointer would add: a break on a card's block, a card on free time.
// Nothing on a handle or a break, or during a drag.
function ghostOffer(event, handlers) {
    const target = ghostTarget(event);
    if (!target) return null;
    const { axis } = target;
    const day = axisDay(axis);
    const rect = axis.getBoundingClientRect();
    const atMs = day.dayStartMs + ((event.clientX - rect.left) / rect.width) * (day.dayEndMs - day.dayStartMs);
    if (target.kind === 'break') {
        const card = handlers.cardById(target.cardId);
        const proposal = card && proposeBreak(card, atMs, Date.now());
        return proposal ? { kind: 'break', cardId: card.id, axis, day, proposal } : null;
    }
    const cards = handlers.cards();
    const proposal = cards && proposeCard(cards, day, atMs, Date.now());
    return proposal ? { kind: 'card', cardId: null, axis, day, proposal } : null;
}

// Where the pointer is, for a ghost: { kind: 'break', cardId } on a card's block, { kind: 'card' }
// on the free time of an axis, null elsewhere.
function ghostTarget(event) {
    const axis = event.target.closest('.timeline');
    if (drag || !axis || event.target.closest('.resize-handle, .break-overlay, .hour-label')) return null;
    const block = event.target.closest('.tc-block');
    return block ? { kind: 'break', cardId: block.dataset.cardId, axis } : { kind: 'card', cardId: null, axis };
}

function updateGhost(event, handlers) {
    const offer = ghostOffer(event, handlers);
    if (!offer) return hideGhost();
    const { axis, day, proposal } = offer;
    if (!ghost) ghost = { element: document.createElement('div') };
    const { element } = ghost;
    element.className = `ghost ${offer.kind}-ghost`;
    place(element, day, proposal.startMs, proposal.endMs);
    if (element.parentElement !== axis) axis.appendChild(element);
    Object.assign(ghost, { kind: offer.kind, cardId: offer.cardId, proposal });
    setArmed(event.ctrlKey);
}

function hideGhost() {
    ghost?.element.remove();
    ghost = null;
}

// ── The tip and the handle's marker ─────────────────────────────────────────

function updateHover(event) {
    const axis = event.target.closest('.timeline');
    hover = drag || !axis ? null : {
        axis,
        x: event.clientX,
        span: event.target.closest('.tc-block, .break-overlay'),
        edge: event.target.closest('.resize-handle')?.dataset.edge ?? null,
    };
}

function clearHover() {
    hover = null;
    drawHover();
}

// The tip for what is under the pointer: a handle's time with its marker, or a span's times and
// length, and what a Ctrl+click there would add. During a drag the drag keeps both up to date.
function drawHover() {
    if (drag) return;
    if (!hover?.axis.isConnected || (hover.span && !hover.span.isConnected)) {
        hover = null;
        hideTip();
        hideMarker();
        return;
    }
    const { axis, span, edge } = hover;
    if (span && edge) {
        const atMs = edgeMs(span, edge);
        showTip(axis, placeMarker(axis, span, atMs, false), [edgeLine(span, edge, atMs)]);
        return;
    }
    hideMarker();
    const lines = span ? spanLines(span) : [];
    if (ghost) lines.push(ghostLine());
    if (lines.length === 0) return hideTip();
    showTip(axis, hover.x, lines);
}

// A line of the tip: { key, text, note, tone }. key is a key to press, shown as one; note is said
// quietly after the text; tone colours it: 'card', 'break', 'danger' or 'muted'.
function spanLines(span) {
    const startMs = Number(span.dataset.startMs);
    const endSet = span.dataset.endMs !== undefined;
    const endMs = endSet ? Number(span.dataset.endMs) : Date.now();
    const range = `${timeFormat.format(startMs)} → ${endSet ? timeFormat.format(endMs) : 'now'}`;
    const isBreak = span.classList.contains('break-overlay');
    const lines = [{ text: isBreak ? `Break ${range}` : range, note: formatDurationShort(endMs - startMs) }];
    if (span.classList.contains('is-long')) lines.push({ text: 'Over 16 hours: check its clock-out', tone: 'danger' });
    return lines;
}

function ghostLine() {
    const { kind, proposal } = ghost;
    const text = `+ ${kind} ${timeFormat.format(proposal.startMs)} → ${timeFormat.format(proposal.endMs)}`;
    return document.body.classList.contains('ctrl-held') ? { text, tone: kind } : { key: 'Ctrl+click', text, tone: 'muted' };
}

const EDGE_NAMES = { card: { start: 'clock in', end: 'clock out' }, break: { start: 'break start', end: 'break end' } };

function edgeLine(span, edge, atMs) {
    return { text: timeFormat.format(atMs), note: EDGE_NAMES[spanKind(span)][edge], tone: spanTone(span) };
}

// An edge with a handle is a set time.
function edgeMs(span, edge) {
    const value = edge === 'start' ? span.dataset.startMs : span.dataset.endMs;
    assert(value !== undefined, 'a handle is on a set time');
    return Number(value);
}

function spanKind(span) {
    return span.classList.contains('break-overlay') ? 'break' : 'card';
}

function spanTone(span) {
    return span.classList.contains('is-long') ? 'danger' : spanKind(span);
}

// x: where on the window the tip points, which it is centred on unless that would leave the window.
function showTip(axis, x, lines) {
    if (!tip) {
        tip = document.createElement('div');
        tip.className = 'timeline-tip';
        tip.setAttribute('role', 'tooltip');
        document.body.appendChild(tip);
    }
    tip.replaceChildren(...lines.map(tipLine));
    tip.hidden = false;
    const rect = axis.getBoundingClientRect();
    const width = tip.offsetWidth;
    const height = tip.offsetHeight;
    const leftMax = Math.max(TIP_MARGIN_PX, window.innerWidth - TIP_MARGIN_PX - width);
    tip.style.left = `${Math.round(clamp(x - width / 2, TIP_MARGIN_PX, leftMax))}px`;
    // Above the axis, or below it where the window has no room above.
    const above = rect.top - TIP_GAP_PX - height;
    tip.style.top = `${Math.round(above >= TIP_MARGIN_PX ? above : rect.bottom + TIP_GAP_PX)}px`;
}

function tipLine({ key, text, note, tone }) {
    const line = document.createElement('div');
    line.className = 'tip-line';
    if (tone) line.dataset.tone = tone;
    if (key) line.append(Object.assign(document.createElement('kbd'), { textContent: key }));
    line.append(Object.assign(document.createElement('span'), { className: 'tip-text', textContent: text }));
    if (note) line.append(Object.assign(document.createElement('span'), { className: 'tip-note', textContent: note }));
    return line;
}

function hideTip() {
    if (tip) tip.hidden = true;
}

// Like a tab stop on a ruler: a marker at the top of the axis over the handle's time (the end of
// the axis for an edge past it), with a guide down through the axis while it is dragged. Answers
// the marker's x on the window.
function placeMarker(axis, span, atMs, dragging) {
    if (!marker) {
        marker = document.createElement('div');
        marker.className = 'edge-marker';
    }
    const percent = clamp(percentOfDay(axisDay(axis), atMs), 0, 100);
    marker.style.left = `${percent}%`;
    marker.dataset.tone = spanTone(span);
    marker.classList.toggle('is-dragging', dragging);
    if (marker.parentElement !== axis) axis.appendChild(marker);
    const rect = axis.getBoundingClientRect();
    return rect.left + (rect.width * percent) / 100;
}

function hideMarker() {
    marker?.remove();
}

// ── Dragging ────────────────────────────────────────────────────────────────

function startDrag(event, element, handle, handlers) {
    assert(!drag, 'one drag at a time');
    const card = handlers.cardById(element.dataset.cardId);
    if (!card) return;
    const axis = element.parentElement;
    drag = {
        element,
        handlers,
        card,
        breakIndex: element.dataset.breakIndex === undefined ? null : Number(element.dataset.breakIndex),
        edge: handle.dataset.edge,
        day: axisDay(axis),
        axisRect: axis.getBoundingClientRect(),
        result: null,
    };
    event.preventDefault();
    // Capture keeps the moves coming outside the window; events made by script have no pointer to
    // capture, so it is best effort. The listeners are on the window, where both kinds arrive.
    try { handle.setPointerCapture(event.pointerId); } catch { /* a synthetic pointer */ }
    window.addEventListener('pointermove', onDragMove);
    window.addEventListener('pointerup', onDragUp);
    window.addEventListener('pointercancel', onDragCancel);
    element.classList.add('is-dragging');
    document.body.classList.add('is-dragging');
    hover = null;
    showDragTip(edgeMs(element, drag.edge));
}

// The tip and the marker on the dragged edge, at the time it would be saved with.
function showDragTip(atMs) {
    const { element, edge } = drag;
    const axis = element.parentElement;
    showTip(axis, placeMarker(axis, element, atMs, true), [edgeLine(element, edge, atMs)]);
}

function onDragMove(event) {
    const { axisRect, day, card, breakIndex, edge, element } = drag;
    const fraction = (event.clientX - axisRect.left) / axisRect.width;
    const targetMs = day.dayStartMs + fraction * (day.dayEndMs - day.dayStartMs);
    const result = applyDrag(card, breakIndex, edge, targetMs, Date.now());
    if (!result) return;
    drag.result = result;
    const moved = breakIndex === null ? result.card : { ...result.card.breaks[breakIndex] };
    const startMs = new Date((breakIndex === null ? moved.clockIn : moved.start).dateTime).getTime();
    const endEvent = breakIndex === null ? moved.clockOut : moved.end;
    place(element, day, startMs, endEvent ? new Date(endEvent.dateTime).getTime() : Date.now());
    showDragTip(result.movedMs);
}

function onDragUp() {
    endDrag(true);
}

function onDragCancel() {
    endDrag(false);
}

function endDrag(commit) {
    const { element, handlers, result } = drag;
    window.removeEventListener('pointermove', onDragMove);
    window.removeEventListener('pointerup', onDragUp);
    window.removeEventListener('pointercancel', onDragCancel);
    element.classList.remove('is-dragging');
    document.body.classList.remove('is-dragging');
    drag = null;
    clearHover();
    if (commit && result) handlers.onCommit(result.card, result.movedMs);
    handlers.onEnd(commit && Boolean(result));
}

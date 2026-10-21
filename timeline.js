// One day's timeline: cards and their breaks drawn across the day (24 hours, or 23 or 25 on
// daylight saving days), and dragging their edges. A drag previews the snapped, clamped time and
// saves only on release; Escape cancels. While a drag runs the app holds back re-renders, so a
// background refresh cannot pull the block from under the pointer.
import { applyDrag, assert, cardEndMs, cardStartMs, dragBounds } from './core.js';

const HOUR_LABEL_EVERY_HOURS = 2;
const BLOCK_WIDTH_MIN_PERCENT = 0.15;

const hourFormat = new Intl.DateTimeFormat(undefined, { hour: 'numeric' });
const timeFormat = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });

let drag = null; // the drag in progress, see startDrag

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

function makeSpan(className, day, { cardId, breakIndex, startMs, endMs, card, nowMs, title }) {
    const element = document.createElement('div');
    element.className = className;
    element.dataset.cardId = cardId;
    if (breakIndex !== null) element.dataset.breakIndex = String(breakIndex);
    element.title = title;
    place(element, day, startMs, endMs);
    // A handle only on an edge inside this day that can move.
    for (const edge of ['start', 'end']) {
        const edgeMs = edge === 'start' ? startMs : endMs;
        const inDay = edgeMs >= day.dayStartMs && edgeMs <= day.dayEndMs;
        if (!inDay || !dragBounds(card, breakIndex, edge, nowMs)) continue;
        const handle = document.createElement('div');
        handle.className = 'resize-handle';
        handle.dataset.edge = edge;
        element.appendChild(handle);
    }
    return element;
}

// day: one entry of core's buildWeekDays.
export function renderTimeline(day, nowMs) {
    const axis = document.createElement('div');
    axis.className = 'timeline';
    axis.dataset.dayStartMs = String(day.dayStartMs);
    axis.dataset.dayEndMs = String(day.dayEndMs);
    for (let hour = HOUR_LABEL_EVERY_HOURS; hour < 24; hour += HOUR_LABEL_EVERY_HOURS) {
        const at = new Date(day.dayStartMs);
        at.setHours(hour);
        const label = document.createElement('div');
        label.className = 'hour-label';
        label.style.left = `${percentOfDay(day, at.getTime())}%`;
        label.textContent = hourFormat.format(at);
        axis.appendChild(label);
    }
    if (nowMs > day.dayStartMs && nowMs < day.dayEndMs) {
        const now = document.createElement('div');
        now.className = 'now-line';
        now.style.left = `${percentOfDay(day, nowMs)}%`;
        axis.appendChild(now);
    }
    day.drawn.forEach(card => appendCard(axis, day, card, nowMs));
    return axis;
}

function appendCard(axis, day, card, nowMs) {
    const startMs = cardStartMs(card);
    const endMs = cardEndMs(card, nowMs);
    const block = makeSpan('tc-block', day, {
        cardId: card.id, breakIndex: null, startMs, endMs, card, nowMs,
        title: `${timeFormat.format(startMs)} → ${card.clockOut ? timeFormat.format(endMs) : 'now'}`,
    });
    if (!card.clockOut) block.classList.add('is-running');
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
        if (breakEndMs <= day.dayStartMs || breakStartMs >= day.dayEndMs) return;
        const overlay = makeSpan('break-overlay', day, {
            cardId: card.id, breakIndex, startMs: breakStartMs, endMs: breakEndMs, card, nowMs,
            title: `Break ${timeFormat.format(breakStartMs)} → ${b.end ? timeFormat.format(breakEndMs) : 'now'}`,
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
        const day = { dayStartMs: Number(axis.dataset.dayStartMs), dayEndMs: Number(axis.dataset.dayEndMs) };
        const breakIndex = element.dataset.breakIndex;
        const startMs = breakIndex === undefined
            ? cardStartMs(card)
            : new Date(card.breaks[Number(breakIndex)].start.dateTime).getTime();
        place(element, day, startMs, nowMs);
    });
}

// ── Dragging ────────────────────────────────────────────────────────────────

// handlers: { cardById(id), onCommit(updatedCard, movedMs), onEnd(), onSelect(cardId), tooltip }
export function attachTimelineDrag(root, handlers) {
    root.addEventListener('pointerdown', event => {
        const span = event.target.closest('.tc-block, .break-overlay');
        if (!span) return;
        handlers.onSelect(span.dataset.cardId);
        const handle = event.target.closest('.resize-handle');
        if (handle && event.button === 0) startDrag(event, span, handle, handlers);
    });
    window.addEventListener('keydown', event => {
        if (event.key === 'Escape' && drag) endDrag(false);
    });
}

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
        day: { dayStartMs: Number(axis.dataset.dayStartMs), dayEndMs: Number(axis.dataset.dayEndMs) },
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
}

function onDragMove(event) {
    const { axisRect, day, card, breakIndex, edge, element, handlers } = drag;
    const fraction = (event.clientX - axisRect.left) / axisRect.width;
    const targetMs = day.dayStartMs + fraction * (day.dayEndMs - day.dayStartMs);
    const result = applyDrag(card, breakIndex, edge, targetMs, Date.now());
    if (!result) return;
    drag.result = result;
    const moved = breakIndex === null ? result.card : { ...result.card.breaks[breakIndex] };
    const startMs = new Date((breakIndex === null ? moved.clockIn : moved.start).dateTime).getTime();
    const endEvent = breakIndex === null ? moved.clockOut : moved.end;
    place(element, day, startMs, endEvent ? new Date(endEvent.dateTime).getTime() : Date.now());
    handlers.tooltip.show(event.clientX, event.clientY, timeFormat.format(result.movedMs));
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
    handlers.tooltip.hide();
    drag = null;
    if (commit && result) handlers.onCommit(result.card, result.movedMs);
    handlers.onEnd(commit && Boolean(result));
}

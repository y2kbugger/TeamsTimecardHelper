// One day's timeline: cards and their breaks drawn across the day (24 hours, or 23 or 25 on
// daylight saving days), dragging their edges, and adding breaks. A drag previews the snapped,
// clamped time and saves only on release; Escape cancels. While a drag runs the app holds back
// re-renders, so a background refresh cannot pull the block from under the pointer. Hovering a
// card shows a faint ghost of the break a Ctrl+click would add there, hovering free time one of a
// card; holding Ctrl arms it. A plain click only selects, so a click meant to select never adds
// anything to a real timesheet.
import { applyDrag, assert, cardEndMs, cardStartMs, dragBounds, isLongCard, proposeBreak, proposeCard } from './core.js';

const HOUR_LABEL_EVERY_HOURS = 2;
const BLOCK_WIDTH_MIN_PERCENT = 0.15;

const hourFormat = new Intl.DateTimeFormat(undefined, { hour: 'numeric' });
const timeFormat = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });

let drag = null; // the drag in progress, see startDrag
// { element, kind: 'break' | 'card', cardId (a break's card), proposal } while something new is
// offered. Never during a drag, so the click that follows a drag's release finds none and adds
// nothing.
let ghost = null;

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
function makeSpan(className, day, { cardId, breakIndex, startMs, endMs, card, nowMs, title, overflow }) {
    const element = document.createElement('div');
    element.className = className;
    element.dataset.cardId = cardId;
    if (breakIndex !== null) element.dataset.breakIndex = String(breakIndex);
    element.title = title;
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
    placeNowLine(axis, day, nowMs);
    day.drawn.forEach(card => appendCard(axis, day, card, nowMs));
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

function appendCard(axis, day, card, nowMs) {
    const startMs = cardStartMs(card);
    const endMs = cardEndMs(card, nowMs);
    const overflow = startMs < day.dayStartMs;
    const block = makeSpan('tc-block', day, {
        cardId: card.id, breakIndex: null, startMs, endMs, card, nowMs, overflow,
        title: `${timeFormat.format(startMs)} → ${card.clockOut ? timeFormat.format(endMs) : 'now'}`,
    });
    if (!card.clockOut) block.classList.add('is-running');
    if (isLongCard(card, nowMs)) {
        block.classList.add('is-long');
        block.title += ' · over 16 hours: check its clock-out';
    }
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
            cardId: card.id, breakIndex, startMs: breakStartMs, endMs: breakEndMs, card, nowMs, overflow,
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
        const day = axisDay(axis);
        const breakIndex = element.dataset.breakIndex;
        if (breakIndex === undefined) element.classList.toggle('is-long', isLongCard(card, nowMs));
        const startMs = breakIndex === undefined
            ? cardStartMs(card)
            : new Date(card.breaks[Number(breakIndex)].start.dateTime).getTime();
        place(element, day, startMs, nowMs);
    });
}

// ── Dragging ────────────────────────────────────────────────────────────────

// handlers: { cardById(id), cards(), onCommit(updatedCard, movedMs), onEnd(committed),
//             onSelect(cardId), onAddBreak(card, { startMs, endMs }), onAddCard({ startMs, endMs }),
//             tooltip }. cardById and cards answer nothing while no change may be made.
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
    root.addEventListener('pointermove', event => updateGhost(event, handlers));
    root.addEventListener('pointerleave', hideGhost);
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
    });
    window.addEventListener('keyup', event => {
        if (event.key === 'Control') armGhost(false);
    });
    window.addEventListener('blur', () => armGhost(false));
}

// Ctrl held: the ghost is what a click adds now, and says so.
function armGhost(armed) {
    document.body.classList.toggle('ctrl-held', armed);
    if (!ghost) return;
    const range = `${timeFormat.format(ghost.proposal.startMs)} → ${timeFormat.format(ghost.proposal.endMs)}`;
    ghost.element.classList.toggle('is-armed', armed);
    ghost.element.title = armed ? `+ ${ghost.kind} ${range}` : `Ctrl+click: ${ghost.kind} ${range}`;
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
    // Its label reaches toward the middle of the axis, so it never leaves the page.
    const middleMs = (proposal.startMs + proposal.endMs) / 2;
    element.dataset.align = middleMs < (day.dayStartMs + day.dayEndMs) / 2 ? 'start' : 'end';
    if (element.parentElement !== axis) axis.appendChild(element);
    Object.assign(ghost, { kind: offer.kind, cardId: offer.cardId, proposal });
    armGhost(event.ctrlKey);
}

function hideGhost() {
    ghost?.element.remove();
    ghost = null;
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

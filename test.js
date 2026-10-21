// Browser tests. Live (signed in, real Graph): probes of how the tenant's Graph behaves. Mock
// (?mock): the real app in an iframe against mockgraph.js, driven as a user would. On localhost
// the condensed result goes to the dev server (.agent/browser-tests.txt or mock-tests.txt).
import * as auth from './auth.js';
import { fetchJoinedTeams, fetchWeekTimeCards, STORAGE_KEYS } from './app.js';
import { addDays, formatDateInputValue, getWeekRange } from './core.js';
import { $, reportToDevServer } from './ui.js';
import { mockCards, mockEnabled, mockGraphFetch, MOCK_TEAM, seedMockDb } from './mockgraph.js';

const WAIT_TIMEOUT_MS = 5000;
const WAIT_POLL_MS = 20;

function check(condition, message) {
    if (!condition) throw new Error(message);
}

// ── Live: Graph as this tenant behaves ──────────────────────────────────────

const timeCardsUrl = (teamId, query) => `/teams/${teamId}/schedule/timeCards?${query}`;
const isRejection = message => /filter|unsupported|not allowed|could not find a property|invalid|orderby|order by|sort/i.test(message || '');

// true: Graph accepted the query; false: it refused it as unsupported; anything else throws.
async function graphAccepts(url) {
    try {
        await auth.graphFetch(url);
        return true;
    } catch (error) {
        if (isRejection(error.message)) return false;
        throw error;
    }
}

const filterAccepted = (teamId, filter) => graphAccepts(timeCardsUrl(teamId, `$top=1&$filter=${encodeURIComponent(filter)}`));

async function liveTeam(context) {
    context.team ||= await (async () => {
        const teams = await fetchJoinedTeams();
        const savedId = context.params.teamId || localStorage.getItem(STORAGE_KEYS.selectedTeamId);
        const team = teams.find(t => t.id === savedId) || teams[0];
        check(team, 'the account is in a team');
        return team;
    })();
    return context.team;
}

const liveTests = [
    {
        name: 'auth-status',
        run: async context => ({ account: context.session.account?.username || null, redirectUri: auth.getRuntimeConfig().redirectUri }),
    },
    {
        name: 'joined-teams',
        run: async () => {
            const teams = await fetchJoinedTeams();
            check(teams.length, 'at least one team');
            return { count: teams.length };
        },
    },
    {
        name: 'documented-filters-accepted',
        run: async context => {
            const team = await liveTeam(context);
            const now = Date.now();
            const before = new Date(now - 30 * 864e5).toISOString();
            const after = new Date(now + 864e5).toISOString();
            const filters = [
                `clockInEvent/dateTime ge ${before}`, `clockInEvent/dateTime le ${after}`,
                `clockOutEvent/dateTime ge ${before}`, `clockOutEvent/dateTime le ${after}`,
                "state eq 'clockedIn'",
            ];
            const refused = [];
            for (const filter of filters) if (!await filterAccepted(team.id, filter)) refused.push(filter);
            check(!refused.length, `refused: ${refused.join('; ')}`);
            return { accepted: filters.length };
        },
    },
    {
        name: 'or-in-filter-refused',
        run: async context => {
            const team = await liveTeam(context);
            const { weekStart, weekEndInclusive } = getWeekRange(new Date());
            const filter = `(clockInEvent/dateTime le ${weekEndInclusive.toISOString()} and clockOutEvent/dateTime ge ${weekStart.toISOString()}) or state eq 'clockedIn'`;
            check(!await filterAccepted(team.id, filter), 'Graph now accepts `or`: the week could be one overlap query');
            return { refused: true };
        },
    },
    {
        name: 'last-modified-unqueryable',
        run: async context => {
            const team = await liveTeam(context);
            check(!await filterAccepted(team.id, 'lastModifiedDateTime ge 1970-01-01T00:00:00Z'), 'lastModifiedDateTime filters now work');
            check(!await graphAccepts(timeCardsUrl(team.id, `$top=1&$orderby=${encodeURIComponent('lastModifiedDateTime desc')}`)), 'lastModifiedDateTime ordering now works');
            return { refused: true };
        },
    },
    {
        name: 'week-fetch',
        run: async context => {
            const team = await liveTeam(context);
            const startedAt = performance.now();
            const week = await fetchWeekTimeCards(team.id, context.params.week || null);
            return {
                week: formatDateInputValue(week.weekStart),
                cards: week.cards.length,
                pages: week.pageCount,
                durationMs: Math.round(performance.now() - startedAt),
                running: Boolean(week.activeCard),
            };
        },
    },
];

// ── Mock: the app, driven ───────────────────────────────────────────────────

async function waitFor(describe, predicate) {
    const deadline = performance.now() + WAIT_TIMEOUT_MS;
    for (;;) {
        const value = predicate();
        if (value) return value;
        if (performance.now() > deadline) throw new Error(`Timed out waiting for ${describe}`);
        await new Promise(resolve => setTimeout(resolve, WAIT_POLL_MS));
    }
}

// A fresh app on freshly seeded mock data, showing this week (or `week`), its first render done.
async function openApp({ week = null, beforeLoad = async () => { } } = {}) {
    seedMockDb();
    sessionStorage.removeItem(auth.MOCK_EXPIRED_KEY);
    sessionStorage.removeItem(STORAGE_KEYS.joinedTeams);
    sessionStorage.removeItem(STORAGE_KEYS.selectedWeek);
    sessionStorage.removeItem(STORAGE_KEYS.lastActivity);
    if (week) {
        sessionStorage.setItem(STORAGE_KEYS.selectedWeek, formatDateInputValue(week));
        sessionStorage.setItem(STORAGE_KEYS.lastActivity, String(Date.now()));
    }
    localStorage.setItem(STORAGE_KEYS.selectedTeamId, MOCK_TEAM.id);
    await beforeLoad();
    const frame = document.createElement('iframe');
    frame.className = 'test-app-frame';
    frame.src = 'index.html';
    document.body.appendChild(frame);
    await new Promise(resolve => frame.addEventListener('load', resolve, { once: true }));
    const app = new AppDriver(frame);
    await app.settled('the first render');
    return app;
}

class AppDriver {
    constructor(frame) {
        this.frame = frame;
        this.win = frame.contentWindow;
        this.doc = frame.contentDocument;
    }

    $(selector) { return this.doc.querySelector(selector); }

    $$(selector) { return [...this.doc.querySelectorAll(selector)]; }

    clockLabel() { return this.doc.getElementById('btn-clock').textContent; }

    weekLabel() { return this.doc.getElementById('week-range-label').textContent; }

    // Nothing on its way to Graph and the week drawn.
    settled(describe) {
        return waitFor(describe, () => this.$('#busy')?.hidden && !this.$('#weeks')?.hidden);
    }

    click(selector) {
        const element = typeof selector === 'string' ? this.$(selector) : selector;
        check(element, `an element for ${selector}`);
        check(!element.disabled, `${selector} is enabled`);
        element.dispatchEvent(new this.win.PointerEvent('pointerdown', { bubbles: true }));
        element.click();
    }

    async previousWeek() {
        const label = this.weekLabel();
        this.click('#btn-week-prev');
        await waitFor('the previous week', () => this.weekLabel() !== label);
        await this.settled('the previous week to load');
    }

    pointer(type, target, clientX) {
        target.dispatchEvent(new this.win.PointerEvent(type, { bubbles: true, cancelable: true, clientX, clientY: 10, button: 0, pointerId: 7, isPrimary: true }));
    }

    // The x on a timeline for a local time of its day.
    xAt(timeline, hours, minutes) {
        const dayStartMs = Number(timeline.dataset.dayStartMs);
        const dayEndMs = Number(timeline.dataset.dayEndMs);
        const at = new Date(dayStartMs);
        at.setHours(hours, minutes, 0, 0);
        const rect = timeline.getBoundingClientRect();
        return rect.left + rect.width * ((at.getTime() - dayStartMs) / (dayEndMs - dayStartMs));
    }

    // Presses an edge's handle and moves it to hh:mm of its day, without letting go.
    grab(span, edge, hours, minutes) {
        const handle = span.querySelector(`.resize-handle[data-edge="${edge}"]`);
        check(handle, `a ${edge} handle`);
        const rect = handle.getBoundingClientRect();
        this.pointer('pointerdown', handle, rect.left + rect.width / 2);
        this.pointer('pointermove', handle, this.xAt(span.parentElement, hours, minutes));
        return handle;
    }

    close() { this.frame.remove(); }
}

async function withApp(run, options) {
    const app = await openApp(options);
    try { return await run(app); }
    finally { app.close(); }
}

const openMockCard = () => mockCards().find(card => card.state !== 'clockedOut');
const localTime = iso => { const d = new Date(iso); return `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`; };
const lastWeek = () => addDays(new Date(), -7);

// The seeded card of a day last week (09:00-12:00, a break 10:30-10:45) and its timeline block.
function seededBlock(app) {
    const block = app.$('.tc-block[data-card-id^="seed-"]');
    check(block, 'a seeded card on the timeline');
    return { block, id: block.dataset.cardId, card: () => mockCards().find(c => c.id === block.dataset.cardId) };
}

const mockTests = [
    {
        name: 'clock-in-break-clock-out',
        run: () => withApp(async app => {
            check(app.clockLabel() === 'Clock In', `starts clocked out, not ${app.clockLabel()}`);
            app.click('#btn-clock');
            await waitFor('Clock Out', () => app.clockLabel() === 'Clock Out');
            check(openMockCard()?.state === 'clockedIn', 'the backend has an open card');
            app.click(await waitFor('Start Break', () => app.$('[data-action="startBreak"]')));
            await waitFor('On Break', () => app.clockLabel() === 'On Break');
            check(app.$('#btn-clock').disabled, 'no clocking out on a break');
            app.click(await waitFor('End Break', () => app.$('[data-action="endBreak"]')));
            await waitFor('Clock Out again', () => app.clockLabel() === 'Clock Out');
            app.click('#btn-clock');
            await waitFor('Clock In again', () => app.clockLabel() === 'Clock In');
            const card = mockCards().at(-1);
            check(card.state === 'clockedOut' && card.breaks.length === 1 && card.breaks[0].end, `ends clocked out with one closed break: ${JSON.stringify(card)}`);
            return {};
        }),
    },
    {
        name: 'past-week-still-shows-the-open-card',
        // The app starts on last week (a reload while there), so it never saw this week's card.
        run: () => withApp(async app => {
            await waitFor('Clock Out on a past week', () => app.clockLabel() === 'Clock Out');
            return {};
        }, { week: lastWeek(), beforeLoad: () => mockGraphFetch(`/teams/${MOCK_TEAM.id}/schedule/timeCards/clockIn`, { method: 'POST', body: '{}' }) }),
    },
    {
        name: 'clock-in-from-a-past-week',
        run: () => withApp(async app => {
            const cards = app.$$('.card').length;
            app.click('#btn-clock');
            await waitFor('Clock Out', () => app.clockLabel() === 'Clock Out');
            await app.settled('the reload');
            check(app.$$('.card').length === cards, 'the past week still shows its own cards');
            return { cards };
        }, { week: lastWeek() }),
    },
    {
        name: 'edit-refuses-a-break-before-clock-in',
        run: () => withApp(async app => {
            const before = JSON.stringify(mockCards());
            app.click('.card [data-action="edit"]');
            await waitFor('the edit dialog', () => !app.$('#edit-modal').hidden);
            app.$('.break-edit-row .break-start').value = `${app.$('#edit-clock-in').value.slice(0, 11)}00:05`;
            app.click('#btn-edit-save');
            await waitFor('the error', () => app.$$('.toast.error').some(t => /Break start/.test(t.textContent)));
            check(JSON.stringify(mockCards()) === before, 'nothing saved');
            check(!app.$('#edit-modal').hidden, 'the dialog stays open');
            return {};
        }),
    },
    {
        name: 'edit-saves-notes-and-keeps-break-ids',
        run: () => withApp(async app => {
            const id = app.$('.card').dataset.cardId;
            app.click('.card [data-action="edit"]');
            await waitFor('the edit dialog', () => !app.$('#edit-modal').hidden);
            app.$('#edit-notes').value = 'flow note';
            app.click('#btn-edit-save');
            await waitFor('the dialog to close', () => app.$('#edit-modal').hidden);
            const card = mockCards().find(c => c.id === id);
            check(card.notes?.content === 'flow note', `notes saved: ${JSON.stringify(card.notes)}`);
            check(card.breaks.length === 1 && card.breaks[0].breakId.startsWith('seed-'), 'the break kept its id');
            await waitFor('the note on the card', () => app.$(`.card[data-card-id="${id}"] .card-notes`)?.textContent === 'flow note');
            return {};
        }),
    },
    {
        name: 'escape-closes-the-edit-dialog-unsaved',
        run: () => withApp(async app => {
            const before = JSON.stringify(mockCards());
            app.click('.card [data-action="edit"]');
            await waitFor('the edit dialog', () => !app.$('#edit-modal').hidden);
            app.$('#edit-notes').value = 'never saved';
            app.doc.dispatchEvent(new app.win.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
            check(app.$('#edit-modal').hidden, 'closed');
            check(JSON.stringify(mockCards()) === before, 'nothing saved');
            return {};
        }),
    },
    {
        name: 'delete-card',
        run: () => withApp(async app => {
            const id = app.$('.card').dataset.cardId;
            app.click('.card [data-action="delete"]');
            await waitFor('the confirm dialog', () => !app.$('#confirm-modal').hidden);
            app.click('#btn-confirm-ok');
            await waitFor('the card to go', () => !app.$(`.card[data-card-id="${id}"]`));
            check(!mockCards().some(c => c.id === id), 'gone from the backend');
            return {};
        }),
    },
    {
        name: 'delete-break',
        run: () => withApp(async app => {
            const id = app.$('.card').dataset.cardId;
            app.click(`.card[data-card-id="${id}"] [data-action="delete-break"]`);
            await waitFor('the confirm dialog', () => !app.$('#confirm-modal').hidden);
            app.click('#btn-confirm-ok');
            await waitFor('the break to go', () => !mockCards().find(c => c.id === id).breaks.length);
            await waitFor('the break row to go', () => !app.$(`.card[data-card-id="${id}"] .break-row`));
            return {};
        }),
    },
    {
        name: 'drag-card-end',
        run: () => withApp(async app => {
            const { block, card } = seededBlock(app);
            app.grab(block, 'end', 12, 31);
            check(app.$('#drag-tooltip').textContent.includes('12:30'), `the tooltip shows the snapped time: ${app.$('#drag-tooltip').textContent}`);
            app.pointer('pointerup', block, 0);
            await waitFor('the new clock-out', () => localTime(card().clockOutEvent.dateTime) === '12:30');
            check(localTime(card().clockInEvent.dateTime) === '9:00', 'clock-in unchanged');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'drag-break-start',
        run: () => withApp(async app => {
            const { id, card } = seededBlock(app);
            const overlay = app.$(`.break-overlay[data-card-id="${id}"]`);
            app.grab(overlay, 'start', 10, 19);
            app.pointer('pointerup', overlay, 0);
            await waitFor('the new break start', () => localTime(card().breaks[0].start.dateTime) === '10:20');
            check(localTime(card().breaks[0].end.dateTime) === '10:45', 'break end unchanged');
            check(card().breaks[0].breakId.startsWith('seed-'), 'the break kept its id');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'drag-escape-cancels',
        run: () => withApp(async app => {
            const before = JSON.stringify(mockCards());
            const { block } = seededBlock(app);
            const width = block.style.width;
            app.grab(block, 'end', 14, 0);
            check(block.style.width !== width, 'the block follows the pointer');
            app.win.dispatchEvent(new app.win.KeyboardEvent('keydown', { key: 'Escape' }));
            await app.settled('the redraw');
            check(JSON.stringify(mockCards()) === before, 'nothing saved');
            check(app.$(`.tc-block[data-card-id="${block.dataset.cardId}"]`).style.width === width, 'back where it was');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'drag-survives-a-background-refresh',
        run: () => withApp(async app => {
            const { block, card } = seededBlock(app);
            app.grab(block, 'end', 13, 0);
            app.doc.dispatchEvent(new Event('visibilitychange'));
            await app.settled('the refresh');
            check(block.isConnected, 'the dragged block is still on the page');
            app.pointer('pointerup', block, 0);
            await waitFor('the new clock-out', () => localTime(card().clockOutEvent.dateTime) === '13:00');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'no-second-change-until-the-first-is-read-back',
        // A second drag built from the card as it was would put the first one's edge back.
        run: () => withApp(async app => {
            const { block, card } = seededBlock(app);
            app.grab(block, 'end', 12, 30);
            app.pointer('pointerup', block, 0);
            check(app.$('#btn-clock').disabled, 'the clock button waits for the save');
            const start = block.querySelector('.resize-handle[data-edge="start"]');
            app.pointer('pointerdown', start, start.getBoundingClientRect().left);
            app.pointer('pointermove', start, app.xAt(block.parentElement, 8, 30));
            check(!app.$('.is-dragging'), 'no second drag while the first is saving');
            app.pointer('pointerup', start, 0);
            await app.settled('the save and its read-back');
            check(localTime(card().clockOutEvent.dateTime) === '12:30', `the first change holds: ${card().clockOutEvent.dateTime}`);
            check(localTime(card().clockInEvent.dateTime) === '9:00', 'the second never happened');
            check(!app.$('#btn-clock').disabled, 'the clock button is back');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'ctrl-click-adds-the-ghost-break',
        run: () => withApp(async app => {
            const { block, card } = seededBlock(app);
            app.pointer('pointermove', block, app.xAt(block.parentElement, 11, 22));
            const ghost = app.$('.break-ghost');
            check(ghost?.title.includes('11:15') && ghost.title.includes('11:25'), `the ghost offers 11:15-11:25: ${ghost?.title}`);
            check(/Ctrl\+click/.test(ghost.title) && !ghost.classList.contains('is-armed'), `a hint until Ctrl is held: ${ghost.title}`);
            app.win.dispatchEvent(new app.win.KeyboardEvent('keydown', { key: 'Control', ctrlKey: true }));
            check(ghost.classList.contains('is-armed') && ghost.title.startsWith('+ break'), `armed with Ctrl: ${ghost.title}`);
            block.dispatchEvent(new app.win.MouseEvent('click', { bubbles: true, ctrlKey: true }));
            await waitFor('the new break', () => card().breaks.length === 2);
            const added = card().breaks.find(b => localTime(b.start.dateTime) !== '10:30');
            check(`${localTime(added.start.dateTime)}-${localTime(added.end.dateTime)}` === '11:15-11:25', `added 11:15-11:25: ${JSON.stringify(added)}`);
            await waitFor('the break on the timeline', () => app.$$(`.break-overlay[data-card-id="${block.dataset.cardId}"]`).length === 2);
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'a-plain-click-adds-no-break',
        run: () => withApp(async app => {
            const { block, card } = seededBlock(app);
            app.pointer('pointermove', block, app.xAt(block.parentElement, 11, 22));
            check(app.$('.break-ghost'), 'the hint shows');
            block.click();
            await app.settled('anything the click set off');
            check(card().breaks.length === 1, 'still one break');
            check(block.classList.contains('is-highlighted') || app.$(`.tc-block.is-highlighted[data-card-id="${block.dataset.cardId}"]`), 'the click selected the card');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'no-ghost-over-a-break',
        run: () => withApp(async app => {
            const { id } = seededBlock(app);
            const overlay = app.$(`.break-overlay[data-card-id="${id}"]`);
            app.pointer('pointermove', overlay, app.xAt(overlay.parentElement, 10, 35));
            check(!app.$('.break-ghost'), 'no ghost on an existing break');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'past-midnight-is-a-ghost-handled-at-home',
        // Last week's Tuesday 22:00 to Wednesday 00:30, a break 23:00-23:15.
        run: () => withApp(async app => {
            const id = 'seed-1';
            const blocks = app.$$(`.tc-block[data-card-id="${id}"]`);
            check(blocks.length === 2, `drawn on both days: ${blocks.length}`);
            const home = blocks.find(b => !b.classList.contains('is-overflow'));
            const ghost = blocks.find(b => b !== home);
            check(ghost.classList.contains('is-overflow') && !ghost.querySelector('.resize-handle'), 'Wednesday has an untouchable ghost');
            check(getComputedStyle(ghost).pointerEvents === 'none', 'the ghost takes no pointer');
            const end = home.querySelector('.resize-handle[data-edge="end"]');
            const axisRight = home.parentElement.getBoundingClientRect().right;
            check(end && Math.abs(end.getBoundingClientRect().right - axisRight) < 3, 'Tuesday has its end handle at the end of the axis');
            app.pointer('pointermove', home, app.xAt(home.parentElement, 23, 50));
            check(app.$('.break-ghost'), 'a break offered near midnight');
            check(app.doc.documentElement.scrollWidth <= app.doc.documentElement.clientWidth, 'its label stays on the page');
            app.grab(home, 'end', 23, 30);
            app.pointer('pointerup', home, 0);
            const card = () => mockCards().find(c => c.id === id);
            await waitFor('the new clock-out', () => localTime(card().clockOutEvent.dateTime) === '23:30');
            return {};
        }, {
            week: lastWeek(),
            beforeLoad: () => {
                const tuesday = addDays(getWeekRange(lastWeek()).weekStart, 2);
                const at = (days, h, m) => { const d = addDays(tuesday, days); d.setHours(h, m, 0, 0); return { dateTime: d.toISOString() }; };
                return mockGraphFetch(`/teams/${MOCK_TEAM.id}/schedule/timeCards/seed-1`, {
                    method: 'PUT',
                    body: JSON.stringify({ clockInEvent: at(0, 22, 0), clockOutEvent: at(1, 0, 30), breaks: [{ breakId: 'seed-1-b', start: at(0, 23, 0), end: at(0, 23, 15) }] }),
                });
            },
        }),
    },
    {
        name: 'the-click-ending-a-drag-adds-no-break',
        run: () => withApp(async app => {
            const { block, card } = seededBlock(app);
            app.grab(block, 'end', 12, 30);
            // Over the block itself, where a ghost would be offered were it not mid-drag.
            app.pointer('pointermove', block, app.xAt(block.parentElement, 11, 30));
            check(!app.$('.break-ghost'), 'no ghost during a drag');
            app.pointer('pointerup', block, 0);
            block.dispatchEvent(new app.win.MouseEvent('click', { bubbles: true, ctrlKey: true }));
            await waitFor('the drag saved', () => localTime(card().clockOutEvent.dateTime) === '11:30');
            await app.settled('the reload');
            check(card().breaks.length === 1, 'still one break');
            return {};
        }, { week: lastWeek() }),
    },
    {
        name: 'expired-session-finishes-the-click',
        run: () => withApp(async app => {
            sessionStorage.setItem(auth.MOCK_EXPIRED_KEY, '1');
            const cards = mockCards().length;
            app.click('#btn-clock');
            await waitFor('the session bar', () => app.$('#session-bar button'));
            check(mockCards().length === cards, 'nothing happens until signed in');
            check(app.$('#btn-clock').disabled, 'the clicked button waits');
            check(app.doc.body.classList.contains('session-expired'), 'the shown data is marked stale');
            app.click('#session-bar button');
            await waitFor('Clock Out after signing in', () => app.clockLabel() === 'Clock Out');
            check(!app.$('#session-bar') && !app.doc.body.classList.contains('session-expired'), 'the bar and the mark are gone');
            return {};
        }),
    },
    {
        name: 'expired-session-shows-on-return',
        run: () => withApp(async app => {
            sessionStorage.setItem(auth.MOCK_EXPIRED_KEY, '1');
            app.doc.dispatchEvent(new Event('visibilitychange'));
            await waitFor('the session bar without a click', () => app.$('#session-bar'));
            app.click('#session-bar button');
            await waitFor('the bar to go', () => !app.$('#session-bar'));
            return {};
        }),
    },
    {
        name: 'idle-returns-to-this-week',
        run: () => withApp(async app => {
            const thisWeek = app.weekLabel();
            await app.previousWeek();
            sessionStorage.setItem(STORAGE_KEYS.lastActivity, String(Date.now() - 2 * 3600 * 1000));
            app.doc.dispatchEvent(new Event('visibilitychange'));
            await waitFor('this week again', () => app.weekLabel() === thisWeek);
            return {};
        }),
    },
];

// ── The runner ──────────────────────────────────────────────────────────────

function setAuthStatus(session) {
    $('test-auth-status').textContent = session.authenticated ? `Signed in: ${session.account?.username}` : 'Not signed in';
    $('btn-test-signin').hidden = session.authenticated;
    $('btn-test-signout').hidden = !session.authenticated;
}

function resultRow(name) {
    const row = document.createElement('section');
    row.className = 'test-result test-result-running';
    row.innerHTML = '<div class="test-result-name"></div><div class="test-result-body">running…</div>';
    row.firstElementChild.textContent = name;
    $('test-results').appendChild(row);
    return (status, details) => {
        row.className = `test-result test-result-${status}`;
        row.lastElementChild.textContent = typeof details === 'string' ? details : JSON.stringify(details, null, 2);
    };
}

async function runSuite() {
    const params = Object.fromEntries(new URLSearchParams(location.search));
    const tests = (mockEnabled ? mockTests : liveTests).filter(test => !params.test || test.name === params.test);
    $('test-meta').textContent = `${mockEnabled ? 'Mock' : 'Live'} tests. Run one with ?test=<name>: ${tests.map(t => t.name).join(', ')}.`;
    $('test-results').replaceChildren();
    const session = await auth.ensureSession();
    setAuthStatus(session);

    const context = { params, session, team: null };
    const startedAt = performance.now();
    const problems = [];
    let passed = 0;
    for (const test of tests) {
        const finish = resultRow(test.name);
        if (!session.authenticated) {
            problems.push(`SKIP ${test.name}: not signed in`);
            finish('skip', 'Not signed in.');
            continue;
        }
        try {
            finish('pass', await test.run(context));
            passed += 1;
        } catch (error) {
            problems.push(`FAIL ${test.name}: ${error.message}`);
            finish('fail', error.stack || error.message);
        }
    }
    const failed = problems.filter(line => line.startsWith('FAIL')).length;
    const skipped = problems.length - failed;
    const mode = mockEnabled ? 'mock' : 'live';
    const summary = `${failed ? 'FAIL' : 'ok'} ${mode} browser tests: ${passed}/${tests.length} passed, ${failed} failed, ${skipped} skipped, ${((performance.now() - startedAt) / 1000).toFixed(1)}s`;
    $('test-summary').textContent = summary;
    reportToDevServer(mockEnabled ? 'mock-tests' : 'browser-tests', [summary, ...problems].join('\n'));
}

$('btn-test-signin').addEventListener('click', () => void auth.signIn({ navigateToApp: false }).finally(runSuite));
$('btn-test-signout').addEventListener('click', () => void auth.signOut({ navigateToLogin: false }).finally(runSuite));
void runSuite();

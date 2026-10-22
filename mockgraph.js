// An in-memory stand-in for the Graph endpoints the app uses, so the app can be driven
// (clock in/out, breaks, drags, edits, deletes) without touching real timecards.
// On with ?mock in the URL (kept for the tab in sessionStorage), off with ?mock=0.
// Behaves like the tenant where it matters: action endpoints answer 204 No Content, PUT
// answers 204, and filters accept only `clockInEvent/dateTime ge|le … and …`.

const MODE_KEY = 'tc_mock';
const DB_KEY = 'tc_mock_db';
export const MOCK_TEAM = { id: 'mock-team', displayName: 'Mock Team', description: '' };
export const MOCK_ACCOUNT = { name: 'Mock User', username: 'mock@localhost', homeAccountId: 'mock', localAccountId: 'mock-user' };

const modeParam = new URLSearchParams(location.search).get('mock');
if (modeParam !== null) {
    if (modeParam === '0') sessionStorage.removeItem(MODE_KEY);
    else sessionStorage.setItem(MODE_KEY, '1');
}

export const mockEnabled = location.hostname === 'localhost' && sessionStorage.getItem(MODE_KEY) === '1';

// Cards seeded relative to now: 09:00-12:00 with a 15 minute break on each of the last 10 days,
// and one that ended half an hour ago today, so this week is never empty (past 01:00).
export function seedMockDb(now = new Date()) {
    const cards = [];
    const iso = ms => new Date(ms).toISOString();
    for (let daysAgo = 10; daysAgo >= 1; daysAgo -= 1) {
        const day = new Date(now);
        day.setDate(day.getDate() - daysAgo);
        const at = (h, m = 0) => { const d = new Date(day); d.setHours(h, m, 0, 0); return d.getTime(); };
        cards.push(makeCard(`seed-${daysAgo}`, iso(at(9)), iso(at(12)), [{ breakId: `seed-${daysAgo}-b`, start: { dateTime: iso(at(10, 30)) }, end: { dateTime: iso(at(10, 45)) } }]));
    }
    const midnight = new Date(now).setHours(0, 0, 0, 0);
    const startMs = Math.max(midnight + 60e3, now.getTime() - 90 * 60e3);
    const endMs = now.getTime() - 30 * 60e3;
    if (endMs - startMs >= 20 * 60e3) {
        const breakMs = startMs + Math.floor((endMs - startMs) / 2);
        cards.push(makeCard('seed-today', iso(startMs), iso(endMs), [{ breakId: 'seed-today-b', start: { dateTime: iso(breakMs) }, end: { dateTime: iso(breakMs + 5 * 60e3) } }]));
    }
    writeDb({ nextId: 1, cards });
}

function makeCard(id, clockIn, clockOut = null, breaks = []) {
    const card = {
        id,
        userId: MOCK_ACCOUNT.localAccountId,
        createdDateTime: clockIn,
        lastModifiedDateTime: clockIn,
        clockInEvent: { dateTime: clockIn, atApprovedLocation: false, notes: null },
        clockOutEvent: clockOut ? { dateTime: clockOut, atApprovedLocation: false, notes: null } : null,
        breaks,
        notes: null,
    };
    card.state = stateOf(card);
    return card;
}

function stateOf(card) {
    if (card.clockOutEvent) return 'clockedOut';
    return card.breaks.some(b => b.start && !b.end) ? 'onBreak' : 'clockedIn';
}

function readDb() {
    const raw = sessionStorage.getItem(DB_KEY);
    if (!raw) seedMockDb();
    return JSON.parse(sessionStorage.getItem(DB_KEY));
}

function writeDb(db) {
    sessionStorage.setItem(DB_KEY, JSON.stringify(db));
}

function graphError(status, message, code = 'BadRequest') {
    const error = new Error(message);
    error.status = status;
    error.code = code;
    return error;
}

function parseClockInFilter(filter) {
    const match = /^clockInEvent\/dateTime ge (\S+) and clockInEvent\/dateTime le (\S+)$/.exec(filter);
    if (!match) throw graphError(400, `Mock Graph does not support this filter: ${filter}`);
    return { geMs: new Date(match[1]).getTime(), leMs: new Date(match[2]).getTime() };
}

export async function mockGraphFetch(url, options = {}) {
    const method = (options.method || 'GET').toUpperCase();
    const { pathname, searchParams } = new URL(url, 'https://graph.microsoft.com');
    const path = pathname.replace(/^\/(v1\.0|beta)/, '');
    const body = options.body ? JSON.parse(options.body) : null;
    const db = readDb();
    const now = new Date().toISOString();
    const nowEvent = () => ({ dateTime: now, atApprovedLocation: Boolean(body?.atApprovedLocation), notes: body?.notes || null });
    await Promise.resolve(); // answer later, as a network does, but on no timer

    if (method === 'GET' && path === '/me/joinedTeams') return { value: [MOCK_TEAM] };

    const match = /^\/teams\/([^/]+)\/schedule\/timeCards(?:\/([^/]+))?(?:\/([^/]+))?$/.exec(path);
    if (!match || match[1] !== MOCK_TEAM.id) throw graphError(404, `Mock Graph has no route for ${method} ${path}`, 'NotFound');
    const [, , idOrAction, action] = match;

    if (!idOrAction && method === 'POST') {
        if (!body?.clockInEvent?.dateTime) throw graphError(400, 'clockInEvent is required.');
        const card = makeCard(`mock-${db.nextId++}`, body.clockInEvent.dateTime, body.clockOutEvent?.dateTime || null, body.breaks || []);
        db.cards.push(card);
        writeDb(db);
        return card; // 201 Created, with the card
    }

    if (!idOrAction) {
        if (method !== 'GET') throw graphError(405, 'Method not allowed');
        const top = Number(searchParams.get('$top') || 100);
        const skip = Number(searchParams.get('$skip') || 0);
        const filter = searchParams.get('$filter');
        let cards = db.cards;
        if (filter) {
            const { geMs, leMs } = parseClockInFilter(filter);
            cards = cards.filter(c => { const t = new Date(c.clockInEvent.dateTime).getTime(); return t >= geMs && t <= leMs; });
        }
        const page = cards.slice(skip, skip + top);
        const next = new URL(url, 'https://graph.microsoft.com');
        next.searchParams.set('$skip', String(skip + top));
        return { value: page, ...(skip + top < cards.length ? { '@odata.nextLink': next.href } : {}) };
    }

    if (idOrAction === 'clockIn' && method === 'POST') {
        if (db.cards.some(c => c.state !== 'clockedOut')) throw graphError(400, 'User is already clocked in.');
        db.cards.push(makeCard(`mock-${db.nextId++}`, now));
        writeDb(db);
        return null;
    }

    const card = db.cards.find(c => c.id === idOrAction);
    if (!card) throw graphError(404, `Timecard ${idOrAction} not found.`, 'NotFound');

    if (!action) {
        if (method === 'GET') return card;
        if (method === 'DELETE') {
            db.cards = db.cards.filter(c => c !== card);
            writeDb(db);
            return null;
        }
        if (method === 'PUT') {
            if (!body?.clockInEvent?.dateTime) throw graphError(400, 'clockInEvent is required.');
            Object.assign(card, {
                clockInEvent: body.clockInEvent,
                clockOutEvent: body.clockOutEvent || null,
                breaks: (body.breaks || []).map((b, i) => ({ breakId: b.breakId || `${card.id}-b${Date.now()}-${i}`, start: b.start || null, end: b.end || null })),
                notes: body.notes || null,
                lastModifiedDateTime: now,
            });
            card.state = stateOf(card);
            writeDb(db);
            return null;
        }
        throw graphError(405, 'Method not allowed');
    }

    if (method !== 'POST') throw graphError(405, 'Method not allowed');
    const openBreak = card.breaks.find(b => b.start && !b.end);
    if (action === 'clockOut') {
        if (card.state !== 'clockedIn') throw graphError(400, `Cannot clock out while ${card.state}.`);
        card.clockOutEvent = nowEvent();
    } else if (action === 'startBreak') {
        if (card.state !== 'clockedIn') throw graphError(400, `Cannot start a break while ${card.state}.`);
        card.breaks.push({ breakId: `${card.id}-b${card.breaks.length + 1}`, start: nowEvent(), end: null });
    } else if (action === 'endBreak') {
        if (!openBreak) throw graphError(400, 'No break to end.');
        openBreak.end = nowEvent();
    } else {
        throw graphError(404, `Unknown action ${action}`, 'NotFound');
    }
    card.state = stateOf(card);
    card.lastModifiedDateTime = now;
    writeDb(db);
    return null;
}

// For tests: every card, newest clock-in last.
export function mockCards() {
    return readDb().cards;
}

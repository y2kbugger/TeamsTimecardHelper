# Agent Notes

## User Preferences

- Keep this as a static app. Do not introduce a build step or install Node packages.
- Code is split into native ES modules (no bundler). Browsers load them directly via `<script type="module">`. They will not load over `file://` — always use `serve.sh` (http://localhost:8999).
- Node is for checks only: `node --check` and the built-in `node --test` runner, with no packages.
- Prefer simpler DOM over decorative structure. Resize handles should stay square, thin, and minimal.
- Keep structure split and simple:
  - `common.css` for shared styles
  - `timecards.css` for app-specific styles
  - `login.html` for sign-in
  - `index.html` for the app shell
  - `auth.js` for auth/session concerns
  - `app.js` for app behavior
- Follow the Boy Scout Rule: leave the code better than you found it.
- Do not change how the app works or looks at a high level unless simplification requires it.
- Do not keep OAuth app registration configurable at runtime. Use `config.js` and remove setup UI.
- Maintain a browser-run self-test harness with `test.html` and `test.js`.
- Self-tests should be easy to read, low-noise, and show authentication status clearly.
- When opening the local app in a browser for validation, use `http://localhost:8999/...` instead of `127.0.0.1` so redirects and auth origin rules stay aligned.

## Project Structure

- `config.js`: hardcoded auth/runtime config (ES module, exports `config`)
- `core.js`: pure logic (calendar math, card shapes and updates, worked time, the weekly average model). No DOM, storage or network; unit-tested in `tests/core.test.js`.
- `ui.js`: shared DOM helpers (`$`, `escHtml`, `toast`, `showError`), and on localhost the side channel to `devserver.py`
- `mockgraph.js`: an in-memory Graph backend for local work, on with `?mock` (see Validation)
- `devserver.py` (run by `serve.sh`): static files served `no-store`, plus `POST /__agent/report` and `/__agent/log`, which write `.agent/` (ignored by git)
- `agent-test.sh`: the text-only test run (see Validation)
- `staticwebapp.config.json`: Azure Static Web Apps headers; `Cache-Control: no-cache` makes browsers revalidate every file, so there are no `?v=` cache-busting strings to bump
- `login.html`: sign-in page (loads `auth.js` directly)
- `index.html`: main app shell (loads `app.js`, which imports `auth.js` and `ui.js`)
- `auth.js`: MSAL bootstrap, session restore, sign-in/out, Graph fetch helpers. Auto-boots only on the login page; the app and test pages drive their own boot via explicit imports.
- `app.js`: team/timecard loading, rendering, interactions (production only — no probes). Exports the small surface area used by `test.js`.
- `common.css`: shared layout, buttons, modals, banners, sign-in, test-page scaffolding
- `timecards.css`: timeline, cards, team picker, app-specific UI
- `test.html`: browser self-test runner (loads `test.js`)
- `test.js`: self-test framework, contract probes, and assertions. Imports real production functions from `app.js`; performs Graph contract probes directly via `auth.graphFetch`.

## Working Approach

- Start from the narrowest deciding code path.
- Prefer small, local edits and validate immediately after each substantive change.
- Self-tests and Graph contract probes live in `test.js`. `app.js` contains no probe or test-only code; `test.js` imports the production functions it needs from `app.js` and `core.js`.
- Keep test pages browser-runnable without any build or package install.
- New logic that needs no DOM goes in `core.js` with a unit test.

## Validation Approach

Three lanes, fastest first. Each prints or writes one condensed line per suite plus the failures.

- `./agent-test.sh` (dev server running): unit tests with `node --test` (TZ is set to New York so daylight saving cases bite), then `test.html?mock` in headless Brave with a throwaway profile. About two seconds. Needs nobody signed in.
- Mock browser tests (`test.html?mock`): the flow tests load the real app in an iframe against `mockgraph.js` and drive it: clock in/out, breaks, edits, deletes, week navigation. They never touch real timecards. `?mock` is kept per tab (`?mock=0` turns it off) and mock mode uses its own `mock_` storage keys.
- Live tests (`test.html` in a signed-in browser): Graph contract probes against the real tenant. The page posts its condensed result to `.agent/browser-tests.txt`; read it with `./agent-test.sh live`.
- Every page on localhost forwards uncaught errors, rejections and `console.error` to `.agent/browser.log`.
- After switching dev servers once, hard-reload: files cached by an older server can outlive it.

## Verified Graph Notes

- Microsoft Graph timeCards filter compounds accept `and`, but this tenant rejects `or` in combined expressions such as overlap-or-open-state weekly queries.
- The accepted production week fetch is a single `clockInEvent/dateTime ge ... and clockInEvent/dateTime le ...` query; keep the rejected combined overlap/state probe as a negative assertion.
- Microsoft Graph `PUT /teams/{teamId}/schedule/timeCards/{timeCardId}` succeeds with `204 No Content`; client update flows must not require a returned card body.
- Microsoft Graph action endpoints can also succeed with `204 No Content`; use a follow-up `GET /timeCards/{id}` for clock-out/start-break/end-break and a current-week fetch to discover the new active card after clock-in.
- The edit modal uses minute-precision `datetime-local` inputs, so unchanged values must reuse the original ISO timestamp to avoid silently rounding away seconds.
- Cache `/me/joinedTeams` in `sessionStorage` for the app boot path so browser refreshes in the same tab do not block on the slow team lookup before the selected-week timecard query.

## TODO
- Drag tests: drive the timeline's resize handles in the mock flows (synthetic pointer events through interact.js), so drag jank can be fixed against a test.
- Recover from an expired refresh token. On an unmanaged PC the refresh token lease is short (8 of 24 hours); a page left open over sleep fails every interaction until the user reloads and signs in again.
- Full overhaul: less app code (JS and CSS), more tests, more DRY, more performant, less janky (drag glitches), and a little better looking.
- add mobile css


# Agent Notes

## User Preferences

- Keep this as a static app. Do not introduce a build step or install Node packages.
- Code is split into native ES modules (no bundler). Browsers load them directly via `<script type="module">`. They will not load over `file://` — always use `serve.sh` (http://localhost:8999).
- Node is acceptable for syntax checks only, such as `node --check`.
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
- `ui.js`: shared DOM helpers (`$`, `escHtml`, `toast`, `showError`)
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
- Self-tests and Graph contract probes live in `test.js`. `app.js` contains no probe or test-only code; `test.js` shares code by calling real production functions on `window.timecardsApp`.
- Keep test pages browser-runnable without any build or package install.
- Prefer visible test output over console-only debugging, since integrated browser auth state matters.

## Validation Approach

- Use `node --check` for JavaScript syntax validation.
- Use the browser self-test page for auth status, Graph reachability, and feature probes.
- Avoid adding dependencies just for validation.

## Verified Graph Notes

- Microsoft Graph timeCards filter compounds accept `and`, but this tenant rejects `or` in combined expressions such as overlap-or-open-state weekly queries.
- The accepted production week fetch is a single `clockInEvent/dateTime ge ... and clockInEvent/dateTime le ...` query; keep the rejected combined overlap/state probe as a negative assertion.
- Microsoft Graph `PUT /teams/{teamId}/schedule/timeCards/{timeCardId}` succeeds with `204 No Content`; client update flows must not require a returned card body.
- Microsoft Graph action endpoints can also succeed with `204 No Content`; use a follow-up `GET /timeCards/{id}` for clock-out/start-break/end-break and a current-week fetch to discover the new active card after clock-in.
- The edit modal uses minute-precision `datetime-local` inputs, so unchanged values must reuse the original ISO timestamp to avoid silently rounding away seconds.
- Cache `/me/joinedTeams` in `sessionStorage` for the app boot path so browser refreshes in the same tab do not block on the slow team lookup before the selected-week timecard query.

## TODO
Agent tooling first; it is the flywheel for everything after it.
- Agent test entrypoint: one text-only command or page mode that runs the tests and side-channels a condensed result (pass count, failures only) to a file the agent reads from the shell, so nobody has to look at the human-readable test page.
- Dev server that never serves stale files (no-cache), so a reload always runs the code on disk.
- Pure logic (week math, durations, card shapes, the weekly model) in a DOM-free module with `node --test` unit tests that run in the terminal in under a second.
- A mock Graph backend (in-browser, opt-in) so the agent can exercise clock in/out, breaks, drags, edits and deletes without touching real timecards.
- Then:
- Recover from an expired refresh token. On an unmanaged PC the refresh token lease is short (8 of 24 hours); a page left open over sleep fails every interaction until the user reloads and signs in again.
- Full overhaul: less app code (JS and CSS), more tests, more DRY, more performant, less janky (drag glitches), and a little better looking.
- add mobile css


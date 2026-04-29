# Agent Notes

## User Preferences

- Keep this as a static app. Do not introduce a build step or install Node packages.
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

- `config.js`: hardcoded auth/runtime config
- `login.html`: sign-in page
- `index.html`: main app shell
- `auth.js`: MSAL bootstrap, session restore, sign-in, sign-out, Graph fetch helpers
- `app.js`: team/timecard loading, rendering, interactions, probe logic
- `common.css`: shared layout, buttons, modals, banners, sign-in, test-page scaffolding
- `timecards.css`: timeline, cards, team picker, app-specific UI
- `test.html`: browser self-test runner
- `test.js`: self-test framework and probe runner

## Working Approach

- Start from the narrowest deciding code path.
- Prefer small, local edits and validate immediately after each substantive change.
- Reuse the existing Graph probe logic instead of duplicating behavior in separate tools.
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

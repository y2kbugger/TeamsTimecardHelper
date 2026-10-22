# Teams Timecards

A small web app for Microsoft Teams Shifts timecards: clock in and out, take breaks, and see and
edit a week of cards on day timelines whose edges drag. A weekly average report compares worked time
with a weekly target. It is static files (HTML, CSS and native ES modules, no build) that talk to
Microsoft Graph with the signed-in user's token, through MSAL.


## Use it

Open https://teamstimecardhelper.y2kbugger.com/ and sign in with a work or school account whose team
uses Shifts.


## The road ahead

- Desktop only: phones are not a target.


## Working on it

- `./serve.sh` serves the app at http://localhost:8999/.
    - Use `localhost`, not `127.0.0.1`: the sign-in redirect URIs are registered for it.
    - Files are served `no-store`, so a reload always runs what is on disk.
- `./agent-test.sh` runs every test that needs nobody signed in, in about two seconds:
    - unit tests of `core.js` with `node --test` (in New York time, so daylight saving cases bite);
    - `test.html?mock` in headless Brave: the real app in an iframe, driven against a mock Graph.
- `test.html` in a signed-in browser probes the real Graph; `./agent-test.sh live` prints its last
  result.
- Pushing to `master` deploys (Azure Static Web Apps uploads the repository as it is).


### Rules

- Static: no build step and no packages. Node is only for `node --check` and `node --test`.
- Pure logic (dates, durations, card shapes, the weekly model) goes in `core.js`, with a unit test.
- Every loop has a named bound; invariants our own code derives are asserted (`assert` in
  `core.js`); input from Graph or the user is validated, never asserted.
- External scripts are pinned to an exact version with an integrity hash. To update one, change the
  version in every page that loads it and replace the hash with the new file's:
  `curl -s URL | openssl dgst -sha384 -binary | openssl base64 -A`.
- Browsers revalidate every file (`staticwebapp.config.json`, and `no-store` locally), so files carry
  no `?v=` cache busting.
- Days and weeks are calendar arithmetic (`addDays`), never `n * 24h`: daylight saving days last 23
  or 25 hours.
- The app registration is fixed in `config.js`; nothing about sign-in is configurable at runtime.
- Simple DOM over decoration; resize handles stay thin and square.


### Behaviour worth knowing

- The team is chosen once, in a dialog that opens by itself when none is saved in this browser
  (one team is chosen without asking); the team's name in the toolbar opens it again.
- Weeks run Sunday to Saturday. The chosen week holds while the user is active; after an hour
  without a click or key (a new sign-in, a tab left open over the weekend) the app returns to the
  current week, so Monday morning opens on the new week.
- The toolbar has Clock In while no card is open. The open card carries Clock Out and its break
  buttons, as a card's own actions; on another week it is not on the page, and Clock In stays
  hidden there too (whether a card is open comes from the current week, whatever week is shown).
- A card past midnight is worked on in the day it clocked in: there an edge past midnight has its
  handle at the end of the axis; on the next day it is a faint ghost that takes no clicks.
- The timeline's language, for cards and breaks alike: a white handle is a set time that drags; a
  sheen flowing toward now is still running. A card over 16 hours is red (most likely a clock-out
  forgotten overnight). Past days and card controls stay dim until pointed at.
- Breaks and missed cards are added on the timeline, every day of the week up to today having
  one: hovering a card shows a ghost of a ten-minute break fitted into its free time under the
  pointer, hovering free time a ghost of a thirty-minute card (ending by now), and Ctrl+click adds
  it (a plain click only selects). Its edges then drag like any other. The edit dialog changes and
  removes breaks.
- Weekly average: expected time is split evenly over the chosen workdays and counts only through
  today.


### Layout

| file | what |
|---|---|
| `index.html`, `app.js` | the app: teams, the week, timelines, the toolbar, the edit dialog |
| `timeline.js` | a day's timeline and dragging the edges of cards and breaks |
| `weeklyavg.js` | the weekly average report |
| `core.js` | pure logic, no DOM, storage or network |
| `auth.js`, `config.js` | MSAL sign-in, the expired-session bar, Graph requests |
| `login.html`, `auth-done.html` | the sign-in page; the empty page popups and silent renewal return to |
| `ui.js` | DOM helpers, and on localhost the error and report channel to the dev server |
| `common.css`, `timecards.css` | shared styles; the app's |
| `test.html`, `test.js` | browser tests: Graph probes (live) and app flows (with `?mock`) |
| `mockgraph.js` | an in-memory Graph for local work, on with `?mock` and off with `?mock=0` |
| `devserver.py`, `serve.sh`, `agent-test.sh` | the dev server and the text-only test run |


### Sign-in

- The app's redirect URIs (Single-page application) are `login.html` and `auth-done.html`, on
  localhost:8999 and on the production host. Any other page gets AADSTS50011.
- Popups and MSAL's hidden renewal iframe come back to `auth-done.html`, which loads nothing:
  its opener or parent reads the response from its URL.
- When a token cannot be renewed silently (an unmanaged PC's refresh token lasts 8 of 24 hours),
  Graph requests wait and a bar offers Sign in. Its click opens the popup (browsers block popups no
  click opened); the waiting requests then finish.
- Sign-out is local: it clears this app's tokens and leaves the Microsoft session alone.
- On localhost, `sessionStorage.dev_no_iframe_renewal = 1` stops silent renewal at the refresh
  token, as on such a PC; with `?mock`, `sessionStorage.mock_tc_session_expired = 1` expires the
  mock session.


### Microsoft Graph, as this tenant behaves

- timeCard filters accept `and` but not `or`, so a week is one query: clock-in within the week.
  A card that clocked in the week before and runs into this one is not fetched.
- `PUT /teams/{id}/schedule/timeCards/{id}` and the clockIn, clockOut, startBreak and endBreak
  actions answer `204 No Content`: the app reads the card back (or, after clock-in, the week).
- Deleting a timecard is beta-only (`DELETE` on `/beta`).
- `/me/joinedTeams` is slow, so it is cached for the browser session.
- The edit dialog's inputs have minute precision: an unchanged value keeps its original timestamp.


### The dev server's side channel

On localhost every page posts uncaught errors, rejections and `console.error` to
`.agent/browser.log`, and the test page posts a condensed result to `.agent/browser-tests.txt` or
`.agent/mock-tests.txt` (`.agent/` is ignored by git).

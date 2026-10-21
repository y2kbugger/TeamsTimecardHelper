#!/usr/bin/env bash
# Text-only test run: unit tests in node, then the browser tests against the mock backend in a
# headless browser with a throwaway profile. Prints one line per suite and only the failures.
# Needs the dev server (./serve.sh) on localhost:8999. Exit status is nonzero on any failure.
#
#   ./agent-test.sh          unit + mock browser tests
#   ./agent-test.sh unit     unit tests only
#   ./agent-test.sh live     read the last live report (test.html run in the signed-in browser)
set -uo pipefail
cd "$(dirname "$0")"

BROWSER="${BROWSER:-brave}"
BASE_URL="http://localhost:8999"
status=0

run_unit() {
    local out
    out=$(TZ=America/New_York node --test --test-reporter=tap tests/ 2>&1)
    local pass fail
    pass=$(grep -c '^ok ' <<<"$out")
    fail=$(grep -c '^not ok ' <<<"$out")
    if [ "$fail" -eq 0 ] && [ "$pass" -gt 0 ]; then
        echo "ok unit tests: $pass passed"
    else
        echo "FAIL unit tests: $pass passed, $fail failed"
        grep -A12 '^not ok ' <<<"$out" | grep -E "^not ok|error:|expected:|actual:|at .*core" | head -40
        status=1
    fi
}

run_mock_browser() {
    if ! curl -sf -o /dev/null "$BASE_URL/test.html"; then
        echo "FAIL mock browser tests: no dev server at $BASE_URL (run ./serve.sh)"
        status=1
        return
    fi
    local report=.agent/mock-tests.txt profile
    profile=$(mktemp -d)
    rm -f "$report" .agent/browser.log
    # Its own process group, so the browser's helper processes go down with it.
    setsid "$BROWSER" --headless=new --disable-gpu --no-first-run --no-default-browser-check \
        --user-data-dir="$profile" "$BASE_URL/test.html?mock" >/dev/null 2>&1 &
    local pid=$!
    for _ in $(seq 120); do
        [ -s "$report" ] && break
        sleep 0.25
    done
    kill -- -"$pid" 2>/dev/null
    wait "$pid" 2>/dev/null
    rm -rf "$profile"
    if [ ! -s "$report" ]; then
        echo "FAIL mock browser tests: no report within 30s"
        status=1
    else
        cat "$report"
        echo
        head -1 "$report" | grep -q '^ok ' || status=1
    fi
    # Entries start with a timestamp; stack lines follow. Self-test failures are reported above.
    local errors
    errors=$(awk '/^[0-9]{4}-/ { keep = !/\[self-test\] FAIL/ } keep' .agent/browser.log 2>/dev/null | cut -c1-300 | head -20)
    if [ -n "$errors" ]; then
        echo "browser errors (.agent/browser.log):"
        echo "$errors"
        status=1
    fi
}

case "${1:-all}" in
    unit) run_unit ;;
    live) cat .agent/browser-tests.txt; echo ;;
    all) run_unit; run_mock_browser ;;
    *) echo "usage: $0 [all|unit|live]"; exit 2 ;;
esac
exit $status

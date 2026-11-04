# _test — checks for the reporting / resume layer (never shipped)

Requirements: Node ≥ 22 (built-in WebSocket and fetch) and Microsoft Edge at
`C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe`. There are no npm dependencies. Each
browser script serves the repo itself over HTTP under a `/sub/` subpath, which stands in for a CDN
folder, then drives headless Edge through the DevTools protocol on a random port.

| script | what it checks | run |
|---|---|---|
| `verify-static.js` | Asset references exist with the exact case and sit in shipped folders; one `?v=`; ids match the metadata; the screen→item map is complete; exactly one `reportQ` per declared question; choice option counts match `answers[]`; the loader letter matches its regex; no name is declared in both the shared layer and a part. | `node _test/verify-static.js` |
| `shoot.mjs` + `cmp.mjs` | Screenshots of every screen of every part (animations and media frozen), plus 404s, console errors, loaded font faces and broken images. `cmp` compares two runs byte-for-byte. | `node _test/shoot.mjs <repo> <out> [01,02,…]` · `node _test/cmp.mjs <baseOut> <newOut> [parts]` |
| `probe.mjs` | A scripted behaviour run from a JSON step list: `nav`, `eval` with `expect`, `wait`, `hits`, `shot`. | `node _test/probe.mjs <repo> _test/scenarios/<file>.json` |
| `resume-test.mjs` | **The resume oracle.** For each scenario it runs the setup, re-enters the screen in-session (screenshot A), reloads so the restore lands (screenshot B), and passes only if A ≈ B (anti-alias tolerant), the landing screen is right, no `answered` / `completed` / `requested` is re-sent, and there are no console errors. | `node _test/resume-test.mjs <repo> _test/scenarios/resume-NN.json <out>` |
| `xapi-720-k.js` | Local stand-in for the CDN library. It keeps statements and state in sessionStorage. Helpers: `__stmts()`, `__state()`, `__dupes()`, `__reset()`, `__failWrites(status)`. | page URL `?xapiLib=../_test/xapi-720-k.js`, localhost only |

## Scenarios

| file | covers |
|---|---|
| `xapi-02.json` | Part 02 statement flow: boot `initialized`, item open/close, Back closes nothing (O-1), `answered` vs `answered.last`, hint once, explicit item score, answer texts, no duplicates. |
| `xapi-mix.json` | Part 01 guided picks, hotspot text, part 03 drag-board text, part 04 one-shot riddle, part 05 hint id, part 06 two-field text. |
| `qa-2026-10-04.json` | Fixes from the first live run: B-2 (Back onto item-less S0 closes nothing; the real close is sent once), F-6 (all-sections items in 01–04), B-5 (05/06 content videos: `played, paused, played`, leaving mid-play, no orphan `paused`, only content video opted in). Videos run at `playbackRate 16` — this server has no Range support, so seeking is ignored. |
| `routing.json` | Production / `?dev=1` / `?dev=1&registration` boots. The last click reports once and disables its button with no hop; the dev hop keeps the query; S0 Back is hidden; part 03's checkpoint-1 fail ends at S5. |
| `report.json` | The issue dialog with `fetch` stubbed: gating, confirm-on-close, the payload fields, the thanks message plus its announcement, Escape. |
| `challenge-60.json` | Part 04 (the challenge, MOE 2026-10-08): `success` only at ≥60% of the 7 declared questions, with the real score; unanswered = wrong. |
| `partnav.json` | The `?dev=1` walkthrough across parts: forward, Back shown, back lands on the same screen, the replay override. |
| `resume-NN.json` | The resume oracle per part. |

## Notes

- **Resetting between runs.** A lomda page that has restored a document writes it back when it is left (`pagehide` flush). `resume-test.mjs` therefore clears storage from a same-origin page that is not the lomda (`metadata/*.json`) before each scenario.
- **Mutation-test every new assertion** (guide §10.4). The resume oracle was checked this way: dropping one `Mark()` from a painter fails with 731 pixels off.

# methodica-math-ratio-03 — reporting & resume: the unit record

> **This repo is the Arabic twin, `methodica-ar-math-ratio-03` (2026-11-15).** The code is the Hebrew
> unit's at `8ddfc35` (14.11.26: 03 scoring 2187cb6, response format ff88372, resume hardening + leave
> save d9296c2, all "same change as the Hebrew twin"), with Arabic text. What differs from Hebrew:
> - Identity: `XAPI_ID_PREFIX = …/720/ar/math/ratio/03/`, `XAPI_UNIT_ID = 'methodica-ar-math-ratio-03'`,
>   folders `methodica-ar-math-ratio-03-0N`, browser keys `methodica_ar_math_ratio_03_*`.
> - Metadata in Arabic; every string the page sends (choice options, matching sources / targets, step
>   names) is the Arabic screen's own, and `_test/statement-flow.mjs` checks each right answer against
>   its pattern.
> - Shared files differ from Hebrew only in: `15-engine.js` (Arabic retry titles), `90-boot.js` (hint
>   pill follows the primary button), `10-base.css` (Arabic font stack), `25-report.js/.css` (Arabic
>   dialog; the Google Form values stay Hebrew on purpose).
> - Part 03 S5: as Hebrew again (owner, 15.11.26) — below 2 of 3 the learner stays on S5 and the component
>   ends there. The Arabic QA change of 10.11.26 (cceb5dd: Continue always to S6) is reverted.
> - Kata: components go into the Hebrew unit (parent-unit mode), see SEND-METADATA.md.
>
> Everything below is the Hebrew unit's record.

Branch `vadimr/feat/reporting-resume`. This file records what was built, what was decided, and
what is still open. The general method is in
`Documentation/reporting-and-resume/ADDING-REPORTING-AND-RESUME.md` (below: "the guide").
The guide is the source for the rules. This file covers only what is specific to this unit.

## 1. Layout

```
unit-assets/   fonts/ (5 Assistant TTFs), img/gesture-hand.svg — shared by more than one part
unit-css/      00-fonts · 10-base (the shared prologue) · 20-motion · 25-report
unit-js/       05-motion · 08-partnav · 10-identity · 15-engine · 20-xapi · 22-scoring ·
               25-report · 40-resume · 45-painters · 50-loader · 90-boot
metadata/      unit + 6 component JSON files (MOE v2.5; the source of truth for ids)
methodica-math-ratio-03-0N/   index.html · js/main.js · css/style.css · assets/   (+ docs/, tests/: never shipped)
_test/         verify-static.js, the local xapi-720-k.js stub, browser harnesses
docs-and-tools/ packaging (allowlist/build/verify), Kata send/retrieve, this record
```

**Load order in every `index.html`:**

- **Head:** `00-fonts`, `10-base`, `css/style.css`, `20-motion`, `25-report`.
- **End of body, in this order:**
  1. `05-motion`, `08-partnav`
  2. `10-identity`, `15-engine`, `20-xapi`, `22-scoring`, `25-report`, `40-resume`, `45-painters`, `50-loader`
  3. `js/main.js`
  4. `90-boot`

**Rules:**

- All code references carry one `?v=`, which is bumped together.
- No top-level name may be declared both in `unit-js/` and in a part's `main.js` (§3.3). `_test/verify-static.js` checks this.

## 2. Identity

| | |
|---|---|
| `XAPI_ID_PREFIX` | `https://lomdot.education.gov.il/metodica/720active/math/ratio/03/` |
| `XAPI_UNIT_ID` | `methodica-math-ratio-03` (bare slug, = `metadata/methodica-math-ratio-03_unit.json` id) |
| component id | `XAPI_ID_PREFIX + 'methodica-math-ratio-03-0N/'` (checked against `METADATA.id` on every load) |
| library | CDN `xapi-720-k.js` (resume on) — `XAPI_USING_G` regex `[ghijk]` |
| state document | `execution-state::methodica-math-ratio-03-0N`, one per {learner, component}, v6 |
| character key | `methodica_math_ratio_03_selectedCharacter` (= `CHARACTER_STORAGE_KEY`, 15-engine.js) |
| verdict keys | `methodica_math_ratio_03_0N_result` → `'pass' / 'fail'` |

## 3. Screen → item maps (`SCREEN_TO_SUBCONTENT`)

| part | screens | items |
|---|---|---|
| 01 (24) | s0–3 → 001 · s4–8 → 002 · s9–19 → 003 · s20–23 → 004 | 001 ungraded (widget); 002/q1 s8; 003 is the guided item (picks on s11/s15/s17): no question, no answer, no score since 11.11.26 (Kata declares none; unit owner 10.11.26); 004/q1–q3 s21–s23 |
| 02 (9) | s0 null · s1–3 → 001 · s4–5 → 002 · s6–8 → 003 | one screen per question |
| 03 (11) | s0 null · s1–2 → 001 · s3 → 002 · s4–5 → 003 · s6–7 → 004 · s8–10 → 005 | s6 (transition) opens 004, so a Back from s7 does not close it (O-1) |
| 04 (12) | s0 null · s1 → 001 · s2 → 002 · s3–5 → 003 · s6 → 004 · s7–11 → 005 | 005/q1 is the one-shot riddle on s8 |
| 05 (9) | s0–8 → 001 | q1 s2 · q2 s4 · q3 s6 · q4 s7 |
| 06 (10) | s0–9 → 001 | q1 s3 · q2 s4 · q3 s5 · q4 s7 |

Every declared question has exactly one `reportQ` site, and the static check enforces it (D-9).

**Answer texts (`result.response`, since 11.11.26):** MOE compares a right answer's response with Kata's
`correctXapiResponsesPattern` character by character (runbook §3). Every site passes a spec, and `xapiResponse`
(`unit-js/20-xapi.js`) builds the string from the metadata by `questionType`. `student_answer` keeps a readable copy.

- **Choice questions** report the metadata's own `answers[]` string for each picked option, matched by position (`xapiChoiceAnswer`). The visible text would garble superscripts and stacked fractions. A multi-select is joined with `[,]` in `answers` order, whatever the click order.
- **Value inputs** report what was typed (`fill-in` / `numeric`). The value is read before any final-wrong reveal rewrites the field.
- **Dropdowns and drag boards** are matching questions: `source[.]target[,]…`, right pairs in `correctAnswers` order (02-003/q2, 03-003/q2, 03-005/q1–q2, 04-002).
- **Two-field inputs** (02-002/q1, 03-004/q1, 03-005/q3, 05-001/q4, 06-001/q3) are `other` in Kata, with no pattern. The repo metadata proposes them as `performance` (`Tools/moe-qa/propose-performance.py`), e.g. `במישור[.]15[,]בעלייה[.]6`, and the page sends that form. They reach Kata only through the owner's send, item by item.
- The metadata carries Kata's `correctXapiResponsesPattern` / `xapiInteractionComponents` (`Tools/moe-qa/add-kata-patterns.py`), and `_test/statement-flow.mjs` fails on a right answer that is not one of them.

## 4. Decisions (2026-10-01)

1. **Metadata** was supplied by the owner and committed as is. No id was invented.
2. **Optimisation depth: a shared engine.** Identical files and functions were hoisted, and the per-screen template code stays per part.
   - No media was re-encoded and no file was renamed. See the open items.
3. **Fail path: the guide's rule** (§6.4, §8.8 rule 12).
   - The last click reports the component `completed`, pass **and** fail, disables its button and stops. Kata routes on the statement (`recommendedAfterFail`).
   - This supersedes the producer's earlier rule that kept "המשך" enabled on a fail (02, 03, 05, 06).
   - Part 03 also ends at a failed checkpoint 1 (S5): `PART_END_SCREENS = [5, 10]`.
4. **Re-entry keeps a half-answered question.** `resetScreenState` resets a screen only while it is pristine (`screenPristine`). Attempts, a pending answer and the retry gate survive a round trip; before this change they were reset on every entry.
5. **Resume uses per-screen painters**, not a DOM diff (§8.8).
   - Each part's `PART_SCREENS` names its state and redraws it with the template's own functions.
   - `45-painters.js` rebuilds a restored screen once, on first entry.
6. **The platform owns routing.**
   - PartNav (`?dev=1` walkthrough only) is a no-op in production, and S0's Back is hidden.
   - Hops carry the query string and save first.
7. **The sessionStorage part snapshots of 01 and 03** were removed and folded into the state document.

## 5. Component results (`partGateResult`)

| part | promised on screen | reported |
|---|---|---|
| 01 | none (warm-up outcome does not gate) | success:true, score = catalogued share |
| 02 | 2 of 3 questions | success = masteryState().passed, score = questions/3 |
| 03 | checkpoint 1 (2 of 3) and checkpoint 2 (both) | success = both passed, score = catalogued share |
| 04 | none | success = catalogued share ≥ 0.6, score = catalogued share (of 7; unanswered = wrong). **Changed 2026-10-08 (MOE, with Maya):** the challenge ("בואו נאתגר") reports its real score, success only at ≥60%. Still no gate on screen, nothing shown |
| 05 | 3 of 4 sections | success = peak.scoredCorrect, score = sections/4 |
| 06 | 3 of 4 sections | success = s9MasteryMet(), score = sections/4 |

## 6. How to test

- **Static:** `node _test/verify-static.js`.
- **Browser** (headless Edge over CDP, Node ≥ 22, no dependencies; see `_test/README.md`):
  - every screen of every part pixel-compared against a baseline;
  - reporting flows against the stub (`?xapiLib=../_test/xapi-720-k.js`, localhost only);
  - the resume oracle: a screen restored after a reload must equal the same screen re-entered in-session, and nothing may be re-sent.
- **Package:** `pwsh -File docs-and-tools/build-package.ps1 -DryRun`, then build, then `verify-package.ps1`.
- **QA regressions:** `node _test/probe.mjs . _test/scenarios/qa-2026-10-04.json` (B-2, B-5, F-6; §6b).
- **Live:** the guide §12 probes 1–7 against Kata, after a deploy only.

## 6a. Verification at the resume-on commit (2026-10-01)

- `node _test/verify-static.js` — 2187 checks, ALL PASS.
- Screenshot regression, every screen of all six parts with the real CDN library and no
  `?slxapi` (the guide's regression gate) — 75/75 byte-identical to the pre-change baseline, 0 x 404,
  0 console errors.
- Resume oracle with the local stub — 01: 32, 02: 17, 03: 26, 04: 22, 05: 22, 06: 23 scenarios, all PASS
  (no answered / completed / requested re-sent by any restore). Mutation-tested on 02 and 03.
- Probes — reporting flows (xapi-02, xapi-mix), routing (three boots), issue dialog, PartNav dev walk:
  all PASS, 0 console errors.
- Package `deployments/2026-10-01` (outside git), built from `d1aa225`: 130 files, 63.97 MB; build self-check
  PASS; `verify-package.ps1` FORWARD/REVERSE/HYGIENE/COMMIT PASS; served from a subpath it shows 75/75 screens
  byte-identical, 0 × 404; package boot probe (`_test/scenarios/package-boot.json`) PASS for all six parts against
  the real CDN `xapi-720-k`; leak scan clean. Its DEPLOY.md carries the upload target, order and the media report.
- Metadata: `send-metadata.ps1 -DryRun` reviewed — `created=26 updated=1 failed=0` (1 unit, 6 components, 19 items;
  the update is 03 → 02 `recommendedAfterFail`). Live send pending the API key; it goes in as draft, no publish.
- Kata preflight 2026-10-01: unit + 6 components ABSENT (exit 0).
- **Live send 2026-10-01 REJECTED, nothing written** (`created=0 updated=0 failed=1`): the unit POST got HTTP 409
  `resource.conflict` — "objective already bound to a unit (strict 1:1)". `MOE.MATH.G8.NUM.RATIO-PROP-SCL.RATIO.ALG-EXPR`
  is already bound to **methodica-math-ratio-02** in Kata (its `metadata-from/` read-back). The objectives catalogue
  (720-metadata-extractor `learning-objectives.json`) assigns ALG-EXPR to **ratio-03** and gives ratio-02 no MOE code
  ("לבטא יחס בדרכים שונות"). Needs an owner decision before any resend — see §7.
- ratio-02's objective cannot be changed through the API: `PATCH /content-units/methodica-math-ratio-02` with
  `learningObjective` returned 200 and ignored it (`updatedAt` unchanged), published or draft. A Kata admin can
  re-bind it (available next week).
- **Sent 2026-10-01 with a TEMPORARY objective — `created=26 updated=1 failed=0`.** By the owner's decision the
  unit went in with `MOE.MATH.G8.NUM.RATIO-PROP-SCL.RATIO.MULTI-FORM` (the code that belongs to ratio-02), sent
  from a scratch copy of `metadata/` that differs in that one field; the repo keeps the final, correct
  `ALG-EXPR`. Read-back (`retrieve-metadata.ps1 -FailOnDrift`): unit=1 components=6 items=19 warnings=0;
  `-Status`: 6/6 draft; nothing published. Logs: `send-metadata-2026-10-01-live-multiform.log`,
  `retrieve-metadata-2026-10-01.log` (git-ignored).
- **CDN upload verified 2026-10-01:** the owner uploaded `deployments/2026-10-01` to `.../720/math/ratio/03/`; all 130
  files byte-identical on the CDN (size + MD5); root index.html / DEPLOY.md / _test / index_dev.html not served; all
  six parts boot live from the CDN (lib k, ids match, cover drops, dialog hidden, DEV_NAV off, 0 console errors).
- Not yet done: the objective swap (Kata admin), publishing (owner) and the live Kata run (§12).
  *(All three done by 2026-10-05 — see §7.)*

## 6b. Fixes from the first live run (QA 2026-10-04 round, branch `vadimr/fix/qa-2026-10-04`)

- **B-2 — Back onto an item-less screen closes nothing** (`unit-js/20-xapi.js`). Direction is the screen's
  (`_xapiLastScreen`), port of ratio-04 `e7a9531`. Live it was 02 s1 → s0 and 04 s1 → s0 sending
  `completed 001 {success:false}`; 03 has the same null s0.
- **B-5 — the content videos are reported** (decision 2026-10-05). `data-xapi-report="001"` on 05 `#s3-video`,
  06 `#s6-video`, `#s8-video`; `xapiWireVideos` now uses the scale-01 played/paused state machine (the old
  `pausedOnce` latch swallowed the first `played`, guide §6.9.2). Character loops stay unwired.
- **F-6 — item success follows the part's all-sections rule** (`unit-js/22-scoring.js` `itemResultFor`). Items
  listed in a part's `XAPI_ITEM_ALL_CORRECT` succeed only when every question is right (score = share):
  01 `004`; 02 `001 002 003`; 03 `001 003 005`; 04 `003`. 05/06 keep the share rule (one item = the whole
  part, "k of 4"). `verify-static` checks each listed item is graded and multi-question.
- `?v=2` → `?v=3` on all 102 references.
- Tests: `_test/scenarios/qa-2026-10-04.json` (B-2, F-6 in 01–04, B-5 in 05/06 incl. no orphan `paused` and
  leaving mid-play) — 11 of its assertions FAIL on the pre-fix code (mutation check). Videos are run at
  `playbackRate 16`: the test server has no Range support, so seeking is ignored there.

## 6c. The challenge reports its real score (MOE 2026-10-08, branch `vadimr/feat/challenge-success-60`)

- **04 — "בואו נאתגר את עצמנו".** MOE decided, with Maya, that the challenge reports its real score and
  `success` only at ≥60%. `partGateResult` now sends `success = partScaled() >= 0.6` (it was `true`), with the
  same real `score.scaled`: the share of the 7 questions metadata declares (001 q1, 002 q1, 003 q1–q3,
  004 q1, 005 q1), where an unanswered question counts as wrong. So 5 of 7 or more passes. Nothing changes on
  screen: no gate, the learner always continues, nothing is shown.
- `?v=3` → `?v=4` on all 102 references.
- Test: `_test/scenarios/challenge-60.json` (0, 2, 4 of 7 fail; 5, 7 of 7 pass, each with the real score).

## 6d. 03 reports what S0 promises (MOE sheet 08.11.26, fixes 1, 2, 4, 6, branch `vadimr/fix/score-success-2026-11`)

- **03 — `partGateResult`.** S0 promises "3 שאלות … ענו נכון על 2 שאלות ומעלה כדי להתקדם". `success` and
  `score.scaled` now come from one count: checkpoint 1's three questions (`mastery1State().correctCount`, each
  correct only when all its sub-parts are, the count the S5 gate reads). `success = correct >= 2`,
  `scaled = correct / 3`. Until now `success` also needed both harder questions (S7–S10) right, although S6
  promises no threshold (fix 4), and the score was the share of all 9 catalogued sub-questions, the harder
  ones included (fixes 1, 6). So 3/3 + harder wrong gave `success:false, 0.44` (fix 2). The harder questions keep
  their own `answered`. The S5 gate and the S10 end are unchanged. 04 keeps the MOE 08.10 60% rule (6c).
- `?v=4` → `?v=5` on all 102 references.
- Test: `_test/scenarios/score-contract.json` (every count 0–3 with the harder questions all wrong / all right,
  success ⇔ scaled ≥ 2/3 = the gate's verdict, a question with one sub-part wrong counts 0, the gate at 1/3 and
  2/3). 4 of its 6 expectations FAIL on the pre-fix code (mutation check).

## 7. Open items

- **Media, not optimised (by decision):**
  - `01/assets/widgets/Hummus Stand.html`: 16.9 MB, of which ~16 MB is four base64 1920×1080 PNGs.
  - `03/assets/video/Character_{1,2}_Proud.mp4`: ~6 MB each, 1280×720 @ 12 Mbps.
  - `06/assets/img/התראה.png` (1.5 MB) and `הצלחה.png` (0.6 MB).
  - Large 4:4:4 JPGs.
- **Hebrew and space-containing asset file names:** referenced raw, and they work. Renaming was not done (by decision).
- **The dropdown icons** `02/assets/img/dq-*.svg` and `03/assets/icons/dropdown-*.svg` are byte-identical. Merging them needs a rename, so it was not done.
- **`$ContentBaseUrl`** (`.../720/math/ratio/03`, Kata `hostedContentRef`): confirmed by the owner 2026-10-01.
- **Known issue (guide):** the leave-page state write can be blocked by Chrome (KNOWN-ISSUE-dismissal-state-write.md). At most ~800 ms of position plus unsubmitted input is lost.
- **Kata objectives — ratio-03 re-bound.** Read back 2026-10-05: ratio-03 holds `…RATIO.ALG-EXPR`, equal to
  `metadata/` (QA/2026-10-04/LINK-AUDIT.md §4); all six components are published. ratio-02's side was not checked
  here, and ratio-02's own repo metadata (`metadata/`, `metadata-src/`) still says ALG-EXPR and needs correcting.
- **Live run done 2026-10-05** (QA/2026-10-04/REPORT.md): 6/6 PASS; one QA-marked issue report sent to the
  shared form. Fixed from it in §6b: B-2, B-5, F-6. Still open from it: C-1 (Kata has 0 questions for 01-003 —
  re-send the 01 metadata), C-2 (03 `isRequired` false in Kata) and the Low items (O-5, F-3, O-9, O-10).

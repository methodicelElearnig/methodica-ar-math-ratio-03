'use strict';
/* ═══════════════════ xAPI — item and component results ═══════════════════
   Ported from methodica-math-ratio-05. Results are derived from XAPI_Q_RESULTS (20-xapi.js:
   '<item>/<qKey>' → correct, written by every answered call) and from the metadata, which is the
   source of truth for how many questions an item declares. So a result can never disagree with the
   answers already reported, and an unanswered question counts as not correct.

   Per-part seams (js/main.js):
     XAPI_EVAL_ITEMS       items that carry a code-graded question, e.g. { '002': 1, '003': 1 }
     XAPI_PASS_SCALED      optional pass threshold for items (default 0.6, MOE's own example)
     partGateResult()      optional — the component result, when the part promises its learner a
                           gate of its own on screen (e.g. "2 of 3 questions"). The gate the
                           learner sees is the gate that is reported (guide §6.4); without the
                           hook partResult() falls back to the metadata-weighted share below. */
var XAPI_DEFAULT_PASS = 0.6;

function xapiPassScaled() {
  return (typeof XAPI_PASS_SCALED === 'number') ? XAPI_PASS_SCALED : XAPI_DEFAULT_PASS;
}

/* Number of questions the metadata declares for an item; falls back to the answers seen so far
   when the metadata is not loaded (reporting off) so the call never throws. */
function itemQuestionCount(item) {
  try {
    var sc = (window.METADATA && window.METADATA.subContent) || [];
    for (var i = 0; i < sc.length; i++) {
      if (String(sc[i].id).replace(/\/+$/, '').slice(-(item.length + 1)) === '-' + item) {
        return (sc[i].questions || []).length;
      }
    }
  } catch (e) {}
  return Object.keys(XAPI_Q_RESULTS).filter(function (k) { return k.split('/')[0] === item; }).length;
}

function itemCorrectCount(item) {
  return Object.keys(XAPI_Q_RESULTS).filter(function (k) {
    return k.split('/')[0] === item && XAPI_Q_RESULTS[k] === true;
  }).length;
}

/* Item 'completed' result: MOE v2.4 requires success + score on any item carrying answered
   interactions; the library's own aggregate is an all-correct AND.

   ── F-6 (QA 2026-10-04): an item whose part scores it as ONE question ──
   Where an authored question is split into sections (01 item 004, 02 items 001–003, 03 items
   001/003/005, 04 item 003), the part — and its progress bar — counts the question correct only
   if EVERY section is right. The share rule above reported such an item as a pass at 2 of 3
   (live: 02 item 003 {success:true, scaled:0.667} while the learner's screen showed it wrong).
   A part lists those items in XAPI_ITEM_ALL_CORRECT (declared in the part, never here); for them
   success is the AND of the final answers. The score stays the share. XAPI_Q_RESULTS holds each
   question's final outcome and is restored by resume, so the AND survives a reload. */
function itemResultFor(item) {
  var n = itemQuestionCount(item);
  if (!n) return null;
  var ok = itemCorrectCount(item);
  var scaled = ok / n;
  var allRule = (typeof XAPI_ITEM_ALL_CORRECT !== 'undefined') && !!XAPI_ITEM_ALL_CORRECT[item];
  return { success: allRule ? ok === n : scaled >= xapiPassScaled(), score: { scaled: scaled } };
}

/* The share of the part's catalogued questions answered correctly — the score every component
   result carries. */
function partScaled() {
  var items = (typeof XAPI_EVAL_ITEMS !== 'undefined') ? Object.keys(XAPI_EVAL_ITEMS) : [];
  var total = 0, ok = 0;
  items.forEach(function (it) { total += itemQuestionCount(it); ok += itemCorrectCount(it); });
  return total ? ok / total : null;
}

/* The component result — sent on EVERY exit, failing ones included (guide §6.5). */
function partResult() {
  if (typeof partGateResult === 'function') {
    try { var r = partGateResult(); if (r) return r; } catch (e) { console.error('[xAPI] partGateResult', e); }
  }
  var scaled = partScaled();
  if (scaled == null) return { success: true };
  return { success: scaled >= xapiPassScaled(), score: { scaled: scaled } };
}

/* Read by 20-xapi.js xapiItemResult(). Built at boot (90-boot.js), after the part has declared
   XAPI_EVAL_ITEMS. */
var XAPI_ITEM_RESULT = {};
function buildItemResults() {
  if (typeof XAPI_EVAL_ITEMS === 'undefined') return;
  Object.keys(XAPI_EVAL_ITEMS).forEach(function (it) {
    XAPI_ITEM_RESULT[it] = function () { return itemResultFor(it); };
  });
}

/* One graded answer, from a question's Check. ref = [itemSuffix, qKey] or null (not reported).
   The resume flush is deferred to the end of the current task (flushAfterCommit, 45-painters.js),
   so it lands after the handler's LAST state write even when the handler returns early (D-13). */
function reportQ(ref, correct, isLast, answer) {
  if (!ref) return;
  xapiAnswered(ref[0], ref[1], correct, isLast, answer);
  try { if (typeof flushAfterCommit === 'function') flushAfterCommit(); } catch (e) {}
}

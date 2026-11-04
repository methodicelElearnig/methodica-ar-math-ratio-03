'use strict';
/* ═══════════════════ Boot — methodica-math-ratio-03 (all parts) ═══════════════════
   Loaded LAST, after the part's js/main.js, so every shared and per-part declaration exists
   (guide §3.2: definitions live in the earlier files; this is the one that runs them).

   The part's own init has already run at the bottom of its main.js (scaleApp, PartNav.init,
   resetScreenState(0)). Order here:
     1. the ?resetState hatch — it rewrites the URL before anything below reads it
     2. PART_LAST, the issue-report dialog and the per-item result functions
     3. the leave / edit save triggers (inert while RESUME_ENABLED is false)
     4. the xAPI layer — its landing-screen item 'initialized' needs currentScreen */
(function boot() {
  try { initResumeResetHatch(); } catch (e) { console.error('[boot] initResumeResetHatch', e); }
  window.PART_LAST = TOTAL_SCREENS - 1;   // read by 20-xapi.js restoreEndedButton()
  try { initReportModal(); } catch (e) { console.error('[boot] initReportModal', e); }
  try { buildItemResults(); } catch (e) { console.error('[boot] buildItemResults', e); }
  try { initResumeLeaveHandlers(); } catch (e) { console.error('[boot] initResumeLeaveHandlers', e); }
  try { initResumeFieldSaves(); } catch (e) { console.error('[boot] initResumeFieldSaves', e); }
  bootXAPI();
})();

// ==UserScript==
// @name         Veracross HPU Duplicate Merger
// @namespace    windward.axiom.dedup
// @version      1.0
// @description  Auto-merges "Potential Duplicates" household groups in Veracross Axiom. Selects each group, opens the lightning-bolt Actions menu, clicks "Merge Duplicate Records", handles the survivor-selection screen if it appears (keeping the lowest ID), pauses 2.5s, repeats 8x, reloads, and loops until no duplicates remain.
// @match        https://axiom.veracross.com/windward/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

/*
  IMPORTANT — READ BEFORE RUNNING
  --------------------------------
  * Merges in Veracross are IRREVERSIBLE. This script clicks the real merge
    control. There is no undo.
  * During discovery, a clean 2-record group merged IMMEDIATELY when the
    "Merge Duplicate Records" menu link was clicked — no separate confirmation
    screen appeared, and the lowest ID survived automatically. This script
    therefore ALSO contains defensive handling for a survivor-selection
    "interstitial" (radio + green button) in case larger groups (3–5 records)
    show one. That interstitial path is INFERRED, not verified — watch the
    first few runs and adjust the selectors in handleInterstitial() if needed.
  * The script does NOT auto-start. Open the Potential Duplicates results page,
    then click the floating "▶ Start Dedup" button. Click "■ Stop" any time to
    abort; it stops after the current group finishes.
  * It stops on its own when a page refresh shows no duplicate groups remain.
*/

(function () {
  'use strict';

  // ------------------------- CONFIG -------------------------
  const CFG = {
    PAUSE_MS: 2500,          // pause after each merge before the next group
    BATCH_SIZE: 8,           // merges per page load before reloading
    MENU_WAIT_MS: 6000,      // max wait for the Actions menu to open
    MERGE_WAIT_MS: 20000,    // max wait for a merge to complete (server round-trip)
    RELOAD_WAIT_MS: 25000,   // max wait for the grid to re-render after reload
    POLL_MS: 120,            // polling interval for waitFor()
    STORAGE_KEY: 'vc_dedup_running' // survives reloads so the loop resumes
  };

  // ------------------------- STATE / UI -------------------------
  let ABORT = false;
  let RUNNING = false;

  const ui = buildPanel();

  function buildPanel() {
    const wrap = document.createElement('div');
    wrap.style.cssText =
      'position:fixed;z-index:2147483647;bottom:16px;right:16px;background:#1f2937;color:#fff;' +
      'font:13px/1.4 system-ui,sans-serif;padding:10px 12px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.35);' +
      'min-width:220px;';
    wrap.innerHTML =
      '<div style="font-weight:600;margin-bottom:6px">Veracross Dedup</div>' +
      '<div id="vcd-status" style="margin-bottom:8px;color:#d1d5db">Idle.</div>' +
      '<button id="vcd-start" style="cursor:pointer;border:0;border-radius:6px;padding:6px 10px;margin-right:6px;background:#10b981;color:#062;font-weight:600">▶ Start Dedup</button>' +
      '<button id="vcd-stop" style="cursor:pointer;border:0;border-radius:6px;padding:6px 10px;background:#ef4444;color:#fff;font-weight:600">■ Stop</button>' +
      '<div id="vcd-count" style="margin-top:8px;color:#9ca3af">Merged this session: 0</div>';
    document.body.appendChild(wrap);
    wrap.querySelector('#vcd-start').onclick = () => start();
    wrap.querySelector('#vcd-stop').onclick = () => stop();
    return {
      status: (t) => (wrap.querySelector('#vcd-status').textContent = t),
      count: (n) => (wrap.querySelector('#vcd-count').textContent = 'Merged this session: ' + n)
    };
  }

  let mergedCount = 0;
  function log(msg) { ui.status(msg); console.log('[VC-Dedup]', msg); }

  // ------------------------- HELPERS -------------------------
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // Poll a predicate until it returns truthy or timeout. Returns the truthy value or null.
  async function waitFor(fn, timeout, label) {
    const deadline = Date.now() + timeout;
    while (Date.now() < deadline) {
      if (ABORT) return null;
      let v;
      try { v = fn(); } catch (e) { v = null; }
      if (v) return v;
      await sleep(CFG.POLL_MS);
    }
    if (label) console.warn('[VC-Dedup] waitFor timeout:', label);
    return null;
  }

  const visible = (el) =>
    !!el && !!(el.offsetParent || el.getClientRects().length) &&
    getComputedStyle(el).visibility !== 'hidden';

  // Dispatch a full, realistic click sequence — Wijmo/Ember often need mousedown+mouseup, not just click().
  function realClick(el) {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    const cx = r.left + r.width / 2;
    const cy = r.top + r.height / 2;
    const opts = { bubbles: true, cancelable: true, view: window, clientX: cx, clientY: cy, button: 0 };
    try { el.dispatchEvent(new PointerEvent('pointerdown', opts)); } catch (e) {}
    el.dispatchEvent(new MouseEvent('mousedown', opts));
    try { el.dispatchEvent(new PointerEvent('pointerup', opts)); } catch (e) {}
    el.dispatchEvent(new MouseEvent('mouseup', opts));
    el.dispatchEvent(new MouseEvent('click', opts));
    return true;
  }

  const byText = (selector, re) =>
    Array.from(document.querySelectorAll(selector)).filter(
      (e) => re.test((e.textContent || '').trim()) && visible(e)
    );

  // ------------------------- GRID QUERIES -------------------------
  function getGridControl() {
    try {
      const host = document.querySelector('.wj-flexgrid');
      return window.wijmo ? window.wijmo.Control.getControl(host) : null;
    } catch (e) { return null; }
  }

  function scrollGridToTop() {
    const g = getGridControl();
    try {
      if (g && g.scrollPosition && window.wijmo) {
        g.scrollPosition = new window.wijmo.Point(g.scrollPosition.x, 0);
      }
    } catch (e) {}
    // DOM fallback: scroll any overflow container inside the grid to the top.
    document.querySelectorAll('.wj-flexgrid .wj-cells, .wj-flexgrid .wj-rowheaders').forEach((c) => (c.scrollTop = 0));
  }

  // Number of duplicate groups currently rendered (each "(N records)" label = one group).
  const groupCount = () => document.querySelectorAll('.wj-group-record-count').length;

  // Total records reported in the footer (e.g. "24 records"). Returns a number or null.
  function footerRecords() {
    const el = Array.from(document.querySelectorAll('*')).find(
      (e) => e.children.length === 0 && /^\s*\d+\s+records\s*$/.test(e.textContent || '')
    );
    if (!el) return null;
    return parseInt(el.textContent.replace(/\D/g, ''), 10);
  }

  // "Nothing returned on refresh": no duplicate groups AND/OR zero records.
  function listIsEmpty() {
    const rec = footerRecords();
    const groups = groupCount();
    if (groups === 0) return true;
    if (rec === 0) return true;
    // some views show an explicit empty-state message
    if (byText('*', /no (results|records|matching)/i).length) return true;
    return false;
  }

  // The checkbox that selects the TOP duplicate group. Verified behaviour: clicking the
  // top-most group checkbox selects the whole group (header + all member records).
  function topGroupCheckbox() {
    const boxes = Array.from(document.querySelectorAll('i.grid-row-checkbox')).filter(visible);
    if (!boxes.length) return null;
    boxes.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
    return boxes[0];
  }

  const checkedCount = () =>
    document.querySelectorAll('i.grid-row-checkbox.fa-check-square').length;

  // ------------------------- WORKFLOW STEPS -------------------------

  // Select the first group. Returns true if >=2 record boxes end up checked.
  async function selectFirstGroup() {
    scrollGridToTop();
    await sleep(200);
    const box = topGroupCheckbox();
    if (!box) return false;
    realClick(box);
    // Wait for the selection to register (>=2 checkmarks: 1 header + >=2 records, or >=2 records).
    const ok = await waitFor(() => checkedCount() >= 2, 2500, 'group selection');
    if (ok) return true;

    // Fallback: the top box may have been a single record. Check the next member boxes too.
    const boxes = Array.from(document.querySelectorAll('i.grid-row-checkbox'))
      .filter(visible)
      .sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top);
    for (let i = 1; i < Math.min(boxes.length, 5); i++) {
      if (boxes[i].classList.contains('fa-square-o')) realClick(boxes[i]);
      await sleep(120);
      if (checkedCount() >= 2) return true;
    }
    return checkedCount() >= 2;
  }

  // Open the lightning-bolt Actions menu.
  async function openBoltMenu() {
    const bolt =
      document.querySelector('a.launcher-toolbar__button.actions') ||
      (document.querySelector('i.fa-bolt') && document.querySelector('i.fa-bolt').closest('a'));
    if (!bolt) return false;
    realClick(bolt);
    const link = await waitFor(
      () => byText('a.action-link', /merge duplicate records/i)[0],
      CFG.MENU_WAIT_MS,
      'actions menu'
    );
    return !!link;
  }

  // Click the blue "Merge Duplicate Records" link in the menu.
  async function clickMergeLink() {
    const link = byText('a.action-link', /merge duplicate records/i)[0];
    if (!link) return false;
    realClick(link);
    return true;
  }

  // If a survivor-selection interstitial appears, keep the LOWEST ID and click the green
  // confirm button. Returns 'interstitial' if handled, 'immediate' if none appeared,
  // 'invalid' if the "must select at least 2" validation error showed instead.
  async function handleInterstitial(groupsBefore) {
    const outcome = await waitFor(() => {
      // (a) validation error — selection was invalid
      if (byText('*', /must select at least .* records/i).length) return 'invalid';
      // (b) grid already changed => immediate merge, no interstitial
      if (groupCount() < groupsBefore) return 'immediate';
      // (c) an interstitial with a green confirm button distinct from the menu link
      const green = greenMergeButton();
      if (green) return 'interstitial';
      return null;
    }, 6000, 'post-merge state');

    if (outcome === 'interstitial') {
      selectLowestIdSurvivor();
      await sleep(300);
      const green = greenMergeButton();
      if (green) realClick(green);
      return 'interstitial';
    }
    return outcome || 'immediate';
  }

  // Find the green "Merge Duplicate Records" confirm button (NOT the blue menu link).
  function greenMergeButton() {
    const cands = Array.from(document.querySelectorAll('button, a, input[type=button], input[type=submit]'))
      .filter(visible)
      .filter((e) => e.className.indexOf('action-link') === -1)
      .filter((e) => /merge duplicate records/i.test((e.textContent || e.value || '')));
    if (!cands.length) return null;
    // Prefer one that looks green / primary; else take the last (usually the confirm).
    const green = cands.find((e) => {
      const bg = getComputedStyle(e).backgroundColor || '';
      const m = bg.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      const greenish = m && +m[2] > 120 && +m[2] > +m[1] && +m[2] > +m[3];
      return greenish || /green|success|primary|confirm/i.test(e.className);
    });
    return green || cands[cands.length - 1];
  }

  // Within the interstitial, pick the radio/row for the lowest household ID.
  function selectLowestIdSurvivor() {
    // Look for a modal/dialog container; fall back to document.
    const scope =
      document.querySelector('.action-menu, .modal, [role=dialog], .veracross-modal') || document;
    const radios = Array.from(scope.querySelectorAll('input[type=radio]')).filter(visible);
    if (!radios.length) return; // interstitial may already default to lowest ID
    let best = null, bestId = Infinity;
    radios.forEach((radio) => {
      // Find an ID number associated with this radio's row.
      const row = radio.closest('tr, li, .row, [class*=row]') || radio.parentElement;
      const txt = (row && row.textContent) || '';
      const nums = (txt.match(/\b\d{4,}\b/g) || []).map(Number);
      const id = nums.length ? Math.min(...nums) : Infinity;
      if (id < bestId) { bestId = id; best = radio; }
    });
    if (best && !best.checked) realClick(best);
  }

  // Close any leftover dialog (validation error, status popup) so the next iteration is clean.
  async function closeAnyDialog() {
    const close = byText('a, button', /^\s*close\s*$/i)[0];
    if (close) { realClick(close); await sleep(300); }
  }

  // Wait for a merge to actually complete: group count drops, or footer count drops.
  async function waitMergeComplete(groupsBefore, recordsBefore) {
    return waitFor(() => {
      if (groupCount() < groupsBefore) return true;
      const r = footerRecords();
      if (recordsBefore != null && r != null && r < recordsBefore) return true;
      return false;
    }, CFG.MERGE_WAIT_MS, 'merge completion');
  }

  // Process exactly one group. Returns true on a successful merge.
  async function mergeOneGroup() {
    if (listIsEmpty()) return false;
    const groupsBefore = groupCount();
    const recordsBefore = footerRecords();

    if (!(await selectFirstGroup())) { log('Could not select a group (need >=2 records).'); return false; }
    if (ABORT) return false;

    if (!(await openBoltMenu())) { log('Actions menu did not open.'); return false; }
    if (ABORT) return false;

    if (!(await clickMergeLink())) { log('Merge link not found.'); return false; }

    const outcome = await handleInterstitial(groupsBefore);
    if (outcome === 'invalid') {
      log('Validation error — selection was invalid; skipping.');
      await closeAnyDialog();
      return false;
    }

    const done = await waitMergeComplete(groupsBefore, recordsBefore);
    await closeAnyDialog();
    if (done) { mergedCount++; ui.count(mergedCount); }
    return !!done;
  }

  // ------------------------- MAIN LOOP -------------------------
  async function runBatchThenMaybeReload() {
    for (let i = 0; i < CFG.BATCH_SIZE; i++) {
      if (ABORT) { finishAbort(); return; }
      if (listIsEmpty()) { finishDone(); return; }
      log(`Merging group ${i + 1}/${CFG.BATCH_SIZE} in this batch…`);
      const ok = await mergeOneGroup();
      if (ABORT) { finishAbort(); return; }
      if (!ok) {
        // Nothing merged — likely empty or a snag. Reload once to re-check.
        break;
      }
      log(`Merged. Pausing ${CFG.PAUSE_MS / 1000}s…`);
      await sleep(CFG.PAUSE_MS);
    }

    if (ABORT) { finishAbort(); return; }
    // Persist the running flag across the reload, then reload.
    sessionStorage.setItem(CFG.STORAGE_KEY, '1');
    log('Batch complete — reloading page…');
    await sleep(500);
    location.reload();
  }

  function finishDone() {
    RUNNING = false;
    sessionStorage.removeItem(CFG.STORAGE_KEY);
    log('Done — no duplicate groups remaining. ✅');
  }
  function finishAbort() {
    RUNNING = false;
    sessionStorage.removeItem(CFG.STORAGE_KEY);
    log('Stopped by user.');
  }

  async function start(resumed) {
    if (RUNNING) return;
    ABORT = false;
    RUNNING = true;
    if (!resumed) { mergedCount = 0; ui.count(0); }
    log('Waiting for grid…');
    const ready = await waitFor(() => document.querySelector('i.grid-row-checkbox') || listIsEmpty(), CFG.RELOAD_WAIT_MS, 'grid ready');
    if (!ready) { RUNNING = false; log('Grid did not load.'); return; }
    if (listIsEmpty()) { finishDone(); return; }
    runBatchThenMaybeReload();
  }

  function stop() {
    ABORT = true;
    sessionStorage.removeItem(CFG.STORAGE_KEY);
    log('Stopping after current group…');
  }

  // Auto-resume after a reload that the script itself triggered.
  if (sessionStorage.getItem(CFG.STORAGE_KEY) === '1') {
    log('Resuming after reload…');
    // give the SPA a moment to boot
    setTimeout(() => start(true), 2500);
  }
})();

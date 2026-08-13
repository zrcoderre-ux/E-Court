/**
 * LA Court E-Court Suite — Documents tab: when each filing was POSTED
 *
 * The Documents tab lists a filing DATE, which is the paper's effective date,
 * not the day it showed up on the site. Those differ: the clerk's intake queue
 * runs during business hours and lands a paper 0-3 court days after it was
 * filed, so a reply filed on the day it was due is routinely invisible until
 * the next court day. This surfaces the day and time eCourt actually posted
 * each filing, decoded out of the doc endpoint's Last-Modified header (see
 * lib/case-status.js), as a hover under the row's Filed date — in place of the
 * site's own "UPDATE DOCUMENT" tooltip, which that cell otherwise shows.
 *
 * The lag in court days is computed for the samples only. On the page the date
 * itself is the useful fact; the lag distribution belongs to the options page.
 *
 * The lookup is one HEAD per document, cached permanently by docId, so a case
 * costs its requests once. Every decoded (filed -> posted) pair is also handed
 * to the background as a lag sample, which the options page totals into the
 * distribution the grace window is set from.
 */

(function () {
  'use strict';

  if (!window.LACCaseStatus) return;
  const { parseDocRows, getIngestTime, fmtIngest, ingestDay, ingestLagCourtDays } = LACCaseStatus;

  const MARK_ATTR = 'data-lac-ingest';
  const CONCURRENCY = 4;
  const LOG = (...a) => { try { console.log('[LACourt-Ingest]', ...a); } catch (_) {} };

  // Only the Documents tab lists openable filings; every other case sub-tab
  // parses to nothing and this is a no-op.
  function docAnchors() {
    const out = new Map(); // docId -> anchor
    for (const a of document.querySelectorAll('a[onclick*="openInNewWindow"]')) {
      const m = (a.getAttribute('onclick') || '')
        .match(/openInNewWindow\('((?:[^'\\]|\\.)*)'/);
      if (!m) continue;
      const idm = m[1].replace(/\\\//g, '/').match(/docId=(\d+)/);
      if (idm && !out.has(idm[1])) out.set(idm[1], a);
    }
    return out;
  }

  const STYLE_ID = 'lac-ingest-styles';
  const CELL_CLASS = 'lac-upload-cell';
  const STAMP_CLASS = 'lac-upload-stamp';

  // Absolutely positioned so revealing it doesn't reflow the table — it drops
  // over the row below the way the native tooltip it replaces would.
  function ensureStyles() {
    if (document.getElementById(STYLE_ID)) return;
    const s = document.createElement('style');
    s.id = STYLE_ID;
    s.textContent = '.' + CELL_CLASS + '{position:relative;}'
      + '.' + STAMP_CLASS + '{display:none;position:absolute;left:0;top:100%;margin-top:2px;'
      + 'z-index:2147483000;padding:3px 7px;border-radius:3px;background:#1f2937;color:#f9fafb;'
      + 'font:600 11px system-ui,sans-serif;white-space:nowrap;pointer-events:none;'
      + 'box-shadow:0 1px 4px rgba(0,0,0,.35);}'
      + '.' + CELL_CLASS + ':hover .' + STAMP_CLASS + '{display:block;}';
    (document.head || document.documentElement).appendChild(s);
  }

  // The row's Filed/Status Date cell. It is column 1 by header order, but the
  // date pattern is checked rather than trusted, and scanned for if that misses.
  function dateCellFor(anchor) {
    const tr = anchor.closest && anchor.closest('tr');
    if (!tr) return null;
    const cells = Array.from(tr.children);
    const isDate = c => /^\s*\d{1,2}\/\d{1,2}\/\d{4}\s*$/.test(c.textContent || '');
    if (cells[1] && isDate(cells[1])) return cells[1];
    return cells.filter(isDate)[0] || null;
  }

  // eCourt hangs an "UPDATE DOCUMENT" tooltip off this cell. Stash the titles
  // and clear them so hovering shows the upload time instead of the site's.
  function suppressNativeTitles(cell) {
    const els = [cell].concat(Array.from(cell.querySelectorAll('[title]')));
    for (const el of els) {
      const t = el.getAttribute && el.getAttribute('title');
      if (t) { el.setAttribute('data-lac-title', t); el.removeAttribute('title'); }
    }
  }

  function stamp(anchor, info) {
    if (!anchor || anchor.getAttribute(MARK_ATTR) === '1') return;
    const text = fmtIngest(info);
    if (!text) return;
    const cell = dateCellFor(anchor);
    if (!cell) return;
    anchor.setAttribute(MARK_ATTR, '1');
    ensureStyles();
    suppressNativeTitles(cell);
    cell.classList.add(CELL_CLASS);
    const prev = cell.querySelector('.' + STAMP_CLASS);
    if (prev) prev.remove();
    const span = document.createElement('span');
    span.className = STAMP_CLASS;
    span.textContent = 'Uploaded ' + text;
    cell.appendChild(span);
  }

  function runWithConcurrency(items, limit, worker) {
    let i = 0;
    const next = async () => { while (i < items.length) { const idx = i++; try { await worker(items[idx]); } catch (_) {} } };
    return Promise.all(Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, next));
  }

  async function annotate() {
    const anchors = docAnchors();
    if (!anchors.size) return;

    // parseDocRows reads the filed date off the same anchors, so the lag is
    // computed against exactly the date the row displays.
    const filedById = new Map();
    for (const r of parseDocRows(document)) {
      if (r.when) filedById.set(r.docId, { when: r.when, name: r.name, dateStr: r.dateStr, filedBy: r.filedBy });
    }

    const jobs = [];
    for (const [docId, anchor] of anchors) {
      if (anchor.getAttribute(MARK_ATTR) === '1') continue;
      jobs.push({ docId, anchor, filed: filedById.get(docId) || null });
    }
    if (!jobs.length) return;

    const samples = [];
    await runWithConcurrency(jobs, CONCURRENCY, async job => {
      const info = await getIngestTime(job.docId);
      if (!info) return;
      const filedWhen = job.filed ? job.filed.when : null;
      stamp(job.anchor, info);
      const day = ingestDay(info);
      const lag = ingestLagCourtDays(filedWhen, info);
      // A paper posted before its own filing date isn't measuring the intake
      // queue (a filing date set forward, a nunc pro tunc entry), so it stays
      // out of the distribution rather than dragging the medians down.
      if (filedWhen && day && lag != null && lag >= 0) {
        samples.push({
          docId: job.docId,
          name: (job.filed && job.filed.name || '').slice(0, 80),
          filedBy: (job.filed && job.filed.filedBy || '').slice(0, 80),
          filed: job.filed.dateStr || '',
          posted: day.getTime(),
          postedText: fmtIngest(info),
          lag,
        });
      }
    });

    LOG('stamped', jobs.length, 'documents;', samples.length, 'lag samples');
    if (samples.length) {
      try {
        chrome.runtime.sendMessage(
          { type: 'recordFilingLag', samples }, () => { void chrome.runtime.lastError; });
      } catch (_) {}
    }
  }

  // The tab paginates in place, so re-annotate when new rows arrive. Guarded
  // against our own inserts by the marker attribute plus a mutation flag.
  let queued = false;
  function schedule() {
    if (queued) return;
    queued = true;
    setTimeout(() => { queued = false; annotate().catch(() => {}); }, 300);
  }

  function start() {
    schedule();
    try {
      const mo = new MutationObserver(muts => {
        for (const m of muts) {
          for (const n of m.addedNodes) {
            if (n.nodeType === 1 && !(n.classList && n.classList.contains(STAMP_CLASS))) { schedule(); return; }
          }
        }
      });
      mo.observe(document.documentElement, { childList: true, subtree: true });
    } catch (_) {}
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();

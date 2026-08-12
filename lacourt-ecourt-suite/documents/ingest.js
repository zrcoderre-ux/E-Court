/**
 * LA Court E-Court Suite — Documents tab: when each filing was POSTED
 *
 * The Documents tab lists a filing DATE, which is the paper's effective date,
 * not the day it showed up on the site. Those differ: the clerk's intake queue
 * runs during business hours and lands a paper 0-3 court days after it was
 * filed, so a reply filed on the day it was due is routinely invisible until
 * the next court day. This annotates each document row with the day and time
 * eCourt actually posted it, plus the lag in court days, decoded out of the doc
 * endpoint's Last-Modified header (see lib/case-status.js).
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

  // "Thu 8/6 10:43a · +1 ct day" — the posting moment and how long it sat.
  function labelFor(info, filedWhen) {
    const when = fmtIngest(info);
    if (!when) return '';
    const lag = ingestLagCourtDays(filedWhen, info);
    if (lag == null) return when;
    if (lag <= 0) return when + ' · same day';
    return when + ' · +' + lag + ' ct day' + (lag === 1 ? '' : 's');
  }

  // A filing that sat longer than this reads amber — it is the case where the
  // paper was on file well before you could have seen it.
  const SLOW_COURT_DAYS = 2;

  function stamp(anchor, info, filedWhen) {
    if (!anchor || anchor.getAttribute(MARK_ATTR) === '1') return;
    const text = labelFor(info, filedWhen);
    if (!text) return;
    anchor.setAttribute(MARK_ATTR, '1');
    const lag = ingestLagCourtDays(filedWhen, info);
    const span = document.createElement('span');
    span.className = 'lac-ingest-stamp';
    span.textContent = text;
    span.title = 'Posted to eCourt at this time (±17 min). The Filed column shows the '
      + 'effective filing date, which can be several court days earlier.';
    span.setAttribute('style',
      'display:inline-block;margin-left:8px;font:600 11px system-ui,sans-serif;'
      + 'white-space:nowrap;color:' + (lag != null && lag > SLOW_COURT_DAYS ? '#b8860b' : '#6b7280') + ';');
    // Sits after the document link, inside the same cell, so it travels with
    // the row through pagination and filtering.
    if (anchor.parentNode) anchor.parentNode.insertBefore(span, anchor.nextSibling);
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
      if (r.when) filedById.set(r.docId, { when: r.when, name: r.name, dateStr: r.dateStr });
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
      stamp(job.anchor, info, filedWhen);
      const day = ingestDay(info);
      if (filedWhen && day) {
        samples.push({
          docId: job.docId,
          name: (job.filed && job.filed.name || '').slice(0, 80),
          filed: job.filed.dateStr || '',
          posted: day.getTime(),
          postedText: fmtIngest(info),
          lag: ingestLagCourtDays(filedWhen, info),
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
            if (n.nodeType === 1 && !(n.classList && n.classList.contains('lac-ingest-stamp'))) { schedule(); return; }
          }
        }
      });
      mo.observe(document.documentElement, { childList: true, subtree: true });
    } catch (_) {}
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();

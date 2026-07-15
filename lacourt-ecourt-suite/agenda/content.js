/**
 * LA Court Agenda Cleaner - Content Script
 *
 * On agenda pages (civil.lacourt.org/ecourt/ecms/agenda...):
 *   - Reads directly from the live #day-table in the DOM
 *   - Extracts hearing text from <b> tags (avoids React/script content)
 *   - Extracts case text from <a> tags
 *   - Manual selection: restricts to rows intersecting the selection
 *   - Ctrl+A / broad selection: processes all visible rows
 *   - Single cell: passes through unchanged
 *   - Applies exclusion list (editable via popup)
 *   - Output: blank col A, hearing col B, case col C, blank row after each
 *   - Font: Times New Roman 22pt
 */

/* -------------------------------------------------
   EXCLUSION TERMS
------------------------------------------------- */

const DEFAULT_EXCLUDED_TERMS = [
  'conference',
  'non-appearance case revie',
  'non-jury trial',
  'order to show cause re: d',
  'ex parte',
  'application for order for',
  'jury trial',
  'post-arbitration status c',
  'post-mediation status con',
  'order to show cause re: s',
  'informal discovery confer',
];

let EXCLUDED_TERMS = [...DEFAULT_EXCLUDED_TERMS];

chrome.storage.sync.get(['excludedTerms'], result => {
  if (result.excludedTerms && Array.isArray(result.excludedTerms)) {
    EXCLUDED_TERMS = result.excludedTerms;
  }
});

chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'sync' && changes.excludedTerms) {
    EXCLUDED_TERMS = changes.excludedTerms.newValue || [...DEFAULT_EXCLUDED_TERMS];
  }
});

/* -------------------------------------------------
   COPY HANDLER
------------------------------------------------- */

document.addEventListener('copy', function (e) {
  const selection = window.getSelection();
  if (!selection || selection.isCollapsed) return;
  handleAgendaCopy(e, selection);
});

/* -------------------------------------------------
   AGENDA HANDLER
------------------------------------------------- */

function handleAgendaCopy(e, selection) {
  const table = document.getElementById('day-table');
  if (!table) return;

  // Single-cell selection: pass through unchanged
  const range = selection.getRangeAt(0);
  const frag = range.cloneContents();
  const tmp = document.createElement('div');
  tmp.appendChild(frag);
  if (tmp.querySelectorAll('td, th').length <= 1) return;

  // All visible data rows
  const allDataRows = Array.from(table.querySelectorAll('tr.js-row'))
    .filter(row => row.style.display !== 'none');
  if (allDataRows.length === 0) return;

  // Restrict to intersected rows for manual selections;
  // fall back to all rows if intersection finds nothing (e.g. Ctrl+A)
  // Restrict to rows that visually overlap the selection rectangle.
  // This is more reliable than intersectsNode for partial table row selections.
  let candidateRows = allDataRows;
  try {
    const selRect = range.getBoundingClientRect();
    // Only filter if we got a real rect (non-zero area means a real selection)
    if (selRect && selRect.width + selRect.height > 0) {
      const intersected = allDataRows.filter(row => {
        const rowRect = row.getBoundingClientRect();
        return rowRect.bottom > selRect.top && rowRect.top < selRect.bottom;
      });
      if (intersected.length > 0) candidateRows = intersected;
    }
  } catch (err) {
    // fall back to all rows
  }

  const outputRows = buildAgendaOutputRows(candidateRows);

  if (outputRows.length === 0) {
    e.clipboardData.setData('text/plain', '');
    e.clipboardData.setData('text/html', '');
    e.preventDefault();
    return;
  }

  const { plainText, html } = buildAgendaPayload(outputRows);
  e.clipboardData.setData('text/plain', plainText);
  e.clipboardData.setData('text/html', html);
  e.preventDefault();
}

/* -------------------------------------------------
   OUTPUT BUILDERS (shared by the copy handler + Copy All button)
------------------------------------------------- */

// Turn day-table rows into cleaned output rows: { col5 hearing, col6 case, url }.
function buildAgendaOutputRows(rows) {
  const outputRows = [];
  rows.forEach(row => {
    const cells = row.querySelectorAll('td');
    if (cells.length < 7) return;

    const hearingText = extractHearingText(cells[5]);
    const caseInfo    = extractCaseText(cells[6]);
    if (!hearingText && !caseInfo.text) return;

    const hearingResult = filterExcluded(hearingText);
    if (hearingResult === null) return;

    // Ensure no newlines remain that could split cells in Excel
    const hearingClean = hearingResult.replace(/\n/g, ' ').replace(/\s{2,}/g, ' ').trim();
    outputRows.push({ col5: hearingClean, col6: caseInfo.text, url: caseInfo.url });
  });
  return outputRows;
}

// Build the plain-text (tab-separated, blank row after each) and Times New Roman
// 22pt HTML clipboard payloads from cleaned output rows.
function buildAgendaPayload(outputRows) {
  const plainLines = outputRows.flatMap(r => [r.col5 + '\t' + r.col6, '']);
  const plainText  = plainLines.join('\n');

  const htmlRows = outputRows.flatMap(r => {
    const c5 = escapeHtml(r.col5);
    const c6 = escapeHtml(r.col6);
    const caseCell = r.url
      ? '<a href="' + escapeHtml(r.url) + '">' + c6 + '</a>'
      : c6;
    return [
      '<tr>' +
        '<td style="font-family:\'Times New Roman\',serif;font-size:22pt;">' + c5 + '</td>' +
        '<td style="font-family:\'Times New Roman\',serif;font-size:22pt;">' + caseCell + '</td>' +
      '</tr>',
      '<tr><td></td><td></td></tr>'
    ];
  }).join('');

  const html = '<html><body><table>' + htmlRows + '</table></body></html>';
  return { plainText, html };
}

/* -------------------------------------------------
   COPY ALL BUTTON (top-right; no Ctrl+A / Ctrl+C needed)
------------------------------------------------- */

const COPY_ALL_BTN_ID = '__lacourt_agenda_copy_all__';

function agendaToast(message, ok) {
  let t = document.getElementById('__lacourt_agenda_toast__');
  if (!t) {
    t = document.createElement('div');
    t.id = '__lacourt_agenda_toast__';
    t.style.cssText = 'position:fixed;top:56px;right:16px;z-index:2147483647;'
      + 'padding:8px 14px;border-radius:6px;font:600 13px system-ui,sans-serif;'
      + 'color:#fff;box-shadow:0 2px 8px rgba(0,0,0,.3);transition:opacity .3s;';
    document.body.appendChild(t);
  }
  t.textContent = message;
  t.style.background = ok === false ? '#c0392b' : '#0a6e6e';
  t.style.opacity = '1';
  clearTimeout(t.__hide);
  t.__hide = setTimeout(() => { t.style.opacity = '0'; }, 2200);
}

async function copyAllAgenda(btn) {
  const table = document.getElementById('day-table');
  if (!table) { agendaToast('No agenda table found', false); return; }

  const label = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = 'Copying…'; }
  // Make sure truncated hearing names are expanded first so the copy is complete.
  try { if (typeof expandTruncatedHearings === 'function') await expandTruncatedHearings(); } catch (_) {}

  const rows = Array.from(table.querySelectorAll('tr.js-row')).filter(r => r.style.display !== 'none');
  const outputRows = buildAgendaOutputRows(rows);
  const restore = () => { if (btn) { btn.disabled = false; btn.textContent = label || 'Copy All'; } };

  if (!outputRows.length) { agendaToast('Nothing to copy', false); restore(); return; }

  const { plainText, html } = buildAgendaPayload(outputRows);
  try {
    await navigator.clipboard.write([new ClipboardItem({
      'text/plain': new Blob([plainText], { type: 'text/plain' }),
      'text/html': new Blob([html], { type: 'text/html' }),
    })]);
    agendaToast('Copied ' + outputRows.length + ' hearing' + (outputRows.length === 1 ? '' : 's'), true);
  } catch (e) {
    // Fallback: plain text only (older browsers / no ClipboardItem).
    try {
      await navigator.clipboard.writeText(plainText);
      agendaToast('Copied ' + outputRows.length + ' (text only)', true);
    } catch (_) {
      agendaToast('Copy failed — use Ctrl+A then Ctrl+C', false);
    }
  }
  restore();
}

function renderCopyAllButton() {
  if (document.getElementById(COPY_ALL_BTN_ID)) return;
  if (!document.getElementById('day-table') || !document.body) return;
  const btn = document.createElement('button');
  btn.id = COPY_ALL_BTN_ID;
  btn.type = 'button';
  btn.textContent = 'Copy All';
  btn.style.cssText = 'position:fixed;top:14px;right:16px;z-index:2147483647;'
    + 'padding:8px 16px;border:none;border-radius:6px;cursor:pointer;'
    + 'background:#0a6e6e;color:#fff;font:600 13px system-ui,sans-serif;'
    + 'box-shadow:0 2px 6px rgba(0,0,0,.3);';
  btn.addEventListener('mouseenter', () => { btn.style.background = '#0d8f8f'; });
  btn.addEventListener('mouseleave', () => { btn.style.background = '#0a6e6e'; });
  btn.addEventListener('click', () => copyAllAgenda(btn));
  document.body.appendChild(btn);
}

(function initCopyAllButton() {
  const start = () => {
    renderCopyAllButton();
    // The day-table can render late / re-render; keep the button pinned.
    try { new MutationObserver(renderCopyAllButton).observe(document.body, { childList: true, subtree: true }); } catch (_) {}
  };
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start, { once: true });
  } else {
    start();
  }
})();

/* -------------------------------------------------
   TEXT EXTRACTION
------------------------------------------------- */

function extractHearingText(cell) {
  const bTags = cell.querySelectorAll('b');
  if (bTags.length === 0) return cell.textContent.trim().replace(/\s+/g, ' ');
  return Array.from(bTags).map(b => b.textContent.trim().replace(/\s+/g, ' ')).join('\n');
}

function extractCaseText(cell) {
  const a = cell.querySelector('a');
  if (a) {
    // a.href returns the resolved absolute URL even if the markup uses a relative href
    const href = a.href || '';
    return {
      text: a.textContent.trim().replace(/\s+/g, ' '),
      url: href
    };
  }
  return {
    text: cell.textContent.trim().replace(/\s+/g, ' '),
    url: ''
  };
}

/* -------------------------------------------------
   EXCLUSION FILTERING
------------------------------------------------- */

function isExcluded(segment) {
  const lower = segment.trim().toLowerCase();
  return EXCLUDED_TERMS.some(term => lower.includes(term));
}

function stripBullet(seg) {
  return seg.replace(/^[\s\u2022\u2013\u2014\*\-•]+/, '').trim();
}

function filterExcluded(text) {
  if (!text) return text;
  const hasSoftReturn = text.includes('\n');
  if (!hasSoftReturn) return isExcluded(text) ? null : text;

  const segments = text.split('\n');
  const kept = segments
    .map(seg => seg.trim())
    .filter(seg => seg.length > 0 && !isExcluded(seg));

  if (kept.length === 0) return null;
  // Strip bullets from all kept segments, then join with a space
  return kept.map(seg => stripBullet(seg)).join(' ');
}

/* -------------------------------------------------
   FULL HEARING NAME EXPANSION

   The agenda truncates long hearing names in the day-table (server-side, e.g.
   "Demurrer - without Motion..."). The per-hearing event page is an AJAX shell
   that doesn't contain the hearing name, so we recover the full name from the
   case's Hearings tab (case?id=<caseId>&formId=395), which lists every hearing
   as "<name> <MM/DD/YYYY HH:MM AM/PM> <status>". We match the Scheduled hearing
   on the agenda's day whose name (minus a leading "Hearing on ") starts with
   the truncated text, and drop the full name back into the <b> in place — so
   both the on-page display and the copy output carry it, and the exclusion
   check runs against the full name. Case fetches are cached by case id.
------------------------------------------------- */

const CASE_HEARINGS_CACHE = new Map(); // caseId -> hearings array or in-flight Promise
const EXPANDED_ATTR = 'data-lac-expanded';

function isTruncatedName(text) { return /(?:\.\.\.|…)\s*$/.test(text); }
function truncatedPrefix(text) { return text.replace(/\s*(?:\.\.\.|…)\s*$/, '').trim(); }
function normDate(s) { const m = (s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? (+m[1]) + '/' + (+m[2]) + '/' + (+m[3]) : ''; }
function stripHearingOn(s) { return (s || '').replace(/^\s*hearing on\s+/i, '').trim(); }
// Drop a trailing purely-numeric parenthetical, e.g. "...defendant's aka (6861)"
// -> "...defendant's aka". Alphanumeric ones like "(CCP 437c)" are kept.
// KEEP IN SYNC with stripTrailingParenNumber in clipboard/content.js.
function stripTrailingParenNumber(s) {
  if (!s) return s;
  return s.replace(/\s*\(\s*\d[\d\s.,\-]*\)\s*$/, '').trim();
}

// The agenda day (single-day view) from the URL's ?day= param, normalized.
function agendaDay() {
  try { return normDate(new URLSearchParams(location.search).get('day') || ''); } catch (_) { return ''; }
}

function fetchWithTimeout(url, ms) {
  return Promise.race([
    fetch(url, { credentials: 'include' }),
    new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), ms)),
  ]);
}

// Parse a fetched case Hearings-tab document into [{ name, dateTime, status }].
// Scans every row for a "MM/DD/YYYY HH:MM AM/PM" cell. Rows interleave EMPTY
// cells (observed live: ["", <name>, "", <date/time>, <status>, …]), so the
// hearing name is the nearest NON-EMPTY cell before the date, and the status is
// the first non-empty cell after it. Doesn't depend on exact column positions.
function parseCaseHearings(doc) {
  const out = [];
  for (const tr of doc.querySelectorAll('tr')) {
    const cells = Array.from(tr.children).map(c => (c.textContent || '').replace(/\s+/g, ' ').trim());
    const dtIdx = cells.findIndex(c => /^\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)\b/i.test(c));
    if (dtIdx < 1) continue;
    let name = '';
    for (let i = dtIdx - 1; i >= 0; i--) { if (cells[i]) { name = cells[i]; break; } }
    if (!name || name.length > 200 || !/[a-z]/i.test(name)) continue;
    let status = '';
    for (let i = dtIdx + 1; i < cells.length; i++) { if (cells[i]) { status = cells[i]; break; } }
    out.push({ name, dateTime: cells[dtIdx], status });
  }
  return out;
}

async function fetchCaseHearings(caseId) {
  try {
    const res = await fetchWithTimeout('/ecourt/ecms/case?id=' + caseId + '&formId=395', 8000);
    if (!res || !res.ok) return [];
    return parseCaseHearings(new DOMParser().parseFromString(await res.text(), 'text/html'));
  } catch (_) { return []; }
}

function getCaseHearings(caseId) {
  let v = CASE_HEARINGS_CACHE.get(caseId);
  if (v === undefined) {
    v = fetchCaseHearings(caseId);
    CASE_HEARINGS_CACHE.set(caseId, v);
    v.then(r => CASE_HEARINGS_CACHE.set(caseId, r), () => CASE_HEARINGS_CACHE.set(caseId, []));
  }
  return Promise.resolve(v);
}

// Full agenda name = a Scheduled hearing (preferably on the agenda's day) whose
// "Hearing on"-stripped name starts with the truncated prefix.
function fullNameForHearing(hearings, day, prefix) {
  const p = prefix.toLowerCase();
  const scheduled = hearings.filter(h => /scheduled/i.test(h.status));
  const find = (list, requireDay) => {
    for (const h of list) {
      if (requireDay && day && normDate(h.dateTime) !== day) continue;
      const clean = stripHearingOn(h.name);
      if (clean.length > prefix.length && clean.toLowerCase().startsWith(p)) return clean;
    }
    return '';
  };
  return find(scheduled, true) || find(scheduled, false) || find(hearings, true) || '';
}

// Run async worker over items with limited concurrency (the case cache dedups
// repeat fetches, so same-case jobs still share one request).
function runWithConcurrency(items, limit, worker) {
  let i = 0;
  const next = async () => { while (i < items.length) { const idx = i++; try { await worker(items[idx]); } catch (_) {} } };
  return Promise.all(Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, next));
}

async function expandTruncatedHearings() {
  const table = document.getElementById('day-table');
  if (!table) return;
  const day = agendaDay();

  const jobs = [];
  for (const row of table.querySelectorAll('tr.js-row')) {
    const caseA = row.querySelector('td a[href*="/ecourt/ecms/case"]');
    const idm = (caseA ? (caseA.getAttribute('href') || caseA.href || '') : '').match(/[?&]id=(\d+)/);
    if (!idm) continue;
    // A row can list several hearings (bulleted) — check every one, not just
    // the first (e.g. "Jury Trial" + a truncated "Motion to Deem Request fo…").
    for (const b of row.querySelectorAll('a[href*="/ecourt/ecms/agenda/event"] b')) {
      if (b.getAttribute(EXPANDED_ATTR) === '1') continue;
      const text = (b.textContent || '').replace(/\s+/g, ' ').trim();
      if (!isTruncatedName(text)) continue;
      jobs.push({ b, caseId: idm[1], prefix: truncatedPrefix(text) });
    }
  }
  if (!jobs.length) return;

  await runWithConcurrency(jobs, 4, async job => {
    const hearings = await getCaseHearings(job.caseId);
    const full = stripTrailingParenNumber(fullNameForHearing(hearings, day, job.prefix));
    if (full && job.b.getAttribute(EXPANDED_ATTR) !== '1') {
      job.b.textContent = full;
      job.b.setAttribute(EXPANDED_ATTR, '1');
      try { console.log('[LACourt-Agenda] expanded:', job.prefix + '…', '->', full); } catch (_) {}
    }
  });
}

(function initHearingExpansion() {
  let pending = false;
  const run = () => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => { pending = false; expandTruncatedHearings(); });
  };
  const start = () => {
    run();
    // Re-run when the day-table re-renders or paginates. Our own text swaps
    // no-op on the next pass (already-expanded / no longer truncated).
    const target = document.getElementById('day-table') || document.body;
    try { new MutationObserver(run).observe(target, { childList: true, subtree: true }); } catch (_) {}
  };
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start, { once: true });
  } else {
    start();
  }
})();

/* -------------------------------------------------
   UTILITIES
------------------------------------------------- */

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

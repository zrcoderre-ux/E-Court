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
  try { applyAgendaChanges([]); } catch (_) {}
});

chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'sync' && changes.excludedTerms) {
    EXCLUDED_TERMS = changes.excludedTerms.newValue || [...DEFAULT_EXCLUDED_TERMS];
    try { applyAgendaChanges([]); } catch (_) {}
  }
});

/* -------------------------------------------------
   COLOR CODING

   Hearings that Copy All WILL take (not on the exclusion list) render green
   instead of the default blue link color; excluded ones are left as-is.
   Idempotent — re-run after name expansion and whenever the exclusion list
   changes.
------------------------------------------------- */

const COPY_GREEN = '#1a6b3a';

function colorizeAgendaRows() {
  const table = document.getElementById('day-table');
  if (!table) return;
  for (const row of table.querySelectorAll('tr.js-row')) {
    if (row.style.display === 'none') continue;
    const cells = row.querySelectorAll('td');
    if (cells.length < 7) continue;
    for (const a of cells[5].querySelectorAll('a')) {
      const b = a.querySelector('b');
      const txt = stripTrailingParenNumber(((b || a).textContent || '').replace(/\s+/g, ' ').trim());
      if (!txt) continue;
      if (!isExcluded(txt)) {
        a.style.setProperty('color', COPY_GREEN, 'important');
        if (b) b.style.setProperty('color', COPY_GREEN, 'important');
      } else {
        a.style.removeProperty('color');
        if (b) b.style.removeProperty('color');
      }
    }
    // Colour the case link green too when the row has a green (will-be-copied)
    // hearing, so the case that corresponds to it matches.
    const rowGreen = rowHasGreenHearing(row);
    for (const a of cells[6].querySelectorAll('a')) {
      if (rowGreen) a.style.setProperty('color', COPY_GREEN, 'important');
      else a.style.removeProperty('color');
    }
  }
}

// Rewrite the VISIBLE hearing labels to drop a trailing event number like
// " (4557)" so it doesn't show on the day-table (it was already stripped from
// the copy output and exclusion checks, but the on-page text still showed it
// for non-truncated names). Only writes when the text actually changes, so it's
// idempotent and doesn't churn the observer.
function stripHearingLabelNumbers() {
  const table = document.getElementById('day-table');
  if (!table) return;
  for (const row of table.querySelectorAll('tr.js-row')) {
    const cells = row.querySelectorAll('td');
    if (cells.length < 7) continue;
    for (const a of cells[5].querySelectorAll('a')) {
      const b = a.querySelector('b');
      const el = b || a;
      const cur = (el.textContent || '').replace(/\s+/g, ' ').trim();
      const stripped = stripTrailingParenNumber(cur);
      if (stripped && stripped !== cur) el.textContent = stripped;
    }
  }
}

/* -------------------------------------------------
   AUTO-SORT + GREEN-ROWS-TO-TOP

   The agenda's "Hearing / Documents" column sort is eCourt's own client-side
   tablesort, fired by the column header's inline onclick
   (eCourt.DayView.Rows.sortByHeader(2, th)). We can't call eCourt.* from the
   isolated content-script world, but clicking the <th> runs its handler in the
   page world. We apply that sort once on load (unless the column is already the
   active sort), then float the green (will-be-copied) hearings to the top while
   preserving the site's sort order within each group.
------------------------------------------------- */

// A row is "green" (Copy All will take it) when any hearing link in the
// Hearing/Documents cell is NOT excluded — mirrors colorizeAgendaRows().
function rowHasGreenHearing(row) {
  const cells = row.querySelectorAll('td');
  if (cells.length < 7) return false;
  for (const a of cells[5].querySelectorAll('a')) {
    const b = a.querySelector('b');
    const txt = stripTrailingParenNumber(((b || a).textContent || '').replace(/\s+/g, ' ').trim());
    if (txt && !isExcluded(txt)) return true;
  }
  return false;
}

// Apply the native "Hearing / Documents" sort once per page load. Only clicks
// when that column isn't already the active sort, so we never toggle a sort the
// user (or the page default) already set to ascending/descending.
let __autoSortApplied = false;
function applyHearingDocsSort() {
  if (__autoSortApplied) return;
  const table = document.getElementById('day-table');
  if (!table) return;
  const th = table.querySelector('th.js-row-header-2');
  if (!th) return;
  __autoSortApplied = true;
  if (th.hasAttribute('aria-sort')) return; // already sorted by this column
  try { th.click(); } catch (_) {}
}

// Move green rows to the top, keeping the site's current order within the green
// and non-green groups (stable partition). No-ops when already ordered so it
// doesn't feed its own MutationObserver into a loop.
function floatGreenRowsToTop() {
  const table = document.getElementById('day-table');
  if (!table) return;
  const rows = Array.from(table.querySelectorAll('tr.js-row'));
  if (rows.length < 2) return;
  const greens = [], others = [];
  for (const r of rows) {
    if (r.style.display !== 'none' && rowHasGreenHearing(r)) greens.push(r);
    else others.push(r);
  }
  if (!greens.length) return;
  const target = greens.concat(others);
  let same = true;
  for (let i = 0; i < rows.length; i++) { if (rows[i] !== target[i]) { same = false; break; } }
  if (same) return; // already in the desired order — don't churn the DOM
  const parent = rows[0].parentNode;
  const anchor = rows[0].previousSibling; // node just before the row block (may be null)
  const frag = document.createDocumentFragment();
  target.forEach(r => frag.appendChild(r));
  if (anchor && anchor.parentNode === parent) anchor.after(frag);
  else parent.insertBefore(frag, parent.firstChild);
}

/* -------------------------------------------------
   COPY BEHAVIOR

   Manual copying (Ctrl+C / Ctrl+A) is left completely untouched — the browser
   copies exactly what's selected. The agenda cleaning (exclusion filtering,
   two-column layout, Times New Roman formatting) happens ONLY through the
   dedicated Copy All button below.
------------------------------------------------- */

/* -------------------------------------------------
   OUTPUT BUILDERS (used by the Copy All button)
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
  try { applyHearingNameSwaps(await fetchHearingNameSwaps()); } catch (_) {}

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

// Find the site's fixed blue header bar by probing what is actually rendered
// at the top of the viewport (elementsFromPoint at a few x positions), then
// taking the TALLEST fixed/sticky, near-full-width element pinned to the top.
// Selector-based matching grabbed a slim inner wrapper (the search-bar row)
// whose top aligned with the blue bar, so the button height came out wrong.
let __topBarEl = null;
function findTopBar() {
  try {
    if (__topBarEl && document.contains(__topBarEl)) {
      const r = __topBarEl.getBoundingClientRect();
      if (r.top <= 2 && r.height >= 20 && r.height <= 140) return __topBarEl;
      __topBarEl = null;
    }
    const w = window.innerWidth;
    let best = null, bestH = 0;
    for (const x of [w * 0.3, w * 0.5, w * 0.7]) {
      for (const el of document.elementsFromPoint(x, 8)) {
        if (el === document.documentElement || el === document.body) continue;
        if (el.id === COPY_ALL_BTN_ID || el.id === '__lacourt_agenda_toast__') continue;
        const pos = getComputedStyle(el).position;
        if (pos !== 'fixed' && pos !== 'sticky') continue;
        const r = el.getBoundingClientRect();
        if (r.top > 2 || r.height < 20 || r.height > 140) continue;
        if (r.width < w * 0.6) continue;
        if (r.height > bestH) { best = el; bestH = r.height; }
      }
    }
    __topBarEl = best;
    return best;
  } catch (_) { return null; }
}

// Size the button to the full height of the top bar and align it flush with
// the bar (falls back to a default size 8px from the viewport top when no bar
// is found). Uses !important so site CSS can't shrink it.
let __barLoggedEl = null;
function positionCopyAllButton(btn) {
  const bar = findTopBar();
  const set = (prop, val) => { try { btn.style.setProperty(prop, val, 'important'); } catch (_) { btn.style[prop] = val; } };
  if (bar) {
    const r = bar.getBoundingClientRect();
    const h = Math.round(r.height);
    if (__barLoggedEl !== bar) { __barLoggedEl = bar; try { console.log('[LACourt-Agenda] top bar:', bar.tagName + '.' + (bar.className || ''), 'height=' + h); } catch (_) {} }
    set('top', Math.max(0, Math.round(r.top)) + 'px');
    set('height', h + 'px');
    set('line-height', h + 'px');
    set('padding', '0 20px');
    set('font-size', Math.max(14, Math.min(18, Math.round(h * 0.4))) + 'px');
    set('border-radius', '0');
    set('box-sizing', 'border-box');
    // No drop shadow when the button fills the header — it would fall onto the
    // slim grey bar directly below the blue header.
    set('box-shadow', 'none');
  } else {
    set('top', '8px');
    btn.style.removeProperty('height');
    set('line-height', '18px');
    set('padding', '4px 12px');
    set('font-size', '12px');
    set('border-radius', '5px');
    btn.style.removeProperty('box-shadow');
  }
}

function renderCopyAllButton() {
  const existing = document.getElementById(COPY_ALL_BTN_ID);
  if (existing) { positionCopyAllButton(existing); return; }
  if (!document.getElementById('day-table') || !document.body) return;
  const btn = document.createElement('button');
  btn.id = COPY_ALL_BTN_ID;
  btn.type = 'button';
  btn.textContent = 'Copy All';
  // Compact enough to sit inside the site's top bar.
  btn.style.cssText = 'position:fixed;top:8px;right:16px;z-index:2147483647;'
    + 'padding:4px 12px;border:none;border-radius:5px;cursor:pointer;'
    + 'background:#0a6e6e;color:#fff;font:600 12px system-ui,sans-serif;'
    + 'line-height:18px;box-shadow:0 1px 4px rgba(0,0,0,.3);';
  btn.addEventListener('mouseenter', () => { btn.style.background = '#0d8f8f'; });
  btn.addEventListener('mouseleave', () => { btn.style.background = '#0a6e6e'; });
  btn.addEventListener('click', () => copyAllAgenda(btn));
  document.body.appendChild(btn);
  positionCopyAllButton(btn);
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
  const clean = s => stripTrailingParenNumber(s.trim().replace(/\s+/g, ' '));
  const bTags = cell.querySelectorAll('b');
  if (bTags.length === 0) return clean(cell.textContent);
  return Array.from(bTags).map(b => clean(b.textContent)).join('\n');
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

// Match one excluded term against a hearing segment. A term wrapped in double
// quotes ("...") requires the whole segment to equal it exactly (trimmed);
// an unquoted term matches as a substring. Terms are already stored lowercased.
// KEEP IN SYNC with clipboard/content.js excludedTermMatches().
function excludedTermMatches(term, lower) {
  if (!term) return false;
  const quoted = term.match(/^"(.*)"$/);
  if (quoted) return lower.trim() === quoted[1].trim();
  return lower.includes(term);
}

function isExcluded(segment) {
  const lower = segment.trim().toLowerCase();
  return EXCLUDED_TERMS.some(term => excludedTermMatches(term, lower));
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
// Drop trailing number-only decorations: a purely-numeric parenthetical
// ("...defendant's aka (6861)") or a dash-number ("Motion to Compel - 3891").
// Repeats so stacked forms ("X (123) - 456") fully strip. Alphanumeric
// parentheticals like "(CCP 437c)" and worded dashes ("Demurrer - without
// Motion to Strike") are kept.
// KEEP IN SYNC with stripTrailingParenNumber in clipboard/content.js.
function stripTrailingParenNumber(s) {
  if (!s) return s;
  let prev;
  do {
    prev = s;
    s = s.replace(/\s*(?:\(\s*\d[\d\s.,\-]*\)|[-–—]\s*\d[\d\s.,]*)\s*$/, '').trim();
  } while (s !== prev);
  return s;
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

// Cell text with spaces preserved across markup line breaks. A hearing name
// that wraps in the Hearings tab renders as separate text nodes with NO
// whitespace between them (e.g. "judgment<br>to reflect"), so plain
// textContent glues words together ("judgmentto"). Join text nodes with a
// space instead, then collapse.
function cellTextSpaced(el) {
  const parts = [];
  const walk = n => {
    for (const c of n.childNodes) {
      if (c.nodeType === 3) { const t = c.nodeValue; if (t && t.trim()) parts.push(t.trim()); }
      else if (c.nodeType === 1) walk(c);
    }
  };
  walk(el);
  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

// Parse a fetched case Hearings-tab document into [{ name, dateTime, status }].
// Scans every row for a "MM/DD/YYYY HH:MM AM/PM" cell. Rows interleave EMPTY
// cells (observed live: ["", <name>, "", <date/time>, <status>, …]), so the
// hearing name is the nearest NON-EMPTY cell before the date, and the status is
// the first non-empty cell after it. Doesn't depend on exact column positions.
function parseCaseHearings(doc) {
  const out = [];
  for (const tr of doc.querySelectorAll('tr')) {
    const cells = Array.from(tr.children).map(cellTextSpaced);
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

// chrome.storage.session persists across the full page reload that navigating to
// the next agenda day triggers (and is shared across tabs), so a case's hearings
// fetched while prefetching the next day are still warm when that page loads.
function sessGet(key) {
  return new Promise(res => {
    try { chrome.storage.session.get([key], r => { if (chrome.runtime.lastError) return res(null); res((r && r[key]) || null); }); }
    catch (_) { res(null); }
  });
}
function sessSet(key, val) {
  try { chrome.storage.session.set({ [key]: val }, () => { void chrome.runtime.lastError; }); } catch (_) {}
}

async function loadCaseHearings(caseId) {
  const key = 'caseHearings:' + caseId;
  const cached = await sessGet(key);
  if (Array.isArray(cached)) return cached; // warmed by a prefetch or an earlier page
  const fetched = await fetchCaseHearings(caseId);
  if (fetched && fetched.length) sessSet(key, fetched); // don't persist empties/failures
  return fetched;
}

function getCaseHearings(caseId) {
  let v = CASE_HEARINGS_CACHE.get(caseId);
  if (v === undefined) {
    v = loadCaseHearings(caseId);
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

/* -------------------------------------------------
   NEXT-PAGE PREFETCH

   Users typically page forward day by day. Navigating to the next agenda day is
   a full reload, which would re-fetch every case's Hearings tab for name
   expansion. So while viewing the current day we prefetch the next day's agenda
   in the background and warm the (persistent) case-hearings cache for its cases,
   making the next page's expansion instant.
------------------------------------------------- */

function dayKeyNum(mdy) {
  const m = (mdy || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? (+m[3]) * 10000 + (+m[1]) * 100 + (+m[2]) : 0;
}
function pad2(n) { return n < 10 ? '0' + n : '' + n; }

// The URL of the next agenda day: prefer the page's own next-day nav link (so
// it respects the court's own day sequence), else +1 calendar day.
function nextAgendaUrl() {
  const curNum = dayKeyNum(agendaDay());
  if (!curNum) return null;
  let best = null, bestNum = Infinity;
  for (const a of document.querySelectorAll('a[href*="day="]')) {
    const href = a.getAttribute('href') || a.href || '';
    const m = href.match(/[?&]day=(\d{1,2}\/\d{1,2}\/\d{4})/);
    if (!m) continue;
    const n = dayKeyNum(m[1]);
    if (n > curNum && n < bestNum) { bestNum = n; best = a.href; } // a.href resolves to absolute
  }
  if (best) return best;
  try {
    const m = agendaDay().match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) return null;
    const d = new Date(+m[3], +m[1] - 1, +m[2] + 1);
    const url = new URL(location.href);
    url.searchParams.set('day', pad2(d.getMonth() + 1) + '/' + pad2(d.getDate()) + '/' + d.getFullYear());
    return url.toString();
  } catch (_) { return null; }
}

let __prefetchStarted = false;
async function prefetchNextAgenda() {
  if (__prefetchStarted) return;
  __prefetchStarted = true;
  const url = nextAgendaUrl();
  if (!url) return;
  // Only prefetch a given next-day once per session (across re-renders / tabs).
  const doneKey = 'agendaPrefetched:' + url;
  if (await sessGet(doneKey)) return;
  sessSet(doneKey, true);
  try {
    const res = await fetchWithTimeout(url, 10000);
    if (!res || !res.ok) return;
    const doc = new DOMParser().parseFromString(await res.text(), 'text/html');
    const table = doc.getElementById('day-table');
    if (!table) return;
    // Warm only cases with a truncated hearing (the ones expansion will fetch).
    const ids = new Set();
    for (const row of table.querySelectorAll('tr.js-row')) {
      let hasTrunc = false;
      for (const b of row.querySelectorAll('a[href*="/ecourt/ecms/agenda/event"] b')) {
        if (isTruncatedName((b.textContent || '').replace(/\s+/g, ' ').trim())) { hasTrunc = true; break; }
      }
      if (!hasTrunc) continue;
      const caseA = row.querySelector('td a[href*="/ecourt/ecms/case"]');
      const idm = (caseA ? (caseA.getAttribute('href') || caseA.href || '') : '').match(/[?&]id=(\d+)/);
      if (idm) ids.add(idm[1]);
    }
    await runWithConcurrency(Array.from(ids), 4, id => getCaseHearings(id));
    try { console.log('[LACourt-Agenda] prefetched next agenda', url, '—', ids.size, 'cases warmed'); } catch (_) {}
  } catch (_) { /* best-effort */ }
}

// Fetch full names for truncated hearings WITHOUT mutating the DOM. Returns an
// array of { b, full } swaps to apply, so the caller can batch the renames with
// the sort/colorize/float in a single reflow instead of swapping each name as
// its fetch resolves (which made the page jump repeatedly).
async function fetchHearingNameSwaps() {
  const table = document.getElementById('day-table');
  if (!table) return [];
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
  if (!jobs.length) return [];

  const swaps = [];
  await runWithConcurrency(jobs, 4, async job => {
    const hearings = await getCaseHearings(job.caseId);
    const full = stripTrailingParenNumber(fullNameForHearing(hearings, day, job.prefix));
    if (full && job.b.getAttribute(EXPANDED_ATTR) !== '1') swaps.push({ b: job.b, full, prefix: job.prefix });
  });
  return swaps;
}

// Apply the collected name swaps to the DOM (idempotent).
function applyHearingNameSwaps(swaps) {
  for (const s of swaps || []) {
    if (s.b.getAttribute(EXPANDED_ATTR) === '1') continue;
    s.b.textContent = s.full;
    s.b.setAttribute(EXPANDED_ATTR, '1');
    try { console.log('[LACourt-Agenda] expanded:', s.prefix + '…', '->', s.full); } catch (_) {}
  }
}

// True while we're applying our own batch of DOM changes, so the MutationObserver
// ignores the mutations we cause (renames, sort, reorder) and doesn't re-enter.
let __agendaBatching = false;

// Apply renames + native sort + colorize + green-float as ONE guarded batch, so
// the page reflows a single time to the final version instead of jumping once
// per async name expansion.
function applyAgendaChanges(swaps) {
  __agendaBatching = true;
  try {
    applyHearingNameSwaps(swaps);
    try { stripHearingLabelNumbers(); } catch (_) {}
    try { applyHearingDocsSort(); } catch (_) {}
    try { colorizeAgendaRows(); } catch (_) {}
    try { floatGreenRowsToTop(); } catch (_) {}
  } finally {
    // Release on the next frame: the observer callbacks queued by the mutations
    // above run as microtasks (before this rAF), so they see the guard still set
    // and skip; genuine later mutations resume normally.
    requestAnimationFrame(() => { __agendaBatching = false; });
  }
}

(function initHearingExpansion() {
  // The day-table renders in phases (empty → partial → full), and name expansion
  // needs async fetches. Debounce so we act only once the DOM has settled, and
  // serialize so a batch never overlaps another — otherwise we'd reorder once for
  // the initial rows and again after expansion (two visible jumps). A single
  // settled batch does renames + sort + colour + float together: one jump.
  let running = false;
  let timer = null;
  const SETTLE_MS = 300;

  const kick = () => {
    timer = null;
    if (running) { schedule(); return; } // a batch is in flight — retry after it
    running = true;
    // Fetch full names for any truncated hearings first (no DOM changes yet),
    // then apply the renames, sort, colouring, and green-float together.
    Promise.resolve(fetchHearingNameSwaps())
      .then(swaps => applyAgendaChanges(swaps))
      .catch(() => applyAgendaChanges([]))
      .finally(() => { running = false; });
  };
  const schedule = () => {
    if (__agendaBatching) return; // ignore the mutations our own batch causes
    if (timer) clearTimeout(timer);
    timer = setTimeout(kick, SETTLE_MS);
  };
  const start = () => {
    schedule();
    // Re-run when the day-table re-renders or paginates. Our own text swaps
    // no-op on the next pass (already-expanded / no longer truncated).
    const target = document.getElementById('day-table') || document.body;
    try { new MutationObserver(schedule).observe(target, { childList: true, subtree: true }); } catch (_) {}
    // Once the current day has had a head start, prefetch the next day in the
    // background so paging forward is instant.
    setTimeout(() => { try { prefetchNextAgenda(); } catch (_) {} }, 3000);
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

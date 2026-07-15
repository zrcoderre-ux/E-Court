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
   "Demurrer - without Motion..."). Each hearing <b> is wrapped in a per-hearing
   event link (/ecourt/ecms/agenda/event?dispatch=eventPage&id=NNNN). We fetch
   that event page for every truncated hearing, recover the full name, and drop
   it back into the <b> in place — so both the on-page display and the copy
   output carry the full name, and the exclusion check always runs against the
   full name rather than a truncated prefix. Fetches are cached by event id.
------------------------------------------------- */

const EVENT_NAME_CACHE = new Map(); // eventId -> full name (''=none) or in-flight Promise
const EXPANDED_ATTR = 'data-lac-expanded';

function isTruncatedName(text) { return /(?:\.\.\.|…)\s*$/.test(text); }
function truncatedPrefix(text) { return text.replace(/\s*(?:\.\.\.|…)\s*$/, '').trim(); }

function eventIdFromAnchor(a) {
  const href = (a && (a.getAttribute('href') || a.href)) || '';
  const m = href.match(/[?&]id=(\d+)/);
  return m ? m[1] : null;
}

// Pull the full hearing name out of a fetched event page: entries render as
// "MM/DD/YYYY <name>" in label/td/span cells; return the one that starts with
// the truncated prefix (so we pick the right hearing, not a related deadline).
function fullNameFromEventDoc(doc, prefix) {
  const p = prefix.toLowerCase();
  const seen = new Set();
  for (const el of doc.querySelectorAll('label, td, span')) {
    let t = (el.textContent || '').replace(/\s+/g, ' ').trim();
    if (!t) continue;
    t = t.replace(/^\d{1,2}\/\d{1,2}\/\d{4}\s+/, '').trim();
    if (!t || seen.has(t)) continue;
    seen.add(t);
    if (t.length > prefix.length && t.toLowerCase().startsWith(p)) return t;
  }
  return '';
}

async function fetchEventFullName(eventId, prefix) {
  try {
    const res = await fetch('/ecourt/ecms/agenda/event?dispatch=eventPage&id=' + eventId, { credentials: 'include' });
    if (!res || !res.ok) return '';
    const doc = new DOMParser().parseFromString(await res.text(), 'text/html');
    return fullNameFromEventDoc(doc, prefix);
  } catch (_) { return ''; }
}

async function expandTruncatedHearings() {
  const table = document.getElementById('day-table');
  if (!table) return;
  const anchors = table.querySelectorAll('tr.js-row a[href*="/ecourt/ecms/agenda/event"]');
  for (const a of anchors) {
    const b = a.querySelector('b');
    if (!b || b.getAttribute(EXPANDED_ATTR) === '1') continue;
    const text = (b.textContent || '').replace(/\s+/g, ' ').trim();
    if (!isTruncatedName(text)) continue;
    // Expand every truncated hearing so the exclusion check (during copy) always
    // runs against the full name, not the truncated prefix.
    const eventId = eventIdFromAnchor(a);
    if (!eventId) continue;
    const prefix = truncatedPrefix(text);

    let full = EVENT_NAME_CACHE.get(eventId);
    if (full === undefined) {
      const promise = fetchEventFullName(eventId, prefix);
      EVENT_NAME_CACHE.set(eventId, promise);
      full = await promise;
      EVENT_NAME_CACHE.set(eventId, full || '');
    } else if (full && typeof full.then === 'function') {
      full = await full;
    }
    if (full) {
      b.textContent = full;
      b.setAttribute(EXPANDED_ATTR, '1');
      try { console.log('[LACourt-Agenda] expanded:', prefix + '…', '->', full); } catch (_) {}
    }
  }
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

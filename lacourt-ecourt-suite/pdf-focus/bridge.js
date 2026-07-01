/**
 * LA Court PDF Focus - Bridge Script (ISOLATED world)
 *
 * 1. Forwards LACOURT_OPEN_DOC events from content.js to the background
 *    worker so PDFs open as background tabs.
 * 2. Tracks clicked doc-view buttons (by docId) and stamps a green
 *    checkmark just to the left of each.
 *
 * Marks live in sessionStorage so they survive AJAX pagination and
 * reloads, but clear when the tab closes.
 */

const DEBUG = true;
const log = (...args) => { if (DEBUG) console.log('[LACourtPDF]', ...args); };

// ============================================================
// Forward window.open intercepts to the background worker
// ============================================================

window.addEventListener('LACOURT_OPEN_DOC', (event) => {
  const url = event.detail && event.detail.url;
  if (!url) return;
  log('LACOURT_OPEN_DOC received, url=', url);
  chrome.runtime.sendMessage({ type: 'OPEN_DOC_BACKGROUND', url: url });

  const docId = extractDocId(url);
  if (docId) {
    rememberOpened(docId);
    scanAndStampAll();
  }
});

// ============================================================
// Session storage of opened doc IDs
// ============================================================

const STORAGE_KEY = 'lacourtPdfFocus.openedDocIds';

function getOpenedSet() {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY);
    if (!raw) return new Set();
    return new Set(JSON.parse(raw));
  } catch (e) {
    return new Set();
  }
}

function rememberOpened(docId) {
  const set = getOpenedSet();
  set.add(docId);
  try {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(Array.from(set)));
    log('remembered docId', docId, '— total:', set.size);
  } catch (e) {
    log('sessionStorage write failed:', e);
  }
}

// ============================================================
// docId extraction
// ============================================================

function extractDocId(str) {
  if (!str) return null;
  const m = String(str).match(/docId=(\d+)/);
  return m ? m[1] : null;
}

function getDocIdForAnchor(anchor) {
  if (!anchor) return null;
  const onclick = anchor.getAttribute('onclick') || '';
  return extractDocId(onclick);
}

// ============================================================
// Checkmark rendering (inline, as a sibling of the button)
// ============================================================

const MARK_CLASS = 'lacourt-pdf-checkmark';
const MARK_ATTR = 'data-lacourt-pdf-mark-for';

function injectStyles() {
  if (document.getElementById('lacourt-pdf-focus-styles')) return;
  const style = document.createElement('style');
  style.id = 'lacourt-pdf-focus-styles';
  style.textContent = `
    .${MARK_CLASS} {
      display: inline-block !important;
      width: 12px !important;
      height: 12px !important;
      border-radius: 50% !important;
      background: #22c55e !important;
      color: #ffffff !important;
      font-size: 9px !important;
      font-weight: bold !important;
      line-height: 12px !important;
      text-align: center !important;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
      box-shadow: 0 0 2px rgba(0, 0, 0, 0.5) !important;
      pointer-events: none !important;
      margin-right: 4px !important;
      padding: 0 !important;
      vertical-align: middle !important;
    }
  `;
  (document.head || document.documentElement).appendChild(style);
}

function stampCheckmark(anchor, docId) {
  if (!anchor || !docId) return;

  // Skip if this anchor already has a mark adjacent to it.
  if (anchor.previousElementSibling &&
      anchor.previousElementSibling.classList &&
      anchor.previousElementSibling.classList.contains(MARK_CLASS)) {
    return;
  }

  const mark = document.createElement('span');
  mark.className = MARK_CLASS;
  mark.setAttribute(MARK_ATTR, docId);
  mark.textContent = '\u2713';

  // Insert as the immediate previous sibling of the button.
  // This way the mark lives in the same table cell/row and disappears
  // automatically when the row is removed by pagination or filtering.
  anchor.parentNode.insertBefore(mark, anchor);
  log('inline checkmark inserted for docId=', docId);
}

function cssEscape(value) {
  if (window.CSS && window.CSS.escape) return window.CSS.escape(value);
  return String(value).replace(/[^a-zA-Z0-9_-]/g, '\\$&');
}

// ============================================================
// Scanning
// ============================================================

function findDocAnchors() {
  // Try several selector patterns since e-court markup may vary.
  const selectors = [
    'a[onclick*="openInNewWindow"]',
    'a[onclick*="ecms/doc"]',
    'a[onclick*="docId"]'
  ];
  const found = new Set();
  selectors.forEach(sel => {
    document.querySelectorAll(sel).forEach(a => found.add(a));
  });
  return Array.from(found);
}

function scanAndStampAll() {
  const opened = getOpenedSet();
  injectStyles();

  const anchors = findDocAnchors();
  log('scan: found', anchors.length, 'doc anchors, opened set size=', opened.size);

  let stamped = 0;
  anchors.forEach(a => {
    const docId = getDocIdForAnchor(a);
    if (docId && opened.has(docId)) {
      stampCheckmark(a, docId);
      stamped++;
    }
  });
  log('scan: stamped', stamped, 'checkmarks');

  removeOrphanMarks(opened);
}

function removeOrphanMarks(openedSet) {
  // For inline marks: a mark is orphaned if (a) its associated anchor
  // is no longer in the DOM, or (b) its docId is no longer in the
  // opened set. Both cases get the mark removed.
  document.querySelectorAll(`.${MARK_CLASS}`).forEach(mark => {
    const id = mark.getAttribute(MARK_ATTR);
    if (!openedSet.has(id)) {
      mark.remove();
      return;
    }
    // Verify the anchor it sits next to still exists and matches.
    const next = mark.nextElementSibling;
    if (!next || getDocIdForAnchor(next) !== id) {
      mark.remove();
    }
  });
}

// ============================================================
// Triggers
// ============================================================

document.addEventListener('click', (e) => {
  const anchor = e.target && e.target.closest && e.target.closest(
    'a[onclick*="openInNewWindow"], a[onclick*="ecms/doc"], a[onclick*="docId"]'
  );
  if (!anchor) return;
  const docId = getDocIdForAnchor(anchor);
  log('click captured, docId=', docId, 'anchor=', anchor);
  if (!docId) return;
  rememberOpened(docId);
  setTimeout(scanAndStampAll, 0);
}, true);

function ready(fn) {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', fn, { once: true });
  } else {
    fn();
  }
}

ready(() => {
  log('ready, initial scan');
  injectStyles();
  scanAndStampAll();

  let pending = false;
  const observer = new MutationObserver((mutations) => {
    // Ignore mutations caused by our own checkmark inserts/removes.
    const meaningful = mutations.some(m => {
      const isOurNode = (n) => n && n.classList && n.classList.contains(MARK_CLASS);
      const addedReal = Array.from(m.addedNodes || []).some(n => !isOurNode(n));
      const removedReal = Array.from(m.removedNodes || []).some(n => !isOurNode(n));
      return addedReal || removedReal;
    });
    if (!meaningful) return;
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      scanAndStampAll();
    });
  });
  observer.observe(document.body, { childList: true, subtree: true });

  log('observers attached');
});

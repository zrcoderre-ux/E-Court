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

  const outputRows = [];

  candidateRows.forEach(row => {
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

  if (outputRows.length === 0) {
    e.clipboardData.setData('text/plain', '');
    e.clipboardData.setData('text/html', '');
    e.preventDefault();
    return;
  }

  // Plain text: hearing tab case, blank row after each
  const plainLines = outputRows.flatMap(r => [r.col5 + '\t' + r.col6, '']);
  const plainText  = plainLines.join('\n');

  // HTML with Times New Roman 22pt
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

  e.clipboardData.setData('text/plain', plainText);
  e.clipboardData.setData('text/html', html);
  e.preventDefault();
}

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
   UTILITIES
------------------------------------------------- */

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

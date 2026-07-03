/**
 * LA Court Cleaner - Options page
 *
 * Manages two user-editable exclusion lists, both stored in chrome.storage.sync:
 *
 *   1. excludedTerms                 — agenda-cleaner hearing-type filter
 *                                       (used by the Agenda Cleaner extension)
 *   2. dismissedMotionExclusions     — motion types for which dismissed
 *                                       parties should be dropped from the
 *                                       Ctrl+A rotation and Fill Microsoft
 *                                       Form output. Substring match,
 *                                       case-insensitive.
 *
 * The two lists share a render/edit pattern, factored into makeListController.
 */

const DEFAULT_EXCLUDED_TERMS = [
  'conference',
  'non-appearance case revie',
  'non-jury trial',
  'order to show cause re: d',
  'ex parte',
  'motion to deem request fo',
  'application for order for',
  'jury trial',
  'post-arbitration status c',
  'motion to compel discover',
  'post-mediation status con',
  'order to show cause re: s',
  'informal discovery confer',
  'motion to be relieved as',
];

// IMPORTANT: keep this in sync with the same constant in content.js
// (DEFAULT_DISMISSED_MOTION_EXCLUSIONS). content.js uses these as its
// fallback when no saved list exists yet.
const DEFAULT_DISMISSED_MOTION_EXCLUSIONS = [
  // Dispositive / merits motions
  'summary judgment',
  'summary adjudication',
  'judgment on the pleadings',
  'directed verdict',
  'nonsuit',
  'new trial',
  'vacate judgment',
  'set aside default',
  // Pleading challenges
  'demurrer',
  'motion to strike',
  'anti-slapp',
  'special motion to strike',
  'leave to amend',
  'leave to file cross-complaint',
  // Discovery
  'motion to compel',
  'protective order',
  'motion to quash',
  'trial preference',
  'motion in limine',
  'bifurcate',
  'consolidate',
  'sever',
  'coordinate',
  // Class / representative actions
  'class certification',
  'decertify',
  // Equitable relief
  'preliminary injunction',
  'temporary restraining order',
  'writ of attachment',
  // Service / jurisdiction
  'quash service',
  'order to show cause re contempt',
];

/**
 * Builds a controller for one editable list (agenda terms, or dismissed
 * motion exclusions). Wires up the corresponding DOM elements by ID.
 */
function makeListController(opts) {
  const {
    storageKey,
    defaults,
    listElId,
    emptyNoteElId,
    addBtnId,
    newInputId,
    saveBtnId,
    resetBtnId,
    saveStatusElId,
    resetConfirmMessage,
  } = opts;

  let terms = [];

  function renderList() {
    const ul = document.getElementById(listElId);
    const emptyNote = document.getElementById(emptyNoteElId);
    ul.innerHTML = '';

    if (terms.length === 0) {
      emptyNote.style.display = 'block';
      return;
    }
    emptyNote.style.display = 'none';

    terms.forEach((term, i) => {
      const li = document.createElement('li');

      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'term-text';
      input.value = term;
      input.setAttribute('aria-label', 'Edit term');
      input.addEventListener('change', () => {
        const val = input.value.trim().toLowerCase();
        if (val) {
          terms[i] = val;
        } else {
          terms.splice(i, 1);
          renderList();
        }
      });
      input.addEventListener('keydown', e => {
        if (e.key === 'Enter') input.blur();
      });

      const delBtn = document.createElement('button');
      delBtn.className = 'btn-delete';
      delBtn.textContent = '×';
      delBtn.title = 'Remove this term';
      delBtn.addEventListener('click', () => {
        terms.splice(i, 1);
        renderList();
      });

      li.appendChild(input);
      li.appendChild(delBtn);
      ul.appendChild(li);
    });
  }

  function showSaved() {
    const status = document.getElementById(saveStatusElId);
    if (!status) return;
    status.classList.add('visible');
    setTimeout(() => status.classList.remove('visible'), 2000);
  }

  document.getElementById(addBtnId).addEventListener('click', () => {
    const input = document.getElementById(newInputId);
    const val = input.value.trim().toLowerCase();
    if (val && !terms.includes(val)) {
      terms.push(val);
      renderList();
      input.value = '';
    }
    input.focus();
  });

  document.getElementById(newInputId).addEventListener('keydown', e => {
    if (e.key === 'Enter') document.getElementById(addBtnId).click();
  });

  document.getElementById(saveBtnId).addEventListener('click', () => {
    chrome.storage.sync.set({ [storageKey]: terms }, () => {
      showSaved();
    });
  });

  document.getElementById(resetBtnId).addEventListener('click', () => {
    if (confirm(resetConfirmMessage)) {
      terms = [...defaults];
      renderList();
      chrome.storage.sync.set({ [storageKey]: terms }, showSaved);
    }
  });

  // Initial load — fall back to defaults if nothing has been saved yet.
  chrome.storage.sync.get([storageKey], result => {
    const saved = result && result[storageKey];
    if (Array.isArray(saved)) {
      terms = saved;
    } else {
      terms = [...defaults];
    }
    renderList();
  });
}

// Card 1: agenda-cleaner hearing-type exclusions
makeListController({
  storageKey: 'excludedTerms',
  defaults: DEFAULT_EXCLUDED_TERMS,
  listElId: 'termList',
  emptyNoteElId: 'emptyNote',
  addBtnId: 'addBtn',
  newInputId: 'newTermInput',
  saveBtnId: 'saveBtn',
  resetBtnId: 'resetBtn',
  saveStatusElId: 'saveStatus',
  resetConfirmMessage: 'Reset to the default agenda exclusion list? This will discard any custom changes.',
});

// Card 2: dismissed-party motion exclusions
makeListController({
  storageKey: 'dismissedMotionExclusions',
  defaults: DEFAULT_DISMISSED_MOTION_EXCLUSIONS,
  listElId: 'dismissedList',
  emptyNoteElId: 'dismissedEmptyNote',
  addBtnId: 'dismissedAddBtn',
  newInputId: 'dismissedNewInput',
  saveBtnId: 'dismissedSaveBtn',
  resetBtnId: 'dismissedResetBtn',
  saveStatusElId: 'dismissedSaveStatus',
  resetConfirmMessage: 'Reset to the default dismissed-party motion exclusion list? This will discard any custom changes.',
});

/* ------------------------------------------------------------------ */
/* Card 3: Documents-button debug tracking (sortable table + CSV)      */
/* ------------------------------------------------------------------ */

// Flattens the opened-document tracking into one row per document:
//   { caseNumber, name, source, downloaded, dlKnown, buttonMissed, at }
// dlKnown is true once the case has >=1 recorded download, which is what makes
// "not downloaded" (over-inclusion) meaningful. buttonMissed flags docs opened
// only manually — candidate under-inclusions the button should have opened.
function loadDocTrackingRows(cb) {
  chrome.storage.local.get(['docTracking'], result => {
    const t = (result && result.docTracking) || {};
    const opened = t.opened || {};
    const downloaded = t.downloaded || {};

    const dlByCase = {};
    Object.keys(opened).forEach(id => {
      const cn = (opened[id] || {}).caseNumber || '(unknown case)';
      if (!(cn in dlByCase)) dlByCase[cn] = 0;
      if (downloaded[id]) dlByCase[cn]++;
    });

    const rows = Object.keys(opened).map(id => {
      const o = opened[id] || {};
      const cn = o.caseNumber || '(unknown case)';
      return {
        caseNumber: cn,
        name: o.name || '(unnamed)',
        source: (o.button && o.manual) ? 'button+manual' : (o.button ? 'button' : 'manual'),
        downloaded: !!downloaded[id],
        dlKnown: dlByCase[cn] > 0,
        buttonMissed: !!(o.manual && !o.button),
        at: o.at || 0,
      };
    });
    cb(rows);
  });
}

function dtEscHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function dtWhen(at) { return at ? new Date(at).toLocaleString() : ''; }

// Rank downloaded state so the column sorts sensibly: 0 unknown, 1 not, 2 yes.
function dtDlRank(r) { return !r.dlKnown ? 0 : (r.downloaded ? 2 : 1); }

let dtSortKey = 'caseNumber';
let dtSortDir = 1;

function dtSortRows(rows) {
  const dir = dtSortDir;
  return rows.slice().sort((a, b) => {
    let av, bv;
    if (dtSortKey === 'name') { av = a.name.toLowerCase(); bv = b.name.toLowerCase(); }
    else if (dtSortKey === 'source') { av = a.source; bv = b.source; }
    else if (dtSortKey === 'downloaded') { av = dtDlRank(a); bv = dtDlRank(b); }
    else if (dtSortKey === 'at') { av = a.at; bv = b.at; }
    else { av = a.caseNumber; bv = b.caseNumber; }
    if (av < bv) return -dir;
    if (av > bv) return dir;
    return b.at - a.at; // stable tiebreak: newest first
  });
}

function renderDocTracking() {
  const view = document.getElementById('docTrackingView');
  if (!view) return;
  loadDocTrackingRows(rows => {
    if (!rows.length) {
      view.innerHTML = '<p class="empty-note">No documents opened yet.</p>';
      return;
    }
    const over = rows.filter(r => r.dlKnown && !r.downloaded).length;
    const under = rows.filter(r => r.buttonMissed).length;
    const arrow = k => dtSortKey === k ? (dtSortDir > 0 ? ' ▲' : ' ▼') : '';
    const cols = [
      ['caseNumber', 'Case'], ['name', 'Document'], ['source', 'Source'],
      ['downloaded', 'Downloaded'], ['at', 'Opened'],
    ];

    let html = '<div class="dt-summary">' + rows.length + ' documents · ' +
      over + ' opened-not-downloaded (over) · ' + under + ' button-missed (under)</div>' +
      '<div class="dt-tablewrap"><table class="dt-table"><thead><tr>' +
      cols.map(c => `<th data-sort="${c[0]}">${c[1]}${arrow(c[0])}</th>`).join('') +
      '</tr></thead><tbody>';

    dtSortRows(rows).forEach(r => {
      const dlCell = !r.dlKnown ? '<span class="dt-muted">—</span>'
        : (r.downloaded ? '<span class="dt-yes">✓ yes</span>' : '<span class="dt-no">✗ no</span>');
      const cls = [];
      if (r.dlKnown && !r.downloaded) cls.push('dt-row-over');
      if (r.buttonMissed) cls.push('dt-row-under');
      html += `<tr class="${cls.join(' ')}">` +
        `<td>${dtEscHtml(r.caseNumber)}</td>` +
        `<td>${dtEscHtml(r.name)}</td>` +
        `<td class="dt-src-cell">${r.source}${r.buttonMissed ? ' <span class="dt-missed">button missed</span>' : ''}</td>` +
        `<td>${dlCell}</td>` +
        `<td class="dt-when">${dtEscHtml(dtWhen(r.at))}</td>` +
        '</tr>';
    });
    html += '</tbody></table></div>';
    view.innerHTML = html;

    view.querySelectorAll('th[data-sort]').forEach(th => {
      th.addEventListener('click', () => {
        const k = th.getAttribute('data-sort');
        if (dtSortKey === k) dtSortDir = -dtSortDir;
        else { dtSortKey = k; dtSortDir = 1; }
        renderDocTracking();
      });
    });
  });
}

// One row per document, sorted by case then newest-first, quoted for CSV.
function dtCsvCell(s) {
  s = String(s == null ? '' : s);
  return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function docTrackingCsv(cb) {
  loadDocTrackingRows(rows => {
    const header = ['Case', 'Document', 'Source', 'Downloaded', 'ButtonMissed', 'OpenedAt'];
    const lines = [header.map(dtCsvCell).join(',')];
    rows.slice()
      .sort((a, b) => (a.caseNumber < b.caseNumber ? -1 : a.caseNumber > b.caseNumber ? 1 : b.at - a.at))
      .forEach(r => {
        const dl = !r.dlKnown ? 'unknown' : (r.downloaded ? 'yes' : 'no');
        lines.push([
          r.caseNumber, r.name, r.source, dl, r.buttonMissed ? 'yes' : '', dtWhen(r.at),
        ].map(dtCsvCell).join(','));
      });
    cb(lines.join('\r\n'));
  });
}

function docTrackingStatus(msg) {
  const s = document.getElementById('docTrackingStatus');
  if (!s) return;
  s.textContent = msg || '✓ Done';
  s.classList.add('visible');
  setTimeout(() => s.classList.remove('visible'), 2000);
}

(function initDocTracking() {
  const refresh = document.getElementById('docTrackingRefresh');
  const downloadBtn = document.getElementById('docTrackingDownload');
  const copyBtn = document.getElementById('docTrackingCopy');
  const clearBtn = document.getElementById('docTrackingClear');
  if (!refresh) return; // card not present

  refresh.addEventListener('click', () => { renderDocTracking(); docTrackingStatus('✓ Refreshed'); });

  downloadBtn.addEventListener('click', () => {
    docTrackingCsv(csv => {
      const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'documents-debug-tracking.csv';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      docTrackingStatus('✓ CSV downloaded');
    });
  });

  copyBtn.addEventListener('click', () => {
    docTrackingCsv(csv => {
      navigator.clipboard.writeText(csv).then(
        () => docTrackingStatus('✓ CSV copied'),
        () => docTrackingStatus('Copy failed')
      );
    });
  });

  clearBtn.addEventListener('click', () => {
    if (!confirm('Clear all Documents-button tracking data?')) return;
    chrome.storage.local.remove('docTracking', () => { renderDocTracking(); docTrackingStatus('✓ Cleared'); });
  });

  renderDocTracking();
})();

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
/* Card 3: Documents-button debug tracking                             */
/* ------------------------------------------------------------------ */

// Groups the opened-document tracking by case. Each entry is
// { docId, name, at, downloaded, button, manual }. Cases with >=1 download flag
// the docs that were opened but not downloaded (candidate over-inclusions).
// Docs opened only manually (button === false) are candidate under-inclusions
// the Documents button should have opened but didn't.
function loadDocTrackingGrouped(cb) {
  chrome.storage.local.get(['docTracking'], result => {
    const t = (result && result.docTracking) || {};
    const opened = t.opened || {};
    const downloaded = t.downloaded || {};
    const byCase = {};
    Object.keys(opened).forEach(docId => {
      const o = opened[docId] || {};
      const cn = o.caseNumber || '(unknown case)';
      (byCase[cn] = byCase[cn] || []).push({
        docId,
        name: o.name || '(unnamed)',
        at: o.at || 0,
        downloaded: !!downloaded[docId],
        button: !!o.button,
        manual: !!o.manual,
      });
    });
    cb(byCase);
  });
}

function renderDocTracking() {
  const view = document.getElementById('docTrackingView');
  if (!view) return;
  loadDocTrackingGrouped(byCase => {
    const cases = Object.keys(byCase);
    if (!cases.length) {
      view.innerHTML = '<p class="empty-note">No documents opened yet.</p>';
      return;
    }
    const lastAt = cn => Math.max.apply(null, byCase[cn].map(d => d.at));
    cases.sort((a, b) => lastAt(b) - lastAt(a));

    view.innerHTML = '';
    cases.forEach(cn => {
      const docs = byCase[cn].slice().sort((a, b) => b.at - a.at);
      const dlCount = docs.filter(d => d.downloaded).length;
      const manualOnly = docs.filter(d => d.manual && !d.button).length;

      const wrap = document.createElement('div');
      wrap.className = 'dt-case';

      const title = document.createElement('div');
      title.className = 'dt-case-title';
      title.textContent = cn;
      const sub = document.createElement('span');
      sub.className = 'dt-sub';
      sub.textContent = '  — ' + docs.length + ' opened' +
        (dlCount > 0 ? (', ' + (docs.length - dlCount) + ' not downloaded') : ' (no downloads recorded)') +
        (manualOnly > 0 ? (', ' + manualOnly + ' button missed') : '');
      title.appendChild(sub);
      wrap.appendChild(title);

      const ul = document.createElement('ul');
      docs.forEach(d => {
        const li = document.createElement('li');
        li.textContent = d.name;
        // Over-inclusion: the button opened it but you didn't download it.
        if (dlCount > 0 && !d.downloaded) li.classList.add('dt-notdl');
        else if (dlCount > 0 && d.downloaded) li.classList.add('dt-dl');
        // Under-inclusion: you opened it yourself; the button never did.
        if (d.manual && !d.button) li.classList.add('dt-manual');

        const tag = document.createElement('span');
        tag.className = 'dt-src';
        tag.textContent = d.button && d.manual ? ' [button + manual]'
          : d.button ? ' [button]'
          : ' [manual]';
        li.appendChild(tag);

        ul.appendChild(li);
      });
      wrap.appendChild(ul);
      view.appendChild(wrap);
    });
  });
}

function docTrackingAsText(cb) {
  loadDocTrackingGrouped(byCase => {
    const lines = [];
    Object.keys(byCase).forEach(cn => {
      const docs = byCase[cn].slice().sort((a, b) => b.at - a.at);
      const dlCount = docs.filter(d => d.downloaded).length;
      const manualOnly = docs.filter(d => d.manual && !d.button).length;
      lines.push(cn + '  (' + docs.length + ' opened' +
        (dlCount > 0 ? (', ' + (docs.length - dlCount) + ' not downloaded') : ', no downloads recorded') +
        (manualOnly > 0 ? (', ' + manualOnly + ' button missed') : '') + ')');
      docs.forEach(d => {
        const mark = dlCount > 0 ? (d.downloaded ? '[downloaded] ' : '[NOT downloaded] ') : '';
        const src = d.button && d.manual ? '[button+manual] '
          : d.button ? '[button] '
          : '[manual] ';
        lines.push('  ' + mark + src + d.name);
      });
      lines.push('');
    });
    cb(lines.join('\n').trim() || '(no data)');
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
  const copyBtn = document.getElementById('docTrackingCopy');
  const clearBtn = document.getElementById('docTrackingClear');
  if (!refresh) return; // card not present

  refresh.addEventListener('click', () => { renderDocTracking(); docTrackingStatus('✓ Refreshed'); });
  copyBtn.addEventListener('click', () => {
    docTrackingAsText(text => {
      navigator.clipboard.writeText(text).then(
        () => docTrackingStatus('✓ Copied'),
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

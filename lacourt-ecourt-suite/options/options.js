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

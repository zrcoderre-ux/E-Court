/**
 * LA Court Clipboard Cleaner - Paste Rotator (v2.5)
 *
 * Runs on all URLs except the case page itself. Listens for paste events.
 * When the pasted text matches the sentinel set by the case-page content
 * script, intercepts the paste and inserts the next entry from the stored
 * rotation sequence instead, then increments the index. The next paste gets
 * the next entry, and so on.
 *
 * Each paste reads the latest sequence/index from chrome.storage.local, so
 * multiple tabs/forms stay in sync.
 *
 * Wrapped in an IIFE so its identifiers don't collide with content.js when
 * both happen to be injected into the same isolated world.
 */
(function () {
  'use strict';

  const SENTINEL = '\u26A1LACOURT_PARTY_ROTATION\u26A1';

  // Per-frame stack of pastes performed by this rotator. Each entry holds
  // a snapshot of the target element's state before our insertion, so a
  // subsequent Ctrl+Z can restore that snapshot and rewind the rotation.
  // Shape: [{ element, kind: 'input'|'editable', beforeValue, afterValue, ts }]
  const pasteStack = [];

  document.addEventListener('paste', function (e) {
    const pasted = (e.clipboardData && e.clipboardData.getData('text/plain')) || '';
    if (pasted.trim() !== SENTINEL.trim()) return; // Normal paste — ignore.

    // Walk up from e.target (which may be a text node or inner span) to find
    // the actual editable element. Microsoft Forms uses contenteditable divs
    // whose inner spans are sometimes the paste target.
    const editable = findEditableHost(e.target);
    if (!editable) {
      console.log('[LACourt-rotator] no editable host for paste target:', e.target);
      return;
    }

    // Block the sentinel from actually landing in the field. Must happen
    // synchronously before the async messaging round-trip.
    e.preventDefault();
    e.stopImmediatePropagation();

    // First, fetch the full rotation. If the form has 2+ labeled fields we
    // recognize, auto-fill all of them at once and exhaust the rotation.
    // Otherwise, fall back to single-field rotation paste.
    try {
      chrome.runtime.sendMessage({ type: 'getRotation' }, response => {
        if (chrome.runtime.lastError) {
          console.log('[LACourt-rotator] getRotation error:', chrome.runtime.lastError.message);
          return;
        }
        const rot = response && response.rotation;
        if (!rot || !Array.isArray(rot.sequence) || rot.sequence.length === 0) {
          console.log('[LACourt-rotator] no rotation in storage');
          return;
        }

        const matches = findLabeledFields(rot.labeled || {});
        console.log('[LACourt-rotator] labeled-field matches found:', matches.length);

        if (matches.length >= 2) {
          // Auto-fill mode: clear all fields first, then fill every matched field,
          // push each to the undo stack so Ctrl+Z still works, then mark the rotation as exhausted.
          for (const m of matches) {
            clearField(m.element);
          }
          
          for (const m of matches) {
            const snap = insertText(m.element, m.value);
            if (snap) {
              snap.autoFillKey = m.key;
              pasteStack.push(snap);
            }
          }

          // Post-fill: Microsoft Forms' Fluent date picker is built on a
          // combobox, not a plain text input. Its parser ignores direct
          // .value writes (which is why the previous insertText path filled
          // every field except the date). Run a Fluent-aware commit pass
          // that, per element, decides whether the standard blur is enough
          // or whether we need to simulate actual typing.
          for (const m of matches) {
            commitFieldChange(m.element, m.value);
          }

          console.log('[LACourt-rotator] auto-filled', matches.length, 'fields:',
            matches.map(m => m.key));

          try {
            chrome.runtime.sendMessage({ type: 'exhaustRotation' }, () => {});
          } catch (_) {}
          return;
        }

        // Auto-fill threshold not met. Dump diagnostic info so we can see
        // what labels are actually present on the page and adjust matching.
        console.log('[LACourt-rotator] auto-fill threshold not met — dumping fields:');
        const allEditable = document.querySelectorAll(
          'input, textarea, [contenteditable="true"], [contenteditable=""]'
        );
        const dump = [];
        for (const el of allEditable) {
          if (el.tagName.toLowerCase() === 'input') {
            const type = (el.getAttribute('type') || 'text').toLowerCase();
            if (!['text', 'search', 'email', 'tel', 'url', 'date', ''].includes(type)) continue;
          }
          dump.push({
            tag: el.tagName.toLowerCase(),
            type: el.getAttribute('type') || null,
            label: findLabelForElement(el),
            ariaLabel: el.getAttribute('aria-label') || null,
            ariaLabelledBy: el.getAttribute('aria-labelledby') || null,
            placeholder: el.getAttribute('placeholder') || null,
            id: el.id || null,
          });
        }
        console.log('[LACourt-rotator] editable field dump:', dump);
        console.log('[LACourt-rotator] available labeled keys:', Object.keys(rot.labeled || {}));

        // Fall back to single-field rotation: advance index and insert into
        // the focused field.
        chrome.runtime.sendMessage({ type: 'advanceRotation' }, advResp => {
          if (chrome.runtime.lastError) {
            console.log('[LACourt-rotator] advanceRotation error:',
              chrome.runtime.lastError.message);
            return;
          }
          if (!advResp || advResp.value == null) {
            console.log('[LACourt-rotator] no rotation value (exhausted or empty)');
            return;
          }
          console.log('[LACourt-rotator] inserting paste #' +
            advResp.rotation.index + ':', advResp.value);
          const snap = insertText(editable, advResp.value);
          if (snap) pasteStack.push(snap);
        });
      });
    } catch (err) {
      console.log('[LACourt-rotator] paste handler threw:', err);
    }
  }, true); // capture phase: run before the page's own paste handlers

  /**
   * Listens for Ctrl+Z / Cmd+Z. If we have any tracked pastes, undo the most
   * recent one (restoring the field's prior state) and rewind the rotation
   * index by 1 so the next paste re-issues that value.
   *
   * Forms doesn't include our programmatic insertions in its own undo stack,
   * so we handle the entire undo ourselves.
   */
  document.addEventListener('keydown', function (e) {
    const isUndo = (e.ctrlKey || e.metaKey) && !e.shiftKey && (e.key === 'z' || e.key === 'Z');
    if (!isUndo) return;
    if (pasteStack.length === 0) return; // nothing of ours to undo

    const last = pasteStack.pop();
    if (!restoreSnapshot(last)) {
      // Restoration failed (the field state has changed since our paste);
      // put the entry back and let the browser handle Ctrl+Z normally.
      pasteStack.push(last);
      return;
    }

    e.preventDefault();
    e.stopImmediatePropagation();

    console.log('[LACourt-rotator] undo: rewinding rotation');
    try {
      chrome.runtime.sendMessage({ type: 'rewindRotation' }, () => {
        if (chrome.runtime.lastError) {
          console.log('[LACourt-rotator] rewind sendMessage error:',
            chrome.runtime.lastError.message);
        }
      });
    } catch (err) {
      console.log('[LACourt-rotator] rewind sendMessage threw:', err);
    }
  }, true);

/**
 * Walks up from any node to find the nearest input/textarea/contenteditable
 * host. Returns null if none found.
 */
function findEditableHost(node) {
  let el = node;
  // If we got a text node, jump to its parent element.
  if (el && el.nodeType === Node.TEXT_NODE) el = el.parentElement;

  for (let i = 0; el && i < 20; i++) {
    if (el.nodeType === Node.ELEMENT_NODE) {
      const tag = (el.tagName || '').toLowerCase();
      if (tag === 'input' || tag === 'textarea') return el;
      if (el.isContentEditable) return el;
    }
    el = el.parentElement;
  }
  return null;
}

/**
 * Mapping of labeled-data keys to label substrings that identify the matching
 * form field. Matching is case-insensitive on the field's computed label text.
 *
 * IMPORTANT: order matters. More specific patterns must come first because the
 * matcher walks this list in order and assigns each field to the FIRST key
 * whose pattern matches. "title cross-defendant" must come before
 * "title defendant" — the latter is a substring of the former.
 *
 * Keys correspond to the labeled object built in content.js. "Hearing Date" is
 * captured even though it's not in the rotation sequence — it's an auto-fill
 * convenience. The user explicitly does NOT want any other unrelated fields
 * touched, so any label not on this list is left alone.
 */
const LABEL_PATTERNS = [
  { key: 'caseNumber',               needles: ['case number', 'case no'] },
  { key: 'hearingDate',              needles: ['hearing date'] },
  { key: 'motionType',               needles: ['motion type'] },
  // Cross-* are SINGLE fields on the user's form — one for all cross-
  // complainants and one for all cross-defendants. The user's form labels
  // them "CrossComplainants" / "CrossDefendants" (no separator), but accept
  // hyphenated and spaced variants too. These must come before the
  // plaintiff/defendant patterns so that a field labeled (e.g.)
  // "CrossDefendants" doesn't slip through to `titleDefendant` /
  // `otherDefendants` first. (The needle includes "cross" so a Cross*
  // label can never collide with plain "defendant"/"plaintiff" needles,
  // but ordering remains defensive in case future patterns are added.)
  { key: 'crossComplainants',        needles: ['crosscomplainants', 'cross-complainants', 'cross complainants', 'crosscomplainant', 'cross-complainant', 'cross complainant'] },
  { key: 'crossDefendants',          needles: ['crossdefendants',   'cross-defendants',   'cross defendants',   'crossdefendant',   'cross-defendant',   'cross defendant'] },
  { key: 'titlePlaintiff',           needles: ['title plaintiff'] },
  { key: 'otherPlaintiffs',          needles: ['other plaintiffs', 'other plaintiff'] },
  { key: 'titleDefendant',           needles: ['title defendant'] },
  { key: 'otherDefendants',          needles: ['other defendants', 'other defendant'] },
];

/**
 * Scans the page for editable fields whose label matches one of our known
 * labeled keys, and which has a corresponding value in the labeled object.
 * Returns an array of { element, key, value } — one entry per matched field.
 *
 * Each field is assigned to AT MOST one key (the first one whose needle
 * matches), and each key is assigned to AT MOST one field (the first match
 * found in document order).
 */
function findLabeledFields(labeled) {
  if (!labeled || typeof labeled !== 'object') return [];

  const candidates = document.querySelectorAll(
    'input, textarea, [contenteditable="true"], [contenteditable=""]'
  );

  const usedKeys = new Set();
  const matches = [];

  for (const el of candidates) {
    // Skip non-text inputs (checkbox/radio/file/etc.).
    if (el.tagName.toLowerCase() === 'input') {
      const type = (el.getAttribute('type') || 'text').toLowerCase();
      if (!['text', 'search', 'email', 'tel', 'url', 'date', ''].includes(type)) continue;
    }

    const label = findLabelForElement(el);
    if (!label) continue;
    const lower = label.toLowerCase();

    for (const { key, needles } of LABEL_PATTERNS) {
      if (usedKeys.has(key)) continue;
      if (labeled[key] == null) continue;
      if (needles.some(n => lower.includes(n))) {
        matches.push({ element: el, key, value: labeled[key] });
        usedKeys.add(key);
        break; // this element is now spoken for
      }
    }
  }

  return matches;
}

/**
 * Computes the visible label text for a form field, used for matching against
 * LABEL_PATTERNS. Tries (in order):
 *   1. aria-labelledby (looking up referenced elements) — Microsoft Forms uses
 *      this to point at the question title element
 *   2. an associated <label for=...>
 *   3. wrapping <label>
 *   4. a [role="heading"] in an ancestor (Microsoft Forms' question container)
 *   5. aria-label IF it doesn't look like a generic input-type descriptor
 *      ("Single line text", "Date picker", etc. are useless for matching)
 *   6. placeholder
 *
 * Microsoft Forms sets aria-label to the input TYPE ("Single line text") rather
 * than the question label, so aria-label is a near-useless signal here. We
 * still consult it as a last resort for non-Forms pages.
 */
const GENERIC_ARIA_LABELS = [
  'single line text',
  'multi line text',
  'multi-line text',
  'long text',
  'short text',
  'text input',
  'date picker',
  'datepicker',
  'choice',
  'rating',
  'number',
  'answer',
];

function findLabelForElement(el) {
  if (!el) return '';

  // 1. aria-labelledby — Forms' QuestionId references live here
  const labelledBy = el.getAttribute && el.getAttribute('aria-labelledby');
  if (labelledBy) {
    const ids = labelledBy.split(/\s+/).filter(Boolean);
    const text = ids.map(id => {
      const ref = document.getElementById(id);
      return ref ? (ref.textContent || '').trim() : '';
    }).filter(Boolean).join(' ');
    if (text) return text;
  }

  // 2. <label for="...">
  if (el.id) {
    try {
      const explicit = document.querySelector('label[for="' + CSS.escape(el.id) + '"]');
      if (explicit) {
        const t = (explicit.textContent || '').trim();
        if (t) return t;
      }
    } catch (_) {}
  }

  // 3. wrapping <label>
  const wrapping = el.closest && el.closest('label');
  if (wrapping) {
    const t = (wrapping.textContent || '').trim();
    if (t) return t;
  }

  // 4. role="heading" in an ancestor (Microsoft Forms' question container)
  let node = el.parentElement;
  for (let i = 0; node && i < 12; i++) {
    const heading = node.querySelector && node.querySelector('[role="heading"]');
    if (heading) {
      const t = (heading.textContent || '').trim();
      if (t) return t;
    }
    node = node.parentElement;
  }

  // 5. aria-label, but only if it isn't a generic input-type descriptor
  const aria = el.getAttribute && el.getAttribute('aria-label');
  if (aria && aria.trim()) {
    const lower = aria.trim().toLowerCase();
    if (!GENERIC_ARIA_LABELS.includes(lower)) return aria.trim();
  }

  // 6. placeholder
  const ph = el.getAttribute && el.getAttribute('placeholder');
  if (ph && ph.trim()) return ph.trim();

  return '';
}

/**
 * Returns true if the element is the input inside a Microsoft Forms Fluent
 * date picker. Identified by ARIA combobox role plus an ancestor with the
 * `.ms-DatePicker` class (or `data-automation-id="dateContainer"`). These
 * inputs require typing-simulation to commit a value; direct .value writes
 * are silently dropped by Fluent's parser.
 */
function isFluentDatePicker(el) {
  if (!el || el.tagName !== 'INPUT') return false;
  if ((el.getAttribute('role') || '') !== 'combobox') return false;
  return !!el.closest('.ms-DatePicker, [data-automation-id="dateContainer"]');
}

/**
 * Post-fill commit pass. Standard text inputs and textareas have already
 * been written via insertText; firing blur/focusout is enough to make
 * Microsoft Forms' Fluent text widgets commit. The Fluent date picker is
 * different — it's a combobox whose parser ignores .value writes — so for
 * those elements we re-enter the value via typing-simulation, then press
 * Enter to commit.
 */
function commitFieldChange(el, value) {
  if (isFluentDatePicker(el)) {
    typeIntoFluentDatePicker(el, value);
    return;
  }

  // Plain text input / textarea / contenteditable: blur is enough.
  try {
    el.dispatchEvent(new FocusEvent('blur',     { bubbles: false }));
    el.dispatchEvent(new FocusEvent('focusout', { bubbles: true  }));
  } catch (_) {
    try { el.dispatchEvent(new Event('blur',     { bubbles: false })); } catch (_) {}
    try { el.dispatchEvent(new Event('focusout', { bubbles: true  })); } catch (_) {}
  }
}

/**
 * Drives a Fluent DatePicker by simulating actual typing. Steps:
 *   1. Reset the input value to empty via the prototype setter so React's
 *      tracker sees a transition.
 *   2. Focus the input. The picker may pop a calendar; that's fine.
 *   3. Select the entire current value (so any leftover characters get
 *      replaced when we type).
 *   4. Use document.execCommand('insertText', false, value) — this is what
 *      drives the most accurate beforeinput / input event sequence that
 *      React + Fluent listen to.
 *   5. Press Enter (keydown + keyup) to commit. Fluent's date picker
 *      parses the typed text on Enter and on blur.
 *   6. Blur to be doubly sure.
 */
function typeIntoFluentDatePicker(el, value) {
  try {
    // Step 1: clear via prototype setter (React-friendly).
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    if (setter && setter.set) {
      setter.set.call(el, '');
    } else {
      el.value = '';
    }
    el.dispatchEvent(new Event('input', { bubbles: true }));

    // Step 2: focus.
    try { el.focus(); } catch (_) {}

    // Step 3: select-all so any current text is replaced by what we type.
    try {
      el.setSelectionRange(0, (el.value || '').length);
    } catch (_) {}

    // Step 4: insert via execCommand to drive the native event pipeline.
    // execCommand is deprecated but still works for insertText in inputs.
    // If it returns false (some browsers refuse on type=text inputs), fall
    // back to the setter + input event combo we already used.
    let inserted = false;
    try {
      inserted = document.execCommand('insertText', false, String(value));
    } catch (_) { inserted = false; }

    if (!inserted) {
      if (setter && setter.set) setter.set.call(el, String(value));
      else el.value = String(value);
      el.dispatchEvent(new Event('input',  { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // Step 5: press Enter to commit. Fluent listens for Enter on date inputs.
    const enterDown = new KeyboardEvent('keydown', {
      key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true,
    });
    const enterUp = new KeyboardEvent('keyup', {
      key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true,
    });
    el.dispatchEvent(enterDown);
    el.dispatchEvent(enterUp);

    // Step 6: blur to commit anything still pending.
    try {
      el.dispatchEvent(new FocusEvent('blur',     { bubbles: false }));
      el.dispatchEvent(new FocusEvent('focusout', { bubbles: true  }));
    } catch (_) {
      try { el.dispatchEvent(new Event('blur',     { bubbles: false })); } catch (_) {}
      try { el.dispatchEvent(new Event('focusout', { bubbles: true  })); } catch (_) {}
    }

    console.log('[LACourt-rotator] typed into Fluent date picker:', value);
  } catch (err) {
    console.log('[LACourt-rotator] Fluent date picker typing failed:', err);
  }
}

/**
 * Clears the content of a field (input, textarea, or contenteditable).
 * Dispatches events so React/Angular/Forms frameworks notice the change.
 */
function clearField(target) {
  try { target.focus(); } catch (_) {}

  const tag = (target.tagName || '').toLowerCase();

  if (tag === 'input' || tag === 'textarea') {
    const proto = tag === 'textarea' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value');
    if (setter && setter.set) {
      setter.set.call(target, '');
    } else {
      target.value = '';
    }
    target.dispatchEvent(new Event('input', { bubbles: true }));
    target.dispatchEvent(new Event('change', { bubbles: true }));
    return;
  }

  if (target.isContentEditable) {
    target.innerHTML = '';
    target.dispatchEvent(new Event('input', { bubbles: true }));
    target.dispatchEvent(new Event('change', { bubbles: true }));
    return;
  }
}

function insertText(target, value) {
  // Make sure the target has focus. Microsoft Forms sometimes loses focus
  // during paste-event processing; execCommand needs an active selection.
  try { target.focus(); } catch (_) {}

  const tag = (target.tagName || '').toLowerCase();

  if (tag === 'input' || tag === 'textarea') {
    // For input/textarea, use the native setter so React notices the change.
    const start = target.selectionStart != null ? target.selectionStart : (target.value || '').length;
    const end = target.selectionEnd != null ? target.selectionEnd : (target.value || '').length;
    const beforeValue = target.value || '';
    const next = beforeValue.substring(0, start) + value + beforeValue.substring(end);

    // Use the prototype setter so React's synthetic-event tracker sees the change.
    const proto = tag === 'textarea' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value');
    if (setter && setter.set) {
      setter.set.call(target, next);
    } else {
      target.value = next;
    }
    const newPos = start + value.length;
    target.selectionStart = target.selectionEnd = newPos;

    target.dispatchEvent(new Event('input', { bubbles: true }));
    target.dispatchEvent(new Event('change', { bubbles: true }));

    return {
      element: target,
      kind: 'input',
      beforeValue,
      afterValue: next,
      caretBefore: start,
      ts: Date.now(),
    };
  }

  // Contenteditable path: prefer execCommand('insertText'), which is the most
  // React/Angular/Forms-friendly way to insert because it fires a native
  // beforeinput+input event sequence with inputType="insertFromPaste"-ish
  // behavior. Falls back to manual range insertion if execCommand is gone.
  if (target.isContentEditable) {
    const beforeHTML = target.innerHTML;
    const beforeText = target.innerText;

    let inserted = false;
    try {
      inserted = document.execCommand('insertText', false, value);
    } catch (_) { inserted = false; }

    if (!inserted) {
      // Fallback: manual DOM insertion.
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0) {
        sel.deleteFromDocument();
        const range = sel.getRangeAt(0);

        const frag = document.createDocumentFragment();
        const parts = value.split('\n');
        parts.forEach((part, i) => {
          if (i > 0) frag.appendChild(document.createElement('br'));
          if (part) frag.appendChild(document.createTextNode(part));
        });

        const lastChild = frag.lastChild;
        range.insertNode(frag);

        if (lastChild) {
          range.setStartAfter(lastChild);
          range.collapse(true);
        }
        sel.removeAllRanges();
        sel.addRange(range);
      }

      target.dispatchEvent(new Event('input', { bubbles: true }));
      target.dispatchEvent(new Event('change', { bubbles: true }));
    }

    return {
      element: target,
      kind: 'editable',
      beforeHTML,
      afterHTML: target.innerHTML,
      beforeText,
      ts: Date.now(),
    };
  }

  return null;
}

/**
 * Restores the field's pre-paste state. Returns true if the snapshot was
 * applied, false if the field has been modified since (in which case the
 * caller should fall through to the browser's default Ctrl+Z behavior).
 */
function restoreSnapshot(snap) {
  if (!snap || !snap.element || !snap.element.isConnected) return false;

  if (snap.kind === 'input') {
    const target = snap.element;
    // Only undo if the field still contains exactly what we left.
    if (target.value !== snap.afterValue) return false;

    const tag = (target.tagName || '').toLowerCase();
    const proto = tag === 'textarea' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value');
    if (setter && setter.set) {
      setter.set.call(target, snap.beforeValue);
    } else {
      target.value = snap.beforeValue;
    }
    try {
      target.selectionStart = target.selectionEnd = snap.caretBefore;
    } catch (_) {}
    target.dispatchEvent(new Event('input', { bubbles: true }));
    target.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  if (snap.kind === 'editable') {
    const target = snap.element;
    if (target.innerHTML !== snap.afterHTML) return false;

    target.innerHTML = snap.beforeHTML;
    target.dispatchEvent(new Event('input', { bubbles: true }));
    target.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  return false;
}

/* ------------------------------------------------------------------ */
/* Auto-fill-on-load: triggered by the popup-driven workflow.         */
/* ------------------------------------------------------------------ */

(function setupAutoFillOnLoad() {
  // Microsoft Forms is a heavy SPA; fields render after first paint. We need
  // to (a) check whether autoFillOnLoad is set, (b) wait for editable fields
  // to render, (c) try to match them against our labeled keys, (d) fill once
  // enough match. A MutationObserver handles the rendering wait, with a hard
  // timeout so we don't watch forever if no labeled fields ever appear.

  const TIMEOUT_MS = 30_000;
  const POLL_INTERVAL_MS = 500;
  let attemptStart = 0;
  let observer = null;
  let pollTimer = null;
  let done = false;

  function tryAutoFill() {
    if (done) return;
    chrome.runtime.sendMessage({ type: 'getRotation' }, response => {
      if (chrome.runtime.lastError) return;
      const rot = response && response.rotation;
      if (!rot || !rot.autoFillOnLoad) {
        // No flag set, or rotation cleared — stop watching.
        teardown();
        return;
      }

      const matches = findLabeledFields(rot.labeled || {});
      if (matches.length < 2) {
        // Not enough fields rendered yet. Keep watching.
        if (Date.now() - attemptStart > TIMEOUT_MS) {
          console.log('[LACourt-rotator] auto-fill-on-load timed out (no fields matched)');
          teardown();
        }
        return;
      }

      done = true;
      console.log('[LACourt-rotator] auto-fill-on-load: clearing and filling',
        matches.length, 'fields:', matches.map(m => m.key));

      // First, clear all matched fields
      for (const m of matches) {
        clearField(m.element);
      }

      // Then fill them with the new values
      for (const m of matches) {
        const snap = insertText(m.element, m.value);
        if (snap) {
          snap.autoFillKey = m.key;
          pasteStack.push(snap);
        }
      }

      // Post-fill commit pass (see analogous block in the paste-event
      // handler for rationale — Fluent date picker needs special handling).
      for (const m of matches) {
        commitFieldChange(m.element, m.value);
      }

      // Clear the flag and exhaust the rotation.
      try {
        chrome.runtime.sendMessage({ type: 'clearAutoFillFlag' }, () => {});
        chrome.runtime.sendMessage({ type: 'exhaustRotation' }, () => {});
      } catch (_) {}

      teardown();
    });
  }

  function teardown() {
    if (observer) { try { observer.disconnect(); } catch (_) {} observer = null; }
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  }

  function start() {
    attemptStart = Date.now();
    // Try immediately in case fields are already there.
    tryAutoFill();
    if (done) return;

    // Watch for any DOM additions; on each addition, retry.
    observer = new MutationObserver(() => { tryAutoFill(); });
    try {
      observer.observe(document.body || document.documentElement, {
        childList: true, subtree: true,
      });
    } catch (_) {}

    // Belt and braces: also poll. MutationObserver alone misses some Forms
    // re-renders that happen via React's reconciliation without a tree mutation
    // we can detect.
    pollTimer = setInterval(tryAutoFill, POLL_INTERVAL_MS);
  }

  // Only set up on pages where it makes sense (skip extension pages, etc.).
  if (location.protocol !== 'http:' && location.protocol !== 'https:') return;

  // Check the flag once on load. If it's set, start watching.
  // We use a small initial delay to let the page begin rendering.
  setTimeout(() => {
    try {
      chrome.runtime.sendMessage({ type: 'getRotation' }, response => {
        if (chrome.runtime.lastError) return;
        const rot = response && response.rotation;
        if (rot && rot.autoFillOnLoad) {
          console.log('[LACourt-rotator] auto-fill-on-load: armed, watching for fields');
          start();
        }
      });
    } catch (_) {}
  }, 250);
})();
})();

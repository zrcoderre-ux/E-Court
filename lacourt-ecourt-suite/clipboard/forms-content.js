/**
 * LA Court Clipboard Cleaner - Microsoft Forms Content Script
 *
 * When case party data has been captured from a LA Court case page (via Ctrl+A),
 * this script intercepts paste events on forms.microsoft.com and inserts the
 * correct value based on the question label of the focused field.
 *
 * Field mapping (case-insensitive label matching):
 *   "case number"       → caseNumber
 *   "title plaintiff"   → titlePlaintiff
 *   "other plaintiff"   → otherPlaintiffs
 *   "title defendant"   → titleDefendant
 *   "other defendant"   → otherDefendants
 */

document.addEventListener('paste', function (e) {
  // Read the latest case data from storage each time
  chrome.storage.local.get(['casePartyData'], result => {
    const data = result.casePartyData;
    if (!data) return; // No case data captured yet — normal paste

    const target = document.activeElement;
    if (!target) return;

    const tag = target.tagName.toLowerCase();
    if (tag !== 'input' && tag !== 'textarea' && !target.isContentEditable) return;

    const label = findFormLabel(target);
    if (!label) return;

    const lower = label.toLowerCase();
    let value = null;

    if (lower.includes('case number') || lower.includes('case no')) {
      value = data.caseNumber;
    } else if (lower.includes('title plaintiff')) {
      value = data.titlePlaintiff;
    } else if (lower.includes('other plaintiff')) {
      value = data.otherPlaintiffs;
    } else if (lower.includes('title defendant')) {
      value = data.titleDefendant;
    } else if (lower.includes('other defendant')) {
      value = data.otherDefendants;
    }

    if (value === null) return; // Unrecognized field — allow normal paste

    e.preventDefault();
    e.stopImmediatePropagation();

    insertText(target, value);
  });
}, true); // capture phase so we run before the form's own handlers

/**
 * Inserts text into a focused element, replacing any current selection.
 * Works for <input>, <textarea>, and contenteditable elements.
 */
function insertText(target, value) {
  if (target.isContentEditable) {
    // contenteditable div (Microsoft Forms uses these)
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      sel.deleteFromDocument();
      const range = sel.getRangeAt(0);
      range.insertNode(document.createTextNode(value));
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    }
  } else {
    // Regular input / textarea
    const start = target.selectionStart;
    const end   = target.selectionEnd;
    const current = target.value;
    target.value = current.substring(0, start) + value + current.substring(end);
    target.selectionStart = target.selectionEnd = start + value.length;
  }

  // Dispatch events so Microsoft Forms' React/Angular picks up the change
  target.dispatchEvent(new Event('input',  { bubbles: true }));
  target.dispatchEvent(new Event('change', { bubbles: true }));
}

/**
 * Finds the question label associated with the currently focused form element.
 * Microsoft Forms renders questions in a structure like:
 *   <div class="question-container">
 *     <div role="heading">Question label text</div>
 *     ...
 *     <input> or contenteditable div
 *   </div>
 */
function findFormLabel(el) {
  // Walk up the DOM looking for a question container
  let node = el;
  for (let i = 0; i < 10; i++) {
    if (!node || node === document.body) break;

    // Microsoft Forms uses role="heading" for question titles
    const heading = node.querySelector('[role="heading"]');
    if (heading) {
      const text = heading.textContent.trim();
      if (text) return text;
    }

    // Also check for aria-label on the element itself or its container
    const ariaLabel = node.getAttribute('aria-label');
    if (ariaLabel) return ariaLabel;

    // Check for a <label> nearby
    const label = node.querySelector('label');
    if (label) {
      const text = label.textContent.trim();
      if (text) return text;
    }

    node = node.parentElement;
  }

  // Fallback: aria-label or placeholder on the input itself
  return el.getAttribute('aria-label') || el.getAttribute('placeholder') || null;
}

/**
 * LA Court Clipboard Cleaner - Microsoft Forms Content Script
 *
 * Listens for paste events on Microsoft Forms pages.
 * When party data has been stored (by Ctrl+A on a case page),
 * detects the focused field label and inserts the matching value.
 *
 * Field labels matched (case-insensitive):
 *   "Title Plaintiff"    → first plaintiff name
 *   "Other Plaintiffs"   → remaining plaintiff names (newline-separated)
 *   "Title Defendant"    → first defendant name
 *   "Other Defendants"   → remaining defendant names (newline-separated)
 *   "Case Number" / "Case No" → case number
 */

document.addEventListener('paste', function (e) {
  // Load the stored party data
  chrome.storage.session.get(['casePartyData'], result => {
    if (!result || !result.casePartyData) return;
    const data = result.casePartyData;

    const target = e.target;
    if (!target) return;

    const tag = target.tagName.toLowerCase();
    if (tag !== 'input' && tag !== 'textarea') return;

    const label = findLabelForElement(target);
    if (!label) return;

    const labelText = label.toLowerCase();
    let value = null;

    if (labelText.includes('title plaintiff')) {
      value = data.titlePlaintiff;
    } else if (labelText.includes('other plaintiff')) {
      value = data.otherPlaintiffs;
    } else if (labelText.includes('title defendant')) {
      value = data.titleDefendant;
    } else if (labelText.includes('other defendant')) {
      value = data.otherDefendants;
    } else if (labelText.includes('case number') || labelText.includes('case no')) {
      value = data.caseNumber;
    }

    if (value === null) return;

    e.preventDefault();
    e.stopPropagation();

    // Insert value at cursor position
    const start = target.selectionStart != null ? target.selectionStart : target.value.length;
    const end = target.selectionEnd != null ? target.selectionEnd : target.value.length;
    const current = target.value || '';
    target.value = current.substring(0, start) + value + current.substring(end);
    target.selectionStart = target.selectionEnd = start + value.length;

    target.dispatchEvent(new Event('input', { bubbles: true }));
    target.dispatchEvent(new Event('change', { bubbles: true }));
  });
});

function findLabelForElement(el) {
  if (el.getAttribute('aria-label')) return el.getAttribute('aria-label');

  const labelledBy = el.getAttribute('aria-labelledby');
  if (labelledBy) {
    const ids = labelledBy.split(/\s+/);
    const texts = ids.map(id => {
      const el2 = document.getElementById(id);
      return el2 ? el2.textContent.trim() : '';
    });
    const joined = texts.filter(Boolean).join(' ');
    if (joined) return joined;
  }

  if (el.id) {
    const explicit = document.querySelector('label[for="' + CSS.escape(el.id) + '"]');
    if (explicit) return explicit.textContent.trim();
  }

  const wrapping = el.closest('label');
  if (wrapping) return wrapping.textContent.trim();

  // Walk up to find Microsoft Forms question title
  let node = el.parentElement;
  for (let i = 0; i < 8; i++) {
    if (!node) break;
    const title = node.querySelector(
      '[role="heading"], .office-form-question-title, ' +
      '[data-automation-id="question-title"], label'
    );
    if (title && title.textContent.trim()) return title.textContent.trim();
    node = node.parentElement;
  }

  if (el.placeholder) return el.placeholder;
  return null;
}

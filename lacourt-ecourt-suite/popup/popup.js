/**
 * LA Court E-Court Suite - Popup
 *
 * Combines the two source popups:
 *   - "Fill Microsoft Form" (from the former Clipboard Cleaner popup): captures
 *     case-party data from the active case tab and opens the target form on the
 *     opposite display. Only works on a civil.lacourt.org case page.
 *   - "Manage Exclusion Lists" (from the former Agenda Cleaner popup): opens the
 *     shared options page, which manages both the agenda exclusion terms and the
 *     dismissed-party motion exclusions.
 *
 * The message types (captureForFormFill, fireMailto, openFormOnOppositeDisplay)
 * are unchanged from the standalone build; the case content script and the
 * service worker still handle them.
 */

const CASE_URL_RE = /^https:\/\/civil\.lacourt\.org\/ecourt\/ecms\/case/;

const btn = document.getElementById('fillFormBtn');
const status = document.getElementById('status');

function setStatus(text, type) {
  status.textContent = text || '';
  status.className = 'status' + (type ? ' ' + type : '');
}

document.getElementById('openOptions').addEventListener('click', () => {
  chrome.runtime.openOptionsPage();
});

btn.addEventListener('click', async () => {
  setStatus('Working...');
  btn.disabled = true;

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab || !tab.url || !CASE_URL_RE.test(tab.url)) {
      setStatus('Open a case page first.', 'error');
      btn.disabled = false;
      return;
    }

    const captured = await new Promise((resolve, reject) => {
      chrome.tabs.sendMessage(tab.id, { type: 'captureForFormFill' }, response => {
        if (chrome.runtime.lastError) { reject(chrome.runtime.lastError); return; }
        resolve(response);
      });
    });

    if (!captured || !captured.ok) {
      const reason = captured && captured.reason ? captured.reason : 'unknown';
      setStatus('Could not capture data (' + reason + ').', 'error');
      btn.disabled = false;
      return;
    }

    // Non-OSC cases now open the in-extension Order Template popup instead of
    // a Microsoft Form; openUrl carries whichever is appropriate. Fall back to
    // formUrl for safety if an older content script is still loaded.
    const openUrl = captured.openUrl || captured.formUrl;
    if (!openUrl) {
      setStatus('No form URL returned.', 'error');
      btn.disabled = false;
      return;
    }

    const openingLabel = captured.isOrderTemplate ? 'Order Template' : 'form';
    setStatus('Captured ' + captured.count + ' values. Opening ' + openingLabel + '...', 'success');

    // Fire mailto (OSC default-judgment cases only). Fire-and-forget - the
    // case page launches the OS mail handler from its own context so popup
    // closure doesn't interrupt it.
    if (captured.mailtoUrl) {
      chrome.tabs.sendMessage(tab.id, { type: 'fireMailto', url: captured.mailtoUrl }, () => {
        void chrome.runtime.lastError;
      });
    }

    await new Promise((resolve, reject) => {
      chrome.runtime.sendMessage({
        type: 'openFormOnOppositeDisplay',
        url: openUrl,
        sourceWindowId: tab.windowId,
      }, response => {
        if (chrome.runtime.lastError) { reject(chrome.runtime.lastError); return; }
        if (!response || !response.ok) {
          reject(new Error(response && response.error || 'open failed'));
          return;
        }
        resolve();
      });
    });

    setTimeout(() => window.close(), 300);
  } catch (err) {
    console.error('[EcourtSuite-popup] error:', err);
    setStatus('Error: ' + (err.message || err), 'error');
    btn.disabled = false;
  }
});

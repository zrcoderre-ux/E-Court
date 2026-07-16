/**
 * LA Court E-Court Suite - Service Worker
 *
 * This is the single background service worker for the combined extension.
 * MV3 allows only one service worker, so the two source extensions that each
 * had their own background script are loaded here via importScripts.
 *
 * Each imported module registers its own chrome.runtime.onMessage listener.
 * Chrome supports multiple onMessage listeners; each inspects the message and
 * ignores ones it does not own, so there is no conflict:
 *
 *   clipboard/background.js  - handles msg.type: getRotation, advanceRotation,
 *                              rewindRotation, exhaustRotation, setRotation,
 *                              openFormOnOppositeDisplay, clearAutoFillFlag
 *   pdf-focus/background.js  - handles msg.type: OPEN_DOC_BACKGROUND
 *
 * The two namespaces do not overlap.
 */

importScripts(
  'clipboard/background.js',
  'pdf-focus/background.js'
);

// Let content scripts read/write chrome.storage.session (used to cache the OSC
// default-judgment status per case, so it is computed once per browser session
// rather than re-fetched on every tab navigation). Default access is
// trusted-contexts-only; widen it to include content scripts.
try {
  if (chrome.storage.session && chrome.storage.session.setAccessLevel) {
    chrome.storage.session.setAccessLevel({ accessLevel: 'TRUSTED_AND_UNTRUSTED_CONTEXTS' });
  }
} catch (_) {}

// ---------------------------------------------------------------------------
// Agenda auto-advance-on-paste
//
// When enabled on an agenda tab, open a persistent native-messaging port to the
// host, which watches the OS clipboard in the background. After the user's Excel
// paste macro clears the clipboard, the host sends {event:'clipboardEmpty'} and
// we tell that agenda tab to advance to the next day. Native messaging is only
// available here (the service worker), not in content scripts, so the agenda
// script registers via runtime messages and we relay back with chrome.tabs.
// ---------------------------------------------------------------------------
(function () {
  const NATIVE_HOST = 'com.lacourt.ecourt_host';
  let port = null;
  let agendaTabId = null;

  function ensurePort() {
    if (port) return port;
    try {
      port = chrome.runtime.connectNative(NATIVE_HOST);
    } catch (e) {
      port = null;
      return null;
    }
    port.onMessage.addListener((msg) => {
      if (!msg) return;
      if (msg.event === 'clipboardEmpty' && agendaTabId != null) {
        chrome.tabs.sendMessage(agendaTabId, { type: 'AGENDA_ADVANCE' }, () => { void chrome.runtime.lastError; });
      }
      // 'keepalive' events just reset this worker's idle timer — nothing to do.
    });
    port.onDisconnect.addListener(() => { void chrome.runtime.lastError; port = null; });
    try { port.postMessage({ action: 'watchClipboard' }); } catch (_) {}
    return port;
  }

  function closePort() {
    if (!port) return;
    try { port.postMessage({ action: 'stopClipboardWatch' }); } catch (_) {}
    try { port.disconnect(); } catch (_) {}
    port = null;
  }

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (!msg || msg.type !== 'AGENDA_AUTOADVANCE_WATCH') return; // not ours
    if (msg.enabled && sender.tab && sender.tab.id != null) {
      agendaTabId = sender.tab.id;
      ensurePort();
    } else {
      if (sender.tab && sender.tab.id === agendaTabId) agendaTabId = null;
      closePort();
    }
    try { sendResponse({ ok: true }); } catch (_) {}
    return false;
  });

  chrome.tabs.onRemoved.addListener((tabId) => {
    if (tabId === agendaTabId) { agendaTabId = null; closePort(); }
  });
})();

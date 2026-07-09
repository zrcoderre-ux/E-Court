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

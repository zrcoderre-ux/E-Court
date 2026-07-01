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

/**
 * LA Court Clipboard Cleaner - Service Worker
 *
 * Content scripts in some frames (e.g., certain Microsoft Forms iframes) have
 * `chrome.storage` undefined. We route storage access through this service
 * worker so all frames can reliably read/write the rotation state.
 *
 * Messages handled:
 *   { type: 'getRotation' }                              -> { rotation }
 *   { type: 'advanceRotation' }                          -> { rotation, value }
 *      Reads rotation, returns the current value, increments index, writes back.
 *      Atomic from the caller's perspective (single round trip).
 *   { type: 'setRotation', sequence }                    -> { ok: true }
 */

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (!msg || typeof msg !== 'object') return false;

  if (msg.type === 'getRotation') {
    chrome.storage.local.get(['lacourtRotation'], result => {
      sendResponse({ rotation: result && result.lacourtRotation || null });
    });
    return true; // keep channel open for async response
  }

  if (msg.type === 'advanceRotation') {
    chrome.storage.local.get(['lacourtRotation'], result => {
      const rot = result && result.lacourtRotation;
      if (!rot || !Array.isArray(rot.sequence) || rot.sequence.length === 0) {
        sendResponse({ rotation: null, value: null });
        return;
      }
      const idx = rot.index || 0;
      if (idx >= rot.sequence.length) {
        sendResponse({ rotation: rot, value: null }); // exhausted
        return;
      }
      const value = rot.sequence[idx];
      const updated = { ...rot, index: idx + 1 };
      chrome.storage.local.set({ lacourtRotation: updated }, () => {
        sendResponse({ rotation: updated, value });
      });
    });
    return true;
  }

  if (msg.type === 'rewindRotation') {
    chrome.storage.local.get(['lacourtRotation'], result => {
      const rot = result && result.lacourtRotation;
      if (!rot || !Array.isArray(rot.sequence) || rot.sequence.length === 0) {
        sendResponse({ rotation: null });
        return;
      }
      const newIdx = Math.max(0, (rot.index || 0) - 1);
      const updated = { ...rot, index: newIdx };
      chrome.storage.local.set({ lacourtRotation: updated }, () => {
        sendResponse({ rotation: updated });
      });
    });
    return true;
  }

  if (msg.type === 'exhaustRotation') {
    chrome.storage.local.get(['lacourtRotation'], result => {
      const rot = result && result.lacourtRotation;
      if (!rot) {
        sendResponse({ ok: false });
        return;
      }
      const updated = { ...rot, index: rot.sequence.length };
      chrome.storage.local.set({ lacourtRotation: updated }, () => {
        sendResponse({ ok: true });
      });
    });
    return true;
  }

  if (msg.type === 'setRotation' && Array.isArray(msg.sequence)) {
    const rotation = {
      sequence: msg.sequence,
      labeled: msg.labeled || {},
      index: 0,
      autoFillOnLoad: !!msg.autoFillOnLoad,
      createdAt: Date.now(),
    };
    chrome.storage.local.set({ lacourtRotation: rotation }, () => {
      sendResponse({ ok: true });
    });
    return true;
  }

  if (msg.type === 'openFormOnOppositeDisplay' && typeof msg.url === 'string') {
    // If sourceWindowId wasn't provided, infer it from the sender (the tab
    // that messaged us is on the source window).
    const srcId = msg.sourceWindowId != null
      ? msg.sourceWindowId
      : (_sender && _sender.tab && _sender.tab.windowId);
    openFormOnOppositeDisplay(msg.url, srcId)
      .then(() => sendResponse({ ok: true }))
      .catch(err => {
        console.error('[LACourt-bg] openFormOnOppositeDisplay failed:', err);
        sendResponse({ ok: false, error: String(err && err.message || err) });
      });
    return true;
  }

  if (msg.type === 'clearAutoFillFlag') {
    chrome.storage.local.get(['lacourtRotation'], result => {
      const rot = result && result.lacourtRotation;
      if (!rot) {
        sendResponse({ ok: false });
        return;
      }
      const updated = { ...rot, autoFillOnLoad: false };
      chrome.storage.local.set({ lacourtRotation: updated }, () => {
        sendResponse({ ok: true });
      });
    });
    return true;
  }

  return false;
});

/**
 * Opens the given URL in a new window placed on a display that does NOT
 * contain the source window. Falls back to default placement if there's only
 * one display or the source window's location can't be determined.
 */
async function openFormOnOppositeDisplay(url, sourceWindowId) {
  let placement = null;
  try {
    if (sourceWindowId != null && chrome.system && chrome.system.display) {
      const [sourceWindow, displays] = await Promise.all([
        chrome.windows.get(sourceWindowId),
        chrome.system.display.getInfo(),
      ]);
      if (displays && displays.length >= 2) {
        const cx = (sourceWindow.left || 0) + (sourceWindow.width  || 0) / 2;
        const cy = (sourceWindow.top  || 0) + (sourceWindow.height || 0) / 2;
        const sourceDisplay = displays.find(d => {
          const b = d.bounds;
          return cx >= b.left && cx < b.left + b.width
              && cy >= b.top  && cy < b.top  + b.height;
        });
        const target = displays.find(d => d !== sourceDisplay) || displays[1];
        if (target) {
          const wa = target.workArea || target.bounds;
          placement = { left: wa.left, top: wa.top, width: wa.width, height: wa.height };
        }
      }
    }
  } catch (err) {
    console.warn('[LACourt-bg] display lookup failed, using default placement:', err);
  }

  const opts = { url, type: 'normal', focused: true };
  if (placement) {
    Object.assign(opts, placement, { state: 'normal' });
  }
  await chrome.windows.create(opts);
}

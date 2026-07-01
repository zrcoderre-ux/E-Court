/**
 * LA Court PDF Focus - Background Service Worker
 *
 * Receives OPEN_DOC_BACKGROUND messages from the bridge script and opens
 * the document URL as a background tab (active: false) right next to the
 * tab that requested it. No focus change, no flash.
 */

chrome.runtime.onMessage.addListener((message, sender) => {
  if (!message || message.type !== 'OPEN_DOC_BACKGROUND') return;
  if (!message.url || !sender.tab) return;

  // Only honor requests from lacourt tabs, as a safety check.
  if (!sender.tab.url || !sender.tab.url.includes('civil.lacourt.org')) return;

  chrome.tabs.create({
    url: message.url,
    active: false,
    windowId: sender.tab.windowId,
    index: sender.tab.index + 1,
    openerTabId: sender.tab.id
  });
});

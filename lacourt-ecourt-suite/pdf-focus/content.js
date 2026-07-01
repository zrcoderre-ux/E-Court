/**
 * LA Court PDF Focus - Content Script (MAIN world)
 *
 * Intercepts window.open() calls for e-court document URLs and suppresses
 * them. The URL is forwarded to the bridge script, which asks the background
 * worker to open it as a background tab instead (no focus change, no flash).
 *
 * Returns a stub window object so any page code that calls .focus() or
 * checks the return value doesn't break.
 */

(function () {
  const originalOpen = window.open.bind(window);

  // Stub Window object that absorbs e-court's post-open API calls
  // (e.g. updateWindowTitle, focus, etc.) without throwing.
  function makeStubWindow() {
    // Stub element returned from any DOM lookup. All methods/properties
    // are no-ops or empty values so chained calls like
    // doc.getElementById('x').getAttribute('y') don't throw.
    function makeStubElement() {
      const el = {
        // Attribute access
        getAttribute: function () { return ''; },
        setAttribute: function () {},
        removeAttribute: function () {},
        hasAttribute: function () { return false; },
        // Class / style
        classList: {
          add: function () {},
          remove: function () {},
          toggle: function () {},
          contains: function () { return false; }
        },
        style: {},
        // Content
        innerHTML: '',
        outerHTML: '',
        textContent: '',
        innerText: '',
        value: '',
        title: '',
        id: '',
        className: '',
        // Tree
        children: [],
        childNodes: [],
        firstChild: null,
        lastChild: null,
        parentNode: null,
        parentElement: null,
        nextSibling: null,
        previousSibling: null,
        // Methods
        appendChild: function (c) { return c; },
        removeChild: function (c) { return c; },
        insertBefore: function (c) { return c; },
        replaceChild: function (c) { return c; },
        cloneNode: function () { return makeStubElement(); },
        contains: function () { return false; },
        querySelector: function () { return makeStubElement(); },
        querySelectorAll: function () { return []; },
        getElementsByTagName: function () { return []; },
        getElementsByClassName: function () { return []; },
        addEventListener: function () {},
        removeEventListener: function () {},
        dispatchEvent: function () { return true; },
        focus: function () {},
        blur: function () {},
        click: function () {},
        getBoundingClientRect: function () {
          return { top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 };
        }
      };
      return el;
    }

    const stubDoc = {
      title: '',
      write: function () {},
      writeln: function () {},
      open: function () {},
      close: function () {},
      // Lookups return stub elements (NOT null) so chained calls don't crash.
      getElementById: function () { return makeStubElement(); },
      getElementsByTagName: function () { return []; },
      getElementsByClassName: function () { return []; },
      getElementsByName: function () { return []; },
      querySelector: function () { return makeStubElement(); },
      querySelectorAll: function () { return []; },
      createElement: function () { return makeStubElement(); },
      createTextNode: function () { return makeStubElement(); },
      addEventListener: function () {},
      removeEventListener: function () {},
      readyState: 'complete',
      body: makeStubElement(),
      head: makeStubElement(),
      documentElement: makeStubElement()
    };

    const stubWin = {
      focus: function () {},
      blur: function () {},
      close: function () {},
      closed: false,
      opener: null,
      location: {
        href: '',
        replace: function () {},
        assign: function () {},
        reload: function () {}
      },
      document: stubDoc,
      postMessage: function () {},
      addEventListener: function () {},
      removeEventListener: function () {},
      setTimeout: function () { return 0; },
      clearTimeout: function () {},
      setInterval: function () { return 0; },
      clearInterval: function () {}
    };

    // Self-references so things like stubWin.window.document also work.
    stubWin.window = stubWin;
    stubWin.self = stubWin;
    stubWin.top = stubWin;
    stubWin.parent = stubWin;

    return stubWin;
  }

  window.open = function (url, target, features) {
    const urlStr = (url || '').toString();

    if (urlStr.includes('/ecourt/ecms/doc')) {
      // Resolve the URL to absolute form so the background worker can fetch it.
      let absoluteUrl;
      try {
        absoluteUrl = new URL(urlStr, window.location.href).href;
      } catch (e) {
        absoluteUrl = urlStr;
      }

      window.dispatchEvent(new CustomEvent('LACOURT_OPEN_DOC', {
        detail: { url: absoluteUrl }
      }));

      // Suppress the original window.open entirely.
      return makeStubWindow();
    }

    return originalOpen(url, target, features);
  };
})();

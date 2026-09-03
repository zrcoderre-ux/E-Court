/**
 * One case-level PDF, in its own tab.
 *
 * Two jobs, both opened by the Documents button:
 *
 *   ?pdf=…   Pass through a PDF the court already builds — the Register of
 *            Actions, which eCourt's report runner produces from the case
 *            number. It is fetched and shown here rather than navigated to
 *            because the server names it after its own batch job
 *            ("RegisterOfActions-PRODUCTION2-2026-09-02"), and that name is
 *            what Chrome would title the tab and the saved file. Held in this
 *            tab, the name is ours.
 *
 *   ?url=…   Render a case sub-tab that has no print endpoint at all — the
 *            Parties tab. The page is FETCHED here rather than read off the
 *            live tab, so what gets rendered is the server's own HTML: the
 *            extension's buttons, header rewrites and status spans never
 *            existed in it.
 *
 * Either way the tab owns the finished PDF, so Export tells this page to save
 * it (savePagePdf) rather than trying to re-derive it in the background: the
 * bytes are already here, under the name they should be filed under.
 *
 * Query string: pdf | url, plus title, subtitle, pdfname. The save-as name is
 * `pdfname` rather than `file` because the background's resolveOpenPdfUrl reads
 * everything after a `file=` as a document URL.
 */
(function () {
  'use strict';

  const params = new URLSearchParams(location.search);
  const sourcePdf = params.get('pdf') || '';
  const pageUrl = params.get('url') || '';
  const title = params.get('title') || 'eCourt page';
  const subtitle = params.get('subtitle') || '';
  const fileName = (params.get('pdfname') || title)
    .replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();

  const $ = id => document.getElementById(id);
  document.title = title;
  $('name').textContent = title;

  const isCourtUrl = u => /^https:\/\/[^/]*lacourt\.org\//i.test(u);

  /* The finished PDF, once there is one, and Export's hook into it.

     The download is started HERE rather than handed back to the background,
     because what this tab holds is a blob of its own: a blob URL belongs to the
     context that made it, and bytes passed back as a data: URL would be the
     whole document through a message channel. So Export says "save yourself",
     this page calls chrome.downloads with the name the document should be filed
     under, and the background waits on the download id it gets back. A tab still
     fetching answers when it is done. */
  let blobUrl = null;
  let resolveReady;
  const ready = new Promise(r => { resolveReady = r; });

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (!msg || msg.type !== 'savePagePdf') return false;
    ready.then(() => {
      if (!blobUrl) { sendResponse(null); return; }
      try {
        chrome.downloads.download({ url: blobUrl, filename: fileName + '.pdf' }, id => {
          void chrome.runtime.lastError;
          sendResponse(id == null ? null : { downloadId: id });
        });
      } catch (_) { sendResponse(null); }
    });
    return true; // the PDF may still be on its way
  });

  // Built as nodes rather than markup: the URL being reported comes in on the
  // query string, and a URL is never trusted as HTML.
  function fail(message, linkUrl) {
    const status = $('status');
    status.textContent = '';
    status.appendChild(document.createTextNode(message));
    if (linkUrl) {
      status.appendChild(document.createElement('br'));
      status.appendChild(document.createElement('br'));
      status.appendChild(document.createTextNode('Open '));
      const a = document.createElement('a');
      a.href = linkUrl;
      a.textContent = 'it directly';
      status.appendChild(a);
      status.appendChild(document.createTextNode(
        ' to check you are still signed in to eCourt, then click Documents again.'));
    }
    status.hidden = false;
    $('frame').hidden = true;
    resolveReady();
  }

  function show(blob) {
    blobUrl = URL.createObjectURL(blob);
    const frame = $('frame');
    frame.src = blobUrl;
    frame.hidden = false;
    $('status').hidden = true;
    const save = $('save');
    save.href = blobUrl;
    save.setAttribute('download', fileName + '.pdf');
    save.hidden = false;
    resolveReady();
  }

  const failed = (what, err) => {
    // The usual cause is a signed-out session: the fetch carries the tab's
    // cookies, so if the court site has logged the user out it answers with the
    // login page instead of the document.
    fail('Could not ' + what + ' — ' + String(err && err.message || err) + '.', sourcePdf || pageUrl);
    console.error('[LACourt-PagePDF]', err);
  };

  if (isCourtUrl(sourcePdf)) {
    fetch(sourcePdf, { credentials: 'include' })
      .then(r => {
        if (!r.ok) throw new Error('the court site answered ' + r.status);
        // A signed-out session answers 200 with the login page, which would
        // otherwise be framed here under the register's name.
        const ct = (r.headers.get('content-type') || '').split(';')[0].trim();
        if (!/pdf/i.test(ct)) throw new Error('the court site returned ' + (ct || 'no content type') + ', not a PDF');
        return r.blob();
      })
      .then(blob => show(blob))
      .catch(err => failed('fetch that report', err));
  } else if (isCourtUrl(pageUrl)) {
    fetch(pageUrl, { credentials: 'include' })
      .then(r => {
        if (!r.ok) throw new Error('the court site answered ' + r.status);
        return r.text();
      })
      .then(html => {
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const sections = LACPagePrint.extractSections(doc);
        if (!sections.length) throw new Error('no tables were found on that page');
        const bytes = LACPagePrint.renderPdf(sections, { title, subtitle, source: 'civil.lacourt.org' });
        show(new Blob([bytes], { type: 'application/pdf' }));
      })
      .catch(err => failed('render that page', err));
  } else {
    fail('No court page to show.');
  }
})();

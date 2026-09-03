/**
 * Renders one eCourt case sub-tab as a PDF in its own tab.
 *
 * Opened by the Documents button for the Parties tab, which eCourt gives no
 * print endpoint of its own. The page is fetched here rather than read off the
 * live tab, so what gets rendered is the server's own HTML — the extension's
 * buttons, header rewrites and status spans never existed in it.
 *
 * Query string: url (the case sub-tab to render), title, subtitle, pdfname.
 * The save-as name is `pdfname` rather than `file` because the background's
 * resolveOpenPdfUrl reads everything after a `file=` as a document URL.
 */
(function () {
  'use strict';

  const params = new URLSearchParams(location.search);
  const pageUrl = params.get('url') || '';
  const title = params.get('title') || 'eCourt page';
  const subtitle = params.get('subtitle') || '';
  const fileName = (params.get('pdfname') || title).replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();

  const $ = id => document.getElementById(id);
  document.title = title;
  $('name').textContent = title;

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
      a.textContent = 'the page itself';
      status.appendChild(a);
      status.appendChild(document.createTextNode(
        ' to check you are still signed in to eCourt, then click Documents again.'));
    }
    status.hidden = false;
    $('frame').hidden = true;
  }

  if (!/^https:\/\/[^/]*lacourt\.org\//i.test(pageUrl)) {
    fail('No court page to render.');
    return;
  }

  fetch(pageUrl, { credentials: 'include' })
    .then(r => {
      if (!r.ok) throw new Error('the court site answered ' + r.status);
      return r.text();
    })
    .then(html => {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const sections = LACPagePrint.extractSections(doc);
      if (!sections.length) throw new Error('no tables were found on that page');
      const bytes = LACPagePrint.renderPdf(sections, {
        title,
        subtitle,
        source: 'civil.lacourt.org',
      });
      const url = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }));
      const frame = $('frame');
      frame.src = url;
      frame.hidden = false;
      $('status').hidden = true;
      const save = $('save');
      save.href = url;
      save.setAttribute('download', fileName + '.pdf');
      save.hidden = false;
    })
    .catch(err => {
      // The usual cause is a signed-out session: the fetch carries the tab's
      // cookies, so if the court site has logged the user out it answers with
      // the login page instead of the case.
      fail('Could not render that page — ' + String(err && err.message || err) + '.', pageUrl);
      console.error('[LACourt-PagePDF]', err);
    });
})();

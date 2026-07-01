/**
 * LA Court E-Court Suite - Order Template Input popup
 *
 * Replaces the old "Order Template Input" Microsoft Form. When the user hits
 * "Fill Microsoft Form" on a non-Default-Judgment case, the case content
 * script parses the party data, stores it under chrome.storage.local
 * `orderTemplateData`, and opens this page in its own window (on the opposite
 * display). Here we render an editable, pre-filled box per question, let the
 * user tweak values and add the Movant by hand, then download a spreadsheet.
 *
 * The download is a real .xlsx whose columns match, byte-for-byte in header
 * text and order, the export the Microsoft Form used to produce — so the
 * user's existing Word mail merge keeps working with no changes:
 *
 *   ID | Start time | Completion time | Email | Name | Case Number  |
 *   Hearing Date | Motion Type | Title Plaintiff | Other Plaintiffs |
 *   Title Defendant | Other Defendants | CrossComplainants |
 *   CrossDefendants | Movant
 *
 * The five leading metadata columns are written blank (the Form filled them
 * with response metadata that the order template doesn't use); every question
 * column is filled from the parsed/edited values.
 */

// Question fields, in export order. `key` matches the storage/labeled key from
// the content script; `header` is the EXACT spreadsheet column header (note the
// trailing space on "Case Number " and the no-hyphen cross-* headers — these
// must match the merge source the user's template expects). `movant` has no
// parsed value; the user fills it in.
const FIELDS = [
  { key: 'caseNumber',        label: 'Case Number',        header: 'Case Number ',    multiline: false },
  { key: 'hearingDate',       label: 'Hearing Date',       header: 'Hearing Date',    multiline: false },
  { key: 'motionType',        label: 'Motion Type',        header: 'Motion Type',     multiline: false },
  { key: 'titlePlaintiff',    label: 'Title Plaintiff',    header: 'Title Plaintiff', multiline: false },
  { key: 'otherPlaintiffs',   label: 'Other Plaintiffs',   header: 'Other Plaintiffs', multiline: true },
  { key: 'titleDefendant',    label: 'Title Defendant',    header: 'Title Defendant', multiline: false },
  { key: 'otherDefendants',   label: 'Other Defendants',   header: 'Other Defendants', multiline: true },
  { key: 'crossComplainants', label: 'Cross-Complainants', header: 'CrossComplainants', multiline: true },
  { key: 'crossDefendants',   label: 'Cross-Defendants',   header: 'CrossDefendants', multiline: true },
  { key: 'movant',            label: 'Movant',             header: 'Movant',          multiline: false },
];

// Leading metadata columns from the Microsoft Forms export. Written blank.
const META_HEADERS = ['ID', 'Start time', 'Completion time', 'Email', 'Name'];

// Native messaging host that launches the Word mail-merge template. Must match
// the "name" in the installed host manifest (see native-host/).
const NATIVE_HOST = 'com.lacourt.ecourt_host';

const statusEl = document.getElementById('status');
const fieldsEl = document.getElementById('fields');

function setStatus(text, type) {
  statusEl.textContent = text || '';
  statusEl.className = 'status' + (type ? ' ' + type : '');
}

// -------------------------------------------------------------------------
// Render + populate
// -------------------------------------------------------------------------

function renderFields(values) {
  const frag = document.createDocumentFragment();

  for (const f of FIELDS) {
    const wrap = document.createElement('div');
    wrap.className = 'field' + (f.key === 'movant' ? ' movant' : '');

    const label = document.createElement('label');
    label.textContent = f.label;
    label.setAttribute('for', 'f_' + f.key);

    const input = f.multiline
      ? document.createElement('textarea')
      : document.createElement('input');
    if (!f.multiline) input.type = 'text';
    input.id = 'f_' + f.key;
    input.dataset.key = f.key;
    input.value = values && values[f.key] != null ? String(values[f.key]) : '';

    wrap.appendChild(label);
    wrap.appendChild(input);
    frag.appendChild(wrap);
  }

  fieldsEl.textContent = '';
  fieldsEl.appendChild(frag);
}

function collectValues() {
  const out = {};
  for (const f of FIELDS) {
    const el = document.getElementById('f_' + f.key);
    out[f.key] = el ? el.value : '';
  }
  return out;
}

function loadData() {
  // Read once; retry a single time shortly after in case the popup window
  // beat the content script's storage write (belt-and-suspenders — the
  // content script already waits for the write before opening us).
  const attempt = (retriesLeft) => {
    chrome.storage.local.get(['orderTemplateData'], result => {
      const rec = result && result.orderTemplateData;
      const fields = rec && rec.fields;
      if (!fields && retriesLeft > 0) {
        setTimeout(() => attempt(retriesLeft - 1), 150);
        return;
      }
      renderFields(fields || {});
    });
  };
  attempt(1);
}

// -------------------------------------------------------------------------
// .xlsx generation (minimal, dependency-free)
// -------------------------------------------------------------------------

function xmlEscape(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function buildSheetXml(headers, dataRow) {
  const cell = (colIdx, rowNum, val) =>
    `<c r="${colLetter(colIdx)}${rowNum}" t="inlineStr">` +
    `<is><t xml:space="preserve">${xmlEscape(val)}</t></is></c>`;

  let headerCells = '';
  headers.forEach((h, i) => { headerCells += cell(i + 1, 1, h); });

  let dataCells = '';
  dataRow.forEach((v, i) => {
    if (v !== '' && v != null) dataCells += cell(i + 1, 2, String(v));
  });

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData>' +
    `<row r="1">${headerCells}</row>` +
    `<row r="2">${dataCells}</row>` +
    '</sheetData></worksheet>';
}

const XLSX_PARTS = {
  '[Content_Types].xml':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '</Types>',
  '_rels/.rels':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>',
  'xl/workbook.xml':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
    '</workbook>',
  'xl/_rels/workbook.xml.rels':
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '</Relationships>',
};

// --- tiny ZIP writer (STORE / no compression) --------------------------------

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xFF];
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function zipStore(files) {
  const enc = new TextEncoder();
  const locals = [];   // Uint8Array chunks in file order
  const central = [];  // central directory records
  let offset = 0;

  for (const f of files) {
    const nameBytes = enc.encode(f.name);
    const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
    const crc = crc32(data);
    const size = data.length;

    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); // local file header signature
    lv.setUint16(4, 20, true);         // version needed
    lv.setUint16(6, 0, true);          // flags
    lv.setUint16(8, 0, true);          // method: 0 = store
    lv.setUint16(10, 0, true);         // mod time
    lv.setUint16(12, 0x21, true);      // mod date = 1980-01-01
    lv.setUint32(14, crc, true);
    lv.setUint32(18, size, true);      // compressed size
    lv.setUint32(22, size, true);      // uncompressed size
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);         // extra length
    local.set(nameBytes, 30);
    locals.push(local, data);

    const cen = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(cen.buffer);
    cv.setUint32(0, 0x02014b50, true); // central dir signature
    cv.setUint16(4, 20, true);         // version made by
    cv.setUint16(6, 20, true);         // version needed
    cv.setUint16(8, 0, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0x21, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, size, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);    // relative offset of local header
    cen.set(nameBytes, 46);
    central.push(cen);

    offset += local.length + size;
  }

  const centralSize = central.reduce((a, c) => a + c.length, 0);
  const centralOffset = offset;

  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, centralOffset, true);
  ev.setUint16(20, 0, true);

  const all = [...locals, ...central, eocd];
  const total = all.reduce((a, c) => a + c.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const chunk of all) { out.set(chunk, pos); pos += chunk.length; }
  return out;
}

function buildWorkbook(values) {
  const headers = META_HEADERS.concat(FIELDS.map(f => f.header));
  const dataRow = META_HEADERS.map(() => '')
    .concat(FIELDS.map(f => values[f.key] || ''));

  const files = [];
  for (const name in XLSX_PARTS) {
    files.push({ name, data: XLSX_PARTS[name] });
  }
  files.push({ name: 'xl/worksheets/sheet1.xml', data: buildSheetXml(headers, dataRow) });

  return zipStore(files);
}

function sanitizeForFilename(s) {
  return String(s || '').trim().replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function filenameFor(values) {
  // The Word macro finds the sheet with Dir("Order*.xlsx") in Downloads, so
  // the name MUST start with "Order" and land in the Downloads root.
  const caseTag = sanitizeForFilename(values.caseNumber);
  return caseTag
    ? `Order_Template_Input_${caseTag}.xlsx`
    : 'Order_Template_Input.xlsx';
}

function workbookBlob(values) {
  const bytes = buildWorkbook(values);
  return new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

// Preferred save path: chrome.downloads into the Downloads root, resolving only
// once the file is fully written — so the native mail-merge trigger can't fire
// before the spreadsheet exists on disk. Rejects if the API errors.
function downloadViaApi(blob, filename) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    chrome.downloads.download(
      { url, filename, saveAs: false, conflictAction: 'uniquify' },
      downloadId => {
        if (chrome.runtime.lastError || downloadId == null) {
          URL.revokeObjectURL(url);
          reject(new Error(chrome.runtime.lastError
            ? chrome.runtime.lastError.message : 'download failed'));
          return;
        }
        const onChanged = delta => {
          if (delta.id !== downloadId || !delta.state) return;
          if (delta.state.current === 'complete') {
            chrome.downloads.onChanged.removeListener(onChanged);
            URL.revokeObjectURL(url);
            resolve(downloadId);
          } else if (delta.state.current === 'interrupted') {
            chrome.downloads.onChanged.removeListener(onChanged);
            URL.revokeObjectURL(url);
            reject(new Error('download interrupted'));
          }
        };
        chrome.downloads.onChanged.addListener(onChanged);
      }
    );
  });
}

// Fallback save path (anchor click) if chrome.downloads is unavailable. No
// completion signal, so callers wait a beat before triggering the merge.
function downloadViaAnchor(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Ask the native host to launch the Word template, which fires the macro's
// Document_New handler. Never throws — resolves to {ok:false,error} when the
// host isn't installed, so the download itself always succeeds regardless.
function triggerNativeMailMerge() {
  return new Promise(resolve => {
    if (!chrome.runtime || !chrome.runtime.sendNativeMessage) {
      resolve({ ok: false, error: 'native messaging unavailable' });
      return;
    }
    try {
      chrome.runtime.sendNativeMessage(
        NATIVE_HOST, { action: 'launchTemplate' },
        response => {
          if (chrome.runtime.lastError) {
            resolve({ ok: false, error: chrome.runtime.lastError.message });
            return;
          }
          resolve(response || { ok: false, error: 'no response from host' });
        }
      );
    } catch (e) {
      resolve({ ok: false, error: String(e && e.message || e) });
    }
  });
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// -------------------------------------------------------------------------
// "Did you pull the case PDFs?" guard
//
// Before exporting we check that at least one PDF has been downloaded from the
// court website since the previous export. If none has, the mail merge would
// have no documents to file into the case folder, so we warn (overridable).
// The window mirrors the Word macro's PDF band: "since my last export", with a
// since-midnight fallback the very first time.
// -------------------------------------------------------------------------

const COURT_DOMAIN_RE = /lacourt\.org/i;

function startOfTodayISO() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d.toISOString();
}

function getLastExportAt() {
  return new Promise(resolve => {
    try {
      chrome.storage.local.get(['lastOrderExportAt'], r =>
        resolve(r && r.lastOrderExportAt ? r.lastOrderExportAt : null));
    } catch (_) {
      resolve(null);
    }
  });
}

function setLastExportAt(ts) {
  try {
    chrome.storage.local.set({ lastOrderExportAt: ts });
  } catch (_) { /* non-critical */ }
}

function isPdfDownload(item) {
  if (item.mime && item.mime.toLowerCase() === 'application/pdf') return true;
  if (item.filename && /\.pdf$/i.test(item.filename)) return true;
  const u = item.finalUrl || item.url || '';
  return /\.pdf(?:[?#]|$)/i.test(u);
}

function isFromCourt(item) {
  // blob: downloads keep the origin in the URL, so this catches those too.
  return COURT_DOMAIN_RE.test(item.url || '')
      || COURT_DOMAIN_RE.test(item.finalUrl || '')
      || COURT_DOMAIN_RE.test(item.referrer || '');
}

// Resolves true if a court PDF was downloaded within the window, OR if we
// simply can't tell (no downloads API) — we never block on uncertainty.
async function hasRecentCourtPdf() {
  if (!chrome.downloads || !chrome.downloads.search) return true;

  const last = await getLastExportAt();
  const sinceISO = last ? new Date(last).toISOString() : startOfTodayISO();

  const items = await new Promise(resolve => {
    try {
      chrome.downloads.search(
        { state: 'complete', startedAfter: sinceISO, orderBy: ['-startTime'], limit: 0 },
        res => {
          if (chrome.runtime.lastError) { resolve([]); return; }
          resolve(res || []);
        }
      );
    } catch (_) {
      resolve([]);
    }
  });

  return items.some(it => isPdfDownload(it) && isFromCourt(it));
}

// -------------------------------------------------------------------------
// Wire up
// -------------------------------------------------------------------------

const downloadBtn = document.getElementById('downloadBtn');
const pdfWarning = document.getElementById('pdfWarning');

function showPdfWarning() {
  if (pdfWarning) pdfWarning.hidden = false;
  setStatus('No court PDF found since your last export.', 'error');
}

function hidePdfWarning() {
  if (pdfWarning) pdfWarning.hidden = true;
}

// Does the actual export: build → save → record time → trigger mail merge.
async function performExport() {
  hidePdfWarning();

  const values = collectValues();
  const filename = filenameFor(values);

  let blob;
  try {
    blob = workbookBlob(values);
  } catch (err) {
    console.error('[OrderTemplate] build failed:', err);
    setStatus('Build failed: ' + (err && err.message || err), 'error');
    return;
  }

  const autoRunEl = document.getElementById('autoRun');
  const wantAutoRun = autoRunEl ? autoRunEl.checked : true;

  downloadBtn.disabled = true;
  setStatus('Saving ' + filename + '…');

  let usedApi = false;
  try {
    if (chrome.downloads && chrome.downloads.download) {
      await downloadViaApi(blob, filename);
      usedApi = true;
    } else {
      downloadViaAnchor(blob, filename);
    }
  } catch (err) {
    // Last-ditch: fall back to the anchor download so a save never regresses.
    console.warn('[OrderTemplate] downloads API failed, using anchor:', err);
    try {
      downloadViaAnchor(blob, filename);
    } catch (e2) {
      setStatus('Download failed: ' + (err && err.message || err), 'error');
      downloadBtn.disabled = false;
      return;
    }
  }

  // Anchor this export so the next run's "since last export" PDF check starts
  // from here. Recorded whether or not the mail merge runs.
  setLastExportAt(Date.now());

  if (!wantAutoRun) {
    setStatus('Downloaded ' + filename, 'success');
    downloadBtn.disabled = false;
    return;
  }

  // If we used the anchor fallback there's no completion signal; give the
  // browser a moment to finish writing before the macro reads the file.
  if (!usedApi) await sleep(1500);

  setStatus('Downloaded. Starting mail merge…');
  const res = await triggerNativeMailMerge();
  if (res && res.ok) {
    setStatus('Downloaded ' + filename + ' — mail merge started.', 'success');
  } else {
    // Download succeeded; only the auto-run bridge is missing. Non-alarming.
    setStatus('Downloaded ' + filename + '. Auto mail-merge did not run ('
      + (res && res.error ? res.error : 'host not set up') + ').');
  }
  downloadBtn.disabled = false;
}

// Download button first runs the "did you pull the PDFs?" guard.
downloadBtn.addEventListener('click', async () => {
  hidePdfWarning();
  downloadBtn.disabled = true;
  setStatus('Checking recent downloads…');

  let ok = true;
  try {
    ok = await hasRecentCourtPdf();
  } catch (err) {
    console.warn('[OrderTemplate] PDF check failed, allowing export:', err);
    ok = true; // never block on a check failure
  }

  downloadBtn.disabled = false;
  if (!ok) {
    showPdfWarning();
    return;
  }
  await performExport();
});

// Warning banner buttons.
document.getElementById('pdfDownloadAnyway').addEventListener('click', () => {
  performExport();
});
document.getElementById('pdfDismiss').addEventListener('click', () => {
  hidePdfWarning();
  setStatus('');
});

document.getElementById('closeBtn').addEventListener('click', () => {
  window.close();
});

loadData();

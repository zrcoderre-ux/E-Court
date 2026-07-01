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

function download(values) {
  const bytes = buildWorkbook(values);
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);

  const caseTag = sanitizeForFilename(values.caseNumber);
  const filename = caseTag
    ? `Order_Template_Input_${caseTag}.xlsx`
    : 'Order_Template_Input.xlsx';

  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  return filename;
}

// -------------------------------------------------------------------------
// Wire up
// -------------------------------------------------------------------------

document.getElementById('downloadBtn').addEventListener('click', () => {
  try {
    const name = download(collectValues());
    setStatus('Downloaded ' + name, 'success');
  } catch (err) {
    console.error('[OrderTemplate] download failed:', err);
    setStatus('Download failed: ' + (err && err.message || err), 'error');
  }
});

document.getElementById('closeBtn').addEventListener('click', () => {
  window.close();
});

loadData();

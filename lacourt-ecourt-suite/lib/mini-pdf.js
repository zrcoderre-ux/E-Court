/**
 * Minimal PDF writer.
 *
 * The extension has no build step and no vendored libraries, so this is a
 * from-scratch PDF 1.4 generator: enough to lay out ruled text tables on
 * Letter pages, and nothing more. It exists because eCourt gives the Parties
 * tab no print endpoint of its own — the only native way to a PDF is Ctrl+P,
 * which cannot run unattended — so the extension renders that page itself.
 *
 * Only the two base-14 fonts are used (Helvetica and Helvetica-Bold), which
 * every PDF reader carries, so nothing has to be embedded. Text is written in
 * WinAnsi, the encoding those base fonts declare.
 *
 * Coordinates are given TOP-LEFT in points and flipped on the way out; PDF's
 * own origin is bottom-left and reasoning about a table upside down is a
 * needless source of mistakes.
 */
var LACPdf = (function () {
  'use strict';

  const PT_PER_IN = 72;
  const LETTER = { w: 8.5 * PT_PER_IN, h: 11 * PT_PER_IN };

  /* Adobe's published widths for the base-14 Helvetica faces, in 1/1000 em,
     for the printable ASCII range. Anything outside it measures as the average
     — width is only used for wrapping and column fitting, so a rare character
     measuring slightly wrong costs a little whitespace, never a broken file. */
  const W_REG = [
    278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,          // 32-47
    556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,          // 48-63
    1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,         // 64-79
    667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,          // 80-95
    333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,          // 96-111
    556,556,333,500,278,556,500,722,500,500,500,334,260,334,584,              // 112-126
  ];
  const W_BOLD = [
    278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,
    556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,
    975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,
    667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,
    333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,
    611,611,389,556,333,611,556,778,556,556,500,389,280,389,584,
  ];
  const AVG = 556;

  // The characters an eCourt page actually carries beyond ASCII, mapped to the
  // WinAnsi byte that draws them. Everything else unmapped degrades to a plain
  // ASCII stand-in rather than a wrong glyph.
  const WINANSI = {
    '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94,
    '–': 0x96, '—': 0x97, '•': 0x95, '…': 0x85,
    '§': 0xA7, '¶': 0xB6, ' ': 0x20, '™': 0x99,
    'é': 0xE9, 'ñ': 0xF1, 'ü': 0xFC, 'á': 0xE1,
    'í': 0xED, 'ó': 0xF3, 'à': 0xE0, 'è': 0xE8,
  };

  // To a string whose every char code is a byte, so the file's byte offsets —
  // which the xref table has to name exactly — equal its string indices.
  function toWinAnsi(s) {
    let out = '';
    for (const ch of String(s == null ? '' : s)) {
      const c = ch.codePointAt(0);
      if (c >= 0x20 && c <= 0x7e) { out += ch; continue; }
      if (ch === '\t') { out += ' '; continue; }
      const mapped = WINANSI[ch];
      if (mapped != null) { out += String.fromCharCode(mapped); continue; }
      if (c >= 0xa0 && c <= 0xff) { out += ch; continue; }
      out += '?';
    }
    return out;
  }

  function widthOf(text, size, bold) {
    const tbl = bold ? W_BOLD : W_REG;
    let total = 0;
    const s = toWinAnsi(text);
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      total += (c >= 32 && c <= 126) ? tbl[c - 32] : AVG;
    }
    return total * size / 1000;
  }

  // Break `text` to fit `maxWidth`, splitting inside a word only when the word
  // alone cannot fit (docket titles run together and party names carry long
  // corporate strings, so a hard break has to be available).
  function wrap(text, maxWidth, size, bold) {
    const words = String(text == null ? '' : text).split(/\s+/).filter(Boolean);
    if (!words.length) return [''];
    const lines = [];
    let line = '';
    const push = () => { if (line) { lines.push(line); line = ''; } };
    for (let w of words) {
      const candidate = line ? line + ' ' + w : w;
      if (widthOf(candidate, size, bold) <= maxWidth) { line = candidate; continue; }
      push();
      while (widthOf(w, size, bold) > maxWidth && w.length > 1) {
        let cut = w.length;
        while (cut > 1 && widthOf(w.slice(0, cut), size, bold) > maxWidth) cut--;
        lines.push(w.slice(0, cut));
        w = w.slice(cut);
      }
      line = w;
    }
    push();
    return lines.length ? lines : [''];
  }

  function escapeText(s) {
    return toWinAnsi(s).replace(/([\\()])/g, '\\$1');
  }

  function fmt(n) {
    // Three decimals is finer than any printer resolves, and keeps the content
    // stream free of float noise like 72.00000000000001.
    const r = Math.round(n * 1000) / 1000;
    return String(Object.is(r, -0) ? 0 : r);
  }

  /**
   * @param {{landscape?:boolean, margin?:number}} opts
   */
  function MiniPdf(opts) {
    opts = opts || {};
    this.landscape = !!opts.landscape;
    this.width = this.landscape ? LETTER.h : LETTER.w;
    this.height = this.landscape ? LETTER.w : LETTER.h;
    this.margin = opts.margin != null ? opts.margin : 36;
    this.pages = [];
    this.addPage();
  }

  MiniPdf.prototype.addPage = function () {
    this.ops = [];
    this.pages.push(this.ops);
    return this;
  };

  MiniPdf.prototype.pageCount = function () { return this.pages.length; };

  /* y is the text BASELINE, measured from the top of the page.

     The fill colour is set on every run, never assumed. PDF carries one
     graphics state forward across the whole content stream, and text is painted
     in the FILL colour — so a shaded table header, which sets the fill to a
     light grey to paint its band, silently recolours every word drawn after it.
     Stating the colour here costs four bytes a line and cannot be forgotten. */
  MiniPdf.prototype.text = function (str, x, y, size, bold, gray) {
    const s = escapeText(str);
    if (!s) return this;
    this.ops.push(fmt(gray == null ? 0 : gray) + ' g BT /' + (bold ? 'F2' : 'F1') + ' ' + fmt(size)
      + ' Tf 1 0 0 1 ' + fmt(x) + ' ' + fmt(this.height - y) + ' Tm (' + s + ') Tj ET');
    return this;
  };

  MiniPdf.prototype.line = function (x1, y1, x2, y2, thickness, gray) {
    this.ops.push(fmt(gray == null ? 0 : gray) + ' G ' + fmt(thickness == null ? 0.5 : thickness) + ' w '
      + fmt(x1) + ' ' + fmt(this.height - y1) + ' m '
      + fmt(x2) + ' ' + fmt(this.height - y2) + ' l S');
    return this;
  };

  MiniPdf.prototype.fillRect = function (x, y, w, h, gray) {
    this.ops.push(fmt(gray) + ' g ' + fmt(x) + ' ' + fmt(this.height - y - h) + ' '
      + fmt(w) + ' ' + fmt(h) + ' re f');
    return this;
  };

  /* Serialize. Objects are numbered in the order written and the xref table
     records each one's byte offset, so the document is built as a byte-per-char
     string and converted at the end. */
  MiniPdf.prototype.build = function () {
    const objs = [];                       // 1-based body objects
    const add = body => { objs.push(body); return objs.length; };

    const pageIds = [];
    const contentIds = [];
    for (const ops of this.pages) {
      const stream = ops.join('\n');
      contentIds.push(add('<< /Length ' + stream.length + ' >>\nstream\n' + stream + '\nendstream'));
      pageIds.push(0); // placeholder, filled once the Pages object has an id
    }
    const fontReg = add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>');
    const fontBold = add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>');
    const pagesId = objs.length + this.pages.length + 1;
    for (let i = 0; i < this.pages.length; i++) {
      pageIds[i] = add('<< /Type /Page /Parent ' + pagesId + ' 0 R /MediaBox [0 0 '
        + fmt(this.width) + ' ' + fmt(this.height) + '] /Resources << /Font << /F1 '
        + fontReg + ' 0 R /F2 ' + fontBold + ' 0 R >> >> /Contents ' + contentIds[i] + ' 0 R >>');
    }
    const realPagesId = add('<< /Type /Pages /Kids [' + pageIds.map(id => id + ' 0 R').join(' ')
      + '] /Count ' + pageIds.length + ' >>');
    const catalogId = add('<< /Type /Catalog /Pages ' + realPagesId + ' 0 R >>');

    let out = '%PDF-1.4\n%\xE2\xE3\xCF\xD3\n';
    const offsets = [0];
    for (let i = 0; i < objs.length; i++) {
      offsets.push(out.length);
      out += (i + 1) + ' 0 obj\n' + objs[i] + '\nendobj\n';
    }
    const xref = out.length;
    out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n';
    for (let i = 1; i <= objs.length; i++) {
      out += String(offsets[i]).padStart(10, '0') + ' 00000 n \n';
    }
    out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root ' + catalogId
      + ' 0 R >>\nstartxref\n' + xref + '\n%%EOF\n';

    const bytes = new Uint8Array(out.length);
    for (let i = 0; i < out.length; i++) bytes[i] = out.charCodeAt(i) & 0xff;
    return bytes;
  };

  return { MiniPdf, widthOf, wrap, toWinAnsi, PT_PER_IN, LETTER };
})();

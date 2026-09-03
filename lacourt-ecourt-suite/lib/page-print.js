/**
 * Render an eCourt case sub-tab as a PDF.
 *
 * eCourt gives the Parties tab no print endpoint — Ctrl+P is the only native
 * route to a PDF of it, and a print dialog cannot run unattended. So the tab is
 * fetched fresh, read down to its tables, and laid out here.
 *
 * Fetching rather than reading the live DOM is deliberate, and is what keeps
 * the extension's own additions out of the result: a background fetch returns
 * the page as the server wrote it, with no floating buttons, no rewritten
 * "Next" header, no injected status spans — there is nothing to strip because
 * nothing was ever added.
 *
 * What gets rendered is whatever tables the page carries, in page order, under
 * whatever headings the page gives them. Nothing is hard-coded to "Parties" or
 * "Representation": the page decides what sections it has (a case with no
 * former representation simply has no such table), and this reproduces them.
 */
var LACPagePrint = (function () {
  'use strict';

  const { MiniPdf, widthOf, wrap } = LACPdf;

  /* ---------------------------------------------------------------- */
  /* Reading the page                                                  */
  /* ---------------------------------------------------------------- */

  const cellText = el => (el.textContent || '').replace(/\s+/g, ' ').trim();

  // eCourt is 1990s HTML: the page frame, the sub-nav and the surrounding
  // furniture are all <table>s wrapping the real ones. A table that contains
  // another table is layout, never data — so only the leaves are read.
  function isLeafTable(t) { return !t.querySelector('table'); }

  // The heading a table sits under: the nearest preceding heading element or
  // bold/caption-ish line. eCourt labels its sections with plain styled text
  // ("Parties", "Representation", "Former Representation") as often as with a
  // real <h*>, so both are accepted.
  const HEADING_SEL = 'h1,h2,h3,h4,h5,h6,caption,legend,b,strong,.sectionHeader,.SectionHeader,.tableHeader';
  function headingFor(table) {
    const cap = table.querySelector(':scope > caption');
    if (cap && cellText(cap)) return cellText(cap);
    let node = table;
    for (let hops = 0; node && hops < 6; hops++) {
      let sib = node.previousElementSibling;
      while (sib) {
        if (sib.matches && sib.matches(HEADING_SEL)) {
          const t = cellText(sib);
          if (t && t.length <= 80) return t;
        }
        const inner = sib.querySelectorAll ? Array.from(sib.querySelectorAll(HEADING_SEL)) : [];
        for (let i = inner.length - 1; i >= 0; i--) {
          const t = cellText(inner[i]);
          if (t && t.length <= 80) return t;
        }
        const own = cellText(sib);
        if (own && own.length <= 80 && !/^\s*$/.test(own)) return own;
        sib = sib.previousElementSibling;
      }
      node = node.parentElement;
    }
    return '';
  }

  // A row's cells as { text, span }, honouring colspan so a heading row that
  // spans the table ("Cross-Complaint filed by … on 04/12/2026" — the parties
  // table interleaves those with the party rows) stays one wide cell.
  function readRow(tr) {
    const cells = [];
    for (const td of tr.children) {
      if (!/^t[dh]$/i.test(td.tagName)) continue;
      const span = Math.max(1, parseInt(td.getAttribute('colspan') || '1', 10) || 1);
      const a = td.querySelector('a');
      cells.push({
        text: cellText(td), span, header: /^th$/i.test(td.tagName),
        // A cell that is nothing but a link — what the case sub-nav is made of.
        link: !!a && !!cellText(a) && cellText(a) === cellText(td),
      });
    }
    return cells;
  }

  function rowWidth(cells) { return cells.reduce((n, c) => n + c.span, 0); }

  /* Is this leaf table worth printing? The page carries one-cell spacer tables,
     a banner of logo-plus-court-name, and the case sub-nav's strip of tab
     links; none of them are the record, and all three are <table>s like
     everything else on a page of this vintage. Two tests clear them out:

       - a section has at least TWO rows carrying a row's worth of text (its
         headings and an entry). The banner has one such row at most and the
         sub-nav is a single row of links above an empty one;
       - a table that is almost entirely links is navigation. No section of the
         record is. */
  function isDataTable(rows) {
    if (rows.length < 2) return false;
    const widest = Math.max(...rows.map(rowWidth));
    if (widest < 2) return false;
    const substantial = rows.filter(r => r.filter(c => c.text).length >= 2).length;
    if (substantial < 2) return false;
    const filled = rows.reduce((n, r) => n + r.filter(c => c.text).length, 0);
    const links = rows.reduce((n, r) => n + r.filter(c => c.text && c.link).length, 0);
    return !(filled && links / filled >= 0.6);
  }

  /* eCourt marks a column-heading row up as <th> on some tabs and as styled
     <td> on others, so the row has to be recognized by what it looks like too:
     every cell filled, every cell short, and no cell holding a value (a date is
     the giveaway) that a heading would never be. */
  function looksLikeHeaderRow(row) {
    return row.length > 1
      && row.every(c => c.text && c.text.length <= 40 && c.span === 1)
      && !row.some(c => /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(c.text));
  }

  /**
   * Every printable section of a fetched page, in page order.
   * @returns {Array<{title:string, columns:string[], rows:Array<Array<{text,span}>>}>}
   */
  function extractSections(doc) {
    const sections = [];
    const seen = new Set();
    for (const table of doc.querySelectorAll('table')) {
      if (!isLeafTable(table)) continue;
      const rows = Array.from(table.querySelectorAll('tr')).map(readRow).filter(r => r.length);
      if (!isDataTable(rows)) continue;

      const width = Math.max(...rows.map(rowWidth));
      // The header row is the first row that spans the table's full width and
      // is either marked up as <th> or is the row the rest of the data follows.
      let columns = [];
      let body = rows;
      const first = rows[0];
      if (rowWidth(first) === width && (first.some(c => c.header) || looksLikeHeaderRow(first))) {
        columns = first.map(c => c.text);
        body = rows.slice(1);
      }
      if (!body.length) continue;

      // eCourt repeats the same table markup in more than one wrapper on some
      // tabs; a section whose header and first row are already rendered is that
      // repeat, not a second section.
      const key = columns.join('|') + '||' + body[0].map(c => c.text).join('|');
      if (seen.has(key)) continue;
      seen.add(key);

      sections.push({ title: headingFor(table), columns, rows: body, width });
    }
    return sections;
  }

  /* ---------------------------------------------------------------- */
  /* Laying it out                                                     */
  /* ---------------------------------------------------------------- */

  const FONT = 8.5;
  const HEAD_FONT = 9;
  const TITLE_FONT = 13;
  const SECTION_FONT = 10.5;
  const LINE_H = 11;
  const CELL_PAD = 4;

  // Column widths proportional to what the column actually holds, so a "Role"
  // column does not get the same inch as a party name. Measured on the widest
  // single word too, so a column is never squeezed narrower than one word.
  function columnWidths(section, avail) {
    const n = section.width;
    const want = new Array(n).fill(0);
    const floor = new Array(n).fill(0);
    const measure = (cells, bold) => {
      let col = 0;
      for (const c of cells) {
        if (c.span === 1 && col < n) {
          want[col] = Math.max(want[col], widthOf(c.text, FONT, bold) + CELL_PAD * 2);
          const longest = c.text.split(/\s+/).reduce((m, w) => Math.max(m, widthOf(w, FONT, bold)), 0);
          floor[col] = Math.max(floor[col], longest + CELL_PAD * 2);
        }
        col += c.span;
      }
    };
    if (section.columns.length) measure(section.columns.map(t => ({ text: t, span: 1 })), true);
    for (const r of section.rows) measure(r, false);

    const cap = avail * 0.42;
    for (let i = 0; i < n; i++) want[i] = Math.max(24, Math.min(want[i] || 24, cap));
    let total = want.reduce((a, b) => a + b, 0);
    if (!total) return new Array(n).fill(avail / n);

    // Scale to the page. Columns are never taken below what one word needs, so
    // shrinking falls on the roomy columns; if even the floors overflow, the
    // whole row scales and the wrapper breaks words.
    if (total > avail) {
      const floorTotal = floor.reduce((a, b) => a + Math.min(b, cap), 0);
      if (floorTotal < avail) {
        const slack = avail - floorTotal;
        const extra = total - floorTotal;
        for (let i = 0; i < n; i++) {
          const f = Math.min(floor[i], cap);
          want[i] = f + (extra > 0 ? (want[i] - f) * slack / extra : 0);
        }
      } else {
        for (let i = 0; i < n; i++) want[i] *= avail / total;
      }
    } else {
      for (let i = 0; i < n; i++) want[i] *= avail / total;
    }
    return want;
  }

  /**
   * @param {Array} sections  from extractSections
   * @param {{title:string, subtitle:string, source:string}} meta
   * @returns {Uint8Array}
   */
  function renderPdf(sections, meta) {
    meta = meta || {};
    // A parties table runs six or seven columns; portrait Letter turns that into
    // a column of single words. Anything wider than four columns goes landscape.
    const widest = sections.reduce((m, s) => Math.max(m, s.width), 0);
    const pdf = new MiniPdf({ landscape: widest > 4, margin: 36 });
    const M = pdf.margin;
    const avail = pdf.width - M * 2;
    const bottom = pdf.height - M - 14;
    let y = M + TITLE_FONT;

    pdf.text(meta.title || 'eCourt', M, y, TITLE_FONT, true);
    y += 6;
    if (meta.subtitle) { y += LINE_H; pdf.text(meta.subtitle, M, y, FONT, false); }
    y += 6;
    pdf.line(M, y, pdf.width - M, y, 1, 0.35);
    y += 8;

    let pageNo = 1;
    const footer = () => {
      const stamp = (meta.source ? meta.source + '  ·  ' : '') + 'page ' + pageNo;
      pdf.text(stamp, M, pdf.height - M + 8, 7.5, false, 0.45);
    };
    const newPage = () => {
      footer();
      pdf.addPage();
      pageNo++;
      y = M + FONT;
    };

    for (const section of sections) {
      const widths = columnWidths(section, avail);
      const xs = [];
      let x = M;
      for (const w of widths) { xs.push(x); x += w; }
      const right = x;

      const drawHeader = () => {
        if (!section.columns.length) return;
        const lines = section.columns.map((t, i) => wrap(t, widths[i] - CELL_PAD * 2, HEAD_FONT, true));
        const h = Math.max(...lines.map(l => l.length)) * LINE_H + 4;
        pdf.fillRect(M, y, right - M, h, 0.9);
        lines.forEach((ls, i) => {
          ls.forEach((ln, k) => pdf.text(ln, xs[i] + CELL_PAD, y + LINE_H * (k + 1) - 2, HEAD_FONT, true));
        });
        y += h;
        pdf.line(M, y, right, y, 0.8, 0.2);
      };

      if (y + 40 > bottom) newPage();
      if (section.title) {
        y += SECTION_FONT + 4;
        pdf.text(section.title, M, y, SECTION_FONT, true);
        y += 5;
      }
      drawHeader();

      for (const row of section.rows) {
        // A lone cell in a multi-column table is a divider row, not data.
        const spanning = row.length === 1 && section.width > 1;
        let lineSets, height;
        if (spanning) {
          lineSets = [wrap(row[0].text, right - M - CELL_PAD * 2, FONT, true)];
          height = lineSets[0].length * LINE_H + 3;
        } else {
          lineSets = [];
          let col = 0;
          for (const c of row) {
            const w = widths.slice(col, col + c.span).reduce((a, b) => a + (b || 0), 0) || widths[col] || avail;
            lineSets.push({ col, text: c.text, w, lines: wrap(c.text, w - CELL_PAD * 2, FONT, false) });
            col += c.span;
          }
          height = Math.max(...lineSets.map(s => s.lines.length)) * LINE_H + 3;
        }

        if (y + height > bottom) { newPage(); drawHeader(); }

        if (spanning) {
          // A heading row inside the table (the parties table interleaves the
          // complaint / cross-complaint headings with the party rows) — banded
          // so it reads as the divider it is.
          pdf.fillRect(M, y, right - M, height, 0.94);
          lineSets[0].forEach((ln, k) => pdf.text(ln, M + CELL_PAD, y + LINE_H * (k + 1) - 2, FONT, true));
        } else {
          for (const s of lineSets) {
            const cx = xs[s.col] != null ? xs[s.col] : M;
            s.lines.forEach((ln, k) => pdf.text(ln, cx + CELL_PAD, y + LINE_H * (k + 1) - 2, FONT, false));
          }
        }
        y += height;
        pdf.line(M, y, right, y, 0.4, 0.75);
      }
      y += 6;
    }
    footer();
    return pdf.build();
  }

  return { extractSections, renderPdf };
})();

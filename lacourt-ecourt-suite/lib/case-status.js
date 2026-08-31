/**
 * Shared case-status engine.
 *
 * Everything needed to say where a case stands on its next hearing — the
 * §1005/§437c briefing deadlines, which papers are actually on file, and the
 * default/dismissal status for an OSC Re: Failure to Prosecute Default Judgment
 * — with no dependency on WHICH page is asking.
 *
 * Two pages ask:
 *   - clipboard/content.js paints it inline on the case page's "Next" header;
 *   - agenda/content.js paints it beside each case name on the daily agenda.
 *
 * A case's data (Documents, Hearings, Parties) is reached through a "case
 * context" (makeCaseCtx) rather than the ambient page, so the same engine can
 * run against the live case page or against any case id fetched in the
 * background. Each context memoizes its fetches, so N callers cost one fetch.
 *
 * Loaded as a plain content script before its consumers, which pull what they
 * need out of the LACCaseStatus object.
 */
var LACCaseStatus = (function () {
'use strict';

/**
 * Parses the parties table.
 * Rows in the parties section each contain an a[title='UPDATE PARTY'] anchor.
 * The party name and role are in nearby cells. Structure may vary, so we walk
 * up to the row, then read its text cells.
 */
function parsePartiesTable(root) {
  root = root || document;
  const anchors = root.querySelectorAll('a[title="UPDATE PARTY"]');
  if (anchors.length === 0) return [];

  // Pattern that marks a party as removed and no-longer-named on the case.
  // Such parties are excluded from the paste output. Match is case-
  // insensitive. Variants observed in the wild:
  //   "Removed - No Longer Named 03/11/2026"  (Removed + dash + phrase)
  //   "Removed-No Longer Named 03/11/2026"    (no spaces around dash)
  //   "REMOVED \u2013 NO LONGER NAMED 03/11/2026" (en-dash, all caps)
  //   "No Longer Named 01/24/2025"            (bare; observed on cross-
  //                                            defendant rows that were
  //                                            dropped from the case without
  //                                            a "Removed -" prefix)
  // We match either form: the bare "no longer named" phrase suffices since
  // it's specific enough to avoid false positives.
  const REMOVED_RE = /\bno\s+longer\s+named\b/i;

  // Heading rows in the parties table introduce a "section" — the original
  // complaint, an amended complaint, or one of potentially several cross-
  // complaints. We only want parties from the original/amended complaint
  // section AND from the FIRST cross-complaint section. Any subsequent
  // cross-complaints are dropped (the user will splice those into the mail
  // merge document by hand).
  //
  // Sample heading row texts:
  //   "Complaint filed by Fuxin Sun on 03/11/2026"
  //   "Amended Complaint (2nd) filed by Fuxin Sun on 03/11/2026"
  //   "Cross-Complaint filed by Try Touch Service on 04/12/2026"
  //
  // The headings are anchored with ^ because some E-court tables include a
  // giant outer "container" row whose textContent concatenates every cell in
  // the table, including the heading text from real heading rows further
  // down. Without the anchor, that container row would falsely match
  // HEADING_CROSS_RE (consuming the "first cross-complaint" slot) and the
  // genuine cross-complaint heading would then be treated as the SECOND one
  // and have its parties dropped. The combination of an anchored regex +
  // requiring the row to have no UPDATE PARTY anchor (real heading rows
  // never do) reliably excludes the container row.
  const HEADING_COMPLAINT_RE = /^\s*(amended\s+)?complaint\s+filed\s+by\b/i;
  const HEADING_CROSS_RE     = /^\s*cross[-\s]?complaint\s+filed\s+by\b/i;
  // An appeal opens its own subcase section, e.g.
  //   "Appeal filed by Adam Roe on 05/14/2026"
  // whose rows re-list parties under their APPELLATE designations —
  // "Appellant (Appellant)" and "Respondent (Respondent)". Those designations
  // are carried IN ADDITION to the party's trial-court role, and the party is
  // already listed under that role in the complaint section above. This is a
  // trial court, so the appellate rows are dropped: keeping them would file the
  // plaintiff/appellate respondent under the defendant/respondent side (wrong
  // caption on export, and a false "not in default/dismissed" defendant on the
  // OSC Re: Failure to Prosecute Default Judgment status line).
  const HEADING_APPEAL_RE    = /^\s*(notice\s+of\s+)?(cross[-\s]?)?appeal\s+filed\s+by\b/i;

  // Locate the parties table. The UPDATE PARTY anchors live inside it; walk
  // up from the first anchor to its enclosing <table>.
  const firstRow = anchors[0].closest('tr');
  const partiesTable = firstRow && firstRow.closest('table');
  if (!partiesTable) return [];

  // Iterate every <tr> in document order so heading rows can be detected
  // and the per-row "section" tracked. anchors-based iteration alone would
  // miss the heading rows entirely.
  const allRows = Array.from(partiesTable.querySelectorAll('tr'));

  const parties = [];
  // Section enum:
  //   'primary'  → original/amended complaint (always included)
  //   'cross-1'  → first cross-complaint (included)
  //   'cross-N'  → subsequent cross-complaints (excluded)
  //   'appeal'   → an appeal subcase (excluded; see HEADING_APPEAL_RE)
  let currentSection = 'primary';
  let crossSectionsSeen = 0;

  for (const row of allRows) {
    const rowText = (row.textContent || '').trim().replace(/\s+/g, ' ');
    if (!rowText) continue;

    // Real heading rows never contain an UPDATE PARTY anchor. Requiring the
    // absence of one is a defense-in-depth guard against rows whose text
    // happens to start with a heading-shaped phrase (e.g. a party-row whose
    // name begins with "Complaint" or similar).
    const hasUpdateAnchor = !!row.querySelector('a[title="UPDATE PARTY"]');

    // Heading-row detection. Cross- check goes first because the cross
    // pattern is a more specific superset (the primary pattern's "complaint
    // filed by" substring also appears inside "Cross-Complaint filed by").
    if (!hasUpdateAnchor && HEADING_APPEAL_RE.test(rowText)) {
      currentSection = 'appeal-skip';
      continue;
    }
    if (!hasUpdateAnchor && HEADING_CROSS_RE.test(rowText)) {
      crossSectionsSeen += 1;
      currentSection = (crossSectionsSeen === 1) ? 'cross-1' : 'cross-skip';
      continue;
    }
    if (!hasUpdateAnchor && HEADING_COMPLAINT_RE.test(rowText)) {
      currentSection = 'primary';
      continue;
    }

    // Party-row detection: must contain an UPDATE PARTY anchor.
    if (!hasUpdateAnchor) continue;

    // Drop any parties belonging to a 2nd+ cross-complaint section, or to an
    // appeal subcase (the party's substantive role comes from the complaint
    // section; the appellate designation is not a caption role here).
    if (currentSection === 'cross-skip' || currentSection === 'appeal-skip') continue;

    // Read each cell's trimmed text.
    const cells = Array.from(row.querySelectorAll('td')).map(td => {
      return (td.textContent || '').trim().replace(/\s+/g, ' ');
    }).filter(Boolean);

    // If any cell on this row indicates the party has been removed, skip
    // the row entirely. The "Party Status" column is column 5 by header
    // order, but empty cells get filtered out above so positional indexing
    // isn't reliable — text-matching across all cells is more robust and
    // there are no other columns whose values would collide with this
    // phrase.
    if (cells.some(c => REMOVED_RE.test(c))) continue;

    // Detect "Dismissed MM/DD/YYYY" in the Party Status column. Unlike
    // "No Longer Named", a dismissed party may still be a party to certain
    // post-judgment motions (attorney fees, costs, sanctions). We keep
    // these parties in the parsed output and flag them as dismissed so the
    // caller can decide whether to drop them based on motion type.
    const DISMISSED_RE = /\bdismissed\b/i;
    const dismissed = cells.some(c => DISMISSED_RE.test(c));

    // Detect an entered default and its date from the status column, e.g.
    // "Defaulted 01/29/2025" / "Default 05/12/2026" / "Default Entered 05/12/2026".
    // Match "default" or "defaulted" (the plain \bdefault\b boundary fails on the
    // "-ed" form eCourt actually uses). Used by the OSC Re: Failure to Prosecute
    // Default Judgment status indicator.
    let defaultDate = null;
    for (const c of cells) {
      if (!/\bdefault(?:ed)?\b/i.test(c)) continue;
      const dm = c.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
      defaultDate = dm ? parseHearingDateTime(dm[1]) : new Date(0); // date if present, else epoch sentinel
      break;
    }

    // Identify the role cell (one of the cells starts with a known role keyword).
    // Note: cross-* alternatives are listed first so the regex engine matches
    // them as a single token rather than letting "Defendant" match the start
    // of "Cross-Defendant" (it can't anyway because of the ^ anchor, but this
    // keeps the intent explicit).
    let role = '';
    let roleIdx = -1;
    const roleRe = /^(cross[-\s]?complainant|cross[-\s]?defendant|plaintiff|defendant|petitioner|respondent)\b/i;
    for (let i = 0; i < cells.length; i++) {
      if (roleRe.test(cells[i])) {
        role = cells[i];
        roleIdx = i;
        break;
      }
    }

    // The name is the first non-role, non-action-button cell.
    let name = '';
    for (let i = 0; i < cells.length; i++) {
      if (i === roleIdx) continue;
      const text = cells[i];
      if (/^(update\s*party|edit|delete|view|action)$/i.test(text)) continue;
      // Skip cells that are purely numeric (party index).
      if (/^\d+\.?$/.test(text)) continue;
      name = text;
      break;
    }

    // Strip parenthetical content from name.
    if (name) {
      // For a defendant / cross-defendant, preserve a "(Doe N)" / "(Does 1-10)"
      // designation — it identifies which fictitious defendant the party was
      // named as — even though the role and entity-type parentheticals are
      // stripped.
      let doe = '';
      if (/\b(?:cross[-\s]?defendant|defendant)\b/i.test(role)) {
        const dm = name.match(/\(\s*does?\b[^)]*\)/i);
        if (dm) doe = dm[0].replace(/\s+/g, ' ').trim();
      }
      const parenIdx = name.indexOf('(');
      if (parenIdx !== -1) name = name.substring(0, parenIdx).trim();
      // Also strip trailing "Update Party" if it leaked in.
      name = name.replace(/\s*update\s*party\s*$/i, '').trim();
      if (doe) name = (name + ' ' + doe).trim();
    }

    if (name) {
      parties.push({ name, role, dismissed, defaultDate });
    }
  }

  return parties;
}

/**
 * Strips a trailing system event id from a Next-event description.
 *
 * e-court appends an event id to the end of the Next-event text. The format
 * has drifted over time — seen both without and with a space after the hyphen:
 *   "Hearing on Motion to Compel Discovery ID-148870297793"
 *   "Hearing on Motion for Summary Judgment ID- 396523215423"
 * The user wants "ID" and everything after it dropped from the motion type /
 * hearing type. We match a standalone "ID" token followed by any mix of
 * separators (hyphen / colon / hash / spaces) and then digits, through end of
 * string. Requiring the trailing digits (and the \b before "ID") keeps real
 * words like "grid-5" or "...Valid" from being clipped.
 */
function stripEventId(desc) {
  if (!desc) return desc;
  return desc.replace(/\s*\bID\b[\s:#-]*\d[\d\s]*$/i, '').trim();
}

/**
 * Drops trailing number-only decorations from a motion type: a purely numeric
 * parenthetical ("Motion to Compel Discovery (12345)", "Demurrer (2)") or a
 * dash-number ("Motion to Compel - 3891"). Repeats so stacked forms
 * ("X (123) - 456") fully strip. Alphanumeric parentheticals such as
 * "(CCP 437c)" / "(Set One)" and worded dashes ("Demurrer - without Motion to
 * Strike") are kept. KEEP IN SYNC with agenda/content.js.
 */
function stripTrailingParenNumber(s) {
  if (!s) return s;
  let prev;
  do {
    prev = s;
    s = s.replace(/\s*(?:\(\s*\d[\d\s.,\-]*\)|[-–—]\s*\d[\d\s.,]*)\s*$/, '').trim();
  } while (s !== prev);
  return s;
}

function movantNormName(s) {
  return (s || '').toUpperCase().replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

// "motion"/"hearing" appear in nearly every motion type and doc name, so they
// carry no signal for matching a hearing to its moving paper — treat them as
// stopwords so the discriminating terms (demurrer, compel, strike, …) decide.
const MOVANT_STOPWORDS = new Set(['the', 'of', 'for', 'and', 'to', 'a', 'an', 'on', 'in', 're', 'with', 'by', 'motion', 'hearing']);

function movantSigTokens(s) {
  return (s || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim()
    .split(' ').filter(t => t.length > 1 && !MOVANT_STOPWORDS.has(t));
}

// eCourt document names carry standardized qualifiers that describe what a paper
// does NOT include and inject tokens absent from the hearing description, e.g.
// "Motion to Strike (not anti-SLAPP) - without Demurrer" or "Demurrer - without
// Motion to Strike". Strip parentheticals and any trailing dash-prefixed
// "with/without …" clause so matching keys on the core motion phrase.
//
// A leading "Motion - Other" is eCourt's catch-all hearing CATEGORY, not part
// of the motion's name — the free text after it is the name ("Motion - Other
// Motion to Substitute Executor of Estate …"). No filing title ever carries
// the word "other", so leaving the wrapper in only dilutes the hearing-side
// coverage; a substitute-executor hearing scored exactly 0.5 against its own
// moving paper and read "No Motion". Strip it too (only when text follows —
// a bare "Motion - Other" keeps its tokens rather than matching nothing).
function stripMotionQualifiers(s) {
  return (s || '')
    .replace(/\([^)]*\)/g, ' ')
    .replace(/^\s*motion\s*[-–—]\s*other\s+(?=\S)/i, '')
    .replace(/\s[-–—]\s*(?:with|without)\b.*$/i, ' ')
    .replace(/\s+/g, ' ').trim();
}

function movantTokenHit(a, b) {
  if (a === b) return true;
  if (a.length >= 4 && b.length >= 4 && (b.startsWith(a) || a.startsWith(b))) return true;
  // Neither "bifurcate" nor "bifurcation" is a prefix of the other, so the
  // startsWith test above misses the pair the shared-stem test catches.
  return sameWordFamily(a, b);
}

// Directional token coverage between a hearing's motion type and a filing
// title: mtCov = the fraction of the motion type's significant tokens the doc
// name accounts for, dnCov = the reverse.
function movantMatchCoverages(motionType, docName) {
  const mt = [...new Set(movantSigTokens(stripMotionQualifiers(motionType)))];
  const dn = [...new Set(movantSigTokens(stripMotionQualifiers(docName)))];
  if (!mt.length || !dn.length) return { mtCov: 0, dnCov: 0 };
  let mtHit = 0;
  for (const t of mt) if (dn.some(d => movantTokenHit(t, d))) mtHit++;
  let dnHit = 0;
  for (const d of dn) if (mt.some(t => movantTokenHit(t, d))) dnHit++;
  return { mtCov: mtHit / mt.length, dnCov: dnHit / dn.length };
}

function movantMatchScore(motionType, docName) {
  const { mtCov, dnCov } = movantMatchCoverages(motionType, docName);
  // Bidirectional: a concise doc name that fully matches the motion type's key
  // terms scores high even when the motion type carries extra descriptors
  // (e.g. "Demurrer - without Motion to Strike" vs a doc named "Demurrer").
  return Math.max(mtCov, dnCov);
}

// Is this document Name an actual moving paper (motion/demurrer/etc.), not a
// response, order, minute order, declaration, etc. that rides alongside it?
function isMovingPaper(name) {
  if (!name) return false;
  const n = name.trim();
  if (/^(Opposition|Reply|Response|Declaration|Proof of (Personal )?Service|Order\b|Minute Order|Notice\b|Brief|Request\b|Certificate|Summons|Appeal\b|Case Management|Ex Parte Proposed Order|Points and Authorities|Memorandum|Stipulation|Objection|Separate Statement)/i.test(n)) {
    return false;
  }
  // A petition is a moving paper however eCourt qualifies its title — "Verified
  // Petition …", "First Amended Petition …" — not only as a bare leading
  // "Petition …". isPetitionDoc carries those prefixes, and rejects the papers
  // that merely reference a petition.
  if (isPetitionDoc(n)) return true;
  return /^(Motion|Demurrer|Petition|Application|Ex Parte Application|Anti-SLAPP|Special Motion|Amended Motion|Renewed Motion|Cross-?Motion)/i.test(n);
}

function bestFilingMatch(motionType, filings) {
  let best = null, bestScore = 0, bestMin = 0;
  for (const f of filings) {
    if (!isMovingPaper(f.name)) continue;
    const { mtCov, dnCov } = movantMatchCoverages(motionType, f.name);
    const s = Math.max(mtCov, dnCov), sMin = Math.min(mtCov, dnCov);
    // Rival motions can cover the hearing type EQUALLY — two "Motion to
    // Substitute Personal Representative…" papers, one per decedent capacity,
    // each set for its own hearing. The coverage in the other direction (how
    // much of the doc's own title the hearing type accounts for) then
    // separates the paper actually noticed for THIS hearing from its
    // longer-titled rival, instead of docket order deciding the tie.
    if (s > bestScore || (s === bestScore && sMin > bestMin)) {
      bestScore = s; bestMin = sMin; best = f;
    }
  }
  // Strictly greater than half: a score of exactly 0.50 is usually a name that
  // overlaps the hearing type by one incidental word and diverges on the rest —
  // "Motion in Limine ... to Exclude Testimony at Trial" against a Motion for
  // New Trial, sharing only "trial". Accepting those picked the wrong moving
  // paper for a hearing whose motion was never filed, and the single-hearing
  // sweep then opened the whole docket after it.
  if (bestScore > 0.5) return best;

  // Except when the shared word IS the species. A half score between papers
  // that classify into the SAME specific rule bucket is a same-family pairing,
  // not an incidental one: "Motion for Summary Judgment" and "Motion for
  // Summary Adjudication" share only "summary", yet § 437c governs both and
  // classifyMotion already puts them together. Rejecting that pairing left an
  // MSA hearing reading "No Motion" with the motion plainly on the docket.
  //
  // 'standard' is deliberately excluded — it is the catch-all, and under it two
  // motions sharing one word ("compel" arbitration vs "compel" further
  // responses) are exactly the incidental pairing the threshold exists to
  // reject.
  if (best && bestScore >= 0.5) {
    const cat = DL.classifyMotion(motionType || '');
    if (cat !== 'standard' && DL.classifyMotion(best.name || '') === cat) return best;
  }
  return null;
}

// Parse a "Filed By" cell into { parties:[{name,role}], truncated }.
function parseFiledByParties(text) {
  if (!text) return { parties: [], truncated: false };
  let s = text.replace(/\s+/g, ' ').trim();
  let truncated = false;
  if (/\bet al\.?\s*$/i.test(s)) { truncated = true; s = s.replace(/\s*\bet al\.?\s*$/i, '').trim(); }
  const parts = s.split(';').map(x => x.trim()).filter(Boolean);
  const parties = [];
  for (const p of parts) {
    const m = p.match(/^(.*?)\s*\(([^)]+)\)\s*$/);
    if (m) parties.push({ name: m[1].trim(), role: m[2].trim() });
    else parties.push({ name: p, role: '' });
  }
  return { parties, truncated };
}

// Keep in sync with agenda/content.js DEFAULT_EXCLUDED_TERMS.
const DEFAULT_EXCLUDED_TERMS = [
  'conference',
  'non-appearance case revie',
  'non-jury trial',
  'order to show cause re: d',
  'ex parte',
  'application for order for',
  'jury trial',
  'post-arbitration status c',
  'post-mediation status con',
  'order to show cause re: s',
  'informal discovery confer',
];

let EXCLUDED_TERMS_CACHE = null;

function loadExcludedTerms() {
  return new Promise(resolve => {
    if (EXCLUDED_TERMS_CACHE) { resolve(EXCLUDED_TERMS_CACHE); return; }
    try {
      chrome.storage.sync.get(['excludedTerms'], r => {
        const terms = (r && Array.isArray(r.excludedTerms) && r.excludedTerms.length)
          ? r.excludedTerms : DEFAULT_EXCLUDED_TERMS;
        EXCLUDED_TERMS_CACHE = terms;
        resolve(terms);
      });
    } catch (_) {
      EXCLUDED_TERMS_CACHE = DEFAULT_EXCLUDED_TERMS;
      resolve(DEFAULT_EXCLUDED_TERMS);
    }
  });
}

// Case-insensitive substring match against the exclusion terms, mirroring the
// agenda cleaner's isExcluded().
// Match one excluded term against a hearing text. A term wrapped in double
// quotes ("...") requires the whole text to equal it exactly (trimmed); an
// unquoted term matches as a substring. Terms are already stored lowercased.
// KEEP IN SYNC with agenda/content.js excludedTermMatches().
function excludedTermMatches(term, lower) {
  if (!term) return false;
  const quoted = term.match(/^"(.*)"$/);
  if (quoted) return lower.trim() === quoted[1].trim();
  return lower.includes(term);
}

// The options page can edit the exclusion list while a page is open; keep the
// cache current so a hearing's workable/excluded verdict follows the edit
// without a reload. (This listener moved here with the cache it maintains.)
try {
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'sync' && changes.excludedTerms) {
      EXCLUDED_TERMS_CACHE = Array.isArray(changes.excludedTerms.newValue) && changes.excludedTerms.newValue.length
        ? changes.excludedTerms.newValue : DEFAULT_EXCLUDED_TERMS;
    }
  });
} catch (_) {}

function isHearingExcluded(text) {
  if (!text) return false;
  const lower = text.toLowerCase();
  const terms = EXCLUDED_TERMS_CACHE || DEFAULT_EXCLUDED_TERMS;
  return terms.some(t => excludedTermMatches(t, lower));
}

function stripHearingOnPrefix(s) {
  return (s || '').replace(/^\s*Hearing on\s+/i, '').trim();
}

// Parses "MM/DD/YYYY HH:MM AM/PM" into a Date (local). Returns null on failure.
function parseHearingDateTime(s) {
  const m = (s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})\s*(AM|PM)?)?/i);
  if (!m) return null;
  const mo = +m[1], d = +m[2], y = +m[3];
  let hh = m[4] ? +m[4] : 0;
  const mm = m[5] ? +m[5] : 0;
  if (m[6]) {
    const up = m[6].toUpperCase();
    if (up === 'PM' && hh < 12) hh += 12;
    if (up === 'AM' && hh === 12) hh = 0;
  }
  return new Date(y, mo - 1, d, hh, mm);
}

// Parses future scheduled hearings from a Hearings-tab document: rows in tables
// whose header has Name / Date/Time / Status columns, kept when Status is
// "Scheduled" and the date is today or later. Returns them soonest-first,
// deduped by type+date. Each: { type, date, when }.
function parseFutureHearings(doc) {
  const rows = [];
  const tables = doc.querySelectorAll('table');
  for (const table of tables) {
    let headerRow = null, nameIdx = -1, dateIdx = -1, statusIdx = -1;
    for (const tr of table.querySelectorAll('tr')) {
      const texts = Array.from(tr.children).map(td => (td.textContent || '').replace(/\s+/g, ' ').trim());
      const ni = texts.indexOf('Name'), di = texts.indexOf('Date/Time'), si = texts.indexOf('Status');
      if (ni !== -1 && di !== -1 && si !== -1) { headerRow = tr; nameIdx = ni; dateIdx = di; statusIdx = si; break; }
    }
    if (!headerRow) continue;

    let started = false;
    const maxIdx = Math.max(nameIdx, dateIdx, statusIdx);
    for (const tr of table.querySelectorAll('tr')) {
      if (tr === headerRow) { started = true; continue; }
      if (!started) continue;
      const cells = Array.from(tr.children);
      if (cells.length <= maxIdx) continue; // skip the continuance sub-rows
      const name = (cells[nameIdx] ? cells[nameIdx].textContent : '').replace(/\s+/g, ' ').trim();
      const dateTime = (cells[dateIdx] ? cells[dateIdx].textContent : '').replace(/\s+/g, ' ').trim();
      const status = (cells[statusIdx] ? cells[statusIdx].textContent : '').replace(/\s+/g, ' ').trim();
      if (!name || !dateTime) continue;
      if (!/^scheduled$/i.test(status)) continue; // only genuinely upcoming
      const when = parseHearingDateTime(dateTime);
      const dm = dateTime.match(/\d{1,2}\/\d{1,2}\/\d{4}/);
      if (!when || !dm) continue;
      // `raw` keeps the docket's own wording (event number and all) so a row the
      // case page has to RENDER reads exactly as eCourt would have written it;
      // `timeText` is the clock time, which the MM/DD/YYYY-only `date` drops.
      const tm = dateTime.match(/\d{1,2}:\d{2}\s*(?:AM|PM)/i);
      rows.push({
        type: stripTrailingParenNumber(stripEventId(name)),
        raw: name,
        date: dm[0],
        timeText: tm ? tm[0].replace(/\s+/g, ' ').toUpperCase() : '',
        when,
      });
    }
  }

  const startOfToday = new Date(); startOfToday.setHours(0, 0, 0, 0);
  const seen = new Set();
  return rows
    .filter(h => h.when >= startOfToday)
    .filter(h => { const k = h.type + '@' + h.date; if (seen.has(k)) return false; seen.add(k); return true; })
    .sort((a, b) => a.when - b.when);
}

// A hearing we "work up": a motion (a "Hearing on <motion>" event, minus the
// few excluded types like ex parte) or an OSC Re: Failure to Prosecute Default
// Judgment. Case management conferences, trials, status conferences, reviews,
// and other OSCs are NOT worked up, so anything keyed on the Next event should
// skip past them to the next workable hearing.
// A motion we work up, named either as the "Next:" banner form ("Hearing on
// <motion>") or as a Hearings-tab entry that names the motion directly
// ("Motion …", "Demurrer …", "Anti-SLAPP …", "Petition …", "Application …").
const MOTION_HEARING_RE = /\bhearing on\b|^(?:amended\s+|renewed\s+|cross-?\s*|special\s+)?motion\b|^demurrer\b|^anti-?slapp\b|^petition\b|^application\b/i;

function isWorkableHearing(type) {
  if (!type) return false;
  if (isOscDefaultJudgment(type)) return true;
  if (isHearingExcluded(type)) return false;
  return MOTION_HEARING_RE.test(type);
}

// A demurrer filed "with Motion to Strike" is ONE work-up that eCourt calendars
// as two hearing rows. The standalone strike row is the duplicate, so it drops
// out when a demurrer shares its date — except a motion to strike or tax COSTS,
// which is its own motion on its own schedule and merely shares a verb.
function isDuplicateStrikeOf(item, sameDay) {
  if (!/\bmotion to strike\b/i.test(item.hearingType || '')) return false;
  if (DL.classifyMotion(item.motionType || item.hearingType || '') === 'costs') return false;
  return sameDay.some(o => o !== item && /\bdemurrer\b/i.test(o.hearingType || ''));
}

/* eCourt writes the SAME hearing's caption two different ways, so it cannot be
   compared letter-for-letter. The page's own "Next" header reads it out of a
   title attribute, spaced as a person typed it; the Hearings tab reads it out of
   table cells whose caption is split across child elements, and textContent
   joins those with nothing at all — "Motion To Seal … In Support Of" on the
   header comes back as "MotionTo Seal … In SupportOf" from the tab, with
   "Judgment(0753)" for "Judgment (0753)" and a straight apostrophe for a curly
   one. Keying on the letters and digits alone makes the two spellings one
   hearing again; before this, the case page showed it twice, once per source.
   Two genuinely different motions never differ only in whitespace. */
function hearingIdentityKey(type) {
  return stripHearingOnPrefix(type || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

// Every hearing on this case that we work up, bundled by hearing DATE: one group
// per calendar day, each holding the motions set for that day soonest-first.
// `native` is the case page's own "Next:" event ({ motionType, hearingType,
// hearingDate, timeText, raw }); `hearings` is the Hearings tab. The Next event
// leads its group when it's workable — it is the one hearing whose wording the
// page already shows — and the Hearings tab supplies the rest. Groups come back
// soonest-first, so index 0 is the hearing the page would have used on its own.
function groupWorkableHearings(native, hearings) {
  const items = [];
  const seen = new Set();
  const push = it => {
    if (!it.hearingDate) return;
    const key = it.hearingDate + '|' + hearingIdentityKey(it.hearingType);
    if (seen.has(key)) return;
    seen.add(key);
    items.push(it);
  };

  if (native && isWorkableHearing(native.hearingType)) {
    push({
      // Left as the page reads it — an OSC Re: Failure to Prosecute Default
      // Judgment is not a "Hearing on <motion>" event and has no motion type,
      // and the order template has always had that box empty for one.
      motionType: native.motionType || '',
      hearingType: native.hearingType,
      hearingDate: native.hearingDate,
      timeText: native.timeText || '',
      raw: native.raw || native.hearingType || '',
      when: parseHearingDateTime(native.hearingDate),
      native: true,
    });
  }
  for (const h of (hearings || [])) {
    if (!isWorkableHearing(h.type)) continue;
    push({
      motionType: stripHearingOnPrefix(h.type),
      hearingType: h.type,
      hearingDate: h.date,
      timeText: h.timeText || '',
      raw: h.raw || h.type,
      when: h.when,
      native: false,
    });
  }

  const byDay = new Map();
  for (const it of items) {
    if (!byDay.has(it.hearingDate)) byDay.set(it.hearingDate, []);
    byDay.get(it.hearingDate).push(it);
  }
  const groups = [];
  for (const [date, list] of byDay) {
    // The page's own event first, then the docket's order (which carries the
    // clock time the MM/DD/YYYY string doesn't).
    list.sort((a, b) => (b.native ? 1 : 0) - (a.native ? 1 : 0)
      || ((a.when ? a.when.getTime() : 0) - (b.when ? b.when.getTime() : 0)));
    const kept = list.filter(it => !isDuplicateStrikeOf(it, list));
    if (!kept.length) continue;
    groups.push({ date, when: kept[0].when, items: kept });
  }
  groups.sort((a, b) => (a.when ? a.when.getTime() : 0) - (b.when ? b.when.getTime() : 0));
  return groups;
}

function fetchWithTimeout(url, ms) {
  return Promise.race([
    fetch(url, { credentials: 'include' }),
    new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), ms)),
  ]);
}

// Fetch a case page (same authenticated origin) and parse it into a Document.
// Returns null on any failure.
async function fetchCaseDoc(url) {
  try {
    const res = await fetchWithTimeout(url, 6000);
    if (!res || !res.ok) return null;
    const html = await res.text();
    return new DOMParser().parseFromString(html, 'text/html');
  } catch (_) {
    return null;
  }
}

const DOC_STOP = new Set(['motion','opposition','reply','response','notice','declaration',
  'memorandum','points','authorities','support','order','proposed','hearing','plaintiff',
  'defendant','plaintiffs','defendants','exhibit','proof','service','with','from','that',
  'this','case','court','filed','amended']);

function docSigTokens(s) {
  return (s || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/)
    .filter(t => t.length >= 4 && !DOC_STOP.has(t));
}

/* Docket titles are typed by hand and carry typos — this docket alone has
   "Motion for New Trail" and "Compelling Abritration". One edit still names the
   same motion, so tokens are compared loosely where a miss would lose the only
   papers that reference a hearing. Adjacent transpositions count as one edit
   (they are the common typo, and "trial"/"trail" is exactly that); the 5-char
   floor keeps short words from colliding. */
function withinOneEdit(a, b) {
  if (a === b) return true;
  const la = a.length, lb = b.length;
  if (Math.abs(la - lb) > 1) return false;
  let i = 0, j = 0, edits = 0;
  while (i < la && j < lb) {
    if (a[i] === b[j]) { i++; j++; continue; }
    if (++edits > 1) return false;
    if (la === lb) {
      if (a[i + 1] === b[j] && a[i] === b[j + 1]) { i += 2; j += 2; continue; } // transposed
      i++; j++;
    } else if (la > lb) { i++; } else { j++; }
  }
  return true;
}

// Does this document's name reference `motionType`, allowing for docket typos?
// Deliberately looser than docWordOverlap, and used only where no moving paper
// is on file — there the referencing papers are the whole record of the hearing,
// so a missed token costs more than a stray match.
function docReferencesMotion(name, motionType) {
  const a = docSigTokens(motionType);
  if (!a.length) return false;
  const dn = docSigTokens(name);
  return dn.some(d => a.some(t => (t.length >= 5 && d.length >= 5) ? withinOneEdit(t, d) : t === d));
}

/* Two tokens name the same thing when one is an inflection of the other:
   "bifurcate"/"bifurcation", "adjudicate"/"adjudication", "dismiss"/"dismissal".
   Motions get captioned with the verb and briefed with the noun, so demanding
   exact equality lost the opposition on a motion to bifurcate.

   The test is a shared prefix of six characters. Six because five would pair
   "appeal" with "appear"; six is still short enough for every motion-word pair
   a docket actually produces ("bifurcat", "adjudicat", "reconsider"). Tokens
   under six characters must match exactly — "trial" and "trail" stay apart, as
   do "stay" and "stab". */
const DOC_STEM_MIN = 6;
function sameWordFamily(a, b) {
  if (a === b) return true;
  if (a.length < DOC_STEM_MIN || b.length < DOC_STEM_MIN) return false;
  const n = Math.min(a.length, b.length);
  let i = 0;
  while (i < n && a[i] === b[i]) i++;
  return i >= DOC_STEM_MIN;
}

function docWordOverlap(name, motionType) {
  const a = docSigTokens(motionType);
  if (!a.length) return false;
  return docSigTokens(name).some(t => a.some(x => sameWordFamily(x, t)));
}

// Common motion abbreviations, so a filing that names the motion by its shorthand
// ("Reply ISO Ford's MSA") still links to the spelled-out motion type ("Motion
// for Summary Adjudication") even though no full word overlaps.
const MOTION_ABBREV_RE = /\bms[ja]\b|summary\s+(?:judgment|adjudication)|\bmil\b|motion\s+in\s+limine|\bmtc\b|motion\s+to\s+compel|\bmtq\b|motion\s+to\s+quash/i;

function motionAbbrevMatch(name, motionType) {
  return MOTION_ABBREV_RE.test(name || '') && MOTION_ABBREV_RE.test(motionType || '');
}

// A filing links to the upcoming motion if it shares a significant word or a
// known abbreviation with the motion type.
function docLinksToMotion(name, motionType) {
  return docWordOverlap(name, motionType) || motionAbbrevMatch(name, motionType);
}

/* A paper whose title says only what KIND of paper it is — "Opposition
   OPPOSITION", "Reply REPLY", "Opposition to Motion" — names no motion at all:
   every word in it is a DOC_STOP word, so docSigTokens comes back empty and
   docLinksToMotion can never be true for it, however the motion is captioned.
   A paper like that has to be attributed by something other than its name (who
   filed it, what is on calendar) rather than dropped — dropping it reported
   "No Opposition" with the opposition sitting on the docket. */
function docNameIsGeneric(name) {
  return !docSigTokens(name).length;
}

// A real opposition — NOT a "Notice of Non-Opposition", which the moving party
// files to note the ABSENCE of an opposition and must not count as one.
function isOppositionDoc(name) {
  return /\bopposition\b/i.test(name || '') && !/\bnon-?\s*opposition\b/i.test(name || '');
}

// A "Notice of Non-Opposition" / "Statement of No Opposition" — a filing that
// affirmatively notes the ABSENCE of any opposition, rather than opposing on the
// merits. Depending on who filed it, it stands in for a Reply (moving party) or
// for the Opposition (non-moving party). See fetchNextDeadlineFilings.
function isNonOppositionDoc(name) {
  return /\bnon-?\s*opposition\b/i.test(name || '') || /\bno\s+opposition\b/i.test(name || '');
}

function docPartyNames(filedBy) {
  return parseFiledByParties(filedBy).parties.map(p => movantNormName(p.name)).filter(Boolean);
}

function docSharesParty(a, b) { const A = new Set(a); return b.some(x => A.has(x)); }

function sameCalendarDay(x, y) { return !!(x && y && x.getTime() === y.getTime()); }

function isComplaintDoc(name) {
  const n = (name || '').trim();
  if (/^amendment to /i.test(n)) return false;              // "Amendment to Complaint (Fictitious/Incorrect Name)"
  if (/fictitious|incorrect\s+name/i.test(n)) return false;
  // eCourt clerks sometimes double up the ordinal word, e.g. "Amended Amended
  // Complaint (1st)" — accept one or more "amended" repeats, not just one.
  return /^(?:(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|\d+(?:st|nd|rd|th))\s+)?(?:amended\s+)+complaint\b/i.test(n)
      || /^complaint\b/i.test(n);
}

function isCrossComplaintDoc(name) {
  const n = (name || '').trim();
  if (/^amendment to /i.test(n)) return false;
  return /^(?:(?:first|second|third|fourth|fifth|\d+(?:st|nd|rd|th))\s+)?(?:amended\s+)*cross-?complaint\b/i.test(n);
}

// An AMENDED complaint (not the original, not a cross-complaint). The
// "\bamended complaint\b" test won't match "amended cross-complaint" (the
// "cross-" breaks adjacency) or the original complaint.
function isAmendedComplaintDoc(name) {
  const n = (name || '').trim();
  if (/^amendment to /i.test(n)) return false;               // "Amendment to Complaint (Fictitious/Incorrect Name)"
  if (/fictitious|incorrect\s+name/i.test(n)) return false;
  return /\bamended\s+complaint\b/i.test(n);
}

// The ordinal of an amended complaint: 1 for "First"/"1st"/"(1st)" or a bare
// "Amended Complaint" (the first amendment is often titled without an ordinal),
// 2 for "Second"/"2nd", etc. Returns null when it isn't an amended complaint.
function amendedComplaintOrdinal(name) {
  if (!isAmendedComplaintDoc(name)) return null;
  const n = name || '';
  const WORDS = ['first','second','third','fourth','fifth','sixth','seventh','eighth','ninth','tenth'];
  const w = n.match(/\b(first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth)\b/i);
  if (w) return WORDS.indexOf(w[1].toLowerCase()) + 1;
  const num = n.match(/\b(\d+)\s*(?:st|nd|rd|th)\b/i);
  if (num) return parseInt(num[1], 10);
  return 1; // bare "Amended Complaint" — the first amendment
}

// A plaintiff may amend once as a matter of course in response to a demurrer
// (CCP 472). That single amendment is the FIRST amended complaint — a Second (or
// later) Amended Complaint is NOT of right and cannot moot a demurrer to the FAC.
// So only a first amended complaint counts as the "in lieu of opposition"
// response. Titled a few ways: "First Amended Complaint", "1st Amended
// Complaint", "Amended Complaint (1st)", or a bare "Amended Complaint".
function isFirstAmendedComplaintDoc(name) { return amendedComplaintOrdinal(name) === 1; }

// A demurrer or motion to strike — the challenges a plaintiff can moot by amending
// once as a matter of course (CCP 472). isMovingPaper excludes the oppositions,
// replies, notices, and orders that merely mention "demurrer"/"motion to strike".
function isDemurrerOrMotionToStrikeDoc(name) {
  return isMovingPaper(name) && (/\bdemurrer\b/i.test(name) || /\bmotion to strike\b/i.test(name));
}

// True when the hearing on calendar is itself a demurrer or motion to strike — a
// first amended complaint filed in lieu of opposition moots either one.
function isDemurrerOrStrikeMotion(motionType) {
  return /\bdemurrer\b|\bmotion to strike\b/i.test(motionType || '');
}

// A challenge whose Result column shows it's already been ruled on / disposed of,
// so it's not the moving paper for an upcoming hearing.
function isResolvedChallengeResult(result) {
  return /sustain|overrul|grant|deni|off[\s-]?calendar|moot|withdrawn|taken off/i.test(result || '');
}

// Resolve the moving paper for a given hearing. bestFilingMatch alone can't tell
// identically-named challenges apart — e.g. a demurrer to the complaint AND a
// demurrer to a cross-complaint, both "Demurrer - without Motion to Strike" — and
// grabs the newest. When the pending challenge docs line up one-to-one with the
// upcoming challenge hearings, pair them by date instead: the Nth challenge
// hearing (soonest first) is answered by the Nth pending challenge (earliest
// first). Falls back to bestFilingMatch whenever that clean pairing doesn't hold,
// so ordinary single-motion cases are unchanged.
function resolveMovingPaper(motionType, hearingWhen, hearings, docs) {
  let md = bestFilingMatch(motionType, docs);
  if (isDemurrerOrStrikeMotion(motionType) && hearingWhen) {
    const challengeHearings = (hearings || []).filter(h => isDemurrerOrStrikeMotion(h.type));
    const pendingChallenges = docs
      .filter(d => d.when && isDemurrerOrMotionToStrikeDoc(d.name) && !isResolvedChallengeResult(d.result))
      .sort((a, b) => a.when - b.when);
    const idx = challengeHearings.findIndex(h => dayMs(h.when) === dayMs(hearingWhen));
    if (idx >= 0 && challengeHearings.length === pendingChallenges.length && pendingChallenges[idx]) {
      md = pendingChallenges[idx];
    }
  }
  // Summary judgment and summary adjudication are coextensive: § 437c governs
  // both, classifyMotion puts them in one bucket, and a hearing captioned one
  // way is routinely briefed by a paper captioned the other — so the caption is
  // not what tells them apart. TIMING is. Where a case carries several, pair
  // them by date exactly as parallel demurrers are paired, so an MSA hearing
  // doesn't adopt the moving paper of an MSJ set months later; where there is
  // one of each, they pair regardless of what either is called.
  const isSummaryJudgmentFamily = t => DL.classifyMotion(t || '') === 'msj';
  if (isSummaryJudgmentFamily(motionType) && hearingWhen) {
    const sjHearings = (hearings || []).filter(h => isSummaryJudgmentFamily(h.type));
    const sjPapers = (docs || [])
      .filter(d => d.when && isMovingPaper(d.name) && isSummaryJudgmentFamily(d.name))
      .sort((a, b) => a.when - b.when);
    const sjIdx = sjHearings.findIndex(h => dayMs(h.when) === dayMs(hearingWhen));
    if (sjIdx >= 0 && sjHearings.length === sjPapers.length && sjPapers[sjIdx]) {
      md = sjPapers[sjIdx];
    }
  }
  // A petition heard in a case that HAS a complaint is a motion, but the
  // calendar's caption and the paper's own caption routinely describe the same
  // relief in different words: eCourt calendars "Petition to Confirm Minor's
  // Compromise" while the paper filed is the Judicial Council form, "Petition to
  // Approve Compromise of Disputed Claim …". Those share only "petition" and
  // "compromise" — a score of exactly 0.50, which bestFilingMatch rejects by
  // design — so the widget read "No Motion" with the petition on the docket.
  // Pair petition hearings to petition papers by date instead, exactly as
  // parallel demurrers are: where a case has one of each they pair whatever
  // either is called, and where it has several the Nth petition answers the Nth
  // petition hearing.
  if (isPetitionMotion(motionType) && hearingWhen && docsHaveComplaint(docs)) {
    const petHearings = (hearings || []).filter(h => isPetitionMotion(h.type));
    const petPapers = (docs || [])
      .filter(d => d.when && isPetitionDoc(d.name) && !isResolvedChallengeResult(d.result))
      .sort((a, b) => a.when - b.when);
    const pIdx = petHearings.findIndex(h => dayMs(h.when) === dayMs(hearingWhen));
    if (pIdx >= 0 && petHearings.length === petPapers.length && petPapers[pIdx]) {
      md = petPapers[pIdx];
    }
  }
  return md;
}

// A petition is another kind of initial pleading (probate, family, writ, etc.),
// used as the operative pleading only when the case has no complaint. Match the
// pleading ITSELF — a name that starts with "Petition" (optionally prefixed by
// "Verified"/"Amended"/an ordinal) — NOT documents that merely reference one
// ("Notice of Hearing on Petition", "Proof of Service of Petition", "Order on
// Petition").
function isPetitionDoc(name) {
  const n = (name || '').trim();
  if (/^amendment to /i.test(n)) return false;
  if (/fictitious|incorrect\s+name/i.test(n)) return false;
  return /^(?:(?:verified|amended|first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|\d+(?:st|nd|rd|th))\s+)*petition\b/i.test(n);
}

// A hearing ON a petition — the Hearings-tab form ("Petition to Confirm Minor's
// Compromise") or the case header's ("Hearing on Verified Petition for Writ of
// Mandate").
function isPetitionMotion(motionType) {
  return isPetitionDoc(stripHearingOnPrefix(motionType || ''));
}

/* Does this docket carry a complaint?

   A PETITION is a MOTION in a case that has one — a minor's compromise, a
   petition to confirm an arbitration award, a petition for approval brought
   inside an ordinary civil action — and is the initiating COMPLAINT in a case
   that has none (a standalone petition to compel arbitration, a writ petition).
   That is the whole of the rule, and it decides two things: whether the petition
   can be the moving paper for its own hearing, and whether that hearing carries
   a § 1005 briefing schedule at all.

   A cross-complaint counts as a complaint: it presupposes the pleading it
   answers, which may sit on a Documents page that failed to fetch. */
function docsHaveComplaint(docs) {
  return (docs || []).some(d => isComplaintDoc(d.name) || isCrossComplaintDoc(d.name));
}

/* Buckets a filing by the KIND of paper it is, for the filing-lag statistics.
   Different papers go through different clerk queues and post at different
   speeds — a minute order is entered by the clerk and appears the same day,
   while a default prove-up packet has been measured at 11 court days — so the
   lag distribution is only meaningful broken out by type.

   Classification is by name (plus the Filed By column for clerk-entered
   papers), so it runs retroactively over samples already collected. Order
   matters: the specific patterns are tested before the generic ones, and
   isMovingPaper() comes last because a moving paper is anything that isn't one
   of the named kinds. */
const DOC_LAG_CATEGORIES = [
  ['Clerk-entered',       n => /^minute order\b|^clerk'?s? certificate\b|^certificate of mailing\b/i.test(n)],
  ['Request for Default', n => /^request for entry of default\b/i.test(n)],
  ['Opposition',          n => /^opposition\b/i.test(n)],
  ['Reply',               n => /^reply\b/i.test(n)],
  ['Separate Statement',  n => /^separate statement\b/i.test(n)],
  ['Case Mgmt Statement', n => /^case management statement\b/i.test(n)],
  ['Proof of Service',    n => /^proof of (?:personal )?service\b|^proof of service\b/i.test(n)],
  ['Declaration',         n => /^declaration\b/i.test(n)],
  ['Stipulation',         n => /^stipulation\b/i.test(n)],
  ['Answer',              n => /^answer\b/i.test(n)],
  ['Pleading',            n => /^summons\b/i.test(n) || isComplaintDoc(n) || isCrossComplaintDoc(n) || isPetitionDoc(n)],
  ['Order',               n => /^order\b|\[\s*proposed\s*\]/i.test(n)],
  ['Notice',              n => /^notice\b/i.test(n)],
  ['Objection',           n => /^objection\b/i.test(n)],
  ['Motion',              n => isMovingPaper(n)],
];

function docLagCategory(name, filedBy) {
  const n = (name || '').trim();
  if (!n) return 'Other';
  // The clerk's own paperwork skips the intake queue entirely, so it is called
  // out however it happens to be titled.
  if (/^clerk\b/i.test((filedBy || '').trim())) return 'Clerk-entered';
  for (const [label, test] of DOC_LAG_CATEGORIES) {
    try { if (test(n)) return label; } catch (_) {}
  }
  return 'Other';
}

// Latest openable doc in a list (operative pleading).
function latestDoc(list) {
  let best = null;
  for (const d of list) {
    if (!d.openUrl) continue;
    if (!best || (d.when && best.when && d.when > best.when) || (!best.when && d.when)) best = d;
  }
  return best;
}

function absoluteDocUrl(u) { try { return new URL(u, location.origin).href; } catch (_) { return u; } }

// Parses openable document rows from a Documents-tab document or paged fragment:
// each openInNewWindow anchor yields
// { docId, openUrl, name, dateStr, when, filedBy, status, result }.
// Columns (by header order): Filed/Status Date(1) Name(2) Status(3) Result(4)
// Result Date(5) Filed By(6). The default prove-up check reads Result.
function parseDocRows(root) {
  const rows = [], seen = new Set();
  const anchors = root.querySelectorAll('a[onclick*="openInNewWindow"]');
  for (const a of anchors) {
    const oc = a.getAttribute('onclick') || '';
    const um = oc.match(/openInNewWindow\('((?:[^'\\]|\\.)*)'\s*,\s*'((?:[^'\\]|\\.)*)'/);
    if (!um) continue;
    const url = um[1].replace(/\\\//g, '/');
    if (!/\/ecourt\/ecms\/doc\?docId=/.test(url)) continue;
    const idm = url.match(/docId=(\d+)/); if (!idm) continue;
    const docId = idm[1]; if (seen.has(docId)) continue; seen.add(docId);
    const title = um[2].replace(/\\\//g, '/').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
    const tm = title.match(/^[^:]*:\s*([\s\S]*?)\s+on\s+(\d{1,2}\/\d{1,2}\/\d{4})\s*$/);
    const name = tm ? tm[1].trim() : title;
    const dateStr = tm ? tm[2] : '';
    let filedBy = '', status = '', result = ''; const tr = a.closest('tr');
    if (tr) {
      const cells = Array.from(tr.children);
      const cellText = i => (cells[i] ? (cells[i].textContent || '').replace(/\s+/g, ' ').trim() : '');
      if (cells.length > 6) filedBy = cellText(6);
      status = cellText(3);
      result = cellText(4);
    }
    rows.push({ docId, openUrl: absoluteDocUrl(url), name, dateStr, when: dateStr ? parseHearingDateTime(dateStr) : null, filedBy, status, result });
  }
  return rows;
}

/* The default prove-up packet for an OSC Re: Failure to Prosecute Default
   Judgment.

   The default itself is taken by a CIV-100 whose Result column reads
   "Entered" — that is the `base`. The prove-up is whatever asks the court to
   ENTER JUDGMENT on that default, and it takes more than one shape:

     - a SECOND "Request for Entry of Default / Judgment" filed afterward (the
       shape this used to assume was the only one);
     - the SAME CIV-100, when its own title says it asked for a court or clerk
       judgment — eCourt appends the relief in brackets, "Request for Entry of
       Default / Judgment [Court Judgment]". One paper both entered the default
       and requested judgment, so there is no second request to wait for, and
       insisting on one reported "No Default Packet" over a complete packet;
     - the judgment papers themselves — the proposed default judgment (JUD-100),
       the § 585(d) declaration, a prove-up declaration — which a plaintiff may
       file without any further CIV-100.

   Once the court has already taken the OSC up — a Minute Order for an Order to
   Show Cause Re: Failure to Prosecute Default Judgment (or a Default Prove-Up
   Hearing, same flow) is on the docket — the packet is gated to papers filed
   after the FIRST such minute order. The OSC exists because judgment was never
   entered despite whatever was filed before it, so a CIV-100 [Court Judgment]
   or judgment papers predating the court's first look are the stale attempt
   the OSC is about, not the packet answering it; reporting them said
   "✓ Default Packet" over a docket the court had already found wanting. With
   no such minute order on the docket (the OSC's first hearing) nothing is
   gated. The earliest minute order is the floor, not the latest, so papers the
   court has seen at a continuance still open for review at the next hearing.

   Returns { base, proveUp, oscFloor, oscLast }: `base` is the most recent
   entered request (null when the default was never entered — a separate fact
   the OSC status reports on its own), `proveUp` is the earliest judgment paper
   on or after it that clears the minute-order gate, `proveUp.when` is the
   floor for the rest of the packet, and `oscFloor` / `oscLast` are the dates
   of the first and most recent OSC minute orders (null when none is on the
   docket; `oscLast` lets the status say "no update since the court last
   looked"). Null when the docket shows no base, no prove-up, and no OSC
   minute order. */
const DEFAULT_REQUEST_RE = /request for entry of default/i;

// eCourt names the relief a CIV-100 asked for in brackets. A request for a
// COURT or CLERK judgment is itself the judgment request; a bare "Request for
// Entry of Default / Judgment" only took the default.
const JUDGMENT_REQUEST_RE = /\b(?:court|clerk)'?s?\s+judgment\b/i;

// A paper that asks the court to enter judgment on an entered default, as
// opposed to the request that took the default.
function isDefaultJudgmentPacketDoc(name) {
  const n = (name || '').trim();
  if (!n) return false;
  // A court minute order is never the plaintiff's paper — "Minute Order
  // ( (Default Prove Up Hearing) )" carries the prove-up phrase in its
  // hearing parenthetical without being part of the packet.
  if (/^\s*minute\s+order\b/i.test(n)) return false;
  if (DEFAULT_REQUEST_RE.test(n)) return JUDGMENT_REQUEST_RE.test(n);
  if (/\bdefault\s+prove[-\s]?up\b/i.test(n)) return true;
  // The proposed judgment itself (JUD-100), however eCourt qualifies it.
  if (/^(?:\[?\s*proposed\s*\]?\s*)?(?:default\s+judgment|judgment\s+by\s+default)\b/i.test(n)) return true;
  // "Declaration Pursuant to 585 CCP in Support of Default Judgment", and the
  // other declarations filed to prove the judgment up.
  if (/\bdeclaration\b/i.test(n) && /\b(?:585|default\s+judgment|prove[-\s]?up)\b/i.test(n)) return true;
  return false;
}

function findDefaultProveUp(docs) {
  const list = (docs || []).filter(d => d && d.when);
  // base = most recent Request for Entry of Default whose Result reads "Entered".
  let base = null;
  for (const d of list) {
    if (!DEFAULT_REQUEST_RE.test(d.name || '')) continue;
    if (!/\bentered\b/i.test(d.result || '')) continue;
    if (!base || d.when > base.when) base = d;
  }
  // The minute-order gate (see the block comment above): the date of the FIRST
  // Minute Order memorializing a hearing of this OSC. Same-day papers stay in —
  // a packet filed the day of the hearing answers it, and midnight-normalized
  // dates can't order the two within the day anyway.
  let oscFloor = null, oscLast = null;
  for (const d of list) {
    if (!/^\s*minute\s+order\b/i.test(d.name || '')) continue;
    if (!OSC_DEFAULT_JUDGMENT_RE.test(d.name || '')) continue;
    if (!oscFloor || d.when < oscFloor) oscFloor = d.when;
    if (!oscLast || d.when > oscLast) oscLast = d.when;
  }
  // proveUp = the earliest judgment paper. A plain second CIV-100 counts by
  // POSITION (filed after the default was entered, so it can only be the
  // judgment request); everything else counts by what its title says it is, so
  // a packet on a docket whose default was never entered is still reported.
  let proveUp = null;
  for (const d of list) {
    if (!d.openUrl) continue;
    if (base && d.when < base.when) continue;   // predates the default it proves up
    if (oscFloor && d.when < oscFloor) continue; // predates the court's first look — the stale attempt
    const secondRequest = !!base && d.docId !== base.docId && DEFAULT_REQUEST_RE.test(d.name || '');
    if (!secondRequest && !isDefaultJudgmentPacketDoc(d.name)) continue;
    if (!proveUp || d.when < proveUp.when) proveUp = d;
  }
  if (!base && !proveUp && !oscFloor) return null;
  return { base, proveUp, oscFloor, oscLast };
}

// POSTs the eCourt tree-table pager to fetch a page of documents at `offset`,
// using the pageData token read from the live Documents page. Returns a parsed
// fragment (wrapped in a table so rows survive) or null.
async function postPanelPage(offset, pageData) {
  try {
    const body = 'offset=' + offset + '&pageData=' + encodeURIComponent(pageData);
    const res = await Promise.race([
      fetch('/ecourt/ecms/forms/support/onPanelPage', {
        method: 'POST', credentials: 'include',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' },
        body,
      }),
      new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 8000)),
    ]);
    if (!res || !res.ok) return null;
    const html = await res.text();
    return new DOMParser().parseFromString('<table><tbody>' + html + '</tbody></table>', 'text/html');
  } catch (_) { return null; }
}

// Fetches ALL documents for a case: page 1 plus every subsequent page via the
// pager (so old pleadings like the complaint are included, not just the newest).
async function fetchAllDocuments(docsUrl) {
  const doc1 = await fetchCaseDoc(docsUrl);
  if (!doc1) return [];
  const rows = parseDocRows(doc1);
  const seen = new Set(rows.map(r => r.docId));

  let pageData = null, count = 0;
  const pgEl = doc1.querySelector('[data-ec-pagedata]');
  if (pgEl) { pageData = pgEl.getAttribute('data-ec-pagedata'); try { count = (JSON.parse(pageData) || {}).count || 0; } catch (_) {} }

  if (pageData && count > rows.length) {
    // Known total: fetch every remaining page in parallel.
    const offsets = [];
    for (let o = 50; o < count && o <= 2000; o += 50) offsets.push(o);
    const results = await Promise.all(offsets.map(o => postPanelPage(o, pageData)));
    for (const frag of results) {
      if (!frag) continue;
      for (const r of parseDocRows(frag)) if (!seen.has(r.docId)) { seen.add(r.docId); rows.push(r); }
    }
  } else if (pageData && count < 0 && rows.length > 0) {
    // Unknown total: eCourt returns count: -1 for lazily-counted lists, so the
    // "count > rows.length" test above never fires and older filings (e.g. the
    // original complaint) on later pages are missed. Page forward sequentially
    // until a page adds nothing new, with a safety cap.
    for (let o = 50; o <= 2000; o += 50) {
      const frag = await postPanelPage(o, pageData);
      if (!frag) break;
      let added = 0;
      for (const r of parseDocRows(frag)) if (!seen.has(r.docId)) { seen.add(r.docId); rows.push(r); added++; }
      if (!added) break; // reached the end (or the pager stopped advancing)
    }
  }
  return rows;
}

/* ------------------------------------------------------------------ */
/* Document ingest time — when eCourt actually POSTED a filing         */
/*                                                                     */
/* The Documents tab shows only the filing DATE, which is the paper's  */
/* effective date. A reply e-filed late on the day it is due carries   */
/* that date (and is timely), but does not appear on the site until    */
/* the clerk's intake queue reaches it — typically the next court day, */
/* sometimes two or three.                                             */
/*                                                                     */
/* The doc endpoint's Last-Modified header carries the real posting    */
/* time, mangled: the server hands an epoch value in SECONDS to an API */
/* expecting MILLISECONDS, so every document dates to January 1970.    */
/* Multiplying the parsed epoch-milliseconds back up by 1000 recovers  */
/* the true instant, whose UTC fields read as the court's PACIFIC wall */
/* clock — confirmed against the PDF /ModDate stamps on documents that */
/* carry one, which land within minutes of the decoded window.         */
/*                                                                     */
/* The header renders whole seconds, so one header-second spans 1000   */
/* real seconds: the true moment falls in a ~16.7-minute window. That  */
/* is enough to place a filing on a day and a part of a day, which is  */
/* all the lag question needs.                                         */
/* ------------------------------------------------------------------ */

const INGEST_WINDOW_MS = 999 * 1000;
// A Last-Modified parsing below this is the 1970 artifact (a genuine modern
// date parses to ~1.7e12). Should the court ever fix the bug, the header is
// then a real instant and is used as-is.
const INGEST_ARTIFACT_MAX_MS = 1e11;

// Decode a Last-Modified header into { start, end, decoded }, or null. When
// `decoded` is true, read the Dates with their UTC getters — those fields are
// the court's Pacific wall clock, and the underlying instant is meaningless.
function decodeIngestTime(lastModified) {
  const t = Date.parse(lastModified || '');
  if (!isFinite(t) || t <= 0) return null;
  if (t >= INGEST_ARTIFACT_MAX_MS) return { start: new Date(t), end: new Date(t), decoded: false };
  const ms = t * 1000;
  return { start: new Date(ms), end: new Date(ms + INGEST_WINDOW_MS), decoded: true };
}

// The calendar fields of an ingest window's start, in whichever clock applies.
function ingestParts(info) {
  const d = info.start;
  return info.decoded
    ? { y: d.getUTCFullYear(), mo: d.getUTCMonth(), da: d.getUTCDate(), h: d.getUTCHours(), mi: d.getUTCMinutes() }
    : { y: d.getFullYear(), mo: d.getMonth(), da: d.getDate(), h: d.getHours(), mi: d.getMinutes() };
}

// The ingest day as a local midnight Date, so it can be compared against filing
// dates and run through the court-day helpers.
function ingestDay(info) {
  if (!info) return null;
  const p = ingestParts(info);
  return new Date(p.y, p.mo, p.da);
}

// "Thu 8/6 10:43a" — the day and part-of-day, which is the resolution we have.
const INGEST_DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
function fmtIngest(info) {
  if (!info) return '';
  const p = ingestParts(info);
  const day = new Date(p.y, p.mo, p.da);
  const h12 = p.h % 12 || 12;
  const mm = p.mi < 10 ? '0' + p.mi : String(p.mi);
  return INGEST_DOW[day.getDay()] + ' ' + (p.mo + 1) + '/' + p.da + ' ' + h12 + ':' + mm + (p.h < 12 ? 'a' : 'p');
}

// Court days from `from` to `to`, counting the days after `from` up to and
// including `to`. Weekends and court holidays don't count.
function courtDaysBetween(from, to) {
  if (!from || !to) return null;
  const a = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  const b = new Date(to.getFullYear(), to.getMonth(), to.getDate());
  if (b < a) return -courtDaysBetween(to, from);
  let n = 0;
  const d = new Date(a);
  while (d < b) { d.setDate(d.getDate() + 1); if (DL.isCourtDay(d)) n++; }
  return n;
}

// How long a document sat between its filing date and eCourt posting it,
// in court days. Null when either end is unknown.
function ingestLagCourtDays(filedWhen, info) {
  const day = ingestDay(info);
  return (filedWhen && day) ? courtDaysBetween(filedWhen, day) : null;
}

/* Ingest times never change once a document is posted, so they're cached
   permanently in chrome.storage.local under one key per docId. A Documents tab
   costs one HEAD per uncached document on its first visit and nothing after. */
const INGEST_KEY = id => 'ingest:' + id;
const INGEST_MEM = new Map(); // docId -> {start,end,decoded} | null | Promise

function ingestFromStore(docId) {
  return new Promise(res => {
    try {
      chrome.storage.local.get([INGEST_KEY(docId)], r => {
        if (chrome.runtime.lastError) return res(undefined);
        const v = r && r[INGEST_KEY(docId)];
        res(typeof v === 'string' ? decodeIngestTime(v) : undefined);
      });
    } catch (_) { res(undefined); }
  });
}

// HEAD is enough — we only want the header, never the PDF (some run to
// megabytes). A server that rejects HEAD gets a one-byte ranged GET instead.
async function headIngest(docId) {
  const url = '/ecourt/ecms/doc?docId=' + encodeURIComponent(docId);
  const tries = [
    { method: 'HEAD' },
    { method: 'GET', headers: { Range: 'bytes=0-0' } },
  ];
  for (const opt of tries) {
    try {
      const res = await Promise.race([
        fetch(url, Object.assign({ credentials: 'include' }, opt)),
        new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), 8000)),
      ]);
      if (!res || !res.ok) continue;
      const lm = res.headers.get('last-modified');
      if (lm) return lm;
    } catch (_) { /* try the next shape */ }
  }
  return null;
}

// The ingest window for one document, or null if the server didn't say.
function getIngestTime(docId) {
  const id = String(docId || '');
  if (!id) return Promise.resolve(null);
  let v = INGEST_MEM.get(id);
  if (v !== undefined) return Promise.resolve(v);
  const p = (async () => {
    const cached = await ingestFromStore(id);
    if (cached !== undefined) return cached;
    const lm = await headIngest(id);
    const info = lm ? decodeIngestTime(lm) : null;
    if (lm) { try { chrome.storage.local.set({ [INGEST_KEY(id)]: lm }, () => { void chrome.runtime.lastError; }); } catch (_) {} }
    return info;
  })();
  INGEST_MEM.set(id, p);
  p.then(r => INGEST_MEM.set(id, r), () => INGEST_MEM.set(id, null));
  return p;
}

/* ---- The intake queue, as it bears on a paper that isn't there yet ---- */

// How many court days after a filing date a paper can still be in the intake
// queue rather than missing. Measured at 0-3 court days (median 1) across the
// sample the options page accumulates; 2 covers all but the tail.
const INGEST_GRACE_COURT_DAYS = 2;

// True when a paper due on `due` could be filed-but-not-yet-posted: its due
// date has passed, but not by enough court days to clear the intake queue.
// Used to hedge an absence rather than call the paper missing.
function withinIngestGrace(due, now) {
  if (!due) return false;
  const today = now || new Date();
  const dd = dayMs(due);
  if (dd == null || dd >= dayMs(today)) return false; // not yet due: not an absence at all
  const lag = courtDaysBetween(due, today);
  return lag != null && lag <= INGEST_GRACE_COURT_DAYS;
}

/* ------------------------------------------------------------------ */
/* Post-judgment motions: new trial / JNOV and reconsideration         */
/*                                                                     */
/* These don't brief off the hearing date, so §1005 doesn't reach them */
/* — their clocks run from papers already on the docket:               */
/*                                                                     */
/*   CCP 659a  moving papers 10 days after the notice of intention is  */
/*             FILED; opposition 10 days after service of that brief;  */
/*             reply 5 days after service of the opposition. A judge   */
/*             may extend for good cause up to 10 more days, which is  */
/*             an order we can't see, so the dates here are the        */
/*             unextended ones.                                        */
/*   CCP 660   the court's power to rule expires 75 days after the     */
/*             earliest of the clerk's mailing of notice of entry      */
/*             (§664.5) or a party's service of written notice of      */
/*             entry; absent either, 75 days after the first notice of */
/*             intention. Undetermined by then, the motion is denied   */
/*             by operation of law.                                    */
/*   CCP 1008  reconsideration within 10 days after service of written */
/*             notice of entry of THE ORDER — a different anchor.      */
/*                                                                     */
/* The periods running from SERVICE carry the electronic-service       */
/* extension (§1010.6(a)(3)(B), 2 court days), consistent with the     */
/* rest of the widget; the 659a moving-papers period runs from FILING  */
/* and carries none. Docket filing dates stand in for service dates —  */
/* the docket doesn't record when a paper was served.                  */
/* ------------------------------------------------------------------ */

const NOTICE_OF_INTENTION_RE = /\bnotice\s+of\s+inten(?:t|tion)\b/i;
const NEW_TRIAL_SUBJECT_RE = /\bnew\s+trial\b|\bjnov\b|judgment\s+notwithstanding/i;
// Anchored at the start of the name on purpose. §660 runs from the clerk's
// MAILING of notice of entry (§664.5) or a party's SERVICE of written notice of
// entry — so the notice itself, or the clerk's certificate of mailing it. A
// "Minute Order (Court Order re: Notice of Entry of Judgment)" merely mentions
// the notice; an unanchored match picked it up, and since the earliest match
// wins, a docket where that order predates the notice anchored the whole
// 75-day computation on the wrong paper.
// DO NOT widen this to a bare clerk-filed "Judgment" on the rule 8.104 analogy.
// The two clocks are not parallel and the Supreme Court has said so directly.
//
// Van Beurden Ins. Services v. Customized Worldwide Weather Ins. Agency (1997)
// 15 Cal.4th 51, 64: the clerk mailed a file-stamped judgment with proof of
// service and it did NOT start § 660's period — to be a mailing "pursuant to
// Section 664.5" the notice must affirmatively state it is given "upon order of
// the court" or "under section 664.5". A file-endorsed copy does neither.
// Palmer v. GTE California (2003) 30 Cal.4th 1265, 1274 reaffirmed that clerk
// holding, and at 1277 answered the appeal-rule analogy itself: rule 8.104 "at
// most confirms our view that the requirements of section 664.5 . . . exceed
// those of sections 659 and 660". The rule was amended in 2002 to add a
// file-stamped copy as a triggering document; §§ 660 and 664.5 never were.
//
// So the same clerk-served paper can start the 60-day appeal clock (and the fee
// clock that borrows it) while not starting this one. What Palmer DID relax is
// the party prong: a file-stamped copy served by a party suffices there.
const NOTICE_OF_ENTRY_JUDGMENT_RE = /^(?:certificate of mailing for\s+)?notice of entry of (?:the\s+)?judgment\b/i;
const NOTICE_OF_ENTRY_ORDER_RE = /^(?:certificate of mailing for\s+)?notice of (?:entry of (?:the\s+)?order|ruling)\b/i;
const PJ_SERVICE_EXT_COURT_DAYS = 2; // §1010.6(a)(3)(B), electronic service

function pjAddCal(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function pjAddCourtDays(d, n) {
  const r = new Date(d); let rem = n;
  while (rem > 0) { r.setDate(r.getDate() + 1); if (DL.isCourtDay(r)) rem--; }
  return r;
}
// CCP 12/12a: a period ending on a holiday runs to the next court day.
function pjRoll(d) { const r = new Date(d); while (!DL.isCourtDay(r)) r.setDate(r.getDate() + 1); return r; }
function pjFromFiling(anchor, days) { return pjRoll(pjAddCal(anchor, days)); }
function pjFromService(anchor, days) { return pjRoll(pjAddCourtDays(pjAddCal(anchor, days), PJ_SERVICE_EXT_COURT_DAYS)); }

function pjEarliest(list) {
  return (list || []).filter(d => d && d.when).sort((a, b) => a.when - b.when)[0] || null;
}

// The schedule for a post-judgment hearing, from the anchoring papers on the
// docket. Null when the anchor isn't there — then there's nothing to compute
// from and the widget falls back to reporting whether the motion is on file.
function computePostJudgmentSchedule(cat, docs) {
  if (cat === 'recon') {
    const notice = pjEarliest((docs || []).filter(d => NOTICE_OF_ENTRY_ORDER_RE.test(d.name || '')));
    if (!notice) return null;
    return { kind: 'recon', anchor: notice, motionDue: pjFromService(notice.when, 10) };
  }
  const intent = pjEarliest((docs || []).filter(d =>
    NOTICE_OF_INTENTION_RE.test(d.name || '') && NEW_TRIAL_SUBJECT_RE.test(d.name || '')));
  if (!intent) return null;
  const out = { kind: 'new_trial', anchor: intent, motionDue: pjFromFiling(intent.when, 10) };
  // Party prong, per Palmer: a file-stamped copy served BY A PARTY is written
  // notice of entry for § 660, with no captioned document required. eCourt shows
  // that as a party-filed "Judgment" — but the same docket line is also what a
  // prevailing party lodging a proposed judgment produces, and the title cannot
  // tell them apart. Admitted anyway, because the error runs the safe way: an
  // extra trigger only makes the period EARLIER, while missing a real one leaves
  // the court believing it still has power it has lost. The tooltip flags it.
  const noticed = pjEarliest((docs || []).filter(d => NOTICE_OF_ENTRY_JUDGMENT_RE.test(d.name || '')));
  const partyCopy = pjEarliest((docs || []).filter(d =>
    ENTRY_JUDGMENT_DOC_RE.test(d.name || '') && !/\bclerk\b/i.test(d.filedBy || '')));
  const entry = (noticed && partyCopy)
    ? (noticed.when <= partyCopy.when ? noticed : partyCopy)
    : (noticed || partyCopy);
  // Deliberately NOT rolled forward: §660 is a jurisdictional cutoff, not a
  // filing deadline, so the 75th day is reported as it falls.
  const from = entry || intent;
  // Rolled forward under CCP 12a. § 660(c) opens "Except as otherwise provided
  // in Section 12a of this code", so the statute incorporates the holiday
  // extension rather than standing apart from it — the fact that the period
  // bounds the COURT's power rather than a party's filing doesn't take it
  // outside § 12a, because § 660 says so itself.
  const raw = pjAddCal(from.when, 75);
  out.ruleExpires = pjRoll(raw);
  out.ruleExpiresRaw = raw;
  out.expiresAnchor = from;
  out.expiresFrom = entry ? 'notice of entry of judgment' : 'the first notice of intention';
  // True when the anchor is a party-filed judgment rather than a captioned
  // notice of entry — the Palmer prong, and the one the docket can't confirm.
  // Flagged only when no proof of service backs it: with one on file the
  // rule's own condition is met and there is nothing left uncertain.
  out.expiresFromPartyCopy = !!(entry && entry === partyCopy && entry !== noticed
    && !hasAccompanyingProofOfService(entry, docs));
  return out;
}

/* A party's service of a filed-endorsed judgment is a trigger under both rule
   8.104(a)(1)(B) and — per Palmer — CCP § 660's party prong. eCourt shows it as
   a party-filed "Judgment", which is also exactly what LODGING a proposed
   judgment produces; the title cannot separate them. What can be checked is the
   condition rule 8.104(a)(1)(B) states outright: a proof of service filed the
   same day by the same party. Present, the document is a served copy and there
   is nothing to flag. Absent, it is admitted anyway and flagged, because the
   reader can resolve in a glance what the docket cannot. */
const PROOF_OF_SERVICE_RE = /^proof of (?:personal |substituted )?service\b/i;

function hasAccompanyingProofOfService(doc, docs) {
  if (!doc || !doc.when) return false;
  const party = docPartyNames(doc.filedBy);
  return (docs || []).some(p => p !== doc && p.name && p.when
    && PROOF_OF_SERVICE_RE.test(p.name)
    && sameCalendarDay(p.when, doc.when)
    && (!party.length || !p.filedBy || docSharesParty(docPartyNames(p.filedBy), party)));
}

/* ---- Appeal-time trigger, and the fee deadline that borrows it ----------
   CRC 3.1702(b)(1): a notice of motion claiming fees for services through
   rendition of judgment is served and filed within the time for filing a NOTICE
   OF APPEAL — rule 8.104(a)(1), the earliest of 60 days from the clerk's
   service of a "Notice of Entry" or a filed-endorsed copy of the judgment,
   60 days from a party's service of either with proof of service, or 180 days
   from entry. Rule 8.104(e) makes an appealable order the judgment, so a
   clerk's "Order - Dismissal" after a settlement is a trigger.

   No service extension: § 1013(a) and § 1010.6(a)(3)(B) both exclude the notice
   of appeal. Rule 8.108 and a rule 3.1702(b)(2) stipulation can extend the
   period and neither is visible from the docket, so this is the unextended
   date. Docket filing dates stand in for service dates. */
const APPEAL_NOTICE_OF_ENTRY_RE = /notice of (entry|ruling)/i;
const APPEAL_DISMISSAL_RE = /^order\s*[-\u2013\u2014:]?\s*dismissal\b|^order of dismissal\b|notice of entry of dismissal\b/i;
const APPEAL_JUDGMENT_RE = /^(?:amended\s+)?judgment\b/i;

// The paper that started the appeal clock, and whether the docket can confirm
// it. Rule 8.104(a)(1) takes the EARLIEST of the triggering events, bounded to
// one appeal period so an older judgment can't claim a later motion's trigger.
function findAppealTimeTrigger(docs, cutoff) {
  const triggers = (docs || []).filter(d => d.name && d.when
    && (!cutoff || d.when <= cutoff)
    && (APPEAL_NOTICE_OF_ENTRY_RE.test(d.name) || APPEAL_DISMISSAL_RE.test(d.name)
        || APPEAL_JUDGMENT_RE.test(d.name)));
  if (!triggers.length) return null;
  let doc = triggers.slice().sort((a, b) => b.when - a.when)[0];
  const windowStart = new Date(doc.when); windowStart.setDate(windowStart.getDate() - 60);
  for (const d of triggers) if (d.when >= windowStart && d.when < doc.when) doc = d;
  // A party-filed judgment with no proof of service beside it may be a lodged
  // proposed judgment rather than the served copy (a)(1)(B) requires.
  const unverified = APPEAL_JUDGMENT_RE.test(doc.name)
    && !/\bclerk\b/i.test(doc.filedBy || '')
    && !hasAccompanyingProofOfService(doc, docs);
  return { doc, unverified };
}

function computeFeeMotionStatus(docs, motionWhen) {
  const trig = findAppealTimeTrigger(docs);
  if (!trig) return null;
  const due = pjRoll(pjAddCal(trig.doc.when, 60));
  const entry = pjEarliest((docs || []).filter(d => ENTRY_JUDGMENT_DOC_RE.test(d.name || '')));
  const outer = entry ? pjRoll(pjAddCal(entry.when, 180)) : null;
  const effective = (outer && outer < due) ? outer : due;
  return {
    trigger: trig.doc, unverified: trig.unverified, due: effective,
    basis: (effective === due ? '60 days after service of ' + trig.doc.name : '180 days after entry of judgment'),
    motionWhen,
    late: motionWhen != null && dayMs(motionWhen) > dayMs(effective),
  };
}

/* ---- Memorandum of costs (CRC 3.1700(a)(1)) ------------------------------
   The prevailing party's memorandum is due on the FIRST of: 15 days after
   service of the notice of entry of judgment or dismissal (the clerk's under
   §664.5, or a party's written notice), or 180 days after entry. Only the
   15-day period runs from service, so only it carries the §1010.6(a)(3)(B)
   electronic-service extension the widget assumes throughout.

   The trigger here is the notice of entry specifically — narrower than rule
   8.104's, which also accepts a filed-endorsed copy of the judgment. Rule
   3.1700(a)(1) names the §664.5 notice, so the bare judgment is deliberately
   not admitted as a trigger.

   Docket filing dates stand in for service dates, which eCourt doesn't record. */
const COSTS_MEMO_RE = /^memorandum of costs\b/i;
const COSTS_TRIGGER_RE = /^(?:certificate of mailing for\s+)?notice of entry of (?:the\s+)?(?:judgment|dismissal)\b/i;
const ENTRY_JUDGMENT_DOC_RE = /^(?:amended\s+)?judgment\b/i;

function computeCostsMemoStatus(docs) {
  const memo = pjEarliest((docs || []).filter(d => COSTS_MEMO_RE.test(d.name || '')));
  if (!memo) return null;
  const trigger = pjEarliest((docs || []).filter(d => COSTS_TRIGGER_RE.test(d.name || '')));
  const entry = pjEarliest((docs || []).filter(d => ENTRY_JUDGMENT_DOC_RE.test(d.name || '')));
  const due15 = trigger ? pjFromService(trigger.when, 15) : null;
  const outer = entry ? pjRoll(pjAddCal(entry.when, 180)) : null;
  const due = (due15 && outer) ? (due15 <= outer ? due15 : outer) : (due15 || outer);
  if (!due) return { memo, due: null, late: false };
  return {
    memo, due, trigger, entry,
    basis: due === due15
      ? '15 days after service of the notice of entry, plus the electronic-service extension'
      : '180 days after entry of judgment',
    late: dayMs(memo.when) > dayMs(due),
  };
}

// The paper a post-judgment motion runs from: the notice of intention for new
// trial / JNOV, the notice of entry of the order for reconsideration. It is a
// floor as well as a clock — nothing filed before it briefs the motion, so the
// Documents button uses it to bound what it opens. Null for any other motion
// category, or when the anchoring paper isn't on the docket.
function postJudgmentAnchor(motionType, docs) {
  const cat = DL.classifyMotion(motionType || '');
  if (cat !== 'new_trial' && cat !== 'recon') return null;
  const sched = computePostJudgmentSchedule(cat, docs);
  return sched ? sched.anchor : null;
}

// California motion-deadline engine, inlined so the content script is
// self-contained (no dependency on a separate file loading first).
// KEEP IN SYNC with lib/deadlines.js, which the Deadline Calculator page uses.
const DL = (function () {
  const holidayCache = {};
  function getHolidays(year) {
    if (holidayCache[year]) return holidayCache[year];
    const h = new Set();
    const obs = d => { const day = d.getDay();
      if (day === 6) { const f = new Date(d); f.setDate(f.getDate() - 1); return f; }
      if (day === 0) { const m = new Date(d); m.setDate(m.getDate() + 1); return m; }
      return d; };
    const fixed = (mo, da) => obs(new Date(year, mo, da));
    const nth = (mo, wd, n) => { let d = new Date(year, mo, 1), c = 0;
      while (d.getMonth() === mo) { if (d.getDay() === wd && ++c === n) return d; d.setDate(d.getDate() + 1); } };
    const last = (mo, wd) => { let d = new Date(year, mo + 1, 0); while (d.getDay() !== wd) d.setDate(d.getDate() - 1); return d; };
    const key = d => d ? `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}` : null;
    const add = d => { if (d) h.add(key(d)); };
    add(fixed(0, 1)); add(nth(0, 1, 3)); add(fixed(1, 12)); add(nth(1, 1, 3)); add(fixed(2, 31));
    add(last(4, 1)); add(fixed(5, 19)); add(fixed(6, 4)); add(nth(8, 1, 1)); add(nth(8, 5, 4));
    add(fixed(10, 11)); const tg = nth(10, 4, 4); add(tg);
    if (tg) { const da = new Date(tg); da.setDate(da.getDate() + 1); add(da); }
    add(fixed(11, 25));
    const nyNext = obs(new Date(year + 1, 0, 1)); if (nyNext.getFullYear() === year) add(nyNext);
    holidayCache[year] = h; return h;
  }
  function isCourtDay(d) { const dow = d.getDay(); if (dow === 0 || dow === 6) return false;
    return !getHolidays(d.getFullYear()).has(`${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`); }
  function nextCourtDay(d) { const r = new Date(d); while (!isCourtDay(r)) r.setDate(r.getDate() + 1); return r; }
  function prevCourtDay(d) { const r = new Date(d); while (!isCourtDay(r)) r.setDate(r.getDate() - 1); return r; }
  function addCD(d, n) { const r = new Date(d), step = n >= 0 ? 1 : -1; let rem = Math.abs(n);
    while (rem > 0) { r.setDate(r.getDate() + step); if (isCourtDay(r)) rem--; } return r; }
  function addCAL(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
  function stdMotion(hearing, svc) { let d = addCD(hearing, -16);
    if (svc === 'electronic') d = addCD(d, -2); else if (svc === 'mail_ca') d = addCAL(d, -5);
    else if (svc === 'mail_state') d = addCAL(d, -10); else if (svc === 'mail_conf') d = addCAL(d, -12);
    else if (svc === 'mail_intl') d = addCAL(d, -20); else if (svc === 'fax') d = addCAL(d, -2);
    return prevCourtDay(d); }
  function msjMotion(hearing, svc) { let d = addCAL(hearing, -81);
    if (svc === 'electronic') d = addCD(d, -2); else if (svc === 'mail_ca') d = addCAL(d, -5);
    else if (svc === 'mail_state') d = addCAL(d, -10); else if (svc === 'mail_conf') d = addCAL(d, -5);
    else if (svc === 'mail_intl') d = addCAL(d, -20); else if (svc === 'fax') d = addCD(d, -2);
    return prevCourtDay(d); }
  function stdOpp(h)   { return prevCourtDay(addCD(h, -9));  }
  function msjOpp(h)   { return prevCourtDay(addCAL(h, -20)); }
  function stdReply(h) { return prevCourtDay(addCD(h, -5));  }
  function msjReply(h) { return prevCourtDay(addCAL(h, -11)); }
  // Unlawful detainer MSJ — CCP § 1170.7: made at any time after the answer, on
  // FIVE days' notice; CRC 3.1351(a) has that notice given in compliance with
  // §§ 1013 and 1170.7, so the ordinary service extensions ride on the 5 days.
  function udMsjMotion(hearing, svc) { let d = addCAL(hearing, -5);
    if (svc === 'electronic') d = addCD(d, -2); else if (svc === 'mail_ca') d = addCAL(d, -5);
    else if (svc === 'mail_state') d = addCAL(d, -10); else if (svc === 'mail_conf') d = addCAL(d, -12);
    else if (svc === 'mail_intl') d = addCAL(d, -20); else if (svc === 'fax') d = addCAL(d, -2);
    return prevCourtDay(d); }
  // CRC 3.1351(b)-(c): opposition may be made orally at the hearing; a WRITTEN
  // opposition to be considered in advance is filed and served on or before the
  // court day before the hearing.
  function udMsjOpp(h)   { return prevCourtDay(addCAL(h, -1)); }
  // No advance written reply schedule — the reply may be made orally at the
  // hearing (CRC 3.1351(b)), so the slot dates to the hearing day itself.
  function udMsjReply(h) { return prevCourtDay(h); }

  /* A motion's title routinely names ANOTHER motion it merely relates to: a
     motion to seal is captioned "Motion to Seal the Exhibits of CCTV Video in
     Support of Defendant's Motion for Summary Judgment." The relief sought is
     sealing — an ordinary § 1005 motion — but the nested reference to the MSJ
     put it on § 437c's 81-day clock and reported a timely motion as LATE. Cut
     the title at the clause that introduces the other motion; what is left is
     this motion's own relief.

     Only the "in <support|opposition|…> of/to" forms are cut, so a paper
     captioned from the start as a response ("Opposition to Motion for Summary
     Judgment") keeps its whole title. */
  function stripAncillaryMotionReference(mt) {
    return (mt || '')
      .replace(/\s+(?:filed\s+)?in\s+(?:support|opposition|connection|conjunction|response|reply|regard|regards|relation)\s+(?:of|to|with)\b.*$/i, ' ')
      .replace(/\s+/g, ' ').trim();
  }

  /* A motion to seal is a § 1005 motion however the papers it seals are
     described — including where the title names them with no "in support of"
     clause to cut ("Motion to Seal Exhibits to Defendant's Motion for Summary
     Judgment"). The relief a title LEADS with is the motion's own, so the test
     is anchored: sealing has to be what the caption reaches within its opening
     words, not a word appearing anywhere in it. */
  const SEAL_MOTION_RE = /^[^;]{0,40}?\b(?:seal|sealing|unseal|unsealing)\b/i;

  function classifyMotion(mt) { const s = stripAncillaryMotionReference(mt).toLowerCase();
    if (SEAL_MOTION_RE.test(s)) return 'standard';
    if (/summary\s+judgment|summary\s+adjudication|\bmsj\b|\bmsa\b/.test(s)) return 'msj';
    if (/new\s+trial|\bjnov\b|judgment\s+notwithstanding|vacate\s+(the\s+)?judgment/.test(s)) return 'new_trial';
    if (/reconsideration|renewed?\s+motion|\bccp?\s*1008\b|\b1008\b/.test(s)) return 'recon';
    // Fees before costs: "strike or tax costs" is a costs motion, but "fees and
    // costs" is a fee motion, and the fee test is the narrower one.
    if (/attorney'?s?\s+fees|\battorney\s+fee\b|\bfees\s+and\s+costs\b|\b3\.1702\b/.test(s)) return 'fees';
    if (/\b(?:strike|tax|taxing)\s+(?:of\s+)?costs\b|\bcosts\b.*\b(?:strike|tax)\b|memorandum\s+of\s+costs|\b3\.1700\b/.test(s)) return 'costs';
    return 'standard'; }
  return { classifyMotion, stripAncillaryMotionReference, stdMotion, msjMotion, stdOpp, msjOpp, stdReply, msjReply,
    udMsjMotion, udMsjOpp, udMsjReply, isCourtDay };
})();

// The title-trimming half of classifyMotion, shared with the case page so the
// motion-type tests there read the same relief the deadline engine does.
const stripAncillaryMotionReference = DL.stripAncillaryMotionReference;

const DL_DEBUG = true;

function dlLog() { if (DL_DEBUG) { try { console.log.apply(console, ['[LACourt-DL]'].concat([].slice.call(arguments))); } catch (_) {} } }

function fmtShortDate(d) {
  return (!d || isNaN(d)) ? '' : d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: '2-digit' });
}

function dayMs(d) { return (d && !isNaN(d)) ? new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() : null; }

const NEXT_DL_GAP = '<span style="display:inline-block;width:20px"></span>';

const DL_YELLOW = '#b8860b'; // late for e-service but timely if personally served

const DL_NEUTRAL = '#0a6e6e'; // due date not yet passed, nothing on file
const DL_BLACK = '#000000';   // nothing to say yet: filing status unknown, or a reply not yet due

// Colour: green = filed on time; red = overdue and nothing timely filed; neutral
// = not yet due; black = we don't know yet.
//
// The dates paint as soon as they're computed, but whether a paper was FILED
// takes a background fetch of the case Documents. Until that lands, a passed due
// date says nothing about the paper — so it reads black, and turns red only once
// the fetch confirms nothing timely is on file. (Colouring from the date alone
// made every overdue-looking paper flash red on load, then go green a moment
// later for the ones that had been filed on time.)
function nextDlColor(due, filed, filedKnown) {
  const dd = dayMs(due);
  if (dd == null) return DL_NEUTRAL;
  if (!filedKnown) return DL_BLACK;
  const fm = dayMs(filed);
  if (fm != null && fm <= dd) return '#1a6b3a';
  // Nothing on file, but the due date passed too recently for the clerk's
  // intake queue to have posted a timely filing — that's not a red yet.
  if (fm == null && withinIngestGrace(due)) return DL_NEUTRAL;
  return dd < dayMs(new Date()) ? '#c0392b' : DL_NEUTRAL;
}

// Motion-specific colour. Like nextDlColor, but when a filed motion missed the
// electronic-service deadline yet lands on/before the (more lenient) personal-
// service deadline, it shows YELLOW instead of red — a cue to check the proof of
// service, since personal service (no notice extension) would make it timely.
function motionDlColor(elecDue, personalDue, filed, filedKnown) {
  const dd = dayMs(elecDue);
  if (dd == null) return DL_NEUTRAL;
  if (!filedKnown) return DL_BLACK;                        // filing status not in yet
  const fm = dayMs(filed);
  if (fm != null) {
    if (fm <= dd) return '#1a6b3a';                        // timely for electronic service
    const pd = dayMs(personalDue);
    if (pd != null && fm <= pd) return DL_YELLOW;          // timely only if personally served
    return '#c0392b';                                      // late even for personal service
  }
  if (withinIngestGrace(elecDue)) return DL_NEUTRAL;       // may be filed but not posted yet
  return dd < dayMs(new Date()) ? '#c0392b' : DL_NEUTRAL;
}

// Escapes quotes as well as markup: these strings go into title="..." attributes
// as often as into text, and a document name carrying a quote silently truncated
// the attribute (and everything after it) rather than failing loudly.
function dlEsc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// Triggers the default-judgment flow: either an OSC Re: Failure to Prosecute
// Default Judgment, or a Default Prove-Up Hearing — both are worked up the same
// way (prove-up packet documents, recommendation email, fee calculator). The OSC
// branch must include "Default Judgment"; the OSC for other reasons (sanctions,
// etc.) doesn't go through this flow.
const OSC_DEFAULT_JUDGMENT_RE = /\border\s+to\s+show\s+cause\s+re:?\s+failure\s+to\s+prosecute\s+default\s+judgment\b|\bprove\s*-?\s*up\b/i;

function isOscDefaultJudgment(hearingType) {
  if (!hearingType) return false;
  return OSC_DEFAULT_JUDGMENT_RE.test(hearingType);
}

/* ------------------------------------------------------------------ */
/* Case-type designation ("Civil Unlimited Unlawful Detainer/…")       */
/* ------------------------------------------------------------------ */
//
// eCourt prints the case-type designation under the case number, e.g.
// "Civil Unlimited Unlawful Detainer/Commercial ..." (routinely truncated).
// The lead is what makes it safe to detect: no filing title or hearing name
// starts "Civil Unlimited"/"Civil Limited", so an anchored test never reads a
// document that merely mentions a case type as the case's own designation.
const CASE_TYPE_LEAD_RE = /^\s*civil\s+(?:unlimited|limited)\b/i;
// The designation and its type read together (bounded, so a page-wide text
// blob that happens to contain both far apart doesn't match).
const UD_CASE_TYPE_RE = /\bcivil\s+(?:unlimited|limited)\b[\s\S]{0,80}?\bunlawful\s+detainer\b/i;

function isUnlawfulDetainerTypeText(text) {
  return UD_CASE_TYPE_RE.test(text || '');
}

// The element carrying the case-type line: the SMALLEST element whose text is
// the designation, searched in the case-header blocks first (the designation
// sits under the case number, near the [class*="case"] element the case name
// lives in) and the whole document only as a fallback. Truncation-tolerant —
// the truncated line still starts with the lead.
function findCaseTypeEl(root) {
  root = root || (typeof document !== 'undefined' ? document : null);
  if (!root || !root.querySelectorAll) return null;
  const scopes = [];
  const seen = new Set();
  for (const el of root.querySelectorAll('[class*="case"]')) {
    for (let n = el, up = 0; n && n.querySelectorAll && up < 3; n = n.parentElement, up++) {
      if (!seen.has(n)) { seen.add(n); scopes.push(n); }
    }
  }
  const whole = root.body || root;
  if (whole && !seen.has(whole)) scopes.push(whole);
  for (const scope of scopes) {
    let best = null, bestLen = Infinity;
    for (const el of scope.querySelectorAll('span, div, td, p, li, b, strong, h1, h2, h3, h4')) {
      const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
      if (!t || t.length > 200 || !CASE_TYPE_LEAD_RE.test(t)) continue;
      if (t.length < bestLen) { best = el; bestLen = t.length; }
    }
    if (best) return best;
  }
  return null;
}

function findCaseTypeLine(root) {
  const el = findCaseTypeEl(root);
  return el ? (el.textContent || '').replace(/\s+/g, ' ').trim() : '';
}

// Is the case an unlawful detainer, per the designation on the page. The
// truncated line still carries "Unlawful Detainer", so this works whether or
// not the designation is cut off.
function isUnlawfulDetainerCase(root) {
  return isUnlawfulDetainerTypeText(findCaseTypeLine(root));
}

/* ------------------------------------------------------------------ */
/* Case context: where one case's Documents / Hearings / Parties come  */
/* from, fetched at most once each.                                    */
/* ------------------------------------------------------------------ */

// eCourt case sub-tabs are the same page with a different formId.
const DOCS_FORM_ID = '279';
const HEARINGS_FORM_ID = '395';

// A sub-tab URL read off a case document by tab label ("parties", "documents").
// Parties has no reliable formId, so its URL has to come from the tab link.
function caseTabUrlFrom(root, label) {
  try {
    for (const a of root.querySelectorAll('a[href*="/ecourt/ecms/case"]')) {
      const href = a.getAttribute('href') || '';
      if (href && (a.textContent || '').trim().toLowerCase() === label) {
        return new URL(href, location.origin).href;
      }
    }
  } catch (_) {}
  return null;
}

function emptyDoc() {
  return new DOMParser().parseFromString('<html><body></body></html>', 'text/html');
}

// Builds a case context. Each source can be given as a URL (string or a thunk
// returning one, possibly async) or as a ready-made loader — the case page
// passes its own memoized getters so the widget and the Documents button share
// one fetch. Every source degrades to empty rather than throwing.
//
//   ctx.docs()     -> [{ docId, name, when, filedBy, result, openUrl }, …]
//   ctx.hearings() -> [{ type, date, when }, …]  (future + scheduled)
//   ctx.parties()  -> a Document to read the parties table out of
function makeCaseCtx(o) {
  const memo = fn => { let p = null; return () => (p || (p = fn())); };
  const urlOf = async v => (typeof v === 'function' ? await v() : v);
  return {
    docs: o.docs || memo(async () => {
      try { const u = await urlOf(o.docsUrl); return u ? await fetchAllDocuments(u) : []; }
      catch (_) { return []; }
    }),
    hearings: o.hearings || memo(async () => {
      try {
        const u = await urlOf(o.hearingsUrl);
        const doc = u ? await fetchCaseDoc(u) : null;
        return doc ? parseFutureHearings(doc) : [];
      } catch (_) { return []; }
    }),
    parties: o.parties || memo(async () => {
      try {
        const u = await urlOf(o.partiesUrl);
        const doc = u ? await fetchCaseDoc(u) : null;
        return doc || emptyDoc();
      } catch (_) { return emptyDoc(); }
    }),
  };
}

// Context for a case known only by its eCourt id (the agenda's case links).
// Documents and Hearings are addressable by formId; Parties needs the case page
// fetched first to read its tab link, and only an OSC status ever asks for it.
function caseCtxForId(caseId) {
  const base = '/ecourt/ecms/case?id=' + encodeURIComponent(caseId);
  return makeCaseCtx({
    docsUrl: base + '&formId=' + DOCS_FORM_ID,
    hearingsUrl: base + '&formId=' + HEARINGS_FORM_ID,
    partiesUrl: async () => {
      const doc = await fetchCaseDoc(base);
      return doc ? caseTabUrlFrom(doc, 'parties') : null;
    },
  });
}

// Briefing deadlines for ONE hearing. `eff` is a resolved hearing
// { motionType, hearingType, hearingDate, lookedAhead } — read off the case
// page's Next header, off the Hearings tab, or off an agenda row. Returns
// { osc } for an OSC Re: Failure to Prosecute Default Judgment, { skip } for a
// motion whose deadlines aren't keyed to the hearing date, null when there's no
// motion to brief, else the dates. `eff` is echoed back when the hearing was
// looked ahead to, so the caller can label which hearing the figures belong to.
function computeDueDatesFor(eff) {
  const D = DL;
  const hearingType = eff.hearingType;
  const tag = eff.lookedAhead ? { eff } : {};
  // An OSC Re: Failure to Prosecute Default Judgment gets a default/dismissal
  // status indicator instead of briefing deadlines.
  if (isOscDefaultJudgment(hearingType)) return Object.assign({ osc: true }, tag);
  // Only "Hearing on <motion>" events are §1005/§437c-briefed. Non-motion
  // events (CMC, OSC, trial, status conference) have no "Hearing on" prefix.
  const motionType = eff.motionType;
  if (!motionType) return null; // not a motion (or the header hasn't rendered yet)
  const cat = D.classifyMotion(motionType);
  // New trial / JNOV / reconsideration aren't hearing-based: their deadlines run
  // from notice of entry of judgment or the original order, not the hearing, so
  // there is no §1005 schedule to show. Whether the moving papers ever arrived
  // is a separate fact and still worth reporting — a hearing whose motion was
  // never filed comes off calendar — so these carry motionOnly rather than
  // dropping out entirely.
  if (cat === 'new_trial' || cat === 'recon') {
    return Object.assign({
      skip: true, motionOnly: true, reason: cat, motionType, cat,
      hearingWhen: parseHearingDateTime(eff.hearingDate),
    }, tag);
  }
  const hearing = parseHearingDateTime(eff.hearingDate);
  if (!hearing) { dlLog('no hearing date parsed for motion:', motionType); return null; }
  // In an unlawful detainer (eff.ud, from the case-type designation) summary
  // judgment is not on § 437c's 81/20/11-day schedule: CCP § 1170.7 gives it
  // five days' notice, and CRC 3.1351 puts the written opposition on the court
  // day before the hearing (or orally at it). Only the MSJ changes — the
  // parties' other motions keep their ordinary schedules.
  const ud = cat === 'msj' && !!eff.ud;
  const motionFn = svc => cat !== 'msj' ? D.stdMotion(hearing, svc)
    : (ud ? D.udMsjMotion(hearing, svc) : D.msjMotion(hearing, svc));
  return Object.assign({
    skip: false, motionType, cat, ud, hearingWhen: hearing,
    motionDue: motionFn('electronic'),
    // Same deadline assuming PERSONAL service, which carries no notice extension
    // (electronic adds 2 court days). A motion filed after the electronic
    // deadline but on/before this one would still be timely if personally served
    // — the widget flags that in yellow so the proof of service can be checked.
    motionDuePersonal: motionFn('personal'),
    oppDue:    cat === 'msj' ? (ud ? D.udMsjOpp(hearing) : D.msjOpp(hearing)) : D.stdOpp(hearing),
    replyDue:  cat === 'msj' ? (ud ? D.udMsjReply(hearing) : D.msjReply(hearing)) : D.stdReply(hearing),
  }, tag);
}

// When the deadlines/OSC belong to a later hearing (the Next event wasn't one we
// work up), lead with that hearing's type + date so it's clear which hearing
// these figures are for.
function effHearingPrefix(c) {
  const eff = c && c.eff;
  if (!eff) return '';
  const when = eff.hearingDate ? parseHearingDateTime(eff.hearingDate) : null;
  const label = stripHearingOnPrefix(eff.hearingType || '') + (when ? ' ' + fmtShortDate(when) : '');
  return '<span style="color:#0a6e6e">▸ ' + dlEsc(label.trim()) + ':</span>' + NEXT_DL_GAP;
}

// The status text for one hearing: `c` from computeDueDatesFor, `filed` from
// computeFiledStatus (null until it resolves — the dates then show in their
// date-only colours), `osc` from computeOscStatus.
function statusHtml(c, filed, osc) {
  const prefix = effHearingPrefix(c);
  if (c.osc) {
    const st = osc;
    if (!st) return prefix + '<span style="color:#0a6e6e">Checking defaults…</span>';
    let html = prefix + '<span style="color:' + st.color + '">' + st.text + '</span>';
    if (st.packetText) {
      html += NEXT_DL_GAP + '<span style="color:' + st.packetColor + '">' + st.packetText + '</span>';
    }
    return html;
  }
  const f = filed || {};
  // The petition IS the pleading in a case with no complaint — nothing to brief.
  if (f.petitionIsPleading) return '';
  const RED = '#c0392b';
  // A motion whose deadlines don't run from the hearing (new trial, JNOV,
  // reconsideration). There is no briefing schedule to paint, so the only thing
  // to say is whether the moving papers are on the docket at all — and the only
  // thing worth saying is when they aren't.
  const item = (label, key, due) =>
    `<span style="color:${nextDlColor(due, f[key], f.filedKnown)}">${label} ${fmtShortDate(due)}</span>`;
  // The reply is the one paper whose absence is routine until its date arrives —
  // most motions never draw one. So while it isn't due yet it reads BLACK rather
  // than the neutral teal the other papers use, keeping the widget's colour
  // vocabulary for what's actually been filed (green), missed (red), or flagged
  // (yellow). Once the date passes unfiled it becomes "No Reply (Due …)" in red.
  const replyItem = due => {
    const c0 = nextDlColor(due, f.reply, f.filedKnown);
    const col = c0 === DL_NEUTRAL ? DL_BLACK : c0;
    return `<span style="color:${col}">Reply Due ${fmtShortDate(due)}</span>`;
  };
  // A paper is "absent" (never filed) — as opposed to filed-but-late — only when
  // the Documents fetch succeeded (filedKnown), found no matching filing, and the
  // due date has already passed. A late filing keeps its "Due <date>" in red.
  const absent = (due, filedDate) => {
    if (!f.filedKnown || filedDate != null) return false;
    const dd = dayMs(due);
    return dd != null && dd < dayMs(new Date());
  };
  // An absent paper shows "No Opposition (Due <date>)" / "No Reply (Due <date>)".
  // Except while the due date is recent enough that a timely filing might still
  // be sitting in the clerk's intake queue: eCourt posts a paper 0-3 court days
  // after its filing date, so a reply filed the day it was due routinely isn't
  // visible until the next court day. Inside that window the absence is
  // reported as unconfirmed, in neutral teal, rather than called missing in red.
  const absentSpan = (noun, due) => {
    if (withinIngestGrace(due)) {
      const t = `Due ${fmtShortDate(due)} — nothing on file yet, but eCourt posts filings up to `
        + `${INGEST_GRACE_COURT_DAYS} court days after they are filed, so a timely ${noun.toLowerCase()} may not be visible yet.`;
      return `<span style="color:${DL_NEUTRAL}" title="${dlEsc(t)}">No ${noun} Posted Yet (Due ${fmtShortDate(due)})</span>`;
    }
    return `<span style="color:${RED}">No ${noun} (Due ${fmtShortDate(due)})</span>`;
  };
  if (c.motionOnly) {
    if (!f.filedKnown) return '';
    const pj = f.pj;
    if (!pj) {
      // No anchoring paper on the docket, so there is no schedule to compute.
      // All that can be said is whether the moving papers arrived.
      if (f.motion != null) return '';
      return prefix + `<span style="color:${RED}" title="No moving papers for this hearing are on the docket, `
        + `and no notice of intention or notice of entry to compute deadlines from.">No Motion</span>`;
    }
    const pjAbsent = (due, when) => f.filedKnown && when == null && dayMs(due) != null && dayMs(due) < dayMs(new Date());
    const pjParts = [];
    const cite = pj.kind === 'recon' ? 'CCP 1008(a), 10 days after service of notice of entry of the order'
      : 'CCP 659a, 10 days after the notice of intention was filed';
    pjParts.push(pjAbsent(pj.motionDue, f.motion)
      ? absentSpan('Motion', pj.motionDue)
      : `<span style="color:${nextDlColor(pj.motionDue, f.motion, f.filedKnown)}" title="${dlEsc(cite)}">Motion Due ${fmtShortDate(pj.motionDue)}</span>`);
    // Only meaningful once the paper they run from is actually on file.
    if (f.motion != null && pj.oppDue) {
      pjParts.push(pjAbsent(pj.oppDue, f.opp)
        ? absentSpan('Opposition', pj.oppDue)
        : `<span style="color:${nextDlColor(pj.oppDue, f.opp, f.filedKnown)}">Opposition Due ${fmtShortDate(pj.oppDue)}</span>`);
    }
    if (f.opp != null && pj.replyDue) {
      if (pjAbsent(pj.replyDue, f.reply)) {
        pjParts.push(absentSpan('Reply', pj.replyDue));
      } else {
        const rc = nextDlColor(pj.replyDue, f.reply, f.filedKnown);
        pjParts.push(`<span style="color:${rc === DL_NEUTRAL ? DL_BLACK : rc}">Reply Due ${fmtShortDate(pj.replyDue)}</span>`);
      }
    }
    // The jurisdictional cutoff. Past it and the motion stands denied by
    // operation of law, which is the whole point of showing it.
    if (pj.ruleExpires) {
      const gone = dayMs(pj.ruleExpires) < dayMs(new Date());
      // Name the paper it counted from: the anchor is inferred from document
      // titles, so the figure should be checkable without reading the code.
      const from = pj.expiresAnchor
        ? '"' + pj.expiresAnchor.name + '" filed ' + fmtShortDate(pj.expiresAnchor.when)
        : pj.expiresFrom;
      // A clerk-filed notice only qualifies if the served document itself recites
      // § 664.5 or a court order (Van Beurden, supra, 15 Cal.4th at p. 64) —
      // which the docket title cannot show. Say so, because if it doesn't
      // recite, the period runs instead from party service or, absent that,
      // from the first notice of intention, and this date is then too early.
      const clerkAnchored = pj.expiresAnchor
        && /\bclerk\b/i.test(pj.expiresAnchor.filedBy || '')
        && pj.expiresFrom === 'notice of entry of judgment';
      const partyCopyCaveat = pj.expiresFromPartyCopy
        ? ' Anchored on a party-filed judgment: under Palmer (2003) 30 Cal.4th 1265, 1274 a file-stamped copy SERVED by a party is written notice of entry, but the docket cannot show whether this one was served or merely lodged — check the proof of service. If it was not served, the period runs from the first notice of intention instead.'
        : '';
      const caveat = clerkAnchored
        ? ' The clerk\'s notice starts this period only if it affirmatively states it is given under § 664.5 or upon order of the court (Van Beurden (1997) 15 Cal.4th 51, 64) — check the served document.'
        : '';
      const rolled = pj.ruleExpiresRaw && dayMs(pj.ruleExpiresRaw) !== dayMs(pj.ruleExpires)
        ? ', which falls on ' + fmtShortDate(pj.ruleExpiresRaw) + ' — a non-court day, carried to the next court day under CCP 12a'
        : '';
      const t = 'CCP 660(c): the court\'s power to rule expires 75 days after ' + from + rolled
        + (gone ? '. That date has passed — undetermined, the motion is denied by operation of law.' : '.')
        + caveat + partyCopyCaveat;
      pjParts.push(`<span style="color:${gone ? RED : DL_NEUTRAL}" title="${dlEsc(t)}">`
        + (gone ? 'Power to Rule Expired ' : 'Power to Rule Expires ') + fmtShortDate(pj.ruleExpires) + '</span>');
    }
    return prefix + pjParts.join(NEXT_DL_GAP);
  }
  const oppAbsent = absent(c.oppDue, f.opp);
  // A hearing on calendar whose moving papers never arrived — the reserved date
  // the party abandoned. This state was always being reported, but only as a red
  // "Motion Due <date>"; it now reads "No Motion (Due <date>)" like the other
  // papers, since the point is that nothing was filed, not that a date passed.
  const motionAbsent = absent(c.motionDue, f.motion);

  // The Motion uses its own colour so a late-but-maybe-personally-served filing
  // reads yellow (a cue to check the proof of service) rather than red.
  const motionColor = motionDlColor(c.motionDue, c.motionDuePersonal, f.motion, f.filedKnown);
  const motionTitle = motionColor === DL_YELLOW
    ? ' title="Late for electronic service, but timely if personally served — check the proof of service."'
    : '';
  const motionSpan = motionAbsent
    ? absentSpan('Motion', c.motionDue)
    : `<span style="color:${motionColor}"${motionTitle}>Motion Due ${fmtShortDate(c.motionDue)}</span>`;

  const parts = [motionSpan];
  // The UD schedule looks nothing like § 437c's, so say which rules the dates
  // come from — otherwise a 5-days-out MSJ deadline reads like a bug.
  if (c.ud) {
    parts.unshift('<span style="color:' + DL_NEUTRAL + '" title="'
      + dlEsc('Unlawful detainer: the MSJ is heard on 5 days’ notice (CCP 1170.7), the written opposition is due '
        + 'the court day before the hearing, and the opposition or reply may instead be made orally at the hearing (CRC 3.1351).')
      + '">UD §1170.7:</span>');
  }
  // A memorandum of costs filed after its own deadline is the fact worth having
  // in front of you on a motion to strike or tax it — the memo may be untimely
  // on its face. Shown ONLY when it is late; a timely memo is the ordinary case
  // and would just be another date to read past.
  // A fee motion filed after rule 3.1702(b)(1)'s deadline. The § 1005 dates
  // beside it are the notice period for the HEARING and say nothing about
  // whether the motion could be brought at all, so the timeliness question
  // leads the line. Shown only when late, like the costs memorandum.
  const fd = f.feeDeadline;
  if (fd && fd.late && fd.due) {
    const t = 'Fee motion filed ' + fmtShortDate(fd.motionWhen || fd.filedWhen || new Date(0))
      + '. Rule 3.1702(b)(1) gives it the time for filing a notice of appeal — '
      + fd.basis + ', so ' + fmtShortDate(fd.due) + '. '
      + (fd.unverified ? 'The trigger is a party-filed judgment with no proof of service beside it, so check whether it was served. ' : '')
      + 'Rule 8.108 (post-trial motion pending) or a rule 3.1702(b)(2) stipulation can extend the period; neither shows on the docket.';
    parts.unshift(`<span style="color:${RED}" title="${dlEsc(t)}">Fee Motion Late `
      + `(due ${fmtShortDate(fd.due)})</span>`);
  }
  const cm = f.costsMemo;
  if (cm && cm.late && cm.due) {
    const t = 'Memorandum of costs filed ' + fmtShortDate(cm.memo.when) + ', after the '
      + fmtShortDate(cm.due) + ' deadline (' + cm.basis + ' — CRC 3.1700(a)(1)). '
      + 'The docket shows filing, not service, so check the proof of service.';
    parts.unshift(`<span style="color:${RED}" title="${dlEsc(t)}">Costs Memo Late `
      + `(${fmtShortDate(cm.memo.when)}, due ${fmtShortDate(cm.due)})</span>`);
  }
  // With no moving papers on file there is nothing to brief, so the opposition
  // and reply deadlines are moot — the same reasoning that drops the reply when
  // no opposition was filed. A motion that is merely still in the clerk's intake
  // queue keeps its schedule: those deadlines are days out and remain useful.
  if (motionAbsent && !withinIngestGrace(c.motionDue)) return prefix + parts.join(NEXT_DL_GAP);
  const nonOpp = f.nonOpp || null;
  const nonOppSpan = '<span style="color:#1a6b3a">Notice of Non-Opposition</span>';
  // A demurrer or motion to strike answered by an amended complaint in lieu of
  // opposition (CCP 472): the amended pleading moots the challenge, so show it in
  // place of the Opposition/Reply slots rather than "No Opposition". RED, unlike
  // the other "this paper is accounted for" entries — it isn't briefing to read,
  // it's a hearing to take off calendar.
  if (f.fac && isDemurrerOrStrikeMotion(c.motionType)) {
    parts.push(`<span style="color:${RED}">${dlEsc(f.fac.label)}</span>`);
  } else if (nonOpp && nonOpp.slot === 'opp') {
    // The non-moving party filed a Notice of Non-Opposition — it takes the place
    // of (and is more meaningful than) its opposition, and moots the reply.
    parts.push(nonOppSpan);
  } else {
    parts.push(oppAbsent
      ? absentSpan('Opposition', c.oppDue)
      : item('Opposition Due', 'opp', c.oppDue));
    if (nonOpp && nonOpp.slot === 'reply') {
      // The moving party filed a Notice of Non-Opposition noting the absence of
      // any opposition — it takes the place of the reply.
      parts.push(nonOppSpan);
    } else if (!oppAbsent) {
      // With no opposition on file, the reply deadline is irrelevant — drop it.
      parts.push(absent(c.replyDue, f.reply)
        ? absentSpan('Reply', c.replyDue)
        : replyItem(c.replyDue));
    }
  }
  return prefix + parts.join(NEXT_DL_GAP);
}

// Which briefing papers are on file for hearing `c`, and when — what turns the
// dates into filed/late/absent colours. Best-effort: any failure resolves with
// filedKnown false and the caller keeps the date-only colours.
async function computeFiledStatus(ctx, c) {
  const filed = { filedKnown: false, motion: null, opp: null, reply: null, nonOpp: null };
  try {
    {
      const docs = await ctx.docs();
      if (docs && docs.length) {
        filed.filedKnown = true;
        // No complaint on the docket, so this petition is the case's initiating
        // PLEADING rather than a motion (see docsHaveComplaint). A hearing on a
        // pleading has no § 1005 schedule to paint and no missing moving paper
        // to report — statusHtml renders nothing at all, the way it does for a
        // conference or a trial.
        if (isPetitionMotion(c.motionType) && !docsHaveComplaint(docs)) {
          filed.petitionIsPleading = true;
          return filed;
        }
        const earliest = list => list.slice().sort((a, b) => a.when - b.when)[0] || null;
        const hearings = await ctx.hearings();

        // On a motion to strike or tax costs, whether the memorandum it attacks
        // was itself timely is worth knowing before reading the motion.
        if (c.cat === 'costs') filed.costsMemo = computeCostsMemoStatus(docs);
        // A fee motion's real deadline is rule 3.1702(b)(1)'s, not the § 1005
        // notice period for its hearing — and the docket can date it.
        if (c.cat === 'fees') {
          const fm = resolveMovingPaper(c.motionType, c.hearingWhen, hearings, docs);
          filed.feeDeadline = computeFeeMotionStatus(docs, fm ? fm.when : null);
        }

        // Post-judgment motions brief off the docket, not the hearing date, so
        // their schedule is built here where the documents are in hand.
        if (c.motionOnly) {
          const sched = computePostJudgmentSchedule(c.cat, docs);
          filed.pj = sched;
          const pjMd = resolveMovingPaper(c.motionType, c.hearingWhen, hearings, docs);
          filed.motion = pjMd ? pjMd.when : null;
          // 659a chains off what was actually served: no moving papers, no
          // opposition period; no opposition, no reply period.
          if (sched && sched.kind === 'new_trial' && pjMd) {
            sched.oppDue = pjFromService(pjMd.when, 10);
            const o = earliest(docs.filter(d => d.when && d.when >= pjMd.when && isOppositionDoc(d.name)));
            filed.opp = o ? o.when : null;
            if (o) {
              sched.replyDue = pjFromService(o.when, 5);
              const r = earliest(docs.filter(d => d.when && d.when >= o.when && /\breply\b/i.test(d.name)));
              filed.reply = r ? r.when : null;
            }
          }
          return filed;
        }

        // Identify the moving paper, disambiguating parallel same-named
        // demurrers/motions to strike by pairing them to their hearings by date
        // (see resolveMovingPaper). bestFilingMatch alone grabs the newest, which
        // for a case with a demurrer to the complaint AND one to a cross-complaint
        // mis-dates the motion and its timeliness.
        const md = resolveMovingPaper(c.motionType, c.hearingWhen, hearings, docs);
        filed.motion = md ? md.when : null;
        const mw = md ? md.when : null;
        const after = docs.filter(d => d.when && (!mw || d.when >= mw));

        // When exactly one live motion has its moving papers on file, an
        // Opposition/Reply on the docket can only belong to it — so match by
        // title + position (Opposition after the motion, Reply after the
        // Opposition) without needing the name to reference the motion. With
        // multiple filed motions, require the name to link to THIS motion so we
        // don't attribute the wrong paper.
        const workable = hearings.filter(h => isWorkableHearing(h.type));
        // Each workable hearing whose moving papers are on file, carrying the
        // party that brought it: WHO moved is what tells a generically titled
        // paper apart from the motions its own name can't name (below).
        const filedMotions = [];
        for (const h of workable) {
          const hm = resolveMovingPaper(h.type, h.when, hearings, docs);
          if (hm) filedMotions.push({ when: h.when, movant: docPartyNames(hm.filedBy) });
        }
        const singleMotion = !!md && filedMotions.length === 1;

        // The movant files the reply, so a "Reply …" by the same party as the
        // motion also counts (covers replies named only by party/abbreviation).
        const movantParties = md ? docPartyNames(md.filedBy) : [];
        const sharesMovant = d => movantParties.length && docSharesParty(docPartyNames(d.filedBy), movantParties);
        const filedByNonMovant = d => {
          const p = docPartyNames(d.filedBy);
          return !!(p.length && movantParties.length && !docSharesParty(p, movantParties));
        };

        /* A generically titled paper ("Opposition OPPOSITION", "Reply REPLY")
           names no motion, so the name test below can never link it and the
           multi-motion branch dropped it. What the name can't say, the calendar
           and the filer can: briefing is filed for the NEXT hearing, so such a
           paper belongs to this motion when this motion's hearing is the
           earliest filed-motion hearing on or after the filing day.

           Motions heard that SAME day compete for it, but two of them are only
           genuinely rival when they come from different sides. Papers filed
           together by one moving party — a demurrer and its companion motion to
           strike, two discovery motions on one calendar — are one side's
           papers, and a generic opposition to them opposes this motion whichever
           of them it meant to name. Treating that as ambiguous is what reported
           "No Opposition" with the opposition sitting on the docket, on the most
           ordinary multi-motion calendar there is.

           Cross-motions set on one date (each side moving) stay separable, by
           the filer: a party neither opposes nor replies to its own motion, so
           drop the same-day motions brought by this paper's filer and the paper
           belongs to this motion if it is the one left standing. Anything still
           spanning two moving parties is left unattributed rather than guessed
           at, as before. */
        const genericPaperIsThisMotions = d => {
          if (!docNameIsGeneric(d.name) || !d.when || !c.hearingWhen) return false;
          const upcoming = filedMotions.filter(h => h.when && dayMs(h.when) >= dayMs(d.when));
          if (!upcoming.length) return false;
          const nearest = Math.min(...upcoming.map(h => dayMs(h.when)));
          if (nearest !== dayMs(c.hearingWhen)) return false;
          const sameDay = upcoming.filter(h => dayMs(h.when) === nearest);
          if (sameDay.length <= 1) return true;
          const isThisSide = h => !!(h.movant.length && movantParties.length
            && docSharesParty(h.movant, movantParties));
          // All of that day's filed motions are this motion's own side's papers.
          if (sameDay.every(isThisSide)) return true;
          // Otherwise let the filer eliminate the motions it cannot be aimed at.
          const filer = docPartyNames(d.filedBy);
          if (!filer.length) return false;
          const aimedAt = sameDay.filter(h => !(h.movant.length && docSharesParty(h.movant, filer)));
          return aimedAt.length === 1 && isThisSide(aimedAt[0]);
        };

        let o, r;
        if (singleMotion) {
          o = earliest(after.filter(d => isOppositionDoc(d.name)));
          const afterOpp = o ? o.when : mw;
          r = earliest(docs.filter(d => d.when && (!afterOpp || d.when >= afterOpp) && /\breply\b/i.test(d.name)));
        } else {
          // The movant does not oppose its own motion, so a generic opposition
          // filed by the movant is opposing something else; a generic reply,
          // conversely, is the movant's paper and can't come from the other side.
          o = earliest(after.filter(d => isOppositionDoc(d.name)
            && (docLinksToMotion(d.name, c.motionType)
                || (genericPaperIsThisMotions(d) && !sharesMovant(d)))));
          r = earliest(after.filter(d => /\breply\b/i.test(d.name)
            && (docLinksToMotion(d.name, c.motionType) || sharesMovant(d)
                || (genericPaperIsThisMotions(d) && !filedByNonMovant(d)))));
        }
        filed.opp = o ? o.when : null;
        filed.reply = r ? r.when : null;

        // A "Notice of Non-Opposition" / "No Opposition" stands in for a briefing
        // paper, matched by WHO filed it (only meaningful when no real opposition
        // is on the docket):
        //   - filed by the NON-moving party, it affirmatively states that party
        //     does not oppose — more meaningful than, and taking the place of, its
        //     Opposition, so it fills the Opposition slot (and moots the reply);
        //   - filed by the MOVING party, it notes the absence of any opposition
        //     and takes the place of the reply — but only once the opposition
        //     deadline has run. Before that day there is no non-opposition to
        //     notice: the opposition is not late and may still be filed. So a
        //     movant's notice dated on or before c.oppDue is premature as a
        //     report of the other side's silence and is not read as one. (The
        //     non-movant's is a statement about ITSELF — a party may concede
        //     whenever it likes — and carries no such cutoff.)
        // Multi-motion cases still require the notice to link to THIS motion (or,
        // for the movant's, be filed by the movant) so it isn't misattributed.
        if (!o) {
          const nonOpps = after.filter(d => isNonOppositionDoc(d.name)
            && (singleMotion || docLinksToMotion(d.name, c.motionType) || sharesMovant(d)
                || genericPaperIsThisMotions(d)));
          /* Which paper the notice stands in for turns on WHO filed it, so it
             takes an identified filer on BOTH sides of the comparison. With no
             movant to compare against, "not the movant" is not a finding —
             reading it as one dropped every notice into the Opposition slot on
             any motion whose moving paper carries no Filed By, and reported the
             movant's own notice as the other side's concession. */
          const byNonMovant = movantParties.length ? earliest(nonOpps.filter(d => {
            const p = docPartyNames(d.filedBy);
            return p.length && !docSharesParty(p, movantParties);
          })) : null;
          if (byNonMovant) {
            filed.nonOpp = { slot: 'opp', when: byNonMovant.when };
          } else {
            /* The movant's own notice (or one whose filer we can't identify),
               dropped when it predates the opposition deadline it purports to
               report on: a caption that names a motion for summary judgment is
               enough to link such a paper to an MSJ still weeks from its
               opposition date. */
            const byMovant = earliest(nonOpps.filter(d => {
              const p = docPartyNames(d.filedBy);
              if (p.length && movantParties.length && !docSharesParty(p, movantParties)) return false;
              if (c.oppDue && d.when && dayMs(d.when) <= dayMs(c.oppDue)) {
                dlLog('non-opposition notice filed', d.when, 'is premature — opposition not due until',
                  c.oppDue, ':', d.name);
                return false;
              }
              return true;
            }));
            if (byMovant) filed.nonOpp = { slot: 'reply', when: byMovant.when };
          }
        }

        // Demurrer or motion to strike answered by a first amended complaint in
        // lieu of opposition (CCP 472): the of-right amendment moots the
        // challenge, so show it in place of the opposition. Three conditions:
        //   - the current hearing is a demurrer / motion to strike;
        //   - the plaintiff did NOT oppose (in lieu = INSTEAD of opposition) — if
        //     an opposition is on file, that's what we show, not the amendment;
        //   - a FIRST amended complaint (only the one of-right amendment, never an
        //     SAC) was filed AFTER the current challenge (mw). One filed on/before
        //     mw is the challenged pleading, not a response — e.g. a demurrer to
        //     the FAC, whose FAC predates it. Pairing mw to the right challenge
        //     above is what keeps a cross-complaint's demurrer from polluting this.
        if (isDemurrerOrStrikeMotion(c.motionType) && !filed.opp && mw) {
          const fac = earliest(docs.filter(d => d.when && d.when > mw && isFirstAmendedComplaintDoc(d.name)));
          filed.fac = fac ? { label: 'First Amended Complaint', when: fac.when } : null;
        }
      }
    }
  } catch (_) { /* keep date-only colours */ }
  return filed;
}

// For an OSC Re: Failure to Prosecute Default Judgment, report whether every
// (non-cross) defendant is either in default or dismissed. A default counts
// only if it was entered on/after the operative (latest) complaint — an earlier
// default was superseded by the amended complaint and must be re-taken.
async function computeOscStatus(ctx) {
  try {
    const parties = parsePartiesTable(await ctx.parties());
    // A party on the claimant side is never a defendant who could be in
    // default, even if some other row labels it "Respondent" — the plaintiff of
    // a case on appeal is listed as the appellate Respondent, and a "Respondent"
    // row is otherwise a writ/petition defendant. Appeal-subcase rows are
    // already dropped by parsePartiesTable; this also covers a subcase heading
    // we don't recognize.
    const claimants = new Set(parties
      .filter(p => /^\s*(plaintiff|petitioner|cross[-\s]?complainant)\b/i.test(p.role || ''))
      .map(p => movantNormName(p.name)));
    const defendants = parties.filter(p => {
      const r = p.role || '';
      if (/cross[-\s]?defendant/i.test(r) || !/^\s*(defendant|respondent)\b/i.test(r)) return false;
      return !claimants.has(movantNormName(p.name));
    });
    if (!defendants.length) return { text: 'No defendants found', color: '#0a6e6e' };

    const docs = await ctx.docs();
    // Operative pleading = latest complaint, falling back to the latest petition
    // when the case has no complaint (mirrors computeRelevantDocuments).
    let op = latestDoc((docs || []).filter(d => isComplaintDoc(d.name)));
    if (!op) op = latestDoc((docs || []).filter(d => isPetitionDoc(d.name)));
    const opWhen = op ? op.when : null;

    // Default prove-up packet: a request for default judgment (second Request
    // for Entry of Default, filed after the one that entered the default).
    const pu = findDefaultProveUp(docs);
    const hasPacket = !!(pu && pu.proveUp);
    // With no packet answering the OSC and the court having looked at least
    // once (a minute order of this OSC on the docket), distinguish "nothing
    // has happened since" from "no packet": when no party paper postdates the
    // most recent minute order, the operative fact for the next hearing is
    // that nothing changed — report that, with the date, instead of "No
    // Default Packet" (which reads as if the docket were bare even when a
    // stale, already-considered packet sits on it). Court-generated papers —
    // the minute order's own certificate of mailing, a signed order, the OSC
    // itself — are not an update; any party filing is, packet or not.
    const oscLast = (pu && pu.oscLast) || null;
    const isCourtPaper = d => /^\s*(?:minute\s+order\b|order\b|certificate\s+of\b|clerk'?s\b)/i.test((d && d.name) || '');
    const noUpdate = !hasPacket && oscLast
      && !(docs || []).some(d => d && d.when && d.when > oscLast && !isCourtPaper(d));
    const packet = {
      packetText: hasPacket ? '✓ Default Packet'
        : noUpdate ? '⚠ No Update Since Last MO (' + fmtShortDate(oscLast) + ')'
        : '⚠ No Default Packet',
      packetColor: hasPacket ? '#1a6b3a' : '#c0392b',
    };
    dlLog('OSC default packet:', {
      hasPacket,
      noUpdate: !!noUpdate,
      base: pu && pu.base && { name: pu.base.name, when: fmtShortDate(pu.base.when), result: pu.base.result },
      proveUp: pu && pu.proveUp && { name: pu.proveUp.name, when: fmtShortDate(pu.proveUp.when) },
      oscFloor: (pu && pu.oscFloor && fmtShortDate(pu.oscFloor)) || null,
      oscLast: (oscLast && fmtShortDate(oscLast)) || null,
    });

    const notAccounted = [];
    let dismissed = 0, defaulted = 0, stale = 0;
    for (const d of defendants) {
      if (d.dismissed) { dismissed++; continue; }
      const dd = d.defaultDate;
      if (dd && opWhen && dd >= opWhen) { defaulted++; continue; }
      if (dd && opWhen && dd < opWhen) stale++;
      notAccounted.push(d.name);
    }
    const N = defendants.length, plural = N === 1 ? '' : 's';
    dlLog('OSC default status:', {
      N, defaulted, dismissed, stale, opWhen: opWhen && fmtShortDate(opWhen), notAccounted,
      defendants: defendants.map(d => ({ name: d.name, dismissed: d.dismissed, defaultDate: d.defaultDate && fmtShortDate(d.defaultDate) })),
    });

    if (!notAccounted.length) {
      return {
        text: '✓ All ' + N + ' defendant' + plural + ' in default or dismissed (' + defaulted + ' default, ' + dismissed + ' dismissed)',
        color: '#1a6b3a',
        ...packet,
      };
    }
    // Nothing has been defaulted at all: no defendant carries an entered
    // default, superseded or otherwise. Naming every defendant as "not in
    // default/dismissed" adds nothing — the single fact that matters is that
    // the default itself was never entered, so say only that. (A default that
    // predates the operative complaint WAS entered, just superseded, so those
    // cases keep the detailed listing and its caveat below.)
    if (!defendants.some(d => d.defaultDate)) {
      return { text: '⚠ Default Not Entered', color: '#c0392b', ...packet };
    }
    let caveat = '';
    if (!opWhen) caveat = ' — operative complaint not found, cannot verify default dates';
    else if (stale) caveat = ' (' + stale + ' default' + (stale === 1 ? '' : 's') + ' predate the operative complaint of ' + fmtShortDate(opWhen) + ')';
    return {
      text: '⚠ ' + notAccounted.length + ' of ' + N + ' defendant' + plural + ' not in default/dismissed: ' + notAccounted.map(dlEsc).join(', ') + caveat,
      color: '#c0392b',
      ...packet,
    };
  } catch (e) {
    dlLog('OSC status error:', e && e.message || e);
    return { text: 'Default status unavailable', color: '#0a6e6e' };
  }
}

return {
  // Case data sources
  makeCaseCtx, caseCtxForId, caseTabUrlFrom, emptyDoc,
  fetchWithTimeout, fetchCaseDoc, fetchAllDocuments, parseDocRows, absoluteDocUrl,
  parsePartiesTable, parseFutureHearings, parseHearingDateTime,
  // Status
  computeDueDatesFor, computeFiledStatus, computeOscStatus, statusHtml, computeCostsMemoStatus,
  findAppealTimeTrigger, computeFeeMotionStatus,
  // Hearing / document classification
  isOscDefaultJudgment, isWorkableHearing, groupWorkableHearings, hearingIdentityKey, isHearingExcluded, excludedTermMatches,
  isUnlawfulDetainerCase, isUnlawfulDetainerTypeText, findCaseTypeEl, findCaseTypeLine,
  loadExcludedTerms, DEFAULT_EXCLUDED_TERMS,
  isMovingPaper, bestFilingMatch, parseFiledByParties, resolveMovingPaper,
  docWordOverlap, docReferencesMotion, postJudgmentAnchor, hasAccompanyingProofOfService, docLinksToMotion, docNameIsGeneric, docPartyNames, docSharesParty,
  isOppositionDoc, isNonOppositionDoc, isComplaintDoc, isCrossComplaintDoc,
  isFirstAmendedComplaintDoc, isDemurrerOrMotionToStrikeDoc, isDemurrerOrStrikeMotion,
  isPetitionDoc, isPetitionMotion, docsHaveComplaint,
  latestDoc, findDefaultProveUp, isDefaultJudgmentPacketDoc, sameCalendarDay,
  // Document ingest time (when eCourt actually posted a filing)
  decodeIngestTime, getIngestTime, fmtIngest, ingestDay, ingestLagCourtDays,
  courtDaysBetween, withinIngestGrace, INGEST_GRACE_COURT_DAYS, docLagCategory,
  // Text helpers
  stripEventId, stripTrailingParenNumber, stripHearingOnPrefix, stripAncillaryMotionReference,
  movantNormName,
  fmtShortDate, dayMs, dlEsc, dlLog,
  // Deadline engine + palette
  DL, NEXT_DL_GAP, DL_YELLOW, DL_NEUTRAL, DL_BLACK, nextDlColor, motionDlColor,
};
})();

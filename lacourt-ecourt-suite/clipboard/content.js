/**
 * LA Court Clipboard Cleaner - Case Page Content Script (v2.0)
 *
 * Copy/paste on case pages is left completely NATIVE — the extension no longer
 * intercepts the copy event to reformat selections or arm a paste rotation.
 *
 * 1) (removed) Manual selection cleaning / party-rotation copy. Copy is native.
 *
 * 2) Export flow (button / popup):
 *    - Parses the case number, hearing date, motion type, and parties from
 *      the live DOM.
 *    - Builds the field set:
 *        1.  Case number
 *        2.  Hearing date                       — only if a Next Event is shown
 *        3.  Motion type                        — only if a "Hearing on ..." event
 *        4.  First plaintiff (or petitioner)
 *        5.  Other plaintiffs (or petitioners)  — only if 2+
 *        6.  First defendant (or respondent)
 *        7.  Other defendants (or respondents)  — only if 2+
 *        8.  All cross-complainants (combined)  — only if any; FIRST cross-
 *                                                 complaint section only
 *        9.  All cross-defendants (combined)    — only if any; same restriction
 *      Petitioners/Respondents replace Plaintiffs/Defendants when the case
 *      uses those role labels (treated identically in order).
 *      Cross-* fields use a single combined value per side (matching the
 *      user's form layout) rather than separate Title/Other fields. If the
 *      case has multiple cross-complaints, only parties belonging to the
 *      first one are captured.
 *    - Formats the party names (title-case, entity suffixes, short-name
 *      parentheticals, collective labels) and hands them to the Export popup.
 */
(function () {
  'use strict';

  // The case-status engine (lib/case-status.js, loaded first) — the deadline
  // maths, document/party parsing, and background case fetches, shared with the
  // agenda page so both show the same figures from one implementation.
  const {
    makeCaseCtx, emptyDoc, caseTabUrlFrom, fetchCaseDoc, fetchAllDocuments,
    parsePartiesTable, parseFutureHearings, parseHearingDateTime,
    computeDueDatesFor, computeFiledStatus, computeOscStatus, statusHtml,
    isOscDefaultJudgment, isWorkableHearing, groupWorkableHearings, loadExcludedTerms,
    isUnlawfulDetainerCase, findCaseTypeEl,
    isMovingPaper, bestFilingMatch, parseFiledByParties, resolveMovingPaper,
    docWordOverlap, docReferencesMotion, docNameIsGeneric, postJudgmentAnchor, findAppealTimeTrigger,
    docPartyNames, docSharesParty, isComplaintDoc, isCrossComplaintDoc,
    isDemurrerOrMotionToStrikeDoc, isPetitionDoc, latestDoc, findDefaultProveUp,
    sameCalendarDay, stripEventId, stripTrailingParenNumber, stripHearingOnPrefix,
    stripAncillaryMotionReference, movantNormName, fmtShortDate, dlLog,
  } = LACCaseStatus;

  /* ------------------------------------------------------------------ */
  /* Default judgment (OSC Re: Failure to Prosecute / Prove-Up) flow     */
  /* ------------------------------------------------------------------ */
  //
  // When the case's next event is an OSC Re: Failure to Prosecute Default
  // Judgment or a Default Prove-Up Hearing, Export runs the SAME flow as a
  // regular motion — the in-extension
  // Order Template popup, spreadsheet export, and Word mail merge — and, in
  // addition, opens a pre-composed mailto: link addressed to Judge Mackenzie
  // with the case info in the subject and a standard body. The DJ recommendation
  // email is the only DJ-specific extra; the export itself is not special-cased.
  //
  // Ctrl+A rotation and manual selection paste are unaffected.
  // Use DesignPageV2 with topview=Preview so the form loads in the owner's
  // authenticated context (the Auto-Export companion extension needs this
  // session to call the owner-API; the public ResponsePage URL returns 401).
  const REGULAR_FORM_URL = 'https://forms.office.com/Pages/DesignPageV2.aspx?prevorigin=rbf&origin=NeoPortalPage&rpring=UsGovGccProduction&subpage=design&id=x8OU3Ei7_0CTBeRz_W9qFt74YgjxwElOsa89AoRCn9FUQzNGQ0NPWVpUMDBVTzcwN1I2Q0JFOVFZVi4u&analysis=false&tab=0&topview=Preview';

  /**
   * Builds the mailto: URL fired automatically when Fill Microsoft Form is
   * used on an OSC Re: Failure to Prosecute Default Judgment case.
   *
   *   To:      AMackenzie@lacourt.ca.gov
   *   Subject: "MM/DD/YYYY – CASENUM – CASE NAME – OSC RE: FAILURE TO
   *             PROSECUTE DEFAULT JUDGMENT"
   *             (em-dash separators; case name in its original casing)
   *   Body:    Three paragraphs separated by blank lines:
   *              "Judge Mackenzie,"
   *              "The default prove-up packet is complete. I recommend
   *               entering a default judgment. I have sent the judgment
   *               to your queue for your signature."
   *              "Best,\nZach"
   *
   * Returns null if essential pieces are missing (case number or date).
   * The case name is non-essential — we'll send the email without it
   * rather than block the workflow.
   */
  function buildOscMailto(caseNumber, hearingDate, caseName, hearingType) {
    if (!caseNumber || !hearingDate) return null;

    const EM = '\u2013'; // en-dash — what the user calls "em dash" colloquially
    const subjectParts = [hearingDate, caseNumber];
    if (caseName) subjectParts.push(caseName);
    // Describe the actual hearing in the subject: a Default Prove-Up Hearing vs.
    // the OSC. Both run the same recommendation-email flow.
    subjectParts.push(/prove\s*-?\s*up/i.test(hearingType || '')
      ? 'DEFAULT PROVE-UP HEARING'
      : 'OSC RE: FAILURE TO PROSECUTE DEFAULT JUDGMENT');
    const subject = subjectParts.join(' ' + EM + ' ');

    // CRLF line breaks so Outlook on Windows handles the mailto body
    // exactly as written. mailto: spec is technically %0A only, but %0D%0A
    // is widely accepted and renders correctly in Outlook.
    const body =
      'Judge Mackenzie,\r\n' +
      '\r\n' +
      'The default prove-up packet is complete. I recommend entering a default judgment. I have sent the judgment to your queue for your signature.\r\n' +
      '\r\n' +
      'Best,\r\n' +
      'Zach';

    return 'mailto:AMackenzie@lacourt.ca.gov' +
      '?subject=' + encodeURIComponent(subject) +
      '&body=' + encodeURIComponent(body);
  }



  /* ------------------------------------------------------------------ */
  /* Dismissed-party motion exclusion list                              */
  /* ------------------------------------------------------------------ */
  //
  // When the parsed motion type matches any term below (case-insensitive
  // substring match), dismissed parties are dropped from the rotation /
  // Fill Microsoft Form output. The intuition: dispositive and pleading-
  // stage motions only affect parties still actively litigating the
  // merits, so a dismissed party shouldn't appear in the caption. But
  // post-judgment cleanup motions (attorney fees, costs, sanctions) can
  // still implicate a dismissed party, so for those we keep them.
  //
  // This default list is also stored in options.js (where the user can
  // edit it). Both lists must stay in sync — see options.js DEFAULT_DISMISSED_MOTION_EXCLUSIONS.
  //
  // The list applies to:
  //   - Ctrl+A rotation copy
  //   - Fill Microsoft Form button (popup + floating)
  //
  // It does NOT apply to:
  //   - Manual subset selection paste (the user is curating by hand)
  //   - "Removed - No Longer Named" / "No Longer Named" parties (those
  //     are ALWAYS dropped, regardless of motion type)
  const DEFAULT_DISMISSED_MOTION_EXCLUSIONS = [
    // Dispositive / merits motions
    'summary judgment',
    'summary adjudication',
    'judgment on the pleadings',
    'directed verdict',
    'nonsuit',
    'new trial',
    'vacate judgment',
    'set aside default',
    // Pleading challenges
    'demurrer',
    'motion to strike',
    'anti-slapp',
    'special motion to strike',
    'leave to amend',
    'leave to file cross-complaint',
    // Discovery
    'motion to compel',
    'protective order',
    'motion to quash',
    'trial preference',
    'motion in limine',
    'bifurcate',
    'consolidate',
    'sever',
    'coordinate',
    // Class / representative actions
    'class certification',
    'decertify',
    // Equitable relief
    'preliminary injunction',
    'temporary restraining order',
    'writ of attachment',
    // Service / jurisdiction
    'quash service',
    'order to show cause re contempt',
  ];

  let dismissedMotionExclusions = DEFAULT_DISMISSED_MOTION_EXCLUSIONS.slice();

  // Load the user-edited list from chrome.storage.sync. Falls back to the
  // embedded defaults if the user hasn't visited the options page yet.
  try {
    chrome.storage.sync.get(['dismissedMotionExclusions'], result => {
      if (chrome.runtime.lastError) return;
      const list = result && result.dismissedMotionExclusions;
      if (Array.isArray(list)) {
        dismissedMotionExclusions = list.map(s => String(s || '').trim().toLowerCase()).filter(Boolean);
      }
    });
  } catch (_) {}

  // Refresh whenever the user saves new edits in the options page.
  try {
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area !== 'sync') return;
      if (!changes.dismissedMotionExclusions) return;
      const list = changes.dismissedMotionExclusions.newValue;
      if (Array.isArray(list)) {
        dismissedMotionExclusions = list.map(s => String(s || '').trim().toLowerCase()).filter(Boolean);
      }
    });
  } catch (_) {}

  /**
   * Returns true if the given motion type should cause dismissed parties to
   * be dropped from the rotation output. Substring-matches against the
   * user-editable exclusion list (case-insensitive). Empty/missing motion
   * type → false (no exclusion, keep dismissed parties).
   */
  function motionExcludesDismissed(motionType) {
    if (!motionType) return false;
    // Read the relief THIS motion seeks, not the motion its title says it
    // supports: a motion to seal exhibits "in support of Defendant's Motion for
    // Summary Judgment" is not itself a summary judgment motion.
    const m = stripAncillaryMotionReference(motionType).toLowerCase();
    return dismissedMotionExclusions.some(term => term && m.includes(term));
  }

/**
 * Parses the case page DOM and builds the rotation sequence + labeled object.
 * Returns { sequence, labeled } or null if there's nothing to capture.
 *
 * Does NOT touch the clipboard or storage — pure data extraction, used by the
 * Export flow.
 */
function buildRotationData(root, hearingOverride) {
  root = root || document;
  const parties = parsePartiesTable(root);
  const caseNumber = parseCaseNumber(root);
  // When Export resolved a different hearing (because the Next event was
  // excluded), use its date/type; otherwise parse the Next event live.
  const hearingDate = hearingOverride ? hearingOverride.hearingDate : parseHearingDate(root);
  const motionType = hearingOverride ? hearingOverride.motionType : parseMotionType(root);

  // If the motion type indicates a proceeding that a dismissed party would
  // not be involved in (e.g. summary judgment, demurrer), drop dismissed
  // parties before classifying. The selection-paste mode does NOT call this
  // function, so manual subset selections are unaffected.
  const dropDismissed = motionExcludesDismissed(motionType);
  const eligibleParties = dropDismissed
    ? parties.filter(p => !p.dismissed)
    : parties;

  if (dropDismissed) {
    const dropped = parties.filter(p => p.dismissed).map(p => `${p.role}: ${p.name}`);
    console.log('[LACourt] motion type "' + motionType +
      '" excludes dismissed parties; dropped:', dropped);
  }

  // Classify each party into one of four buckets. Order of checks matters:
  // "Cross-Defendant" contains "Defendant", so cross-* must be tested first.
  const primaryClaimants  = []; // Plaintiff OR Petitioner
  const primaryRespondents = []; // Defendant OR Respondent
  const crossClaimants    = []; // Cross-Complainant
  const crossRespondents  = []; // Cross-Defendant

  for (const p of eligibleParties) {
    const role = p.role || '';
    if (/cross[-\s]?complainant/i.test(role)) {
      crossClaimants.push(p);
    } else if (/cross[-\s]?defendant/i.test(role)) {
      crossRespondents.push(p);
    } else if (/^\s*(plaintiff|petitioner)\b/i.test(role)) {
      primaryClaimants.push(p);
    } else if (/^\s*(defendant|respondent)\b/i.test(role)) {
      primaryRespondents.push(p);
    }
  }

  console.log('[LACourt] parsed:', {
    caseNumber, hearingDate, motionType,
    primaryClaimants: primaryClaimants.map(p => `${p.role}: ${p.name}`),
    primaryRespondents: primaryRespondents.map(p => `${p.role}: ${p.name}`),
    crossClaimants: crossClaimants.map(p => `${p.role}: ${p.name}`),
    crossRespondents: crossRespondents.map(p => `${p.role}: ${p.name}`),
    rawPartiesParsed: parties.length,
  });

  const sequence = [];
  const labeled = {};

  if (caseNumber)   { sequence.push(caseNumber);   labeled.caseNumber  = caseNumber; }
  if (hearingDate)  { sequence.push(hearingDate);  labeled.hearingDate = hearingDate; }
  if (motionType)   { sequence.push(motionType);   labeled.motionType  = motionType; }

  // Primary side. Petitioners share form labels with plaintiffs.
  // Resolve short names for all plaintiffs on this side at once so
  // colliding parties (e.g. two "Globex" entities) get widened.
  const plaintiffShortNames = resolveShortNames(primaryClaimants.map(p => p.name));
  let titlePlaintiffShortName = null;
  if (primaryClaimants.length >= 1) {
    const result = formatPartyName(primaryClaimants[0].name);
    sequence.push(result.formatted);
    labeled.titlePlaintiff = result.formatted;
    titlePlaintiffShortName = plaintiffShortNames.get(primaryClaimants[0].name) || null;
  }
  if (primaryClaimants.length >= 2) {
    const others = formatOthers(
      primaryClaimants.slice(1).map(p => p.name),
      'Plaintiffs',
      titlePlaintiffShortName,
      plaintiffShortNames
    );
    sequence.push(others);
    labeled.otherPlaintiffs = others;
  }

  const defendantShortNames = resolveShortNames(primaryRespondents.map(p => p.name));
  let titleDefendantShortName = null;
  if (primaryRespondents.length >= 1) {
    const result = formatPartyName(primaryRespondents[0].name);
    sequence.push(result.formatted);
    labeled.titleDefendant = result.formatted;
    titleDefendantShortName = defendantShortNames.get(primaryRespondents[0].name) || null;
  }
  if (primaryRespondents.length >= 2) {
    const others = formatOthers(
      primaryRespondents.slice(1).map(p => p.name),
      'Defendants',
      titleDefendantShortName,
      defendantShortNames
    );
    sequence.push(others);
    labeled.otherDefendants = others;
  }

  // Cross side. The user's Microsoft Form has a single field per side
  // (CrossComplainants / CrossDefendants) rather than separate Title /
  // Other fields, so we build one combined value containing every cross-*
  // party on that side. The collective "(collectively …)" suffix only
  // appears when there are 2+ parties.
  //
  // Notes:
  // - parsePartiesTable() restricts cross-* parties to those belonging to
  //   the FIRST cross-complaint section in the eCourt parties table — any
  //   2nd, 3rd, etc. cross-complaint's parties are dropped on the floor.
  // - If there are no eligible cross-defendants left (because all of them
  //   were dismissed or marked No Longer Named, or none existed), we also
  //   suppress the cross-complainants output. Without any cross-defendants
  //   in play there's no live cross-action for the cross-complainants to
  //   appear in, so emitting them in the order would be misleading.
  const haveCrossRespondents = crossRespondents.length >= 1;

  if (crossClaimants.length >= 1 && haveCrossRespondents) {
    const map = resolveShortNames(crossClaimants.map(p => p.name));
    const combined = formatCombinedList(
      crossClaimants.map(p => p.name),
      'Cross-Complainants',
      map
    );
    sequence.push(combined);
    labeled.crossComplainants = combined;
  } else if (crossClaimants.length >= 1 && !haveCrossRespondents) {
    console.log('[LACourt] suppressing cross-complainants — no eligible cross-defendants remain');
  }

  if (haveCrossRespondents) {
    const map = resolveShortNames(crossRespondents.map(p => p.name));
    const combined = formatCombinedList(
      crossRespondents.map(p => p.name),
      'Cross-Defendants',
      map
    );
    sequence.push(combined);
    labeled.crossDefendants = combined;
  }

  // Every non-party name (attorneys, firms, "Non-Party" entities) for the
  // pseudonym generator, minus anyone already in the caption. Goes in `labeled`
  // only (never the Ctrl+A rotation sequence) so it lands in a trailing
  // spreadsheet column that the mail merge ignores.
  const captionNames = new Set(
    [...primaryClaimants, ...primaryRespondents, ...crossClaimants, ...crossRespondents]
      .map(p => movantNormName(p.name))
  );
  const otherNames = parseNonPartyNames(root).filter(n => !captionNames.has(movantNormName(n)));
  // Dismissed parties dropped by the motion-type exclusion still belong in the
  // pseudonym pool — add them to Other Names (deduped) since they no longer
  // appear in any party field.
  if (dropDismissed) {
    const seen = new Set(otherNames.map(n => movantNormName(n)));
    for (const p of parties) {
      if (!p.dismissed || !p.name) continue;
      const key = movantNormName(p.name);
      if (captionNames.has(key) || seen.has(key)) continue;
      seen.add(key);
      otherNames.push(p.name);
    }
  }
  if (otherNames.length) labeled.otherNames = otherNames.join('; ');

  console.log('[LACourt] rotation sequence:', sequence);
  console.log('[LACourt] labeled:', labeled);

  if (sequence.length === 0) return null;
  return { sequence, labeled };
}

/**
 * Stores the rotation via the service worker, with a direct-storage fallback.
 */
function storeRotation(data, extra) {
  const payload = { type: 'setRotation', sequence: data.sequence, labeled: data.labeled };
  if (extra) Object.assign(payload, extra);
  try {
    chrome.runtime.sendMessage(payload, () => {
      if (chrome.runtime.lastError) {
        try {
          chrome.storage.local.set({
            lacourtRotation: { ...data, index: 0, createdAt: Date.now() },
          });
        } catch (_) {}
      }
    });
  } catch (_) {
    try {
      chrome.storage.local.set({
        lacourtRotation: { ...data, index: 0, createdAt: Date.now() },
      });
    } catch (_) {}
  }
}

/**
 * Stores the parsed field values for the Order Template popup to read.
 *
 * The popup (order-template/order-template.html) opens in its own window and
 * reads this key on load to pre-fill its editable boxes. We store the labeled
 * object verbatim; the popup maps each key to its form question / export
 * column. `movant` is intentionally NOT included — the user fills it in the
 * popup by hand.
 *
 * Returns a Promise so callers can wait for the write to land before opening
 * the popup window (avoids a load-vs-write race).
 */
function storeOrderTemplateData(labeled) {
  return new Promise(resolve => {
    try {
      chrome.storage.local.set(
        { orderTemplateData: { fields: labeled || {}, createdAt: Date.now() } },
        () => { void chrome.runtime.lastError; resolve(); }
      );
    } catch (_) {
      resolve();
    }
  });
}

/**
 * Builds the context object both Fill-Microsoft-Form entry points need.
 *
 * Detects whether the case is an OSC Re: Failure to Prosecute Default
 * Judgment and, if so:
 *   - Selects the OSC form URL instead of the regular one.
 *   - Trims the rotation data down to only Case Number + Hearing Date
 *     (the OSC form has only those two fields).
 *   - Builds a mailto: URL pre-addressed to Judge Mackenzie.
 *
 * For non-OSC (regular) cases the Order Template Input Microsoft Form has been
 * retired in favor of an in-extension popup: `openUrl` points at the packaged
 * order-template page and `isOrderTemplate` is true. OSC / Default Judgment
 * Checklist cases are unchanged — they still open the real Microsoft Form.
 *
 * Returns null if there's no rotation data at all.
 * Returns { data, formUrl, openUrl, isOrderTemplate, mailtoUrl, isOsc,
 *           hearingType } otherwise.
 *
 * mailtoUrl is null for non-OSC cases.
 */
function getFillFormContext(root, hearingOverride) {
  root = root || document;
  const data = buildRotationData(root, hearingOverride);
  if (!data) return null;

  const hearingType = hearingOverride ? hearingOverride.hearingType : parseHearingType(root);
  const isOsc = isOscDefaultJudgment(hearingType);

  // An OSC Re: Failure to Prosecute Default Judgment runs the SAME export as a
  // regular motion — the in-extension Order Template popup, spreadsheet export,
  // and Word mail merge — but ALSO fires the recommendation email to Judge
  // Mackenzie. So the only DJ-specific extra is the mailto; everything else is
  // the ordinary order-template flow (full data, not the trimmed OSC form).
  let mailtoUrl = null;
  if (isOsc) {
    // parseCaseName accepts an optional case-number hint so it can pin the
    // location precisely.
    const caseName = parseCaseName(data.labeled.caseNumber, root);
    mailtoUrl = buildOscMailto(
      data.labeled.caseNumber,
      data.labeled.hearingDate,
      caseName,
      hearingType
    );
    console.log('[LACourt] OSC default-judgment flow:', {
      hearingType, caseName, mailtoUrl: !!mailtoUrl,
    });
  }

  return {
    data,
    formUrl: REGULAR_FORM_URL,
    openUrl: chrome.runtime.getURL('order-template/order-template.html'),
    isOrderTemplate: true,
    mailtoUrl,
    isOsc,
    hearingType,
  };
}

/**
 * Fires a mailto: URL from the case page, handing off to the OS default
 * mail handler (Outlook on Windows). We use a hidden anchor with
 * target="_blank" + click() rather than window.location or window.open so
 * the case page itself doesn't navigate and no blank tab is left behind.
 *
 * Returns true if the click was dispatched; the actual OS handoff is
 * fire-and-forget (we have no signal back from the mail handler).
 */
function triggerMailto(mailtoUrl) {
  if (!mailtoUrl) return false;
  try {
    const a = document.createElement('a');
    a.href = mailtoUrl;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    // Defer removal so the click event finishes propagating.
    setTimeout(() => { try { a.remove(); } catch (_) {} }, 100);
    return true;
  } catch (err) {
    console.error('[LACourt] triggerMailto failed:', err);
    return false;
  }
}

/**
 * Listens for messages from the popup. The popup invokes this to capture the
 * current case-page data, store the rotation with autoFillOnLoad=true, and
 * reply so the popup can open the form URL.
 *
 * On OSC Re: Failure to Prosecute Default Judgment cases, the response
 * also carries the OSC form URL and a mailto URL — the popup will then
 * fire the mailto from THIS frame's context (via openMailto message
 * routed back here) so the case page handles the OS handoff.
 */
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg && msg.type === 'captureForFormFill') {
    // getExportContext is async (it may background-fetch the Parties page when
    // Export is pressed from a non-Parties case page), so we always keep the
    // message channel open and reply from the promise.
    getExportContext().then(result => {
      if (!result) {
        sendResponse({ ok: false, reason: 'no-data' });
        return;
      }
      const ctx = result.ctx;
      storeRotation(ctx.data, { autoFillOnLoad: true });

      const reply = {
        ok: true,
        count: ctx.data.sequence.length,
        formUrl: ctx.formUrl,
        openUrl: ctx.openUrl,
        isOrderTemplate: ctx.isOrderTemplate,
        mailtoUrl: ctx.mailtoUrl,
        isOsc: ctx.isOsc,
      };

      if (ctx.isOrderTemplate) {
        // Auto-detect the Movant (background-fetches the Documents page; roster
        // read from the same parties root), then stash the fields and reply
        // only once the write has landed so the popup can't load first.
        computeMovant(ctx.data.labeled.motionType, result.partiesRoot, ctx.data.labeled.hearingDate).then(movant => {
          if (movant) ctx.data.labeled.movant = movant;
          storeOrderTemplateData(ctx.data.labeled).then(() => sendResponse(reply));
        });
      } else {
        sendResponse(reply);
      }
    });
    return true; // async response
  }

  if (msg && msg.type === 'fireMailto' && typeof msg.url === 'string') {
    triggerMailto(msg.url);
    sendResponse({ ok: true });
    return false;
  }

  return false;
});

/**
 * Smart name formatter that handles business entities and proper casing.
 * Returns an object with:
 *   - formatted: The full formatted name (entity suffix + any location
 *     phrase). NEVER includes the short-name parenthetical — the
 *     parenthetical is the responsibility of formatOthers(), which appends
 *     "(Short Name)" to non-title entities and prepends "(TitleShort)" when
 *     the title party is an entity and there are 2+ parties on that side.
 *   - shortName: The short name (if business entity), null otherwise.
 *
 * Business entity handling:
 * - Detects entity types (inc., LLC, Corp., etc.) at end of name
 * - Extracts entity name (up to 3 words before entity type)
 * - Strips trailing punctuation (e.g. comma in "Monsters, Inc.") so the
 *   short-name parenthetical doesn't end up as "(Monsters,)"
 *
 * Case handling:
 * - Preserves mixed case (e.g., "McDonald's", "eBay")
 * - Converts all-caps or all-lowercase to title case
 * - Preserves known acronyms (LLC, Inc., USA, FBI, etc.)
 * - Preserves 2-letter combos with & (e.g., "A&E", "H&R")
 * - Preserves 3-letter all-caps that look like acronyms
 */
function formatPartyName(name, wordCount) {
  if (!name || !name.trim()) return { formatted: '', shortName: null };

  const original = name.trim();

  // Step 1: Strip a trailing location phrase like "a Delaware Corporation" /
  // "A NEW YORK COMPANY" / "TEXAS". The location phrase pattern intentionally
  // does NOT match bare "Inc"/"Incorporated" because those are entity
  // suffixes that come BEFORE the location phrase, not part of it.
  const STATE_COUNTRY = '(?:alabama|alaska|arizona|arkansas|california|colorado|connecticut|delaware|florida|georgia|hawaii|idaho|illinois|indiana|iowa|kansas|kentucky|louisiana|maine|maryland|massachusetts|michigan|minnesota|mississippi|missouri|montana|nebraska|nevada|new\\s+hampshire|new\\s+jersey|new\\s+mexico|new\\s+york|north\\s+carolina|north\\s+dakota|ohio|oklahoma|oregon|pennsylvania|rhode\\s+island|south\\s+carolina|south\\s+dakota|tennessee|texas|utah|vermont|virginia|washington|west\\s+virginia|wisconsin|wyoming|canada|mexico|united\\s+kingdom|uk|france|germany|japan|china|india)';
  // Word-form trailing entity in a location phrase: "Corporation" / "Company"
  // / "Corp" / "Co" / "Limited Liability Company". We deliberately do NOT
  // include "Inc"/"Incorporated" here; "INC A DELAWARE CORPORATION" should
  // be parsed as INC + (a Delaware Corporation), not as the location phrase
  // alone.
  const LOC_TRAILING = '(?:limited\\s+liability\\s+company|corporation|company|corp\\.?|co\\.?)';
  const locationRe = new RegExp(`\\b(?:a\\s+)?${STATE_COUNTRY}(?:\\s+${LOC_TRAILING})?\\s*$`, 'i');

  let workingName = original;
  let locationPhrase = '';
  const locMatch = workingName.match(locationRe);
  if (locMatch) {
    locationPhrase = locMatch[0];
    workingName = workingName.substring(0, locMatch.index).trim();
    // Strip any trailing comma left behind, e.g. "Acme Inc., a Delaware Corp."
    workingName = workingName.replace(/[,;:]+$/, '').trim();
  }

  // Step 2: Detect an entity suffix at the end of what remains. For each
  // group, the *render* is determined by what the user actually wrote:
  //   - Word-form full suffix → preserve as-written, just title-case it
  //     (Incorporated, Corporation, Company, Limited, Association)
  //   - Abbreviated suffix → preserve the user's period choice (Inc / Inc.,
  //     Corp / Corp., Co / Co., Ltd / Ltd., Assn / Assn.)
  //   - Acronym suffix → emit canonical all-caps (LLC, LLP, LP, PLLC, LLLP,
  //     PLC) or canonical mixed-case (GmbH, S.A.).
  // Order within the array matters only for ambiguous cases ("LP" must come
  // after "LLP" / "LLLP" / "PLLC" since those contain it; we handle that by
  // requiring word boundaries and by using the longest match first).
  const SUFFIX_GROUPS = [
    // Acronyms — canonical all-caps regardless of how user wrote them.
    { re: /\b(l\.l\.l\.p\.?|lllp)\s*$/i,  render: () => 'LLLP' },
    { re: /\b(p\.l\.l\.c\.?|pllc)\s*$/i,  render: () => 'PLLC' },
    { re: /\b(l\.l\.c\.?|llc)\s*$/i,       render: () => 'LLC' },
    { re: /\b(l\.l\.p\.?|llp)\s*$/i,       render: () => 'LLP' },
    { re: /\b(l\.p\.?|lp)\s*$/i,           render: () => 'LP' },
    { re: /\bplc\s*$/i,                    render: () => 'PLC' },
    { re: /\b(p\.c\.?|pc)\s*$/i,           render: () => 'PC' },
    { re: /\bgmbh\s*$/i,                   render: () => 'GmbH' },
    { re: /\b(s\.a\.|sa)\s*$/i,            render: () => 'S.A.' },

    // Word-form full suffixes — preserve as a real word, casing-normalized.
    { re: /\bincorporated\s*$/i,  render: m => titleCaseWord(m[0].trim()) }, // "Incorporated"
    { re: /\bcorporation\s*$/i,   render: m => titleCaseWord(m[0].trim()) }, // "Corporation"
    { re: /\bcompany\s*$/i,       render: m => titleCaseWord(m[0].trim()) }, // "Company"
    { re: /\blimited\s*$/i,       render: m => titleCaseWord(m[0].trim()) }, // "Limited"
    { re: /\bassociation\s*$/i,   render: m => titleCaseWord(m[0].trim()) }, // "Association"

    // Abbreviated word-form suffixes — preserve user's period choice.
    { re: /\binc(\.?)\s*$/i,   render: m => 'Inc'  + (m[1] || '') },
    { re: /\bcorp(\.?)\s*$/i,  render: m => 'Corp' + (m[1] || '') },
    { re: /\bco(\.?)\s*$/i,    render: m => 'Co'   + (m[1] || '') },
    { re: /\bltd(\.?)\s*$/i,   render: m => 'Ltd'  + (m[1] || '') },
    { re: /\bassn(\.?)\s*$/i,  render: m => 'Assn' + (m[1] || '') },
  ];

  let entityRendered = null;
  let coreName = workingName;
  for (const { re, render } of SUFFIX_GROUPS) {
    const m = workingName.match(re);
    if (m) {
      entityRendered = render(m);
      coreName = workingName.substring(0, m.index).trim();
      // NOTE: do NOT strip a trailing comma from coreName here. It is part
      // of the display ("Monsters, Inc."). The short-name extraction strips
      // its own trailing punctuation downstream.
      break;
    }
  }

  if (entityRendered === null && !locationPhrase) {
    // Not a recognized business entity — just smart-case the original.
    return { formatted: smartCase(original), shortName: null };
  }

  // If we matched a location phrase but no entity suffix, treat the original
  // as not-an-entity (we don't want to mangle "City of Los Angeles" type
  // names). Defensive — locationRe requires a state/country word so this is
  // unlikely to fire on personal names.
  if (entityRendered === null) {
    return { formatted: smartCase(original), shortName: null };
  }

  return formatBusinessEntityWithShortName(coreName, entityRendered, locationPhrase, wordCount);
}

/**
 * Title-case a single word, preserving any internal punctuation. Used for
 * full-word entity suffixes like "Incorporated" / "Corporation".
 */
function titleCaseWord(word) {
  if (!word) return word;
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

/**
 * Connective stop-words that should never be the LAST word of a short name.
 * If a short-name slice ends on one of these (e.g. "Bank of" from "Bank of
 * America Corp"), one more word is pulled in.
 */
const SHORT_NAME_STOP_WORDS = new Set([
  'of', 'and', 'the', 'for', '&', 'a', 'an', 'in', 'on', 'at', 'to', 'by',
]);

/**
 * Take the first `n` words of `words`, but never end on a stop-word — extend
 * by additional words as needed (still bounded by the array length). This is
 * shared between formatBusinessEntityWithShortName (default 2-word slice)
 * and resolveShortNames (variable widths during disambiguation).
 */
function sliceShortName(words, n) {
  let take = Math.min(n, words.length);
  while (take < words.length && SHORT_NAME_STOP_WORDS.has(words[take - 1].toLowerCase())) {
    take += 1;
  }
  return words.slice(0, take);
}

/**
 * Formats a business entity and extracts the short name.
 *
 * Returns: { formatted, shortName }
 *   - formatted: the full display name including entity suffix and any
 *     "a Delaware Corporation"-style location phrase. Does NOT include
 *     the short-name parenthetical — that is appended only when this
 *     party is referenced as the title party from formatOthers().
 *   - shortName: the FIRST `wordCount` words of the core name (default 2),
 *     extended by one if the slice would end on a stop-word (so "Bank of
 *     Commerce Corp" yields "Bank of Commerce" not "Bank of"). Trailing
 *     punctuation on the core (e.g. "Monsters," from "Monsters, Inc.") is
 *     stripped before extraction. Callers may pass a larger `wordCount`
 *     to disambiguate parties whose first 2 words collide.
 *
 * Examples (default 2-word width):
 *   "ACME CORPORATION LLC"
 *     → formatted: "Acme Corporation LLC"
 *       shortName: "Acme Corporation"
 *   "MONSTERS, INC."
 *     → formatted: "Monsters, Inc."
 *       shortName: "Monsters"          (only 1 word in core)
 *   "GLOBEX STORES INC A DELAWARE CORPORATION"
 *     → formatted: "Globex Stores Inc a Delaware Corporation"
 *       shortName: "Globex Stores"
 *   "BANK OF COMMERCE CORP"
 *     → formatted: "Bank of Commerce Corp"
 *       shortName: "Bank of Commerce"   (extended past stop-word "of")
 */
function formatBusinessEntityWithShortName(coreName, entityRendered, locationPhrase, wordCount) {
  // Strip trailing punctuation (commas, semicolons, etc.) from the core name
  // before deriving the short name. e.g. "MONSTERS, INC." leaves the core as
  // "MONSTERS," after the entity strip — without this trim the short-name
  // parenthetical would render as "(Monsters,)".
  const coreForShortName = coreName.replace(/[\s,;:.\-]+$/, '');

  // Extract SHORT name. Default is the FIRST `wordCount` words of the core
  // name (typically 2). buildRotationData() may pass a larger value to
  // disambiguate parties on the same side that would otherwise collide.
  // If the slice would end on a stop-word like "of"/"and"/"the"/"for"/"&"
  // (e.g. "Bank of Commerce" → "Bank of"), extend by one more word so the
  // short name doesn't dangle on a connective.
  const n = wordCount || 2;
  const words = coreForShortName.split(/\s+/).filter(Boolean);
  const shortName = sliceShortName(words, n).join(' ');
  const formattedShortName = smartCase(shortName);

  // Format the display name. Note: the entity suffix (entityRendered) is
  // already in its final form — do NOT pass it through smartCase, which
  // would mangle "Inc" → "Inc." or "CORP" → "CORP".
  let displayName = smartCase(coreName) + ' ' + entityRendered;
  if (locationPhrase) {
    let renderedLocation = smartCase(locationPhrase);
    // Location phrases like "a Delaware Corporation" are sentence fragments
    // that follow the entity suffix — the leading article "a" should always
    // be lowercase, but smartCase capitalizes first words by rule. Override
    // here.
    renderedLocation = renderedLocation.replace(/^A\s+/, 'a ');
    displayName += ' ' + renderedLocation;
  }

  return { formatted: displayName, shortName: formattedShortName };
}

/**
 * Smart case conversion with legal-specific acronym and abbreviation detection.
 * Based on the Legal Citation Linker capitalization rules.
 * 
 * Heuristics for legal document formatting:
 * 1. Protected acronyms (always all-caps): FAC, SAC, TAC, CEQA, LLC, LLP, etc.
 * 2. Protected abbreviations (canonical mixed-case): Inc., MtS
 * 3. Statute code abbreviations: Civ., Pen., Evid., etc.
 * 4. Multi-tier heuristic detection for unlisted acronyms
 * 5. Common lowercase words (except first/last position)
 * 6. Mixed case preservation
 */
function smartCase(text) {
  if (!text) return '';
  
  // If text has mixed case (not all upper or all lower), preserve it
  const hasUpper = /[A-Z]/.test(text);
  const hasLower = /[a-z]/.test(text);
  const isMixedCase = hasUpper && hasLower;
  
  if (isMixedCase) {
    return text; // Already mixed case, preserve it
  }
  
  // Build casing map from original input
  const rawCasingMap = new Map();
  text.split(/\s+/).forEach(word => {
    const key = word.toLowerCase().replace(/[^a-z]/g, '');
    if (key && !rawCasingMap.has(key)) {
      rawCasingMap.set(key, { original: word, hadPeriod: word.endsWith('.') });
    }
  });
  
  // Protected acronyms - always all-caps (true acronyms/initialisms)
  const PROTECTED_ACRONYMS = new Map([
    ['fac', 'FAC'], ['sac', 'SAC'], ['tac', 'TAC'], ['ceqa', 'CEQA'],
    ['cd', 'CD'], ['ceo', 'CEO'], ['iied', 'IIED'],
    ['llc', 'LLC'], ['llp', 'LLP'], ['lp', 'LP'], ['lllp', 'LLLP'],
    ['pc', 'PC'], ['gp', 'GP'], ['feha', 'FEHA'], ['iso', 'ISO'], ['msj', 'MSJ'],
    ['dba', 'DBA'], ['aka', 'AKA'], ['fka', 'FKA'], ['nka', 'NKA'],
    // Common business/government acronyms
    ['usa', 'USA'], ['us', 'US'], ['uk', 'UK'], ['eu', 'EU'], ['un', 'UN'],
    ['fbi', 'FBI'], ['cia', 'CIA'], ['nsa', 'NSA'], ['dea', 'DEA'], ['atf', 'ATF'],
    ['irs', 'IRS'], ['dmv', 'DMV'], ['dot', 'DOT'], ['epa', 'EPA'], ['fda', 'FDA'],
    ['cfo', 'CFO'], ['cto', 'CTO'], ['coo', 'COO'], ['cpa', 'CPA'], ['cfa', 'CFA'],
    ['mba', 'MBA'], ['phd', 'PhD'], ['md', 'MD'], ['rn', 'RN'],
    ['it', 'IT'], ['hr', 'HR'], ['pr', 'PR'], ['rv', 'RV'], ['tv', 'TV'],
    ['ibm', 'IBM'], ['hp', 'HP'], ['gm', 'GM'], ['ge', 'GE'],
  ]);
  
  // Protected abbreviations - canonical mixed-case form
  const PROTECTED_ABBREVIATIONS = new Map([
    ['inc', 'Inc.'],
    ['mts', 'MtS'],
  ]);
  
  // Statute code dotted forms - require trailing period in original
  const STATUTE_CODE_DOTTED = new Map([
    ['civ', 'Civ.'], ['pen', 'Pen.'], ['evid', 'Evid.'], ['bus', 'Bus.'],
    ['prof', 'Prof.'], ['fam', 'Fam.'], ['gov', 'Gov.'], ['govt', 'Govt.'],
    ['saf', 'Saf.'], ['lab', 'Lab.'], ['prob', 'Prob.'], ['veh', 'Veh.'],
    ['welf', 'Welf.'], ['inst', 'Inst.'], ['corp', 'Corp.'], ['ins', 'Ins.'],
    ['rev', 'Rev.'], ['tax', 'Tax.'], ['educ', 'Educ.'], ['elec', 'Elec.'],
    ['fin', 'Fin.'], ['agric', 'Agric.'], ['agr', 'Agr.'], ['harb', 'Harb.'],
    ['nav', 'Nav.'], ['mil', 'Mil.'], ['vet', 'Vet.'], ['cont', 'Cont.'],
    ['contract', 'Contract.'], ['res', 'Res.'], ['util', 'Util.'], ['sts', 'Sts.'],
    ['hy', 'Hy.'], ['unemp', 'Unemp.'], ['wat', 'Wat.'], ['com', 'Com.'],
    ['proc', 'Proc.'],
  ]);
  
  // Statute code acronyms - standalone (no period)
  const STATUTE_CODE_ACRONYMS = new Set([
    'bpc', 'com', 'civ', 'ccp', 'corp', 'edc', 'elec', 'evid', 'fam', 'fin',
    'fgc', 'fac', 'gov', 'hnc', 'hsc', 'ins', 'lab', 'mvc', 'pen', 'prob',
    'pcc', 'prc', 'puc', 'rtc', 'shc', 'uic', 'veh', 'wat', 'wic'
  ]);
  
  // Common lowercase words (except first/last position)
  const LOWERCASE_WORDS = new Set([
    'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in', 'of', 'on',
    'or', 'the', 'to', 'with', 'from', 'into', 'onto', 'upon', 'over', 'under',
    'via', 'per', 'up', 'down', 'out', 'off', 'through'
  ]);
  
  const words = text.split(/\s+/);
  
  return words.map((word, index) => {
    if (!word) return word;
    
    // Strip leading/trailing punctuation for analysis
    const match = word.match(/^([^a-zA-Z]*)([a-zA-Z]+)([^a-zA-Z]*)$/);
    if (!match) return word; // No letters, return as-is
    
    const [, lead, bare, trail] = match;
    const lowerBare = bare.toLowerCase();
    const isFirst = index === 0;
    const isLast = index === words.length - 1;
    
    // Check if original had trailing period
    const hadPeriod = rawCasingMap.get(lowerBare)?.hadPeriod || false;
    
    // Rule 2.3: Protected acronyms - always canonical
    if (PROTECTED_ACRONYMS.has(lowerBare)) {
      return lead + PROTECTED_ACRONYMS.get(lowerBare) + trail;
    }
    
    // Rule 2.4: Protected abbreviations - canonical form
    if (PROTECTED_ABBREVIATIONS.has(lowerBare)) {
      const canonical = PROTECTED_ABBREVIATIONS.get(lowerBare);
      // Remove first period from trail to avoid double period
      const cleanTrail = trail.replace(/^\./, '');
      return lead + canonical + cleanTrail;
    }
    
    // Rule 2.5: Statute code dotted forms - only if had period in original
    if (hadPeriod && STATUTE_CODE_DOTTED.has(lowerBare)) {
      const canonical = STATUTE_CODE_DOTTED.get(lowerBare);
      const cleanTrail = trail.replace(/^\./, '');
      return lead + canonical + cleanTrail;
    }
    
    // Rule 2.6: Heuristic acronym detection
    if (looksLikeAcronym(bare, isMixedCase, rawCasingMap)) {
      const original = rawCasingMap.get(lowerBare)?.original || bare.toUpperCase();
      return lead + original + trail;
    }
    
    // Common lowercase words (except first/last position)
    if (!isFirst && !isLast && LOWERCASE_WORDS.has(lowerBare)) {
      return word.toLowerCase();
    }
    
    // Default: title case (capitalize first letter)
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(' ');
}

/**
 * Heuristic acronym detector with four tiers.
 * Returns true if the word should be treated as an acronym.
 */
function looksLikeAcronym(bare, inputIsMixedCase, rawCasingMap) {
  const lowerBare = bare.toLowerCase();
  
  // Tier 1: Statute code initialisms
  const STATUTE_CODE_ACRONYMS = new Set([
    'bpc', 'com', 'civ', 'ccp', 'corp', 'edc', 'elec', 'evid', 'fam', 'fin',
    'fgc', 'fac', 'gov', 'hnc', 'hsc', 'ins', 'lab', 'mvc', 'pen', 'prob',
    'pcc', 'prc', 'puc', 'rtc', 'shc', 'uic', 'veh', 'wat', 'wic'
  ]);
  
  if (STATUTE_CODE_ACRONYMS.has(lowerBare)) {
    return true;
  }
  
  // Tier 2: Short all-consonant tokens (length <= 5)
  // y is treated as vowel here
  if (bare.length <= 5 && /^[bcdfghjklmnpqrstvwxz]+$/i.test(bare)) {
    return true;
  }
  
  // Tier 3: Single repeated letter (AA, BB, etc.) length 2-4
  if (bare.length >= 2 && bare.length <= 4) {
    const firstChar = bare.charAt(0).toLowerCase();
    if (bare.toLowerCase().split('').every(c => c === firstChar)) {
      return true;
    }
  }
  
  // Tier 4: Short token (2-5 chars) with doubled letters in mixed-case context
  // Only fires if input is mixed-case AND this token was written in caps
  if (bare.length >= 2 && bare.length <= 5 && inputIsMixedCase) {
    const wasAllCaps = rawCasingMap.get(lowerBare)?.original === bare.toUpperCase();
    if (wasAllCaps) {
      // Check for doubled letters
      const hasDoubled = /([a-z])\1/i.test(bare);
      if (hasDoubled) {
        // Exclude common words with doubled letters
        const commonDoubled = new Set([
          'app', 'add', 'ann', 'att', 'bee', 'book', 'call', 'cell', 'cool',
          'deep', 'door', 'egg', 'feed', 'feel', 'feet', 'fill', 'food', 'free',
          'good', 'hill', 'jazz', 'keep', 'kiss', 'less', 'look', 'meet', 'mood',
          'moon', 'need', 'noon', 'pass', 'peek', 'pool', 'poor', 'pull', 'room',
          'seek', 'seem', 'seen', 'sell', 'soon', 'tall', 'tell', 'took', 'tool',
          'tree', 'week', 'well', 'will', 'wood', 'wool', 'zoo'
        ]);
        if (!commonDoubled.has(lowerBare)) {
          return true;
        }
      }
    }
  }
  
  return false;
}

/**
 * Resolves short names for a list of party names on a single side of the
 * caption (e.g. all plaintiffs, or all defendants). Default short-name
 * width is 2 words, but if two entities on the same side would collide
 * at width 2, those colliding parties are widened (independently of the
 * non-colliding ones) until the side is unambiguous or we run out of
 * words to add.
 *
 * Returns a Map<originalName, resolvedShortName>. Individuals are omitted
 * from the map (they have no short name).
 */
function resolveShortNames(rawNames) {
  // Step 1: for each entity name, compute its core-name word array (the
  // raw material from which the short name is sliced). Non-entities are
  // skipped entirely.
  const entityRecords = []; // { name, words }
  for (const raw of rawNames) {
    if (!raw) continue;
    // Use a probe call with a generous wordCount to derive maximum-width
    // short name → split that back into words. This re-uses formatPartyName's
    // entity detection so we don't duplicate the parsing logic.
    const probe = formatPartyName(raw, 999);
    if (!probe.shortName) continue; // individual, no short name needed
    const words = probe.shortName.split(/\s+/).filter(Boolean);
    entityRecords.push({ name: raw, words });
  }

  // Step 2: assign each entity an initial width of 2 (or fewer if the core
  // is shorter than 2 words), then iteratively widen any group that still
  // collides until they're unique or we hit the maximum width.
  const widths = entityRecords.map(r => Math.min(2, r.words.length));

  const candidateAt = (i) => sliceShortName(entityRecords[i].words, widths[i]).join(' ');

  // Iterate at most maxWidth times — bounded by the longest core name.
  const maxWidth = entityRecords.reduce((m, r) => Math.max(m, r.words.length), 0);
  for (let pass = 0; pass < maxWidth; pass++) {
    // Group indices by current candidate.
    const groups = new Map();
    for (let i = 0; i < entityRecords.length; i++) {
      const c = candidateAt(i);
      if (!groups.has(c)) groups.set(c, []);
      groups.get(c).push(i);
    }
    // Any group with > 1 member collides → widen each, if possible.
    let widened = false;
    for (const [, indices] of groups) {
      if (indices.length < 2) continue;
      for (const i of indices) {
        if (widths[i] < entityRecords[i].words.length) {
          widths[i] += 1;
          widened = true;
        }
      }
    }
    if (!widened) break; // either everyone's unique, or no one can grow
  }

  // Step 3: build the result map.
  const result = new Map();
  for (let i = 0; i < entityRecords.length; i++) {
    result.set(entityRecords[i].name, candidateAt(i));
  }
  return result;
}

/**
 * Joins a list of "other" party names with proper formatting.
 *
 * Output structure:
 *   [(TitleShort), ]Name1[, Name2[, and Name3]] (collectively "RoleType")
 *
 * - If `titleShortName` is provided, it's emitted as a leading `(TitleShort), `
 *   parenthetical. This is used when the title party is an entity AND there
 *   are 2+ parties on that side, so the short-name reference for the title
 *   party lives in the "Other" field instead of polluting the heading.
 * - Each entity in the names list renders as "Full Name LLC (Short Name)",
 *   pulling the resolved short name from `shortNameMap` (which the caller
 *   builds via resolveShortNames so colliding parties on the same side
 *   each get unique short names). Individuals render as just their
 *   formatted name.
 * - Oxford comma for 3+ items.
 */
function formatOthers(names, roleType, titleShortName, shortNameMap) {
  // Render each name; entities get a trailing "(Short Name)" parenthetical.
  const cleaned = names
    .map(n => {
      const r = formatPartyName(n);
      if (!r.formatted) return '';
      const resolvedShort = shortNameMap && shortNameMap.get(n);
      return resolvedShort
        ? r.formatted + ' (' + resolvedShort + ')'
        : r.formatted;
    })
    .filter(Boolean);
  
  if (cleaned.length === 0) return '';
  
  // Build the result starting with the title short-name parenthetical if any.
  let result = '';
  if (titleShortName) {
    result = '(' + titleShortName + ')';
  }
  
  if (cleaned.length === 1) {
    if (result) result += ', ';
    else result = ', ';
    result += cleaned[0];
    return roleType ? result + ' (collectively "' + roleType + '")' : result;
  }
  
  // For multiple parties, add comma separator after the title short name.
  if (result) result += ', ';
  else result = ', ';
  
  // Format the list
  if (cleaned.length === 2) {
    result += cleaned[0] + ' and ' + cleaned[1];
  } else {
    // 3+ names: use Oxford comma
    const allButLast = cleaned.slice(0, -1).join(', ');
    result += allButLast + ', and ' + cleaned[cleaned.length - 1];
  }
  
  // Add collective label if specified
  if (roleType) {
    result += ' (collectively "' + roleType + '")';
  }
  
  return result;
}

/**
 * Formats every party on a side into a single combined list — used when the
 * form has one field per side rather than separate "Title" / "Other" fields.
 * (Cross-* parties on the user's Microsoft Form work this way.)
 *
 * Output structure:
 *   single party:   "Name1"   (or "Name1 LLC (Name1Short)" for an entity)
 *   two parties:    "Name1 and Name2"
 *   3+ parties:     "Name1, Name2, and Name3"
 *   2+ parties also get a trailing ` (collectively "RoleType")` suffix.
 *
 * Unlike formatOthers, there's no leading short-name parenthetical (no title
 * party lives in a separate field) and no leading comma.
 */
function formatCombinedList(names, roleType, shortNameMap) {
  const cleaned = names
    .map(n => {
      const r = formatPartyName(n);
      if (!r.formatted) return '';
      const resolvedShort = shortNameMap && shortNameMap.get(n);
      return resolvedShort
        ? r.formatted + ' (' + resolvedShort + ')'
        : r.formatted;
    })
    .filter(Boolean);

  if (cleaned.length === 0) return '';

  let result;
  if (cleaned.length === 1) {
    result = cleaned[0];
  } else if (cleaned.length === 2) {
    result = cleaned[0] + ' and ' + cleaned[1];
  } else {
    const allButLast = cleaned.slice(0, -1).join(', ');
    result = allButLast + ', and ' + cleaned[cleaned.length - 1];
  }

  // Collective suffix only when there are multiple parties — a single party
  // doesn't need a "collectively" label.
  if (roleType && cleaned.length >= 2) {
    result += ' (collectively "' + roleType + '")';
  }

  return result;
}


/**
 * Collects every NON-party name on the Parties page — attorneys, law firms, and
 * "Non-Party" type entities (e.g. receivers) — for the pseudonym generator.
 * Column positions are found by HEADER TEXT (so layout order doesn't matter):
 * an Attorney/Counsel column and, for Non-Party detection, the Party Type +
 * Name columns. Names are split on ';'/newlines, stripped of bar numbers /
 * "Esq." / "in pro per" noise and parentheticals, and de-duplicated. Caller
 * removes any that coincide with caption parties.
 */
function parseNonPartyNames(root) {
  root = root || document;
  const anchors = root.querySelectorAll('a[title="UPDATE PARTY"]');
  if (!anchors.length) return [];
  const firstRow = anchors[0].closest('tr');
  const table = firstRow && firstRow.closest('table');
  if (!table) return [];
  const allRows = Array.from(table.querySelectorAll('tr'));

  // Locate columns by header text (skip party rows; headers have no anchor).
  let nameIdx = -1, typeIdx = -1, attyIdx = -1;
  for (const r of allRows) {
    if (r.querySelector('a[title="UPDATE PARTY"]')) continue;
    const cells = Array.from(r.children);
    if (cells.length < 3) continue;
    const texts = cells.map(c => (c.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase());
    const ai = texts.findIndex(t => t.length < 40 && /attorney|counsel|represent/.test(t));
    const ni = texts.findIndex(t => t.length < 30 && /\bname\b/.test(t));
    const ti = texts.findIndex(t => t.length < 30 && /party[\s-]*type|^type$/.test(t));
    if (ai !== -1 || (ni !== -1 && ti !== -1)) { attyIdx = ai; nameIdx = ni; typeIdx = ti; break; }
  }
  try { console.log('[LACourt] non-party columns:', { nameIdx, typeIdx, attyIdx }); } catch (_) {}
  if (attyIdx === -1 && typeIdx === -1) return [];

  const out = [], seen = new Set();
  const add = raw => {
    const s = (raw || '').replace(/\s+/g, ' ').trim();
    if (!s) return;
    for (let part of s.split(/\s*;\s*|\n+/)) {
      part = part
        .replace(/\([^)]*\)/g, ' ')                                   // "(SBN 12345)" etc.
        .replace(/\b(state\s*)?bar\s*(no\.?|#)?\s*\d+/ig, ' ')          // bar numbers
        .replace(/\besq\.?/ig, ' ')                                     // "Esq." / "Esq"
        .replace(/\b(in\s+)?pro\s+per\b|\bpro\s+se\b|\bself[-\s]?represented(?:\s+litigant)?\b/ig, ' ')
        .replace(/[,\s]+$/, '').replace(/^[,\s]+/, '').replace(/\s+/g, ' ').trim();
      // "Self-represented Litigant" is a representation status, not a person —
      // drop it (and the bare "Litigant" residue) rather than treat it as a name.
      if (part.length >= 2 && /[a-z]/i.test(part) && !/^litigant$/i.test(part)) {
        const key = part.toLowerCase();
        if (!seen.has(key)) { seen.add(key); out.push(part); }
      }
    }
  };

  for (const r of allRows) {
    if (!r.querySelector('a[title="UPDATE PARTY"]')) continue;
    const cells = Array.from(r.children);
    if (attyIdx !== -1 && cells[attyIdx]) add(cells[attyIdx].textContent);
    if (typeIdx !== -1 && cells[typeIdx]) {
      const type = (cells[typeIdx].textContent || '').toLowerCase();
      if (/non-?party/.test(type) && nameIdx !== -1 && cells[nameIdx]) {
        add((cells[nameIdx].textContent || '').replace(/\([^)]*\)/g, ' '));
      }
    }
  }

  // The parties table's Representation column lists attorney names but not their
  // firms. Firm names live in a separate REPRESENTATION table whose header has a
  // "Firm Name" column — pull the attorney Name and Firm Name from each row.
  for (const t of root.querySelectorAll('table')) {
    let firmIdx = -1, repNameIdx = -1, headerSeen = false;
    for (const r of t.querySelectorAll('tr')) {
      const cells = Array.from(r.children);
      if (cells.length < 2) continue;
      const lower = cells.map(c => (c.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase());
      if (!headerSeen) {
        const fi = lower.findIndex(s => s.length < 30 && /firm\s*name/.test(s));
        if (fi === -1) continue;
        firmIdx = fi;
        repNameIdx = lower.findIndex(s => s.length < 30 && /\bname\b/.test(s) && !/firm/.test(s));
        headerSeen = true;
        continue;
      }
      if (firmIdx !== -1 && cells[firmIdx]) add(cells[firmIdx].textContent);
      if (repNameIdx !== -1 && cells[repNameIdx]) add(cells[repNameIdx].textContent);
    }
  }
  return out;
}

// Matches both LA Superior Court case-number formats:
//   - Current year-first: 2 digits + location/type letters + sequence digits,
//     e.g. "25STCV12345", "21STCR00001".
//   - Legacy district-prefix: a district letter + a case-type letter
//     (C civil, D family, F paternity, P probate, Q DV, S special, T adoption)
//     + a six-digit sequence, e.g. "BC654321", "SC123456". Unlimited civil in
//     Central (Stanley Mosk) is the familiar "BC" prefix. Cases filed before
//     the 2017/2018 rollout use this legacy form.
const CASE_NUMBER_RE = /\b(?:\d{2}[A-Z]{4,6}\d{4,6}|[A-Z][CDFPQST]\d{6})\b/;

/**
 * Finds the case number on the page. Prefers the authoritative `caseNumber`
 * URL query param (present on every eCourt case page), then the page title,
 * then prominent header elements, then a whole-page scan — matching both the
 * current and legacy formats.
 */
function parseCaseNumber(root) {
  root = root || document;

  // 1) The URL query param is authoritative and format-agnostic. It reflects
  //    the current case on every eCourt case page, whatever `root` we parse.
  try {
    const q = (new URLSearchParams(location.search).get('caseNumber') || '').trim();
    if (q && /^[0-9A-Z]{5,20}$/i.test(q)) return q;
  } catch (_) {}

  // 2) The page title leads with the case number, e.g. "BC123456: DOCUMENTS ...".
  const titleM = (root.title || '').match(CASE_NUMBER_RE);
  if (titleM) return titleM[0];

  // 3) Prominent header elements.
  const candidates = ['#caseNumber', '.case-number', '[data-case-number]', 'h1', 'h2', 'h3'];
  for (const sel of candidates) {
    const el = root.querySelector(sel);
    if (el) {
      const m = (el.textContent || '').match(CASE_NUMBER_RE);
      if (m) return m[0];
    }
  }

  // 4) Fallback: scan the whole document for a case-number-shaped token.
  //    Use textContent (a fetched/parsed doc has no layout, so innerText is '').
  const body = root.body;
  const m = (body ? (body.innerText || body.textContent || '') : '').match(CASE_NUMBER_RE);
  return m ? m[0] : '';
}



/**
 * Finds the motion type from the "Next Event" indicator. Returns the text
 * following "Hearing on" up to (but not including) " in Department ...".
 * Returns '' if no Hearing-on event is shown.
 *
 * Looks at both the title attribute (preferred — usually has "in Department"
 * suffix that bounds the match) and the visible text content as fallback.
 */
function parseMotionType(root) {
  root = root || document;
  const re = /Hearing on\s+(.+?)(?:\s+in\s+Department\b.*)?$/i;

  // Look at every span with a title (cheap; the page has few of them).
  const spans = root.querySelectorAll('span[title]');
  for (const span of spans) {
    const title = (span.getAttribute('title') || '').trim();
    if (title) {
      const m = title.match(re);
      if (m) return stripTrailingParenNumber(stripEventId(m[1]));
    }
    const text = (span.textContent || '').trim().replace(/\s+/g, ' ');
    if (text) {
      const m = text.match(re);
      if (m) return stripTrailingParenNumber(stripEventId(m[1]));
    }
  }
  return '';
}

/**
 * Returns the full hearing-type description from the "Next:" event
 * indicator. Unlike parseMotionType, this is NOT gated on the "Hearing on"
 * prefix — events like "Order to Show Cause Re: Failure to Prosecute
 * Default Judgment" don't have that prefix but still need to be detected
 * for the OSC alternate-form flow.
 *
 * Strategy: find the same span(s) parseMotionType / parseHearingDate look
 * at, then grab everything after the date+time prefix.
 *
 * Example input "Next: 05/27/2026 8:30 AM Order to Show Cause Re:
 * Failure to Prosecute Default Judgment" → returns "Order to Show Cause
 * Re: Failure to Prosecute Default Judgment".
 *
 * Example input "Next: 05/27/2026 8:30 AM Hearing on Motion for Summary
 * Judgment in Department 73" → returns "Hearing on Motion for Summary
 * Judgment in Department 73" (the caller can do further parsing if
 * needed).
 */
function parseHearingType(root) {
  root = root || document;
  // Match "Next:" prefix + date + time, then capture everything that
  // follows. Tolerant of optional "in Department NN" suffix.
  const re = /Next:?\s*\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s*(?:AM|PM)\s+(.+?)\s*$/i;

  const spans = root.querySelectorAll('span[title]');
  for (const span of spans) {
    const title = (span.getAttribute('title') || '').trim();
    if (title) {
      const m = title.match(re);
      if (m) return stripTrailingParenNumber(stripEventId(m[1]));
    }
    const text = (span.textContent || '').trim().replace(/\s+/g, ' ');
    if (text) {
      const m = text.match(re);
      if (m) return stripTrailingParenNumber(stripEventId(m[1]));
    }
  }
  return '';
}

/**
 * Parses the case name from the e-court page header. The case name is
 * rendered alongside the case number in an element matched by
 * [class*="case"]. Observed textContent format:
 *
 *   "25STCV12345 ACME HOLDINGS, LP, et al. vs TAYLOR ROE ReactDOM.render(...)"
 *
 * The case number prefix and the trailing ReactDOM.render(…) noise both
 * leak into textContent because of how the page is built. We strip both.
 *
 * Returns the case name with its original casing preserved (so "et al."
 * stays lowercase, "vs" stays lowercase, etc.) Returns '' if nothing
 * recognizable is found.
 */
function parseCaseName(caseNumberHint, root) {
  root = root || document;
  const caseNumberRe = caseNumberHint
    ? new RegExp('\\b' + caseNumberHint.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b')
    : CASE_NUMBER_RE;

  // Cut at trailing JS noise. ReactDOM.render is the observed leak; we
  // also defensively cut at any '<' (HTML), 'function(' (JS), or a stray
  // semicolon followed by space (statement separator).
  const noiseRe = /\s+(?:ReactDOM\.|React\.createElement|function\s*\(|var\s+\w+\s*=|window\.)/;

  const candidates = root.querySelectorAll('[class*="case"]');
  for (const el of candidates) {
    let text = (el.textContent || '').replace(/\s+/g, ' ').trim();
    if (!text) continue;

    // Cut at JS noise first.
    const noiseMatch = text.match(noiseRe);
    if (noiseMatch) text = text.substring(0, noiseMatch.index).trim();

    // Find the case number — accept it appearing anywhere, take everything
    // immediately after it.
    const m = text.match(caseNumberRe);
    if (!m) continue;
    let name = text.substring(m.index + m[0].length).trim();
    if (!name) continue;

    // Strip a stray "v.", ",", etc. at the very start — shouldn't happen
    // in practice but defensive.
    name = name.replace(/^[,;:\-\s]+/, '');

    if (name) return name;
  }
  return '';
}


/**
 * Finds the hearing date in the "Next Event" indicator (same span as the
 * motion type). Returns MM/DD/YYYY or '' if no Next-Event date is shown.
 *
 * Captured for ANY upcoming event (Hearing on, Status Conference, etc.) —
 * not gated on "Hearing on" the way motionType is — so a Status Conference
 * date still flows into the form's Hearing Date field.
 */
function parseHearingDate(root) {
  root = root || document;
  const dateRe = /\b(\d{1,2}\/\d{1,2}\/\d{4})\b/;
  const eventRe = /next\b/i;

  const spans = root.querySelectorAll('span[title]');
  for (const span of spans) {
    const title = (span.getAttribute('title') || '').trim();
    const text = (span.textContent || '').trim().replace(/\s+/g, ' ');

    // Only consider spans that look like the Next Event indicator.
    if (eventRe.test(title)) {
      const m = title.match(dateRe);
      if (m) return m[1];
    }
    if (eventRe.test(text)) {
      const m = text.match(dateRe);
      if (m) return m[1];
    }
  }
  return '';
}

/* ------------------------------------------------------------------ */
/* Automatic Movant detection                                          */
/* ------------------------------------------------------------------ */
//
// At Export time we background-fetch the case's Documents page (same
// authenticated origin), find the moving paper that matches the upcoming
// hearing's motion type, and read its "Filed By" party(ies). We then resolve
// each filer's role against the live Parties roster (so a receiver shows as
// "Receiver", not the grid's generic "Non-Party") and apply the user's rule:
//   - all parties of a role moving  -> the role, pluralized ("Defendants")
//   - only some moving               -> their names, without the role
//   - "et al." (truncated filer list) -> treat as all of that role
// Everything degrades to '' (blank, manual) on any failure so Export never
// breaks.










function pluralizeRole(role, count) {
  if (!role) return role;
  if (count <= 1) return role;
  return role.endsWith('s') ? role : role + 's';
}

function joinMovantNames(names) {
  const a = names.filter(Boolean);
  if (a.length <= 1) return a[0] || '';
  if (a.length === 2) return a[0] + ' and ' + a[1];
  return a.slice(0, -1).join(', ') + ', and ' + a[a.length - 1];
}

function movantNameMatch(a, b) {
  return a === b || a.includes(b) || b.includes(a);
}

function canonicalMovantRole(raw) {
  const r = (raw || '').toLowerCase();
  if (/cross[-\s]?complainant/.test(r)) return 'Cross-Complainant';
  if (/cross[-\s]?defendant/.test(r)) return 'Cross-Defendant';
  if (r.startsWith('plaintiff')) return 'Plaintiff';
  if (r.startsWith('defendant')) return 'Defendant';
  if (r.startsWith('petitioner')) return 'Petitioner';
  if (r.startsWith('respondent')) return 'Respondent';
  return raw;
}

// "Appellant" is an appellate designation a party carries IN ADDITION to its
// trial-court role (e.g. a Defendant who appeals is listed as both). This is a
// trial court, so the movant is never labeled "Appellant" — the designation is
// dropped wherever it would otherwise become a movant's role.
function isAppellateRole(role) { return /^appellants?$/i.test((role || '').trim()); }

// Reads the parties table into a movant roster: every party row's name and an
// "effective role" for labeling. Unlike parsePartiesTable (which only tracks
// the standard caption roles for the rotation/fill flow), this also captures
// party types like "Non-Party (Receiver)" -> "Receiver" so a receiver movant
// renders the way it appears on eCourt. Works on the live page (default) or a
// fetched Parties document (root override).
//
// roster: { byName: Map(normName -> role), byRole: Map(role -> Set(normName)) }
function buildMovantRoster(root) {
  root = root || document;
  const byName = new Map(), byRole = new Map();
  const rolesByName = new Map(); // normName -> [roles], resolved to one below

  let anchors = [];
  try { anchors = Array.from(root.querySelectorAll('a[title="UPDATE PARTY"]')); } catch (_) {}

  const STD_ROLE_RE = /^(cross[-\s]?complainant|cross[-\s]?defendant|plaintiff|defendant|petitioner|respondent)\b/i;
  const TYPE_QUALIFIER_RE = /^(?:non-?party|other|interested\s+party)\s*\(([^)]+)\)/i;
  const BARE_TYPE_RE = /^(non-?party|receiver|trustee|guardian|intervenor|claimant|creditor|appellant|garnishee)\b/i;

  const seenRows = new Set();
  for (const a of anchors) {
    const row = a.closest('tr');
    if (!row || seenRows.has(row)) continue;
    seenRows.add(row);

    const cells = Array.from(row.querySelectorAll('td'))
      .map(td => (td.textContent || '').trim().replace(/\s+/g, ' ')).filter(Boolean);
    if (!cells.length) continue;

    // Effective role: prefer a specific party-type qualifier like
    // "Non-Party (Receiver)"; then a standard caption role; then a bare
    // non-standard party type.
    let role = '';
    for (const c of cells) {
      const m = c.match(TYPE_QUALIFIER_RE);
      if (m) { role = m[1].trim(); break; }
    }
    if (!role) {
      for (const c of cells) {
        const m = c.match(STD_ROLE_RE);
        if (m) { role = canonicalMovantRole(m[1]); break; }
      }
    }
    if (!role) {
      for (const c of cells) {
        if (BARE_TYPE_RE.test(c) && c.length < 40) {
          role = c.replace(/\s*\([^)]*\)\s*$/, '').trim();
          break;
        }
      }
    }

    // Name: first cell that isn't a role/type/action/index cell; strip any
    // trailing parenthetical (e.g. "Alex Roe (Non-Party)" -> "Alex Roe").
    let name = '';
    for (const c of cells) {
      if (/^(update\s*party|edit|delete|view|action)$/i.test(c)) continue;
      if (/^\d+\.?$/.test(c)) continue;
      if (STD_ROLE_RE.test(c) || TYPE_QUALIFIER_RE.test(c) || BARE_TYPE_RE.test(c)) continue;
      name = c; break;
    }
    if (name) { const p = name.indexOf('('); if (p !== -1) name = name.substring(0, p).trim(); }
    if (!name || !role) continue;

    const nn = movantNormName(name);
    if (!rolesByName.has(nn)) rolesByName.set(nn, []);
    rolesByName.get(nn).push(role);
  }

  // Resolve one effective role per party. A party can appear under several roles
  // (e.g. "Defendant" and "Appellant" when a defendant appeals). This is a trial
  // court, so the movant is NEVER labeled "Appellant" — that appellate
  // designation is dropped and the party's substantive role is used. Among the
  // remaining roles the last one wins, matching prior behavior. A party whose
  // only captured role is "Appellant" is left out of the roster so formatMovant
  // resolves it by the role it filed under instead.
  for (const [nn, roles] of rolesByName) {
    const nonAppellate = roles.filter(r => !isAppellateRole(r));
    if (!nonAppellate.length) continue;
    const chosen = nonAppellate[nonAppellate.length - 1];
    byName.set(nn, chosen);
    if (!byRole.has(chosen)) byRole.set(chosen, new Set());
    byRole.get(chosen).add(nn);
  }

  return { byName, byRole };
}

function formatMovant(parties, truncated, roster) {
  const groups = new Map(); // role -> [display names]
  for (const p of parties) {
    if (!p.name || /^clerk$/i.test(p.name)) continue;
    let role = roster.byName.get(movantNormName(p.name)) || p.role || '';
    // Trial court: the movant is never "Appellant". If that's the only label
    // available, drop it and group the party by name alone rather than emit it.
    if (isAppellateRole(role)) role = '';
    if (!groups.has(role)) groups.set(role, []);
    groups.get(role).push(p.name);
  }
  const out = [];
  for (const [role, names] of groups) {
    const rosterSet = roster.byRole.get(role);
    let all = false;
    if (truncated) {
      all = true; // user rule: "et al." means all of that role
    } else if (rosterSet && rosterSet.size) {
      const fn = names.map(movantNormName);
      all = [...rosterSet].every(rn => fn.some(f => movantNameMatch(f, rn)));
    }
    if (all && role) out.push(pluralizeRole(role, rosterSet ? rosterSet.size : names.length));
    else out.push(joinMovantNames(names));
  }
  return out.filter(Boolean).join('; ');
}

// Finds a case tab's URL by its visible link text (e.g. "Documents",
// "Parties"). The case sub-nav is present on every case page.
function getCaseTabUrl(label) {
  return caseTabUrlFrom(document, label);
}

// Documents-page URL: the "Documents" tab link, else swap formId=279 in.
function getDocumentsUrl() {
  const link = getCaseTabUrl('documents');
  if (link) return link;
  try {
    const u = new URL(location.href);
    u.searchParams.set('formId', '279');
    return u.toString();
  } catch (_) {
    return null;
  }
}

// Parties-page URL: the "Parties" tab link (no reliable formId fallback).
function getPartiesUrl() {
  return getCaseTabUrl('parties');
}

// Hearings-page URL: the "Hearings" tab link, else swap formId=395 in.
function getHearingsUrl() {
  const link = getCaseTabUrl('hearings');
  if (link) return link;
  try {
    const u = new URL(location.href);
    u.searchParams.set('formId', '395');
    return u.toString();
  } catch (_) {
    return null;
  }
}

/* ------------------------------------------------------------------ */
/* Hearing selection via the agenda exclusion terms                    */
/* ------------------------------------------------------------------ */
//
// The case header's "Next" event is sometimes a routine hearing (a
// conference, an OSC re: sanctions, an ex parte) that isn't the motion the
// order is for. When the Next event matches the agenda exclusion list, we
// look at the Hearings tab and use the soonest FUTURE, SCHEDULED, non-excluded
// hearing instead. The exclusion list is the same `excludedTerms` the agenda
// cleaner uses (chrome.storage.sync), so editing it in options affects both.







// Resolves the effective hearing: the Next event when it's something we work up,
// otherwise the soonest future scheduled hearing on the Hearings tab that IS a
// motion we work up or a default judgment. Returns { motionType, hearingDate,
// hearingType }.
// The hearing everything on the page is keyed to: the first (or only) motion on
// the selected hearing date. Falls back to reading the header directly when the
// hearing list can't be built (no workable hearing on the case, or the Hearings
// tab was unreachable).
async function resolveEffectiveHearing(root) {
  const g = await getSelectedHearingGroup().catch(() => null);
  if (g && g.items.length) {
    const it = g.items[0];
    return {
      motionType: it.motionType,
      hearingDate: it.hearingDate,
      hearingType: it.hearingType,
      lookedAhead: !it.native,
    };
  }
  return resolveEffectiveHearingFromHeader(root);
}

async function resolveEffectiveHearingFromHeader(root) {
  root = root || document;
  const nextType = parseHearingType(root);
  const base = {
    motionType: parseMotionType(root),
    hearingDate: parseHearingDate(root),
    hearingType: nextType,
  };
  try {
    await loadExcludedTerms();
    if (isWorkableHearing(nextType)) return base;

    const url = getHearingsUrl();
    if (!url) return base;
    const doc = await fetchCaseDoc(url);
    if (!doc) return base;

    const hearings = parseFutureHearings(doc).filter(h => isWorkableHearing(h.type));
    if (!hearings.length) return base;

    // Take the soonest workable hearing. When several fall on that same date,
    // prefer a Demurrer (with Motion to Strike) over a standalone Motion to
    // Strike — a demurrer + motion to strike is filed together but shows as two
    // hearing entries, and the demurrer is the one we want.
    const soonestDate = hearings[0].date;
    const sameDay = hearings.filter(h => h.date === soonestDate);
    const pick = sameDay.find(h => /demurrer/i.test(h.type)) || sameDay[0];

    console.log('[LACourt] Next event not worked up (' + nextType +
      '); using Hearings-tab pick:', pick.type, pick.date);
    return {
      motionType: stripHearingOnPrefix(pick.type),
      hearingDate: pick.date,
      hearingType: pick.type,
      lookedAhead: true,
    };
  } catch (err) {
    console.warn('[LACourt] hearing resolution failed:', err);
    return base;
  }
}

// Given a parsed HTML Document, pull moving-paper filings: [{name, filedBy}].
function parseDocumentsFilingsFrom(doc) {
  const tables = doc.querySelectorAll('table');
  for (const table of tables) {
    let headerRow = null, nameIdx = -1, filedByIdx = -1;
    for (const tr of table.querySelectorAll('tr')) {
      const texts = Array.from(tr.children).map(td => (td.textContent || '').replace(/\s+/g, ' ').trim());
      const fb = texts.indexOf('Filed By');
      const nm = texts.indexOf('Name');
      if (fb !== -1 && nm !== -1) { headerRow = tr; filedByIdx = fb; nameIdx = nm; break; }
    }
    if (!headerRow) continue;

    const filings = [];
    let started = false;
    for (const tr of table.querySelectorAll('tr')) {
      if (tr === headerRow) { started = true; continue; }
      if (!started) continue;
      const cells = Array.from(tr.children);
      if (cells.length <= filedByIdx) continue;
      const name = (cells[nameIdx] ? cells[nameIdx].textContent : '').replace(/\s+/g, ' ').trim();
      const filedBy = (cells[filedByIdx] ? cells[filedByIdx].textContent : '').replace(/\s+/g, ' ').trim();
      if (!name || !filedBy) continue;
      if (!isMovingPaper(name)) continue;
      filings.push({ name, filedBy });
    }
    if (filings.length) return filings;
  }
  return [];
}



// Async: resolves to the Movant string, or '' on any failure / no match.
// `partiesRoot` is the document the party roster is read from (the live page
// when on Parties, or a fetched Parties document otherwise).
async function computeMovant(motionType, partiesRoot, hearingDate) {
  try {
    if (!motionType) return '';

    // Primary: the Documents page's first view (dedicated header-aware parser).
    let best = null;
    const url = getDocumentsUrl();
    if (url) {
      const doc = await fetchCaseDoc(url);
      if (doc) best = bestFilingMatch(motionType, parseDocumentsFilingsFrom(doc));
    }

    // Fallback: the full paginated document cache, so a moving paper filed
    // beyond the first Documents page is still found (busy cases).
    if (!best) {
      const all = await getAllDocumentsCached();
      const filings = (all || [])
        .filter(d => d && d.name && d.filedBy && isMovingPaper(d.name))
        .map(d => ({ name: d.name, filedBy: d.filedBy }));
      best = bestFilingMatch(motionType, filings);
    }

    // Last resort: let the CALENDAR name the moving paper. A petition's hearing
    // caption and the petition's own caption routinely describe the same relief
    // in different words ("Petition to Confirm Minor's Compromise" heard on a
    // "Petition to Approve Compromise of Disputed Claim …"), so neither name
    // match above finds it; resolveMovingPaper pairs them by hearing date.
    if (!best) {
      const [all, hearings] = await Promise.all([getAllDocumentsCached(), getFutureHearingsCached()]);
      const md = resolveMovingPaper(motionType, parseHearingDateTime(hearingDate), hearings, all || []);
      if (md && md.filedBy) best = md;
    }
    if (!best) return '';

    const { parties, truncated } = parseFiledByParties(best.filedBy);
    if (!parties.length) return '';
    const roster = buildMovantRoster(partiesRoot || document);
    const movant = formatMovant(parties, truncated, roster);
    console.log('[LACourt] movant detected:', { motionType, doc: best.name, filedBy: best.filedBy, movant });
    return movant || '';
  } catch (err) {
    console.warn('[LACourt] movant detection failed:', err);
    return '';
  }
}

// Builds the Export context, pulling party data from the current page when it
// has the parties table, or by background-fetching the Parties page otherwise
// (so Export works from Documents/Summary/any case page). Resolves to
// { ctx, partiesRoot } or null.
async function getExportContext() {
  // 1) Resolve which document to read the party roster from (live page when it
  //    has the parties table, else a background-fetched Parties page).
  let partiesRoot = document;
  if (!document.querySelector('a[title="UPDATE PARTY"]')) {
    const url = getPartiesUrl();
    if (url) {
      const doc = await fetchCaseDoc(url);
      if (doc) partiesRoot = doc;
    }
  }

  // 2) Resolve the effective hearing (may fetch the Hearings tab when the Next
  //    event is excluded), then build the context with that override.
  const hearing = await resolveEffectiveHearing(partiesRoot);
  // Two motions heard the same morning are worked up together and go out on ONE
  // order template, so the form's Motion Type box names both, separated by "; ".
  const group = await getSelectedHearingGroup().catch(() => null);
  if (group && group.items.length > 1) {
    const both = group.items.map(it => it.motionType).filter(Boolean).join('; ');
    if (both) hearing.motionType = both;
  }
  const ctx = getFillFormContext(partiesRoot, hearing);
  return ctx ? { ctx, partiesRoot } : null;
}

/* ------------------------------------------------------------------ */
/* Documents button: open the documents relevant to the motion         */
/* ------------------------------------------------------------------ */
//
// Identifies and opens (as background tabs) the documents relevant to the
// selected motion, all sourced from the Documents tab (deduped by docId):
//   - the complaint chain (original + every amended complaint through the
//     operative one) + operative cross-complaint (not fictitious-name
//     amendments)
//   - the moving paper + anything the moving party filed the same day
//   - documents the Hearings tab lists for that motion
//   - one upcoming hearing  -> everything filed after the motion
//   - multiple hearings     -> documents after the motion whose title shares a
//     meaningful word with the motion type, plus each Opposition/Reply and its
//     same-day co-filings.




function computeRelevantDocuments(docs, motionType, hearingDocBlob, singleHearing, movingPaper) {
  const rel = new Map();
  const add = d => { if (d && d.docId && d.openUrl) rel.set(d.docId, d); };

  // The complaint chain: the original complaint AND every amended complaint
  // through the operative one — an operative Second Amended Complaint implies a
  // First (often titled just "Amended Complaint (1st)") that is read alongside
  // it. isComplaintDoc anchors to the START of the title, so filings that merely
  // mention "complaint" (answers, demurrers, proofs of service, cross-complaints,
  // fictitious-name amendments) stay out. When the case has no complaint at all,
  // fall back to the operative petition (another initial pleading).
  const complaintChain = docs.filter(d => isComplaintDoc(d.name));
  if (complaintChain.length) for (const d of complaintChain) add(d);
  else add(latestDoc(docs.filter(d => isPetitionDoc(d.name))));
  add(latestDoc(docs.filter(d => isCrossComplaintDoc(d.name))));

  // Identify the moving paper up front so the Hearings-tab blob match below can
  // use its filing date as a floor. (The full motion-doc handling still runs in
  // its own block later.) The caller resolves it (pairing parallel same-named
  // demurrers to the right hearing); fall back to bestFilingMatch.
  const motionDoc = movingPaper || bestFilingMatch(motionType, docs);
  const motionFloor = motionDoc && motionDoc.when ? motionDoc.when : null;

  // Documents the Hearings tab lists for this motion (substring containment).
  // The Hearings tab is authoritative, so this runs regardless of whether we
  // can independently identify the moving paper below — otherwise a motion
  // whose filing name doesn't match (bestFilingMatch returns null) would leave
  // only the operative pleading.
  //
  // The blob lists document NAMES, not identities, and filing names are not
  // unique — a generic "Separate Statement" appears on every MSJ. So a match by
  // name alone would pull in an OLD filing that merely shares a name with one of
  // this motion's papers (e.g. a "Separate Statement" filed months earlier for a
  // different motion). Two guards prevent that:
  //   1. When we know the motion's filing date, never add a filing dated before
  //      it — nothing filed before the motion briefs the motion (the operative
  //      pleadings are added separately, above).
  //   2. When we don't (bestFilingMatch returned null, so no date), and several
  //      filings share the same name, keep only the latest.
  // Documents the Hearings tab actually listed for THIS hearing. Authoritative,
  // so they're exempt from the other-motion guard below.
  const blobDocIds = new Set();
  if (hearingDocBlob) {
    const blob = movantNormName(hearingDocBlob);
    const hits = [];
    for (const d of docs) {
      const nn = movantNormName(d.name);
      if (nn && nn.length >= 6 && blob.indexOf(nn) !== -1) hits.push({ d, nn });
    }
    for (const { d, nn } of hits) {
      if (motionFloor) {
        if (d.when && d.when < motionFloor) continue;
      } else if (hits.some(o => o.nn === nn && o.d.when && d.when && o.d.when > d.when)) {
        continue;
      }
      add(d);
      blobDocIds.add(d.docId);
    }
  }

  // For an OSC Re: Failure to Prosecute Default Judgment, the prove-up packet is
  // the paper asking the court to enter judgment on the default — a second
  // Request for Entry of Default, a CIV-100 whose title says it asked for a
  // court judgment, or the judgment papers themselves (see findDefaultProveUp)
  // — plus everything the plaintiff filed on or after it.
  if (isOscDefaultJudgment(motionType)) {
    // Every Request for Dismissal is relevant on a default-judgment prove-up:
    // dismissals of individual defendants drive severability (CCP 578-579) and
    // confirm the remaining defendants are all in default. Always flag them,
    // independent of whether the prove-up packet below is detected.
    for (const d of docs) if (/\brequest for dismissal\b/i.test(d.name || '')) add(d);

    const pu = findDefaultProveUp(docs);
    if (pu && pu.proveUp) {
      add(pu.proveUp);
      const pw = pu.proveUp.when;
      for (const d of docs) {
        if (d.when && pw && d.when >= pw && /\bplaintiff\b/i.test(d.filedBy || '')) add(d);
      }
    }
  }

  // No moving paper on file for this hearing — never filed, or withdrawn. The
  // papers that DO reference it are then the entire record of it: the notice of
  // intent, the notice of hearing, a notice that the motion was never served.
  // Those are what the button should open. Other motions' moving papers stay
  // out, for the same reason as in the sweeps below — they belong to their own
  // hearings — unless the Hearings tab listed them for this one.
  if (!motionDoc) {
    // A post-judgment motion's anchoring paper is also a floor: the notice of
    // intention (or, for reconsideration, the notice of entry of the order)
    // is what starts the matter, so nothing filed before it briefs it. Without
    // that floor a new trial motion pulled in every "…Trial" document on the
    // docket — the trial minute orders, months of them — on the shared word.
    const pjAnchorDoc = postJudgmentAnchor(motionType, docs);
    const floor = pjAnchorDoc && pjAnchorDoc.when ? pjAnchorDoc.when : null;
    for (const d of docs) {
      if (isMovingPaper(d.name) && !blobDocIds.has(d.docId)) continue;
      if (floor && d.when && d.when < floor) continue;
      if (docReferencesMotion(d.name, motionType)) add(d);
    }
  }

  // The companion motion to strike a "Demurrer - with Motion to Strike" hearing
  // picked up from the window around the demurrer (below) — exempted from the
  // parallel-challenge guard further down, which otherwise keeps only
  // challenges filed the demurrer's own day.
  const strikeCompanionIds = new Set();

  if (motionDoc) {
    add(motionDoc);
    const mov = docPartyNames(motionDoc.filedBy), mw = motionDoc.when;

    // Same-day filings by the moving party (incl. just before the motion).
    for (const d of docs) if (sameCalendarDay(d.when, mw) && docSharesParty(docPartyNames(d.filedBy), mov)) add(d);

    // A "Demurrer - with Motion to Strike" is one work-up whose motion to
    // strike is its own filing — usually the demurrer's day, but it lands a few
    // days before or after often enough that the same-day sweep above misses
    // it. Hunt the window around the demurrer for the strike's moving paper
    // (however the clerk keyed the title — "Motion re: to Strike Portions
    // of …" included) and take it plus its same-day supporting papers, the
    // same way the demurrer's own co-filings ride in. Documents button only —
    // the briefing-deadline widget is unaffected.
    if (/\bdemurrer\b/i.test(motionType || '') && /\bwith\s+motion\s+to\s+strike\b/i.test(motionType || '')) {
      for (const d of docs) {
        if (d.docId === motionDoc.docId || !isCompanionStrikeDoc(d.name)) continue;
        if (!withinCalendarDays(d.when, mw, STRIKE_COMPANION_WINDOW_DAYS)) continue;
        add(d);
        strikeCompanionIds.add(d.docId);
        const P = docPartyNames(d.filedBy);
        for (const co of docs) {
          if (!sameCalendarDay(co.when, d.when) || !docSharesParty(docPartyNames(co.filedBy), P)) continue;
          // Its declarations and memoranda, not some OTHER motion the same
          // party happened to file that day.
          if (isMovingPaper(co.name) && co.docId !== d.docId) continue;
          add(co);
        }
      }
    }

    if (singleHearing) {
      // One upcoming hearing: everything after the motion is fair game.
      for (const d of docs) if (d.when && mw && d.when > mw) add(d);
    } else {
      // A document that is itself a moving paper belongs to the hearing IT
      // noticed. The name-similarity sweeps below can't tell "Motion to Compel
      // Arbitration" from "Motion to Compel Further Responses" — docWordOverlap
      // matches on the single shared token "compel" — so a case carrying
      // parallel motions set on different dates would open every one of them.
      // Keep those sweeps to papers that aren't somebody's moving paper. This
      // hearing's own motion and anything the Hearings tab listed for it are
      // exempt (both authoritative, and both already added above). Oppositions,
      // replies, separate statements and declarations are not moving papers, so
      // the briefing this is meant to collect is unaffected.
      const isOtherMotion = d => isMovingPaper(d.name)
        && d.docId !== (motionDoc && motionDoc.docId)
        && !blobDocIds.has(d.docId);

      // Multiple hearings: match by shared words + Opposition/Reply co-filings.
      for (const d of docs) if (d.when && mw && d.when >= mw && docWordOverlap(d.name, motionType) && !isOtherMotion(d)) add(d);
      const after = docs.filter(d => d.when && mw && d.when >= mw);
      // A generically titled paper ("Opposition OPPOSITION", "Reply REPLY") has
      // no word to overlap with — every word in it is a stop word — so it takes
      // the filer instead: the movant doesn't oppose its own motion, and the
      // reply is the movant's paper. Looser than the status engine's calendar
      // test, deliberately: here a wrong guess costs one extra tab, and missing
      // the opposition costs the briefing.
      const genericOpp = d => docNameIsGeneric(d.name)
        && !docSharesParty(docPartyNames(d.filedBy), mov);
      const genericReply = d => docNameIsGeneric(d.name)
        && !(docPartyNames(d.filedBy).length && mov.length
             && !docSharesParty(docPartyNames(d.filedBy), mov));
      for (const opp of after) if (/\bopposition\b/i.test(opp.name)
          && (docWordOverlap(opp.name, motionType) || genericOpp(opp))) {
        add(opp); const P = docPartyNames(opp.filedBy);
        for (const d of docs) if (sameCalendarDay(d.when, opp.when) && docSharesParty(docPartyNames(d.filedBy), P) && !isOtherMotion(d)) add(d);
      }
      // The movant files the reply, so a "Reply …" by the same party as the
      // motion also counts — covers replies that name the motion only by the
      // party or a mangled title (e.g. "…BERNARDSDEMURRER" with no space, which
      // no word-overlap can catch).
      for (const rep of after) if (/\breply\b/i.test(rep.name)
          && (docWordOverlap(rep.name, motionType) || docSharesParty(docPartyNames(rep.filedBy), mov)
              || genericReply(rep))) {
        add(rep); const P = docPartyNames(rep.filedBy);
        for (const d of docs) if (sameCalendarDay(d.when, rep.when) && docSharesParty(docPartyNames(d.filedBy), P) && !isOtherMotion(d)) add(d);
      }
    }
  }

  // With parallel same-named challenges (e.g. a demurrer to the complaint AND a
  // demurrer to a cross-complaint), the name-based paths above can pull in the
  // OTHER demurrer/motion-to-strike, which belongs to a different hearing. Keep
  // only the current motion's moving paper among challenge documents — but keep
  // a challenge filed the SAME day as it (a demurrer + motion to strike filed
  // together belong to this hearing), and keep the companion strike the window
  // sweep above vouched for.
  if (motionDoc) {
    for (const [id, d] of rel) {
      if (d.docId !== motionDoc.docId && isDemurrerOrMotionToStrikeDoc(d.name)
          && !sameCalendarDay(d.when, motionDoc.when)
          && !strikeCompanionIds.has(d.docId)) rel.delete(id);
    }
  }

  // A fee-waiver request (and its "additional fees" variant) is an
  // administrative filing that is never relevant to a motion — always drop it.
  for (const [id, d] of rel) if (ALWAYS_IRRELEVANT_RE.test(d.name || '')) rel.delete(id);

  // Ex parte papers belong to their own proceeding, not to the noticed motion:
  // an application for an order shortening time, an application to advance or
  // continue the hearing, the opposition to one. They ride in on the word-overlap
  // sweeps (an ex parte application to continue an MSJ names the MSJ) and are
  // noise on a law-and-motion work-up — "ex parte" is already an excluded HEARING
  // term for the same reason. The Hearings tab is authoritative, as everywhere
  // else here: a paper it listed for THIS hearing stays in.
  for (const [id, d] of rel) {
    if (EX_PARTE_RE.test(d.name || '') && !blobDocIds.has(d.docId)) rel.delete(id);
  }

  // The papers that ride in on an ex parte application — a supporting
  // declaration, a memorandum, a proposed order — are routinely titled without
  // the words "Ex Parte" at all. A document filed by the SAME party as an actual
  // ex parte paper, on the SAME calendar day, belongs to that ex parte matter
  // too and gets swept out on the same theory — unless its own title is a close
  // match to the hearing being worked up, meaning it really is this motion's
  // paper and just happens to share a filing date with an unrelated ex parte.
  const exParteDocs = docs.filter(d => EX_PARTE_RE.test(d.name || ''));
  if (exParteDocs.length) {
    for (const [id, d] of rel) {
      if (blobDocIds.has(d.docId) || EX_PARTE_RE.test(d.name || '')) continue;
      if (docWordOverlap(d.name, motionType)) continue;
      const dParty = docPartyNames(d.filedBy);
      if (!dParty.length) continue;
      const isExParteCompanion = exParteDocs.some(ep =>
        sameCalendarDay(ep.when, d.when) && docSharesParty(docPartyNames(ep.filedBy), dParty));
      if (isExParteCompanion) rel.delete(id);
    }
  }

  return Array.from(rel.values());
}

// How far either side of the demurrer's filing date the companion motion to
// strike is hunted for — the parties file the pair a few days apart often
// enough that same-day-only missed it. Calendar days; 5 clears a weekend plus
// a court holiday in both directions.
const STRIKE_COMPANION_WINDOW_DAYS = 5;

// Both `when` values are Dates already normalized to midnight (see
// sameCalendarDay), so a plain difference counts whole days.
function withinCalendarDays(x, y, days) {
  return !!(x && y && Math.abs(x.getTime() - y.getTime()) <= days * 86400000);
}

// The moving paper of the strike that rides with a demurrer, however the clerk
// keyed the title: "Motion to Strike …", "Motion re: to Strike Portions of
// Plaintiff's Verified Third Amended Complaint …", "Notice of Motion and
// Motion to Strike …". A motion to strike or tax COSTS is its own motion (see
// the agenda grouping rule), and an anti-SLAPP special motion to strike is its
// own beast too — neither is ever the demurrer's companion.
function isCompanionStrikeDoc(name) {
  const n = (name || '').trim();
  // eCourt's parentheticals and trailing dash qualifiers describe what a paper
  // is NOT ("Demurrer - without Motion to Strike", "(not anti-SLAPP)") — judge
  // the core title without them, the way movant matching does.
  const core = n.replace(/\([^)]*\)/g, ' ')
    .replace(/\s[-–—]\s*(?:with|without)\b.*$/i, ' ')
    .replace(/\s+/g, ' ').trim();
  if (!/\bmotion\b(?:\s+re:?)?\s+to\s+strike\b/i.test(core)) return false;
  if (/\b(?:strike|tax|taxing)\s+(?:of\s+)?costs\b|\bcosts\b.*\b(?:strike|tax)\b|memorandum\s+of\s+costs/i.test(n)) return false;
  if (/anti-?slapp|special\s+motion/i.test(core)) return false;
  return isMovingPaper(n) || /^notice of motion and motion\b/i.test(n);
}

// Documents that are never relevant to any motion, regardless of motion type:
//   - "Request to Waive Court Fees" / "...Additional Court Fees": administrative
//     fee-waiver filings, not substantive to any hearing.
//   - "Proposed Order" / "[Proposed] Order": a lodged draft order, not a paper
//     that briefs or supports the motion.
//   - anything mentioning "Jury Fees" (e.g. "Notice of Posting of Jury Fees"):
//     an administrative fee posting, not substantive to any hearing.
const ALWAYS_IRRELEVANT_RE = /request to waive (additional )?court fees|\[?\s*proposed\s*\]?\s+order\b|jury fees/i;

// An ex parte paper, wherever "Ex Parte" falls in the title — the application
// itself ("Ex Parte Application to Continue Trial"), the opposition to one, the
// order on one. The separator is optional so a run-together "ExParte" (docket
// titles are typed by hand) is caught too.
const EX_PARTE_RE = /\bex[\s-]?parte\b/i;






// From a Hearings-tab document, returns the "Document" column text of the
// hearing matching the motion type (used to mark those documents relevant).
function findHearingDocBlob(hearingsDoc, motionType) {
  const tables = hearingsDoc.querySelectorAll('table');
  for (const table of tables) {
    let headerRow = null, nameIdx = -1, docIdx = -1;
    for (const tr of table.querySelectorAll('tr')) {
      const texts = Array.from(tr.children).map(td => (td.textContent || '').replace(/\s+/g, ' ').trim());
      const ni = texts.indexOf('Name'), di = texts.indexOf('Document');
      if (ni !== -1 && di !== -1) { headerRow = tr; nameIdx = ni; docIdx = di; break; }
    }
    if (!headerRow) continue;
    let started = false;
    for (const tr of table.querySelectorAll('tr')) {
      if (tr === headerRow) { started = true; continue; }
      if (!started) continue;
      const cells = Array.from(tr.children);
      if (cells.length <= docIdx) continue;
      const name = stripEventId((cells[nameIdx] ? cells[nameIdx].textContent : '').replace(/\s+/g, ' ').trim());
      if (name && docWordOverlap(name, motionType)) {
        return (cells[docIdx] ? cells[docIdx].textContent : '').replace(/\s+/g, ' ').trim();
      }
    }
  }
  return '';
}

// Memoized fetch of ALL documents for this case page. Shared by the relevant-
// documents computation and the inline Next-header deadline check so the page is
// fetched at most once. Resolves to [] on failure (allowing graceful degrade).
let __allDocsPromise = null;
function getAllDocumentsCached() {
  if (__allDocsPromise) return __allDocsPromise;
  const docsUrl = getDocumentsUrl();
  __allDocsPromise = (docsUrl ? fetchAllDocuments(docsUrl) : Promise.resolve([])).catch(() => []);
  return __allDocsPromise;
}

// Memoized future scheduled hearings from the Hearings tab (fetched at most once).
// Resolves to [] when the tab isn't reachable, so callers degrade gracefully.
let __futureHearingsPromise = null;
function getFutureHearingsCached() {
  if (__futureHearingsPromise) return __futureHearingsPromise;
  const url = getHearingsUrl();
  __futureHearingsPromise = (url ? fetchCaseDoc(url) : Promise.resolve(null))
    .then(doc => (doc ? parseFutureHearings(doc) : []))
    .catch(() => []);
  return __futureHearingsPromise;
}

// Orchestrates: resolve the motion, fetch all documents + hearings, compute the
// relevant set. Resolves to { relevant, motionType, docCount, singleHearing }.
async function getRelevantDocuments() {
  // Every motion set for the selected hearing date — two motions heard the same
  // morning are read together, so the button opens both sets at once (deduped).
  const group = await getSelectedHearingGroup().catch(() => null);
  const hearings = (group && group.items.length)
    ? group.items
    : [await resolveEffectiveHearing(document)];

  // The type used to match the Hearings-tab row and any moving paper. Normally
  // the motion type; for an OSC Re: Failure to Prosecute Default Judgment (which
  // isn't a "Hearing on <motion>" event, so has no motion type) fall back to the
  // OSC hearing type — the operative complaint and the Hearings-tab documents are
  // still the relevant set.
  const matchTypeOf = h => {
    const mt = h && h.motionType;
    if (mt) return mt;
    return (h && isOscDefaultJudgment(h.hearingType)) ? h.hearingType : '';
  };
  const targets = hearings.filter(h => matchTypeOf(h));
  if (!targets.length) return { relevant: [], reason: 'no-motion' };

  const docs = await getAllDocumentsCached();
  if (!docs.length) return { relevant: [], reason: 'no-documents' };

  const hearingsUrl = getHearingsUrl();
  const hearingsDoc = hearingsUrl ? await fetchCaseDoc(hearingsUrl) : null;
  const futureHearings = hearingsDoc ? parseFutureHearings(hearingsDoc) : [];
  const singleHearing = futureHearings.length <= 1;

  const merged = new Map();
  const types = [];
  for (const hearing of targets) {
    const matchType = matchTypeOf(hearing);
    types.push(matchType);
    const hearingDocBlob = hearingsDoc ? findHearingDocBlob(hearingsDoc, matchType) : '';
    // Pick the moving paper for THIS hearing, disambiguating parallel same-named
    // demurrers/motions to strike by pairing them to their hearings by date.
    const hearingWhen = hearing.hearingDate ? parseHearingDateTime(hearing.hearingDate) : null;
    const movingPaper = resolveMovingPaper(matchType, hearingWhen, futureHearings, docs);
    for (const d of computeRelevantDocuments(docs, matchType, hearingDocBlob, singleHearing, movingPaper)) {
      if (d && d.docId && !merged.has(d.docId)) merged.set(d.docId, d);
    }
  }

  const relevant = Array.from(merged.values());
  const motionType = types.join('; ');
  console.log('[LACourt] relevant documents:', {
    motionType, docCount: docs.length, singleHearing, relevant: relevant.map(d => d.name),
  });
  return { relevant, motionType, docCount: docs.length, singleHearing };
}

// Memoized wrapper so the Documents button opens instantly: the fetch + relevance
// computation runs once (prefetched on page load) and the result is reused. NO
// tabs are opened here — the button decides when to open.
let __relevantDocsPromise = null;
function getRelevantDocumentsCached() {
  if (!__relevantDocsPromise) __relevantDocsPromise = getRelevantDocuments();
  return __relevantDocsPromise;
}

/* ------------------------------------------------------------------ */
/* Subtle on-page toast confirmation                                   */
/* ------------------------------------------------------------------ */

function showToast(message) {
  const existing = document.getElementById('__lacourt_toast__');
  if (existing) existing.remove();

  const toast = document.createElement('div');
  toast.id = '__lacourt_toast__';
  toast.textContent = '⚖ ' + message;
  Object.assign(toast.style, {
    position: 'fixed',
    bottom: '24px',
    right: '24px',
    background: '#1a365d',
    color: 'white',
    padding: '10px 16px',
    borderRadius: '6px',
    fontFamily: 'Georgia, serif',
    fontSize: '13px',
    zIndex: '999999',
    boxShadow: '0 4px 12px rgba(0,0,0,0.25)',
    opacity: '0',
    transition: 'opacity 0.2s',
  });
  document.body.appendChild(toast);
  requestAnimationFrame(() => { toast.style.opacity = '1'; });
  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 2200);
}

/* ------------------------------------------------------------------ */
/* Floating "Fill Microsoft Form" button (top-right corner)            */
/* Always visible on parties pages, resets after each use.             */
/* ------------------------------------------------------------------ */

// Note: REGULAR_FORM_URL is defined at the top of the IIFE. Both regular and
// OSC Default Judgment cases run the in-extension Order Template flow;
// getFillFormContext() adds the recommendation mailto for the DJ case.

function renderFillFormButton() {
  if (document.getElementById('__lacourt_fill_btn__')) return;

  // Show on any case page once the case sub-nav has rendered. Export works
  // from anywhere now: if the current page lacks the parties table, the click
  // handler background-fetches the Parties page. Waiting on the sub-nav avoids
  // rendering before the page is usable.
  const caseReady = document.querySelector('a[href*="/ecourt/ecms/case"]');
  if (!caseReady) return;

  const btn = document.createElement('button');
  btn.id = '__lacourt_fill_btn__';
  btn.type = 'button';

  // Export icon: a curved arrow rising and curving up-and-right out of a box.
  const EXPORT_ICON =
    '<span class="lac-btn-icon" style="vertical-align:middle">' +
    '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" ' +
    'stroke="currentColor" stroke-width="3" stroke-linecap="round" ' +
    'stroke-linejoin="round" style="vertical-align:middle;margin:-4px 0">' +
    '<path d="M5 12v6a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-6"/>' +
    '<path d="M12 14V8C12 5 14 4 19 4"/>' +
    '<path d="M16 1.5 19 4l-3 2.5"/>' +
    '</svg></span>';
  const setLabel = (text) => {
    btn.innerHTML = EXPORT_ICON +
      '<span class="lac-btn-text" style="vertical-align:middle">' + text + '</span>';
  };
  setLabel('Export');
  Object.assign(btn.style, {
    position: 'fixed',
    top: '0px',
    right: '16px',
    zIndex: '999998',
    padding: '6px 16px',
    background: '#1a365d',
    color: 'white',
    border: 'none',
    borderRadius: '0 0 6px 6px',
    fontFamily: 'Georgia, serif',
    fontSize: '13px',
    fontWeight: 'bold',
    cursor: 'pointer',
    boxShadow: '0 2px 8px rgba(0,0,0,0.25)',
    transition: 'background 0.15s, opacity 0.2s',
  });
  btn.addEventListener('mouseover', () => { btn.style.background = '#2a4a7f'; });
  btn.addEventListener('mouseout',  () => { btn.style.background = '#1a365d'; });

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    setLabel('Working...');
    btn.style.opacity = '0.7';

    // Download any court PDFs the user has open in this window. The background
    // resolves once the downloads have settled; keep the button greyed out until
    // then (a timeout backs it up so the button never sticks).
    const downloadsDone = new Promise(resolve => {
      let done = false;
      const finish = () => { if (!done) { done = true; resolve(); } };
      const safety = setTimeout(finish, 60000);
      try {
        chrome.runtime.sendMessage({ type: 'downloadOpenPdfs' }, res => {
          void chrome.runtime.lastError;
          clearTimeout(safety);
          if (res && res.count) console.log('[LACourt] downloaded ' + res.count + ' open PDF(s)');
          finish();
        });
      } catch (_) { clearTimeout(safety); finish(); }
    });

    try {
      const result = await getExportContext();
      if (!result) {
        setLabel('No data found');
        setTimeout(() => {
          btn.disabled = false;
          setLabel('Export');
          btn.style.opacity = '1';
        }, 2000);
        return;
      }
      const ctx = result.ctx;

      storeRotation(ctx.data, { autoFillOnLoad: true });

      // Fire mailto first (OSC cases only) — it's instant from the user's
      // perspective and the form-open is the longer-running step. Order
      // doesn't really matter since both are async fire-and-forget.
      if (ctx.mailtoUrl) {
        triggerMailto(ctx.mailtoUrl);
      }

      const openedLabel = ctx.isOrderTemplate ? 'Order Template Opened!' : 'Form Opened!';
      const openWindow = () => {
        chrome.runtime.sendMessage(
          { type: 'openFormOnOppositeDisplay', url: ctx.openUrl },
          response => {
            if (chrome.runtime.lastError || !response || !response.ok) {
              setLabel('Error opening');
              setTimeout(() => {
                btn.disabled = false;
                setLabel('Export');
                btn.style.opacity = '1';
              }, 2000);
              return;
            }

            // Popup opened. If court PDFs are still downloading, keep the button
            // greyed out until they finish; then show confirmation and reset.
            setLabel('Downloading…');
            downloadsDone.then(() => {
              setLabel(openedLabel);
              btn.style.opacity = '1';
              setTimeout(() => {
                btn.disabled = false;
                setLabel('Export');
              }, 2000);
            });
          }
        );
      };

      // For the Order Template popup, auto-detect the Movant and wait until the
      // parsed fields are stored before opening the popup window. OSC cases open
      // the real form.
      if (ctx.isOrderTemplate) {
        computeMovant(ctx.data.labeled.motionType, result.partiesRoot, ctx.data.labeled.hearingDate).then(movant => {
          if (movant) ctx.data.labeled.movant = movant;
          storeOrderTemplateData(ctx.data.labeled).then(openWindow);
        });
      } else {
        openWindow();
      }
    } catch (err) {
      console.error('[LACourt] fill button error:', err);
      setLabel('Error');
      setTimeout(() => {
        btn.disabled = false;
        setLabel('Export');
        btn.style.opacity = '1';
      }, 2000);
    }
  });

  document.body.appendChild(btn);
}

/* ------------------------------------------------------------------ */
/* Floating "Documents" button (left of Export)                        */
/* ------------------------------------------------------------------ */

const MAX_DOCS_TO_OPEN = 60; // safety cap on how many tabs to open at once

function renderDocumentsButton() {
  if (document.getElementById('__lacourt_docs_btn__')) return;
  const caseReady = document.querySelector('a[href*="/ecourt/ecms/case"]');
  if (!caseReady) return;

  const btn = document.createElement('button');
  btn.id = '__lacourt_docs_btn__';
  btn.type = 'button';
  const setDocLabel = (text) => {
    btn.innerHTML = '<span class="lac-btn-icon" style="vertical-align:middle">📂</span>' +
      '<span class="lac-btn-text" style="vertical-align:middle">' + text + '</span>';
  };
  setDocLabel('Documents');
  Object.assign(btn.style, {
    position: 'fixed',
    top: '0px',
    right: '150px',
    zIndex: '999998',
    padding: '6px 16px',
    background: '#1a5d3a',
    color: 'white',
    border: 'none',
    borderRadius: '0 0 6px 6px',
    fontFamily: 'Georgia, serif',
    fontSize: '13px',
    fontWeight: 'bold',
    cursor: 'pointer',
    boxShadow: '0 2px 8px rgba(0,0,0,0.25)',
    transition: 'background 0.15s, opacity 0.2s',
  });
  btn.addEventListener('mouseover', () => { btn.style.background = '#248250'; });
  btn.addEventListener('mouseout',  () => { btn.style.background = '#1a5d3a'; });

  const reset = (text, ms) => {
    setDocLabel(text);
    setTimeout(() => { btn.disabled = false; setDocLabel('Documents'); btn.style.opacity = '1'; }, ms || 2500);
  };

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    setDocLabel('Finding…');
    btn.style.opacity = '0.7';
    try {
      // Use the result prefetched on page load (instant if ready). If it came
      // back empty — e.g. the prefetch ran before the page was ready — drop the
      // cache and recompute once so a real result isn't missed.
      let res = await getRelevantDocumentsCached();
      let opened = (res.relevant || []).filter(d => d.openUrl);
      if (!opened.length) {
        __relevantDocsPromise = null;
        res = await getRelevantDocuments();
        opened = (res.relevant || []).filter(d => d.openUrl);
      }
      if (!opened.length) { reset('None found'); return; }
      let capped = false;
      if (opened.length > MAX_DOCS_TO_OPEN) { capped = true; opened = opened.slice(0, MAX_DOCS_TO_OPEN); }

      // The background decides whether to open or close: if every relevant
      // document is already open it closes them all (no download); otherwise it
      // opens just the ones not yet open. It returns which docIds it opened so we
      // only record those for debug tracking (nothing is recorded on a close).
      const docsPayload = opened.map(d => ({ docId: d.docId, name: d.name, openUrl: d.openUrl }));
      chrome.runtime.sendMessage({ type: 'toggleDocsBackground', docs: docsPayload }, resp => {
        void chrome.runtime.lastError;
        const r = resp || {};
        if (r.action === 'closed') {
          // Un-mark these docs' green checkmarks (pdf-focus/bridge.js) — but only
          // the ones the button itself opened; a doc the user opened by hand stays
          // checked even though its tab just closed.
          try {
            window.dispatchEvent(new CustomEvent('LACOURT_DOCS_BUTTON_CLOSED', {
              detail: { docIds: opened.map(d => String(d.docId)) },
            }));
          } catch (_) {}
          reset('Closed ' + (r.count || 0)); return;
        }

        // Debug tracking: record only the documents actually opened.
        const openedIds = new Set((r.openedDocIds || []).map(String));
        const recorded = opened.filter(d => openedIds.has(String(d.docId)));
        if (recorded.length) {
          try {
            chrome.runtime.sendMessage({
              type: 'recordOpenedDocs',
              source: 'button',
              caseNumber: parseCaseNumber(),
              docs: recorded.map(d => ({ docId: d.docId, name: d.name })),
            }, () => void chrome.runtime.lastError);
          } catch (_) {}
        }
        if (r.action === 'opened') {
          // Mark every relevant doc's green checkmark, not just the ones newly
          // opened this click — the rest were already open (button or manual).
          try {
            window.dispatchEvent(new CustomEvent('LACOURT_DOCS_BUTTON_OPENED', {
              detail: { docIds: opened.map(d => String(d.docId)) },
            }));
          } catch (_) {}
        }
        reset('Opened ' + (r.count || 0) + (capped ? '+' : ''));
      });
    } catch (err) {
      console.error('[LACourt] documents button error:', err);
      reset('Error');
    }
  });

  document.body.appendChild(btn);
}

/* ------------------------------------------------------------------ */
/* Floating "Deadlines" button (left of Documents)                     */
/* ------------------------------------------------------------------ */
//
// Opens the in-extension Motion Deadline Calculator in its own window (on the
// opposite display, like Export). Uses the same hearing detection as Export and
// Documents to seed the calculator with the effective motion type + date.

// Reconsideration (§ 1008) and new trial / JNOV / vacate judgment (§§ 659, 663a)
// run from service of the notice of entry, not the upcoming hearing. For those
// we scan the case's Documents for the operative notice-of-entry filing so the
// calculator can seed the correct trigger date.
// The appeal-time trigger regexes and their selection live in lib/case-status.js
// (findAppealTimeTrigger) so the calculator hand-off and the briefing widget's
// fee-deadline computation can't drift apart.
const ENTRY_OF_JUDGMENT_RE = /\bjudgment\b/i;
function isTriggerBasedMotion(motionType) {
  // Same rule as classifyMotion: a motion filed "in support of" a new trial
  // motion is not itself keyed to the notice of entry.
  motionType = stripAncillaryMotionReference(motionType);
  return /reconsideration|renewed?\s+motion|\b1008\b|new\s+trial|\bjnov\b|judgment\s+notwithstanding|vacate\s+(the\s+)?judgment/i.test(motionType || '')
    // Fees (CRC 3.1702) and costs (3.1700) run off the same notice-of-entry
    // trigger, so they need the dates detected too.
    || /attorney'?s?\s+fees|\battorney\s+fee\b|\bfees\s+and\s+costs\b|\b(?:strike|tax|taxing)\s+(?:of\s+)?costs\b|memorandum\s+of\s+costs/i.test(motionType || '');
}
// The entry-of-judgment filing itself (for the § 659 180-day outer limit), as
// opposed to notices/proposed/supporting papers that merely mention "judgment".
function isEntryOfJudgmentDoc(name) {
  return ENTRY_OF_JUDGMENT_RE.test(name) &&
    !/notice|proposed|request|application|memorandum|points|declaration|stipulat|objection|opposition|\breply\b|\bmotion\b|abstract|assignment|renewal/i.test(name);
}
// Latest matching doc filed on or before the hearing (the challenged order/
// judgment predates the motion); falls back to the latest overall.
function latestDocOnOrBefore(matches, cutoff) {
  let pool = cutoff ? matches.filter(d => d.when <= cutoff) : matches;
  if (!pool.length) pool = matches;
  if (!pool.length) return null;
  pool.sort((a, b) => b.when - a.when);
  return pool[0];
}
// One Documents fetch that finds both the notice-of-entry filing (the § 1008 /
// § 659 15-day trigger) and the entry-of-judgment filing (the § 659 180-day
// outer limit).
async function detectTriggerDates(hearingDateStr) {
  const out = { noticeOfEntryDate: '', noticeOfEntryDoc: '', entryOfJudgmentDate: '', entryOfJudgmentDoc: '',
                memoServedDate: '', memoDoc: '', noticeOfEntryUnverified: false };
  try {
    const docsUrl = getDocumentsUrl();
    if (!docsUrl) return out;
    const docs = await fetchAllDocuments(docsUrl);
    if (!docs || !docs.length) return out;
    const cutoff = hearingDateStr ? parseHearingDateTime(hearingDateStr) : null;
    // Shared with the briefing widget's fee-deadline computation, so the two
    // can't disagree about which paper started the clock.
    const trig = findAppealTimeTrigger(docs, cutoff);
    if (trig) {
      out.noticeOfEntryDate = trig.doc.dateStr;
      out.noticeOfEntryDoc = trig.doc.name;
      out.noticeOfEntryUnverified = trig.unverified;
    }
    const eoj = latestDocOnOrBefore(docs.filter(d => d.name && d.when && isEntryOfJudgmentDoc(d.name)), cutoff);
    if (eoj) { out.entryOfJudgmentDate = eoj.dateStr; out.entryOfJudgmentDoc = eoj.name; }
    // The memorandum of costs a strike/tax motion attacks — the earliest one,
    // since that is the memo whose service starts the 15 days.
    const memos = docs.filter(d => d.name && d.when && COSTS_MEMO_DOC_RE.test(d.name)
      && (!cutoff || d.when <= cutoff)).sort((a, b) => a.when - b.when);
    if (memos[0]) { out.memoServedDate = memos[0].dateStr; out.memoDoc = memos[0].name; }
    return out;
  } catch (_) { return out; }
}

// The Deadlines hand-off, built once and reused. Detecting the statutory
// trigger dates means fetching the case Documents, which took a beat every time
// the button was pressed; building it when the page settles makes the press
// instant. Never stamps createdAt — the caller does, at click time.
let __deadlinePayloadPromise = null;
async function buildDeadlinePayload() {
  const hearing = await resolveEffectiveHearing(document);
  const motionType = (hearing && hearing.motionType) || '';
  // Trigger-based motions (fees, costs, new trial, reconsideration, vacate) run
  // off a paper on the docket rather than the hearing date. Skipped for ordinary
  // motions, which need no document fetch at all.
  let trig = { noticeOfEntryDate: '', noticeOfEntryDoc: '', noticeOfEntryUnverified: false,
               entryOfJudgmentDate: '', entryOfJudgmentDoc: '', memoServedDate: '', memoDoc: '' };
  if (isTriggerBasedMotion(motionType)) {
    trig = await detectTriggerDates(hearing && hearing.hearingDate);
  }
  return {
    motionType,
    hearingDate: (hearing && hearing.hearingDate) || '',
    hearingType: (hearing && hearing.hearingType) || '',
    caseNumber: parseCaseNumber() || '',
    unlawfulDetainer: pageIsUnlawfulDetainer(),
    noticeOfEntryDate: trig.noticeOfEntryDate,
    noticeOfEntryDoc: trig.noticeOfEntryDoc,
    noticeOfEntryUnverified: trig.noticeOfEntryUnverified,
    entryOfJudgmentDate: trig.entryOfJudgmentDate,
    entryOfJudgmentDoc: trig.entryOfJudgmentDoc,
    memoServedDate: trig.memoServedDate,
    memoDoc: trig.memoDoc,
  };
}
function getDeadlinePayloadCached() {
  if (!__deadlinePayloadPromise) {
    __deadlinePayloadPromise = buildDeadlinePayload();
    // Don't cache a failure: a transient fetch error would otherwise leave the
    // button permanently handing over an empty payload.
    __deadlinePayloadPromise.catch(() => { __deadlinePayloadPromise = null; });
  }
  return __deadlinePayloadPromise;
}

function renderDeadlineButton() {
  if (document.getElementById('__lacourt_deadline_btn__')) return;
  const caseReady = document.querySelector('a[href*="/ecourt/ecms/case"]');
  if (!caseReady) return;

  const btn = document.createElement('button');
  btn.id = '__lacourt_deadline_btn__';
  btn.type = 'button';
  const setDlLabel = (text) => {
    btn.innerHTML = '<span class="lac-btn-icon" style="vertical-align:middle">📅</span>' +
      '<span class="lac-btn-text" style="vertical-align:middle">' + text + '</span>';
  };
  setDlLabel('Deadlines');
  Object.assign(btn.style, {
    position: 'fixed',
    top: '0px',
    right: '290px',
    zIndex: '999998',
    padding: '6px 16px',
    background: '#0a6e6e',
    color: 'white',
    border: 'none',
    borderRadius: '0 0 6px 6px',
    fontFamily: 'Georgia, serif',
    fontSize: '13px',
    fontWeight: 'bold',
    cursor: 'pointer',
    boxShadow: '0 2px 8px rgba(0,0,0,0.25)',
    transition: 'background 0.15s, opacity 0.2s',
  });
  btn.addEventListener('mouseover', () => { btn.style.background = '#0d8f8f'; });
  btn.addEventListener('mouseout',  () => { btn.style.background = '#0a6e6e'; });

  const reset = (text, ms) => {
    setDlLabel(text);
    setTimeout(() => { btn.disabled = false; setDlLabel('Deadlines'); btn.style.opacity = '1'; }, ms || 2000);
  };

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    setDlLabel('Opening…');
    btn.style.opacity = '0.7';
    try {
      // Prefetched on page load, so this is normally already resolved. The
      // timestamp is stamped HERE rather than in the payload: the calculator
      // drops hand-offs older than ten minutes, and one prepared when the page
      // loaded could easily be older than that by the time it is used.
      const payload = Object.assign({}, await getDeadlinePayloadCached(), { createdAt: Date.now() });
      await new Promise(res => {
        try {
          chrome.storage.local.set({ deadlineCalcData: payload }, () => { void chrome.runtime.lastError; res(); });
        } catch (_) { res(); }
      });
      const url = chrome.runtime.getURL('deadline-calculator/deadline-calculator.html');
      chrome.runtime.sendMessage({ type: 'openFormOnOppositeDisplay', url }, () => {
        void chrome.runtime.lastError;
        reset('Opened');
      });
    } catch (err) {
      console.error('[LACourt] deadline button error:', err);
      reset('Error');
    }
  });

  document.body.appendChild(btn);
}

/* ------------------------------------------------------------------ */
/* Floating "Fees" button (default-judgment pages only, left of Deadlines) */
/* Opens the LASC Rule 3.214 attorney-fee calculator.                   */
/* ------------------------------------------------------------------ */

// Shown only when the effective hearing is an OSC Re: Failure to Prosecute
// Default Judgment (resolved by the inline deadline widget).
function isDefaultJudgmentPage() {
  return slotsHaveOsc();
}

function renderDefaultJudgmentFeesButton() {
  if (document.getElementById('__lacourt_djfees_btn__')) return;
  if (!isDefaultJudgmentPage()) return;
  const caseReady = document.querySelector('a[href*="/ecourt/ecms/case"]');
  if (!caseReady) return;

  const btn = document.createElement('button');
  btn.id = '__lacourt_djfees_btn__';
  btn.type = 'button';
  const setLabel = (text) => {
    btn.innerHTML = '<span class="lac-btn-icon" style="vertical-align:middle">🧮</span>' +
      '<span class="lac-btn-text" style="vertical-align:middle">' + text + '</span>';
  };
  setLabel('Fees');
  Object.assign(btn.style, {
    position: 'fixed',
    top: '0px',
    right: '430px',
    zIndex: '999998',
    padding: '6px 16px',
    background: '#0a6e6e',
    color: 'white',
    border: 'none',
    borderRadius: '0 0 6px 6px',
    fontFamily: 'Georgia, serif',
    fontSize: '13px',
    fontWeight: 'bold',
    cursor: 'pointer',
    boxShadow: '0 2px 8px rgba(0,0,0,0.25)',
    transition: 'background 0.15s, opacity 0.2s',
  });
  btn.title = 'Default judgment attorney fees (LASC Local Rule 3.214)';
  btn.addEventListener('mouseover', () => { btn.style.background = '#0d8f8f'; });
  btn.addEventListener('mouseout',  () => { btn.style.background = '#0a6e6e'; });

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    setLabel('Opening…');
    btn.style.opacity = '0.7';
    const reset = (text) => {
      setLabel(text);
      setTimeout(() => { btn.disabled = false; setLabel('Fees'); btn.style.opacity = '1'; }, 1500);
    };
    try {
      const payload = { caseNumber: parseCaseNumber() || '', createdAt: Date.now() };
      await new Promise(res => {
        try { chrome.storage.local.set({ djFeesData: payload }, () => { void chrome.runtime.lastError; res(); }); }
        catch (_) { res(); }
      });
      const url = chrome.runtime.getURL('default-judgment-fees/default-judgment-fees.html');
      chrome.runtime.sendMessage({ type: 'openFormOnOppositeDisplay', url }, () => {
        void chrome.runtime.lastError;
        reset('Opened');
      });
    } catch (err) {
      console.error('[LACourt] DJ fees button error:', err);
      reset('Error');
    }
  });

  document.body.appendChild(btn);
  try { scheduleButtonCollapse(); } catch (_) {}
}

/* ------------------------------------------------------------------ */
/* Collapse the floating buttons to icons when zoomed in enough that     */
/* their expanded labels would sit over e-court text.                    */
/* ------------------------------------------------------------------ */

const BTN_EDGE = 16;  // Export button's right offset (px from viewport edge)
const BTN_GAP = 8;    // gap between the two floating buttons

// Inject the icon/text collapse rules once. Hiding .lac-btn-text (and tightening
// horizontal padding) leaves just the icon; scoped to our two button IDs.
function ensureButtonStyles() {
  if (document.getElementById('__lacourt_btn_styles__')) return;
  const st = document.createElement('style');
  st.id = '__lacourt_btn_styles__';
  st.textContent =
    '#__lacourt_fill_btn__ .lac-btn-text,#__lacourt_docs_btn__ .lac-btn-text,' +
    '#__lacourt_deadline_btn__ .lac-btn-text,#__lacourt_djfees_btn__ .lac-btn-text{margin-left:6px}' +
    '#__lacourt_fill_btn__[data-collapsed="1"] .lac-btn-text,' +
    '#__lacourt_docs_btn__[data-collapsed="1"] .lac-btn-text,' +
    '#__lacourt_deadline_btn__[data-collapsed="1"] .lac-btn-text,' +
    '#__lacourt_djfees_btn__[data-collapsed="1"] .lac-btn-text{display:none}' +
    '#__lacourt_fill_btn__[data-collapsed="1"],#__lacourt_docs_btn__[data-collapsed="1"],' +
    '#__lacourt_deadline_btn__[data-collapsed="1"],#__lacourt_djfees_btn__[data-collapsed="1"]' +
    '{padding-left:9px!important;padding-right:9px!important}';
  (document.head || document.documentElement).appendChild(st);
}

// True if viewport point (x,y) lands on rendered text of some non-button element.
function lacPointOverText(x, y, btns) {
  const el = document.elementFromPoint(x, y);
  if (!el || el === document.body || el === document.documentElement) return false;
  if (btns.indexOf(el) !== -1) return false;
  for (let i = 0; i < el.childNodes.length; i++) {
    const node = el.childNodes[i];
    if (node.nodeType === 3 && node.nodeValue && node.nodeValue.trim()) {
      const r = document.createRange();
      r.selectNodeContents(node);
      const rects = r.getClientRects();
      for (let j = 0; j < rects.length; j++) {
        const rc = rects[j];
        if (x >= rc.left && x <= rc.right && y >= rc.top && y <= rc.bottom) return true;
      }
    }
  }
  return false;
}

// Scans the buttons' expanded footprint for any e-court text beneath it.
function lacRegionHasText(left, right, top, bottom, btns) {
  btns.forEach(b => { b.style.pointerEvents = 'none'; });
  const ys = [top, (top + bottom) / 2, bottom];
  let hit = false;
  for (let yi = 0; yi < ys.length && !hit; yi++) {
    for (let x = right; x >= left; x -= 6) {
      if (lacPointOverText(x, ys[yi], btns)) { hit = true; break; }
    }
  }
  btns.forEach(b => { b.style.pointerEvents = ''; });
  return hit;
}

// Find the site's fixed blue header bar by probing what is actually rendered at
// the top of the viewport (elementsFromPoint at a few x positions), then taking
// the TALLEST fixed/sticky, near-full-width element pinned to the top. Mirrors
// agenda/content.js findTopBar() so the case-page buttons match the agenda
// Copy All button's full-bar sizing.
let __caseTopBarEl = null;
const __LAC_BTN_IDS = ['__lacourt_fill_btn__', '__lacourt_docs_btn__', '__lacourt_deadline_btn__', '__lacourt_djfees_btn__'];
function findCaseTopBar() {
  try {
    if (__caseTopBarEl && document.contains(__caseTopBarEl)) {
      const r = __caseTopBarEl.getBoundingClientRect();
      if (r.top <= 2 && r.height >= 20 && r.height <= 140) return __caseTopBarEl;
      __caseTopBarEl = null;
    }
    const w = window.innerWidth;
    let best = null, bestH = 0;
    for (const x of [w * 0.3, w * 0.5, w * 0.7]) {
      for (const el of document.elementsFromPoint(x, 8)) {
        if (el === document.documentElement || el === document.body) continue;
        if (__LAC_BTN_IDS.indexOf(el.id) !== -1) continue;
        const pos = getComputedStyle(el).position;
        if (pos !== 'fixed' && pos !== 'sticky') continue;
        const r = el.getBoundingClientRect();
        if (r.top > 2 || r.height < 20 || r.height > 140) continue;
        if (r.width < w * 0.6) continue;
        if (r.height > bestH) { best = el; bestH = r.height; }
      }
    }
    __caseTopBarEl = best;
    return best;
  } catch (_) { return null; }
}

// Size the floating buttons to the full height of the top bar (flush, no bottom
// radius) so they fill the blue header like the agenda Copy All button. Falls
// back to the compact default (rounded tab hanging from top:0) when no bar is
// found. Uses !important so site CSS can't shrink them.
let __caseBarLoggedEl = null;
// { top, h } of the last successful bar measurement. Seeded from sessionStorage
// (per-tab, survives the full page reload that a case sub-tab switch triggers)
// so the buttons paint at the right size immediately instead of flashing to the
// compact default and only correcting once the live probe settles.
const __CASE_BAR_SS_KEY = 'lacourt.caseBarSize';
let __caseBarLastSize = (() => {
  try { const v = JSON.parse(sessionStorage.getItem(__CASE_BAR_SS_KEY) || 'null');
    return (v && typeof v.h === 'number' && typeof v.top === 'number') ? v : null; } catch (_) { return null; }
})();
function sizeButtonsToBar(btns) {
  // While the tab is hidden, getBoundingClientRect()/elementsFromPoint report
  // stale or zero layout, so a probe here would wrongly conclude "no bar" and
  // revert the buttons. Leave the current sizing untouched; a visibilitychange
  // listener re-sizes once the tab is visible again.
  if (document.hidden) return;

  const bar = findCaseTopBar();
  let size = null;
  if (bar) {
    const r = bar.getBoundingClientRect();
    size = { top: Math.max(0, Math.round(r.top)), h: Math.round(r.height) };
    if (!__caseBarLastSize || __caseBarLastSize.top !== size.top || __caseBarLastSize.h !== size.h) {
      __caseBarLastSize = size;
      try { sessionStorage.setItem(__CASE_BAR_SS_KEY, JSON.stringify(size)); } catch (_) {}
    }
  } else if (__caseBarLastSize) {
    // Transient probe failure (mid-render, layout not settled): keep the last
    // known good bar size rather than snapping back to the compact default.
    size = __caseBarLastSize;
  }
  btns.forEach(btn => {
    const set = (p, v) => { try { btn.style.setProperty(p, v, 'important'); } catch (_) { btn.style[p] = v; } };
    const clear = (p) => { try { btn.style.removeProperty(p); } catch (_) {} };
    if (size) {
      set('top', size.top + 'px');
      set('height', size.h + 'px');
      set('line-height', size.h + 'px');
      set('padding-top', '0');
      set('padding-bottom', '0');
      set('font-size', Math.max(13, Math.min(18, Math.round(size.h * 0.4))) + 'px');
      set('border-radius', '0 0 0 0');
      set('box-sizing', 'border-box');
      // No drop shadow when the button fills the header — it would fall onto the
      // slim grey bar directly below the blue header.
      set('box-shadow', 'none');
    } else {
      ['top', 'height', 'line-height', 'padding-top', 'padding-bottom', 'font-size', 'border-radius', 'box-sizing', 'box-shadow']
        .forEach(clear);
    }
  });
  if (bar && __caseBarLoggedEl !== bar) {
    __caseBarLoggedEl = bar;
    try { console.log('[LACourt] top bar:', bar.tagName + '.' + (bar.className || ''), 'height=' + Math.round(bar.getBoundingClientRect().height)); } catch (_) {}
  }
}

// Docks the buttons in a row from the right edge inward (rightmost first) and
// returns the left edge of the leftmost button.
function dockButtonsRow(btns) {
  const vw = document.documentElement.clientWidth;
  let rightPx = BTN_EDGE, leftEdge = vw - BTN_EDGE;
  for (const b of btns) {
    b.style.right = rightPx + 'px';
    const w = b.offsetWidth;
    leftEdge = Math.min(leftEdge, vw - rightPx - w);
    rightPx += w + BTN_GAP;
  }
  return leftEdge;
}

// Measures the expanded footprint (toggling attributes synchronously so nothing
// paints mid-measurement), then collapses all floating buttons to icons if that
// footprint would overlap e-court text. Keeps the buttons docked in a row
// (Export rightmost, then Documents, Deadlines, and — on default-judgment pages
// — DJ Fees) in whichever state.
function updateButtonCollapse() {
  const ids = ['__lacourt_fill_btn__', '__lacourt_docs_btn__', '__lacourt_deadline_btn__', '__lacourt_djfees_btn__'];
  const btns = ids.map(id => document.getElementById(id)).filter(Boolean);
  if (!btns.length) return;

  // Size to the blue header bar first so width/height measurements below reflect
  // the final rendered buttons.
  sizeButtonsToBar(btns);

  // Force expanded and dock across the row to measure the full footprint.
  btns.forEach(b => b.removeAttribute('data-collapsed'));
  const leftEdge = dockButtonsRow(btns);

  const vw = document.documentElement.clientWidth;
  let maxH = 0;
  btns.forEach(b => { if (b.offsetHeight > maxH) maxH = b.offsetHeight; });

  const overlap = lacRegionHasText(Math.max(0, leftEdge), vw - BTN_EDGE, 3, Math.max(3, maxH - 3), btns);

  if (overlap) {
    btns.forEach(b => b.setAttribute('data-collapsed', '1'));
    dockButtonsRow(btns); // re-dock at the narrower collapsed widths
  }
}

let lacCollapseTimer = null;
function scheduleButtonCollapse() {
  if (lacCollapseTimer) clearTimeout(lacCollapseTimer);
  lacCollapseTimer = setTimeout(() => { try { updateButtonCollapse(); } catch (_) {} }, 120);
}
window.addEventListener('resize', scheduleButtonCollapse);
// Returning to a background tab: layout is only live once the tab is visible
// again, so re-measure the bar and re-size the buttons (they may have been left
// at the wrong size if the page re-rendered while hidden).
document.addEventListener('visibilitychange', () => { if (!document.hidden) scheduleButtonCollapse(); });
window.addEventListener('focus', scheduleButtonCollapse);
window.addEventListener('pageshow', scheduleButtonCollapse);

// A cold load can render the buttons before the site paints its blue header:
// the probe finds no bar, sessionStorage carries no remembered size yet (nothing
// has measured it in this tab), and the buttons fall back to the compact tab.
// The listeners above only fire on resize/refocus/reload, so that wrong size
// then sticks for the life of the page — the symptom being buttons that don't
// fill the header until you reload a second time. Retry on a short ladder after
// first paint, stopping the moment the bar turns up (so a warm load, which wins
// the race on the first rung, costs one extra probe).
const BAR_RETRY_MS = [150, 400, 1000, 2500];
let __barRetryStarted = false;
function retryBarSizingUntilFound() {
  if (__barRetryStarted) return;
  __barRetryStarted = true;
  let i = 0;
  const tick = () => {
    // A hidden tab reports stale layout, so a miss there proves nothing; the
    // visibilitychange listener re-sizes when the tab comes back.
    if (!document.hidden && findCaseTopBar()) {
      try { updateButtonCollapse(); } catch (_) {}
      return;
    }
    if (i < BAR_RETRY_MS.length) setTimeout(tick, BAR_RETRY_MS[i++]);
  };
  setTimeout(tick, BAR_RETRY_MS[i++]);
}

/* ------------------------------------------------------------------ */
/* Inline Opposition / Reply (and Motion) deadlines on the "Next" header */
/* ------------------------------------------------------------------ */
//
// For a briefable motion on calendar (a "Hearing on <motion>"), compute the
// § 1005 / § 437c briefing deadlines from the hearing date and show them inline
// next to the Next-event indicator. The moving-papers deadline assumes
// electronic service. Each paper is checked against the case Documents: if the
// paper was filed on or before its due date it shows GREEN (filed on time); if
// its due date has passed with no timely filing it shows RED (overdue); if it
// isn't due yet it shows in the neutral colour. The Motion additionally shows
// YELLOW when it missed the electronic-service deadline but would still be timely
// under personal service (no notice extension) — a cue to check the proof of
// service.



// Finds the element that visibly shows the "Next: <date> ... Hearing on ..."
// indicator, to anchor the inline deadlines. Tries span[title] first (what the
// parsers use), then falls back to scanning for the smallest visible element
// whose own text starts with "Next" and contains a date.
let __dlDumped = false;
function findNextHeaderSpan() {
  const hasNextDate = s => /\bnext\b/i.test(s) && /\d{1,2}\/\d{1,2}\/\d{4}/.test(s);

  const titled = document.querySelectorAll('span[title]');
  for (const span of titled) {
    const hay = ((span.getAttribute('title') || '') + ' ' + (span.textContent || '')).replace(/\s+/g, ' ');
    if (hasNextDate(hay)) return span;
  }

  // Fallback: any smallish visible element whose text reads "Next: <date> …".
  const all = document.querySelectorAll('span, div, td, p, a, b, strong, label, li');
  let best = null;
  for (const el of all) {
    const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
    if (t.length > 300 || !/^next\b/i.test(t) || !/\d{1,2}\/\d{1,2}\/\d{4}/.test(t)) continue;
    if (!el.getClientRects || !el.getClientRects().length) continue; // must be visible
    // Prefer the deepest (smallest) matching element.
    if (!best || (el.textContent || '').length < (best.textContent || '').length) best = el;
  }
  if (best) return best;

  if (!__dlDumped) {
    __dlDumped = true;
    const titles = [].slice.call(titled).map(s => (s.getAttribute('title') || '').slice(0, 80)).filter(Boolean);
    dlLog('no header found. span[title] count=', titled.length, 'titles=', titles.slice(0, 12));
  }
  return null;
}

/* ------------------------------------------------------------------ */
/* Which hearing (or hearings) the page is working up                   */
/* ------------------------------------------------------------------ */
//
// eCourt's header names exactly ONE upcoming event, and a case routinely has
// several hearings we work up — on different days, or two set for the same
// morning. So every workable hearing is bundled one group per hearing DATE:
// ‹ › arrows on the header step between days, and each motion set for the
// selected day gets its own "Next:" line, its own briefing deadlines, its
// documents opened alongside the others', and its motion type on the Export
// form.

const NEXT_HEADER_RE = /Next:?\s*(\d{1,2}\/\d{1,2}\/\d{4})\s+(\d{1,2}:\d{2}\s*(?:AM|PM))\s+(.+?)\s*$/i;

// The page's own "Next:" line, split into date / time / event name.
function nextHeaderMatch(root) {
  root = root || document;
  for (const span of root.querySelectorAll('span[title]')) {
    const title = (span.getAttribute('title') || '').trim();
    let m = title && title.match(NEXT_HEADER_RE);
    if (m) return m;
    const text = (span.textContent || '').trim().replace(/\s+/g, ' ');
    m = text && text.match(NEXT_HEADER_RE);
    if (m) return m;
  }
  return null;
}

// The Next event, in the shape groupWorkableHearings wants. Read ONCE and kept:
// once the arrows have rewritten the header to a later hearing, the live DOM no
// longer says what eCourt itself put there, and the original is what tells us
// whether the page's own line is a hearing we work up (and what to restore).
let __nativeSnapshot = null;
function nativeSnapshot() {
  if (__nativeSnapshot) return __nativeSnapshot;
  const hearingType = parseHearingType(document);
  if (!hearingType) return null; // header hasn't rendered yet
  const m = nextHeaderMatch(document);
  __nativeSnapshot = {
    motionType: parseMotionType(document),
    hearingType,
    hearingDate: parseHearingDate(document),
    timeText: m ? m[2].replace(/\s+/g, ' ').toUpperCase() : '',
    // "… in Department 73" is where the event is heard, not part of its name.
    raw: m ? stripEventId(m[3]).replace(/\s+in\s+Department\b.*$/i, '').trim() : '',
  };
  return __nativeSnapshot;
}

// The header is rendered by site scripts a beat after the rest of the page, and
// the hearing list is worthless without it, so wait rather than resolve early.
function waitForNativeSnapshot(maxMs) {
  const deadline = Date.now() + (maxMs || 10000);
  return new Promise(resolve => {
    const tick = () => {
      const s = nativeSnapshot();
      if (s || Date.now() > deadline) { resolve(s); return; }
      setTimeout(tick, 250);
    };
    tick();
  });
}

let __hearingGroups = null;
let __hearingGroupsPromise = null;
let __selGroupIdx = 0;

function getHearingGroupsCached() {
  if (__hearingGroupsPromise) return __hearingGroupsPromise;
  __hearingGroupsPromise = (async () => {
    await loadExcludedTerms();
    const native = await waitForNativeSnapshot();
    let hearings = [];
    try { hearings = await getFutureHearingsCached(); } catch (_) { hearings = []; }
    return groupWorkableHearings(native, hearings);
  })();
  // Don't cache a failure: a transient Hearings-tab error would otherwise pin
  // the page to an empty hearing list for the rest of its life.
  __hearingGroupsPromise.catch(() => { __hearingGroupsPromise = null; });
  return __hearingGroupsPromise;
}

// The selected day survives a sub-tab reload (each one is a full page load), so
// stepping to a later hearing and then opening Documents or Parties keeps it.
const __HSEL_PREFIX = 'lacourt.hsel.';
function hselKey() {
  const cn = parseCaseNumber(document) || '';
  return cn ? __HSEL_PREFIX + cn : null;
}
function readSelDate() {
  const k = hselKey(); if (!k) return '';
  try { return sessionStorage.getItem(k) || ''; } catch (_) { return ''; }
}
function writeSelDate(d) {
  const k = hselKey(); if (!k) return;
  try { if (d) sessionStorage.setItem(k, d); else sessionStorage.removeItem(k); } catch (_) {}
}

function selectedGroup() {
  const groups = __hearingGroups || [];
  if (!groups.length) return null;
  if (__selGroupIdx < 0 || __selGroupIdx >= groups.length) __selGroupIdx = 0;
  return groups[__selGroupIdx];
}

// Resolves the hearing groups once and restores the remembered day.
let __groupsReadyPromise = null;
function getSelectedHearingGroup() {
  // Answer with the day selected RIGHT NOW — the arrows move it after this has
  // already resolved once, and every caller wants the current one.
  if (__hearingGroups) return Promise.resolve(selectedGroup());
  if (!__groupsReadyPromise) {
    __groupsReadyPromise = getHearingGroupsCached().then(groups => {
      __hearingGroups = groups;
      const want = readSelDate();
      const i = want ? groups.findIndex(g => g.date === want) : -1;
      __selGroupIdx = i >= 0 ? i : 0;
    });
    __groupsReadyPromise.catch(() => { __groupsReadyPromise = null; });
  }
  return __groupsReadyPromise.then(() => selectedGroup());
}

// Every hearing on the selected day as an effective-hearing record. Index 0 is
// the one the page would have worked up on its own.
function selectedHearingEffs() {
  const g = selectedGroup();
  if (!g) return [];
  // The ▸ "these figures are for…" prefix belongs only on the page's own header
  // line, and only while that line names a DIFFERENT event than the figures
  // describe (a CMC, a trial, another OSC). Lines we render name their own.
  const snap = __nativeSnapshot;
  const nativeShows = !!(snap && isWorkableHearing(snap.hearingType));
  return g.items.map((it, i) => ({
    motionType: it.motionType,
    hearingType: it.hearingType,
    hearingDate: it.hearingDate,
    timeText: it.timeText || '',
    raw: it.raw || '',
    native: !!it.native,
    lookedAhead: i === 0 && !nativeShows,
    ud: pageIsUnlawfulDetainer(),
  }));
}

// Whether this case's type designation reads unlawful detainer — an MSJ then
// runs on CCP § 1170.7 / CRC 3.1351 rather than § 437c. A positive answer is
// cached; a negative one is re-checked (the header can render after our first
// look), and each check is a bounded scan of the case-header blocks.
let __udPageCached = false;
function pageIsUnlawfulDetainer() {
  if (__udPageCached) return true;
  try { if (isUnlawfulDetainerCase(document)) { __udPageCached = true; dlLog('unlawful detainer case detected'); } } catch (_) {}
  return __udPageCached;
}

/* ------------------------------------------------------------------ */
/* Inline deadlines — one set per hearing on the selected day           */
/* ------------------------------------------------------------------ */

// One slot per hearing: { eff, computed, filed, osc, fetchStarted }. Slot 0
// rides the page's own header line; the rest get header lines of their own
// directly beneath it.
let __dlSlots = [];

function slotKey(eff) {
  return (eff.hearingDate || '') + '|' + (eff.hearingType || '');
}
function makeSlot(eff) {
  return { eff, computed: computeDueDatesFor(eff), filed: null, osc: null, fetchStarted: false };
}
// Same test the single-hearing widget always used: an OSC or a briefable motion
// paints; a motion whose deadlines aren't hearing-based paints only if it still
// has to report whether its moving papers arrived.
function paintableSlot(s) {
  const c = s && s.computed;
  return !!c && !(c.skip && !c.motionOnly);
}
function slotsHaveOsc() {
  return __dlSlots.some(s => s.computed && s.computed.osc);
}

let __dlNoMotionLogged = 0;

// This case, as the shared engine reaches its Documents / Hearings / Parties.
// Documents and Hearings reuse the page's own memoized fetches, so the widget
// and the Documents button share one round trip; Parties reads the live DOM when
// we're already on that tab.
let __pageCaseCtx = null;
function pageCaseCtx() {
  if (__pageCaseCtx) return __pageCaseCtx;
  __pageCaseCtx = makeCaseCtx({
    docs: getAllDocumentsCached,
    hearings: getFutureHearingsCached,
    parties: async () => {
      if (document.querySelector('a[title="UPDATE PARTY"]')) return document;
      const url = getPartiesUrl();
      const doc = url ? await fetchCaseDoc(url) : null;
      return doc || emptyDoc();
    },
  });
  return __pageCaseCtx;
}

// The slot the answer belongs to RIGHT NOW. applySelectedHearings rebuilds
// __dlSlots while an early fetch can still be in flight, copying fetchStarted
// into the fresh slot — so a completion that writes to its captured slot
// object lands on a discarded one, the widget never recolours, and the cache
// stores no filing status (the "colours only appear after switching tabs"
// bug). Re-locate by key at completion instead.
function liveSlotFor(slot) {
  const key = slotKey(slot.eff);
  return __dlSlots.find(s => slotKey(s.eff) === key) || slot;
}

// Fetch the case Documents once and recolour by whether each paper was filed on
// time. Best-effort — the dates are already shown regardless.
async function fetchSlotFilings(slot) {
  if (!slot || slot.fetchStarted || !slot.computed) return;
  if (slot.computed.osc) return fetchSlotOsc(slot);
  // motionOnly hearings have no deadlines, but still need the Documents fetch:
  // whether the moving papers are on file is the whole of what they report.
  if (slot.computed.skip && !slot.computed.motionOnly) return;
  slot.fetchStarted = true;
  let filed = null;
  try {
    filed = await computeFiledStatus(pageCaseCtx(), slot.computed);
  } catch (e) { dlLog('filed-status fetch failed:', (e && e.message) || e); return; }
  const live = liveSlotFor(slot);
  live.filed = filed;
  live.fetchStarted = true;
  dlLog('filed status resolved:', slot.eff.hearingType, 'known=', !!(filed && filed.filedKnown));
  dlCacheWrite();        // persist so the next sub-tab load paints final colours
  // Recolour now, and once more a beat later — a repaint that lands while
  // eCourt is mid-re-render can be wiped before it is seen.
  try { injectNextDeadlines(); } catch (e) { dlLog('recolour failed:', (e && e.message) || e); }
  setTimeout(() => { try { injectNextDeadlines(); } catch (_) {} }, 500);
}

// Idempotent: injects the widget if missing, else refreshes its colours. Re-finds
// the header each time so it survives e-court's React re-renders.
function injectNextDeadlines() {
  if (!__dlSlots.length) return;
  const anchor = findNextHeaderAnchor();
  if (!anchor || !anchor.row.parentNode) return;
  paintNativeHeaderLine(anchor);
  renderHearingArrows(anchor);
  syncExtraHeaderRows(anchor);
  const rows = [anchor.row].concat(
    [].slice.call(extraRowHost(anchor).querySelectorAll('.' + EXTRA_ROW_CLASS)));
  __dlSlots.forEach((slot, i) => {
    const row = rows[i];
    if (!row) return;
    const label = row === anchor.row ? anchor.span : (row.querySelector('.' + NEXT_LABEL_CLASS) || row.firstElementChild);
    paintSlotWidget(label, slot);
  });
}

function paintSlotWidget(labelEl, slot) {
  if (!labelEl || !labelEl.parentNode) return;
  const host = labelEl.parentNode;
  let el = null;
  for (const n of host.children) {
    if (n.classList && n.classList.contains(DL_CLASS)) { el = n; break; }
  }
  if (!paintableSlot(slot)) { if (el) el.remove(); return; }
  const html = statusHtml(slot.computed, slot.filed, slot.osc);
  if (el) { if (el.innerHTML !== html) el.innerHTML = html; return; }
  el = document.createElement('span');
  el.className = DL_CLASS;
  // OSC status can be a longer sentence, so let it wrap; deadlines stay on one line.
  const ws = slot.computed.osc ? 'white-space:normal' : 'white-space:nowrap';
  el.setAttribute('style', 'margin-left:22px;font-weight:600;' + ws + ';font-family:inherit;display:inline-block;');
  el.innerHTML = html;
  // Sit after the arrows when they're there, so the hearing text and the control
  // that changes it stay together.
  let ref = labelEl;
  const nx = labelEl.nextElementSibling;
  if (nx && nx.classList && nx.classList.contains(NAV_CLASS)) ref = nx;
  host.insertBefore(el, ref.nextSibling);
  dlLog('injected deadlines next to header:', (labelEl.textContent || '').slice(0, 60));
}

/* ---- The header line, the green band, and the lines we add to it ---- */

const DL_CLASS = '__lacourt_next_dl__';
const NAV_CLASS = '__lacourt_hnav__';
const EXTRA_ROW_CLASS = '__lacourt_extra_next__';
const NEXT_LABEL_CLASS = '__lacourt_next_label__';

// The header line and the band it sits in. The band is the nearest ancestor that
// actually PAINTS a background — eCourt's green strip — so a line inserted into
// it picks up that background and pushes the rest of the page down, exactly as a
// second native hearing line would.
function findNextHeaderAnchor() {
  const span = findNextHeaderSpan();
  if (!span || !span.parentElement) return null;
  const painted = el => {
    const bg = (getComputedStyle(el).backgroundColor || '').replace(/\s+/g, '');
    return !!bg && bg !== 'transparent' && bg !== 'rgba(0,0,0,0)';
  };
  let row = span, band = null;
  for (let i = 0; i < 8 && row.parentElement; i++) {
    if (painted(row.parentElement)) { band = row.parentElement; break; }
    row = row.parentElement;
  }
  if (!band) { row = span; band = span.parentElement; }
  return { span, row, band };
}

// Index path from an ancestor to a descendant, so the same node can be found
// again inside a clone of that ancestor.
function childPath(root, node) {
  const path = [];
  let n = node;
  while (n && n !== root) {
    const p = n.parentNode;
    if (!p) return null;
    path.unshift([].indexOf.call(p.childNodes, n));
    n = p;
  }
  return n === root ? path : null;
}
function nodeAtPath(root, path) {
  let n = root;
  for (const i of path) { n = n && n.childNodes[i]; if (!n) return null; }
  return n;
}

// The line eCourt would itself have written for this hearing, had it been the
// Next event: "Next: 08/31/2026 8:30 AM Hearing on Motion to Strike Costs (6258)".
function hearingHeaderLabel(eff) {
  let name = stripEventId(eff.raw || eff.hearingType || '').trim();
  if (!name) name = 'Hearing';
  // The Hearings tab names a motion directly; the header prefixes it. Events
  // that aren't "Hearing on" anything (an OSC) are written as they stand.
  if (!/^hearing on\b/i.test(name) && !/^(?:order to show cause|osc)\b/i.test(name)) {
    name = 'Hearing on ' + name;
  }
  return 'Next: ' + (eff.hearingDate || '') + (eff.timeText ? ' ' + eff.timeText : '') + ' ' + name;
}

// What actually holds the hearing text: the header element itself when it is a
// plain text line, else the text node inside it that carries the date — so a
// header built out of markup ("<b>Next:</b> <span>08/31/2026 …</span>") is
// rewritten without disturbing that markup.
function nativeLineCarrier(span) {
  if (!span.children || !span.children.length) {
    return { get: () => span.textContent || '', set: v => { if (span.textContent !== v) span.textContent = v; } };
  }
  for (const n of span.childNodes) {
    if (n.nodeType === 3 && /\d{1,2}\/\d{1,2}\/\d{4}/.test(n.nodeValue || '')) {
      return { get: () => n.nodeValue || '', set: v => { if (n.nodeValue !== v) n.nodeValue = v; } };
    }
  }
  return null;
}

// The page's own line shows whichever hearing the arrows have selected, and is
// restored verbatim once the selection is back on the event eCourt named.
function paintNativeHeaderLine(anchor) {
  const snap = __nativeSnapshot;
  const slot = __dlSlots[0];
  if (!snap || !slot) return;
  // A line naming an event we DON'T work up (a CMC, a trial) is left as eCourt
  // wrote it — the widget beside it leads with a ▸ naming the hearing its
  // figures belong to, which is what that line has always done.
  if (!isWorkableHearing(snap.hearingType)) return;
  const span = anchor.span;
  const carrier = span.dataset ? nativeLineCarrier(span) : null;
  if (!carrier) return;
  if (span.dataset.lacNextOrig === undefined) {
    span.dataset.lacNextOrig = carrier.get();
    span.dataset.lacNextOrigTitle = span.getAttribute('title') || '';
  }
  const isNative = slot.eff.hearingDate === snap.hearingDate
    && (slot.eff.hearingType || '') === (snap.hearingType || '');
  let text = span.dataset.lacNextOrig;
  if (!isNative) {
    text = hearingHeaderLabel(slot.eff);
    // Don't repeat a "Next:" that lives in markup beside the text we're rewriting.
    if (!/^\s*next\b/i.test(span.dataset.lacNextOrig)) {
      const lead = (span.dataset.lacNextOrig.match(/^\s*/) || [''])[0] || ' ';
      text = text.replace(/^Next:\s*/i, lead);
    }
  }
  carrier.set(text);
  if (span.hasAttribute('title')) {
    const t = isNative ? span.dataset.lacNextOrigTitle : hearingHeaderLabel(slot.eff);
    if (t && span.getAttribute('title') !== t) span.setAttribute('title', t);
  }
}

const squash = el => ((el && el.textContent) || '').replace(/\s+/g, ' ').trim();

// Where added lines live. Normally they are siblings of the native line inside
// the green band, so the band grows around them. When the band is a table row
// they can't be — a <div> is not a legal sibling of a <td> — so they go inside
// the native line's own cell instead, which renders the same way.
function extraRowHost(anchor) {
  return /^(?:TR|TBODY|THEAD|TFOOT|TABLE)$/.test(anchor.band.tagName || '')
    ? anchor.row : anchor.band;
}

// One added header line per further hearing on the selected day, cloned from the
// native line so it carries the page's own font, colour and spacing.
function buildExtraHeaderRow(anchor, eff) {
  const wrap = document.createElement('div');
  wrap.className = EXTRA_ROW_CLASS;
  // Force its own line whether the band lays its children out as blocks, flex
  // items or grid cells.
  wrap.setAttribute('style', 'display:block;width:100%;flex:0 0 100%;grid-column:1/-1;');
  const text = hearingHeaderLabel(eff);

  const clone = anchor.row.cloneNode(true);
  // Drop anything of ours the clone inherited, and every id (they must stay unique).
  clone.querySelectorAll('.' + DL_CLASS + ', .' + NAV_CLASS + ', .' + EXTRA_ROW_CLASS)
    .forEach(n => n.remove());
  clone.removeAttribute('id');
  clone.querySelectorAll('[id]').forEach(n => n.removeAttribute('id'));

  // Clone the line only when the line really IS the hearing text. If the nearest
  // painted ancestor sits well above it, `row` is a container holding much more
  // than the "Next:" line, and duplicating it would duplicate the whole header —
  // so build a bare line and copy the native text's typography onto it instead.
  const path = childPath(anchor.row, anchor.span);
  const label = path ? nodeAtPath(clone, path) : null;
  const tight = squash(clone).length <= squash(anchor.span).length + 80;
  if (tight && label && label.nodeType === 1) {
    label.textContent = text;
    if (label.hasAttribute('title')) label.setAttribute('title', text);
    label.classList.add(NEXT_LABEL_CLASS);
    wrap.appendChild(clone);
    return wrap;
  }

  const line = document.createElement('span');
  line.className = NEXT_LABEL_CLASS;
  line.textContent = text;
  try {
    const cs = getComputedStyle(anchor.span);
    line.style.font = cs.font || '';
    if (!line.style.fontSize) {
      line.style.fontFamily = cs.fontFamily;
      line.style.fontSize = cs.fontSize;
      line.style.fontWeight = cs.fontWeight;
      line.style.fontStyle = cs.fontStyle;
    }
    line.style.color = cs.color;
    line.style.letterSpacing = cs.letterSpacing;
    // Line the added text up under the native one.
    const rs = getComputedStyle(anchor.row);
    wrap.style.paddingLeft = rs.paddingLeft;
    wrap.style.marginLeft = rs.marginLeft;
  } catch (_) {}
  wrap.appendChild(line);
  return wrap;
}

let __extraRowSig = '';
function syncExtraHeaderRows(anchor) {
  const extras = __dlSlots.slice(1);
  const sig = extras.map(s => slotKey(s.eff)).join('||');
  const host = extraRowHost(anchor);
  const present = host.querySelectorAll('.' + EXTRA_ROW_CLASS);
  if (sig === __extraRowSig && present.length === extras.length) return;
  present.forEach(n => n.remove());
  __extraRowSig = sig;
  let after = host === anchor.row ? null : anchor.row;
  // Step past our own additions to the native line (the arrows, the deadlines)
  // so the added lines land beneath the whole line rather than inside it.
  if (after) {
    let n = after.nextElementSibling;
    while (n && n.classList && (n.classList.contains(NAV_CLASS) || n.classList.contains(DL_CLASS))) {
      after = n;
      n = n.nextElementSibling;
    }
  }
  for (const s of extras) {
    const wrap = buildExtraHeaderRow(anchor, s.eff);
    if (after) {
      if (!after.parentNode) break;
      after.parentNode.insertBefore(wrap, after.nextSibling);
    } else {
      host.appendChild(wrap);
    }
    after = wrap;
  }
  // A band sized for exactly one line has to grow to hold the ones we added.
  if (extras.length) {
    try { anchor.band.style.height = 'auto'; anchor.band.style.overflow = 'visible'; } catch (_) {}
  }
}

/* ---- ‹ › : step between hearing dates ---- */

function renderHearingArrows(anchor) {
  const groups = __hearingGroups || [];
  const host = anchor.span.parentNode;
  if (!host) return;
  let nav = null;
  for (const n of host.children) {
    if (n.classList && n.classList.contains(NAV_CLASS)) { nav = n; break; }
  }
  if (groups.length < 2) { if (nav) nav.remove(); return; }
  if (!nav) {
    nav = document.createElement('span');
    nav.className = NAV_CLASS;
    nav.setAttribute('style', 'margin-left:8px;font-family:inherit;white-space:nowrap;user-select:none;');
    const mk = (glyph, delta, title) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = glyph;
      b.title = title;
      b.setAttribute('style', 'background:transparent;border:none;color:inherit;font:inherit;'
        + 'font-weight:bold;font-size:1.15em;cursor:pointer;padding:0 4px;line-height:1;');
      b.addEventListener('click', e => {
        e.preventDefault(); e.stopPropagation();
        stepHearingGroup(delta);
      });
      return b;
    };
    nav.appendChild(mk('‹', -1, 'Earlier hearing date'));
    const lbl = document.createElement('span');
    lbl.className = '__lac_hnav_count__';
    lbl.setAttribute('style', 'font-size:0.85em;opacity:0.8;');
    nav.appendChild(lbl);
    nav.appendChild(mk('›', 1, 'Later hearing date'));
    host.insertBefore(nav, anchor.span.nextSibling);
  }
  const btns = nav.querySelectorAll('button');
  const setState = (b, off) => {
    if (!b) return;
    b.disabled = off;
    b.style.opacity = off ? '0.3' : '1';
    b.style.cursor = off ? 'default' : 'pointer';
  };
  setState(btns[0], __selGroupIdx <= 0);
  setState(btns[1], __selGroupIdx >= groups.length - 1);
  const lbl = nav.querySelector('.__lac_hnav_count__');
  const count = (__selGroupIdx + 1) + '/' + groups.length;
  if (lbl && lbl.textContent !== count) lbl.textContent = count;
}

function stepHearingGroup(delta) {
  const groups = __hearingGroups || [];
  if (!groups.length) return;
  const next = Math.max(0, Math.min(groups.length - 1, __selGroupIdx + delta));
  if (next === __selGroupIdx) return;
  __selGroupIdx = next;
  writeSelDate(groups[next].date);
  dlLog('hearing date selected:', groups[next].date);
  // A different hearing means different deadlines, different documents to open
  // and a different motion type on the order template, so everything keyed to
  // the old one is dropped and rebuilt.
  __relevantDocsPromise = null;
  __deadlinePayloadPromise = null;
  applySelectedHearings();
  try { getRelevantDocumentsCached(); } catch (_) {}
  try { getDeadlinePayloadCached(); } catch (_) {}
}

/* ---- Committing a selection ---- */

// Rebuild the slots for the selected day and paint them. A slot already resolved
// for the same hearing keeps its filing status, so stepping away and back
// doesn't re-fetch the Documents tab.
function applySelectedHearings() {
  const effs = selectedHearingEffs();
  if (!effs.length) return;
  const prev = __dlSlots;
  __dlSlots = effs.map(eff => {
    const slot = makeSlot(eff);
    const hit = prev.find(s => slotKey(s.eff) === slotKey(eff));
    if (hit) {
      // Keep what's already known for this hearing: a resolved slot doesn't
      // re-fetch, and a slot seeded from the per-tab cache keeps its colours
      // while the fresh fetch runs behind it.
      slot.filed = hit.filed;
      slot.osc = hit.osc;
      slot.fetchStarted = hit.fetchStarted;
    }
    return slot;
  });
  if (!__dlSlots.some(paintableSlot) && __dlNoMotionLogged++ < 2) {
    dlLog('no briefable hearing on', effs[0].hearingDate, '— hearingType=', effs[0].hearingType);
  }
  dlLog('hearings on the selected day:',
    __dlSlots.map(s => (s.eff.hearingType || '') + ' @ ' + (s.eff.hearingDate || '')));
  dlCacheWrite();
  injectNextDeadlines();
  __dlSlots.forEach(s => { fetchSlotFilings(s); });
  syncDefaultJudgmentFeesButton();
}

// The Fees calculator belongs to a default-judgment OSC, so it comes and goes
// with the selected hearing rather than sticking once shown.
function syncDefaultJudgmentFeesButton() {
  if (slotsHaveOsc()) { try { renderDefaultJudgmentFeesButton(); } catch (_) {} return; }
  const btn = document.getElementById('__lacourt_djfees_btn__');
  if (btn) btn.remove();
}

/* ---- OSC Re: Failure to Prosecute Default Judgment ---- */

// Session-scoped cache for the OSC status, keyed by case number, so it is
// computed once per browser session instead of re-fetched on every tab change.
// (Requires the service worker to widen storage.session access to content
// scripts; falls back to recomputing if that isn't available.)
function oscCacheKey() {
  const cn = parseCaseNumber(document) || '';
  return cn ? 'oscStatus:' + cn : null;
}
function oscCacheGet(key) {
  return new Promise(resolve => {
    try {
      chrome.storage.session.get([key], r => {
        if (chrome.runtime.lastError) { resolve(null); return; }
        resolve((r && r[key]) || null);
      });
    } catch (_) { resolve(null); }
  });
}
function oscCacheSet(key, val) {
  try { chrome.storage.session.set({ [key]: val }, () => { void chrome.runtime.lastError; }); } catch (_) {}
}

async function fetchSlotOsc(slot) {
  if (!slot || slot.fetchStarted) return;
  slot.fetchStarted = true;
  const apply = osc => {
    // Same rebuild race as fetchSlotFilings: land the answer on the slot the
    // widget is painting from now, not the one captured before the rebuild.
    const live = liveSlotFor(slot);
    live.osc = osc;
    live.fetchStarted = true;
    try { injectNextDeadlines(); } catch (e) { dlLog('OSC recolour failed:', (e && e.message) || e); }
    setTimeout(() => { try { injectNextDeadlines(); } catch (_) {} }, 500);
  };
  const key = oscCacheKey();
  if (key) {
    const cached = await oscCacheGet(key);
    if (cached && cached.text) { apply(cached); return; }
  }
  const osc = await computeOscStatus(pageCaseCtx());
  apply(osc);
  // Cache real answers only — not transient failures, which should retry.
  if (key && osc && osc.text && osc.text !== 'Default status unavailable') {
    oscCacheSet(key, osc);
  }
}

/* ---- Per-tab cache so a sub-tab reload paints the final answer ---- */

// A case sub-tab switch is a full page reload, so without this the widget
// recomputes (and re-fetches Documents) on every tab — making the dates/colours
// flash each time. The deadlines themselves are pure arithmetic and are simply
// recomputed; only the filing status, which costs a fetch, is carried over. OSC
// status has its own (chrome.storage.session) cache and is skipped here.
const __DL_CACHE_PREFIX = 'lacourt.dl.';
function dlCacheKey() {
  const cn = parseCaseNumber(document) || '';
  return cn ? __DL_CACHE_PREFIX + cn : null;
}
function dlEpoch(d) { return (d && !isNaN(d)) ? d.getTime() : null; }
function dlUnepoch(v) { return (typeof v === 'number') ? new Date(v) : null; }
function dlSerFiled(f) {
  return f ? {
    filedKnown: !!f.filedKnown, motion: dlEpoch(f.motion), opp: dlEpoch(f.opp), reply: dlEpoch(f.reply),
    fac: f.fac ? { label: f.fac.label, when: dlEpoch(f.fac.when) } : null,
    nonOpp: f.nonOpp ? { slot: f.nonOpp.slot, when: dlEpoch(f.nonOpp.when) } : null,
  } : null;
}
function dlDeserFiled(s) {
  return {
    filedKnown: !!s.filedKnown, motion: dlUnepoch(s.motion), opp: dlUnepoch(s.opp), reply: dlUnepoch(s.reply),
    fac: s.fac ? { label: s.fac.label, when: dlUnepoch(s.fac.when) } : null,
    nonOpp: s.nonOpp ? { slot: s.nonOpp.slot, when: dlUnepoch(s.nonOpp.when) } : null,
  };
}
function dlCacheWrite() {
  const key = dlCacheKey(); if (!key) return;
  const slots = __dlSlots.map(s => ({
    eff: {
      motionType: s.eff.motionType || '', hearingType: s.eff.hearingType || '',
      hearingDate: s.eff.hearingDate || '', timeText: s.eff.timeText || '',
      raw: s.eff.raw || '', native: !!s.eff.native, lookedAhead: !!s.eff.lookedAhead,
    },
    // Ordinary briefing status only: an OSC has its own cache, and a
    // post-judgment slot's schedule is rebuilt from the docket, not from here.
    filed: (s.computed && !s.computed.osc && !s.computed.skip) ? dlSerFiled(s.filed) : null,
  }));
  if (!slots.length) return;
  try { sessionStorage.setItem(key, JSON.stringify({ slots })); } catch (_) {}
}
// Seed the slots from the cache (once) so the first paint is the final answer.
function seedDeadlinesFromCache() {
  if (__dlSlots.length) return false;
  const key = dlCacheKey(); if (!key) return false;
  let raw; try { raw = JSON.parse(sessionStorage.getItem(key) || 'null'); } catch (_) { raw = null; }
  if (!raw || !Array.isArray(raw.slots) || !raw.slots.length) return false;
  __dlSlots = raw.slots.map(r => {
    const slot = makeSlot(r.eff);
    if (r.filed) slot.filed = dlDeserFiled(r.filed);
    return slot;
  });
  return true;
}

let __dlComputeStarted = false;
function renderNextHeaderDeadlines() {
  try {
    // Read the page's own Next line before anything of ours can overwrite it.
    nativeSnapshot();

    // On a sub-tab reload the module state is cold; paint instantly from the
    // per-tab cache so the dates/colours don't visibly recompute.
    if (!__dlSlots.length) seedDeadlinesFromCache();
    if (__dlSlots.length) injectNextDeadlines();

    if (__dlComputeStarted) return;
    const snap = nativeSnapshot();
    if (!snap) return; // header not ready — a later poll/observer will retry
    __dlComputeStarted = true;

    // When the Next event is one we work up, paint it at once: the Hearings-tab
    // round trip that finds its companions shouldn't hold up the first paint.
    if (isWorkableHearing(snap.hearingType) && !__dlSlots.length) {
      __dlSlots = [makeSlot({
        motionType: snap.motionType, hearingType: snap.hearingType,
        hearingDate: snap.hearingDate, timeText: snap.timeText, raw: snap.raw,
        native: true, lookedAhead: false, ud: pageIsUnlawfulDetainer(),
      })];
      injectNextDeadlines();
      __dlSlots.forEach(s => { fetchSlotFilings(s); });
      syncDefaultJudgmentFeesButton();
    }

    // Then resolve every workable hearing on the case and commit the selection —
    // adding the arrows, and a header line for each further motion set that day.
    getSelectedHearingGroup()
      .then(() => applySelectedHearings())
      .catch(e => dlLog('hearing resolution failed:', e && e.message || e));
  } catch (e) { dlLog('render error:', e && e.message || e); }
}


// Re-inject if e-court re-renders the header and strips our node (the render
// poll stops after ~10s, so an observer keeps it pinned thereafter).
let __nextDlObserver = null;
function observeNextHeader() {
  if (__nextDlObserver || typeof MutationObserver === 'undefined' || !document.body) return;
  let pending = false;
  const obs = new MutationObserver(() => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      try { renderNextHeaderDeadlines(); } catch (_) {}
      try { initHeaderExpanders(); } catch (_) {}
      // Keep the floating buttons sized to the blue bar if the SPA re-rendered it.
      try { scheduleButtonCollapse(); } catch (_) {}
    });
  });
  try { obs.observe(document.body, { childList: true, subtree: true }); __nextDlObserver = obs; } catch (_) {}
}

/* ------------------------------------------------------------------ */
/* Expandable case name / case type in the header                      */
/* ------------------------------------------------------------------ */
//
// eCourt truncates both the case name and the case-type designation in the
// case header (server-side, with a literal "..." — same habit as the agenda's
// hearing names). Clicking either now expands it; clicking again restores the
// original. eCourt keeps no full copy of these strings on the page, so the
// full text comes from, in order:
//   1. a title attribute on or near the element, when eCourt provides one;
//   2. for the CASE TYPE — the designations are the standard civil case types
//      (the CM-010 / LASC catalog), a finite list embedded below, matched by
//      prefix against the truncated text;
//   3. for the CASE NAME — the Parties tab, which carries every party in
//      full: the title is rebuilt as "<first claimant>[, et al.] vs
//      <first defendant>[, et al.]" (the format eCourt itself uses) and used
//      only when the truncated text is a prefix of it;
//   4. failing all of those, the click still lifts any CSS clipping so a
//      style-truncated header shows its own full text.

// The standard civil case-type names (CM-010 / LASC). Matching is punctuation-
// insensitive, so "Unlawful Detainer/Commercial" and "Unlawful Detainer -
// Commercial" both hit the same entry. Extend the list if a type on a real
// case fails to expand (the miss is logged to the console).
const CASE_TYPE_CATALOG = [
  'Motor Vehicle - Personal Injury/Property Damage/Wrongful Death',
  'Uninsured Motorist - Personal Injury/Property Damage/Wrongful Death',
  'Asbestos Property Damage',
  'Asbestos - Personal Injury/Wrongful Death',
  'Product Liability (not asbestos or toxic/environmental)',
  'Medical Malpractice - Physicians & Surgeons',
  'Other Professional Health Care Malpractice',
  'Premises Liability (e.g., slip and fall)',
  'Intentional Bodily Injury/Property Damage/Wrongful Death (e.g., assault, vandalism)',
  'Intentional Infliction of Emotional Distress',
  'Negligent Infliction of Emotional Distress',
  'Other Personal Injury/Property Damage/Wrongful Death',
  'Business Tort/Unfair Business Practice',
  'Civil Rights (e.g., discrimination, false arrest) (not civil harassment)',
  'Defamation (e.g., slander, libel)',
  'Fraud (no contract)',
  'Intellectual Property',
  'Legal Malpractice',
  'Other Professional Malpractice (not medical or legal)',
  'Other Non-Personal Injury/Property Damage Tort',
  'Wrongful Termination',
  'Other Employment Complaint Case',
  'Labor Commissioner Appeals',
  'Breach of Rental/Lease Contract (not unlawful detainer or wrongful eviction)',
  'Contract/Warranty Breach - Seller Plaintiff (no fraud/negligence)',
  'Negligent Breach of Contract/Warranty (no fraud)',
  'Other Breach of Contract/Warranty (not fraud or negligence)',
  'Collections Case - Seller Plaintiff',
  'Other Promissory Note/Collections Case',
  'Insurance Coverage (not complex)',
  'Contractual Fraud',
  'Tortious Interference',
  'Other Contract Dispute (not breach/insurance/fraud/negligence)',
  'Eminent Domain/Inverse Condemnation',
  'Wrongful Eviction Case',
  'Mortgage Foreclosure',
  'Quiet Title',
  'Other Real Property (not eminent domain, landlord/tenant, foreclosure)',
  'Unlawful Detainer/Commercial (not drugs or wrongful eviction)',
  'Unlawful Detainer/Residential (not drugs or wrongful eviction)',
  'Unlawful Detainer/Post-Foreclosure',
  'Unlawful Detainer/Drugs',
  'Asset Forfeiture Case',
  'Petition re Arbitration Award',
  'Writ - Administrative Mandamus',
  'Writ - Mandamus on Limited Court Case Matter',
  'Writ - Other Limited Court Case Review',
  'Other Writ/Judicial Review',
  'Antitrust/Trade Regulation',
  'Construction Defect',
  'Claims Involving Mass Tort',
  'Securities Litigation',
  'Toxic Tort/Environmental',
  'Insurance Coverage Claims from Provisionally Complex Case',
  'Sister State Judgment',
  'Abstract of Judgment',
  'Confession of Judgment (non-domestic relations)',
  'Administrative Agency Award (not unpaid taxes)',
  'Petition/Certificate for Entry of Judgment on Unpaid Taxes',
  'Other Enforcement of Judgment Case',
  'RICO Case',
  'Declaratory Relief Only',
  'Injunctive Relief Only (not domestic/harassment)',
  'Other Commercial Complaint Case (non-tort/non-complex)',
  'Other Civil Complaint (non-tort/non-complex)',
  'Partnership and Corporate Governance Case',
  'Civil Harassment',
  'Workplace Violence',
  'Elder/Dependent Adult Abuse Case',
  'Election Contest',
  'Petition for Change of Name/Change of Gender',
  'Petition for Relief from Late Claim Law',
  'Other Civil Petition',
];

const TRUNC_TEXT_RE = /(?:\.{3,}|…)\s*$/;
// The ellipsis can fall MID-text, not only at the end: a truncated plaintiff
// reads "GLOBEX HOLDING COMPA... vs TAYLOR ROE", and the type line can carry
// more header text after its "...". Everything below works on "an ellipsis
// anywhere", splitting the text into the segments around it.
const TRUNC_ANY_RE = /(?:\.{3,}|…)/;
const TRUNC_SPLIT_RE = /(?:\.{3,}|…)/g;

function expNorm(s) { return (s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

function isCssClipped(el) {
  try { return el.scrollWidth > el.clientWidth + 1; } catch (_) { return false; }
}

// An element whose own box visibly cuts its text off: wider content than box,
// or an ellipsis style actually engaged. This is the shape eCourt's header
// really uses — each party name and the case type sit in fixed-width boxes
// with the FULL text in the DOM and CSS painting the "…" — so this, not a
// literal "..." in the text, is the primary thing to look for.
function isReallyClipped(el) {
  try {
    if (el.clientWidth > 0 && el.scrollWidth > el.clientWidth + 1) return true;
    const cs = getComputedStyle(el);
    if (cs && cs.textOverflow === 'ellipsis' && cs.overflow !== 'visible'
        && el.scrollWidth > el.clientWidth && el.clientWidth > 0) return true;
  } catch (_) {}
  return false;
}

// Text just before/after the element on its line — its previous/next siblings'
// text, then the parent's, a couple of levels up. What classifies a clipped
// fragment: the case number sits just BEFORE the name, "Civil Unlimited" just
// before the type, "vs" beside a party-name box.
function nearbyText(el, dir) {
  let out = '';
  let node = el, depth = 0;
  while (node && depth < 3 && out.length < 100) {
    let sib = dir === 'prev' ? node.previousSibling : node.nextSibling;
    while (sib && out.length < 100) {
      const t = sib.textContent || '';
      out = dir === 'prev' ? t + ' ' + out : out + ' ' + t;
      sib = dir === 'prev' ? sib.previousSibling : sib.nextSibling;
    }
    node = node.parentElement; depth++;
  }
  return out.replace(/\s+/g, ' ').trim();
}

// What a clipped header fragment is, from its own text and its neighbors:
//   'type' — carries or directly follows the "Civil Unlimited/Limited"
//            designation (the case type is always right after it);
//   'name' — carries or directly follows the case number (the name is
//            immediately right of the number), or sits beside a "vs";
//   'other' — some other clipped header text; still expandable, generically.
function classifyHeaderClip(el) {
  const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
  const pre = nearbyText(el, 'prev');
  const post = nearbyText(el, 'next');
  if (/\bcivil\s+(?:unlimited|limited)\b/i.test(t)
      || /\bcivil\s+(?:unlimited|limited)\b\s*[:\-–—]?\s*$/i.test(pre)) return 'type';
  const cn = parseCaseNumber(document);
  if (cn && (t.indexOf(cn) !== -1 || pre.slice(-40).indexOf(cn) !== -1)) return 'name';
  if (/\s(?:vs?\.?|versus)\s/i.test(t)) return 'name';
  if (/^(?:vs?\.?|versus)\b/i.test(post)) return 'name';  // a plaintiff box — "vs" follows it
  if (/\b(?:vs?\.?|versus)\s*$/i.test(pre)) return 'name'; // a defendant box — "vs" precedes it
  return 'other';
}

// Every CSS-clipped text fragment in the case-header area. Scoped to the
// [class*="case"] blocks and a few of their ancestors (never the whole page,
// where site chrome legitimately clips), skipping containers.
function findClippedHeaderEls() {
  const out = [];
  const seenScope = new Set();
  const scopes = [];
  for (const box of document.querySelectorAll('[class*="case"]')) {
    for (let n = box, up = 0; n && n.querySelectorAll && up < 3; n = n.parentElement, up++) {
      if (!seenScope.has(n)) { seenScope.add(n); scopes.push(n); }
    }
  }
  const seenEl = new Set();
  for (const scope of scopes) {
    let all;
    try { all = scope.getElementsByTagName('*'); } catch (_) { continue; }
    if (all.length > 800) continue; // a page-level container, not the header
    for (const el of all) {
      if (seenEl.has(el)) continue;
      seenEl.add(el);
      const t = (el.textContent || '').trim();
      if (!t || t.length > 400) continue;
      if (el.children.length > 2) continue; // a fragment, not a container
      if (!isReallyClipped(el)) continue;
      out.push(el);
    }
  }
  return out;
}

// Does `full` genuinely fill in `truncated`? The truncated text is split at
// its ellipses; the first segment must start `full`, every later segment must
// follow in order, and — unless the text ends at an ellipsis — the last
// segment must end it. So "GLOBEX HOLDING COMPA... vs TAYLOR ROE" accepts
// "GLOBEX HOLDING COMPANY, LLC vs TAYLOR ROE" and rejects a title whose tail
// doesn't match. Punctuation-insensitive throughout.
function extendsTruncated(full, truncated) {
  const nf = expNorm(full);
  const rawParts = (truncated || '').split(TRUNC_SPLIT_RE);
  const parts = rawParts.map(expNorm);
  if (!parts.length || !parts[0] || !nf) return false;
  if (!nf.startsWith(parts[0])) return false;
  let pos = parts[0].length;
  for (let i = 1; i < parts.length; i++) {
    if (!parts[i]) continue; // empty tail when the text ends at the ellipsis
    const at = nf.indexOf(parts[i], pos);
    if (at === -1) return false;
    pos = at + parts[i].length;
  }
  const lastRaw = rawParts[rawParts.length - 1];
  if (expNorm(lastRaw) && !nf.endsWith(parts[parts.length - 1])) return false;
  return nf.length > expNorm(parts.join(' ')).length; // actually adds something
}

// Expand a truncated case-type line against the catalog. The "Civil
// Unlimited"/"Civil Limited" lead is kept, the text UP TO the first ellipsis
// is prefix-matched punctuation-insensitively (shortest hit wins), and any
// text AFTER the ellipsis — the rest of the header line — is kept, so the
// expansion pushes it along rather than swallowing it.
function expandCaseTypeText(displayed) {
  // pre = anything before the designation (a "Case Type:" label…), lead = the
  // "Civil Unlimited/Limited" phrase, then the truncated type up to the FIRST
  // ellipsis, then the rest of the line — kept, so the expansion pushes it
  // along rather than swallowing it.
  const m = (displayed || '').match(/^([^]*?)\b(civil\s+(?:unlimited|limited)\b)\s*([^]*?)\s*(?:\.{3,}|…)([^]*)$/i);
  if (!m) return ''; // no designation, or nothing truncated — nothing to expand
  const pre = m[1].replace(/\s+/g, ' ').trim();
  const lead = m[2].replace(/\s+/g, ' ').trim();
  const p = expNorm(m[3]);
  const tail = (m[4] || '').replace(/^[\s.]+/, '').trim();
  if (!p) return '';
  const hits = CASE_TYPE_CATALOG
    .filter(tp => { const n = expNorm(tp); return n.startsWith(p) && n.length > p.length; })
    .sort((a, b) => a.length - b.length);
  if (!hits.length) { dlLog('case-type expansion: no catalog hit for', m[3]); return ''; }
  if (hits.length > 1) dlLog('case-type expansion: multiple catalog hits for', m[3], '→ using', hits[0]);
  return (pre ? pre + ' ' : '') + lead + ' ' + hits[0] + (tail ? ' ' + tail : '');
}

// A title attribute on or near the element that carries the displayed text in
// full — eCourt uses title attributes for full text elsewhere (the Next
// header), so check before anything costlier.
function titleAttrExpansion(el, truncatedCore) {
  if (!expNorm(truncatedCore)) return '';
  const pool = [el];
  for (let n = el.parentElement, i = 0; n && i < 3; n = n.parentElement, i++) pool.push(n);
  try { for (const d of el.querySelectorAll('[title]')) pool.push(d); } catch (_) {}
  for (const n of pool) {
    const t = ((n.getAttribute && n.getAttribute('title')) || '').replace(/\s+/g, ' ').trim();
    if (!t || t === 'Click to expand') continue;
    if (extendsTruncated(t, truncatedCore)) return t;
  }
  return '';
}

// The parties, parsed once per call from the same memoized fetch the status
// engine uses. Logged when empty — an empty list is why no expansion could be
// built, and that is the fact worth seeing in the console.
async function captionParties() {
  try {
    const parties = parsePartiesTable(await pageCaseCtx().parties()) || [];
    if (!parties.length) dlLog('case-name expansion: parties table came back empty');
    return parties;
  } catch (e) { dlLog('case-name expansion: parties fetch failed:', (e && e.message) || e); return []; }
}

// Rebuild the case title from the Parties tab: "<first claimant>[, et al.] vs
// <first defendant>[, et al.]" — the format eCourt's own headers use.
function rebuildCaptionFrom(parties) {
  const claim = parties.filter(p => /^\s*(?:plaintiff|petitioner|cross[-\s]?complainant)\b/i.test(p.role || ''));
  const resp = parties.filter(p => /^\s*(?:defendant|respondent)\b/i.test(p.role || ''));
  if (!claim.length || !resp.length) return '';
  const side = list => list[0].name + (list.length > 1 ? ', et al.' : '');
  return side(claim) + ' vs ' + side(resp);
}

async function reconstructCaseNameFromParties() {
  return rebuildCaptionFrom(await captionParties());
}

// Complete the trailing words of `segment` from one of `candidates` (party
// names): the longest suffix of the segment that is a prefix of a candidate
// wins, and the candidate replaces it. "26STCP01881 J.G. WENTWORTH
// ORIGINATIONS" + ["J.G. WENTWORTH ORIGINATIONS, LLC", …] →
// "26STCP01881 J.G. WENTWORTH ORIGINATIONS, LLC".
function completeTruncatedSegment(segment, candidates) {
  const words = (segment || '').split(/(\s+)/); // keep separators so offsets survive
  for (let i = 0; i < words.length; i++) {      // longest suffix first
    const suffix = words.slice(i).join('').trim();
    if (!suffix) continue;
    const ns = expNorm(suffix);
    if (!ns) continue;
    for (const c of candidates) {
      const nc = expNorm(c);
      if (nc.length > ns.length && nc.startsWith(ns)) {
        const cut = segment.lastIndexOf(suffix);
        return segment.slice(0, cut) + c;
      }
    }
  }
  return null;
}

// Fill each "<prefix>..." in the caption from a party name that extends the
// prefix, leaving everything else exactly as displayed. This is what handles
// the caption whose OTHER side doesn't match the Parties tab — e.g. a
// respondent shown by initials ("… vs E. F.") while the parties list the full
// name: the whole-caption rebuild rightly refuses to rewrite "E. F.", but the
// truncated plaintiff can still be completed on its own.
function expandNameSegmentsFrom(name, parties) {
  const candidates = parties.map(p => p && p.name).filter(Boolean);
  if (!candidates.length) return '';
  const re = /(\.{3,}|…)/g;
  let out = '', last = 0, changed = false, m;
  while ((m = re.exec(name)) !== null) {
    const seg = name.slice(last, m.index);
    const completed = completeTruncatedSegment(seg, candidates);
    if (completed != null) { out += completed; changed = true; }
    else out += seg + m[1];
    last = m.index + m[1].length;
  }
  out += name.slice(last);
  return changed ? out : '';
}

function resolveFullCaseTypeText(el, displayed) {
  const t = titleAttrExpansion(el, displayed);
  if (t) return t;
  // A title attribute may carry the full TYPE without the rest of the line —
  // match it against the truncated head alone and re-attach the tail.
  const m = (displayed || '').match(/^([^]*?)(?:\.{3,}|…)([^]*)$/);
  if (m) {
    const tail = (m[2] || '').replace(/^[\s.]+/, '').trim();
    const t2 = titleAttrExpansion(el, m[1].trim() + '...');
    if (t2) return t2 + (tail ? ' ' + tail : '');
  }
  return expandCaseTypeText(displayed);
}

async function resolveFullCaseNameText(el, displayed) {
  // The name as THIS element shows it: whatever follows the case number in
  // its own text (the element may carry "26STCV01234 NAME…" in one run).
  // parseCaseName is deliberately not used here — it reads the first header
  // candidate on the page, which can be a coarser container than the element
  // being expanded.
  const cn = parseCaseNumber(document);
  let name = displayed;
  if (cn) {
    const i = displayed.indexOf(cn);
    if (i !== -1) name = displayed.slice(i + cn.length).trim().replace(/^[,;:\-\s]+/, '');
  }
  if (!expNorm(name)) return '';
  const swapIn = full => (name && displayed.indexOf(name) !== -1) ? displayed.replace(name, full) : full;
  const t = titleAttrExpansion(el, name);
  if (t) return swapIn(t);
  const parties = await captionParties();
  // Whole-caption rebuild first — used only when it genuinely fills the
  // truncated header in: the segments around the ellipsis must all line up.
  const recon = rebuildCaptionFrom(parties);
  if (recon && extendsTruncated(recon, name)) return swapIn(recon);
  // Then segment-wise: complete just the truncated chunk(s) from party names,
  // preserving the rest of the caption verbatim — the path that handles a
  // side shown differently in the caption than in the parties list (initials
  // for a protected party, added descriptors, et al. groupings).
  const seg = expandNameSegmentsFrom(name, parties);
  if (seg && expNorm(seg) !== expNorm(name)) {
    dlLog('case-name expansion: segment-completed from parties —', { header: name, expanded: seg });
    return swapIn(seg);
  }
  if (recon) dlLog('case-name expansion: parties rebuild does not fill the header in —', { header: name, rebuilt: recon });
  return '';
}

// The header element carrying the (truncated) case name. Two shapes are
// searched, smallest element wins, case-header blocks before the whole page:
//   1. the case number followed by the name in one text run;
//   2. the caption on its own — the only header text shaped "SOMEONE vs
//      SOMEONE" (the number lives in a separate element).
// An element that also swallows the case-type line ("Civil Unlimited …") is a
// container, not the name — expanding it would rewrite the type line's
// element out from under its own toggle, so those are skipped.
function findCaseNameEl() {
  const cn = parseCaseNumber(document);
  const scopes = [];
  const seen = new Set();
  for (const box of document.querySelectorAll('[class*="case"]')) {
    for (let n = box, up = 0; n && n.querySelectorAll && up < 3; n = n.parentElement, up++) {
      if (!seen.has(n)) { seen.add(n); scopes.push(n); }
    }
  }
  if (document.body && !seen.has(document.body)) scopes.push(document.body);
  const TAGS = 'span, div, b, i, em, strong, h1, h2, h3, h4, a, td, p, label';
  const pick = test => {
    for (const scope of scopes) {
      let best = null, bestLen = Infinity;
      for (const el of scope.querySelectorAll(TAGS)) {
        const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
        if (!t || t.length > 300) continue;
        if (/\bcivil\s+(?:unlimited|limited)\b/i.test(t)) continue; // the type line rides along — too coarse
        if (!test(t)) continue;
        // <= so that of nested elements with the SAME text, the deepest wins.
        if (t.length <= bestLen) { best = el; bestLen = t.length; }
      }
      if (best) return best;
    }
    return null;
  };
  // The caption shape goes FIRST: the smallest "SOMEONE vs SOMEONE" element is
  // the caption itself whether or not the number shares it, while the
  // number-anchored match can land on a container holding the number and the
  // caption as separate children — expanding that would flatten them.
  const byVs = pick(t => /\S.+\s(?:vs?\.?|versus)\s.+\S/i.test(t));
  if (byVs) return byVs;
  if (cn) {
    return pick(t => { const i = t.indexOf(cn); return i !== -1 && !!t.slice(i + cn.length).trim(); });
  }
  return null;
}

// Expand/collapse. Expanding swaps the full text in place (so any text after
// it in the line reflows right by exactly the added width) and lifts CSS
// clipping; collapsing restores the exact original text and inline styles.
// When nothing can be resolved AND nothing is CSS-clipped, the element stays
// collapsed — so a later click retries instead of toggling a no-op state.
async function toggleHeaderExpansion(el, resolver) {
  if (el.getAttribute('data-lac-exp-open') === '1') {
    const orig = el.getAttribute('data-lac-exp-text');
    if (orig != null) el.textContent = orig;
    el.style.cssText = el.getAttribute('data-lac-exp-css') || '';
    el.style.setProperty('cursor', 'pointer');
    el.setAttribute('data-lac-exp-open', '0');
    return false;
  }
  const displayed = (el.textContent || '').replace(/\s+/g, ' ').trim();
  let full = el.getAttribute('data-lac-exp-full') || '';
  if (!full) {
    try { full = (await resolver(el, displayed)) || ''; } catch (_) { full = ''; }
    if (full) el.setAttribute('data-lac-exp-full', full);
  }
  const clipped = isReallyClipped(el) || isCssClipped(el);
  if (!full && !clipped) {
    if (TRUNC_ANY_RE.test(displayed)) dlLog('no full text found for truncated header text:', displayed);
    return false;
  }
  if (el.getAttribute('data-lac-exp-text') == null) el.setAttribute('data-lac-exp-text', el.textContent);
  if (el.getAttribute('data-lac-exp-css') == null) el.setAttribute('data-lac-exp-css', el.style.cssText || '');
  if (full && full !== displayed) el.textContent = full;
  // Lift any CSS clipping too, so a style-truncated header shows everything —
  // including a fixed width, which is how eCourt sizes the boxes it clips.
  el.style.setProperty('white-space', 'normal', 'important');
  el.style.setProperty('overflow', 'visible', 'important');
  el.style.setProperty('text-overflow', 'clip', 'important');
  el.style.setProperty('max-width', 'none', 'important');
  el.style.setProperty('width', 'auto', 'important');
  el.style.setProperty('height', 'auto', 'important');
  el.style.setProperty('cursor', 'pointer', 'important');
  el.setAttribute('data-lac-exp-open', '1');
  return true;
}

// Bind the click toggle. Returns 'bound' when the element is truncated and
// now clickable, 'done' when nothing about it is cut off, and 'retry' when
// the header hasn't laid out yet (so clipping can't be measured). `force`
// skips the truncation test — the clipped-fragment scan has already verified
// the element is cut off.
function bindHeaderExpander(el, resolver, force) {
  if (!el) return 'retry';
  if (el.getAttribute('data-lac-exp-bound') === '1') return 'bound';
  if (!force) {
    if (el.clientWidth === 0 && el.offsetParent === null) return 'retry'; // not laid out yet
    const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
    if (!TRUNC_ANY_RE.test(t) && !isReallyClipped(el) && !isCssClipped(el)) return 'done'; // nothing cut off
  }
  el.setAttribute('data-lac-exp-bound', '1');
  el.style.setProperty('cursor', 'pointer');
  if (!el.getAttribute('title')) el.setAttribute('title', 'Click to expand');
  el.addEventListener('click', ev => {
    try { ev.preventDefault(); ev.stopPropagation(); } catch (_) {}
    dlLog('header expander clicked:', describeExpEl(el));
    toggleHeaderExpansion(el, resolver);
  });
  dlLog('header expander bound:', describeExpEl(el));
  return 'bound';
}

// Tag.class + leading text, for the console — enough to see which element the
// expander latched onto (or to report which one it should have).
function describeExpEl(el) {
  if (!el) return '(none)';
  const cls = el.className ? '.' + String(el.className).trim().replace(/\s+/g, '.') : '';
  const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
  return el.tagName + cls + ' "' + t.slice(0, 100) + (t.length > 100 ? '…' : '') + '"';
}

// Idempotent; retried from the render poll and the header observer for a
// while (elements render and lay out late). Two discovery paths feed it:
//   - the literal-ellipsis finders (findCaseTypeEl / findCaseNameEl), for a
//     header whose text carries a server-side "...";
//   - the clipped-fragment scan (findClippedHeaderEls), for eCourt's real
//     shape — fixed-width boxes CSS-clipping full text, classified by their
//     neighbors (the name sits right of the case number, the type right
//     after "Civil Unlimited").
// Defaults: every NAME element auto-expands (once — collapsing it by hand is
// respected); the TYPE and anything else stay collapsed until clicked.
const __exp = { autoQueue: [], bound: [], logged: false };
let __expanderTries = 0;
function queueAutoExpand(el, resolver) {
  if (__exp.autoQueue.some(q => q.el === el)) return;
  __exp.autoQueue.push({ el, resolver, done: false, busy: false, tries: 0 });
}
function runAutoQueue() {
  for (const q of __exp.autoQueue) {
    if (q.done || q.busy) continue;
    if (q.el.getAttribute('data-lac-exp-open') === '1') { q.done = true; continue; }
    // A caption that can't be resolved stays truncated: stop auto-retrying
    // after a few attempts (a manual click still retries the lookup).
    if (q.tries >= 6) { q.done = true; continue; }
    q.tries++;
    q.busy = true;
    toggleHeaderExpansion(q.el, q.resolver)
      .then(opened => { if (opened) q.done = true; })
      .catch(() => {})
      .then(() => { q.busy = false; });
  }
}
function noteBound(el, kind) {
  if (!__exp.bound.some(b => b.el === el)) __exp.bound.push({ el, kind });
}
function initHeaderExpanders() {
  if (__expanderTries > 60) { return; }
  __expanderTries++;
  try {
    // Server-truncated whole lines.
    const typeEl = findCaseTypeEl(document);
    if (typeEl && bindHeaderExpander(typeEl, resolveFullCaseTypeText) === 'bound') noteBound(typeEl, 'type');
    const nameEl = findCaseNameEl();
    if (nameEl && bindHeaderExpander(nameEl, resolveFullCaseNameText) === 'bound') {
      noteBound(nameEl, 'name');
      queueAutoExpand(nameEl, resolveFullCaseNameText);
    }
    // CSS-clipped fragments.
    for (const el of findClippedHeaderEls()) {
      const kind = classifyHeaderClip(el);
      const resolver = kind === 'type' ? resolveFullCaseTypeText
        : kind === 'name' ? resolveFullCaseNameText
        : titleAttrExpansion; // generic: a title attr if there is one, else just unclip
      if (bindHeaderExpander(el, resolver, true) === 'bound') {
        noteBound(el, kind);
        if (kind === 'name') queueAutoExpand(el, resolver);
      }
    }
    runAutoQueue();
    // One console line when the search settles, so a page where nothing bound
    // says so — paste it (with the header's outerHTML) to debug.
    if (!__exp.logged && __expanderTries >= 20) {
      __exp.logged = true;
      dlLog('header expanders settled:', __exp.bound.length
        ? __exp.bound.map(b => b.kind + ': ' + describeExpEl(b.el))
        : 'NOTHING BOUND — no truncated or clipped header element found');
    }
  } catch (_) {}
}

function setupFillFormButton() {
  // The parties table loads after initial page render. Try once on DOMContentLoaded
  // and once on full load, then poll briefly until it appears (cap at ~10s).
  const tryRender = () => {
    ensureButtonStyles();
    renderFillFormButton();
    renderDocumentsButton();
    renderDeadlineButton();
    renderDefaultJudgmentFeesButton(); // default-judgment pages only
    renderNextHeaderDeadlines();
    try { initHeaderExpanders(); } catch (_) {}
    observeNextHeader();
    // Size immediately from the remembered bar dimensions (no flash to default),
    // then schedule the debounced pass that measures the live bar and re-docks.
    try { updateButtonCollapse(); } catch (_) {}
    scheduleButtonCollapse();
    // ...and keep looking for the header if it hasn't rendered yet.
    retryBarSizingUntilFound();
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', tryRender, { once: true });
  } else {
    tryRender();
  }
  window.addEventListener('load', tryRender, { once: true });

  // Poll for up to 10s in case the table is rendered late by site scripts.
  let polls = 0;
  const interval = setInterval(() => {
    polls++;
    if (document.getElementById('__lacourt_docs_btn__') || polls > 20) {
      clearInterval(interval);
      return;
    }
    tryRender();
  }, 500);

  // Prefetch the relevant-documents set once the page has settled, so pressing
  // the Documents button opens instantly. This only fetches/computes — it never
  // opens tabs; the button does that.
  const prefetchDocs = () => {
    try { getRelevantDocumentsCached(); } catch (_) {}
    // Same for the Deadlines hand-off: detecting the trigger dates reads the
    // Documents tab, and doing it now means the button opens without a pause.
    try { getDeadlinePayloadCached(); } catch (_) {}
  };
  if (document.readyState === 'complete') setTimeout(prefetchDocs, 1000);
  else window.addEventListener('load', () => setTimeout(prefetchDocs, 1000), { once: true });
}

setupFillFormButton();
})();

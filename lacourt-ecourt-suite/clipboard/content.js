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

  /* ------------------------------------------------------------------ */
  /* OSC Re: Failure to Prosecute Default Judgment — alternate flow     */
  /* ------------------------------------------------------------------ */
  //
  // When the case's next event is an OSC Re: Failure to Prosecute Default
  // Judgment, the Fill Microsoft Form button:
  //   1. Routes to a different (shorter) Microsoft Form that only collects
  //      Case Number and Hearing Date.
  //   2. Strips the rotation down to just those two fields so auto-fill
  //      doesn't try to fill non-existent party fields.
  //   3. Also opens a pre-composed mailto: link addressed to Judge
  //      Mackenzie with the case info in the subject and a standard body.
  //
  // Ctrl+A rotation and manual selection paste are unaffected.
  // Use DesignPageV2 with topview=Preview so the form loads in the owner's
  // authenticated context (the Auto-Export companion extension needs this
  // session to call the owner-API; the public ResponsePage URL returns 401).
  const OSC_FORM_URL = 'https://forms.office.com/Pages/DesignPageV2.aspx?prevorigin=rbf&origin=NeoPortalPage&rpring=UsGovGccProduction&subpage=design&id=x8OU3Ei7_0CTBeRz_W9qFt74YgjxwElOsa89AoRCn9FUQldVVTI0OUlZSUc0UTNMTDdISDNWU0JUNS4u&analysis=false&tab=0&topview=Preview';
  const REGULAR_FORM_URL = 'https://forms.office.com/Pages/DesignPageV2.aspx?prevorigin=rbf&origin=NeoPortalPage&rpring=UsGovGccProduction&subpage=design&id=x8OU3Ei7_0CTBeRz_W9qFt74YgjxwElOsa89AoRCn9FUQzNGQ0NPWVpUMDBVTzcwN1I2Q0JFOVFZVi4u&analysis=false&tab=0&topview=Preview';

  // Strict trigger: must include "Default Judgment" — the OSC for other
  // reasons (sanctions, etc.) doesn't go through this flow.
  const OSC_DEFAULT_JUDGMENT_RE = /\border\s+to\s+show\s+cause\s+re:?\s+failure\s+to\s+prosecute\s+default\s+judgment\b/i;

  function isOscDefaultJudgment(hearingType) {
    if (!hearingType) return false;
    return OSC_DEFAULT_JUDGMENT_RE.test(hearingType);
  }

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
  function buildOscMailto(caseNumber, hearingDate, caseName) {
    if (!caseNumber || !hearingDate) return null;

    const EM = '\u2013'; // en-dash — what the user calls "em dash" colloquially
    const subjectParts = [hearingDate, caseNumber];
    if (caseName) subjectParts.push(caseName);
    subjectParts.push('OSC RE: FAILURE TO PROSECUTE DEFAULT JUDGMENT');
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
    const m = motionType.toLowerCase();
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
  // colliding parties (e.g. two "Marriott" entities) get widened.
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

  if (!isOsc) {
    return {
      data,
      formUrl: REGULAR_FORM_URL,
      openUrl: chrome.runtime.getURL('order-template/order-template.html'),
      isOrderTemplate: true,
      mailtoUrl: null,
      isOsc: false,
      hearingType,
    };
  }

  // OSC flow — trim everything except case number and hearing date so
  // paste-rotator's auto-fill matcher only fills those two OSC form
  // fields. The labeled object's other keys are deleted; the rotation
  // sequence is rebuilt from the two surviving values.
  const trimmedLabeled = {};
  if (data.labeled.caseNumber)  trimmedLabeled.caseNumber  = data.labeled.caseNumber;
  if (data.labeled.hearingDate) trimmedLabeled.hearingDate = data.labeled.hearingDate;
  const trimmedSequence = [];
  if (trimmedLabeled.caseNumber)  trimmedSequence.push(trimmedLabeled.caseNumber);
  if (trimmedLabeled.hearingDate) trimmedSequence.push(trimmedLabeled.hearingDate);

  // Build mailto using values from the labeled object (already cleaned)
  // plus a fresh case-name parse. parseCaseName accepts an optional case
  // number hint so it can pin the location precisely.
  const caseName = parseCaseName(trimmedLabeled.caseNumber, root);
  const mailtoUrl = buildOscMailto(
    trimmedLabeled.caseNumber,
    trimmedLabeled.hearingDate,
    caseName
  );

  console.log('[LACourt] OSC default-judgment flow:', {
    hearingType, caseName, mailtoUrl: !!mailtoUrl,
  });

  return {
    data: { sequence: trimmedSequence, labeled: trimmedLabeled },
    formUrl: OSC_FORM_URL,
    openUrl: OSC_FORM_URL,
    isOrderTemplate: false,
    mailtoUrl,
    isOsc: true,
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
        computeMovant(ctx.data.labeled.motionType, result.partiesRoot).then(movant => {
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
 *     America Corp" yields "Bank of America" not "Bank of"). Trailing
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
 *   "WALMART STORES INC A DELAWARE CORPORATION"
 *     → formatted: "Walmart Stores Inc a Delaware Corporation"
 *       shortName: "Walmart Stores"
 *   "BANK OF AMERICA CORP"
 *     → formatted: "Bank of America Corp"
 *       shortName: "Bank of America"   (extended past stop-word "of")
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
  // (e.g. "Bank of America" → "Bank of"), extend by one more word so the
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

    // Drop any parties belonging to a 2nd+ cross-complaint section.
    if (currentSection === 'cross-skip') continue;

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
      const parenIdx = name.indexOf('(');
      if (parenIdx !== -1) name = name.substring(0, parenIdx).trim();
      // Also strip trailing "Update Party" if it leaked in.
      name = name.replace(/\s*update\s*party\s*$/i, '').trim();
    }

    if (name) {
      parties.push({ name, role, dismissed });
    }
  }

  return parties;
}

// Matches both LA Superior Court case-number formats:
//   - Current year-first: 2 digits + location/type letters + sequence digits,
//     e.g. "25STCV32877", "21STCR00001".
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

  // 2) The page title leads with the case number, e.g. "BC717394: DOCUMENTS ...".
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
      if (m) return stripEventId(m[1]);
    }
    const text = (span.textContent || '').trim().replace(/\s+/g, ' ');
    if (text) {
      const m = text.match(re);
      if (m) return stripEventId(m[1]);
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
      if (m) return stripEventId(m[1]);
    }
    const text = (span.textContent || '').trim().replace(/\s+/g, ' ');
    if (text) {
      const m = text.match(re);
      if (m) return stripEventId(m[1]);
    }
  }
  return '';
}

/**
 * Parses the case name from the e-court page header. The case name is
 * rendered alongside the case number in an element matched by
 * [class*="case"]. Observed textContent format:
 *
 *   "25STCV36868 THE FOUNTAIN GROUP, LP, et al. vs POK SUK KWON ReactDOM.render(...)"
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

function movantNormName(s) {
  return (s || '').toUpperCase().replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

const MOVANT_STOPWORDS = new Set(['the', 'of', 'for', 'and', 'to', 'a', 'an', 'on', 'in', 're', 'with', 'by']);

function movantSigTokens(s) {
  return (s || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim()
    .split(' ').filter(t => t.length > 1 && !MOVANT_STOPWORDS.has(t));
}

function movantTokenHit(a, b) {
  if (a === b) return true;
  if (a.length >= 4 && b.length >= 4 && (b.startsWith(a) || a.startsWith(b))) return true;
  return false;
}

function movantMatchScore(motionType, docName) {
  const mt = movantSigTokens(motionType);
  const dn = movantSigTokens(docName);
  if (!mt.length || !dn.length) return 0;
  const uniq = new Set(mt);
  let hit = 0;
  for (const t of uniq) if (dn.some(d => movantTokenHit(t, d))) hit++;
  return hit / uniq.size;
}

// Is this document Name an actual moving paper (motion/demurrer/etc.), not a
// response, order, minute order, declaration, etc. that rides alongside it?
function isMovingPaper(name) {
  if (!name) return false;
  const n = name.trim();
  if (/^(Opposition|Reply|Response|Declaration|Proof of Service|Order\b|Minute Order|Notice\b|Brief|Request\b|Certificate|Summons|Appeal\b|Case Management|Ex Parte Proposed Order|Points and Authorities|Memorandum|Stipulation|Objection|Separate Statement)/i.test(n)) {
    return false;
  }
  return /^(Motion|Demurrer|Petition|Application|Ex Parte Application|Anti-SLAPP|Special Motion|Amended Motion|Renewed Motion|Cross-?Motion)/i.test(n);
}

function bestFilingMatch(motionType, filings) {
  let best = null, bestScore = 0;
  for (const f of filings) {
    if (!isMovingPaper(f.name)) continue;
    const s = movantMatchScore(motionType, f.name);
    if (s > bestScore) { bestScore = s; best = f; }
  }
  return bestScore >= 0.5 ? best : null;
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
    // trailing parenthetical (e.g. "Kevin Singer (Non-Party)" -> "Kevin Singer").
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
    byName.set(nn, role);
    if (!byRole.has(role)) byRole.set(role, new Set());
    byRole.get(role).add(nn);
  }

  return { byName, byRole };
}

function formatMovant(parties, truncated, roster) {
  const groups = new Map(); // role -> [display names]
  for (const p of parties) {
    if (!p.name || /^clerk$/i.test(p.name)) continue;
    const role = roster.byName.get(movantNormName(p.name)) || p.role || '';
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
  try {
    const links = document.querySelectorAll('a[href*="/ecourt/ecms/case"]');
    for (const a of links) {
      if ((a.textContent || '').trim().toLowerCase() === label && a.href) return a.href;
    }
  } catch (_) {}
  return null;
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

try {
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'sync' && changes.excludedTerms) {
      EXCLUDED_TERMS_CACHE = Array.isArray(changes.excludedTerms.newValue) && changes.excludedTerms.newValue.length
        ? changes.excludedTerms.newValue : DEFAULT_EXCLUDED_TERMS;
    }
  });
} catch (_) {}

// Case-insensitive substring match against the exclusion terms, mirroring the
// agenda cleaner's isExcluded().
function isHearingExcluded(text) {
  if (!text) return false;
  const lower = text.toLowerCase();
  const terms = EXCLUDED_TERMS_CACHE || DEFAULT_EXCLUDED_TERMS;
  return terms.some(t => t && lower.includes(t));
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
      rows.push({ type: stripEventId(name), date: dm[0], when });
    }
  }

  const startOfToday = new Date(); startOfToday.setHours(0, 0, 0, 0);
  const seen = new Set();
  return rows
    .filter(h => h.when >= startOfToday)
    .filter(h => { const k = h.type + '@' + h.date; if (seen.has(k)) return false; seen.add(k); return true; })
    .sort((a, b) => a.when - b.when);
}

// Resolves the effective hearing for Export: the Next event unless it's
// excluded, in which case the soonest future scheduled non-excluded hearing
// from the Hearings tab. Returns { motionType, hearingDate, hearingType }.
async function resolveEffectiveHearing(root) {
  root = root || document;
  const nextType = parseHearingType(root);
  const base = {
    motionType: parseMotionType(root),
    hearingDate: parseHearingDate(root),
    hearingType: nextType,
  };
  try {
    await loadExcludedTerms();
    if (!nextType || !isHearingExcluded(nextType)) return base;

    const url = getHearingsUrl();
    if (!url) return base;
    const doc = await fetchCaseDoc(url);
    if (!doc) return base;

    const hearings = parseFutureHearings(doc).filter(h => !isHearingExcluded(h.type));
    if (!hearings.length) return base;

    // Take the soonest non-excluded hearing. When several fall on that same
    // date, prefer a Demurrer (with Motion to Strike) over a standalone Motion
    // to Strike — a demurrer + motion to strike is filed together but shows as
    // two hearing entries, and the demurrer is the one we want.
    const soonestDate = hearings[0].date;
    const sameDay = hearings.filter(h => h.date === soonestDate);
    const pick = sameDay.find(h => /demurrer/i.test(h.type)) || sameDay[0];

    console.log('[LACourt] Next hearing excluded (' + nextType +
      '); using Hearings-tab pick:', pick.type, pick.date);
    return {
      motionType: stripHearingOnPrefix(pick.type),
      hearingDate: pick.date,
      hearingType: pick.type,
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

// Async: resolves to the Movant string, or '' on any failure / no match.
// `partiesRoot` is the document the party roster is read from (the live page
// when on Parties, or a fetched Parties document otherwise).
async function computeMovant(motionType, partiesRoot) {
  try {
    if (!motionType) return '';
    const url = getDocumentsUrl();
    if (!url) return '';
    const doc = await fetchCaseDoc(url);
    if (!doc) return '';
    const filings = parseDocumentsFilingsFrom(doc);
    if (!filings.length) return '';
    const best = bestFilingMatch(motionType, filings);
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
  const ctx = getFillFormContext(partiesRoot, hearing);
  return ctx ? { ctx, partiesRoot } : null;
}

/* ------------------------------------------------------------------ */
/* Documents button: open the documents relevant to the motion         */
/* ------------------------------------------------------------------ */
//
// Identifies and opens (as background tabs) the documents relevant to the
// selected motion, all sourced from the Documents tab (deduped by docId):
//   - the operative complaint + cross-complaint (latest, not fictitious-name
//     amendments)
//   - the moving paper + anything the moving party filed the same day
//   - documents the Hearings tab lists for that motion
//   - one upcoming hearing  -> everything filed after the motion
//   - multiple hearings     -> documents after the motion whose title shares a
//     meaningful word with the motion type, plus each Opposition/Reply and its
//     same-day co-filings.

const DOC_STOP = new Set(['motion','opposition','reply','response','notice','declaration',
  'memorandum','points','authorities','support','order','proposed','hearing','plaintiff',
  'defendant','plaintiffs','defendants','exhibit','proof','service','with','from','that',
  'this','case','court','filed','amended']);

function docSigTokens(s) {
  return (s || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/)
    .filter(t => t.length >= 4 && !DOC_STOP.has(t));
}
function docWordOverlap(name, motionType) {
  const a = new Set(docSigTokens(motionType));
  if (!a.size) return false;
  return docSigTokens(name).some(t => a.has(t));
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
  return /^(?:(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|\d+(?:st|nd|rd|th))\s+)?amended\s+complaint\b/i.test(n)
      || /^complaint\b/i.test(n);
}
function isCrossComplaintDoc(name) {
  const n = (name || '').trim();
  if (/^amendment to /i.test(n)) return false;
  return /^(?:(?:first|second|third|fourth|fifth|\d+(?:st|nd|rd|th))\s+)?(?:amended\s+)?cross-?complaint\b/i.test(n);
}
// A petition is another kind of initial pleading (probate, family, writ, etc.),
// used as the operative pleading only when the case has no complaint. Word-
// boundary match, so "Petitioner" in other filings doesn't count.
function isPetitionDoc(name) {
  const n = (name || '').trim();
  if (/^amendment to /i.test(n)) return false;
  if (/fictitious|incorrect\s+name/i.test(n)) return false;
  return /\bpetition\b/i.test(n);
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

function computeRelevantDocuments(docs, motionType, hearingDocBlob, singleHearing) {
  const rel = new Map();
  const add = d => { if (d && d.docId && d.openUrl) rel.set(d.docId, d); };

  // Operative complaint + cross-complaint. When the case has no complaint at
  // all, fall back to the operative petition (another initial pleading).
  let initialPleading = latestDoc(docs.filter(d => isComplaintDoc(d.name)));
  if (!initialPleading) initialPleading = latestDoc(docs.filter(d => isPetitionDoc(d.name)));
  add(initialPleading);
  add(latestDoc(docs.filter(d => isCrossComplaintDoc(d.name))));

  const motionDoc = bestFilingMatch(motionType, docs);
  if (motionDoc) {
    add(motionDoc);
    const mov = docPartyNames(motionDoc.filedBy), mw = motionDoc.when;

    // Same-day filings by the moving party (incl. just before the motion).
    for (const d of docs) if (sameCalendarDay(d.when, mw) && docSharesParty(docPartyNames(d.filedBy), mov)) add(d);

    // Documents the Hearings tab lists for this motion (substring containment).
    if (hearingDocBlob) {
      const blob = movantNormName(hearingDocBlob);
      for (const d of docs) { const nn = movantNormName(d.name); if (nn && nn.length >= 6 && blob.indexOf(nn) !== -1) add(d); }
    }

    if (singleHearing) {
      // One upcoming hearing: everything after the motion is fair game.
      for (const d of docs) if (d.when && mw && d.when > mw) add(d);
    } else {
      // Multiple hearings: match by shared words + Opposition/Reply co-filings.
      for (const d of docs) if (d.when && mw && d.when >= mw && docWordOverlap(d.name, motionType)) add(d);
      const after = docs.filter(d => d.when && mw && d.when >= mw);
      for (const opp of after) if (/\bopposition\b/i.test(opp.name) && docWordOverlap(opp.name, motionType)) {
        add(opp); const P = docPartyNames(opp.filedBy);
        for (const d of docs) if (sameCalendarDay(d.when, opp.when) && docSharesParty(docPartyNames(d.filedBy), P)) add(d);
      }
      for (const rep of after) if (/\breply\b/i.test(rep.name) && docWordOverlap(rep.name, motionType)) {
        add(rep); const P = docPartyNames(rep.filedBy);
        for (const d of docs) if (sameCalendarDay(d.when, rep.when) && docSharesParty(docPartyNames(d.filedBy), P)) add(d);
      }
    }
  }

  // Proof-of-service documents are noise for most motions. Keep them only when
  // the motion is one where service itself tends to be at issue.
  const mtl = (motionType || '').toLowerCase();
  const keepProofOfService = POS_KEEP_TERMS.some(t => mtl.indexOf(t) !== -1);
  if (!keepProofOfService) {
    for (const [id, d] of rel) if (/proof of service/i.test(d.name || '')) rel.delete(id);
  }

  return Array.from(rel.values());
}

// Motion-type terms for which "Proof of Service" documents stay relevant.
// For any other motion, proof-of-service filings are excluded.
const POS_KEEP_TERMS = [
  'paga', 'settlement', 'transfer', 'order to show cause re',
  'default', 'quash', 'set aside', 'vacate',
];

function absoluteDocUrl(u) { try { return new URL(u, location.origin).href; } catch (_) { return u; } }

// Parses openable document rows from a Documents-tab document or paged fragment:
// each openInNewWindow anchor yields { docId, openUrl, name, dateStr, when, filedBy }.
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
    let filedBy = ''; const tr = a.closest('tr');
    if (tr) { const cells = Array.from(tr.children); if (cells.length > 6) filedBy = (cells[6].textContent || '').replace(/\s+/g, ' ').trim(); }
    rows.push({ docId, openUrl: absoluteDocUrl(url), name, dateStr, when: dateStr ? parseHearingDateTime(dateStr) : null, filedBy });
  }
  return rows;
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
    const offsets = [];
    for (let o = 50; o < count && o <= 2000; o += 50) offsets.push(o);
    const results = await Promise.all(offsets.map(o => postPanelPage(o, pageData)));
    for (const frag of results) {
      if (!frag) continue;
      for (const r of parseDocRows(frag)) if (!seen.has(r.docId)) { seen.add(r.docId); rows.push(r); }
    }
  }
  return rows;
}

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

// Orchestrates: resolve the motion, fetch all documents + hearings, compute the
// relevant set. Resolves to { relevant, motionType, docCount, singleHearing }.
async function getRelevantDocuments() {
  const hearing = await resolveEffectiveHearing(document);
  const motionType = hearing && hearing.motionType;
  if (!motionType) return { relevant: [], reason: 'no-motion' };

  const docs = await getAllDocumentsCached();
  if (!docs.length) return { relevant: [], reason: 'no-documents' };

  const hearingsUrl = getHearingsUrl();
  const hearingsDoc = hearingsUrl ? await fetchCaseDoc(hearingsUrl) : null;
  const singleHearing = hearingsDoc ? parseFutureHearings(hearingsDoc).length <= 1 : true;
  const hearingDocBlob = hearingsDoc ? findHearingDocBlob(hearingsDoc, motionType) : '';

  const relevant = computeRelevantDocuments(docs, motionType, hearingDocBlob, singleHearing);
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

// Note: form URLs are defined at the top of the IIFE (REGULAR_FORM_URL
// and OSC_FORM_URL); getFillFormContext() picks the right one based on
// the case's next event.

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

            // Success — show confirmation and reset.
            setLabel(openedLabel);
            btn.style.opacity = '1';
            setTimeout(() => {
              btn.disabled = false;
              setLabel('Export');
            }, 2000);
          }
        );
      };

      // For the Order Template popup, auto-detect the Movant and wait until the
      // parsed fields are stored before opening the popup window. OSC cases open
      // the real form.
      if (ctx.isOrderTemplate) {
        computeMovant(ctx.data.labeled.motionType, result.partiesRoot).then(movant => {
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
      const urls = opened.map(d => d.openUrl);

      // Debug tracking: record the documents the button opened.
      try {
        chrome.runtime.sendMessage({
          type: 'recordOpenedDocs',
          source: 'button',
          caseNumber: parseCaseNumber(),
          docs: opened.map(d => ({ docId: d.docId, name: d.name })),
        }, () => void chrome.runtime.lastError);
      } catch (_) {}

      chrome.runtime.sendMessage({ type: 'openDocsBackground', urls }, () => {
        void chrome.runtime.lastError;
        reset('Opened ' + urls.length + (capped ? '+' : ''));
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
const NOTICE_OF_ENTRY_RE = /notice of (entry|ruling)/i;
const ENTRY_OF_JUDGMENT_RE = /\bjudgment\b/i;
function isTriggerBasedMotion(motionType) {
  return /reconsideration|renewed?\s+motion|\b1008\b|new\s+trial|\bjnov\b|judgment\s+notwithstanding|vacate\s+(the\s+)?judgment/i.test(motionType || '');
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
  const out = { noticeOfEntryDate: '', noticeOfEntryDoc: '', entryOfJudgmentDate: '', entryOfJudgmentDoc: '' };
  try {
    const docsUrl = getDocumentsUrl();
    if (!docsUrl) return out;
    const docs = await fetchAllDocuments(docsUrl);
    if (!docs || !docs.length) return out;
    const cutoff = hearingDateStr ? parseHearingDateTime(hearingDateStr) : null;
    const noe = latestDocOnOrBefore(docs.filter(d => d.name && d.when && NOTICE_OF_ENTRY_RE.test(d.name)), cutoff);
    if (noe) { out.noticeOfEntryDate = noe.dateStr; out.noticeOfEntryDoc = noe.name; }
    const eoj = latestDocOnOrBefore(docs.filter(d => d.name && d.when && isEntryOfJudgmentDoc(d.name)), cutoff);
    if (eoj) { out.entryOfJudgmentDate = eoj.dateStr; out.entryOfJudgmentDoc = eoj.name; }
    return out;
  } catch (_) { return out; }
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
    setDlLabel('Detecting…');
    btn.style.opacity = '0.7';
    try {
      const hearing = await resolveEffectiveHearing(document);
      const motionType = (hearing && hearing.motionType) || '';

      // For trigger-based motions, detect the statutory trigger dates (notice of
      // entry, and entry of judgment for the new-trial 180-day cap) from the
      // case documents. Skipped for ordinary motions so the common path is fast.
      let trig = { noticeOfEntryDate: '', noticeOfEntryDoc: '', entryOfJudgmentDate: '', entryOfJudgmentDoc: '' };
      if (isTriggerBasedMotion(motionType)) {
        trig = await detectTriggerDates(hearing && hearing.hearingDate);
      }

      const payload = {
        motionType,
        hearingDate: (hearing && hearing.hearingDate) || '',
        hearingType: (hearing && hearing.hearingType) || '',
        caseNumber: parseCaseNumber() || '',
        noticeOfEntryDate: trig.noticeOfEntryDate,
        noticeOfEntryDoc: trig.noticeOfEntryDoc,
        entryOfJudgmentDate: trig.entryOfJudgmentDate,
        entryOfJudgmentDoc: trig.entryOfJudgmentDoc,
        createdAt: Date.now(),
      };
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
    '#__lacourt_deadline_btn__ .lac-btn-text{margin-left:6px}' +
    '#__lacourt_fill_btn__[data-collapsed="1"] .lac-btn-text,' +
    '#__lacourt_docs_btn__[data-collapsed="1"] .lac-btn-text,' +
    '#__lacourt_deadline_btn__[data-collapsed="1"] .lac-btn-text{display:none}' +
    '#__lacourt_fill_btn__[data-collapsed="1"],#__lacourt_docs_btn__[data-collapsed="1"],' +
    '#__lacourt_deadline_btn__[data-collapsed="1"]' +
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
// (Export rightmost, then Documents, then Deadlines) in whichever state.
function updateButtonCollapse() {
  const ids = ['__lacourt_fill_btn__', '__lacourt_docs_btn__', '__lacourt_deadline_btn__'];
  const btns = ids.map(id => document.getElementById(id)).filter(Boolean);
  if (!btns.length) return;

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
// isn't due yet it shows in the neutral colour.

function fmtShortDate(d) {
  return (!d || isNaN(d)) ? '' : d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: '2-digit' });
}
function findNextHeaderSpan() {
  const spans = document.querySelectorAll('span[title]');
  for (const span of spans) {
    const hay = ((span.getAttribute('title') || '') + ' ' + (span.textContent || '')).replace(/\s+/g, ' ');
    if (/\bnext\b/i.test(hay) && /\d{1,2}\/\d{1,2}\/\d{4}/.test(hay)) return span;
  }
  return null;
}
function dayMs(d) { return (d && !isNaN(d)) ? new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() : null; }

const NEXT_DL_GAP = '<span style="display:inline-block;width:20px"></span>';
// Computed once per page: { skip } or { motionType, cat, motionDue, oppDue, replyDue }.
let __nextDlComputed = null;
// Filing status once the Documents fetch resolves: { filedKnown, motion, opp, reply }.
let __nextDlFiled = null;
let __nextDlFetchStarted = false;

function computeDueDates() {
  const D = window.LACourtDeadlines;
  if (!D) return null;
  // Only "Hearing on <motion>" events are §1005/§437c-briefed. Non-motion
  // events (CMC, OSC, trial, status conference) have no "Hearing on" prefix.
  const motionType = parseMotionType(document);
  if (!motionType) return null; // header not rendered yet (or not a motion)
  const cat = D.classifyMotion(motionType);
  if (cat !== 'standard' && cat !== 'msj') return { skip: true }; // new trial / recon aren't hearing-based
  const hearing = parseHearingDateTime(parseHearingDate(document));
  if (!hearing) return null;
  return {
    skip: false, motionType, cat,
    motionDue: cat === 'msj' ? D.msjMotion(hearing, 'electronic') : D.stdMotion(hearing, 'electronic'),
    oppDue:    cat === 'msj' ? D.msjOpp(hearing) : D.stdOpp(hearing),
    replyDue:  cat === 'msj' ? D.msjReply(hearing) : D.stdReply(hearing),
  };
}

// Colour: green = filed on time; red = overdue and not timely filed; neutral =
// not yet due (or filing status not yet known / unavailable).
function nextDlColor(due, filed) {
  const dd = dayMs(due);
  if (dd == null) return '#0a6e6e';
  if (__nextDlFiled && __nextDlFiled.filedKnown) {
    const fm = dayMs(filed);
    if (fm != null && fm <= dd) return '#1a6b3a';
  }
  return dd < dayMs(new Date()) ? '#c0392b' : '#0a6e6e';
}

function nextDlHtml() {
  const c = __nextDlComputed;
  const f = __nextDlFiled || {};
  const item = (label, key, due) =>
    `<span style="color:${nextDlColor(due, f[key])}">${label} ${fmtShortDate(due)}</span>`;
  return item('Motion Due', 'motion', c.motionDue) + NEXT_DL_GAP +
    item('Opposition Due', 'opp', c.oppDue) + NEXT_DL_GAP +
    item('Reply Due', 'reply', c.replyDue);
}

// Idempotent: injects the widget if missing, else refreshes its colours. Re-finds
// the header each time so it survives e-court's React re-renders.
function injectNextDeadlines() {
  if (!__nextDlComputed || __nextDlComputed.skip) return;
  const span = findNextHeaderSpan();
  if (!span || !span.parentNode) return;
  let el = span.parentNode.querySelector('.__lacourt_next_dl__');
  if (el) { el.innerHTML = nextDlHtml(); return; }
  el = document.createElement('span');
  el.className = '__lacourt_next_dl__';
  el.setAttribute('style', 'margin-left:22px;font-weight:600;white-space:nowrap;font-family:inherit;');
  el.innerHTML = nextDlHtml();
  span.parentNode.insertBefore(el, span.nextSibling);
}

// Fetch the case Documents once and recolour by whether each paper was filed on
// time. Best-effort — the dates are already shown regardless.
async function fetchNextDeadlineFilings() {
  if (__nextDlFetchStarted || !__nextDlComputed || __nextDlComputed.skip) return;
  __nextDlFetchStarted = true;
  const c = __nextDlComputed;
  const filed = { filedKnown: false, motion: null, opp: null, reply: null };
  try {
    {
      const docs = await getAllDocumentsCached();
      if (docs && docs.length) {
        filed.filedKnown = true;
        const md = bestFilingMatch(c.motionType, docs);
        filed.motion = md ? md.when : null;
        const mw = md ? md.when : null;
        const after = docs.filter(d => d.when && (!mw || d.when >= mw));
        const earliest = list => list.slice().sort((a, b) => a.when - b.when)[0] || null;
        const o = earliest(after.filter(d => /\bopposition\b/i.test(d.name) && docWordOverlap(d.name, c.motionType)));
        const r = earliest(after.filter(d => /\breply\b/i.test(d.name) && docWordOverlap(d.name, c.motionType)));
        filed.opp = o ? o.when : null;
        filed.reply = r ? r.when : null;
      }
    }
  } catch (_) { /* keep date-only colours */ }
  __nextDlFiled = filed;
  injectNextDeadlines(); // recolour
}

function renderNextHeaderDeadlines() {
  try {
    if (!window.LACourtDeadlines) return;
    if (!__nextDlComputed) {
      const c = computeDueDates();
      if (!c) return; // header not ready — a later poll/observer will retry
      __nextDlComputed = c;
      if (c.skip) return;
      injectNextDeadlines();        // paint dates immediately (date-only colours)
      fetchNextDeadlineFilings();   // then refine with filing status
      return;
    }
    injectNextDeadlines(); // keep present + coloured across re-renders
  } catch (_) { /* best-effort UI */ }
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
    requestAnimationFrame(() => { pending = false; try { renderNextHeaderDeadlines(); } catch (_) {} });
  });
  try { obs.observe(document.body, { childList: true, subtree: true }); __nextDlObserver = obs; } catch (_) {}
}

function setupFillFormButton() {
  // The parties table loads after initial page render. Try once on DOMContentLoaded
  // and once on full load, then poll briefly until it appears (cap at ~10s).
  const tryRender = () => {
    ensureButtonStyles();
    renderFillFormButton();
    renderDocumentsButton();
    renderDeadlineButton();
    renderNextHeaderDeadlines();
    observeNextHeader();
    scheduleButtonCollapse();
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
  const prefetchDocs = () => { try { getRelevantDocumentsCached(); } catch (_) {} };
  if (document.readyState === 'complete') setTimeout(prefetchDocs, 1000);
  else window.addEventListener('load', () => setTimeout(prefetchDocs, 1000), { once: true });
}

setupFillFormButton();
})();

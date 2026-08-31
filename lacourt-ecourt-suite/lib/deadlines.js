/**
 * Shared California motion-deadline engine.
 *
 * Used by BOTH the case-page content script (inline Opposition/Reply on the
 * "Next" header) and the Deadline Calculator page, so the two never drift.
 * Attaches to the global as `LACourtDeadlines` — visible across content-script
 * files in the isolated world and to the calculator page.
 *
 * Authorities: CCP §§ 1005, 437c, 659, 659a, 663a, 1008, 1013, 1010.6; CRC
 * 3.1700 (costs); court holidays per CCP § 135 / Gov. Code § 6700 / CRC 1.11.
 */
(function () {
  'use strict';

  // ── HOLIDAYS ──────────────────────────────────────────────────────────────
  // California judicial holidays. Columbus Day (2nd Monday of October) is
  // expressly NOT a judicial holiday.
  const holidayCache = {};
  function getHolidays(year) {
    if (holidayCache[year]) return holidayCache[year];
    const h = new Set();
    const obs = d => {
      const day = d.getDay();
      if (day === 6) { const f = new Date(d); f.setDate(f.getDate() - 1); return f; } // Sat → Fri
      if (day === 0) { const m = new Date(d); m.setDate(m.getDate() + 1); return m; } // Sun → Mon
      return d;
    };
    const fixed = (mo, da) => obs(new Date(year, mo, da));
    const nth = (mo, wd, n) => {
      let d = new Date(year, mo, 1), c = 0;
      while (d.getMonth() === mo) { if (d.getDay() === wd && ++c === n) return d; d.setDate(d.getDate() + 1); }
    };
    const last = (mo, wd) => {
      let d = new Date(year, mo + 1, 0);
      while (d.getDay() !== wd) d.setDate(d.getDate() - 1);
      return d;
    };
    const key = d => d ? `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}` : null;
    const add = d => { if (d) h.add(key(d)); };

    add(fixed(0, 1));      // New Year's Day
    add(nth(0, 1, 3));     // MLK Jr. Day — 3rd Monday Jan
    add(fixed(1, 12));     // Lincoln's Birthday — Feb 12
    add(nth(1, 1, 3));     // Presidents' Day — 3rd Monday Feb
    add(fixed(2, 31));     // Farmworkers (Cesar Chavez) Day — Mar 31
    add(last(4, 1));       // Memorial Day — last Monday May
    add(fixed(5, 19));     // Juneteenth — Jun 19
    add(fixed(6, 4));      // Independence Day — Jul 4
    add(nth(8, 1, 1));     // Labor Day — 1st Monday Sep
    add(nth(8, 5, 4));     // Native American Day — 4th Friday Sep
    add(fixed(10, 11));    // Veterans Day — Nov 11
    const tg = nth(10, 4, 4); add(tg);                                 // Thanksgiving
    if (tg) { const da = new Date(tg); da.setDate(da.getDate() + 1); add(da); } // Day after
    add(fixed(11, 25));    // Christmas — Dec 25

    // Next year's New Year's Day observed on Dec 31 of THIS year (Sat → Fri).
    const nyNext = obs(new Date(year + 1, 0, 1));
    if (nyNext.getFullYear() === year) add(nyNext);

    holidayCache[year] = h;
    return h;
  }
  function isCourtDay(d) {
    const dow = d.getDay();
    if (dow === 0 || dow === 6) return false;
    const h = getHolidays(d.getFullYear());
    return !h.has(`${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`);
  }
  function nextCourtDay(d) { const r = new Date(d); while (!isCourtDay(r)) r.setDate(r.getDate() + 1); return r; }
  function prevCourtDay(d) { const r = new Date(d); while (!isCourtDay(r)) r.setDate(r.getDate() - 1); return r; }
  function addCD(d, n) { // court days
    const r = new Date(d), step = n >= 0 ? 1 : -1;
    let rem = Math.abs(n);
    while (rem > 0) { r.setDate(r.getDate() + step); if (isCourtDay(r)) rem--; }
    return r;
  }
  function addCAL(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; } // calendar days

  // ── DEADLINE LOGIC ────────────────────────────────────────────────────────
  // Standard noticed motion — CCP § 1005(b): 16 court days + service addition.
  function stdMotion(hearing, svc) {
    let d = addCD(hearing, -16);
    if (svc === 'electronic') d = addCD(d, -2);        // +2 court days
    else if (svc === 'mail_ca') d = addCAL(d, -5);
    else if (svc === 'mail_state') d = addCAL(d, -10);
    else if (svc === 'mail_conf') d = addCAL(d, -12);  // +12 (Safe at Home)
    else if (svc === 'mail_intl') d = addCAL(d, -20);
    else if (svc === 'fax') d = addCAL(d, -2);          // +2 calendar
    return prevCourtDay(d);
  }
  // MSJ / MSA — CCP § 437c(a)(2): 81 days + service addition. No 12-day Safe at
  // Home tier; fax/express/overnight are +2 COURT days (a § 437c vs § 1005
  // asymmetry).
  function msjMotion(hearing, svc) {
    let d = addCAL(hearing, -81);
    if (svc === 'electronic') d = addCD(d, -2);
    else if (svc === 'mail_ca') d = addCAL(d, -5);
    else if (svc === 'mail_state') d = addCAL(d, -10);
    else if (svc === 'mail_conf') d = addCAL(d, -5);
    else if (svc === 'mail_intl') d = addCAL(d, -20);
    else if (svc === 'fax') d = addCD(d, -2);
    return prevCourtDay(d);
  }
  function stdOpp(hearing)   { return prevCourtDay(addCD(hearing, -9));  } // § 1005(b)
  function msjOpp(hearing)   { return prevCourtDay(addCAL(hearing, -20)); } // § 437c(b)(2)
  function stdReply(hearing) { return prevCourtDay(addCD(hearing, -5));  } // § 1005(b)
  function msjReply(hearing) { return prevCourtDay(addCAL(hearing, -11)); } // § 437c(b)(4)
  // Unlawful detainer MSJ — CCP § 1170.7: made at any time after the answer, on
  // FIVE days' notice; CRC 3.1351(a) has that notice given in compliance with
  // §§ 1013 and 1170.7, so the ordinary service extensions ride on the 5 days.
  function udMsjMotion(hearing, svc) {
    let d = addCAL(hearing, -5);
    if (svc === 'electronic') d = addCD(d, -2);
    else if (svc === 'mail_ca') d = addCAL(d, -5);
    else if (svc === 'mail_state') d = addCAL(d, -10);
    else if (svc === 'mail_conf') d = addCAL(d, -12);
    else if (svc === 'mail_intl') d = addCAL(d, -20);
    else if (svc === 'fax') d = addCAL(d, -2);
    return prevCourtDay(d);
  }
  // CRC 3.1351(b)-(c): opposition may be made orally at the hearing; a WRITTEN
  // opposition to be considered in advance is filed and served on or before the
  // court day before the hearing.
  function udMsjOpp(hearing)   { return prevCourtDay(addCAL(hearing, -1)); }
  // No advance written reply schedule — the reply may be made orally at the
  // hearing (CRC 3.1351(b)), so the slot dates to the hearing day itself.
  function udMsjReply(hearing) { return prevCourtDay(hearing); }
  function newTrialDL(notice){ return nextCourtDay(addCAL(notice, 15));  } // § 659(a)(2)

  // ── ATTORNEY FEES (CRC 3.1702(b)) ─────────────────────────────────────────
  // A notice of motion claiming fees for services through rendition of judgment
  // in the trial court is served and filed within the time for filing a NOTICE
  // OF APPEAL (rule 3.1702(b)(1), incorporating rules 8.104 and 8.108). Under
  // rule 8.104(a)(1) that is the earliest of: 60 days after the clerk serves a
  // "Notice of Entry" or a filed-endorsed copy of the judgment; 60 days after a
  // party serves either with proof of service; or 180 days after entry.
  //
  // Deliberately NO service extension. The period is measured by the time to
  // appeal, and both extension statutes carve that out expressly — § 1013(a)
  // for mail and fax, § 1010.6(a)(3)(B) for electronic service ("the extension
  // shall not apply to extend the time for filing ... a notice of appeal").
  // This is the difference from the costs memorandum, which does carry them.
  //
  // Rule 8.108 can extend the period when a valid post-trial motion is pending,
  // and rule 3.1702(b)(2) lets the parties stipulate to more time. Neither is
  // visible from the docket, so the date here is the unextended one.
  function feesDL(triggerService) { return nextCourtDay(addCAL(triggerService, 60)); }
  function feesOuterDL(entryOfJudgment) { return nextCourtDay(addCAL(entryOfJudgment, 180)); }

  // Service extension for a period that runs FORWARD from service of a document:
  // § 1013(a) adds calendar days for mail, § 1010.6(a)(3)(B) adds two court days
  // for electronic service. Personal service adds nothing.
  function addServiceExtension(d, svc) {
    if (svc === 'electronic') return addCD(d, 2);
    if (svc === 'mail_ca') return addCAL(d, 5);
    if (svc === 'mail_state') return addCAL(d, 10);
    if (svc === 'mail_conf') return addCAL(d, 12);
    if (svc === 'mail_intl') return addCAL(d, 20);
    if (svc === 'fax') return addCD(d, 2);
    return d;
  }

  // § 1008(a): 10 days after service of notice of entry; §§ 1013 / 1010.6 apply.
  function reconDL(notice, svc) {
    return nextCourtDay(addServiceExtension(addCAL(notice, 10), svc));
  }

  // ── COSTS (CRC 3.1700) ────────────────────────────────────────────────────
  // A prevailing party's memorandum of costs is due on the FIRST of two dates:
  // 15 days after service of the notice of entry of judgment or dismissal
  // (rule 3.1700(a)(1)), or 180 days after entry of judgment. Only the 15-day
  // period runs from service, so only it carries the §§ 1013 / 1010.6
  // extensions; the 180-day outer limit runs from entry and does not.
  function costsMemoDL(noticeOfEntry, svc) {
    return nextCourtDay(addServiceExtension(addCAL(noticeOfEntry, 15), svc));
  }
  function costsMemoOuterDL(entryOfJudgment) {
    return nextCourtDay(addCAL(entryOfJudgment, 180));
  }
  // Rule 3.1700(b)(1): a notice of motion to strike or tax costs is served and
  // filed 15 days after service of the cost memorandum, extended for mail
  // (§ 1013) and for electronic service (§ 1010.6(a)(3)(B) — two court days).
  function costsTaxDL(memoServed, svc) {
    return nextCourtDay(addServiceExtension(addCAL(memoServed, 15), svc));
  }

  // ── CLASSIFICATION ────────────────────────────────────────────────────────
  // Map an e-court motion-type string to a rule bucket. Most motions use the
  // standard § 1005 schedule; only these carry their own counting rules.

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

  function classifyMotion(mt) {
    const s = stripAncillaryMotionReference(mt).toLowerCase();
    if (SEAL_MOTION_RE.test(s)) return 'standard';
    if (/summary\s+judgment|summary\s+adjudication|\bmsj\b|\bmsa\b/.test(s)) return 'msj';
    if (/new\s+trial|\bjnov\b|judgment\s+notwithstanding|vacate\s+(the\s+)?judgment/.test(s)) return 'new_trial';
    if (/reconsideration|renewed?\s+motion|\bccp?\s*1008\b|\b1008\b/.test(s)) return 'recon';
    // Fees before costs: "motion to strike or tax costs" is a costs motion, but
    // "attorney fees and costs" is a fee motion, and the fee test is narrower.
    if (/attorney'?s?\s+fees|\battorney\s+fee\b|\bfees\s+and\s+costs\b|\b3\.1702\b/.test(s)) return 'fees';
    if (/\b(?:strike|tax|taxing)\s+(?:of\s+)?costs\b|\bcosts\b.*\b(?:strike|tax)\b|memorandum\s+of\s+costs|\b3\.1700\b/.test(s)) return 'costs';
    return 'standard';
  }

  // Parse "yyyy-mm-dd" (date input) or "m/d/yyyy" (e-court) as a LOCAL date.
  function parseDateFlexible(s) {
    if (!s) return null;
    let m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) return new Date(+m[3], +m[1] - 1, +m[2]);
    const d = new Date(s);
    return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  const api = {
    getHolidays, isCourtDay, nextCourtDay, prevCourtDay, addCD, addCAL,
    stdMotion, msjMotion, stdOpp, msjOpp, stdReply, msjReply,
    udMsjMotion, udMsjOpp, udMsjReply, newTrialDL, reconDL,
    addServiceExtension, costsMemoDL, costsMemoOuterDL, costsTaxDL, feesDL, feesOuterDL,
    classifyMotion, stripAncillaryMotionReference, parseDateFlexible,
  };
  (typeof window !== 'undefined' ? window : globalThis).LACourtDeadlines = api;
})();

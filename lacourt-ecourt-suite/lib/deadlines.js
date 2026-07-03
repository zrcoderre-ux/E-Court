/**
 * Shared California motion-deadline engine.
 *
 * Used by BOTH the case-page content script (inline Opposition/Reply on the
 * "Next" header) and the Deadline Calculator page, so the two never drift.
 * Attaches to the global as `LACourtDeadlines` — visible across content-script
 * files in the isolated world and to the calculator page.
 *
 * Authorities: CCP §§ 1005, 437c, 659, 659a, 663a, 1008, 1013, 1010.6; court
 * holidays per CCP § 135 / Gov. Code § 6700 / CRC 1.11.
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
  function newTrialDL(notice){ return nextCourtDay(addCAL(notice, 15));  } // § 659(a)(2)
  // § 1008(a): 10 days after service of notice of entry; §§ 1013 / 1010.6 apply.
  function reconDL(notice, svc) {
    let d = addCAL(notice, 10);
    if (svc === 'electronic') d = addCD(d, 2);
    else if (svc === 'mail_ca') d = addCAL(d, 5);
    else if (svc === 'mail_state') d = addCAL(d, 10);
    else if (svc === 'mail_conf') d = addCAL(d, 12);
    else if (svc === 'mail_intl') d = addCAL(d, 20);
    else if (svc === 'fax') d = addCD(d, 2);
    return nextCourtDay(d);
  }

  // ── CLASSIFICATION ────────────────────────────────────────────────────────
  // Map an e-court motion-type string to a rule bucket. Most motions use the
  // standard § 1005 schedule; only these carry their own counting rules.
  function classifyMotion(mt) {
    const s = (mt || '').toLowerCase();
    if (/summary\s+judgment|summary\s+adjudication|\bmsj\b|\bmsa\b/.test(s)) return 'msj';
    if (/new\s+trial|\bjnov\b|judgment\s+notwithstanding|vacate\s+(the\s+)?judgment/.test(s)) return 'new_trial';
    if (/reconsideration|renewed?\s+motion|\bccp?\s*1008\b|\b1008\b/.test(s)) return 'recon';
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
    stdMotion, msjMotion, stdOpp, msjOpp, stdReply, msjReply, newTrialDL, reconDL,
    classifyMotion, parseDateFlexible,
  };
  (typeof window !== 'undefined' ? window : globalThis).LACourtDeadlines = api;
})();

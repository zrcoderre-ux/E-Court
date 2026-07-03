/**
 * LA Court E-Court Suite — Motion Deadline Calculator
 *
 * Opened in its own window by the "Deadlines" floating button on a case page.
 * The content script resolves the effective hearing (same detection the Export
 * and Documents buttons use) and stashes { motionType, hearingDate, caseNumber }
 * under chrome.storage.local `deadlineCalcData`; this page reads it on load to
 * pre-fill the trigger date and pre-select the matching motion rule.
 *
 * California authorities: CCP §§ 1005, 437c, 659, 659a, 663a, 1008, 1013,
 * 1010.6; court holidays per CCP § 135 / Gov. Code § 6700 / CRC 1.11.
 */

// ── STATE ──────────────────────────────────────────────────────────────────
let state = {
  baseDate: null,
  mode: 'A',
  expanded: false,
  motionType: 'standard',
  service: 'electronic',
  mailRegion: 'ca',
  detected: null, // { rawMotion, hearingDate, caseNumber, category }
};

// ── HOLIDAYS ───────────────────────────────────────────────────────────────
// California judicial holidays (CCP § 135, Gov. Code § 6700, CRC 1.11).
// Columbus Day (2nd Monday of October) is expressly NOT a judicial holiday.
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

  add(fixed(0, 1));      // New Year's Day — Jan 1
  add(nth(0, 1, 3));     // Dr. MLK Jr. Day — 3rd Monday Jan
  add(fixed(1, 12));     // Lincoln's Birthday — Feb 12
  add(nth(1, 1, 3));     // Presidents' Day — 3rd Monday Feb
  add(fixed(2, 31));     // Farmworkers (Cesar Chavez) Day — Mar 31
  add(last(4, 1));       // Memorial Day — last Monday May
  add(fixed(5, 19));     // Juneteenth — Jun 19
  add(fixed(6, 4));      // Independence Day — Jul 4
  add(nth(8, 1, 1));     // Labor Day — 1st Monday Sep
  add(nth(8, 5, 4));     // Native American Day — 4th Friday Sep
  add(fixed(10, 11));    // Veterans Day — Nov 11
  const tg = nth(10, 4, 4); add(tg); // Thanksgiving — 4th Thursday Nov
  if (tg) { const da = new Date(tg); da.setDate(da.getDate() + 1); add(da); } // Day after Thanksgiving
  add(fixed(11, 25));    // Christmas — Dec 25

  holidayCache[year] = h;
  return h;
}
function isCourtDay(d) {
  const dow = d.getDay();
  if (dow === 0 || dow === 6) return false;
  const h = getHolidays(d.getFullYear());
  return !h.has(`${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`);
}
function nextCourtDay(d) {
  const r = new Date(d);
  while (!isCourtDay(r)) r.setDate(r.getDate() + 1);
  return r;
}
function addCD(d, n) { // court days
  const r = new Date(d), step = n >= 0 ? 1 : -1;
  let rem = Math.abs(n);
  while (rem > 0) { r.setDate(r.getDate() + step); if (isCourtDay(r)) rem--; }
  return r;
}
function addCAL(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; } // calendar days
function prevCourtDay(d) { const r = new Date(d); while (!isCourtDay(r)) r.setDate(r.getDate() - 1); return r; }

// ── DEADLINE LOGIC ─────────────────────────────────────────────────────────
// Standard noticed motion — CCP § 1005(b). Moving papers 16 court days before
// the hearing, plus the service-method addition of § 1005(b) itself.
function stdMotion(hearing, svc) {
  let d = addCD(hearing, -16);
  if (svc === 'electronic') d = addCD(d, -2);        // § 1010.6(a)(3)(B): +2 court days
  else if (svc === 'mail_ca') d = addCAL(d, -5);     // § 1005(b): +5 calendar
  else if (svc === 'mail_state') d = addCAL(d, -10); // +10 calendar (US, outside CA)
  else if (svc === 'mail_conf') d = addCAL(d, -12);  // +12 calendar (Safe at Home)
  else if (svc === 'mail_intl') d = addCAL(d, -20);  // +20 calendar (outside US)
  else if (svc === 'fax') d = addCAL(d, -2);         // fax/express/overnight: +2 calendar
  return prevCourtDay(d);
}
// MSJ / MSA — CCP § 437c(a)(2): moving papers served 81 days before the hearing,
// plus service additions (mail +5/+10/+12/+20 calendar; fax/express/overnight
// +2 calendar; electronic +2 court days).
function msjMotion(hearing, svc) {
  let d = addCAL(hearing, -81);
  if (svc === 'electronic') d = addCD(d, -2);
  else if (svc === 'mail_ca') d = addCAL(d, -5);
  else if (svc === 'mail_state') d = addCAL(d, -10);
  else if (svc === 'mail_conf') d = addCAL(d, -12);
  else if (svc === 'mail_intl') d = addCAL(d, -20);
  else if (svc === 'fax') d = addCAL(d, -2);
  return prevCourtDay(d);
}
function stdOpp(hearing)   { return prevCourtDay(addCD(hearing, -9));  } // § 1005(b)
function msjOpp(hearing)   { return prevCourtDay(addCAL(hearing, -20)); } // § 437c(b)(2)
function stdReply(hearing) { return prevCourtDay(addCD(hearing, -5));  } // § 1005(b)
function msjReply(hearing) { return prevCourtDay(addCAL(hearing, -11)); } // § 437c(b)(4)
function newTrialDL(notice){ return nextCourtDay(addCAL(notice, 15));  } // § 659(a)(2)
function reconDL(notice)   { return nextCourtDay(addCAL(notice, 10));  } // § 1008(a)

// ── MOTION CLASSIFICATION ───────────────────────────────────────────────────
// Map the e-court motion-type string to one of the calculator's rule buckets.
// Most motions (demurrer, strike, compel, quash, etc.) use the standard
// § 1005 briefing schedule; only these carry their own counting rules.
function classifyMotion(mt) {
  const s = (mt || '').toLowerCase();
  if (/summary\s+judgment|summary\s+adjudication|\bmsj\b|\bmsa\b/.test(s)) return 'msj';
  if (/new\s+trial|\bjnov\b|judgment\s+notwithstanding|vacate\s+(the\s+)?judgment/.test(s)) return 'new_trial';
  if (/reconsideration|renewed?\s+motion|\bccp?\s*1008\b|\b1008\b/.test(s)) return 'recon';
  return 'standard';
}
const CATEGORY_LABEL = {
  standard: 'Standard noticed motion (CCP § 1005)',
  msj: 'Summary judgment / adjudication (CCP § 437c)',
  new_trial: 'New trial / JNOV / vacate judgment (CCP §§ 659, 629, 663a)',
  recon: 'Motion for reconsideration (CCP § 1008)',
};

// ── FORMATTING ─────────────────────────────────────────────────────────────
function fmt(d) {
  if (!d || isNaN(d)) return '—';
  return d.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
}
function fmtFull(d) {
  if (!d || isNaN(d)) return '';
  return d.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
}
function parseDate(s) {
  if (!s) return null;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/); // yyyy-mm-dd (date input)
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  const m2 = String(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // m/d/yyyy (e-court)
  if (m2) return new Date(+m2[3], +m2[1] - 1, +m2[2]);
  const d = new Date(s);
  if (isNaN(d)) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function toInputValue(d) {
  if (!d || isNaN(d)) return '';
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${mm}-${dd}`;
}
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── EVENTS ─────────────────────────────────────────────────────────────────
function onDateChange() {
  const v = document.getElementById('baseDate').value;
  state.baseDate = parseDate(v);
  const fd = document.getElementById('dateFriendly');
  fd.textContent = state.baseDate ? fmtFull(state.baseDate) : '';
  render();
}
function setMode(m) {
  state.mode = m;
  document.getElementById('tabA').className = 'tab' + (m === 'A' ? ' active' : '');
  document.getElementById('tabB').className = 'tab' + (m === 'B' ? ' active' : '');
  document.getElementById('modeA').style.display = m === 'A' ? '' : 'none';
  document.getElementById('modeB').style.display = m === 'B' ? '' : 'none';
  render();
}
function toggleExpand() { state.expanded = !state.expanded; render(); }
function setMotionType(t) { state.motionType = t; render(); }
function setService(s) { state.service = s; render(); }
function setMailRegion(r) { state.mailRegion = r; render(); }

function effectiveSvc() {
  if (state.service !== 'mail') return state.service;
  return 'mail_' + state.mailRegion;
}

// ── RENDER ─────────────────────────────────────────────────────────────────
function render() {
  renderDetectedBanner();
  if (state.mode === 'A') renderTableMode();
  else renderInteractiveMode();
}

function renderDetectedBanner() {
  const el = document.getElementById('detected');
  if (!el) return;
  const d = state.detected;
  if (!d) { el.style.display = 'none'; el.innerHTML = ''; return; }
  el.style.display = '';
  const bits = [];
  if (d.caseNumber) bits.push(`<strong>${esc(d.caseNumber)}</strong>`);
  if (d.rawMotion) bits.push(esc(d.rawMotion));
  if (d.hearingDate) bits.push(`hearing ${esc(d.hearingDate)}`);
  const rule = CATEGORY_LABEL[d.category] || CATEGORY_LABEL.standard;
  el.innerHTML =
    `<span class="det-tag">Detected</span> ${bits.join(' · ')}` +
    `<div class="det-rule">Applying: <strong>${esc(rule)}</strong></div>`;
}

// ── TABLE SECTION DATA ─────────────────────────────────────────────────────
function getSectionData(baseDate) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const DEFAULT_SVCS = ['electronic'];
  const ALL_SVCS = ['electronic', 'personal', 'mail_ca', 'mail_state', 'mail_conf', 'mail_intl', 'fax'];
  const SVC_LABELS = {
    electronic: 'Electronic', personal: 'Personal', mail_ca: 'Mail CA→CA',
    mail_state: 'Mail Other State', mail_conf: 'Mail Confidential',
    mail_intl: 'Mail International', fax: 'Fax/Overnight',
  };
  const svcs = state.expanded ? ALL_SVCS : DEFAULT_SVCS;

  const sections = [
    {
      id: 'std', label: 'Standard Motion',
      motionFn: (svc) => stdMotion(baseDate, svc), motionRule: '16 court days',
      opp: stdOpp(baseDate), oppNote: '9 court days',
      reply: stdReply(baseDate), replyNote: '5 court days', warn: null,
    },
    {
      id: 'msj', label: 'MSJ / MSA',
      motionFn: (svc) => msjMotion(baseDate, svc), motionRule: '81 cal. days',
      opp: msjOpp(baseDate), oppNote: '20 cal. days',
      reply: msjReply(baseDate), replyNote: '11 cal. days', warn: null,
    },
    {
      id: 'newtrial', label: 'New Trial / JNOV',
      motionFn: () => newTrialDL(baseDate), motionRule: '15 cal. days from notice',
      opp: stdOpp(baseDate), oppNote: '9 court days',
      reply: null, replyNote: '5 days after opp. served',
      warn: newTrialDL(baseDate) < today ? '⚠ Deadline passed — check pre-judgment alternative' : null,
    },
    {
      id: 'recon', label: 'Reconsideration',
      motionFn: () => reconDL(baseDate), motionRule: '10 cal. days from notice',
      opp: stdOpp(baseDate), oppNote: '9 court days',
      reply: stdReply(baseDate), replyNote: '5 court days', warn: null,
    },
  ];
  return { sections, svcs, SVC_LABELS };
}

function renderTableMode() {
  const el = document.getElementById('modeA');
  if (!state.baseDate) {
    el.innerHTML = `<div class="empty-state"><div class="big">📅</div>Enter a date above to generate the deadline table.</div>`;
    return;
  }
  const { sections, svcs, SVC_LABELS } = getSectionData(state.baseDate);

  let html = `
  <div class="section-bar">
    <div class="info">Base date: <strong>${fmt(state.baseDate)}</strong> &nbsp;·&nbsp; Deadlines counted back from this date.</div>
    <button class="expand-btn" data-action="expand">${state.expanded ? '↑ Fewer Columns' : '↓ All Service Types'}</button>
  </div>
  <div class="table-wrap">
  <table>
    <thead>
      <tr class="group-row">
        <th colspan="2" class="gh-empty"></th>
        <th colspan="${svcs.length}" class="gh-motion">Motion — deadline by service method</th>
        <th colspan="2" class="gh-opp-reply">Opposition &amp; Reply</th>
      </tr>
      <tr class="col-row">
        <th class="left" style="width:108px">Section</th>
        <th class="left" style="width:72px; color:var(--muted)">Rule</th>
        ${svcs.map((s, i) => `<th style="min-width:148px;${i === 0 ? ' border-left:2px solid var(--border);' : ''}">${SVC_LABELS[s]}</th>`).join('')}
        <th style="min-width:148px; border-left:2px solid var(--accent); background:var(--accent-lt); color:var(--accent)">Opposition</th>
        <th style="min-width:148px; background:var(--accent-lt); color:var(--accent)">Reply</th>
      </tr>
    </thead>
    <tbody>`;

  sections.forEach((sec, si) => {
    const rowCls = si % 2 === 0 ? 'row-a' : 'row-b';
    const motionCells = svcs.map(svc => sec.motionFn(svc));

    html += `<tr class="${rowCls}">
      <td class="section-cell">${sec.label}</td>
      <td style="font-size:0.68rem; color:var(--muted); white-space:nowrap">${sec.motionRule}</td>`;

    motionCells.forEach((d, ci) => {
      const isFirst = ci === 0;
      const warnClass = (sec.warn && sec.id === 'newtrial') ? 'warn-cell' : '';
      html += `<td class="motion-cell ${isFirst ? 'first-service' : ''} ${warnClass}">`;
      if (sec.warn && sec.id === 'newtrial' && ci === 0) html += `<div class="warn-badge">PAST DUE</div>`;
      html += `<div class="date-main">${fmt(d)}</div>`;
      if (sec.warn && sec.id === 'newtrial' && ci === 0) html += `<div class="date-note" style="color:var(--red)">${sec.warn}</div>`;
      html += `</td>`;
    });

    html += `<td class="opp-cell">`;
    if (sec.opp) html += `<div class="date-main">${fmt(sec.opp)}</div><div class="date-note">${sec.oppNote}</div>`;
    else html += `<span class="italic-note">${sec.oppNote}</span>`;
    html += `</td><td class="reply-cell">`;
    if (sec.reply) html += `<div class="date-main">${fmt(sec.reply)}</div><div class="date-note">${sec.replyNote}</div>`;
    else html += `<span class="italic-note">${sec.replyNote}</span>`;
    html += `</td></tr>`;
  });

  html += `</tbody></table></div>
  <div class="table-meta">
    * If a computed deadline falls on a non-court day, it moves to the preceding court day (motion / opp / reply)
    or the next court day (New Trial, Reconsideration — triggered by notice of entry date).
    Court days exclude weekends and California judicial holidays (CCP § 135; Columbus Day is not a court holiday).
  </div>`;

  el.innerHTML = html;
}

// ── MODE B ─────────────────────────────────────────────────────────────────
function renderInteractiveMode() {
  const el = document.getElementById('modeB');
  if (!state.baseDate) {
    el.innerHTML = `<div class="empty-state"><div class="big">📅</div>Enter a date above to begin.</div>`;
    return;
  }

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const svc = effectiveSvc();
  const mt = state.motionType;
  const triggerBased = mt === 'new_trial' || mt === 'recon';

  const MOTION_OPTS = [
    { v: 'standard', l: 'Standard Motion' },
    { v: 'msj', l: 'MSJ / MSA' },
    { v: 'new_trial', l: 'New Trial / JNOV' },
    { v: 'recon', l: 'Reconsideration' },
  ];
  const SVC_OPTS = [
    { v: 'electronic', l: 'Electronic' },
    { v: 'personal', l: 'Personal' },
    { v: 'mail', l: 'Mail' },
    { v: 'fax', l: 'Fax / Overnight' },
  ];
  const MAIL_OPTS = [
    { v: 'ca', l: 'CA → CA' },
    { v: 'state', l: 'Other State' },
    { v: 'conf', l: 'Confidential Addr.' },
    { v: 'intl', l: 'International' },
  ];

  let res = {};
  if (mt === 'standard') {
    res = {
      motion: stdMotion(state.baseDate, svc), motionNote: '16 court days before hearing (+ service)',
      opp: stdOpp(state.baseDate), oppNote: '9 court days before hearing',
      reply: stdReply(state.baseDate), replyNote: '5 court days before hearing',
    };
  } else if (mt === 'msj') {
    res = {
      motion: msjMotion(state.baseDate, svc), motionNote: '81 calendar days before hearing (+ service)',
      opp: msjOpp(state.baseDate), oppNote: '20 calendar days before hearing',
      reply: msjReply(state.baseDate), replyNote: '11 calendar days before hearing',
    };
  } else if (mt === 'new_trial') {
    const d = newTrialDL(state.baseDate);
    res = {
      motion: d, motionNote: '15 calendar days from notice of entry (no service extension)',
      opp: stdOpp(state.baseDate), oppNote: 'briefing per CCP § 659a',
      reply: null, replyNote: 'reply within 5 days after opposition is served',
      warn: d < today ? '⚠ This deadline has passed. Check whether the motion could have been filed before entry of judgment.' : null,
    };
  } else {
    res = {
      motion: reconDL(state.baseDate), motionNote: '10 calendar days from service of notice of entry (service extensions apply)',
      opp: stdOpp(state.baseDate), oppNote: '9 court days before hearing',
      reply: stdReply(state.baseDate), replyNote: '5 court days before hearing',
    };
  }

  const chipsMT = MOTION_OPTS.map(o =>
    `<button class="chip${mt === o.v ? ' active' : ''}" data-action="motion" data-value="${o.v}">${o.l}</button>`
  ).join('');
  const chipsSVC = SVC_OPTS.map(o =>
    `<button class="chip${state.service === o.v ? ' active' : ''}" data-action="service" data-value="${o.v}">${o.l}</button>`
  ).join('');
  const mailSub = state.service === 'mail' ? `
    <div class="mail-sub">
      ${MAIL_OPTS.map(o => `<button class="chip${state.mailRegion === o.v ? ' active' : ''}" data-action="mail" data-value="${o.v}">${o.l}</button>`).join('')}
    </div>` : '';
  const warnHtml = res.warn ? `<div class="warn-box">${res.warn}</div>` : '';

  const card = (label, date, note, highlight) => {
    const cls = 'result-card' + (highlight ? ' highlight' : '');
    if (date) {
      return `<div class="${cls}">
        <div class="card-label">${label}</div>
        <div class="card-date">${fmt(date)}</div>
        <div class="card-note">${note}</div>
      </div>`;
    }
    return `<div class="${cls}">
        <div class="card-label">${label}</div>
        <div class="card-italic">${note}</div>
      </div>`;
  };

  el.innerHTML = `<div class="mode-b">
    <div class="field-group">
      <span class="field-label">Motion Type</span>
      <div class="chips">${chipsMT}</div>
    </div>
    ${!triggerBased ? `
    <div class="field-group">
      <span class="field-label">Service Method <span class="sub">&nbsp;(affects motion deadline only)</span></span>
      <div class="chips">${chipsSVC}</div>
      ${mailSub}
    </div>` : `<div class="svc-note">Service method does not affect these deadlines — they run from the notice-of-entry date. Set the date above to the notice of entry, not the hearing.</div>`}
    <div class="result-cards">
      ${warnHtml}
      ${card('Motion — Serve &amp; File By', res.motion, res.motionNote, true)}
      ${card('Opposition — Serve &amp; File By', res.opp, res.oppNote, false)}
      ${card('Reply — Serve &amp; File By', res.reply, res.replyNote, false)}
    </div>
  </div>`;
}

// ── PRE-FILL FROM DETECTED HEARING ──────────────────────────────────────────
function applyDetected(data) {
  if (!data) return;
  const rawMotion = data.motionType || '';
  const category = classifyMotion(rawMotion);
  state.detected = {
    rawMotion,
    hearingDate: data.hearingDate || '',
    caseNumber: data.caseNumber || '',
    category,
  };
  const hd = parseDate(data.hearingDate);
  if (hd) {
    state.baseDate = hd;
    const input = document.getElementById('baseDate');
    if (input) input.value = toInputValue(hd);
    const fd = document.getElementById('dateFriendly');
    if (fd) fd.textContent = fmtFull(hd);
  }
  // Pre-select the matching rule and open the step-by-step view on it.
  state.motionType = category;
  state.mode = 'B';
}

function loadDetected() {
  return new Promise(resolve => {
    try {
      if (!chrome.storage || !chrome.storage.local) { resolve(null); return; }
      chrome.storage.local.get(['deadlineCalcData'], r => {
        void (chrome.runtime && chrome.runtime.lastError);
        const d = r && r.deadlineCalcData;
        // Ignore stale hand-offs (e.g. a bookmarked reopen) older than 10 min.
        if (d && d.createdAt && (Date.now() - d.createdAt) > 10 * 60 * 1000) { resolve(null); return; }
        resolve(d || null);
      });
    } catch (_) { resolve(null); }
  });
}

// ── INIT ────────────────────────────────────────────────────────────────────
document.addEventListener('click', e => {
  const t = e.target.closest && e.target.closest('[data-action]');
  if (!t) return;
  const a = t.getAttribute('data-action');
  const v = t.getAttribute('data-value');
  if (a === 'mode') setMode(v);
  else if (a === 'expand') toggleExpand();
  else if (a === 'motion') setMotionType(v);
  else if (a === 'service') setService(v);
  else if (a === 'mail') setMailRegion(v);
});

const baseInput = document.getElementById('baseDate');
if (baseInput) baseInput.addEventListener('input', onDateChange);

loadDetected().then(data => {
  if (data) applyDetected(data);
  // Reflect the mode selection in the tab UI, then render.
  setMode(state.mode);
});

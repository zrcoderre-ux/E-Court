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
  entryDate: null, // entry-of-judgment date — new trial's 180-day outer limit
  detected: null, // { rawMotion, hearingDate, caseNumber, category, ... }
  showAllTypes: false, // when detected, hide the non-applicable motion types
};

// ── DEADLINE ENGINE ──────────────────────────────────────────────────────────
// The holiday + counting + classification logic lives in the shared
// lib/deadlines.js (loaded before this script) so the case-page inline
// deadlines and this calculator can never drift.
const {
  getHolidays, isCourtDay, nextCourtDay, prevCourtDay, addCD, addCAL,
  stdMotion, msjMotion, stdOpp, msjOpp, stdReply, msjReply, newTrialDL, reconDL,
  classifyMotion,
} = LACourtDeadlines;

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
  const triggerBased = d.category === 'new_trial' || d.category === 'recon';
  let extra = '';
  if (triggerBased) {
    if (d.noticeOfEntryDate) {
      extra = `<div class="det-rule">Base set to detected notice of entry: <strong>${esc(d.noticeOfEntryDate)}</strong>` +
        (d.noticeOfEntryDoc ? ` (${esc(d.noticeOfEntryDoc)})` : '') +
        `. Adjust if a different order controls.</div>`;
    } else {
      extra = `<div class="det-rule">No notice of entry found in the case documents — set the date to the notice of entry of the order being challenged.</div>`;
    }
    if (d.category === 'new_trial' && d.entryOfJudgmentDate) {
      extra += `<div class="det-rule">Entry of judgment detected: <strong>${esc(d.entryOfJudgmentDate)}</strong>` +
        (d.entryOfJudgmentDoc ? ` (${esc(d.entryOfJudgmentDoc)})` : '') +
        ` — used for the 180-day outer limit.</div>`;
    }
  }
  el.innerHTML =
    `<span class="det-tag">Detected</span> ${bits.join(' · ')}` +
    `<div class="det-rule">Applying: <strong>${esc(rule)}</strong></div>` + extra;
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
      // Table shows the 15-day-from-notice date; the 180-day-from-entry outer
      // limit needs the entry-of-judgment date, so it lives in Step-by-Step.
      motionFn: () => newTrialDL(baseDate), motionRule: '15 days from notice, or 180 from entry (§ 659(a)(2))',
      // § 659a briefing runs from filing/service of the notice of intention, not
      // from the notice-of-entry base — so these can't be dated from it alone.
      opp: null, oppNote: 'Opposition 10 days after service of moving brief (§ 659a)',
      reply: null, replyNote: 'Reply 5 days after opposition served (§ 659a)',
      warn: newTrialDL(baseDate) < today ? '⚠ Deadline passed — check pre-judgment alternative' : null,
    },
    {
      id: 'recon', label: 'Reconsideration',
      motionFn: (svc) => reconDL(baseDate, svc), motionRule: '10 days from notice + svc. ext.',
      // Opp/reply run from the eventual hearing (§ 1005), which isn't the
      // notice-of-entry base date, so they're stated as rules, not dated.
      opp: null, oppNote: '9 court days before the eventual hearing (§ 1005(b))',
      reply: null, replyNote: '5 court days before the eventual hearing', warn: null,
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
    or the next court day (New Trial, Reconsideration — triggered by the notice-of-entry date).
    Court days exclude weekends and California judicial holidays (CCP § 135), including Lincoln's Birthday (Feb 12)
    and Native American Day (4th Friday of Sept); Columbus Day is not a court holiday.
    Reconsideration reflects the CCP §§ 1013 / 1010.6 service extensions. For MSJ, "Mail Confidential" uses the
    5-day in-state mail tier because CCP § 437c(a)(2) has no 12-day Safe at Home tier.
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
  // Only new trial is service-immune (§ 659(b)); reconsideration IS extended by
  // service method (§§ 1013 / 1010.6), so it keeps the service chips.
  const triggerBased = mt === 'new_trial';

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
    // § 659(a)(2): notice of intention is due the EARLIEST of 15 days after
    // service of notice of entry, or 180 days after entry of judgment.
    const d15 = state.baseDate ? newTrialDL(state.baseDate) : null;
    const d180 = state.entryDate ? nextCourtDay(addCAL(state.entryDate, 180)) : null;
    let controlling = null, which = '';
    if (d15 && d180) {
      if (d15 <= d180) { controlling = d15; which = '15 days after service of notice of entry controls'; }
      else { controlling = d180; which = '180 days after entry of judgment controls'; }
    } else if (d15) {
      controlling = d15; which = '15 days after service of notice of entry (add the entry-of-judgment date for the 180-day cap)';
    } else if (d180) {
      controlling = d180; which = '180 days after entry of judgment';
    }
    const parts = [];
    if (d15) parts.push('15 days from notice of entry: ' + fmt(d15));
    if (d180) parts.push('180 days from entry of judgment: ' + fmt(d180));
    res = {
      motion: controlling,
      motionNote: 'Earliest of the § 659(a)(2) triggers — ' + which + '. ' +
        (parts.length ? '(' + parts.join(' · ') + '.) ' : '') +
        'Not extended by service (§ 659(b)). Power to rule expires 75 days after notice of entry (§ 660(c)).',
      opp: null, oppNote: 'Opposition 10 days after service of the moving brief (§ 659a)',
      reply: null, replyNote: 'Reply 5 days after opposition is served (§ 659a)',
      warn: controlling && controlling < today ? '⚠ This deadline has passed. Check whether the motion could have been filed before entry of judgment.' : null,
    };
  } else {
    res = {
      motion: reconDL(state.baseDate, svc), motionNote: '10 calendar days from service of notice of entry, plus service extension (§§ 1008(a), 1013, 1010.6)',
      opp: null, oppNote: '9 court days before the eventual hearing (§ 1005(b))',
      reply: null, replyNote: '5 court days before the eventual hearing',
    };
  }

  // When the motion type was detected from the case, hide the non-applicable
  // types by default (show only the detected one) with a toggle to reveal all.
  const detected = !!state.detected;
  const showAll = state.showAllTypes || !detected;
  const visibleOpts = showAll ? MOTION_OPTS : MOTION_OPTS.filter(o => o.v === mt);
  const chipsMT = visibleOpts.map(o =>
    `<button class="chip${mt === o.v ? ' active' : ''}" data-action="motion" data-value="${o.v}">${o.l}</button>`
  ).join('') + (detected
    ? `<button class="chip chip-toggle" data-action="toggleTypes">${showAll ? '− fewer' : '+ other types'}</button>`
    : '');
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
    </div>` : `
    <div class="field-group">
      <span class="field-label">Entry of Judgment Date <span class="sub">&nbsp;(for the 180-day outer limit, § 659(a)(2))</span></span>
      <input type="date" id="entryDate" value="${state.entryDate ? toInputValue(state.entryDate) : ''}">
      <div class="svc-note">Service method does not affect this deadline (§ 659(b)). The date above is the notice of entry (15-day trigger); add the entry-of-judgment date here for the 180-day cap. The earlier of the two controls.</div>
    </div>`}
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
  const triggerBased = category === 'new_trial' || category === 'recon';
  state.detected = {
    rawMotion,
    hearingDate: data.hearingDate || '',
    caseNumber: data.caseNumber || '',
    category,
    noticeOfEntryDate: data.noticeOfEntryDate || '',
    noticeOfEntryDoc: data.noticeOfEntryDoc || '',
    entryOfJudgmentDate: data.entryOfJudgmentDate || '',
    entryOfJudgmentDoc: data.entryOfJudgmentDoc || '',
  };
  // Seed the new-trial 180-day outer limit from the detected entry-of-judgment.
  if (category === 'new_trial' && data.entryOfJudgmentDate) {
    state.entryDate = parseDate(data.entryOfJudgmentDate);
  }
  // New trial and reconsideration run from the notice of entry, not the
  // hearing — prefer the notice-of-entry date the content script detected in
  // the case documents. Everything else keys off the hearing date.
  const baseStr = (triggerBased && data.noticeOfEntryDate) ? data.noticeOfEntryDate : data.hearingDate;
  const bd = parseDate(baseStr);
  if (bd) {
    state.baseDate = bd;
    const input = document.getElementById('baseDate');
    if (input) input.value = toInputValue(bd);
    const fd = document.getElementById('dateFriendly');
    if (fd) fd.textContent = fmtFull(bd);
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
  else if (a === 'toggleTypes') { state.showAllTypes = !state.showAllTypes; render(); }
});

const baseInput = document.getElementById('baseDate');
if (baseInput) baseInput.addEventListener('input', onDateChange);

// The entry-of-judgment field is rendered inside Mode B (rebuilt each render),
// so wire it by delegation rather than a direct listener.
document.addEventListener('input', e => {
  if (e.target && e.target.id === 'entryDate') {
    state.entryDate = parseDate(e.target.value);
    render();
  }
});

loadDetected().then(data => {
  if (data) applyDetected(data);
  // Reflect the mode selection in the tab UI, then render.
  setMode(state.mode);
});

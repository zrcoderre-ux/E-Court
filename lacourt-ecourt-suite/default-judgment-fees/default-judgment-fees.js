/**
 * Default-judgment attorney-fee calculator — LASC Local Rule 3.214(a).
 *
 * Schedule text (Rule 3.214(a), "Contract Provision or Note", eff. July 1, 2011):
 *
 *   Default case:
 *     $0.01–$1,000       15% (minimum $75.00)
 *     $1,000.01–$10,000  $150   + 6% of the excess over $1,000
 *     $10,000.01–$50,000 $690   + 3% of the excess over $10,000
 *     $50,000.01–$100,000$1,890 + 2% of the excess over $50,000
 *     Over $100,000      $2,890 + 1% of the excess over $100,000
 *
 *   Contested case:
 *     $0.01–$1,000       15% (minimum $100.00)
 *     $1,000.01–$10,000  $150   + 8% of the excess over $1,000
 *     $10,000.01–$50,000 $870   + 6% of the excess over $10,000
 *     $50,000.01–$100,000$3,270 + 4% of the excess over $50,000
 *     Over $100,000      $5,270 + 2% of the excess over $100,000
 *
 *   Rule 3.214(b): a foreclosed mortgage or trust deed increases the applicable
 *   fee in (a) by 10%.
 *
 * A default judgment uses the "Default case" column; "Contested" is offered for
 * completeness. The court may allow a different amount, and extraordinary
 * services require an itemized application — this is the schedule figure only.
 */

const SCHEDULES = {
  default: {
    label: 'Default case',
    minFee: 75,
    tiers: [
      { upTo: 1000,     base: 0,    rate: 0.15, over: 0 },
      { upTo: 10000,    base: 150,  rate: 0.06, over: 1000 },
      { upTo: 50000,    base: 690,  rate: 0.03, over: 10000 },
      { upTo: 100000,   base: 1890, rate: 0.02, over: 50000 },
      { upTo: Infinity, base: 2890, rate: 0.01, over: 100000 },
    ],
  },
  contested: {
    label: 'Contested case',
    minFee: 100,
    tiers: [
      { upTo: 1000,     base: 0,    rate: 0.15, over: 0 },
      { upTo: 10000,    base: 150,  rate: 0.08, over: 1000 },
      { upTo: 50000,    base: 870,  rate: 0.06, over: 10000 },
      { upTo: 100000,   base: 3270, rate: 0.04, over: 50000 },
      { upTo: Infinity, base: 5270, rate: 0.02, over: 100000 },
    ],
  },
};

// Compute the schedule fee. Returns null for an empty/invalid/zero amount.
function computeFee(amount, scheduleKey, mortgage) {
  const sched = SCHEDULES[scheduleKey] || SCHEDULES.default;
  const a = Number(amount);
  if (!isFinite(a) || a <= 0) return null;

  let tier = sched.tiers[sched.tiers.length - 1];
  for (const t of sched.tiers) { if (a <= t.upTo) { tier = t; break; } }

  let scheduleFee = tier.base + (a - tier.over) * tier.rate;
  const isFirstTier = tier.over === 0;
  const minApplied = isFirstTier && scheduleFee < sched.minFee;
  if (minApplied) scheduleFee = sched.minFee;

  const mortgageAdj = mortgage ? scheduleFee * 0.10 : 0;
  return { amount: a, sched, tier, scheduleFee, minApplied, mortgageAdj, total: scheduleFee + mortgageAdj };
}

const money = n => n.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
const money0 = n => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 });
const pct = r => (r * 100).toFixed(0) + '%';

// Human-readable label for a tier, e.g. "$50,000.01–$100,000".
function tierRangeLabel(sched, i) {
  const t = sched.tiers[i];
  const lo = i === 0 ? 0.01 : sched.tiers[i - 1].upTo + 0.01;
  if (t.upTo === Infinity) return 'Over ' + money0(sched.tiers[i - 1].upTo);
  return money0(lo) + '–' + money0(t.upTo);
}
function tierFormulaLabel(sched, i) {
  const t = sched.tiers[i];
  if (t.over === 0) return pct(t.rate) + ' (min ' + money0(sched.minFee) + ')';
  return money0(t.base) + ' + ' + pct(t.rate) + ' of excess over ' + money0(t.over);
}

function render() {
  const amountRaw = document.getElementById('amount').value.replace(/[$,\s]/g, '');
  const scheduleKey = document.querySelector('input[name="sched"]:checked').value;
  const mortgage = document.getElementById('mortgage').checked;

  const res = computeFee(amountRaw, scheduleKey, mortgage);
  const feeEl = document.getElementById('feeValue');
  const breakdownEl = document.getElementById('breakdown');

  if (!res) {
    feeEl.textContent = '—';
    breakdownEl.innerHTML = '<span class="muted">Enter a judgment amount to compute the attorney fee.</span>';
    highlightScheduleRow(scheduleKey, null);
    return;
  }

  feeEl.textContent = money(res.total);

  const i = res.sched.tiers.indexOf(res.tier);
  const lines = [];
  lines.push('<div class="bd-row"><span>Judgment amount</span><b>' + money(res.amount) + '</b></div>');
  lines.push('<div class="bd-row"><span>Bracket (' + res.sched.label + ')</span><b>' + tierRangeLabel(res.sched, i) + '</b></div>');
  if (res.tier.over === 0) {
    lines.push('<div class="bd-row"><span>' + pct(res.tier.rate) + ' of ' + money(res.amount) + '</span><b>' + money(res.amount * res.tier.rate) + '</b></div>');
    if (res.minApplied) lines.push('<div class="bd-row muted"><span>Minimum applied</span><b>' + money0(res.sched.minFee) + '</b></div>');
  } else {
    lines.push('<div class="bd-row"><span>Base</span><b>' + money0(res.tier.base) + '</b></div>');
    lines.push('<div class="bd-row"><span>' + pct(res.tier.rate) + ' of excess over ' + money0(res.tier.over)
      + ' (' + money(res.amount - res.tier.over) + ')</span><b>' + money((res.amount - res.tier.over) * res.tier.rate) + '</b></div>');
  }
  lines.push('<div class="bd-row sub"><span>Schedule fee</span><b>' + money(res.scheduleFee) + '</b></div>');
  if (res.mortgageAdj > 0) {
    lines.push('<div class="bd-row"><span>Mortgage / trust-deed +10% (Rule 3.214(b))</span><b>' + money(res.mortgageAdj) + '</b></div>');
    lines.push('<div class="bd-row sub"><span>Total attorney fee</span><b>' + money(res.total) + '</b></div>');
  }
  breakdownEl.innerHTML = lines.join('');
  highlightScheduleRow(scheduleKey, i);
}

// Build / refresh the reference schedule table and highlight the active bracket.
function buildScheduleTable() {
  const key = document.querySelector('input[name="sched"]:checked').value;
  const sched = SCHEDULES[key];
  const rows = sched.tiers.map((t, i) =>
    '<tr data-tier="' + i + '"><td>' + tierRangeLabel(sched, i) + '</td><td>' + tierFormulaLabel(sched, i) + '</td></tr>'
  ).join('');
  document.getElementById('scheduleBody').innerHTML = rows;
}
function highlightScheduleRow(key, i) {
  document.querySelectorAll('#scheduleBody tr').forEach(tr => {
    tr.classList.toggle('active', i != null && Number(tr.getAttribute('data-tier')) === i);
  });
}

function init() {
  // Prefill case number from the case page, if the button passed it along.
  try {
    chrome.storage.local.get(['djFeesData'], r => {
      const d = r && r.djFeesData;
      if (d && d.caseNumber) {
        const banner = document.getElementById('caseBanner');
        banner.textContent = 'Case ' + d.caseNumber;
        banner.style.display = 'block';
      }
    });
  } catch (_) {}

  document.getElementById('amount').addEventListener('input', render);
  document.querySelectorAll('input[name="sched"]').forEach(r => r.addEventListener('change', () => { buildScheduleTable(); render(); }));
  document.getElementById('mortgage').addEventListener('change', render);

  buildScheduleTable();
  render();
  document.getElementById('amount').focus();
}

document.addEventListener('DOMContentLoaded', init);

# Repository instructions

## Workflow: merge when done

When the changes for a task are complete and verified, open a pull request and
merge it into the default branch (`main`) without waiting to be asked. This is a
standing instruction from the repo owner: **always merge when done.**

Exceptions — pause and confirm first when a change is risky or ambiguous, or when
the owner has said to hold off on that specific item.

---

# Project: LA Court E-Court Suite (Chrome extension)

A Manifest V3 Chrome extension (`lacourt-ecourt-suite/`) that adds tooling to the
Los Angeles Superior Court civil eCourt site (`civil.lacourt.org`). The owner is
a **trial-court** research attorney working law-and-motion; the extension speeds
up reading case pages, working up the daily agenda/calendar, computing briefing
deadlines, exporting order-template data, and drafting recommendation emails.

## How to ship a change (follow every time)

1. Make the change. Keep it self-contained and match surrounding code style.
2. Bump `lacourt-ecourt-suite/manifest.json` `version` (patch bump per change,
   e.g. `3.43.7` → `3.43.8`). Every shipped change gets its own version + PR.
3. `node --check` any `.js` you touched (there is no build/test suite).
4. Commit on branch `claude/chrome-popup-order-template-33tdnw` with these
   trailers (chat identity only — never put the model id in code/commits/PRs):

   ```
   Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>
   Claude-Session: <session url>
   ```

   Set `git config user.email noreply@anthropic.com && git config user.name Claude`
   before committing so authorship verifies.
5. `git push -u origin claude/chrome-popup-order-template-33tdnw --force-with-lease`
6. Open a PR (base `main`) and **squash-merge** it (repo rule: always merge when
   done). Repo is `zrcoderre-ux/e-court`; use the GitHub MCP tools.
7. Reset the working branch onto the merged main so the next task starts clean:
   `git fetch origin main && git checkout -B claude/chrome-popup-order-template-33tdnw origin/main`

**Known false positive:** after a squash-merge, a stop-hook warns that the tip
commit (GitHub's squash commit, committer `noreply@github.com`) is "Unverified."
This is expected and must NOT be fixed — the branch only mirrors `origin/main`,
`HEAD == origin/main`, there are no unpushed commits, and amending it would
rewrite already-merged history. Just confirm `HEAD == origin/main` and move on.

## Architecture

eCourt is a server-rendered SPA where **each case sub-tab is a full page reload**
— the content script re-runs cold on every tab, so cross-tab state is cached in
`sessionStorage`/`chrome.storage`, and the Documents/Parties/Hearings tabs are
read by background `fetch(credentials:'include')` + `DOMParser` rather than DOM
scraping the current tab.

| Path | Role |
|---|---|
| `manifest.json` | MV3 manifest; single source of the version number. |
| `service-worker.js` | Background worker: opens PDFs/forms on the opposite display, relays native-host clipboard events to agenda auto-advance. |
| `lib/case-status.js` | **The engine.** Shared by both content scripts: the deadline maths, document/party parsing, background case fetches (via a per-case "context", so it can run against the live page or any case id), the filing-status and OSC default-status computations, and the status HTML. |
| `clipboard/content.js` | Case-page script: floating Copy/Deadlines/Documents/Export buttons, movant detection, relevant-document opening, the inline "Next"-header briefing-deadline widget, OSC default-judgment flow. Pulls the status engine out of `LACCaseStatus`. |
| `clipboard/paste-rotator.js` | Fills the order-template values into forms (rotating paste). |
| `agenda/content.js` | Agenda/calendar page: Copy All (cleaned two-column output), auto-copy on load, name expansion + sort + green-rows-to-top batching, auto-advance to next day, next-day prefetch, and the per-case status shown beside each case name (same engine as the case page). |
| `documents/ingest.js` | Documents-tab script: stamps each filing with the day/time eCourt actually **posted** it (decoded from the doc endpoint's `Last-Modified`) and reports lag samples to the background. |
| `lib/mini-pdf.js` | From-scratch PDF 1.4 writer (base-14 Helvetica, ruled text tables). No build step, so nothing is vendored. |
| `lib/page-print.js` | Reads a fetched case sub-tab down to its tables and lays them out through `mini-pdf`. |
| `page-pdf/` | The tab that renders a case sub-tab as a PDF — fetches the page itself, so the extension's own UI is never in the output. |
| `order-template/` | In-extension Order Template Input popup (replaced the old Microsoft Form) + spreadsheet export. |
| `deadline-calculator/` | Standalone CA motion-deadline calculator page. |
| `lib/deadlines.js` | The deadline engine as the standalone calculator page uses it. **KEEP IN SYNC** with the `DL` copy inside `lib/case-status.js`. |
| `default-judgment-fees/` | LASC Local Rule 3.214 attorney-fee calculator (button on DJ pages). |
| `native-host/` | Python native-messaging host + Word/Excel VBA. Fires the Word mail merge after Export and watches the OS clipboard for agenda auto-advance. See its `README.md`. |
| `popup/`, `options/`, `pdf-focus/`, `icons/` | Toolbar popup, options page, background-tab PDF focus helper, icons. |

## Domain rules / standing decisions

- **Trial court — the movant is never labeled "Appellant."** A party who appeals
  is listed under both its trial role and "Appellant"; export/movant logic drops
  the appellate designation and uses the substantive role (`isAppellateRole`,
  `buildMovantRoster`, `formatMovant` in `clipboard/content.js`).
- **Default Judgment (OSC Re: Failure to Prosecute Default Judgment) export** runs
  the *same* Order Template flow as a regular motion **and** also fires a
  recommendation email to Judge Mackenzie (`mailtoUrl`). The email is the only
  DJ-specific extra; the export is not otherwise special-cased.
- **Briefing-deadline widget** (inline on the "Next" header) colors each paper by
  filing status: green = filed on time (assuming electronic service), red =
  overdue/late, neutral = not yet due. The **Motion** additionally shows **yellow**
  when it missed the electronic-service deadline but would be timely under personal
  service (no notice extension) — a cue to check the proof of service. A paper
  past due with nothing on file reads **"No Motion/Opposition/Reply"**; new trial,
  JNOV and reconsideration have no §1005 schedule (their deadlines run from notice
  of entry of judgment) and carry `motionOnly`, so the widget says only whether the
  moving papers are on the docket at all.
- **Several hearings we work up = one group per hearing DATE.** `groupWorkableHearings`
  (in `lib/case-status.js`) bundles the Next event plus every workable hearing on
  the Hearings tab into groups keyed by date; ‹ › arrows on the "Next" header step
  between days. The page's own header line is rewritten to the selected hearing
  (and restored verbatim on return); every FURTHER motion set for that same day
  gets its own full "Next: <date> <time> Hearing on <motion>" line cloned from the
  native one, inside the green band, each with its own Motion/Opposition/Reply
  display. Same-day hearings are worked up together: Documents opens the union of
  both sets (deduped by docId) and Export puts both motion types in the Motion
  Type box separated by "; ". A demurrer and a standalone motion to strike set the
  same day stay ONE work-up (the strike row collapses into the demurrer) — but a
  motion to strike or tax COSTS is its own motion and is never collapsed.
- **The Documents button also opens the case itself.** Two case-level papers
  lead the set (`caseLevelDocuments` in `clipboard/content.js`), before any
  filing and outside the tab cap: the **Register of Actions** as the court's own
  PDF — eCourt's report runner builds it from the case number alone, at the
  stable URL the tab's Print button drives (`ROA_PDF_URL`), so nothing is
  fetched, discovered or rendered — and the **Parties** tab (parties,
  representation, former representation), which eCourt gives no print endpoint
  at all. Ctrl+P is its
  only native route to a PDF and a print dialog cannot run unattended, so
  `page-pdf/view.html` renders it. That page **re-fetches** the tab instead of
  reading the live DOM, which is what keeps the extension's own buttons and
  header rewrites out of the PDF: nothing was ever added to the HTML the server
  sent. Both carry non-numeric docIds (`lac-roa`, `lac-parties`) and a
  `caseLevel` flag so they never collide with a filing or land in the
  opened-documents stats.
- **Motions in limine are out of scope.** They are the trial judge's, carry no
  § 1005 schedule on our calendar, and their papers land in a block around the
  final status conference where they get mistaken for briefing on the motion
  actually on calendar. Any document whose title says "in limine"
  (`isInLimineText`) is dropped from the Documents button and never counts
  toward a briefing deadline, and `bestFilingMatch` will not nominate one as
  another hearing's moving paper. Two exceptions, both meaning the paper really
  does belong to a hearing we work up: the Hearings tab's Document column lists
  it for that hearing (`hearingListsDocument`, `blobDocIds`), or the hearing
  being worked up is itself a motion in limine.
- **A paper filed the same day as the moving papers must NAME the motion.**
  Briefing filed on the motion's own filing day is almost never a response to
  it — that is the day a party files its briefing on other matters — so
  position and a generic title prove nothing there and `sameDayPaperNamesMotion`
  requires `movantMatchScore` above 0.5. Papers filed later keep the looser
  tests.
- **A hearing whose motion was never filed** is a real and common state (a reserved
  date the party abandoned). `bestFilingMatch` requires a score **strictly above
  0.5** so one incidental shared word can't nominate the wrong moving paper —
  legitimate pairings measure 0.75–1.00. With no moving paper, the Documents button
  opens the papers that *reference* the hearing (notice of intent, notice of
  hearing, notice of failure to serve) via `docReferencesMotion`, which tolerates
  one edit because docket titles carry typos ("New Trail", "Abritration").
- **Not sensitive / intentionally in the repo:** Judge Mackenzie's name and the
  `AMackenzie@lacourt.ca.gov` address (an elected official + standing naming
  convention), and the Windows username `ZCoderre` in native-host paths. Do not
  scrub these. Real *party* names / case numbers, however, should stay
  pseudonymized (Doe/Roe, ACME/Globex, etc.).
- **Filing date ≠ posting date.** eCourt's Documents tab shows the effective
  filing date; the clerk's intake queue posts a paper 0–3 court days later
  (median 1, measured), during business hours. The doc endpoint's
  `Last-Modified` carries the real posting time, mangled — the server passes
  epoch **seconds** to an API expecting **milliseconds**, so every document
  dates to January 1970. Multiply the parsed epoch-ms by 1000; the result's
  **UTC fields are the court's Pacific wall clock** (verified against PDF
  `/ModDate` stamps). Resolution is ~17 minutes. `INGEST_GRACE_COURT_DAYS` in
  `lib/case-status.js` is the window during which a missing paper is reported
  as "not posted yet" rather than absent; revisit it if the options page's
  Filing Lag distribution shifts. Lag varies by **document type** (`docLagCategory`)
  — clerk-entered minute orders post same-day, default prove-up packets have
  been measured at 11 court days — so read the per-type breakdown, not the
  pooled median, before changing the window.
- **The § 660 clock and the appeal clock are NOT parallel — do not "fix" this.**
  A clerk-served file-endorsed copy of the judgment starts the 60-day appeal
  period under CRC 8.104(a)(1)(A) (and so the CRC 3.1702 fee deadline that
  borrows it), but does **not** start CCP § 660's 75 days: to be a mailing
  "pursuant to Section 664.5" the notice must affirmatively state it is given
  "upon order of the court" or "under section 664.5" (*Van Beurden* (1997) 15
  Cal.4th 51, 64, clerk holding reaffirmed in *Palmer v. GTE California* (2003)
  30 Cal.4th 1265, 1274). *Palmer* answered the appeal-rule analogy directly at
  p. 1277 — rule 8.104 "at most confirms" that § 664.5's requirements *exceed*
  those of §§ 659 and 660; the rule was amended in 2002 to add a file-stamped
  copy, the statutes never were. So the same paper can start one clock and not
  the other, and `NOTICE_OF_ENTRY_JUDGMENT_RE` stays narrow while the 8.104
  trigger in `clipboard/content.js` accepts a clerk-filed `Judgment`. *Palmer*
  **did** relax the party prong: a file-stamped copy served by a party suffices
  there. § 660(c) is **75** days (amended eff. 1/1/2019; *Kabran* (2017) 2
  Cal.5th 330 addresses the old 60-day version) and rolls under § 12a, which its
  own opening clause incorporates.
- **New Outlook** is in use (no COM/`.oft`/`CreateItemFromTemplate`), so email
  drafting uses `mailto:` and the Word mail merge is launched via the native host.

## Domain reference (California law-and-motion)

- Standard motion notice: 16 court days before hearing (§1005(b)); +2 court days
  electronic, +5/+10/… calendar days for mail, +0 for personal service.
- MSJ: 81 calendar days in this engine (§437c), plus the same service extensions.
- Opposition/Reply deadlines and the MSJ variants live in `lib/deadlines.js`.
- Costs: the memorandum is due 15 days after service of the notice of entry of
  judgment/dismissal or 180 days after entry, whichever is first (CRC
  3.1700(a)(1)); a motion to strike or tax costs is due 15 days after the
  memorandum is served (3.1700(b)(1)). Both 15-day periods run from service, so
  both carry the §§ 1013 / 1010.6(a)(3)(B) extensions; the 180-day limit runs
  from entry and does not.

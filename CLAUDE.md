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
| `clipboard/content.js` | **The big one.** Case-page script: floating Copy/Deadlines/Documents/Export buttons, movant detection, relevant-document opening, the inline "Next"-header briefing-deadline widget, OSC default-judgment flow. |
| `clipboard/paste-rotator.js` | Fills the order-template values into forms (rotating paste). |
| `agenda/content.js` | Agenda/calendar page: Copy All (cleaned two-column output), auto-copy on load, name expansion + sort + green-rows-to-top batching, auto-advance to next day, next-day prefetch. |
| `order-template/` | In-extension Order Template Input popup (replaced the old Microsoft Form) + spreadsheet export. |
| `deadline-calculator/` | Standalone CA motion-deadline calculator page. |
| `lib/deadlines.js` | The deadline engine. **KEEP IN SYNC** with the inlined copy `DL` inside `clipboard/content.js`. |
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
  service (no notice extension) — a cue to check the proof of service.
- **Not sensitive / intentionally in the repo:** Judge Mackenzie's name and the
  `AMackenzie@lacourt.ca.gov` address (an elected official + standing naming
  convention), and the Windows username `ZCoderre` in native-host paths. Do not
  scrub these. Real *party* names / case numbers, however, should stay
  pseudonymized (Doe/Roe, ACME/Globex, etc.).
- **New Outlook** is in use (no COM/`.oft`/`CreateItemFromTemplate`), so email
  drafting uses `mailto:` and the Word mail merge is launched via the native host.

## Domain reference (California law-and-motion)

- Standard motion notice: 16 court days before hearing (§1005(b)); +2 court days
  electronic, +5/+10/… calendar days for mail, +0 for personal service.
- MSJ: 81 calendar days in this engine (§437c), plus the same service extensions.
- Opposition/Reply deadlines and the MSJ variants live in `lib/deadlines.js`.

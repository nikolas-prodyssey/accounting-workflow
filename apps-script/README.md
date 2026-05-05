# Accounting Apps Script

Google Apps Script (container-bound) for the **Accounting Manual** workbook. Menu title: **📒 Accounting**.

---

## Project summary

This codebase drives spreadsheet workflows for accounting: which tabs are visible, clearing user-editable areas without breaking validations, a formal **sign-off** with audit trail and optional AI-oriented **snapshots**, sheet-structure protection, and a full **re-initialize** path. Code lives under **`apps-script/`** and is deployed to the live spreadsheet with **clasp**.

---

## Repository layout

| Path | Role |
|------|------|
| `src/Code.gs` | All script logic |
| `appsscript.json` | Manifest, timezone, OAuth scopes |
| `.claspignore` | Files excluded from `clasp push` |
| `.cursor/rules/accounting-apps-script.mdc` | Cursor rule for AI onboarding (committed in repo; copy to workspace-root `.cursor/rules/` if your Cursor workspace is the parent folder) |

Local **`apps-script/.clasp.json`** (script id) is gitignored.

---

## Main features

### Generate Workflow (`generateWorkflow`)

- Must be run with the **Business Summary** tab active.
- If sign-off is current (**Sign off Page `F10` = Yes**), prompts to invalidate before proceeding.
- Confirms destructive reset.
- Clears prior error borders/tab styling.
- Reads **Business Summary**: **F17** (VAT → unhide **VAT Rates**), **C20:H20** merged multi-select (unhide matching sheet names).
- Always considers **Reporting** and **Sign off Page** visible in the workflow set.
- Clears **unprotected** cell values on workflow sheets only (**Business Summary** excluded), preserving **data validation** (`clearContent()`).
- Shows selected sheets, hides others; refreshes **sheet guard** snapshot.

### Sign off (`signOff`)

- Ensures installable **onChange** / **onEdit** triggers exist.
- Deletes any prior **Sign-off snapshot** sheet before a new run.
- Requires **Sign off Page `H5` = Yes**; validates empty required inputs on visible sheets (validated cells + merged green `#00bc70` blocks; note text `#ff9900` excluded from snapshot only, not validation UI).
- Inserts log row above row 10; writes **C10** (timestamp), **D10** (user email).
- Sets previous latest **`F11` = No**, new latest **`F10` = Yes**; syncs **green tab colour** `#00bc70` when up to date.
- Builds **Sign-off snapshot** (hidden tab): plain-text dump of unhidden sheets (rows 3+ except full **Sign off Page**); note-coloured text excluded; chunked in column A.
- Refreshes sheet guard; navigates to **Sign off Page**.

### Re-initiated spreadsheet (`reinitiatedSpreadsheet`)

- Confirms, then: deletes snapshot sheet, clears all tab colours, hides everything except **Business Summary**, deletes **Sign off Page** rows from **10** downward, clears **all** unprotected cells (including Business Summary), updates guard.

### Signed-off state & protection

- **`F10` = Yes** means latest log is up to date → green tabs and “invalidate before edit” prompts via **`onSignedOffEdit_`**.
- **`invalidateSignOff_`**: **`F10` = No**, deletes snapshot sheet, clears tab colours.
- **Sheet guard** (`onSheetChange_`): restores expected sheet names / hide state from snapshot; if user deletes **Sign-off snapshot** while signed off, invalidates sign-off and alerts.

### Progress feedback

- **`progressStep_` / `progressDone_`**: spreadsheet **toast** messages (typically bottom-right) with emoji + short pauses so steps remain readable.

### Performance notes

- Trigger installation uses **one** `getProjectTriggers()` pass.
- Validation uses **`getDisplayValues()`** where possible to reduce range reads.
- Empty visible sheets skipped in validation scans.

---

## Setup & deploy

### Prerequisites

- [clasp](https://github.com/google/clasp) (`npm i -g @google/clasp`), `clasp login`
- Container-bound script linked to the spreadsheet (clone or existing project)

### Push to Apps Script

```bash
cd "apps-script"
clasp push
```

Reload the spreadsheet after pushing.

### GitHub

```bash
cd "apps-script"
git add .
git commit -m "Your message"
git push origin main
```

---

## OAuth / permissions

Scopes are declared in **`appsscript.json`** (e.g. spreadsheets, UI, **ScriptApp** for triggers, **userinfo.email** for `Session.getActiveUser().getEmail()`). Users may need to re-authorize after scope changes.

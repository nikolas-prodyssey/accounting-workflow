# Accounting Apps Script

This Apps Script adds an **Accounting** menu to a Google Sheet and provides **Generate Workflow**.

## What it does

- Adds top menu: **Accounting → Generate Workflow**
- Enforces spreadsheet name: **Business Summary** (file name)
- Asks for confirmation before running
- Unhides sheets based on:
  - `Business Summary!F17` equals **Yes** → unhide `VAT Rates`
  - `Business Summary!C20:H20` (merged) multi-select value → unhide sheets with those names
  - Always unhides `Reporting` and `Sign off Page`
- Resets **unprotected** cells (preserves data validation rules) across all sheets **except** `Business Summary`
- Hides all other sheets

## One-time setup (local)

1) Install clasp

```bash
npm i -g @google/clasp
clasp login
```

2) Create a container-bound Apps Script attached to your target spreadsheet

- Get the spreadsheet ID from its URL:
  - `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`

Then run from this folder:

```bash
cd "apps-script"
clasp create --type sheets --title "Accounting - Workflow" --parentId "<SPREADSHEET_ID>" --rootDir .
```

3) Push code to Apps Script

```bash
clasp push
```

4) In the spreadsheet, reload the page

You should see **Accounting** in the menu bar.

## Updating Apps Script

Edit files in `apps-script/src/` and run:

```bash
cd "apps-script"
clasp push
```

## GitHub workflow (suggested)

From repo root:

```bash
git init
git add .
git commit -m "Add Accounting Apps Script workflow"
git branch -M main
git remote add origin <YOUR_GITHUB_REPO_GIT_URL>
git push -u origin main
```

## Notes

- `EXPECTED_SPREADSHEET_NAME` checks the **spreadsheet file name** (not a tab name).
- Data validations are preserved because the script uses `clearContent()` (not `clear()`).


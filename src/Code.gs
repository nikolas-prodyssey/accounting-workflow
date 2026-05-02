const EXPECTED_SHEET_NAME = 'Business Summary';
const MENU_NAME = '📒 Accounting';

const SHEET_VAT_RATES = 'VAT Rates';
const SHEET_REPORTING = 'Reporting';
const SHEET_SIGN_OFF = 'Sign off Page';
const SHEET_BUSINESS_SUMMARY = 'Business Summary';

const CELL_VAT_ENABLED = 'F17';
const RANGE_WORKFLOW_SELECTION = 'C20:H20'; // merged

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('📝 Generate Workflow', 'generateWorkflow')
    .addToUi();
}

function generateWorkflow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const activeSheetName = ss.getActiveSheet().getName();
  if (activeSheetName !== EXPECTED_SHEET_NAME) {
    ui.alert(
      'Wrong sheet',
      `This function must be run from the sheet (tab) named "${EXPECTED_SHEET_NAME}".\n\n` +
        `Current active sheet: "${activeSheetName}".`,
      ui.ButtonSet.OK
    );
    throw new Error(
      `generateWorkflow must run from the "${EXPECTED_SHEET_NAME}" sheet, not "${activeSheetName}".`
    );
  }

  const confirm = ui.alert(
    'Confirm',
    'This will reset any data of the existing spreadsheet. Are you sure you want to proceed?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const bs = ss.getSheetByName(SHEET_BUSINESS_SUMMARY);
  if (!bs) throw new Error(`Missing required sheet "${SHEET_BUSINESS_SUMMARY}".`);

  const sheetsToShow = new Set([
    SHEET_BUSINESS_SUMMARY,
    SHEET_REPORTING,
    SHEET_SIGN_OFF
  ]);

  // (1) VAT Rates conditional visibility.
  const vatEnabled = normalizeYesNo_(bs.getRange(CELL_VAT_ENABLED).getDisplayValue());
  if (vatEnabled === 'yes') sheetsToShow.add(SHEET_VAT_RATES);

  // (2) Based on the merged selection C20:H20.
  const selectionValue = bs.getRange(RANGE_WORKFLOW_SELECTION).getDisplayValue();
  for (const name of parseMultiSelect_(selectionValue)) {
    sheetsToShow.add(name);
  }

  // (4) Reset unprotected fields across spreadsheet (excluding Business Summary).
  resetUnprotectedContent_(ss, { excludeSheetNames: new Set([SHEET_BUSINESS_SUMMARY]) });

  // (3) Unhide required sheets, then (5) hide all other sheets.
  for (const sh of ss.getSheets()) {
    const shouldShow = sheetsToShow.has(sh.getName());
    if (shouldShow) sh.showSheet();
    else sh.hideSheet();
  }
}

/**
 * Clears contents (not data validation) of cells that users can edit.
 * This intentionally avoids altering protected content.
 */
function resetUnprotectedContent_(ss, opts) {
  const excludeSheetNames = (opts && opts.excludeSheetNames) || new Set();

  for (const sh of ss.getSheets()) {
    if (excludeSheetNames.has(sh.getName())) continue;

    // If a sheet protection exists, we only clear its declared unprotected ranges.
    const sheetProtections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (sheetProtections.length > 0) {
      const p = sheetProtections[0];
      const unprotected = p.getUnprotectedRanges() || [];
      for (const r of unprotected) {
        // Keep validations: clearContent() preserves data validation rules.
        r.clearContent();
      }
      continue;
    }

    // If only range protections exist, clear everything and restore protected ranges.
    // This is heavier, but keeps protected content intact.
    const protectedRanges = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    if (protectedRanges.length > 0) {
      const snapshots = protectedRanges.map((prot) => snapshotRange_(prot.getRange()));

      const dataRange = sh.getDataRange();
      dataRange.clearContent();

      for (const snap of snapshots) restoreRangeSnapshot_(snap);
      continue;
    }

    // No protections: safe to clear all content in the used range.
    sh.getDataRange().clearContent();
  }
}

function snapshotRange_(range) {
  return {
    a1: range.getA1Notation(),
    sheetId: range.getSheet().getSheetId(),
    values: range.getValues(),
    formulasR1C1: range.getFormulasR1C1(),
    // Keep validations and formatting of protected ranges intact as much as possible.
    dataValidations: range.getDataValidations(),
    numberFormats: range.getNumberFormats(),
    richTextValues: range.getRichTextValues()
  };
}

function restoreRangeSnapshot_(snap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets().find((s) => s.getSheetId() === snap.sheetId);
  if (!sheet) return;

  const range = sheet.getRange(snap.a1);

  // Restore rich text first (where present), else values/formulas.
  // Rich text restoration will overwrite values in those cells.
  if (snap.richTextValues) range.setRichTextValues(snap.richTextValues);

  range.setNumberFormats(snap.numberFormats);
  range.setDataValidations(snap.dataValidations);

  const hasAnyFormula = snap.formulasR1C1.some((row) => row.some((f) => !!f));
  if (hasAnyFormula) {
    range.setFormulasR1C1(snap.formulasR1C1);
  } else {
    range.setValues(snap.values);
  }
}

function normalizeYesNo_(value) {
  const v = String(value || '').trim().toLowerCase();
  if (v === 'yes' || v === 'y' || v === 'true') return 'yes';
  if (v === 'no' || v === 'n' || v === 'false') return 'no';
  return 'unknown';
}

function parseMultiSelect_(value) {
  const raw = String(value || '').trim();
  if (!raw) return [];

  // Common multi-select separators for validated/merged cells.
  const parts = raw
    .split(/[\n,;]+/g)
    .map((s) => s.trim())
    .filter(Boolean);

  // De-dup while preserving order.
  const out = [];
  const seen = new Set();
  for (const p of parts) {
    if (seen.has(p)) continue;
    seen.add(p);
    out.push(p);
  }
  return out;
}


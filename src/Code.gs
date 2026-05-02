const EXPECTED_SHEET_NAME = 'Business Summary';
const MENU_NAME = '📒 Accounting';

const SHEET_VAT_RATES = 'VAT Rates';
const SHEET_REPORTING = 'Reporting';
const SHEET_SIGN_OFF = 'Sign off Page';
const SHEET_BUSINESS_SUMMARY = 'Business Summary';
const SHEET_SIGN_OFF_PAGE = 'Sign off Page';

const CELL_VAT_ENABLED = 'F17';
const RANGE_WORKFLOW_SELECTION = 'C20:H20'; // merged

const SIGN_OFF_VALIDATION_CELL = 'H5';
const OK_SHEET_COLOR = '#00bc70';
const ERROR_SHEET_COLOR = '#ff0000';
const ERROR_BORDER_COLOR = '#ff0000';
// Cells with this fill are treated as required inputs when merged (even without data validation).
const REQUIRED_INPUT_FILL = '#00bc70';

const SHEET_GUARD_STATE_KEY = 'ACCOUNTING_SHEET_GUARD_STATE_V1';

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('📝 Generate Workflow', 'generateWorkflow')
    .addItem('✅ Sign off', 'signOff')
    .addToUi();
}

function generateWorkflow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ensureSheetGuardInstalled_(ss);

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

  // Clear any previous error borders/tab colors before reshaping the workbook.
  resetErrorFormattingAllSheets_(ss);

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

  // Update the guard snapshot so manual changes revert to this state.
  snapshotSheetGuardState_(ss);
}

function signOff() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ensureSheetGuardInstalled_(ss);

  const signOffSheet = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!signOffSheet) throw new Error(`Missing required sheet "${SHEET_SIGN_OFF_PAGE}".`);

  // (3) Reset previous run errors: remove borders + reset tab colors.
  resetSignOffErrorFormatting_(ss);

  // (1) Sign off request validation.
  const h5 = normalizeYesNo_(signOffSheet.getRange(SIGN_OFF_VALIDATION_CELL).getDisplayValue());
  if (h5 !== 'yes') {
    markErrorCell_(signOffSheet.getRange(SIGN_OFF_VALIDATION_CELL));
    signOffSheet.setTabColor(ERROR_SHEET_COLOR);
    ui.alert(
      'Sign off request is not validated',
      `Sign off request is not validated. Check sheet "${SHEET_SIGN_OFF_PAGE}".`,
      ui.ButtonSet.OK
    );
    throw new Error('Sign off request is not validated. Check sheet Sign off Page.');
  }

  // (2) Validate all unprotected cells on all unhidden sheets.
  const validationResult = validateVisibleSheets_(ss);
  const validationFailures = validationResult.failures;
  if (validationFailures > 0) {
    if (validationResult.firstFailure) {
      try {
        ss.setActiveSheet(validationResult.firstFailure.sheet);
        validationResult.firstFailure.range.activate();
      } catch (e) {
        // ignore
      }
    }
    ui.alert(
      'Validation errors found',
      `There are ${validationFailures} validation error(s) or empty required field(s). ` +
        `Please review the highlighted cells (red borders) and sheets (red tabs).`,
      ui.ButtonSet.OK
    );
    throw new Error('Sign off blocked due to validation errors.');
  }

  // (4) Append sign-off log (insert row above 10, write C10/D10).
  signOffSheet.insertRowBefore(10);
  const now = new Date();
  const email =
    (Session.getActiveUser && Session.getActiveUser().getEmail()) ||
    (Session.getEffectiveUser && Session.getEffectiveUser().getEmail()) ||
    '';
  signOffSheet.getRange('C10').setValue(now);
  signOffSheet.getRange('D10').setValue(email);
  signOffSheet.getRange('C10').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // Navigate to Sign off Page after completion.
  ss.setActiveSheet(signOffSheet);
  signOffSheet.getRange('A1').activate();
}

function resetSignOffErrorFormatting_(ss) {
  for (const sh of ss.getSheets()) {
    if (sh.isSheetHidden()) continue;
    sh.setTabColor(OK_SHEET_COLOR);
    for (const r of getUnprotectedRanges_(sh)) {
      // Explicit full signature avoids runtime signature mismatch across environments.
      r.setBorder(false, false, false, false, false, false, null, null);
    }
  }
}

function resetErrorFormattingAllSheets_(ss) {
  for (const sh of ss.getSheets()) {
    // Only clear the error signals we add (borders + tab color).
    sh.setTabColor(null);
    for (const r of getUnprotectedRanges_(sh)) {
      r.setBorder(false, false, false, false, false, false, null, null);
    }
  }
}

/**
 * Installs a spreadsheet onChange trigger and stores a baseline snapshot.
 * We cannot truly block UI actions like rename/hide/delete, but we can revert them.
 */
function ensureSheetGuardInstalled_(ss) {
  ensureOnChangeTrigger_();
  if (!PropertiesService.getDocumentProperties().getProperty(SHEET_GUARD_STATE_KEY)) {
    snapshotSheetGuardState_(ss);
  }
}

function ensureOnChangeTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some((t) => t.getHandlerFunction() === 'onSheetChange_');
  if (exists) return;
  ScriptApp.newTrigger('onSheetChange_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
}

function snapshotSheetGuardState_(ss) {
  const snapshot = ss.getSheets().map((sh) => ({
    id: sh.getSheetId(),
    name: sh.getName(),
    hidden: sh.isSheetHidden()
  }));
  PropertiesService.getDocumentProperties().setProperty(
    SHEET_GUARD_STATE_KEY,
    JSON.stringify(snapshot)
  );
}

/**
 * Installable onChange trigger handler. Reverts rename/hide/unhide/delete attempts.
 * Note: UI alerts are not available here; we use toast notifications.
 */
function onSheetChange_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = PropertiesService.getDocumentProperties().getProperty(SHEET_GUARD_STATE_KEY);
  if (!raw) return;

  let desired;
  try {
    desired = JSON.parse(raw);
  } catch (err) {
    return;
  }

  const desiredById = new Map(desired.map((d) => [d.id, d]));
  const currentSheets = ss.getSheets();
  const currentById = new Map(currentSheets.map((sh) => [sh.getSheetId(), sh]));

  let changed = false;

  // Restore expected sheets.
  for (const d of desired) {
    const sh = currentById.get(d.id);
    if (!sh) {
      // Cannot recover formatting/content; recreate sheet as a placeholder.
      const recreated = ss.insertSheet(d.name);
      if (d.hidden) recreated.hideSheet();
      changed = true;
      continue;
    }

    if (sh.getName() !== d.name) {
      sh.setName(d.name);
      changed = true;
    }

    const isHidden = sh.isSheetHidden();
    if (isHidden !== d.hidden) {
      if (d.hidden) sh.hideSheet();
      else sh.showSheet();
      changed = true;
    }
  }

  // Any unexpected new sheets: hide them.
  for (const sh of currentSheets) {
    if (desiredById.has(sh.getSheetId())) continue;
    if (!sh.isSheetHidden()) {
      sh.hideSheet();
      changed = true;
    }
  }

  if (changed) {
    ss.toast(
      'Sheet structure changes are restricted. The workbook reverted your change.',
      MENU_NAME,
      8
    );
  }
}

function validateVisibleSheets_(ss) {
  let failures = 0;
  let firstFailure = null;

  for (const sh of ss.getSheets()) {
    if (sh.isSheetHidden()) continue;

    const unprotectedRanges = getUnprotectedRanges_(sh);
    for (const r of unprotectedRanges) {
      const res = validateRange_(r);
      failures += res.failures;
      if (!firstFailure && res.firstFailure) firstFailure = res.firstFailure;
    }

    if (failures > 0) {
      // If any failures happened on any range, we will color individual sheets red when we mark cells.
      // (no-op here)
    }
  }

  return { failures, firstFailure };
}

function validateRange_(range) {
  const sh = range.getSheet();
  const values = range.getValues();
  const formulas = range.getFormulas();
  const validations = range.getDataValidations();
  const mergedRanges = range.getMergedRanges();

  const badCells = [];
  let firstFailureCell = null;

  // 1) Required merged "input blocks" (even without validation)
  // We treat a merged range as required if:
  // - Any cell inside has data validation, OR
  // - The top-left cell has the configured input fill color.
  for (const mr of mergedRanges) {
    const tl = mr.getCell(1, 1);
    const tlBg = String(tl.getBackground() || '').toLowerCase();

    const mrValidations = mr.getDataValidations();
    const hasAnyValidation = mrValidations.some((row) => row.some((v) => !!v));
    const isRequiredMerged = hasAnyValidation || tlBg === REQUIRED_INPUT_FILL;
    if (!isRequiredMerged) continue;

    const display = String(tl.getDisplayValue() || '').trim();
    if (!display) {
      // mark only the top-left cell for merged ranges
      const rowOffset = tl.getRow() - range.getRow();
      const colOffset = tl.getColumn() - range.getColumn();
      badCells.push({ rowOffset, colOffset });
      if (!firstFailureCell) firstFailureCell = { rowOffset, colOffset };
    }
  }

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      const hasValidation = !!validations[r][c];
      if (!hasValidation) continue; // only treat validated cells as required inputs

      const value = values[r][c];
      const formula = formulas[r][c];

      const isEmpty =
        (value === '' || value === null) && (formula === '' || formula === null);

      // Only check empties on "input" cells (those with data validation).
      if (isEmpty) {
        badCells.push({ rowOffset: r, colOffset: c });
        if (!firstFailureCell) firstFailureCell = { rowOffset: r, colOffset: c };
        continue;
      }
    }
  }

  if (badCells.length === 0) return { failures: 0, firstFailure: null };

  sh.setTabColor(ERROR_SHEET_COLOR);
  for (const cell of badCells) {
    const a1 = range.offset(cell.rowOffset, cell.colOffset, 1, 1);
    markErrorCell_(a1);
  }

  const firstFailure = firstFailureCell
    ? {
        sheet: sh,
        range: range.offset(firstFailureCell.rowOffset, firstFailureCell.colOffset, 1, 1)
      }
    : null;

  return { failures: badCells.length, firstFailure };
}

function markErrorCell_(range) {
  // setBorder(top, left, bottom, right, vertical, horizontal, color, style)
  range.setBorder(
    true,
    true,
    true,
    true,
    false,
    false,
    ERROR_BORDER_COLOR,
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
}

function getUnprotectedRanges_(sheet) {
  const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (sheetProtections.length > 0) {
    return sheetProtections[0].getUnprotectedRanges() || [];
  }

  const base = sheet.getDataRange();
  const protectedRanges = (sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [])
    .map((p) => p.getRange())
    .filter(Boolean);

  if (protectedRanges.length === 0) return [base];
  return subtractProtectedFromBase_(base, protectedRanges);
}

function subtractProtectedFromBase_(baseRange, protectedRanges) {
  const sh = baseRange.getSheet();
  const baseRect = rectFromRange_(baseRange);

  let rects = [baseRect];
  for (const pr of protectedRanges) {
    const protRect = rectFromRange_(pr);
    rects = rects.flatMap((r) => subtractRect_(r, protRect));
  }

  return rects
    .filter((r) => r.r1 <= r.r2 && r.c1 <= r.c2)
    .map((r) => sh.getRange(r.r1, r.c1, r.r2 - r.r1 + 1, r.c2 - r.c1 + 1));
}

function rectFromRange_(range) {
  const r1 = range.getRow();
  const c1 = range.getColumn();
  const r2 = r1 + range.getNumRows() - 1;
  const c2 = c1 + range.getNumColumns() - 1;
  return { r1, c1, r2, c2 };
}

function intersectRect_(a, b) {
  const r1 = Math.max(a.r1, b.r1);
  const c1 = Math.max(a.c1, b.c1);
  const r2 = Math.min(a.r2, b.r2);
  const c2 = Math.min(a.c2, b.c2);
  if (r1 > r2 || c1 > c2) return null;
  return { r1, c1, r2, c2 };
}

function subtractRect_(rect, cut) {
  const i = intersectRect_(rect, cut);
  if (!i) return [rect];

  const out = [];

  // top band
  if (rect.r1 < i.r1) out.push({ r1: rect.r1, c1: rect.c1, r2: i.r1 - 1, c2: rect.c2 });
  // bottom band
  if (i.r2 < rect.r2) out.push({ r1: i.r2 + 1, c1: rect.c1, r2: rect.r2, c2: rect.c2 });
  // left band
  if (rect.c1 < i.c1) out.push({ r1: i.r1, c1: rect.c1, r2: i.r2, c2: i.c1 - 1 });
  // right band
  if (i.c2 < rect.c2) out.push({ r1: i.r1, c1: i.c2 + 1, r2: i.r2, c2: rect.c2 });

  return out;
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


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
/**
 * Latest sign-off log row is always row 10 (after insertRowBefore(10)).
 * Column F: Yes = up to date (new entries get Yes on sign-off; drives #00bc70 tabs + edit warnings); No = invalidated.
 */
const SIGN_OFF_LATEST_LOG_STATUS_CELL = 'F10';
const OK_SHEET_COLOR = '#00bc70';
const ERROR_SHEET_COLOR = '#ff0000';
const ERROR_BORDER_COLOR = '#ff0000';
// Cells with this fill are treated as required inputs when merged (even without data validation).
const REQUIRED_INPUT_FILL = '#00bc70';

const SHEET_GUARD_STATE_KEY = 'ACCOUNTING_SHEET_GUARD_STATE_V1';
const SKIP_SIGN_OFF_EDIT_KEY = 'ACCOUNTING_SKIP_SIGN_OFF_EDIT';

const SIGNED_OFF_EDIT_WARNING =
  'This working is already signed off. Any changes will invalidated the signing off of this document.';

/** Dump of unhidden sheet data for AI / audit (full grid ranges, not protection-filtered). */
const SIGN_OFF_SNAPSHOT_SHEET_NAME = 'Sign-off snapshot';
/** Apps Script cell character limit is 50,000; chunk below that. */
const SNAPSHOT_CHUNK_CHARS = 49000;
/** Font colour used for notes — excluded from sign-off snapshot text. */
const NOTE_TEXT_COLOR = '#ff9900';

/** Toast duration (seconds). Sheets shows these toward the bottom-right of the window. */
const PROGRESS_TOAST_SECONDS = 7;
/** Brief pause so the next toast does not replace the previous before the user reads it. */
const PROGRESS_STEP_PAUSE_MS = 800;

/**
 * Status update via spreadsheet toast (native Sheets UI, typically bottom-right).
 * @param {string} message Full line including emoji.
 */
function progressToast_(ss, message, durationSeconds) {
  ss.toast(message, MENU_NAME, durationSeconds != null ? durationSeconds : PROGRESS_TOAST_SECONDS);
}

function progressStep_(ss, emoji, message, pauseMs) {
  progressToast_(ss, emoji + ' ' + message, PROGRESS_TOAST_SECONDS);
  if (pauseMs !== false && (pauseMs === undefined || pauseMs > 0)) {
    Utilities.sleep(pauseMs === undefined ? PROGRESS_STEP_PAUSE_MS : pauseMs);
  }
}

function progressDone_(ss, emoji, message) {
  progressToast_(ss, emoji + ' ' + message, 10);
}

function onOpen(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    syncTabColorsToSignOffStatus_(ss);
    if (isSignedOff_(ss) && !ss.getSheetByName(SIGN_OFF_SNAPSHOT_SHEET_NAME)) {
      invalidateSignOff_(ss);
      SpreadsheetApp.getUi().alert(
        'Sign-off invalidated',
        'The file needs to be signed off again. The latest sign off has been invalidated.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (err) {
    // ignore
  }

  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('📝 Generate Workflow', 'generateWorkflow')
    .addItem('✅ Sign off', 'signOff')
    .addItem('Re-initiated spreadsheet', 'reinitiatedSpreadsheet')
    .addToUi();
}

/**
 * Full reset: snapshot removed, tab colours cleared, only Business Summary visible,
 * sign-off log rows (10+) removed, unprotected cell values cleared (validations kept).
 */
function reinitiatedSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    'Confirm',
    'This will re-initialize the spreadsheet: delete the sign-off snapshot, clear all tab colours, ' +
      'hide every sheet except Business Summary, remove Sign off Page rows from row 10 downward, ' +
      'and clear unprotected cells across the workbook (data validation rules stay). Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  const bs = ss.getSheetByName(SHEET_BUSINESS_SUMMARY);
  if (!bs) {
    ui.alert('Missing sheet', `Sheet "${SHEET_BUSINESS_SUMMARY}" was not found.`, ui.ButtonSet.OK);
    return;
  }

  progressStep_(ss, '🔧', 'Starting re-init…', PROGRESS_STEP_PAUSE_MS);

  PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
  try {
    deleteSignOffSnapshotSheetIfExists_(ss);
    progressStep_(ss, '🗑️', 'Removed Sign-off snapshot sheet (if it existed).', PROGRESS_STEP_PAUSE_MS);

    // Tab colours + visibility in one pass (fewer sheet iterations).
    for (const sh of ss.getSheets()) {
      sh.setTabColor(null);
      if (sh.getName() === SHEET_BUSINESS_SUMMARY) sh.showSheet();
      else sh.hideSheet();
    }
    progressStep_(ss, '🎨', 'Cleared tab colours. Only Business Summary stays visible.', PROGRESS_STEP_PAUSE_MS);

    deleteSignOffLogsFromRow10_(ss);
    progressStep_(ss, '📋', 'Removed Sign off logs from row 10 downward.', PROGRESS_STEP_PAUSE_MS);

    resetUnprotectedContent_(ss, { excludeSheetNames: new Set() });
    progressStep_(ss, '🧹', 'Cleared unprotected cells (validations kept).', PROGRESS_STEP_PAUSE_MS);

    ss.setActiveSheet(bs);
    snapshotSheetGuardState_(ss);
    progressDone_(ss, '✅', 'Re-init finished. Workbook is ready to rebuild.');
  } finally {
    PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
  }
}

/** Deletes all rows from row 10 to the end of used range on Sign off Page. */
function deleteSignOffLogsFromRow10_(ss) {
  const sh = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 10) return;
  sh.deleteRows(10, lastRow - 9);
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

  if (isSignedOff_(ss)) {
    const inv = promptInvalidateSignedOff_(ui);
    if (!inv) return;
    invalidateSignOff_(ss);
  }

  const confirm = ui.alert(
    'Confirm',
    'This will reset any data of the existing spreadsheet. Are you sure you want to proceed?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  progressStep_(ss, '🚀', 'Generate Workflow running — resetting styling…', PROGRESS_STEP_PAUSE_MS);

  resetErrorFormattingAllSheets_(ss);

  const bs = ss.getSheetByName(SHEET_BUSINESS_SUMMARY);
  if (!bs) throw new Error(`Missing required sheet "${SHEET_BUSINESS_SUMMARY}".`);

  progressStep_(ss, '📖', 'Reading VAT & workflow selections from Business Summary…', PROGRESS_STEP_PAUSE_MS);

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

  progressStep_(ss, '🧹', 'Clearing unprotected cells on workflow sheets (keeping Business Summary).', PROGRESS_STEP_PAUSE_MS);

  resetUnprotectedContent_(ss, { excludeSheetNames: new Set([SHEET_BUSINESS_SUMMARY]) });

  progressStep_(ss, '👁️', 'Showing/hiding tabs for your selected workflow…', PROGRESS_STEP_PAUSE_MS);

  for (const sh of ss.getSheets()) {
    const shouldShow = sheetsToShow.has(sh.getName());
    if (shouldShow) sh.showSheet();
    else sh.hideSheet();
  }

  snapshotSheetGuardState_(ss);
  progressDone_(ss, '✅', 'Workflow ready — relevant sheets are visible.');
}

function signOff() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  progressStep_(ss, '✍️', 'Sign off — checking triggers & guards…', PROGRESS_STEP_PAUSE_MS);

  ensureSheetGuardInstalled_(ss);

  const signOffSheet = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!signOffSheet) throw new Error(`Missing required sheet "${SHEET_SIGN_OFF_PAGE}".`);

  deleteSignOffSnapshotSheetIfExists_(ss);
  progressStep_(ss, '🗑️', 'Old snapshot removed — preparing a fresh sign-off.', PROGRESS_STEP_PAUSE_MS);

  resetSignOffErrorFormatting_(ss);

  progressStep_(ss, '🔍', 'Checking Sign off request (H5)…', PROGRESS_STEP_PAUSE_MS);

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

  progressStep_(ss, '📋', 'Scanning visible sheets for empty required fields…', PROGRESS_STEP_PAUSE_MS);

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

  progressStep_(ss, '📝', 'Writing sign-off log row & updating status (F10/F11)…', PROGRESS_STEP_PAUSE_MS);

  signOffSheet.insertRowBefore(10);
  const now = new Date();
  const email =
    (Session.getActiveUser && Session.getActiveUser().getEmail()) ||
    (Session.getEffectiveUser && Session.getEffectiveUser().getEmail()) ||
    '';
  signOffSheet.getRange('C10').setValue(now);
  signOffSheet.getRange('D10').setValue(email);
  signOffSheet.getRange('C10').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  // Previous latest log was row 10 before insertRowBefore; it is now row 11 — no longer the current entry.
  signOffSheet.getRange('F11').setValue('No');

  // Latest log row (row 10): column F = Yes (up to date for this new entry); tabs match via sync below.
  PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
  try {
    signOffSheet.getRange(SIGN_OFF_LATEST_LOG_STATUS_CELL).setValue('Yes');
    syncTabColorsToSignOffStatus_(ss);
  } finally {
    PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
  }

  progressStep_(ss, '📸', 'Building hidden Sign-off snapshot for export / AI…', PROGRESS_STEP_PAUSE_MS);

  writeSignOffSnapshotToSheet_(ss);
  snapshotSheetGuardState_(ss);

  ss.setActiveSheet(signOffSheet);
  signOffSheet.getRange('A1').activate();
  progressDone_(ss, '✅', 'Sign off complete — tabs updated, snapshot saved.');
}

/**
 * Builds plain-text export: every unhidden sheet uses full data range (values as displayed).
 * - Non–Sign-off sheets: rows 3+ only (columns 1..lastColumn).
 * - Sign off Page: entire used range (all rows).
 */
function buildSignOffSnapshotText_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss zzz");
  const lines = [];
  lines.push('# Sign-off snapshot');
  lines.push('# Generated: ' + stamp);
  lines.push('');

  const visible = ss
    .getSheets()
    .filter((sh) => !sh.isSheetHidden() && sh.getName() !== SIGN_OFF_SNAPSHOT_SHEET_NAME);

  for (const sh of visible) {
    const title = sh.getName();
    lines.push('');
    lines.push('=== SHEET: ' + title + ' ===');

    if (title === SHEET_SIGN_OFF_PAGE) {
      const dr = sh.getDataRange();
      if (!snapshotRangeHasContent_(dr)) {
        lines.push('(empty)');
        continue;
      }
      appendSnapshotLinesFromRange_(lines, dr);
    } else {
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow < 3 || lastCol < 1) {
        lines.push('(no data below row 2)');
        continue;
      }
      const body = sh.getRange(3, 1, lastRow, lastCol);
      appendSnapshotLinesFromRange_(lines, body);
    }
  }

  return lines.join('\n');
}

function snapshotRangeHasContent_(range) {
  const v = range.getDisplayValues();
  for (let r = 0; r < v.length; r++) {
    for (let c = 0; c < v[r].length; c++) {
      if (String(v[r][c]).length) return true;
    }
  }
  return false;
}

function appendSnapshotLinesFromRange_(lines, range) {
  const values = range.getDisplayValues();
  const fonts = range.getFontColors();
  for (let r = 0; r < values.length; r++) {
    const row = [];
    for (let c = 0; c < values[r].length; c++) {
      if (isNoteTextColor_(fonts[r] && fonts[r][c])) row.push('');
      else row.push(values[r][c]);
    }
    lines.push(row.join('\t'));
  }
}

function isNoteTextColor_(fontColor) {
  const n = normalizeHexColor_(fontColor);
  return n === normalizeHexColor_(NOTE_TEXT_COLOR);
}

function normalizeHexColor_(color) {
  if (color == null || color === '') return '';
  let s = String(color).trim().toLowerCase();
  if (s.charAt(0) !== '#') s = '#' + s;
  if (s.length === 4) {
    s = '#' + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
  }
  return s;
}

/** Writes snapshot into a sheet tab (one chunk per row in column A). */
function writeSignOffSnapshotToSheet_(ss) {
  const text = buildSignOffSnapshotText_(ss);
  let sh = ss.getSheetByName(SIGN_OFF_SNAPSHOT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SIGN_OFF_SNAPSHOT_SHEET_NAME);
  sh.clearContents();
  let row = 1;
  for (let i = 0; i < text.length; i += SNAPSHOT_CHUNK_CHARS) {
    const chunk = text.substring(i, Math.min(i + SNAPSHOT_CHUNK_CHARS, text.length));
    sh.getRange(row, 1).setValue(chunk);
    row++;
  }
  sh.hideSheet();
}

function deleteSignOffSnapshotSheetIfExists_(ss) {
  const snap = ss.getSheetByName(SIGN_OFF_SNAPSHOT_SHEET_NAME);
  if (snap) ss.deleteSheet(snap);
}

function resetSignOffErrorFormatting_(ss) {
  for (const sh of ss.getSheets()) {
    if (sh.isSheetHidden()) continue;
    for (const r of getUnprotectedRanges_(sh)) {
      // Explicit full signature avoids runtime signature mismatch across environments.
      r.setBorder(false, false, false, false, false, false, null, null);
    }
  }
  syncTabColorsToSignOffStatus_(ss);
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
  ensureTriggersInstalled_();
  if (!PropertiesService.getDocumentProperties().getProperty(SHEET_GUARD_STATE_KEY)) {
    snapshotSheetGuardState_(ss);
  }
}

/** One scan of project triggers instead of two separate passes. */
function ensureTriggersInstalled_() {
  const triggers = ScriptApp.getProjectTriggers();
  const handlers = new Set(triggers.map((t) => t.getHandlerFunction()));
  const spreadsheet = SpreadsheetApp.getActive();
  if (!handlers.has('onSheetChange_')) {
    ScriptApp.newTrigger('onSheetChange_').forSpreadsheet(spreadsheet).onChange().create();
  }
  if (!handlers.has('onSignedOffEdit_')) {
    ScriptApp.newTrigger('onSignedOffEdit_').forSpreadsheet(spreadsheet).onEdit().create();
  }
}

/**
 * When F10 is Yes (latest log up to date), any edit prompts to invalidate (F10 → No, clear tab colors) or revert.
 * Requires installable onEdit trigger (see ensureTriggersInstalled_).
 */
function onSignedOffEdit_(e) {
  if (!e) return;
  if (PropertiesService.getDocumentProperties().getProperty(SKIP_SIGN_OFF_EDIT_KEY)) return;

  const ss = e.source;
  const signOffSh = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!signOffSh) return;

  if (!isSignedOff_(ss)) return;

  const range = e.range;
  if (isManualInvalidateLatestLogStatusEdit_(signOffSh, range, e.value)) return;

  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (err) {
    return;
  }

  const response = ui.alert('Signed off', SIGNED_OFF_EDIT_WARNING, ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    invalidateSignOff_(ss);
    return;
  }

  try {
    if (e.oldValue !== undefined) {
      range.setValue(e.oldValue);
    } else {
      ss.toast(
        'Edit cancelled. Paste or multi-cell changes could not be reverted automatically.',
        MENU_NAME,
        10
      );
    }
  } catch (err2) {
    // ignore
  }
}

function isManualInvalidateLatestLogStatusEdit_(signOffSh, range, newValue) {
  if (range.getSheet().getSheetId() !== signOffSh.getSheetId()) return false;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return false;
  const ref = signOffSh.getRange(SIGN_OFF_LATEST_LOG_STATUS_CELL);
  if (range.getRow() !== ref.getRow() || range.getColumn() !== ref.getColumn()) return false;
  return normalizeYesNo_(String(newValue)) === 'no';
}

function isSignedOff_(ss) {
  const sh = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!sh) return false;
  return normalizeYesNo_(sh.getRange(SIGN_OFF_LATEST_LOG_STATUS_CELL).getDisplayValue()) === 'yes';
}

function syncTabColorsToSignOffStatus_(ss) {
  if (isSignedOff_(ss)) {
    for (const sh of ss.getSheets()) sh.setTabColor(OK_SHEET_COLOR);
  } else {
    for (const sh of ss.getSheets()) sh.setTabColor(null);
  }
}

function invalidateSignOff_(ss) {
  const sh = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!sh) return;
  PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
  try {
    sh.getRange(SIGN_OFF_LATEST_LOG_STATUS_CELL).setValue('No');
    deleteSignOffSnapshotSheetIfExists_(ss);
    syncTabColorsToSignOffStatus_(ss);
  } finally {
    PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
  }
}

function promptInvalidateSignedOff_(ui) {
  const response = ui.alert('Signed off', SIGNED_OFF_EDIT_WARNING, ui.ButtonSet.OK_CANCEL);
  return response === ui.Button.OK;
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
      if (d.name === SIGN_OFF_SNAPSHOT_SHEET_NAME && isSignedOff_(ss)) {
        PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
        try {
          const so = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
          if (so) so.getRange(SIGN_OFF_LATEST_LOG_STATUS_CELL).setValue('No');
          syncTabColorsToSignOffStatus_(ss);
        } finally {
          PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
        }
        try {
          SpreadsheetApp.getUi().alert(
            'Sign-off invalidated',
            'The file needs to be signed off again. The latest sign off has been invalidated.',
            SpreadsheetApp.getUi().ButtonSet.OK
          );
        } catch (alertErr) {
          ss.toast('The latest sign off has been invalidated. Sign off again.', MENU_NAME, 12);
        }
        snapshotSheetGuardState_(ss);
        return;
      }
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
    if (sh.getLastRow() === 0) continue;

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
  // Single read for displayed text (one round-trip vs getValues + getFormulas).
  const displayValues = range.getDisplayValues();
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

  const numCols = displayValues[0] ? displayValues[0].length : 0;
  for (let r = 0; r < displayValues.length; r++) {
    for (let c = 0; c < numCols; c++) {
      const hasValidation = !!validations[r][c];
      if (!hasValidation) continue;

      const dv = String(displayValues[r][c] ?? '').trim();
      if (dv === '') {
        badCells.push({ rowOffset: r, colOffset: c });
        if (!firstFailureCell) firstFailureCell = { rowOffset: r, colOffset: c };
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


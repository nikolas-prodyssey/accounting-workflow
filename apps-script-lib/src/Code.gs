/**
 * Accounting Library — accounting-lib (Apps Script standalone library).
 *
 * Public API used by the bound shim ("AccountingLib.*"):
 *   - onOpen(e)
 *   - signOff()
 *   - generateWorkflow()
 *   - reinitiatedSpreadsheet()
 *   - onSheetChange(e)        (installable trigger handler body)
 *   - onSignedOffEdit(e)      (installable trigger handler body)
 *   - bootstrap(ss)           (silent, non-interactive variant of reinit; used by add-on after copy)
 *
 * The shim is responsible for:
 *   - Installing/repairing installable triggers (`onSheetChange_`, `onSignedOffEdit_`).
 *   - Hosting the menu (via simple onOpen) and the trigger handlers that delegate to this library.
 *
 * Shared OpenAI API key + prompt id live in this library's Script Properties:
 *   - OPENAI_API_KEY  (required)
 *   - OPENAI_ACCOUNTING_MANUAL_PROMPT_ID  (optional; otherwise default below)
 */

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

/** Business Summary — merged client name for OpenAI `client` and generated Doc filename (same range). */
const RANGE_BUSINESS_SUMMARY_CLIENT = 'C7:H7';

/** OpenAI Responses API + stored prompt (dashboard prompt id `pmpt_…`). */
const OPENAI_RESPONSES_ENDPOINT = 'https://api.openai.com/v1/responses';
/** Default if Script property `OPENAI_ACCOUNTING_MANUAL_PROMPT_ID` is not set. */
const OPENAI_ACCOUNTING_MANUAL_PROMPT_ID = 'pmpt_69f83e11be9c81938a5a7909db1612120cb63d58fcc22996';
/** Optional override (library Script properties). */
const OPENAI_PROMPT_ID_PROPERTY = 'OPENAI_ACCOUNTING_MANUAL_PROMPT_ID';
/** Required script property: OpenAI API key (single, shared key in library). */
const OPENAI_API_KEY_PROPERTY = 'OPENAI_API_KEY';

/** Optional overrides for OpenAI prompt variable names. */
const OPENAI_PROMPT_VAR_CLIENT_PROPERTY = 'OPENAI_PROMPT_VAR_CLIENT';
const OPENAI_PROMPT_VAR_SNAPSHOT_PROPERTY = 'OPENAI_PROMPT_VAR_SIGN_OFF_SNAPSHOT';
/** Optional A1 range override on Business Summary (legacy alias kept for back-compat). */
const OPENAI_CLIENT_RANGE_A1_PROPERTY = 'OPENAI_CLIENT_RANGE_A1';
const OPENAI_SIGN_OFF_CLIENT_RANGE_LEGACY_PROPERTY = 'OPENAI_SIGN_OFF_CLIENT_RANGE';

/** Default variable names (used if Script properties above are not set). */
const OPENAI_PROMPT_VAR_CLIENT_DEFAULT = 'client';
// Stored prompt expects {{sign_off_page}} (populated with the full sign-off snapshot text).
const OPENAI_PROMPT_VAR_SNAPSHOT_DEFAULT = 'sign_off_page';
/** Tab created/updated with the generated accounting manual text (chunked in column A). */
const SHEET_ACCOUNTING_MANUAL_OUTPUT = 'Accounting manual';

/** Toast duration (seconds). Sheets shows these toward the bottom-right of the window. */
const PROGRESS_TOAST_SECONDS = 7;
/** Brief pause so the next toast does not replace the previous before the user reads it. */
const PROGRESS_STEP_PAUSE_MS = 800;

// ------------------------------------------------------------------
// Public exports
// ------------------------------------------------------------------

/**
 * Public alias for installable onChange trigger handler. The shim's onSheetChange_ delegates here.
 */
function onSheetChange(e) {
  return onSheetChange_(e);
}

/**
 * Public alias for installable onEdit trigger handler. The shim's onSignedOffEdit_ delegates here.
 */
function onSignedOffEdit(e) {
  return onSignedOffEdit_(e);
}

/**
 * Non-interactive bootstrap used by the add-on right after a workbook is copied from the master.
 * Performs the same effect as `reinitiatedSpreadsheet()` minus the confirmation dialog and minus
 * trigger installation (which the shim does lazily on first user action).
 */
function bootstrap(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('AccountingLib.bootstrap: no active spreadsheet provided.');

  const bs = ss.getSheetByName(SHEET_BUSINESS_SUMMARY);
  if (!bs) throw new Error(`Missing required sheet "${SHEET_BUSINESS_SUMMARY}".`);

  PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
  try {
    deleteSignOffSnapshotSheetIfExists_(ss);
    deleteAccountingManualSheetIfExists_(ss);

    for (const sh of ss.getSheets()) {
      sh.setTabColor(null);
      if (sh.getName() === SHEET_BUSINESS_SUMMARY) sh.showSheet();
      else sh.hideSheet();
    }

    deleteSignOffLogsFromRow10_(ss);
    resetUnprotectedContent_(ss, { excludeSheetNames: new Set() });

    ss.setActiveSheet(bs);
    snapshotSheetGuardState_(ss);
  } finally {
    PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
  }
}

// ------------------------------------------------------------------
// Toast helpers
// ------------------------------------------------------------------

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

// ------------------------------------------------------------------
// Menu + entry points
// ------------------------------------------------------------------

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
    deleteAccountingManualSheetIfExists_(ss);
    progressStep_(ss, '🗑️', 'Removed Sign-off snapshot + Accounting manual sheets (if any).', PROGRESS_STEP_PAUSE_MS);

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

  const sheetsToShow = new Set([SHEET_BUSINESS_SUMMARY, SHEET_REPORTING, SHEET_SIGN_OFF]);

  const vatEnabled = normalizeYesNo_(bs.getRange(CELL_VAT_ENABLED).getDisplayValue());
  if (vatEnabled === 'yes') sheetsToShow.add(SHEET_VAT_RATES);

  const selectionValue = bs.getRange(RANGE_WORKFLOW_SELECTION).getDisplayValue();
  for (const name of parseMultiSelect_(selectionValue)) {
    sheetsToShow.add(name);
  }

  progressStep_(
    ss,
    '🧹',
    'Clearing unprotected cells on workflow sheets (keeping Business Summary).',
    PROGRESS_STEP_PAUSE_MS
  );

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

  progressStep_(ss, '✍️', 'Sign off — checking guards…', PROGRESS_STEP_PAUSE_MS);

  ensureSheetGuardInstalled_(ss);

  const signOffSheet = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!signOffSheet) throw new Error(`Missing required sheet "${SHEET_SIGN_OFF_PAGE}".`);

  resetSignOffErrorFormatting_(ss);

  progressStep_(ss, '🔍', 'Checking Sign off request (H5)…', PROGRESS_STEP_PAUSE_MS);

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

  progressStep_(ss, '🗑️', 'Removing old snapshot (if any)…', PROGRESS_STEP_PAUSE_MS);

  PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
  try {
    deleteSignOffSnapshotSheetIfExists_(ss);

    progressStep_(ss, '📝', 'Writing sign-off log row & updating status (F10/F11)…', PROGRESS_STEP_PAUSE_MS);

    const prevLogC10 = String(signOffSheet.getRange('C10').getDisplayValue() || '').trim();
    const prevLogD10 = String(signOffSheet.getRange('D10').getDisplayValue() || '').trim();
    const hadPriorSignOffLog = prevLogC10 !== '' || prevLogD10 !== '';

    signOffSheet.insertRowBefore(10);
    const now = new Date();
    const email =
      (Session.getActiveUser && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser && Session.getEffectiveUser().getEmail()) ||
      '';
    signOffSheet.getRange('C10').setValue(now);
    signOffSheet.getRange('D10').setValue(email);
    signOffSheet.getRange('C10').setNumberFormat('yyyy-mm-dd hh:mm:ss');

    if (hadPriorSignOffLog) {
      signOffSheet.getRange('F11').setValue('No');
    } else {
      signOffSheet.getRange('F11').clearContent();
    }

    signOffSheet.getRange(SIGN_OFF_LATEST_LOG_STATUS_CELL).setValue('Yes');
    syncTabColorsToSignOffStatus_(ss);

    progressStep_(ss, '📸', 'Building hidden Sign-off snapshot for export / AI…', PROGRESS_STEP_PAUSE_MS);

    writeSignOffSnapshotToSheet_(ss);

    try {
      ui.alert(
        'Sign off in progress',
        'Sign off is complete and the Accounting Manual is now generating. ' +
          'You can safely close this spreadsheet/tab now — the process will continue and update the file when finished.',
        ui.ButtonSet.OK
      );
    } catch (e) {
      // ignore
    }

    progressStep_(
      ss,
      '🤖',
      'Generating accounting manual (OpenAI)… this may take 1–2 minutes.',
      PROGRESS_STEP_PAUSE_MS
    );
    try {
      generateAccountingManualAfterSignOff_(ss);
    } catch (openAiErr) {
      ss.toast(
        'Sign-off saved; OpenAI manual failed: ' + String(openAiErr.message || openAiErr),
        MENU_NAME,
        12
      );
      ui.alert(
        'OpenAI accounting manual',
        'Sign-off completed, but generating the Accounting manual tab failed:\n\n' +
          String(openAiErr.message || openAiErr),
        ui.ButtonSet.OK
      );
    }

    snapshotSheetGuardState_(ss);
  } finally {
    PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
  }

  ss.setActiveSheet(signOffSheet);
  signOffSheet.getRange('A1').activate();
  progressDone_(ss, '✅', 'Sign off complete — tabs updated, snapshot saved.');
}

// ------------------------------------------------------------------
// Sign-off snapshot (sheet) + plain-text export
// ------------------------------------------------------------------

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

function writeSignOffSnapshotToSheet_(ss) {
  const text = buildSignOffSnapshotText_(ss);
  let sh = ss.getSheetByName(SIGN_OFF_SNAPSHOT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet();
  if (sh.getName() !== SIGN_OFF_SNAPSHOT_SHEET_NAME) {
    try {
      sh.setName(SIGN_OFF_SNAPSHOT_SHEET_NAME);
    } catch (e) {
      // best-effort rename
    }
  }
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
      r.setBorder(false, false, false, false, false, false, null, null);
    }
  }
  syncTabColorsToSignOffStatus_(ss);
}

function resetErrorFormattingAllSheets_(ss) {
  for (const sh of ss.getSheets()) {
    sh.setTabColor(null);
    for (const r of getUnprotectedRanges_(sh)) {
      r.setBorder(false, false, false, false, false, false, null, null);
    }
  }
}

// ------------------------------------------------------------------
// Sheet guard / sign-off status
// ------------------------------------------------------------------

function ensureSheetGuardInstalled_(ss) {
  if (!PropertiesService.getDocumentProperties().getProperty(SHEET_GUARD_STATE_KEY)) {
    snapshotSheetGuardState_(ss);
  }
}

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
    deleteAccountingManualSheetIfExists_(ss);
    syncTabColorsToSignOffStatus_(ss);
  } finally {
    PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
  }
}

function deleteAccountingManualSheetIfExists_(ss) {
  const sh = ss.getSheetByName(SHEET_ACCOUNTING_MANUAL_OUTPUT);
  if (sh) ss.deleteSheet(sh);
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

  for (const d of desired) {
    const sh = currentById.get(d.id);
    if (!sh) {
      if (d.name === SIGN_OFF_SNAPSHOT_SHEET_NAME && isSignedOff_(ss)) {
        if (PropertiesService.getDocumentProperties().getProperty(SKIP_SIGN_OFF_EDIT_KEY)) {
          return;
        }
        PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
        try {
          const so = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
          if (so) so.getRange(SIGN_OFF_LATEST_LOG_STATUS_CELL).setValue('No');
          deleteAccountingManualSheetIfExists_(ss);
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
      const recreated = ss.insertSheet();
      try {
        recreated.setName(d.name);
      } catch (e2) {
        // best-effort
      }
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

// ------------------------------------------------------------------
// Validation + protection-aware helpers
// ------------------------------------------------------------------

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
  }

  return { failures, firstFailure };
}

function validateRange_(range) {
  const sh = range.getSheet();
  const displayValues = range.getDisplayValues();
  const validations = range.getDataValidations();
  const mergedRanges = range.getMergedRanges();

  const badCells = [];
  let firstFailureCell = null;

  for (const mr of mergedRanges) {
    const tl = mr.getCell(1, 1);
    const tlBg = String(tl.getBackground() || '').toLowerCase();

    const mrValidations = mr.getDataValidations();
    const hasAnyValidation = mrValidations.some((row) => row.some((v) => !!v));
    const isRequiredMerged = hasAnyValidation || tlBg === REQUIRED_INPUT_FILL;
    if (!isRequiredMerged) continue;

    const display = String(tl.getDisplayValue() || '').trim();
    if (!display) {
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

      const dv = String(displayValues[r][c] || '').trim();
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
  if (rect.r1 < i.r1) out.push({ r1: rect.r1, c1: rect.c1, r2: i.r1 - 1, c2: rect.c2 });
  if (i.r2 < rect.r2) out.push({ r1: i.r2 + 1, c1: rect.c1, r2: rect.r2, c2: rect.c2 });
  if (rect.c1 < i.c1) out.push({ r1: i.r1, c1: rect.c1, r2: i.r2, c2: i.c1 - 1 });
  if (i.c2 < rect.c2) out.push({ r1: i.r1, c1: i.c2 + 1, r2: i.r2, c2: rect.c2 });
  return out;
}

function resetUnprotectedContent_(ss, opts) {
  const excludeSheetNames = (opts && opts.excludeSheetNames) || new Set();

  for (const sh of ss.getSheets()) {
    if (excludeSheetNames.has(sh.getName())) continue;

    const sheetProtections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (sheetProtections.length > 0) {
      const p = sheetProtections[0];
      const unprotected = p.getUnprotectedRanges() || [];
      for (const r of unprotected) {
        r.clearContent();
      }
      continue;
    }

    const protectedRanges = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    if (protectedRanges.length > 0) {
      const snapshots = protectedRanges.map((prot) => snapshotRange_(prot.getRange()));

      const dataRange = sh.getDataRange();
      dataRange.clearContent();

      for (const snap of snapshots) restoreRangeSnapshot_(snap);
      continue;
    }

    sh.getDataRange().clearContent();
  }
}

function snapshotRange_(range) {
  return {
    a1: range.getA1Notation(),
    sheetId: range.getSheet().getSheetId(),
    values: range.getValues(),
    formulasR1C1: range.getFormulasR1C1(),
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

  const parts = raw
    .split(/[\n,;]+/g)
    .map((s) => s.trim())
    .filter(Boolean);

  const out = [];
  const seen = new Set();
  for (const p of parts) {
    if (seen.has(p)) continue;
    seen.add(p);
    out.push(p);
  }
  return out;
}

// ------------------------------------------------------------------
// OpenAI + Google Doc generation
// ------------------------------------------------------------------

function getOpenAIPromptVarName_(propertyName, defaultName) {
  const v = PropertiesService.getScriptProperties().getProperty(propertyName);
  const s = (v && String(v).trim()) || defaultName;
  return String(s).trim();
}

function getBusinessSummaryClientRangeA1_() {
  const props = PropertiesService.getScriptProperties();
  const v =
    props.getProperty(OPENAI_CLIENT_RANGE_A1_PROPERTY) ||
    props.getProperty(OPENAI_SIGN_OFF_CLIENT_RANGE_LEGACY_PROPERTY);
  const s = (v && String(v).trim()) || RANGE_BUSINESS_SUMMARY_CLIENT;
  return String(s).trim();
}

function readBusinessSummaryClient_(ss) {
  const bs = ss.getSheetByName(SHEET_BUSINESS_SUMMARY);
  if (!bs) return '';
  const a1 = getBusinessSummaryClientRangeA1_();
  return readMergedDisplayValue_(bs.getRange(a1));
}

function readMergedDisplayValue_(range) {
  try {
    const dv = range.getDisplayValue();
    if (dv != null && String(dv).trim() !== '') return String(dv).trim();
  } catch (e) {
    // ignore
  }
  try {
    const v = range.getValue();
    if (v != null && String(v).trim() !== '') return String(v).trim();
  } catch (e2) {
    // ignore
  }
  const grid = range.getDisplayValues();
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const s = String(grid[r][c] || '').trim();
      if (s) return s;
    }
  }
  return '';
}

function generateAccountingManualAfterSignOff_(ss) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROPERTY);
  if (!apiKey || !String(apiKey).trim()) {
    ss.toast(
      'Set library Script property OPENAI_API_KEY (apps-script-lib → Project Settings → Script properties) to generate the Accounting manual.',
      MENU_NAME,
      12
    );
    return;
  }

  const signOffSh = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!signOffSh) throw new Error('Sheet "Sign off Page" not found.');

  const client = readBusinessSummaryClient_(ss);
  if (!client) {
    ss.toast(
      'Client name was blank on Business Summary (' +
        RANGE_BUSINESS_SUMMARY_CLIENT +
        '). Set OPENAI_CLIENT_RANGE_A1 if your client cell differs.',
      MENU_NAME,
      12
    );
  }

  progressStep_(ss, '📦', 'Preparing sign-off snapshot for the Accounting Manual…', PROGRESS_STEP_PAUSE_MS);
  const signOffSnapshotText = buildSignOffSnapshotText_(ss);

  progressStep_(ss, '🤖', 'Calling OpenAI…', PROGRESS_STEP_PAUSE_MS);
  const manualText = callOpenAIAccountingManual_(String(apiKey).trim(), client, signOffSnapshotText);

  let parsedManualJson = null;
  try {
    parsedManualJson = JSON.parse(String(manualText || ''));
  } catch (e) {
    parsedManualJson = null;
  }

  const clientForDoc = resolveClientNameForAccountingManualDoc_(ss, client, parsedManualJson);
  const logVersion = getSignOffLogCount_(signOffSh);

  progressStep_(ss, '🧾', 'Saving OpenAI response to Accounting manual sheet…', PROGRESS_STEP_PAUSE_MS);
  writeLongTextToSheetChunked_(ss, SHEET_ACCOUNTING_MANUAL_OUTPUT, manualText);
  const outSh = ss.getSheetByName(SHEET_ACCOUNTING_MANUAL_OUTPUT);
  if (outSh) outSh.showSheet();

  progressStep_(ss, '📄', 'Building Accounting Manual Google Doc…', PROGRESS_STEP_PAUSE_MS);
  const docId = createAccountingManualDocFromOpenAIJson_(ss, clientForDoc, manualText, logVersion);
  if (docId) {
    const url = 'https://docs.google.com/document/d/' + docId + '/edit';
    setSignOffAccountingManualLink_(ss, url);
    ss.toast('Accounting manual doc created: ' + url, MENU_NAME, 12);
  } else {
    ss.toast('Accounting manual doc was not created (no doc id).', MENU_NAME, 10);
  }
}

function getSignOffLogCount_(signOffSh) {
  const lr = signOffSh.getLastRow();
  if (lr < 10) return 1;
  return lr - 9;
}

function resolveClientNameForAccountingManualDoc_(ss, sheetClient, parsedJson) {
  if (parsedJson && typeof parsedJson === 'object') {
    const j = parsedJson;
    const meta = j.metadata && typeof j.metadata === 'object' ? j.metadata : null;
    const fromJson = String(j.client || j.clientName || (meta && meta.client) || '').trim();
    if (fromJson) return fromJson;
  }
  const sc = String(sheetClient || '').trim();
  if (sc) return sc;
  return resolveClientNameFromBusinessSummaryCells_(ss);
}

function resolveClientNameFromBusinessSummaryCells_(ss) {
  const bs = ss.getSheetByName(SHEET_BUSINESS_SUMMARY);
  if (!bs) return '';
  const a1 = getBusinessSummaryClientRangeA1_();
  const merged = readMergedDisplayValue_(bs.getRange(a1));
  if (merged) return merged;
  const c7d = String(bs.getRange('C7').getDisplayValue() || '').trim();
  if (c7d) return c7d;
  const c7v = String(bs.getRange('C7').getValue() || '').trim();
  if (c7v) return c7v;
  for (let col = 3; col <= 8; col++) {
    const v = String(bs.getRange(7, col).getDisplayValue() || '').trim();
    if (v) return v;
  }
  return '';
}

function setSignOffAccountingManualLink_(ss, url) {
  const sh = ss.getSheetByName(SHEET_SIGN_OFF_PAGE);
  if (!sh) return;
  const safeUrl = String(url || '').replace(/"/g, '""');
  PropertiesService.getDocumentProperties().setProperty(SKIP_SIGN_OFF_EDIT_KEY, '1');
  try {
    sh.getRange('E10').setFormula('=HYPERLINK("' + safeUrl + '","Accounting Manual")');
  } finally {
    PropertiesService.getDocumentProperties().deleteProperty(SKIP_SIGN_OFF_EDIT_KEY);
  }
}

function formatAccountingManualDocName_(ss, client, version) {
  const tz = ss.getSpreadsheetTimeZone();
  const date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const safeClient = String(client || '').trim() || 'Unknown client';
  return `Accounting Manual - ${safeClient} - V${version} - ${date}`;
}

function getSpreadsheetParentFolder_(ss) {
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  if (parents && parents.hasNext()) return parents.next();
  return null;
}

function moveDocToSpreadsheetFolder_(ss, docId) {
  const folder = getSpreadsheetParentFolder_(ss);
  if (!folder) return;
  const docFile = DriveApp.getFileById(docId);
  docFile.moveTo(folder);
}

function createAccountingManualDocFromOpenAIJson_(ss, client, jsonText, version) {
  const name = formatAccountingManualDocName_(ss, client, version);

  let data;
  try {
    data = JSON.parse(String(jsonText || ''));
  } catch (e) {
    throw new Error(
      'OpenAI did not return valid JSON for the doc builder. First 500 chars: ' +
        String(jsonText || '').substring(0, 500)
    );
  }

  const doc = DocumentApp.create(name);
  const body = doc.getBody();
  body.clear();

  const title = (data && data.title && String(data.title).trim()) || name;
  const titlePara = body.appendParagraph(title);
  titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);

  const blocks = (data && (data.blocks || data.sections)) || [];
  renderAccountingManualBlocksToDocBody_(body, blocks);

  doc.saveAndClose();

  moveDocToSpreadsheetFolder_(ss, doc.getId());
  return doc.getId();
}

function renderAccountingManualBlocksToDocBody_(body, blocks) {
  if (!blocks || !blocks.length) return;

  for (const b of blocks) {
    if (!b) continue;

    if (b.heading) {
      appendHeading_(body, b.heading.level, b.heading.text);
      continue;
    }
    if (b.type === 'heading') {
      appendHeading_(body, b.level, b.text);
      continue;
    }

    if (b.paragraph != null) {
      body.appendParagraph(String(b.paragraph));
      continue;
    }
    if (b.type === 'paragraph') {
      body.appendParagraph(String(b.text || ''));
      continue;
    }

    const bullets = b.bullets || b.items;
    if (Array.isArray(bullets)) {
      for (const it of bullets) body.appendListItem(String(it)).setGlyphType(DocumentApp.GlyphType.BULLET);
      continue;
    }
    if (b.type === 'bullets' && Array.isArray(b.items)) {
      for (const it of b.items) body.appendListItem(String(it)).setGlyphType(DocumentApp.GlyphType.BULLET);
      continue;
    }

    const table = b.table || (b.type === 'table' ? b : null);
    if (table) {
      const headers = Array.isArray(table.headers) ? table.headers.map((x) => String(x)) : null;
      const rows = Array.isArray(table.rows) ? table.rows : [];
      const allRows = [];
      if (headers && headers.length) allRows.push(headers);
      for (const r of rows) {
        if (!Array.isArray(r)) continue;
        allRows.push(r.map((x) => String(x)));
      }
      if (allRows.length) body.appendTable(allRows);
      continue;
    }

    body.appendParagraph(String(typeof b === 'string' ? b : JSON.stringify(b)));
  }
}

function appendHeading_(body, level, text) {
  const t = String(text || '').trim();
  if (!t) return;
  const lvl = parseInt(level, 10);
  let heading = DocumentApp.ParagraphHeading.HEADING2;
  if (lvl === 1) heading = DocumentApp.ParagraphHeading.HEADING1;
  else if (lvl === 2) heading = DocumentApp.ParagraphHeading.HEADING2;
  else if (lvl === 3) heading = DocumentApp.ParagraphHeading.HEADING3;
  else if (lvl === 4) heading = DocumentApp.ParagraphHeading.HEADING4;
  body.appendParagraph(t).setHeading(heading);
}

function getOpenAIPromptId_() {
  const override = PropertiesService.getScriptProperties().getProperty(OPENAI_PROMPT_ID_PROPERTY);
  const id = (override && String(override).trim()) || OPENAI_ACCOUNTING_MANUAL_PROMPT_ID;
  return String(id).trim();
}

function callOpenAIAccountingManual_(apiKey, client, signOffSnapshotText) {
  const promptId = getOpenAIPromptId_();
  const clientVar = getOpenAIPromptVarName_(OPENAI_PROMPT_VAR_CLIENT_PROPERTY, OPENAI_PROMPT_VAR_CLIENT_DEFAULT);
  const snapshotVar = getOpenAIPromptVarName_(
    OPENAI_PROMPT_VAR_SNAPSHOT_PROPERTY,
    OPENAI_PROMPT_VAR_SNAPSHOT_DEFAULT
  );

  const variables = {};
  variables[clientVar] = client;
  variables[snapshotVar] = signOffSnapshotText;

  const payload = {
    prompt: {
      id: promptId,
      variables: variables
    }
  };

  const res = UrlFetchApp.fetch(OPENAI_RESPONSES_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    let hint = '';
    if (code === 404) {
      hint =
        '\n\n(404: confirm the prompt id in OpenAI Platform → Prompts (or Storage → Prompts), copy the full id, ' +
        'and set library Script property OPENAI_ACCOUNTING_MANUAL_PROMPT_ID to that value. The id must exist in the same ' +
        'organization as your API key.)';
    }
    throw new Error('OpenAI HTTP ' + code + ': ' + body.substring(0, 800) + hint);
  }

  let parsed;
  try {
    parsed = JSON.parse(body);
  } catch (e) {
    throw new Error('OpenAI returned non-JSON: ' + body.substring(0, 400));
  }

  if (parsed.error) {
    throw new Error(parsed.error.message || JSON.stringify(parsed.error));
  }

  const out = extractResponsesApiOutputText_(parsed);
  if (!out) throw new Error('OpenAI returned no assistant text. Raw keys: ' + Object.keys(parsed).join(', '));
  return out;
}

function extractResponsesApiOutputText_(parsed) {
  if (parsed.output_text) return parsed.output_text;
  const items = parsed.output;
  if (!items || !items.length) return '';
  const chunks = [];
  for (const item of items) {
    if (item.type === 'message' && item.content) {
      for (const part of item.content) {
        if (part.type === 'output_text' && part.text) chunks.push(part.text);
      }
    }
  }
  return chunks.join('');
}

function writeLongTextToSheetChunked_(ss, sheetName, text) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clearContents();
  let row = 1;
  for (let i = 0; i < text.length; i += SNAPSHOT_CHUNK_CHARS) {
    const chunk = text.substring(i, Math.min(i + SNAPSHOT_CHUNK_CHARS, text.length));
    sh.getRange(row, 1).setValue(chunk);
    row++;
  }
}

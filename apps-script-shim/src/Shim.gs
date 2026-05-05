/**
 * Accounting Manual — bound shim.
 *
 * This file is auto-attached to each client spreadsheet by the internal Editor Add-on.
 * All real logic lives in the library (AccountingLib). The shim only:
 *   - Builds the spreadsheet menu (simple onOpen).
 *   - Hosts installable trigger handlers required to live in the bound project.
 *   - Lazily installs the two installable triggers on first authorized action.
 */

function onOpen(e) {
  return AccountingLib.onOpen(e);
}

function signOff() {
  ensureTriggersInstalled_();
  return AccountingLib.signOff();
}

function generateWorkflow() {
  ensureTriggersInstalled_();
  return AccountingLib.generateWorkflow();
}

function reinitiatedSpreadsheet() {
  return AccountingLib.reinitiatedSpreadsheet();
}

function onSheetChange_(e) {
  return AccountingLib.onSheetChange(e);
}

function onSignedOffEdit_(e) {
  return AccountingLib.onSignedOffEdit(e);
}

/**
 * Install missing installable triggers on this spreadsheet (idempotent).
 * Triggers must live in the bound script project, so they are created here, not in the library.
 */
function ensureTriggersInstalled_() {
  const triggers = ScriptApp.getProjectTriggers();
  const handlers = new Set(triggers.map(function (t) {
    return t.getHandlerFunction();
  }));
  const ss = SpreadsheetApp.getActive();
  if (!handlers.has('onSheetChange_')) {
    ScriptApp.newTrigger('onSheetChange_').forSpreadsheet(ss).onChange().create();
  }
  if (!handlers.has('onSignedOffEdit_')) {
    ScriptApp.newTrigger('onSignedOffEdit_').forSpreadsheet(ss).onEdit().create();
  }
}

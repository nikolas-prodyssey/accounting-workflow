/**
 * The shim source pushed into each new bound script project.
 * Mirrors apps-script-shim/src/Shim.gs — keep in sync when changing.
 */

const SHIM_TEMPLATE_APPSSCRIPT_JSON = JSON.stringify(
  {
    timeZone: 'Asia/Nicosia',
    dependencies: {
      libraries: [
        {
          userSymbol: 'AccountingLib',
          libraryId: ACCOUNTING_LIB_SCRIPT_ID,
          version: ACCOUNTING_LIB_VERSION,
          developmentMode: false
        }
      ]
    },
    exceptionLogging: 'STACKDRIVER',
    runtimeVersion: 'V8',
    oauthScopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/script.container.ui',
      'https://www.googleapis.com/auth/script.scriptapp',
      'https://www.googleapis.com/auth/userinfo.email',
      'https://www.googleapis.com/auth/script.external_request',
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/documents'
    ]
  },
  null,
  2
);

const SHIM_TEMPLATE_CODE_GS = [
  '/**',
  ' * Accounting Manual — bound shim (auto-attached by the internal Editor Add-on).',
  ' * All real logic is in the AccountingLib library.',
  ' */',
  '',
  'function onOpen(e) {',
  '  return AccountingLib.onOpen(e);',
  '}',
  '',
  'function signOff() {',
  '  ensureTriggersInstalled_();',
  '  return AccountingLib.signOff();',
  '}',
  '',
  'function generateWorkflow() {',
  '  ensureTriggersInstalled_();',
  '  return AccountingLib.generateWorkflow();',
  '}',
  '',
  'function reinitiatedSpreadsheet() {',
  '  return AccountingLib.reinitiatedSpreadsheet();',
  '}',
  '',
  'function onSheetChange_(e) {',
  '  return AccountingLib.onSheetChange(e);',
  '}',
  '',
  'function onSignedOffEdit_(e) {',
  '  return AccountingLib.onSignedOffEdit(e);',
  '}',
  '',
  'function ensureTriggersInstalled_() {',
  '  var triggers = ScriptApp.getProjectTriggers();',
  '  var handlers = {};',
  '  for (var i = 0; i < triggers.length; i++) handlers[triggers[i].getHandlerFunction()] = true;',
  '  var ss = SpreadsheetApp.getActive();',
  "  if (!handlers['onSheetChange_']) {",
  "    ScriptApp.newTrigger('onSheetChange_').forSpreadsheet(ss).onChange().create();",
  '  }',
  "  if (!handlers['onSignedOffEdit_']) {",
  "    ScriptApp.newTrigger('onSignedOffEdit_').forSpreadsheet(ss).onEdit().create();",
  '  }',
  '}',
  ''
].join('\n');

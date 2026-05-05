/**
 * CardService UI for the internal Editor Add-on.
 *
 * Two flows:
 *   1) Create new client workbook (copies the master, attaches shim, runs bootstrap).
 *   2) Attach / upgrade plugin on the currently open spreadsheet.
 */

function buildHomepageCard(e) {
  if (!isAllowedDomain_()) return buildDeniedCard_();

  const builder = CardService.newCardBuilder().setHeader(
    CardService.newCardHeader()
      .setTitle('Accounting Manual')
      .setSubtitle('Internal provisioner')
  );

  const intro = CardService.newCardSection().addWidget(
    CardService.newTextParagraph().setText(
      'Use this add-on to create or upgrade Accounting Manual workbooks for new clients.'
    )
  );

  const createSection = CardService.newCardSection()
    .setHeader('Create new client workbook')
    .addWidget(
      CardService.newTextInput()
        .setFieldName('clientName')
        .setTitle('Client name')
        .setHint('Used in spreadsheet name')
    )
    .addWidget(
      CardService.newTextInput()
        .setFieldName('destinationFolderId')
        .setTitle('Destination folder ID')
        .setHint('Drive folder ID where the new workbook will live')
    )
    .addWidget(
      CardService.newTextButton()
        .setText('Create workbook')
        .setOnClickAction(CardService.newAction().setFunctionName('handleCreateClientWorkbook'))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    );

  const attachSection = CardService.newCardSection()
    .setHeader('Attach / upgrade plugin on this workbook')
    .addWidget(
      CardService.newTextParagraph().setText(
        'Open the spreadsheet you want to attach the plugin to, then click below.'
      )
    )
    .addWidget(
      CardService.newTextButton()
        .setText('Attach to current spreadsheet')
        .setOnClickAction(CardService.newAction().setFunctionName('handleAttachToCurrent'))
    );

  builder.addSection(intro);
  builder.addSection(createSection);
  builder.addSection(attachSection);
  return [builder.build()];
}

function buildDeniedCard_() {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Access denied'))
    .addSection(
      CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText(
          'This add-on is restricted to internal users.'
        )
      )
    )
    .build();
  return [card];
}

function buildResultCard_(title, lines) {
  const section = CardService.newCardSection();
  for (const line of lines) {
    section.addWidget(CardService.newTextParagraph().setText(line));
  }
  section.addWidget(
    CardService.newTextButton()
      .setText('Back')
      .setOnClickAction(CardService.newAction().setFunctionName('buildHomepageCard'))
  );
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(title))
    .addSection(section)
    .build();
}

function linkifyLabel_(label, url, text) {
  const safeUrl = String(url || '').trim();
  if (!safeUrl) return label + ': (missing)';
  const safeText = String(text || safeUrl).replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return label + ': <a href="' + safeUrl + '">' + safeText + '</a>';
}

// ------------------------------------------------------------------
// Action handlers
// ------------------------------------------------------------------

function handleCreateClientWorkbook(e) {
  if (!isAllowedDomain_()) return navigateToCard_(buildDeniedCard_());

  const inputs = (e && e.commonEventObject && e.commonEventObject.formInputs) || {};
  const client = readFormValue_(inputs, 'clientName');
  const folderId = readFormValue_(inputs, 'destinationFolderId');

  if (!client) {
    return navigateToCard_(buildResultCard_('Missing client name', ['Please enter the client name.']));
  }
  if (!folderId) {
    return navigateToCard_(
      buildResultCard_('Missing folder', [
        'Please paste the Drive folder ID where the new workbook should be created.'
      ])
    );
  }

  let result;
  try {
    result = provisionNewClientWorkbook(client, folderId);
  } catch (err) {
    return navigateToCard_(
      buildResultCard_('Provisioning failed', [String((err && err.message) || err)])
    );
  }

  const bootLines =
    result.bootstrapped
      ? ['Bootstrap (re-init): OK']
      : [
          'Bootstrap (re-init): Skipped' +
            (result.bootstrapNote ? ' — ' + result.bootstrapNote : '')
        ];

  return navigateToCard_(
    buildResultCard_('Workbook created', [
      linkifyLabel_('Spreadsheet', result.spreadsheetUrl, 'Open spreadsheet'),
      'Bound script id: ' + result.boundScriptId,
      'Library version: ' + ACCOUNTING_LIB_VERSION,
      ...bootLines,
      ...formatStaleCleanupLines_(result)
    ])
  );
}

function handleAttachToCurrent(e) {
  if (!isAllowedDomain_()) return navigateToCard_(buildDeniedCard_());

  const ssId = e && e.commonEventObject && e.commonEventObject.parameters
    ? e.commonEventObject.parameters.spreadsheetId
    : null;

  let activeId = ssId;
  try {
    if (!activeId) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) activeId = ss.getId();
    }
  } catch (err) {
    activeId = null;
  }

  if (!activeId) {
    return navigateToCard_(
      buildResultCard_('No active spreadsheet', [
        'Open a Google Sheet first, then run "Attach to current spreadsheet".'
      ])
    );
  }

  let result;
  try {
    // Same re-init as after a master copy: only succeeds if this file is an Accounting Manual
    // template (e.g. it has a "Business Summary" sheet). Blank or random sheets will skip
    // bootstrap; use Create new client workbook to copy the master.
    result = attachOrUpgradeShimOnSpreadsheet(activeId, { runBootstrap: true });
  } catch (err) {
    const msg = String((err && err.message) || err || '');
    return navigateToCard_(
      buildResultCard_('Attach failed', [msg])
    );
  }

  const lines = [
    'Spreadsheet id: ' + activeId,
    'Bound script id: ' + result.boundScriptId,
    'Library version pinned: ' + ACCOUNTING_LIB_VERSION,
    result.created ? 'Created new bound script.' : 'Updated existing bound script.'
  ];
  if (result.bootstrapped) {
    lines.push('Bootstrap (re-init): OK — workbook initialized.');
  } else if (result.bootstrapNote) {
    lines.push('Bootstrap (re-init): not run — ' + result.bootstrapNote);
    lines.push(
      'This file does not look like an Accounting Manual workbook copied from the master (it needs a "Business Summary" sheet). Use Create new client workbook to copy the master, or duplicate the master spreadsheet in Drive — Attach only installs/upgrades the plugin.'
    );
  }

  for (const l of formatStaleCleanupLines_(result)) lines.push(l);

  return navigateToCard_(buildResultCard_('Plugin attached / upgraded', lines));
}

function formatStaleCleanupLines_(result) {
  const stale = (result && result.staleNames) || [];
  if (!stale.length) return [];
  const neutralized = result.neutralizedCount || 0;
  const trashed = result.trashedCount || 0;
  return [
    'Removed ' + stale.length + ' inherited bound script(s) so only the shim runs:',
    '• ' + stale.join(', '),
    'Neutralized: ' + neutralized + ' / ' + stale.length + ' · Trashed: ' + trashed + ' / ' + stale.length
  ];
}

// ------------------------------------------------------------------
// Helpers
// ------------------------------------------------------------------

function navigateToCard_(card) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card))
    .build();
}

function readFormValue_(inputs, name) {
  const f = inputs[name];
  if (!f || !f.stringInputs || !f.stringInputs.value || !f.stringInputs.value.length) return '';
  return String(f.stringInputs.value[0] || '').trim();
}

function isAllowedDomain_() {
  try {
    const email = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
    if (!email) return false;
    for (const d of ALLOWED_DOMAINS) {
      if (email.endsWith('@' + d)) return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}

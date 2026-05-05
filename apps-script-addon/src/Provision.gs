/**
 * Provisioning helpers used by the add-on UI.
 *
 * - provisionNewClientWorkbook(clientName, folderId): copy master, attach shim, run bootstrap.
 * - attachOrUpgradeShimOnSpreadsheet(spreadsheetId, opts): attach a fresh bound shim or upgrade
 *   the existing bound script's shim files (idempotent).
 */

/**
 * @param {string} clientName
 * @param {string} folderId
 * @returns {{spreadsheetId: string, spreadsheetUrl: string, boundScriptId: string, bootstrapped: boolean, bootstrapNote: string}}
 */
function provisionNewClientWorkbook(clientName, folderId) {
  if (!clientName || !folderId) throw new Error('clientName and folderId are required.');

  const newName = formatClientWorkbookName_(clientName);
  const newFile = copyMasterSpreadsheet_(MASTER_SPREADSHEET_ID, folderId, newName);
  const newSpreadsheetId = newFile.id;

  const attachResult = attachOrUpgradeShimOnSpreadsheet(newSpreadsheetId, { runBootstrap: true });

  return {
    spreadsheetId: newSpreadsheetId,
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + newSpreadsheetId + '/edit',
    boundScriptId: attachResult.boundScriptId,
    bootstrapped: attachResult.bootstrapped,
    bootstrapNote: attachResult.bootstrapNote || '',
    neutralizedCount: attachResult.neutralizedCount || 0,
    trashedCount: attachResult.trashedCount || 0,
    staleNames: attachResult.staleNames || []
  };
}

function formatClientWorkbookName_(clientName) {
  const tz = Session.getScriptTimeZone() || 'Etc/UTC';
  const date = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return 'Accounting Manual - ' + String(clientName).trim() + ' - ' + date;
}

/**
 * @param {string} spreadsheetId
 * @param {{runBootstrap?: boolean}} [opts]
 * @returns {{boundScriptId: string, created: boolean, bootstrapped: boolean, bootstrapNote: string, neutralizedCount: number, trashedCount: number, staleNames: string[]}}
 */
function attachOrUpgradeShimOnSpreadsheet(spreadsheetId, opts) {
  opts = opts || {};

  // 1. Find or create the bound shim project, identified by title (BOUND_PROJECT_TITLE).
  //    We deliberately do NOT reuse arbitrary bound scripts on the workbook — copies
  //    of the master inherit a stale bound script (e.g. "Accounting-Workflow") that
  //    must be neutralized, not converted into the shim.
  let boundScriptId = findBoundScriptIdByTitle_(spreadsheetId, BOUND_PROJECT_TITLE);
  let created = false;
  if (!boundScriptId) {
    boundScriptId = createBoundScriptProject_(spreadsheetId, BOUND_PROJECT_TITLE);
    created = true;
  }
  pushShimContent_(boundScriptId);

  // 2. Neutralize every other bound script attached to this spreadsheet so only our
  //    shim runs (no duplicate `📒 Accounting` menus, no old monolithic logic).
  const stale = listOtherBoundScripts_(spreadsheetId, boundScriptId);
  const staleNames = [];
  let neutralizedCount = 0;
  let trashedCount = 0;
  for (const f of stale) {
    staleNames.push(f.name);
    try {
      neutralizeBoundScript_(f.id);
      neutralizedCount++;
    } catch (err) {
      console.error('Failed to neutralize bound script ' + f.id + ': ' + (err && err.message ? err.message : err));
    }
    try {
      trashDriveFile_(f.id);
      trashedCount++;
    } catch (errTrash) {
      // Trashing is best-effort — neutralization above is the real safety net.
      console.error('Failed to trash bound script ' + f.id + ': ' + (errTrash && errTrash.message ? errTrash.message : errTrash));
    }
  }

  let bootstrapped = false;
  let bootstrapNote = '';
  if (opts.runBootstrap) {
    try {
      const ss = SpreadsheetApp.openById(spreadsheetId);
      AccountingLib.bootstrap(ss);
      bootstrapped = true;
    } catch (err) {
      bootstrapNote = String((err && err.message) || err || '');
      console.error('Bootstrap failed: ' + bootstrapNote);
    }
  }

  return {
    boundScriptId: boundScriptId,
    created: created,
    bootstrapped: bootstrapped,
    bootstrapNote: bootstrapNote,
    neutralizedCount: neutralizedCount,
    trashedCount: trashedCount,
    staleNames: staleNames
  };
}

// ------------------------------------------------------------------
// Drive (copy master)
// ------------------------------------------------------------------

function copyMasterSpreadsheet_(masterFileId, parentFolderId, newName) {
  const url =
    'https://www.googleapis.com/drive/v3/files/' +
    encodeURIComponent(masterFileId) +
    '/copy?supportsAllDrives=true&fields=id,name,parents';

  const payload = { name: newName, parents: [parentFolderId] };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Drive copy HTTP ' + code + ': ' + body.substring(0, 800));
  }
  return JSON.parse(body);
}

// ------------------------------------------------------------------
// Apps Script API (create + update bound project)
// ------------------------------------------------------------------

/** Returns the FIRST bound script id for a Google Sheet (any title), or null if none exists. */
function findBoundScriptId_(spreadsheetId) {
  const list = listBoundScripts_(spreadsheetId);
  return list.length ? list[0].id : null;
}

/** Returns the bound script id for a Google Sheet whose project title matches `title`, or null. */
function findBoundScriptIdByTitle_(spreadsheetId, title) {
  const list = listBoundScripts_(spreadsheetId);
  const target = String(title || '').trim();
  for (const f of list) {
    if (String(f.name || '').trim() === target) return f.id;
  }
  return null;
}

/** Returns every bound script attached to the spreadsheet EXCEPT the one matching `excludeId`. */
function listOtherBoundScripts_(spreadsheetId, excludeId) {
  const list = listBoundScripts_(spreadsheetId);
  return list.filter((f) => f.id !== excludeId);
}

/** Lists all (non-trashed) bound script projects attached to a Google Sheet. */
function listBoundScripts_(spreadsheetId) {
  // Drive search: bound scripts are Drive files of mimeType=application/vnd.google-apps.script
  // with the host spreadsheet listed in their `parents` field.
  const q = encodeURIComponent(
    "mimeType='application/vnd.google-apps.script' and '" + spreadsheetId + "' in parents and trashed=false"
  );
  const url =
    'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&includeItemsFromAllDrives=true&q=' +
    q +
    '&fields=files(id,name)';
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Drive search HTTP ' + code + ': ' + body.substring(0, 800));
  }
  const parsed = JSON.parse(body);
  return (parsed && parsed.files) || [];
}

/**
 * Neutralizes a bound script project so that nothing it contains will ever execute again
 * (no `onOpen`, no triggers, no menu). We replace its content with an inert manifest +
 * a single empty-comment file.
 */
function neutralizeBoundScript_(scriptId) {
  const url =
    'https://script.googleapis.com/v1/projects/' +
    encodeURIComponent(scriptId) +
    '/content';
  const inertManifest = JSON.stringify(
    {
      timeZone: 'Asia/Nicosia',
      exceptionLogging: 'STACKDRIVER',
      runtimeVersion: 'V8'
    },
    null,
    2
  );
  const inertCode =
    '// Deactivated by Accounting Manual Provisioner.\n' +
    '// All logic now lives in the AccountingLib library, attached via the bound shim.\n';
  const payload = {
    files: [
      { name: 'appsscript', type: 'JSON', source: inertManifest },
      { name: 'Code', type: 'SERVER_JS', source: inertCode }
    ]
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Apps Script neutralize content HTTP ' + code + ': ' + body.substring(0, 600));
  }
}

/** Best-effort move-to-trash for a Drive file (used to also remove neutralized script projects). */
function trashDriveFile_(fileId) {
  const url =
    'https://www.googleapis.com/drive/v3/files/' +
    encodeURIComponent(fileId) +
    '?supportsAllDrives=true';
  const res = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ trashed: true }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Drive trash HTTP ' + code + ': ' + body.substring(0, 600));
  }
}

function createBoundScriptProject_(spreadsheetId, title) {
  assertDriveFileAccessible_(spreadsheetId);
  const url = 'https://script.googleapis.com/v1/projects';
  const payload = { title: title || BOUND_PROJECT_TITLE, parentId: spreadsheetId };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Apps Script API create HTTP ' + code + ': ' + body.substring(0, 800));
  }
  const parsed = JSON.parse(body);
  if (!parsed || !parsed.scriptId) throw new Error('Apps Script API create returned no scriptId.');
  return parsed.scriptId;
}

function assertDriveFileAccessible_(fileId) {
  if (!fileId) throw new Error('Missing spreadsheet id.');
  const url =
    'https://www.googleapis.com/drive/v3/files/' +
    encodeURIComponent(fileId) +
    '?supportsAllDrives=true&fields=id,name,mimeType,driveId,owners(emailAddress),capabilities(canEdit)';

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    const who = safeExecutingUserEmail_();
    let hint = '';
    if (code === 404) {
      hint =
        '\nIf you recently authorized with only "See/edit files created by the app" (drive.file), Drive may return 404 for this spreadsheet.\n' +
        'This add-on needs full Google Drive access for the user (https://www.googleapis.com/auth/drive). Remove/re-install the add-on and approve Drive access, or revoke prior authorization and try again.\n';
    }
    throw new Error(
      'Cannot access spreadsheet file in Drive (HTTP ' +
        code +
        '). This usually means the add-on is executing as a user who lacks edit access to the file, ' +
        'or the file is in a Shared Drive / restricted context.\n' +
        'Executing user: ' +
        who +
        '\n' +
        'File id: ' +
        fileId +
        hint +
        '\n' +
        body.substring(0, 800)
    );
  }
  const parsed = JSON.parse(body);
  if (parsed && parsed.capabilities && parsed.capabilities.canEdit === false) {
    throw new Error(
      'You do not have edit access to this spreadsheet (Drive capabilities.canEdit=false). ' +
        'Edit access is required to attach a bound Apps Script project.\n' +
        'Executing user: ' +
        safeExecutingUserEmail_()
    );
  }
}

function safeExecutingUserEmail_() {
  try {
    return String(Session.getEffectiveUser().getEmail() || '');
  } catch (e) {
    return '(unknown)';
  }
}

function pushShimContent_(scriptId) {
  const url = 'https://script.googleapis.com/v1/projects/' + encodeURIComponent(scriptId) + '/content';
  const payload = {
    files: [
      {
        name: 'appsscript',
        type: 'JSON',
        source: SHIM_TEMPLATE_APPSSCRIPT_JSON
      },
      {
        name: 'Shim',
        type: 'SERVER_JS',
        source: SHIM_TEMPLATE_CODE_GS
      }
    ]
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Apps Script API update content HTTP ' + code + ': ' + body.substring(0, 800));
  }
}

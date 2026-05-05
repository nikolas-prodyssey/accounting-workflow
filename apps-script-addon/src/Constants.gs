/**
 * Add-on constants. Edit these in one place.
 */

/** Master template spreadsheet (Drive file id). Each provisioning copies this file. */
const MASTER_SPREADSHEET_ID = '17UeCZbA6e_M9XdBihPq0OzRXOE1oVOdvYsm1Ii-VxCw';

/** Library script id used by the shim and by this add-on. */
const ACCOUNTING_LIB_SCRIPT_ID = '1YEFSuMVd2zTzPr2DoPcbXGNtzMts0mR0YWMfWGtEZkhVQUVXXFo2GwHh';

/** Library version pinned by the shim (same value as in shim's appsscript.json). Bump when promoting a new library version. */
const ACCOUNTING_LIB_VERSION = '2';

/** Domain restriction (best-effort soft check; real enforcement is via internal add-on deployment). */
const ALLOWED_DOMAINS = ['prodyssey.solutions'];

/** Title of the bound script project that gets attached to each client spreadsheet. */
const BOUND_PROJECT_TITLE = 'Accounting Manual — Bound Shim';

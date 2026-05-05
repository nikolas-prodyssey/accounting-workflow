# Accounting Manual — Multi-Client Architecture

Three Apps Script projects + one master spreadsheet:

| Project | Path | Role |
|---|---|---|
| **Library** | `apps-script-lib/` | All accounting workflow logic. Standalone Apps Script project published as a versioned **library**. Holds the shared **OpenAI API key** + **prompt id** in Script Properties. |
| **Shim template** | `apps-script-shim/` | Tiny bound-script template (delegations + lazy trigger install). Used as the source the **Add-on** pushes into each client spreadsheet via the Apps Script API. **Not pushed by clasp** — it's content. |
| **Internal Add-on** | `apps-script-addon/` | Editor Add-on with two flows: **Create new client workbook** (copy master + attach shim + bootstrap) and **Attach/upgrade plugin on this workbook**. |
| **Master spreadsheet bound script** | `apps-script/` | Intentionally **empty**. The master must not contain any logic so copies start clean. |

Master spreadsheet ID: `17UeCZbA6e_M9XdBihPq0OzRXOE1oVOdvYsm1Ii-VxCw`
Library script ID: `1YEFSuMVd2zTzPr2DoPcbXGNtzMts0mR0YWMfWGtEZkhVQUVXXFo2GwHh` (pinned at version `1`)

---

## Operator workflow

1. Open the **Accounting Manual Provisioner** add-on (Sheets → Extensions).
2. **Create new client workbook**: enter client name + destination Drive folder ID → add-on copies master, attaches shim, runs bootstrap (re-init).
3. The new workbook now has a `📒 Accounting` menu and uses the central library.
4. To **upgrade** all clients to a new library version, publish a new library version in `apps-script-lib/`, bump `ACCOUNTING_LIB_VERSION` in `apps-script-addon/src/Constants.gs` and the shim template's manifest, then re-push the add-on. Use **Attach/upgrade plugin on this workbook** to roll forward each existing client (or rerun the add-on action across them).

---

## Configuration (one-time)

### 1. Library Script Properties (shared across all clients)

In the library project (https://script.google.com/d/1YEFSuMVd2zTzPr2DoPcbXGNtzMts0mR0YWMfWGtEZkhVQUVXXFo2GwHh):

| Key | Value | Notes |
|---|---|---|
| `OPENAI_API_KEY` | `sk-...` | **Required**. One key for the org. |
| `OPENAI_ACCOUNTING_MANUAL_PROMPT_ID` | `pmpt_…` | Optional override. Defaults to the constant in `apps-script-lib/src/Code.gs`. |

### 2. GCP project (placeholder — operator to set)

Editor Add-ons published to a Workspace domain require a Standard GCP project linked to the add-on script project.

- Create / pick a GCP project under your Workspace domain.
- Enable: **Apps Script API** and **Google Drive API**.
- OAuth consent screen: set **User type → Internal** so it’s domain-restricted.
- In the add-on Apps Script project (https://script.google.com/d/1BgvX2nVv1TZB0en7ymJX0Mu0lTcU23M0bdjxn43oXR8IKhYkz_i3Kyp4/edit) → Project Settings → **Google Cloud Platform (GCP) Project** → "Change project" → paste the GCP **project number**.

### 3. Domain allowlist

`apps-script-addon/src/Constants.gs` → `ALLOWED_DOMAINS` defaults to `prodyssey.solutions`. Update if needed.

---

## Architecture rationale

The library + bound shim model is required (vs. add-on only) because the workflow relies on:
- Native Sheets `📒 Accounting` menu (only buildable from a container-bound `onOpen`).
- `PropertiesService.getDocumentProperties()` for per-workbook state — only available in container-bound scripts.
- Two installable triggers (`onSheetChange_`, `onSignedOffEdit_`) that must reside in the bound project.

A standalone add-on has none of these per-workbook abilities for sheets it didn't create.

The shim is intentionally minimal so updates ship via library version bumps, not by editing each workbook's bound script.

---

## Local development

```bash
# Library
cd apps-script-lib && clasp push --force

# Add-on
cd apps-script-addon && clasp push --force

# Master bound script (kept intentionally empty)
cd apps-script && clasp push --force
```

The shim template is **not** clasp-managed — its source lives in `apps-script-shim/src/` and is mirrored in `apps-script-addon/src/ShimTemplate.gs` for the add-on to push.

---

## Git / GitHub

Source control: [github.com/nikolas-prodyssey/accounting-workflow](https://github.com/nikolas-prodyssey/accounting-workflow) (`main`). One repository at the project root covers all folders above. **Do not commit** `.clasp.json` (gitignored). Day-to-day: `git add`, `git commit`, `git push origin main`. For remote setup (HTTPS vs SSH), merge notes, and layout rules, see **`.cursor/rules/accounting-apps-script.mdc`** (section **Git and GitHub**).

---

## Future hardening ideas

- Move shim template content out of `apps-script-addon/src/ShimTemplate.gs` and read at build time so it can't drift from `apps-script-shim/src/Shim.gs`.
- Add a "rolling upgrade" action in the add-on that walks a folder and upgrades all bound scripts.
- Server-side audit log of provisioning events (timestamp, operator email, spreadsheet id, library version).

# Outlook 365 Integration for CiviCRM — civico fork

This extension enables a user to query CiviCRM as an address book when using
either Microsoft Outlook 365 Web and Desktop clients. It generates a Microsoft
Outlook 365 add-in which connects Outlook 365 to the CiviCRM site; the add-in
can be deployed organisation-wide or per user.

Features include:
* Search CiviCRM for contacts or groups.
* Search for contacts within a CiviCRM group.
* Ability to send email to all contacts in a mailing group.
* Select contacts across different mailing groups.
* Ability to CC or BCC email to one group while sending email to another group.
* Create CiviCRM contacts from the email recipients (To, CC, Sender).
* Create multiple CiviCRM contacts with a single click.
* Save multiple contacts to an existing group or create a new group in CiviCRM.
* Store the current email in CiviCRM as an Activity and associate with the email recipients.

This extension is licensed under [AGPL-3.0](LICENSE.txt).

## This is a fork

This fork is <https://github.com/civico-de/outlook365>. Upstream is
<https://lab.civicrm.org/extensions/outlook365>, kept as the remote `upstream`,
so `git log upstream/master..main` is exactly the delta.

The fork exists because the add-in authenticates with a keyring token, and
keyring refuses any token that appears in a URL or a request parameter.
Upstream sends the credential as `api_key=` in every query string.

The delta, in full:

| Change | Where |
| --- | --- |
| One shared transport that sets `X-Civi-Auth: Bearer <token>`; every call routed through it | `assets/civicrm.js`, `assets/taskpane.js`, `assets/readtaskpane.js`, `outlook365/settings/saveContactInGroup.js` |
| The credential removed from every URL and body | as above |
| The settings dialog is told *whether* a token is stored, never the token — upstream passed it in the dialog's own query string | `assets/taskpane.js`, `assets/readtaskpane.js`, `outlook365/settings/dialog.js` |
| The site key dropped entirely — field, storage and request | `outlook365/settings/dialog.*`, `assets/UIStrings.js` |
| The settings field renamed to an API token, and stored as `civicrm_token` | `outlook365/settings/dialog.*`, `assets/UIStrings.js` |
| `com.cividesk.apikey` dropped from `<requires>`, `keyring` added; fork metadata and `<ver>6.17` | `info.xml` |
| A dev stack and an e2e suite proving the above | `.docker/`, `tests/e2e/` |
| A browser suite driving the real task panes with a stubbed office.js | `tests/browser/`, `playwright.config.mjs`, `package.json` |
| Failed CiviCRM calls shown as an error, APIv3 `is_error` answers included; the save-to-group dialog gets its group list from the pane | `assets/*.js`, `outlook365/settings/saveContactInGroup.*`, `templates/CRM/Outlook365/Page/Taskpane.tpl` |
| The read pane's "Save Folder in CiviCRM" removed: it listed and tagged folders through the Outlook REST API, which Microsoft has retired | `assets/readtaskpane.js`, `templates/CRM/Outlook365/Page/ReadTaskpane.tpl` |
| The task panes made accessible: tab labels and icon-only controls exposed, the contact and group lists operable from the keyboard, the document language declared, and an axe-core check over both panes | `templates/CRM/Outlook365/Page/*.tpl`, `assets/taskpane.js`, `assets/readtaskpane.js`, `assets/taskpane.css`, `assets/UIStrings.js`, `tests/browser/a11y.spec.mjs` |
| The group list paged on scroll instead of repeating its first page | `assets/taskpane.js` |
| The inherited civix-upgrade workflow removed — it fires on every push, calls a third-party reusable workflow and names an extension key this repository does not use | `.github/` |

Everything else — the data processors, the manifest and the Office glue — is
upstream's, unchanged.

### Rebasing on upstream

```bash
git fetch upstream
git rebase upstream/master
bash tests/e2e/run.sh             # against a running .docker stack
```

Conflicts land almost entirely in the four JavaScript files above. The rule
when resolving them: a new upstream call site is wrong until it goes through
`CiviCRM.api3` / `CiviCRM.api4` in `assets/civicrm.js`, and no `api_key` or
`key` may reappear in a URL or a body. `tests/e2e/10-header-auth.sh` fails if
one does.

## Configuring it

Server side, once:

* Enable `authx`, `dataprocessor`, `keyring` and this extension.
* Grant **Access Outlook 365 pages** to anonymous visitors — Outlook loads the
  task pane with no CiviCRM session. On Standalone that is the `everyone` role.
* Leave `authx_guards` containing `perm`, and `authx_xheader_cred` containing
  `api_key`. Both are the stock values.

Per user:

* Issue a keyring token for their contact. The contact needs
  `authenticate with api key` on top of the token's own scope, because the
  add-in sends no site key.
* In the add-in's Settings, enter the CiviCRM site URL and the token. The site
  URL must be the origin that serves the task pane; a different origin would
  need a CORS preflight that CiviCRM does not answer.

Deriving a token's scope from real traffic is keyring's job — see its
operations guide on report mode in <https://github.com/civico-de/keyring>.

## Development

```bash
.docker/fetch-ext.sh                    # once: clone dataprocessor
cd .docker && docker compose up -d      # CiviCRM on http://localhost:8200 (admin/admin)

bash tests/e2e/run.sh                   # from the host, against that stack

npm ci && npx playwright install chromium   # once
npm run test:e2e                        # browser suite, against that stack
```

The suite needs `curl`, `python3` and `node` on the host; `tests/e2e/40-uistrings.sh`
resolves `assets/UIStrings.js` with node and runs without the stack.

`tests/browser/` loads `civicrm/outlook365/taskpane.html` and `readtaskpane.html`
from the stack in Chromium. `office.js` is replaced by `tests/browser/office-stub.js`,
which holds the fixture mail and the settings and throws on any Office member it
does not model; dialogs open as iframes that may call only what Microsoft's dialog
API allows there. Each run mints its own keyring token for the seeded
non-superuser through `tests/e2e/lib.sh` and checks every request (token only in
`X-Civi-Auth`, no site key), the pane's output and the rows written. Fabric comes
from the `office-ui-fabric-js` package at the version the pages load; any other
request to an external host fails the test. `CIVICRM_BASE_URL` (default
`http://localhost:8200`) and `CK_COMPOSE_FILE` point it at another stack.

The compose project is named `outlook365` and mounts this checkout, the sibling
`keyring` checkout and the cached `dataprocessor` clone.
`docker compose down -v && docker compose up -d` re-runs provisioning from
scratch.

### CI

Every push and pull request against `main` runs, against a fresh civikitchen
stack:

```bash
cklint --all && ckconform && ckcivix --check && ckfmt --check && ckeslint
phpstan analyse && ckcompat && ckdeps && cktaint && cksmarty
bash tests/e2e/run.sh
npm run test:e2e
```

plus `ckinit --check` for template drift and a disable/enable/uninstall/install
pass (`lifecycle: true`). The e2e job mounts `keyring` from a private checkout
and boots the CI stack with port 8200 published; `dataprocessor` resolves
through the pinned archive in `civikitchen.yaml` (`policy.extension_sources`),
same as the `ci` job's own stack. The browser suite runs in civikitchen's
`playwright` job on its own stack (port 8080, `keyring` as `sibling_repo`).

## Documentation

| Document | What it covers |
| --- | --- |
| [docs/GettingStarted.md](docs/GettingStarted.md) | Installing the extension and adding the add-in to Outlook 365 |
| [docs/HowToAddPhoneNumber.md](docs/HowToAddPhoneNumber.md) | Adding the phone number, or any other CiviCRM data, to Outlook 365 |
| [docs/measurements.md](docs/measurements.md) | What was measured on CiviCRM 6.17.3: installability, the header flow, why the site key could go, required authx settings |

## Requirements

* CiviCRM 6.17 or newer (measured on 6.17.3)
* PHP 8.4 (measured on 8.4.25)
* Data Processor 1.1.0 or newer
* [keyring](https://github.com/civico-de/keyring)
* Microsoft Office 365

## About the Authors

CiviCRM Outlook 365 integration was developed by:
* Jaap Jansma from [CiviCoop](https://civicoop.org)
* Sarvesh Mehta as part of the [Google Summer of Code 2020](https://summerofcode.withgoogle.com/archive/2020/projects/6234630803423232/) with mentoring from Justin Freeman from [Agileware](https://agileware.com.au)
* With new features and refactoring by the team at [Agileware](https://agileware.com.au)

You can say "Thanks!" and support your CiviCRM developers by contacting these
companies for paid CiviCRM support and CiviCRM development services:
[CiviCoop](https://civicoop.org) and [Agileware](https://agileware.com.au).

Re-development of the extension was funded and performed as part of
[Google Summer of Code 2020](https://summerofcode.withgoogle.com/archive/2020/projects/6234630803423232/).

This fork is maintained by civico GmbH.

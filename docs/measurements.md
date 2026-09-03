# Measurements

Every statement here was produced by running `tests/e2e/` against the
`.docker` stack, not by reading source. Raw evidence stays in `tests/e2e/out/`
(gitignored; regenerate with `bash tests/e2e/run.sh`).

## Environment

| Item | Value |
| --- | --- |
| CiviCRM | 6.17.3 |
| CMS | Standalone (`standaloneusers`) |
| PHP | 8.4.25 |
| Smarty | 5.7.0 |
| Image | `ghcr.io/jfilter/civikitchen:v1` |
| Compose project | `outlook365`, app on `http://localhost:8200` |
| Extensions enabled | `authx`, `dataprocessor`, `keyring`, `outlook365` |

The measurement subject is a non-superuser contact with a Standalone user and a
role, seeded by `tests/e2e/seed.php`. Standalone's admin holds
`all CiviCRM permissions and ACLs`, which short-circuits
`CRM_Core_Permission::check` before a token scope is ever consulted.

## Installability on 6.17.3

The add-in installs and enables with `com.cividesk.apikey` dropped from
`info.xml`. Nothing else had to change for the boot to succeed — the Smarty 5
failure previously observed belonged to that dependency, not to this
extension.

```
Enabling extensions: authx,dataprocessor,keyring,outlook365
Enabling extension "authx"
Enabling extension "dataprocessor"
Enabling extension "keyring"
Enabling extension "outlook365"
```

`Extension.get` afterwards reports `outlook365` `2.2.7.1` `installed`, and both
shipped data processors (`outlook_365_contacts`, `outlook_365_groups`) are
present and active, so `CRM_Outlook365_Upgrader::postInstall` ran despite
`info.xml` naming `CiviMix\Schema\Outlook365\AutomaticUpgrader` as the
upgrader.

Both task panes render under Smarty 5 with no unrendered tags left in the
output (`tests/e2e/20-install.sh` greps the served HTML for `{ts`, `{literal`,
`{crmScope` and `{$`). The `smarty-v2@1.0.3` mixin the extension already
declared is what carries the v2 syntax in its templates.

`<compatibility><ver>` is therefore **6.17** — the version this was measured
on. Upstream's `5.38` is not a claim this fork can defend: it predates both
Smarty 5 and keyring.

## The requests the add-in now sends

`assets/civicrm.js` is the only place that talks to CiviCRM. It POSTs

- APIv3: `entity`, `action`, `json` as a form body to `civicrm/ajax/rest`
- APIv4: `params` as a form body to `civicrm/ajax/api4/<Entity>/<action>`

with `X-Civi-Auth: Bearer <token>` and nothing else. `tests/e2e/lib.sh` builds
the same shapes with curl.

| Call | Result |
| --- | --- |
| APIv3 `Contact.get`, header only | 200, `is_error:0` |
| APIv3 `Outlook365Contact.get`, header only | 200, `is_error:0` |
| APIv3 `Outlook365Group.get`, header only | 200, `is_error:0` |
| APIv4 `Email.get`, header only | 200, `count` returned |
| APIv4 `Activity.create` (POST), header only | 200, `count:1` |

The write is attributed to the token's contact:
`civicrm_activity_contact.record_type_id = 2` holds the subject's contact id.
APIv4 `Activity.create` needs an explicit `source_contact_id` — it derives
nothing from the session — which is why `assets/readtaskpane.js` keeps passing
`user_contact_id`.

Reads go over POST because one transport serves both; the legacy endpoint
accepts POST for reads and *requires* it for writes
(`SECURITY: All requests that modify the database must be http POST, not GET`).

## The URL is refused

keyring rejects a `kr_` credential that appears anywhere but a header, and the
message names the header:

```
FATAL: A keyring token must be sent as a request header
("X-Civi-Auth: Bearer <token>" or "Authorization: Bearer <token>"),
never as a URL or form parameter.
```

Measured refusals, each one a shape upstream produced:

| Shape | Where it came from upstream | Result |
| --- | --- | --- |
| `?…&api_key=<token>&key=<sitekey>` | every APIv3 call | refused |
| `/civicrm/ajax/api4/Email/get?api_key=<token>&key=<sitekey>` | `readtaskpane.js`, `api4()` | refused |
| correct header **and** `api_key=` in the URL | — | refused |
| `api_key=<token>` in the POST body | — | refused |

The fourth case matters for the transport chosen here: moving the credential
from the query string into the POST body would not have been enough.

## The site key

**Dropped, from the request and from the settings dialog.** The measurement
that allows it:

- `authx_guards` is `["site_key","perm"]` and the guards are OR'd.
- keyring adds `authenticate with api key` to every token scope, so the perm
  guard is satisfied whenever the token's contact holds that permission.
- Every passing call above sent no site key at all. Adding
  `X-Civi-Key: <sitekey>` on top changes nothing.

The counter-measurement fixes the deployment requirement. With
`authenticate with api key` removed from the subject's role, the same
header-only request fails:

```
HTTP 401 Login not permitted. Must satisfy guard (site_key, perm).
```

and sending `X-Civi-Key` instead makes it pass again. So: the token's contact
**must** hold `authenticate with api key`, or the site key would have to come
back. `tests/e2e/30-deployment.sh` asserts both directions and restores the
role.

## Authx settings the deployment needs

| Setting | Required value | Why |
| --- | --- | --- |
| `authx_xheader_cred` | contains `api_key` | carries `X-Civi-Auth: Bearer` |
| `authx_header_cred` | contains `api_key` | carries `Authorization: Bearer`, the fallback keyring also accepts |
| `authx_guards` | contains `perm` | the site key is no longer sent |

All three ship this way on a stock Standalone install; nothing had to be
changed on the stack. `authx_legacyrest_cred` and `authx_param_cred` are
irrelevant to this fork — keyring refuses tokens arriving on those flows by
design.

Two CiviCRM permissions are also required:

- the token's contact needs `authenticate with api key` (above), plus whatever
  the token's scope allows it to read and write;
- `access outlook 365 pages` must be granted to anonymous visitors. Outlook
  loads the task pane with no CiviCRM session, so without it the pane is a 403.
  On Standalone that is the `everyone` role.

## Cross-origin

`X-Civi-Auth` is not a CORS-safelisted request header, so a task pane
configured with a *different* origin than the one serving it would preflight.
Measured: an `OPTIONS` request to `civicrm/ajax/rest` with
`Access-Control-Request-Headers: x-civi-auth` comes back with **no**
`Access-Control-Allow-Origin` header. CiviCRM does not answer the preflight,
so such a configuration cannot work.

This is not a regression the fork introduced — it is the consequence of moving
the credential into a header. In practice the constraint is already satisfied:
the task pane is served by the CiviCRM site it talks to, and the settings
dialog is prefilled with that site's own base URL. The site URL field must be
that same origin.

## What was not measured

- Behaviour inside a real Outlook client. Everything here is curl against the
  same request shapes; no Office host was driven.
- Any CiviCRM version other than 6.17.3, and any CMS other than Standalone.
- The Office roaming-settings migration. The stored key changed from
  `civicrm_apikey` to `civicrm_token`, so every user re-enters the credential
  once — which they must anyway, an api_key not being a keyring token.

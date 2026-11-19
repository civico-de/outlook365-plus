---
title: Reference
order: 30
---

# Reference

## Permissions

| Permission | Who needs it |
|---|---|
| Access Outlook 365 pages (`access outlook 365 pages`) | Anonymous visitors. It opens the add-in's panes, which Outlook loads without a CiviCRM session. The panes themselves hold no data; every data call carries the user's token. |
| `access CiviCRM` | Staff who download the manifest from the Administer menu. |
| `authenticate with api key` | Every contact who uses the add-in, through their role, on top of the token's scope. |

## Token scope

The keyring token of each user needs these permissions in its scope:

- `access CiviCRM`
- `access AJAX API`
- `view all contacts`
- `edit all contacts`
- `add contacts`
- `view all activities`
- `edit groups`, to add people to a group or create one

A call outside the scope fails, and the pane shows an error.

## Contact and group lists

The add-in reads both lists through APIv4 `Contact.get` with the user's own permissions, so contact ACLs apply. Both list contacts that are not deleted and have a primary email address, each with display name, primary email address and contact ID. The contact list also leaves out deceased contacts and those with "do not email" or "do not trade". A group lists its members with the status Added.

Versions that read the lists through two data processors (`Outlook365Contact`, `Outlook365Group`) leave them behind when Data Processor is installed; the upgrade deletes both.

## URLs

| Path | Purpose |
|---|---|
| `civicrm/outlook365/manifest.xml` | Downloads the manifest. |
| `civicrm/outlook365/taskpane.html` | The pane for writing a mail. |
| `civicrm/outlook365/readtaskpane.html` | The pane for reading a mail. |

## What the add-in stores in Outlook

The site URL and the token are stored in the add-in's roaming settings in the user's mailbox (`civicrm_url`, `civicrm_token`), so they follow the user to every Outlook client. The settings dialog only learns whether a token is stored, never the token itself.

## What the add-in sends

The add-in sends every call as a POST to CiviCRM's API endpoints (`civicrm/ajax/rest` for APIv3, `civicrm/ajax/api4/<Entity>/<action>` for APIv4), with the token in the header `X-Civi-Auth: Bearer <token>`. The token never appears in a URL or a form field; keyring refuses a token sent that way.

---
title: Referenz
order: 30
---

# Referenz

## Berechtigungen

| Berechtigung | Wer sie braucht |
|---|---|
| Access Outlook 365 pages (`access outlook 365 pages`) | Anonyme Besucher. Sie öffnet die Bereiche des Add-ins, die Outlook ohne CiviCRM-Sitzung lädt. Die Bereiche selbst enthalten keine Daten; jeder Datenaufruf trägt das Token des Benutzers. |
| `access CiviCRM` | Mitarbeitende, die das Manifest aus dem Menü Administer herunterladen. |
| `authenticate with api key` | Jeder Kontakt, der das Add-in nutzt, über seine Rolle, zusätzlich zum Scope des Tokens. |

## Token-Scope

Das keyring-Token jedes Benutzers braucht diese Berechtigungen in seinem Scope:

- `access CiviCRM`
- `access AJAX API`
- `view all contacts`
- `edit all contacts`
- `add contacts`
- `view all activities`
- `edit groups`, um Personen einer Gruppe hinzuzufügen oder eine Gruppe anzulegen

Ein Aufruf außerhalb des Scopes schlägt fehl, und der Bereich zeigt einen Fehler.

## Kontakt- und Gruppenlisten

Das Add-in liest beide Listen über APIv4 `Contact.get` mit den Berechtigungen des Benutzers, Kontakt-ACLs gelten also. Beide listen Kontakte, die nicht gelöscht sind und eine primäre E-Mail-Adresse besitzen, jeweils mit Anzeigename, primärer E-Mail-Adresse und Kontakt-ID. Die Kontaktliste lässt außerdem verstorbene Kontakte aus und solche mit „keine E-Mails“ oder „kein Handel“. Eine Gruppe listet ihre Mitglieder mit dem Status „Added“.

Versionen, die die Listen über zwei Data Processors (`Outlook365Contact`, `Outlook365Group`) lasen, lassen diese bei installiertem Data Processor zurück; das Upgrade löscht beide.

## URLs

| Pfad | Zweck |
|---|---|
| `civicrm/outlook365/manifest.xml` | Lädt das Manifest herunter. |
| `civicrm/outlook365/taskpane.html` | Der Bereich zum Schreiben einer Mail. |
| `civicrm/outlook365/readtaskpane.html` | Der Bereich zum Lesen einer Mail. |

## Was das Add-in in Outlook speichert

Website-URL und Token liegen in den Roaming-Einstellungen des Add-ins im Postfach des Benutzers (`civicrm_url`, `civicrm_token`) und folgen ihm so in jeden Outlook-Client. Der Einstellungsdialog erfährt nur, ob ein Token gespeichert ist, nie das Token selbst.

## Was das Add-in sendet

Das Add-in schickt jeden Aufruf als POST an die API-Endpunkte von CiviCRM (`civicrm/ajax/rest` für APIv3, `civicrm/ajax/api4/<Entity>/<action>` für APIv4), mit dem Token im Header `X-Civi-Auth: Bearer <token>`. Das Token steht nie in einer URL oder einem Formularfeld; keyring weist ein so übertragenes Token ab.

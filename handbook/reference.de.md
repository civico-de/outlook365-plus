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

## Data Processors

Beim Aktivieren installiert die Extension zwei Data Processors, die das Add-in über APIv3 liest:

| Data Processor | API-Entität | Inhalt |
|---|---|---|
| Outlook 365 Contacts | `Outlook365Contact` | Kontakte mit Anzeigename und primärer E-Mail-Adresse. Lässt gelöschte und verstorbene Kontakte aus, solche mit „keine E-Mails“ oder „kein Handel“ und solche, die der Benutzer nicht sehen darf. |
| Outlook 365 Groups | `Outlook365Group` | Gruppenmitglieder mit Anzeigename und primärer E-Mail-Adresse, für Mitglieder mit dem Status „Added“ in der Gruppe, die nicht gelöscht sind. |

Die Kontaktliste beim Schreiben zeigt jedes Feld des Kontakt-Data-Processors. Um mehr zu zeigen, etwa die Telefonnummer, bearbeiten Sie **Outlook 365 Contacts** unter **Administer → Data Processor → Manage Data Processors**: Fügen Sie eine Datenquelle für die gewünschten Daten hinzu, verknüpft mit dem Kontakt, und fügen Sie ihr Feld hinzu. **Revert** in der Liste der Data Processors stellt die mitgelieferte Definition wieder her.

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

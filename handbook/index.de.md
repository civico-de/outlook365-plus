---
title: Outlook 365 Integration
---

# Outlook 365 Integration

Diese Extension macht CiviCRM zum Adressbuch in Microsoft Outlook 365, im Web wie im Desktop-Programm. Sie erzeugt ein Outlook-Add-in, das Outlook mit Ihrer CiviCRM-Website verbindet. Wer Mails in Outlook schreibt, kann CiviCRM-Kontakte und -Gruppen adressieren, die Personen einer Mail in CiviCRM aufnehmen und eine Mail als Aktivität ablegen, ohne Outlook zu verlassen.

Das Add-in lässt sich für die ganze Organisation oder von jedem Benutzer selbst einrichten.

## Was es kann

Beim Schreiben einer Mail:

- CiviCRM-Kontakte suchen und als An, Cc oder Bcc eintragen.
- CiviCRM-Gruppen suchen, eine Gruppe öffnen, einige oder alle Mitglieder auswählen und als Empfänger eintragen, sodass eine Gruppe in An und eine andere in Bcc stehen kann.

Beim Lesen einer Mail:

- Sehen, wer von Absender und Empfängern schon Kontakt in CiviCRM ist.
- Für die anderen Kontakte anlegen, einzeln oder alle ausgewählten auf einmal.
- Die ausgewählten Personen einer bestehenden oder einer neuen Gruppe hinzufügen.
- Die Mail als E-Mail-Aktivität speichern, mit den ausgewählten Personen als Empfängern.

Einzelheiten stehen unter [Das Add-in benutzen](usage.md).

## Wie es sich anmeldet

Jeder Benutzer trägt in den Einstellungen des Add-ins die URL der CiviCRM-Website und ein persönliches API-Token ein. Das Token stellt die Extension [keyring](https://github.com/civico-de/keyring) aus; es wird nur in einem Request-Header übertragen, nie in einer URL oder einem Formularfeld. Alles, was das Add-in tut, wird dem Kontakt dieses Benutzers zugeschrieben.

Dies ist ein gepflegter Fork der [ursprünglichen Extension](https://lab.civicrm.org/extensions/outlook365). Der Unterschied ist genau diese Anmeldung: Das Original schickt API-Key und Site-Key in jeder URL mit.

## Voraussetzungen

- CiviCRM in der Version aus der `info.xml` der Extension oder neuer.
- PHP 8.2 oder neuer.
- Die Extensions [Data Processor](https://civicrm.org/extensions/data-processor) ab 1.1.0 und [keyring](https://github.com/civico-de/keyring) sowie die Core-Extension `authx`.
- Microsoft Outlook 365.

## Installation

Laden Sie das Release-Archiv von der [Release-Seite](https://github.com/civico-de/outlook365/releases), entpacken Sie es in das Extension-Verzeichnis von CiviCRM und aktivieren Sie es nach seinen Voraussetzungen:

```bash
cv ext:enable authx dataprocessor keyring outlook365
```

Beim Aktivieren werden die zwei Data Processors installiert, über die das Add-in Kontakte und Gruppen liest. Weiter geht es mit der [Einrichtung](setup.md).

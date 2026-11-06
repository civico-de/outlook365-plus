---
title: Einrichtung
order: 10
---

# Einrichtung

## CiviCRM, einmalig

1. Aktivieren Sie `authx`, `dataprocessor`, `keyring` und diese Extension.
2. Geben Sie anonymen Besuchern die Berechtigung **Access Outlook 365 pages**. Outlook lädt die Bereiche des Add-ins ohne CiviCRM-Sitzung.
    - Standalone: die Rolle **Everyone, including anonymous users**.
    - WordPress: **Anonymous User**.
    - Drupal: sowohl **Anonymous User** als auch **Authenticated User**.
3. Lassen Sie die authx-Einstellungen auf ihren Standardwerten: `authx_guards` enthält `perm`, und `authx_xheader_cred` enthält `api_key`.

## Ein Token für jeden Benutzer

Jede Person, die das Add-in nutzt, braucht ein eigenes keyring-Token, ausgestellt im Reiter **API tokens** ihres Kontakts. Das Token wird einmal angezeigt, beim Ausstellen; ein verlorenes Token wird ersetzt, nie wiederhergestellt. Übergeben Sie es auf einem Weg, der es nicht aufbewahrt.

Der Scope des Tokens braucht die Berechtigungen unter [Token-Scope](reference.md#token-scope). Zusätzlich braucht der Kontakt über seine Rolle die Berechtigung **authenticate with api key**, weil das Add-in keinen Site-Key sendet.

Ein Token gehört zu genau einer Person: Alles, was das Add-in tut, wird als Handlung dieses Kontakts erfasst, und das Token lässt sich einzeln widerrufen.

## Das Manifest

Öffnen Sie in CiviCRM **Administer → Download Outlook 365 Manifest.xml** und speichern Sie die Datei `manifest.xml`. Mit ihr wird das Add-in in Outlook installiert. Das Add-in trägt den Namen des Kontakts Ihrer eigenen Organisation in CiviCRM, und das Manifest verweist Outlook auf die URL Ihrer Website; laden Sie es also von der Website herunter, mit der das Add-in sprechen soll.

## Das Add-in einrichten

Es gibt zwei Wege. Microsoft ändert seine Verwaltungsseiten häufig, die Bezeichnungen können daher leicht abweichen.

### Für die ganze Organisation

1. Öffnen Sie das Microsoft 365 Admin Center.
2. Gehen Sie zu **Einstellungen → Integrierte Apps** und öffnen Sie die Seite der Add-ins.
3. Wählen Sie **Add-In bereitstellen**, dann **Benutzerdefinierte Apps hochladen**.
4. Laden Sie `manifest.xml` hoch, legen Sie fest, wer das Add-in bekommt, und stellen Sie es bereit.

### Für einen Benutzer

1. Öffnen Sie in Outlook eine Mail oder beginnen Sie eine neue.
2. Öffnen Sie das Menü **…** und wählen Sie **Add-Ins abrufen**.
3. Wählen Sie unter **Meine Add-Ins** den Punkt **Benutzerdefiniertes Add-In hinzufügen → Aus Datei hinzufügen** und wählen Sie `manifest.xml`.
4. Bestätigen Sie die Installation.

## Erste Verwendung

Jeder Benutzer richtet das Add-in einmal ein:

1. Öffnen Sie in Outlook eine Mail oder beginnen Sie eine neue.
2. Öffnen Sie das Menü **…** und wählen Sie das Add-in, erkennbar am CiviCRM-Logo und am Namen Ihrer Organisation.
3. Wählen Sie im Bereich **Settings**.
4. Tragen Sie **CiviCRM Site URL** und **CiviCRM API token (keyring)** ein und wählen Sie **Save**.

Die Website-URL ist die Basisadresse der Website, zum Beispiel `https://crm.example.org`, keine Seite darunter. Es muss dieselbe Website sein, die die Bereiche des Add-ins ausliefert: Das Token steht in einem Request-Header, und ein Aufruf an eine andere Website bräuchte eine Cross-Origin-Vorabanfrage, die CiviCRM nicht beantwortet.

Um die URL später zu ändern, öffnen Sie wieder **Settings**. Bleibt das Token-Feld leer, behält das Add-in das gespeicherte Token, solange die Website-URL gleich bleibt.

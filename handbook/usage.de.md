---
title: Das Add-in benutzen
order: 20
---

# Das Add-in benutzen

Sie öffnen das Add-in über das Menü **…** einer Mail. Je nachdem, ob Sie die Mail schreiben oder lesen, öffnet sich ein anderer Bereich. Solange das Add-in nicht eingerichtet ist, bittet der Bereich darum, **Settings** zu öffnen, siehe [Erste Verwendung](setup.md#erste-verwendung).

Schlägt ein Aufruf an CiviCRM fehl, etwa weil der Scope des Tokens ihn nicht erlaubt, zeigt der Bereich einen Fehler, und die Aktion wird nicht ausgeführt.

## Eine Mail schreiben

Der Bereich hat zwei Reiter, **Contacts** und **Groups**. Beide Listen laden 25 Einträge auf einmal und beim Scrollen weitere.

### Kontakte

Geben Sie einen Namen in das Suchfeld ein und drücken Sie die Eingabetaste. Die Liste zeigt passende Kontakte mit ihrer primären E-Mail-Adresse. **To**, **Cc** und **Bcc** tragen den Kontakt in das jeweilige Feld der Mail ein; eine Adresse, die schon dort steht, wird nicht doppelt eingetragen. Das Kontakt-Symbol öffnet den Kontakt in CiviCRM.

Die Liste lässt Kontakte aus, die gelöscht oder verstorben sind, bei denen „keine E-Mails“ oder „kein Handel“ gesetzt ist und die der angemeldete Benutzer nicht sehen darf.

### Gruppen

Gruppen suchen Sie genauso nach ihrem Titel. Öffnen Sie eine Gruppe, um ihre Mitglieder mit primärer E-Mail-Adresse zu sehen; ein Suchfeld in der Gruppe filtert die Mitglieder nach Namen. Wählen Sie Mitglieder einzeln oder mit **Select All** aus und tragen Sie die ausgewählten mit **To**, **Cc** oder **Bcc** ein. Die Knöpfe einer Gruppe übernehmen nur deren Auswahl; so kann eine Gruppe in An und eine andere in Bcc stehen.

Das Einstellungs-Symbol neben einer Gruppe öffnet die Einstellungen der Gruppe in CiviCRM.

## Eine Mail lesen

Der Bereich listet den Absender und alle Empfänger aus An, Cc und Bcc auf, zunächst alle ausgewählt.

- Eine Adresse, die zu einem CiviCRM-Kontakt gehört, zeigt den Namen des Kontakts und einen Link zu ihm.
- Eine Adresse ohne Kontakt hat einen Speichern-Knopf. Er legt einen Kontakt vom Typ Einzelperson mit dieser E-Mail-Adresse an. Der Anzeigename aus der Mail wird am ersten Leerzeichen (oder, ohne Leerzeichen, am ersten Komma) in Vor- und Nachname geteilt. Ein einzelnes Wort wird zum Vornamen. Ist der Name eine E-Mail-Adresse, bleiben Vor- und Nachname leer, und CiviCRM zeigt die E-Mail-Adresse als Namen des Kontakts. Die E-Mail-Adresse des neuen Kontakts ist seine primäre, mit dem Standard-Adresstyp.

Unter der Liste:

- **Save Contacts** legt Kontakte für alle ausgewählten Adressen an, die noch nicht in CiviCRM sind.
- **Save Contacts to Group** fragt nach einer Gruppe: einer bestehenden aktiven Gruppe oder dem Namen einer neuen. Es legt die fehlenden Kontakte an und fügt alle ausgewählten Personen der Gruppe hinzu.
- **Save Email** legt die Mail in CiviCRM als Aktivität vom Typ E-Mail ab. Ausgewählte Personen, die noch nicht in CiviCRM sind, werden vorher angelegt, und die Liste zeigt sie danach als bekannt. Der Knopf ist nutzbar, sobald die Liste geladen ist; nach einer Änderung der Einstellungen wartet er auf die Liste der neuen Website. Eine Mail wird einmal gespeichert: Nach dem Speichern, oder wenn die Mail schon die Kategorie „Saved in CiviCRM“ trägt, bleibt der Knopf deaktiviert, und der Bereich zeigt „This email is saved in CiviCRM.“ an.

Die E-Mail-Aktivität enthält:

| Feld | Wert |
|---|---|
| Betreff | Der Betreff der Mail |
| Details | Der HTML-Text der Mail |
| Datum | Wann die Mail erstellt wurde |
| Hinzugefügt von | Der Kontakt, dem das Token gehört |
| Mit Kontakt | Die ausgewählten Personen |

Nach dem Speichern bekommt die Mail die Outlook-Kategorie **Saved in CiviCRM**. Das Add-in legt die Kategorie in Ihrem Postfach an, wenn es sie zum ersten Mal braucht.

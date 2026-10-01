# Wareneingang · Stichprobenmessung

Übersicht für alle Mitarbeitenden: Welche Lieferung muss zu Vogt zum Messen, welche hat einen Freipass, und was passiert mit der nächsten Lieferung eines Artikels.

Datengrundlage ist die Excel-Liste **„Lieferungen zu Vogt“**. Das ist die Pivot „Pivot Alle Artikel“ aus dem BI-Würfel `AMS01_BI`, die sich beim Öffnen selbst aktualisiert. Jedes ✓ (Bestellmenge geliefert > 0) ist eine Lieferung pro Artikel, Lieferant und Geschäftsmonat. Die Messergebnisse werden nicht mehr über Zellfarben geführt, sondern im Dashboard erfasst.

## Einrichten

Benötigt wird nur Node.js (ab Version 18), keine weiteren Pakete.

1. Die Excel-Liste an einen Ort legen, den der Server lesen kann, z. B. ein Netzlaufwerk.
2. **Einmalig** die bisherigen Ergebnisse aus den Zellfarben übernehmen:
   ```
   node import-altbestand.js "P:\Qualität\Lieferungen_zu_Vogt.xlsx"
   ```
   - grün → i.O., rot → Abweichung (über 0.002 mm), gelb → nicht zu Vogt
   - die Tabelle „Massabweichung“ unter der Pivot → Details zur Abweichung
   - Kommentare an den ✓ → Bemerkung
3. Server starten:
   ```
   set LISTE=P:\Qualität\Lieferungen_zu_Vogt.xlsx
   node server.js
   ```
   Danach ist das Dashboard im Browser unter `http://<servername>:8080` erreichbar.

| Variable | Bedeutung                     | Standard                          |
|----------|-------------------------------|-----------------------------------|
| `PORT`   | Port des Servers              | `8080`                            |
| `LISTE`  | Pfad zur Excel-Liste (.xlsx)  | `data/Lieferungen_zu_Vogt.xlsx`   |

## Aktualisierung

Die Excel-Liste aktualisiert sich beim Öffnen selbst aus dem BI. Sobald sie danach **gespeichert** wird, liest der Server sie neu ein. Die Seite fragt jede Minute nach, neue ✓ erscheinen also kurz nach dem Speichern.

Die Ergebnisse (bei Vogt abgegeben, i.O., ausserhalb Toleranz, ging nicht zu Vogt) liegen zentral in `data/rueckmeldungen.json`. Diese Datei bitte sichern. Sie gehört nicht ins Git, weil sie Firmendaten enthält.

## Zählung pro Bestellung

Ohne weitere Einstellung zeigt die Pivot ein ✓ pro Monat. Zwei Lieferungen im selben Monat zählen dann als eine. Damit jede Bestellung einzeln zählt, in Excel einmalig:

1. In die Pivot klicken, dann rechts in der Feldliste unter **Purchase Orders** das Feld **„Bestell-Nr. - Position“** suchen.
2. Es in den Bereich **Zeilen** ziehen, und zwar ganz nach unten, unter „Lieferant Bez.“.
3. Speichern.

Das Dashboard erkennt die neue Spalte automatisch. Bereits erfasste Ergebnisse pro Monat gelten dann für alle Bestellungen dieses Monats. Neue Ergebnisse werden pro Bestellung erfasst. Wichtig: Unter der Pivot steht die Tabelle „Massabweichung“. Die Pivot wird nicht breiter, aber länger. Steht die Tabelle direkt darunter, meldet Excel beim Aktualisieren einen Konflikt. Die Tabelle deshalb vorher auf ein eigenes Blatt verschieben oder löschen, die Abweichungen sind nach dem Import im Dashboard.

## Regeln (pro Artikel + Lieferant)

1. Qualifizierung: jede Lieferung geht zu Vogt, bis 3 Messungen in Folge i.O. sind. Wurde ein Artikel noch nie bei Vogt gemessen, muss die nächste Lieferung zu Vogt.
2. 3 Lieferungen mit Freipass.
3. Die 4. Lieferung geht zur Kontrollmessung zu Vogt. Ist sie i.O., folgen nochmals 3 Freipässe.
4. Danach folgt die Requalifizierung. Ist sie i.O., geht es wieder bei Schritt 2 weiter.
5. Jede Abweichung, auch eine nachträglich gemeldete bei einer Freipass-Lieferung, führt zurück zu Schritt 1.
6. Eine zusätzliche Messung i.O. während der Freipass-Phase verbraucht keinen Freipass.
7. Solange ein Messergebnis aussteht, gehen auch folgende Lieferungen desselben Artikels zu Vogt.

Die Logik steht in `regeln.js`, das Einlesen der Excel-Liste in `liste.js` und `xlsx.js`. Tests: `node --test test/*.test.js`

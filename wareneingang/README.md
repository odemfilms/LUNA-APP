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
   - Artikel, bei denen alle Lieferungen gelb sind → „nicht messpflichtig“ (im Dashboard änderbar)
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

Der Server liest die Excel-Datei neu ein, sobald sie sich geändert hat. Die Seite fragt jede Minute nach. Neue ✓ erscheinen also, sobald jemand die Liste aktualisiert und **speichert**. Damit das ohne Zutun passiert, sollte die Liste automatisch aktualisiert und gespeichert werden, z. B. stündlich über die Windows-Aufgabenplanung oder Power Automate.

Ergebnisse (bei Vogt abgegeben, i.O., ausserhalb Toleranz, ging nicht zu Vogt) und die Einstellung „nicht messpflichtig“ liegen zentral in `data/rueckmeldungen.json` und `data/artikel.json`. Diese Dateien bitte sichern. Sie gehören nicht ins Git, weil sie Firmendaten enthalten.

## Genauigkeit pro Bestellung

Die Pivot zeigt ein ✓ pro Monat. Kommen in einem Monat zwei Lieferungen desselben Artikels, zählen sie als eine. Für eine Zählung pro Bestellung kann man in der Pivot das Feld **„Bestell-Nr. - Position“** (Dimension *Purchase Orders*) zusätzlich als Zeilenfeld einblenden. Das Dashboard erkennt die Spalte automatisch und zählt dann jede Bestellposition einzeln. In diesem Fall den Altbestand-Import erneut prüfen.

## Regeln (pro Artikel + Lieferant)

1. Qualifizierung: jede Lieferung geht zu Vogt, bis 3 Messungen in Folge i.O. sind.
2. 3 Lieferungen mit Freipass.
3. Die 4. Lieferung geht zur Kontrollmessung zu Vogt. Ist sie i.O., folgen nochmals 3 Freipässe.
4. Danach folgt die Requalifizierung. Ist sie i.O., geht es wieder bei Schritt 2 weiter.
5. Jede Abweichung, auch eine nachträglich gemeldete bei einer Freipass-Lieferung, führt zurück zu Schritt 1.
6. Eine zusätzliche Messung i.O. während der Freipass-Phase verbraucht keinen Freipass.
7. Solange ein Messergebnis aussteht, gehen auch folgende Lieferungen desselben Artikels zu Vogt.

Die Logik steht in `regeln.js`, das Einlesen der Excel-Liste in `liste.js` und `xlsx.js`. Tests: `node --test test/*.test.js`

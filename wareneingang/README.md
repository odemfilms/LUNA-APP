# Wareneingang · Stichprobenmessung

Übersicht für alle Mitarbeitenden: Welche Lieferung muss zu Vogt zum Messen, welche hat einen Freipass, und was passiert mit der nächsten Lieferung eines Artikels.

## Starten

```
cd wareneingang
node server.js            # http://localhost:8080
```

Es wird nur Node.js (ab Version 18) benötigt, keine weiteren Pakete.

| Variable     | Bedeutung                                              | Standard               |
|--------------|--------------------------------------------------------|------------------------|
| `PORT`       | Port des Servers                                       | `8080`                 |
| `ERP_EXPORT` | CSV-Datei, die das ERP regelmässig ablegt/überschreibt | `data/erp-export.csv`  |

## Daten

- **ERP-Export (CSV)**: Das ERP schreibt gebuchte Wareneingänge in eine CSV-Datei (`;` oder `,` getrennt). Der Server liest sie bei jeder Anfrage neu ein, und die Seite lädt sich jede Minute selbst neu. Erwartete Spalten: `WE-Nr; Eingangsdatum; Bestellnummer; Position; Artikelnummer; Bezeichnung; Lieferant; Menge; Lieferschein`. Abweichende Spaltennamen trägt man in `SPALTEN` in `server.js` ein.
- **Rückmeldungen** (bei Vogt abgegeben / i.O. / ausserhalb Toleranz) werden zentral in `data/rueckmeldungen.json` gespeichert, damit alle denselben Stand sehen.

Die mitgelieferten Dateien in `data/` sind **Beispieldaten**. Vor dem Produktivbetrieb `data/rueckmeldungen.json` löschen und `ERP_EXPORT` auf den echten Export setzen.

## Regeln (pro Artikel + Lieferant)

1. Qualifizierung: jede Lieferung geht zu Vogt, bis 3 Messungen in Folge i.O. sind.
2. 3 Lieferungen mit Freipass.
3. Die 4. Lieferung geht zur Kontrollmessung zu Vogt. Ist sie i.O., folgen nochmals 3 Freipässe.
4. Danach folgt die Requalifizierung. Ist sie i.O., geht es wieder bei Schritt 2 weiter.
5. Jede Abweichung, auch eine nachträglich gemeldete bei einer Freipass-Lieferung, führt zurück zu Schritt 1.
6. Solange ein Messergebnis aussteht, gehen auch folgende Lieferungen desselben Artikels zu Vogt.

Die Logik steht in `regeln.js`, die Tests laufen mit `node --test test/*.test.js`.

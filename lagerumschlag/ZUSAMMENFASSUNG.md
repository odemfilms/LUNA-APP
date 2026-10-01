# Lagerumschlag-Tool – Zusammenfassung

**Dateien:** `Lagerumschlag_Tool.xlsx` (Excel, nur Formeln, keine Makros) · `build_tool.py` (Neuaufbau aus neuen Exporten, übernimmt alle gelben Eingaben) · `input/` (Exporte)

## Datenstand
- Kennzahlen-Auszug: 1'464 Artikel (keine Dubletten). Hauptlager: KTL 992, PAL 365, A-BEZ 46, LIFT1 28, OCC 19, VERB-P 6, ohne Hauptlager 5, EK-TV 2, BEKLEI 1. ENTSORGEN!/DUMMY kommen im Auszug nicht vor.
- Verbrauch: Pivot «Auswertung Verbrauch», Filter Kalenderjahr **2025** (01.2025 – 12.2025), negative Werte = 0. 374 Artikel ohne Verbrauch.
- Gebinde vs. lose: KTL 855 / 137 (14 % lose), PAL 164 / 201 (55 % lose), LIFT1 1 / 27 (96 % lose). Also weniger lose Artikel als erwartet (PAL nicht über 80 %, KTL nicht ein Drittel).
- 49 Artikel (Berufskleidung BEKLEI) haben Verbrauch, kommen aber im Bestandsauszug nicht vor → nicht im Tool (wären ohnehin «Nicht im Lift»).
- MB-Blatt: 10 Artikelnummern doppelt (mehrere Lager) → MB aus Spalte AK, sonst Zeile mit Lager = Hauptlager.

## ⚠️ Wichtigster offener Punkt: Ist-Bestand
Spalte G «Bestandsmenge» ist bei **711 von 1'464 Artikeln negativ**, und zwar ab Zeile 306, nicht erst ab Zeile 1441. G hat praktisch keinen Zusammenhang mit dem Bestand im Artikelstamm (Rangkorrelation ≈ 0). **Jahresanfangsbestand 2025 + G** ist dagegen fast nie negativ und passt gut zum Artikelstamm (Rangkorrelation 0.78). Der Pivot-Zeitraum steht auf 01.01.2025 – 20.08.2026. Vermutlich zeigt G deshalb die **Netto-Bewegung seit 01.01.2025** und nicht den Lagerbestand.
- H4 = 747'814.70 CHF ist SUM(H7:H1440), negative Werte mitgerechnet. Mit «negativ = 0», wie verlangt, ergibt sich 1'331'183 CHF (Export-Werte) bzw. 1'371'980 CHF (Menge × Preis). Die Kontrolle «≈ 747'815» lässt sich also nur mit den negativen Werten erreichen. Kontrollblock: Parameter!I10.
- Umschaltbar in **Parameter B20**: «Export (Bestandsmenge)» (Standard, wie im Auftrag) oder «Jahresanfang 2025 + Export». Bitte an 2–3 Artikeln im ERP prüfen.

## Szenarien (Auslastung Lift 1 / Lift 2 / Lift 3 / Lift 2+3 / Total)
Ist-Bestand = Export G (Standard):

| | Lift 1 | Lift 2 | Lift 3 | 2+3 | Total | Passt? | Aussenlager |
|---|---|---|---|---|---|---|---|
| Ist | 25 % | 63 % | 15 % | 39 % | 34 % | JA | 0 m² |
| B · LU 2 | 85 % | 195 % | 36 % | 116 % | 106 % | NEIN | 55 m² |
| C · LU 2 | 85 % | 202 % | 37 % | 120 % | 108 % | NEIN | 68 m² |
| B · LU 3 | 79 % | 179 % | 35 % | 107 % | 98 % | NEIN | 25 m² |
| C · LU 3 | 79 % | 186 % | 37 % | 111 % | 101 % | NEIN | 39 m² |

Ist-Bestand = Jahresanfang 2025 + G:

| | Lift 1 | Lift 2 | Lift 3 | 2+3 | Total | Passt? |
|---|---|---|---|---|---|---|
| Ist | 41 % | 133 % | 34 % | 83 % | 69 % | JA (mit Ausgleich 2↔3) |
| B · LU 2 | 67 % | 210 % | 36 % | 123 % | 104 % | NEIN |
| C · LU 2 | 67 % | 229 % | 41 % | 135 % | 112 % | NEIN |
| B · LU 3 | 63 % | 191 % | 35 % | 113 % | 96 % | NEIN |
| C · LU 3 | 63 % | 210 % | 41 % | 125 % | 105 % | NEIN |

Engpass ist immer **Lift 2 + 3**, also Lift 2: dorthin gehen lose PAL-Artikel, Trennbleche und grosse Euroboxen. Lift 3 ist nur zu ~35–40 % belegt. Mit den aktuellen Annahmen passt bei **keinem LU bis 6** alles hinein, weil die Untergrenze MB + LG den MAX hochhält. Lift 1 allein passt ab LU 1.5.

## Annahmen
1. Lose Artikel: KTL 0.1 m², PAL 0.5 m², LIFT1 3.479 m² (1 Tablar), andere 0.5 m² je Artikel. Rund die Hälfte der belegten Fläche beruht auf diesen Schätzungen (Cockpit Zeilen 33–37).
2. Fehlt «Menge pro Gebinde» (2 Artikel, S71): Der ganze Ist-Bestand gilt als 1 Gebinde, mit Hinweis.
3. Lagerwert = Menge × Preis GLD Aktuell. Hauptlager aus dem Kennzahlen-Auszug (23 Artikel haben im Artikelstamm ein anderes Hauptlager).
4. MAX final und Ziel-LU je Artikel gelten in den Varianten B und C. MAX final ersetzt dort den berechneten Wert.
5. Reihenfolge für Aussenlager: je Liftgruppe nach Verbrauch pro m² absteigend.

## Offene Punkte
- Ist-Bestand (siehe oben).
- MAX = Verbrauch / LU, **mindestens MB + LG**. Bei 1'249 Artikeln bestimmt MB + LG den Wert. Dadurch ist Variante B bei LU 2 rund 6.2 Mio. CHF wert. Ist «mindestens MB + LG» so gewollt, oder eher MB + LG/2?
- Bei losen Artikeln mit kleinem Ist-Bestand wird die Fläche in B/C stark hochskaliert, z. B. 110820 Sicherungsscheibe: Ist 1, MAX 100 → 50 m². Gegebenenfalls eigene Flächen eintragen oder einen Gebinde-Typ zuweisen.
- Höhe: «max. Ladehöhe je Tablar» ist leer. Sobald sie eingetragen ist, erscheint bei zu hohen Gebinden der Hinweis «zu hoch für Tablar».

## Prüfung
- LibreOffice-Neuberechnung: 78'181 Formeln, 0 Fehler.
- LU 2 → 3 senkt die Auslastung in B und C. Variante A bleibt unabhängig vom LU. C ≥ B und C ≥ Ist gilt bei jedem Artikel.
- Ansicht 2/3 im Cockpit = Szenario-Zeile C·LU2 / C·LU3.
- 5 Handrechnungen stimmen: 101749 (Lift 1, lose), 100058 (Lift 2, S71, ohne Ist), 107997 (Lift 3, S51), 002049 (lose PAL), 005101 (kein Verbrauch).
- Neuaufbau mit `build_tool.py` übernimmt MAX final, Lift manuell, Ziel-LU, Kommentare, Parameter, Cockpit-Regler und neue Artikel.

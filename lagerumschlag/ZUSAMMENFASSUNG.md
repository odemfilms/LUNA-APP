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

## Liftbericht Lift 1 (Modula «Artikelbestand für Maschine 1», erhalten am 05.10.2026, Bericht ohne Datum)
- Datei `input/Liftbestand_Lift1_RPT_ART_GIAC_MACCHINA_MOD.prnx`, wird von `build_tool.py` eingelesen. Weitere Berichte (Maschine 2/3) einfach in `input/` legen.
- In Lift 1 liegen **48 Artikel** mit 531 Stk Fach-Kapazität und 293 Stk Bestand. Davon stehen 27 im Export unter LIFT1, 15 unter PAL, und 6 fehlen im Kennzahlen-Auszug (005301, 103683, 106582, 107412, 107414, 108768). Diese 6 sind ergänzt. Artikel 107415 steht im Export unter LIFT1, aber nicht im Bericht.
- Neue Logik: Artikel im Liftbericht bekommen dessen **Lift und Bestand**. Ihr Platz ist das **reservierte Fach**. Die gemessene Belegung (Parameter L5, **Lift 1 = 85 %**) wird im Verhältnis der Fach-Kapazität auf die Artikel verteilt. In B/C wächst die Fläche erst, wenn MAX > Fach-Kapazität. Ein-/ausschaltbar in Parameter B21. Für Lift 2/3 kann in L6/L7 ebenfalls ein gemessener Wert eingetragen werden.
- Lift 2 und Lift 3 werden erst im **Januar 2027** aufgestellt. Für sie gibt es noch keinen Liftbericht; ihr «Ist» ist die geplante Belegung, wenn der heutige KTL/PAL-Bestand dorthin umzieht (Gebinde-Flächen + Schätzwerte für lose Artikel).
- Ergebnis: Lift 1 «Ist» = **86 %** (85 % aus dem Bericht + 1 % für 107415).
- Bestandsvergleich (42 Artikel): Der Liftbericht stimmt bei 17 Artikeln mit «Jahresanfang 2025 + G» überein, mit der Export-Bestandsmenge G nur bei 5. Das bestätigt, dass G nicht der Lagerbestand ist.

## Füllgrad über Gebinde → Tablare (neu ab Version 05.10.2026)
- Gebindedaten aus der neuen Liste `input/Artikelstamm_Umschlag_2_20260623.csv`. Geprüft: 1'200 Artikel mit Gebinde, gleich wie in der Mai-Liste (nur 2 Artikel mit neuer Menge pro Gebinde: 103764, 106938). «Anzahl Gebinde» = AUFRUNDEN(Max/MpG) und «Alles Bestand» = AUFRUNDEN(Bestandsmenge/MpG) stimmen; 1 Abweichung (109201: Anzahl Gebinde leer). 2 Artikel ohne Menge pro Gebinde (105456, 111719).
- Lift 1 ist in der Liste nicht erfasst (keine Gebinde) → Lift 1 weiterhin aus dem Liftbericht (85 % der Tablare).
- **Rechenweg:** Gebinde je Tablar = beste Anordnung des Gebinde-Stellmasses auf dem Tablar 4'060 × 857 mm (längs oder quer) × Lagen. Tablare je Artikel = Anzahl Gebinde / Gebinde je Tablar. **Füllgrad = Tablare Bedarf / Tablare vorhanden (50 je Lift).**
- Gebinde je Tablar (Standard): S21 414 · S22 198 · S32 132 · S33 66 · S41 52 · S51/S52 26 · S61–S63 12 · S71–S73 6 · S81–S83 3 · P20–P25 3.
- **Neue Optionen:** Tablar-Ausnutzung (Parameter M21, z. B. 90 % = Reserve für Lücken) · Tablarmass von Lift 1/2/3 (M22) · im Blatt «Gebinde-Kategorie» je Gebinde Stellmass, Lagen (stapeln) und manuelle Anzahl je Tablar · lose Artikel als Tablar-Anteil je Hauptlager (KTL 0.03, PAL 0.15, LIFT1 1.0).

## Szenarien (Füllgrad Lift 1 / Lift 2 / Lift 3 / Lift 2+3 / Total, in Tablaren)
Mit Liftbericht, übrige Artikel Ist-Bestand = Export G (Standard):

| | Lift 1 | Lift 2 | Lift 3 | 2+3 | Total | Passt? | Aussenlager |
|---|---|---|---|---|---|---|---|
| Ist | 86 % | 73 % | 19 % | 46 % | 59 % | JA | 0 Tablare |
| B · LU 2 | 100 % | 223 % | 47 % | 135 % | 123 % | NEIN | 35 Tablare |
| C · LU 2 | 100 % | 231 % | 48 % | 140 % | 126 % | NEIN | 40 Tablare |
| B · LU 3 | 94 % | 204 % | 46 % | 125 % | 114 % | NEIN | 25 Tablare |
| C · LU 3 | 94 % | 212 % | 47 % | 130 % | 118 % | NEIN | 30 Tablare |

Mit Liftbericht, übrige Artikel Ist-Bestand = Jahresanfang 2025 + G:

| | Lift 1 | Lift 2 | Lift 3 | 2+3 | Total | Passt? |
|---|---|---|---|---|---|---|
| Ist | 86 % | 154 % | 44 % | 99 % | 95 % | JA (knapp, mit Ausgleich 2↔3) |
| B · LU 2 | 100 % | 236 % | 47 % | 141 % | 127 % | NEIN |
| C · LU 2 | 100 % | 259 % | 53 % | 156 % | 137 % | NEIN |
| B · LU 3 | 94 % | 214 % | 46 % | 130 % | 118 % | NEIN |
| C · LU 3 | 94 % | 238 % | 52 % | 145 % | 128 % | NEIN |

Engpass ist **Lift 2**. Treiber bei C · LU 2: lose Artikel (≈ 41 Tablare, geschätzt), Trennbleche S71 (≈ 34 Tablare, nur 6 je Tablar), S81 (≈ 12). Lift 3 (Kleinteile-Boxen) bleibt unter 55 %.

## Annahmen
1. Lose Artikel: Tablar-Anteil je Artikel KTL 0.03, PAL 0.15, LIFT1 1.0, andere 0.15. In Lift 2 sind rund 35 % (Ist) bis 40 % (C · LU 2) der Tablare geschätzt (Cockpit Zeilen 33–37).
2. Fehlt «Menge pro Gebinde» (2 Artikel, S71): Der ganze Ist-Bestand gilt als 1 Gebinde, mit Hinweis.
3. Lagerwert = Menge × Preis GLD Aktuell. Hauptlager aus dem Kennzahlen-Auszug (23 Artikel haben im Artikelstamm ein anderes Hauptlager).
4. MAX final und Ziel-LU je Artikel gelten in den Varianten B und C. MAX final ersetzt dort den berechneten Wert.
5. Reihenfolge für Aussenlager: je Liftgruppe nach Verbrauch pro m² absteigend.

## Offene Punkte
- Ist-Bestand (siehe oben). Für Lift 1 gilt jetzt der Liftbericht.
- Die 85 % für Lift 1 sind als **Fläche** (bzw. belegte Tablare) interpretiert. Bitte bestätigen. Liftberichte für Lift 2 und 3 würden die Schätzung dort ebenfalls ersetzen.
- MAX = Verbrauch / LU, **mindestens MB + LG**. Bei 1'249 Artikeln bestimmt MB + LG den Wert. Dadurch ist Variante B bei LU 2 rund 6.2 Mio. CHF wert. Ist «mindestens MB + LG» so gewollt, oder eher MB + LG/2?
- Bei losen Artikeln mit kleinem Ist-Bestand wird der Platz in B/C stark hochskaliert (Faktor MAX / Ist). Besser: Gebinde zuweisen (Spalte «Gebindekategorie neu») oder eigene Tablare eintragen.
- Höhe: «max. Ladehöhe je Tablar» ist leer. Sobald sie eingetragen ist, erscheint bei zu hohen Gebinden der Hinweis «zu hoch für Tablar».

## Prüfung
- LibreOffice-Neuberechnung: 80'193 Formeln, 0 Fehler (Tablar-Version). Regler-Tests erneut bestanden.
- LU 2 → 3 senkt die Auslastung in B und C. Variante A bleibt unabhängig vom LU. C ≥ B und C ≥ Ist gilt bei jedem Artikel.
- Ansicht 2/3 im Cockpit = Szenario-Zeile C·LU2 / C·LU3.
- 5 Handrechnungen stimmen: 101749 (Lift 1, lose), 100058 (Lift 2, S71, ohne Ist), 107997 (Lift 3, S51), 002049 (lose PAL), 005101 (kein Verbrauch).
- Neuaufbau mit `build_tool.py` übernimmt MAX final, Lift manuell, Ziel-LU, Kommentare, Parameter, Cockpit-Regler und neue Artikel.

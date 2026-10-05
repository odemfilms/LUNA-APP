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

## Lager in Zukunft + grobe Schätzung loser Artikel (05.10.2026)
- **Lift-Zuordnung nach Spalte «Lager in Zukunft»** der neuen Liste (Parameter I34, Schalter B19): Kardex Schwer → Lift 1 (73 Artikel), Kardex Kleinteil → Lift 2/3 nach Gebinde, PAL / Nicht NLZ / SVC Verpackung / Verpackungsmaterial / Kisten Leer / Steuerschrank & Gestell / Kompressor → nicht im Lift. Leer → Regel nach Hauptlager. Artikel im Liftbericht bleiben in Lift 1.
- **Nicht stapeln:** Lagen = 1 für alle Gebinde.
- **Lose Artikel grob geschätzt** (Spalte «Schätzklasse lose», Vorschlag aus der Bezeichnung, änderbar): Klein = Eurobox S51 à 20 Stk, Mittel = Eurobox S61 à 4 Stk, Gross = Trennblech S81 à 2 Stk (Parameter I27:K29). Mindestens der Ist-Bestand bzw. eine Losgrösse passt in 1 Gebinde.

## Ist-Bestand korrigiert (05.10.2026)
- Spalte G «Bestandsmenge» im Kennzahlen-Auszug ist die Bewegung seit 01.01.2025, nicht der Bestand → Tool zeigte nur ≈ 1.4–1.6 Mio. CHF.
- **Neu: Ist-Bestand = Jahresanfang 2025 + G** (Parameter B20). Dazu **272 KTL/PAL-Artikel ergänzt**, die Bestand im Artikelstamm (23.06.2026) haben, aber im Kennzahlen-Auszug fehlen. Lift 1: Bestand aus Liftbericht.
- **Lagerwert im Tool: 3.45 Mio. CHF** (Vorgabe aus dem ERP ≈ 3.6 Mio.; Rest v. a. Mietgeräte, Berufskleidung, Entsorgen – nicht im Lift).

## Füllgrad über Lifthöhe + Lifte gemischt (05.10.2026)
- Rechenweg wie Lagerplanung (Tabelle am Ende der Gebindeliste): Tablare = Gebinde / Gebinde je Tablar (S21 280, S22 140, S32 104, S33/S41 52, S51/S52 26, S61–S63 12, S71–S73 6, S81–S83 4). Jedes Tablar braucht **Gebindehöhe + 35 mm** (S21–S51 120 mm, S52/S62 210, S61 110, S63 310, S71/S81 200, S72/S82 400, S73/S83 600). **Füllgrad = Summe Höhe / 11'900 mm je Lift.**
- **Alle 3 Lifte gemischt** (Parameter E27 = Ja): «Passt?» und Aussenlager über die Summe aller 3 Lifte (35'700 mm). Lift 1 heute aus Liftbericht (85 % = 10'115 mm).
- Kontrolle Lagerplanung: 104 Tablare / 24'700 mm (Gebinde-Artikel) + Lift 1 heute 10'115 mm = 34'815 mm = 98 % → passt. Tool (Ist, inkl. lose Artikel geschätzt): 32'898 mm = **92 % → passt**.

## Szenarien (Füllgrad Lifthöhe, Lifte gemischt)

| | Lift 1 | Lift 2 | Lift 3 | Total | Passt? | Aussenlager |
|---|---|---|---|---|---|---|
| Ist | 107 % | 123 % | 46 % | **92 %** | **JA** (2'800 mm frei ≈ 12 Tablare) | 0 |
| B · LU 2 | 139 % | 160 % | 53 % | 117 % | NEIN | 6'229 mm ≈ 27 Tablare |
| B · LU 3 | 127 % | 144 % | 52 % | 108 % | NEIN | 2'744 mm ≈ 12 Tablare |
| C · LU 2 | 140 % | 184 % | 59 % | 128 % | NEIN | 9'883 mm ≈ 42 Tablare |
| C · LU 3 | 128 % | 170 % | 58 % | 118 % | NEIN | 6'558 mm ≈ 28 Tablare |

Die einzelnen Lift-% zeigen die Zuordnung nach «Lager in Zukunft»; gemischt wird Lift 3 mit Ware aus Lift 1/2 aufgefüllt. Variante B passt erst bei LU 6 knapp nicht (101 %).

## Annahmen
1. Lose Artikel: Tablar-Anteil je Artikel KTL 0.03, PAL 0.15, LIFT1 1.0, andere 0.15. In Lift 2 sind rund 35 % (Ist) bis 40 % (C · LU 2) der Tablare geschätzt (Cockpit Zeilen 33–37).
2. Fehlt «Menge pro Gebinde» (2 Artikel, S71): Der ganze Ist-Bestand gilt als 1 Gebinde, mit Hinweis.
3. Lagerwert = Menge × Preis GLD Aktuell. Hauptlager aus dem Kennzahlen-Auszug (23 Artikel haben im Artikelstamm ein anderes Hauptlager).
4. MAX final und Ziel-LU je Artikel gelten in den Varianten B und C. MAX final ersetzt dort den berechneten Wert.
5. Reihenfolge für Aussenlager: je Liftgruppe nach Verbrauch pro m² absteigend.

## Entscheide (05.10.2026)
- MAX mindestens MB + ganze Losgrösse: bleibt so (Vorgabe Vorgesetzter).
- PAL-Artikel ohne Eintrag «Lager in Zukunft» gehören in den Lift (Lift 2/3 nach Gebinde) – so umgesetzt.
- Gebinde werden nicht gestapelt.

## Offene Punkte
- Ist-Bestand (siehe oben). Für Lift 1 gilt jetzt der Liftbericht.
- Die 85 % für Lift 1 sind als **Fläche** (bzw. belegte Tablare) interpretiert. Bitte bestätigen. Liftberichte für Lift 2 und 3 würden die Schätzung dort ebenfalls ersetzen.
- Bei losen Artikeln mit kleinem Ist-Bestand wird der Platz in B/C stark hochskaliert (Faktor MAX / Ist). Besser: Gebinde zuweisen (Spalte «Gebindekategorie neu») oder eigene Tablare eintragen.
- Höhe: «max. Ladehöhe je Tablar» ist leer. Sobald sie eingetragen ist, erscheint bei zu hohen Gebinden der Hinweis «zu hoch für Tablar».

## Prüfung
- LibreOffice-Neuberechnung: 104'570 Formeln, 0 Fehler. Regler-Tests bestanden. Regler-Tests erneut bestanden.
- LU 2 → 3 senkt die Auslastung in B und C. Variante A bleibt unabhängig vom LU. C ≥ B und C ≥ Ist gilt bei jedem Artikel.
- Ansicht 2/3 im Cockpit = Szenario-Zeile C·LU2 / C·LU3.
- 5 Handrechnungen stimmen: 101749 (Lift 1, lose), 100058 (Lift 2, S71, ohne Ist), 107997 (Lift 3, S51), 002049 (lose PAL), 005101 (kein Verbrauch).
- Neuaufbau mit `build_tool.py` übernimmt MAX final, Lift manuell, Ziel-LU, Kommentare, Parameter, Cockpit-Regler und neue Artikel.

const test = require("node:test");
const assert = require("node:assert");
const path = require("path");
const { leseListe, altbestand } = require("../liste");
const { auswerten, STATUS } = require("../regeln");

const DATEI = path.join(__dirname, "fixtures", "liste-beispiel.xlsx");

test("liest jedes ✓ als Lieferung pro Artikel, Lieferant und Monat", () => {
  const { lieferungen } = leseListe(DATEI);
  assert.strictEqual(lieferungen.length, 10);
  const l = lieferungen.find((x) => x.id === "900001|2|2026-01");
  assert.strictEqual(l.lieferant, "Beispiel & Co AG");
  assert.strictEqual(l.bezeichnung, "Testflansch");
  assert.strictEqual(l.eingang, "2026-01");
});

test("die Abweichungstabelle unter der Pivot zählt nicht als Lieferung", () => {
  const { lieferungen, abweichungen } = leseListe(DATEI);
  assert.ok(!lieferungen.some((l) => l.artikel === "Artikel"));
  assert.deepStrictEqual(abweichungen[0], {
    artikel: "900001", massabweichung: "0.004mm", stueckzahl: "2 von 15", datum: "2026-01-20",
    bestellnummer: "80000001", lieferant: "Beispiel & Co AG", kommentar: "",
  });
});

test("Altbestand: Farben, Kommentare und Abweichungstabelle werden übernommen", () => {
  const { rueckmeldungen, nichtZugeordnet } = altbestand(leseListe(DATEI));
  assert.strictEqual(rueckmeldungen["900001|1|2026-01"].ergebnis, "OK");
  assert.strictEqual(rueckmeldungen["900001|1|2026-04"].ergebnis, "NICHT_GEMESSEN");
  assert.strictEqual(rueckmeldungen["900001|1|2026-02"].bemerkung, "Komplettmessung beim Lieferanten");
  assert.strictEqual(rueckmeldungen["900001|2|2026-01"].ergebnis, "NOK");
  assert.strictEqual(rueckmeldungen["900001|2|2026-01"].massabweichung, "0.004mm");
  assert.strictEqual(rueckmeldungen["900001|1|2026-05"], undefined);
  assert.deepStrictEqual(nichtZugeordnet, []);
});

test("neue Lieferung ohne Farbe bekommt den Status nach Regel", () => {
  const liste = leseListe(DATEI);
  const { lieferungen, artikel } = auswerten(liste.lieferungen, altbestand(liste).rueckmeldungen);
  // 3× i.O., 1 Freipass (gelb) → Mai ist Freipass 2/3
  assert.strictEqual(lieferungen.find((l) => l.id === "900001|1|2026-05").auswertung.status, STATUS.FREIPASS);
  // nach Abweichung erst 1 Messung i.O. → nächste muss zu Vogt
  assert.strictEqual(artikel["900001|2"].naechsteMussZuVogt, true);
});

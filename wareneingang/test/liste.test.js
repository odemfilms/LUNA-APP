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
  const { rueckmeldungen } = altbestand(leseListe(DATEI));
  assert.strictEqual(rueckmeldungen["900001|1|2026-01"].ergebnis, "OK");
  assert.strictEqual(rueckmeldungen["900001|1|2026-04"].ergebnis, "NICHT_GEMESSEN");
  assert.strictEqual(rueckmeldungen["900001|1|2026-02"].bemerkung, "Komplettmessung beim Lieferanten");
  assert.strictEqual(rueckmeldungen["900001|2|2026-01"].ergebnis, "NOK");
  assert.strictEqual(rueckmeldungen["900001|2|2026-01"].massabweichung, "0.004mm");
  assert.strictEqual(rueckmeldungen["900001|1|2026-05"], undefined);
});

test("neue Lieferung ohne Farbe bekommt den Status nach Regel", () => {
  const liste = leseListe(DATEI);
  const { lieferungen, artikel } = auswerten(liste.lieferungen, altbestand(liste).rueckmeldungen);
  // 3× i.O., 1 Freipass (gelb) → Mai ist Freipass 2/3
  assert.strictEqual(lieferungen.find((l) => l.id === "900001|1|2026-05").auswertung.status, STATUS.FREIPASS);
  // nach Abweichung erst 1 Messung i.O. → nächste muss zu Vogt
  assert.strictEqual(artikel["900001|2"].naechsteMussZuVogt, true);
});

test("mit Bestell-Nr. in der Pivot zählt jede Bestellung einzeln", () => {
  const liste = leseListe(path.join(__dirname, "fixtures", "liste-bestellungen.xlsx"));
  assert.deepStrictEqual(liste.lieferungen.map((l) => l.id), [
    "900001|1|80000001 - 10|2026-01",
    "900001|1|80000002 - 10|2026-01",
    "900001|1|80000003 - 10|2026-02",
  ]);
  assert.ok(liste.lieferungen.every((l) => l.lieferant === "Muster AG" && l.monatsId.startsWith("900001|1|")));
  const { rueckmeldungen } = altbestand(liste);
  assert.strictEqual(rueckmeldungen["900001|1|80000001 - 10|2026-01"].ergebnis, "OK");
  assert.strictEqual(rueckmeldungen["900001|1|80000002 - 10|2026-01"].massabweichung, "0.005mm");
  const { artikel } = auswerten(liste.lieferungen, rueckmeldungen);
  assert.match(artikel["900001|1"].naechsteLieferung, /Zu Vogt – Qualifizierung 2\/3/);
});

const test = require("node:test");
const assert = require("node:assert");
const { auswerten, STATUS } = require("../regeln");

function lieferungen(anzahl) {
  return Array.from({ length: anzahl }, (_, i) => ({
    id: "WE" + String(i + 1).padStart(3, "0"),
    eingang: "2026-01-" + String(i + 1).padStart(2, "0"),
    artikel: "A-100",
    lieferant: "Muster AG",
  }));
}

function statusFolge(liste, rueckmeldungen) {
  return auswerten(liste, rueckmeldungen).lieferungen
    .slice().reverse()
    .map((l) => l.auswertung.status);
}

function ok(...ids) {
  return Object.fromEntries(ids.map((id) => [id, { ergebnis: "OK" }]));
}

test("erste Lieferung eines neuen Artikels muss zu Vogt", () => {
  const [l] = auswerten(lieferungen(1), {}).lieferungen;
  assert.strictEqual(l.auswertung.status, STATUS.ZU_VOGT);
  assert.strictEqual(l.auswertung.regelPhase, "Qualifizierung 1/3");
});

test("voller Zyklus: 3 Quali, 3 Freipass, Kontrolle, 3 Freipass, Requali, 3 Freipass", () => {
  const liste = lieferungen(14);
  const r = ok("WE001", "WE002", "WE003", "WE007", "WE011");
  assert.deepStrictEqual(statusFolge(liste, r), [
    "OK", "OK", "OK",
    "FREIPASS", "FREIPASS", "FREIPASS",
    "OK",
    "FREIPASS", "FREIPASS", "FREIPASS",
    "OK",
    "FREIPASS", "FREIPASS", "FREIPASS",
  ]);
  const phasen = auswerten(liste, r).lieferungen.slice().reverse().map((l) => l.auswertung.regelPhase);
  assert.strictEqual(phasen[6], "Kontrollmessung");
  assert.strictEqual(phasen[10], "Requalifizierung");
});

test("Abweichung bei Kontrollmessung erzwingt 3 neue Messungen", () => {
  const liste = lieferungen(11);
  const r = Object.assign(ok("WE001", "WE002", "WE003", "WE008", "WE009", "WE010"), { WE007: { ergebnis: "NOK" } });
  assert.deepStrictEqual(statusFolge(liste, r), [
    "OK", "OK", "OK",
    "FREIPASS", "FREIPASS", "FREIPASS",
    "NOK",
    "OK", "OK", "OK",
    "FREIPASS",
  ]);
});

test("Abweichung in der Qualifizierung setzt den Zähler zurück", () => {
  const liste = lieferungen(6);
  const r = Object.assign(ok("WE001", "WE002", "WE004", "WE005"), { WE003: { ergebnis: "NOK" } });
  const res = auswerten(liste, r).lieferungen.slice().reverse();
  assert.strictEqual(res[3].auswertung.regelPhase, "Qualifizierung 1/3");
  assert.strictEqual(res[5].auswertung.regelPhase, "Qualifizierung 3/3");
  assert.strictEqual(res[5].auswertung.status, STATUS.ZU_VOGT);
});

test("nachträglich gemeldete Abweichung bei Freipass-Lieferung setzt zurück", () => {
  const liste = lieferungen(6);
  const r = Object.assign(ok("WE001", "WE002", "WE003"), { WE004: { ergebnis: "NOK" } });
  assert.deepStrictEqual(statusFolge(liste, r), [
    "OK", "OK", "OK", "FREIPASS_NOK", "ZU_VOGT", "ZU_VOGT",
  ]);
});

test("offenes Messergebnis blockiert Freipässe der Folgelieferungen", () => {
  const liste = lieferungen(5);
  const r = ok("WE001", "WE002");
  r.WE003 = { beiVogtSeit: "2026-01-03" };
  const res = auswerten(liste, r).lieferungen.slice().reverse();
  assert.strictEqual(res[2].auswertung.status, STATUS.BEI_VOGT);
  assert.strictEqual(res[3].auswertung.status, STATUS.ZU_VOGT);
  assert.match(res[3].auswertung.grund, /2026-01-03/);

  r.WE003 = { ergebnis: "OK" };
  assert.strictEqual(auswerten(liste, r).lieferungen.slice().reverse()[3].auswertung.status, STATUS.FREIPASS);
});

test("Artikel werden pro Lieferant getrennt bewertet", () => {
  const liste = lieferungen(4).concat([{ id: "WE900", eingang: "2026-02-01", artikel: "A-100", lieferant: "Andere GmbH" }]);
  const res = auswerten(liste, ok("WE001", "WE002", "WE003")).lieferungen;
  assert.strictEqual(res.find((l) => l.id === "WE004").auswertung.status, STATUS.FREIPASS);
  assert.strictEqual(res.find((l) => l.id === "WE900").auswertung.status, STATUS.ZU_VOGT);
});

test("Zusammenfassung zeigt, was die nächste Lieferung braucht", () => {
  const { artikel } = auswerten(lieferungen(6), ok("WE001", "WE002", "WE003"));
  const a = artikel["A-100|Muster AG"];
  assert.strictEqual(a.naechsteMussZuVogt, true);
  assert.match(a.naechsteLieferung, /Kontrollmessung/);
});

test("freiwillige Messung während Freipass verbraucht keinen Freipass", () => {
  const liste = lieferungen(8);
  const r = ok("WE001", "WE002", "WE003", "WE004", "WE005");
  assert.deepStrictEqual(statusFolge(liste, r), ["OK", "OK", "OK", "OK", "OK", "FREIPASS", "FREIPASS", "FREIPASS"]);
});

test("nicht gemessene Lieferung in einer Messphase wird markiert und zählt nicht", () => {
  const liste = lieferungen(3);
  const r = { WE001: { ergebnis: "NICHT_GEMESSEN" }, WE002: { ergebnis: "OK" } };
  const res = auswerten(liste, r).lieferungen.slice().reverse();
  assert.strictEqual(res[0].auswertung.status, STATUS.NICHT_GEMESSEN);
  assert.strictEqual(res[1].auswertung.regelPhase, "Qualifizierung 1/3");
  assert.strictEqual(res[2].auswertung.regelPhase, "Qualifizierung 2/3");
});

test("Artikel ohne frühere Messung bei Vogt: nächste Lieferung muss zu Vogt", () => {
  const liste = lieferungen(4);
  const r = Object.fromEntries(liste.map((l) => [l.id, { ergebnis: "NICHT_GEMESSEN" }]));
  const { artikel } = auswerten(liste, r);
  assert.strictEqual(artikel["A-100|Muster AG"].naechsteMussZuVogt, true);
  assert.match(artikel["A-100|Muster AG"].naechsteLieferung, /Qualifizierung 1\/3/);
});

test("Rückmeldung pro Monat gilt für alle Bestellungen dieses Monats", () => {
  const liste = [1, 2].map((n) => ({ id: "A|1|8000" + n + "|2026-01", monatsId: "A|1|2026-01", eingang: "2026-01", artikel: "A", lieferantId: "1" }));
  const res = auswerten(liste, { "A|1|2026-01": { ergebnis: "OK" } }).lieferungen;
  assert.ok(res.every((l) => l.auswertung.status === STATUS.OK));
});

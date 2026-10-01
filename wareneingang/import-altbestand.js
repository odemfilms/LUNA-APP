// Einmalige Übernahme der bisher in der Excel-Liste geführten Ergebnisse.
//
//   node import-altbestand.js "Lieferungen_zu_Vogt.xlsx"
//
// - Zellfarben: grün → i.O., rot → Abweichung, gelb → nicht zu Vogt
// - Tabelle "Massabweichung" → Details zur Abweichung
// - Kommentare an den ✓-Zellen → Bemerkung
// - Artikel, deren Lieferungen alle gelb sind → "nicht messpflichtig"
//
// Bereits im Dashboard erfasste Rückmeldungen werden nicht überschrieben.
// Danach werden Ergebnisse nur noch im Dashboard erfasst, nicht mehr über Farben.

const fs = require("fs");
const path = require("path");
const { leseListe, altbestand } = require("./liste");
const { artikelSchluessel } = require("./regeln");

const DATA = path.join(__dirname, "data");
const datei = process.argv[2];
if (!datei) {
  console.error('Aufruf: node import-altbestand.js "Pfad/zur/Liste.xlsx"');
  process.exit(1);
}

function laden(name) {
  try { return JSON.parse(fs.readFileSync(path.join(DATA, name), "utf8")); } catch (e) { return {}; }
}
function speichern(name, daten) {
  fs.mkdirSync(DATA, { recursive: true });
  fs.writeFileSync(path.join(DATA, name), JSON.stringify(daten, null, 2));
}

const liste = leseListe(datei);
const { rueckmeldungen, nichtZugeordnet } = altbestand(liste);

const vorhanden = laden("rueckmeldungen.json");
let neu = 0;
for (const [id, r] of Object.entries(rueckmeldungen)) {
  if (vorhanden[id]) continue;
  vorhanden[id] = r;
  neu++;
}
speichern("rueckmeldungen.json", vorhanden);

const einstellungen = laden("artikel.json");
const gruppen = {};
for (const l of liste.lieferungen) (gruppen[artikelSchluessel(l)] = gruppen[artikelSchluessel(l)] || []).push(l);
const ausgenommen = [];
for (const [k, ls] of Object.entries(gruppen)) {
  if (einstellungen[k]) continue;
  if (ls.every((l) => rueckmeldungen[l.id] && rueckmeldungen[l.id].ergebnis === "NICHT_GEMESSEN")) {
    einstellungen[k] = { ausgenommen: true, grund: "Übernahme Excel-Liste: alle Lieferungen gelb" };
    ausgenommen.push(ls[0].artikel + " " + ls[0].bezeichnung + " (" + ls[0].lieferant + ")");
  }
}
speichern("artikel.json", einstellungen);

const zaehler = {};
for (const r of Object.values(rueckmeldungen)) zaehler[r.ergebnis] = (zaehler[r.ergebnis] || 0) + 1;
console.log("Lieferungen in der Liste:", liste.lieferungen.length);
console.log("Aus Farben/Tabelle gelesen:", zaehler, "→ neu übernommen:", neu);
if (ausgenommen.length) console.log("Als nicht messpflichtig markiert:\n  " + ausgenommen.join("\n  "));
for (const a of nichtZugeordnet) {
  console.log("Abweichung ohne passendes ✓ in der Pivot:", a.artikel, a.lieferant, a.datum, a.massabweichung, "Best.", a.bestellnummer);
}

// Einmalige Übernahme der bisher in der Excel-Liste geführten Ergebnisse.
//
//   node import-altbestand.js "Lieferungen_zu_Vogt.xlsx"
//
// - Zellfarben: grün → i.O., rot → Abweichung, gelb → nicht zu Vogt
// - Tabelle "Massabweichung" → Details zur Abweichung
// - Kommentare an den ✓-Zellen → Bemerkung
//
// Bereits im Dashboard erfasste Rückmeldungen werden nicht überschrieben.
// Danach werden Ergebnisse nur noch im Dashboard erfasst, nicht mehr über Farben.

const fs = require("fs");
const path = require("path");
const { leseListe, altbestand } = require("./liste");

const DATEI_RUECKMELDUNGEN = path.join(__dirname, "data", "rueckmeldungen.json");
const datei = process.argv[2];
if (!datei) {
  console.error('Aufruf: node import-altbestand.js "Pfad/zur/Liste.xlsx"');
  process.exit(1);
}

let vorhanden = {};
try { vorhanden = JSON.parse(fs.readFileSync(DATEI_RUECKMELDUNGEN, "utf8")); } catch (e) {}

const liste = leseListe(datei);
const { rueckmeldungen } = altbestand(liste);
let neu = 0;
for (const [id, r] of Object.entries(rueckmeldungen)) {
  if (vorhanden[id]) continue;
  vorhanden[id] = r;
  neu++;
}
fs.mkdirSync(path.dirname(DATEI_RUECKMELDUNGEN), { recursive: true });
fs.writeFileSync(DATEI_RUECKMELDUNGEN, JSON.stringify(vorhanden, null, 2));

const zaehler = {};
for (const r of Object.values(rueckmeldungen)) zaehler[r.ergebnis] = (zaehler[r.ergebnis] || 0) + 1;
console.log("Lieferungen in der Liste:", liste.lieferungen.length);
console.log("Aus Farben/Tabelle gelesen:", zaehler, "→ neu übernommen:", neu);

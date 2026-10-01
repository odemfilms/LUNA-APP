// Liest die Excel-Liste "Lieferungen zu Vogt" (Pivot aus dem BI, Artikel × Lieferant × Geschäftsmonat).
//
// Jedes ✓ in einer Monatsspalte ist eine Lieferung. Sind in der Pivot zusätzlich
// "Bestell-Nr." oder "Bestell-Nr. - Position" als Zeilenfelder eingeblendet, wird
// jede Bestellposition einzeln gezählt. Ergebnisse, die noch pro Monat erfasst wurden,
// gelten dann über "monatsId" für alle Bestellungen dieses Monats.

const { lesen, excelDatum } = require("./xlsx");

const KOPF = {
  artikel: ["Artikel ID", "Artikel-Nr.", "Artikel"],
  bezeichnung: ["Artikelbezeichnung"],
  materialgruppe: ["Materialgruppe"],
  lieferantId: ["Lieferant ID"],
  lieferant: ["Lieferant Bez.", "Lieferant"],
  bestellung: ["Bestell-Nr. - Position", "Bestell-Nr.", "Bestellnummer"],
};
const MONAT = /^\d{4}-\d{2}$/;

function text(z) {
  return z && z.wert != null ? String(z.wert).trim() : "";
}

// Ordnet eine Füllfarbe der bisherigen Legende zu.
function farbErgebnis(farbe) {
  if (!farbe) return null;
  const r = parseInt(farbe.slice(1, 3), 16), g = parseInt(farbe.slice(3, 5), 16), b = parseInt(farbe.slice(5, 7), 16);
  if (r > 200 && g > 200 && b < 120) return "NICHT_GEMESSEN"; // gelb: geht nicht zu Vogt
  if (g > r && g > b) return "OK";                              // grün: keine Abweichung
  if (r > g && r > b) return "NOK";                             // rot: über 0.002 mm Abweichung
  return null;
}

function findeKopf(blatt) {
  for (const [k, z] of blatt.zellen) {
    if (text(z) !== "Artikel ID") continue;
    const zeile = +k.split(":")[0];
    const spalten = {};
    const monate = [];
    for (const [k2, z2] of blatt.zellen) {
      const [r, c] = k2.split(":").map(Number);
      if (r !== zeile) continue;
      const t = text(z2);
      if (MONAT.test(t)) monate.push({ spalte: c, monat: t });
      for (const [feld, namen] of Object.entries(KOPF)) if (namen.includes(t) && !spalten[feld]) spalten[feld] = c;
    }
    if (monate.length) return { zeile, spalten, monate: monate.sort((a, b) => a.spalte - b.spalte) };
  }
  return null;
}

function lesePivot(blatt, kopf) {
  const pivot = blatt.pivots.find((p) => p.r1 <= kopf.zeile && p.r2 > kopf.zeile);
  const letzteZeile = pivot ? pivot.r2 : Infinity;
  const felder = Object.keys(kopf.spalten);
  const lieferungen = [];
  const vorher = {};
  for (let r = kopf.zeile + 1; r <= letzteZeile; r++) {
    const werte = {};
    for (const f of felder) werte[f] = text(blatt.zelle(r, kopf.spalten[f]));
    if (!pivot && (!felder.some((f) => werte[f]) || werte.artikel === "Artikel")) break;
    if (/^(Gesamt|Grand Total)/i.test(werte.artikel)) break;
    // Leere Zeilenbeschriftungen der Pivot von oben übernehmen
    for (const f of felder) { if (werte[f]) vorher[f] = werte[f]; else werte[f] = vorher[f] || ""; }

    for (const { spalte, monat } of kopf.monate) {
      const z = blatt.zelle(r, spalte);
      const geliefert = z && (z.wert === "✓" || (typeof z.wert === "number" && z.wert > 0));
      if (!geliefert) continue;
      const lieferantKey = werte.lieferantId || werte.lieferant;
      lieferungen.push({
        id: [werte.artikel, lieferantKey, werte.bestellung, monat].filter(Boolean).join("|"),
        monatsId: [werte.artikel, lieferantKey, monat].join("|"),
        eingang: monat,
        artikel: werte.artikel,
        bezeichnung: werte.bezeichnung,
        materialgruppe: werte.materialgruppe,
        lieferantId: werte.lieferantId,
        lieferant: werte.lieferant,
        bestellung: werte.bestellung,
        _farbe: z.farbe,
        _kommentar: (z.kommentar || "").replace(/\s*\r?\n\s*/g, " "),
      });
    }
  }
  return lieferungen;
}

// Bisher manuell geführte Tabelle "Artikel | Massabweichung | Stückzahl | Datum | Bestellnummer | Lieferant"
function leseAbweichungen(blatt) {
  for (const [k, z] of blatt.zellen) {
    if (text(z) !== "Massabweichung") continue;
    const zeile = +k.split(":")[0];
    const spalten = {};
    for (const [k2, z2] of blatt.zellen) {
      const [r, c] = k2.split(":").map(Number);
      if (r === zeile && text(z2)) spalten[text(z2)] = c;
    }
    const liste = [];
    for (let r = zeile + 1; ; r++) {
      const artikel = text(blatt.zelle(r, spalten.Artikel));
      if (!artikel) break;
      const massZelle = blatt.zelle(r, spalten.Massabweichung);
      liste.push({
        artikel: artikel.replace(/\.0$/, ""),
        massabweichung: text(massZelle),
        stueckzahl: text(blatt.zelle(r, spalten["Stückzahl"])),
        datum: excelDatum(blatt.zelle(r, spalten.Datum)?.wert),
        bestellnummer: text(blatt.zelle(r, spalten.Bestellnummer)),
        lieferant: text(blatt.zelle(r, spalten.Lieferant)),
        kommentar: ((massZelle && massZelle.kommentar) || "").replace(/\s*\r?\n\s*/g, " "),
      });
    }
    return liste;
  }
  return [];
}

function leseListe(pfad) {
  for (const blatt of lesen(pfad)) {
    const kopf = findeKopf(blatt);
    if (!kopf) continue;
    return { lieferungen: lesePivot(blatt, kopf), abweichungen: leseAbweichungen(blatt) };
  }
  throw new Error("In " + pfad + " wurde keine Pivot mit 'Artikel ID' und Monatsspalten (JJJJ-MM) gefunden");
}

// Für die Übernahme des Altbestands: Farben, Kommentare und Abweichungstabelle → Rückmeldungen
function altbestand(liste) {
  const rueckmeldungen = {};
  for (const l of liste.lieferungen) {
    const ergebnis = farbErgebnis(l._farbe);
    if (!ergebnis) continue;
    rueckmeldungen[l.id] = { ergebnis, bemerkung: l._kommentar, erfasstVon: "Übernahme Excel-Liste", datum: null };
  }
  for (const a of liste.abweichungen) {
    const monat = a.datum.slice(0, 7);
    const passend = (x) => x.artikel === a.artikel && x.eingang === monat && x.lieferant === a.lieferant;
    // Mit Bestellnummer in der Pivot genau die Bestellung, sonst der Monat
    const l = liste.lieferungen.find((x) => passend(x) && x.bestellung && a.bestellnummer && x.bestellung.startsWith(a.bestellnummer))
      || liste.lieferungen.find(passend);
    if (!l) continue; // Abweichung ohne passendes ✓ wird ignoriert
    rueckmeldungen[l.id] = Object.assign({}, rueckmeldungen[l.id], {
      ergebnis: "NOK",
      massabweichung: a.massabweichung,
      stueckzahl: a.stueckzahl,
      bestellnummer: a.bestellnummer,
      bemerkung: [rueckmeldungen[l.id] && rueckmeldungen[l.id].bemerkung, a.kommentar].filter(Boolean).join(" · "),
      erfasstVon: "Übernahme Excel-Liste",
      datum: a.datum,
    });
  }
  return { rueckmeldungen };
}

function ohneIntern(l) {
  const { _farbe, _kommentar, ...rest } = l;
  return rest;
}

module.exports = { leseListe, altbestand, farbErgebnis, ohneIntern };

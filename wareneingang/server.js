// Kleiner Server ohne Abhängigkeiten für das Wareneingangs-Dashboard.
//
//   node server.js
//
// - liest den ERP-Export (CSV) bei jeder Anfrage neu ein → Liste aktualisiert sich selbst
// - speichert Rückmeldungen von Vogt (i.O. / Abweichung) zentral in data/rueckmeldungen.json,
//   damit alle Mitarbeitenden denselben Stand sehen
//
// Einstellungen über Umgebungsvariablen:
//   PORT        Port (Standard 8080)
//   ERP_EXPORT  Pfad zur CSV-Datei, die das ERP regelmässig ablegt (Standard data/erp-export.csv)

const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;
const ERP_EXPORT = process.env.ERP_EXPORT || path.join(ROOT, "data", "erp-export.csv");
const RUECKMELDUNGEN = path.join(ROOT, "data", "rueckmeldungen.json");

// Spaltennamen im ERP-Export → interne Felder. Bei Bedarf an den Export anpassen.
const SPALTEN = {
  id: ["WE-Nr", "Wareneingang", "Wareneingangsnummer"],
  eingang: ["Eingangsdatum", "Datum", "Buchungsdatum"],
  bestellung: ["Bestellnummer", "Bestellung", "Bestell-Nr"],
  position: ["Position", "Pos"],
  artikel: ["Artikelnummer", "Artikel", "Art-Nr"],
  bezeichnung: ["Bezeichnung", "Artikelbezeichnung"],
  lieferant: ["Lieferant", "Lieferantenname"],
  menge: ["Menge"],
  lieferschein: ["Lieferschein", "Lieferscheinnummer"],
};

const MIME = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".svg": "image/svg+xml", ".json": "application/json; charset=utf-8" };
const STATISCH = new Set(["/index.html", "/app.js", "/regeln.js"]);

function csvZeilen(text) {
  text = text.replace(/^﻿/, "");
  const erste = text.split(/\r?\n/, 1)[0];
  const trenner = (erste.match(/;/g) || []).length >= (erste.match(/,/g) || []).length ? ";" : ",";
  const zeilen = [];
  let feld = "", zeile = [], inAnf = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inAnf) {
      if (c === '"' && text[i + 1] === '"') { feld += '"'; i++; }
      else if (c === '"') inAnf = false;
      else feld += c;
    } else if (c === '"') inAnf = true;
    else if (c === trenner) { zeile.push(feld); feld = ""; }
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      zeile.push(feld); feld = "";
      if (zeile.some((f) => f.trim() !== "")) zeilen.push(zeile);
      zeile = [];
    } else feld += c;
  }
  zeile.push(feld);
  if (zeile.some((f) => f.trim() !== "")) zeilen.push(zeile);
  return zeilen;
}

// Akzeptiert 2026-03-14, 14.03.2026 und 14.03.26
function datum(wert) {
  wert = (wert || "").trim();
  let m = wert.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  m = wert.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})/);
  if (m) {
    const jahr = m[3].length === 2 ? "20" + m[3] : m[3];
    return jahr + "-" + m[2].padStart(2, "0") + "-" + m[1].padStart(2, "0");
  }
  return wert;
}

function erpLesen() {
  const zeilen = csvZeilen(fs.readFileSync(ERP_EXPORT, "utf8"));
  const kopf = zeilen.shift().map((k) => k.trim().toLowerCase());
  const index = {};
  for (const [feld, namen] of Object.entries(SPALTEN)) {
    index[feld] = kopf.findIndex((k) => namen.some((n) => n.toLowerCase() === k));
  }
  return zeilen.map((z) => {
    const l = {};
    for (const feld of Object.keys(SPALTEN)) l[feld] = index[feld] >= 0 ? (z[index[feld]] || "").trim() : "";
    l.eingang = datum(l.eingang);
    if (!l.id) l.id = [l.bestellung, l.position, l.lieferschein].filter(Boolean).join("-");
    return l;
  });
}

function rueckmeldungenLesen() {
  try { return JSON.parse(fs.readFileSync(RUECKMELDUNGEN, "utf8")); }
  catch (e) { return {}; }
}

function rueckmeldungenSchreiben(daten) {
  const tmp = RUECKMELDUNGEN + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(daten, null, 2));
  fs.renameSync(tmp, RUECKMELDUNGEN);
}

function json(res, status, daten) {
  res.writeHead(status, { "Content-Type": MIME[".json"], "Cache-Control": "no-store" });
  res.end(JSON.stringify(daten));
}

function body(req) {
  return new Promise((resolve, reject) => {
    let s = "";
    req.on("data", (c) => { s += c; if (s.length > 100000) req.destroy(); });
    req.on("end", () => { try { resolve(JSON.parse(s || "{}")); } catch (e) { reject(e); } });
    req.on("error", reject);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://localhost");
  try {
    if (url.pathname === "/api/daten" && req.method === "GET") {
      const stat = fs.statSync(ERP_EXPORT);
      return json(res, 200, { lieferungen: erpLesen(), rueckmeldungen: rueckmeldungenLesen(), erpStand: stat.mtime.toISOString() });
    }

    if (url.pathname === "/api/rueckmeldung" && req.method === "POST") {
      const { id, aktion, bemerkung, erfasstVon } = await body(req);
      if (!id || !["BEI_VOGT", "OK", "NOK", "ZURUECKSETZEN"].includes(aktion)) return json(res, 400, { fehler: "id und gültige aktion nötig" });
      const alle = rueckmeldungenLesen();
      const jetzt = new Date().toISOString();
      if (aktion === "ZURUECKSETZEN") delete alle[id];
      else if (aktion === "BEI_VOGT") alle[id] = Object.assign({}, alle[id], { beiVogtSeit: jetzt });
      else alle[id] = Object.assign({}, alle[id], { ergebnis: aktion, datum: jetzt, bemerkung: bemerkung || "", erfasstVon: erfasstVon || "" });
      rueckmeldungenSchreiben(alle);
      return json(res, 200, { ok: true, rueckmeldung: alle[id] || null });
    }

    const datei = url.pathname === "/" ? "/index.html" : url.pathname;
    if (req.method === "GET" && STATISCH.has(datei)) {
      res.writeHead(200, { "Content-Type": MIME[path.extname(datei)], "Cache-Control": "no-cache" });
      return fs.createReadStream(path.join(ROOT, datei)).pipe(res);
    }
    json(res, 404, { fehler: "nicht gefunden" });
  } catch (e) {
    json(res, 500, { fehler: e.message });
  }
});

if (require.main === module) {
  server.listen(PORT, () => console.log("Wareneingang-Dashboard: http://localhost:" + PORT + "  (ERP-Export: " + ERP_EXPORT + ")"));
}

module.exports = { csvZeilen, datum, server };

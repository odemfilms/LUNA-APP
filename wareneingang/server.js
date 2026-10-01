// Kleiner Server ohne Abhängigkeiten für das Wareneingangs-Dashboard.
//
//   node server.js
//
// - liest die Excel-Liste "Lieferungen zu Vogt" (Pivot aus dem BI) neu ein, sobald sie sich ändert
// - speichert Rückmeldungen von Vogt (i.O. / Abweichung) zentral in data/rueckmeldungen.json
//   und Artikel-Einstellungen in data/artikel.json, damit alle denselben Stand sehen
//
// Einstellungen über Umgebungsvariablen:
//   PORT   Port (Standard 8080)
//   LISTE  Pfad zur Excel-Liste (Standard data/Lieferungen_zu_Vogt.xlsx)

const http = require("http");
const fs = require("fs");
const path = require("path");
const { leseListe, ohneIntern } = require("./liste");

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;
const LISTE = process.env.LISTE || path.join(ROOT, "data", "Lieferungen_zu_Vogt.xlsx");
const RUECKMELDUNGEN = path.join(ROOT, "data", "rueckmeldungen.json");
const ARTIKEL = path.join(ROOT, "data", "artikel.json");

const MIME = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".json": "application/json; charset=utf-8" };
const STATISCH = new Set(["/index.html", "/app.js", "/regeln.js"]);

// Die Excel-Liste wird nur neu eingelesen, wenn sich die Datei geändert hat.
let cache = { mtime: 0, daten: null };
function listeLesen() {
  const stat = fs.statSync(LISTE);
  if (stat.mtimeMs !== cache.mtime) {
    cache = { mtime: stat.mtimeMs, daten: leseListe(LISTE).lieferungen.map(ohneIntern), stand: stat.mtime.toISOString() };
  }
  return cache;
}

function jsonLesen(datei) {
  try { return JSON.parse(fs.readFileSync(datei, "utf8")); }
  catch (e) { return {}; }
}

function jsonSchreiben(datei, daten) {
  fs.mkdirSync(path.dirname(datei), { recursive: true });
  const tmp = datei + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(daten, null, 2));
  fs.renameSync(tmp, datei);
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
      const liste = listeLesen();
      return json(res, 200, { lieferungen: liste.daten, rueckmeldungen: jsonLesen(RUECKMELDUNGEN), artikel: jsonLesen(ARTIKEL), listeStand: liste.stand });
    }

    if (url.pathname === "/api/rueckmeldung" && req.method === "POST") {
      const b = await body(req);
      const aktionen = ["BEI_VOGT", "OK", "NOK", "NICHT_GEMESSEN", "ZURUECKSETZEN"];
      if (!b.id || !aktionen.includes(b.aktion)) return json(res, 400, { fehler: "id und gültige aktion nötig" });
      const alle = jsonLesen(RUECKMELDUNGEN);
      const jetzt = new Date().toISOString();
      if (b.aktion === "ZURUECKSETZEN") delete alle[b.id];
      else if (b.aktion === "BEI_VOGT") alle[b.id] = Object.assign({}, alle[b.id], { beiVogtSeit: jetzt });
      else {
        const felder = {};
        for (const f of ["bemerkung", "erfasstVon", "massabweichung", "stueckzahl", "bestellnummer"]) felder[f] = String(b[f] || "").slice(0, 500);
        alle[b.id] = Object.assign({}, alle[b.id], felder, { ergebnis: b.aktion, datum: jetzt });
      }
      jsonSchreiben(RUECKMELDUNGEN, alle);
      return json(res, 200, { ok: true, rueckmeldung: alle[b.id] || null });
    }

    if (url.pathname === "/api/artikel" && req.method === "POST") {
      const b = await body(req);
      if (!b.schluessel) return json(res, 400, { fehler: "schluessel nötig" });
      const alle = jsonLesen(ARTIKEL);
      if (b.ausgenommen) alle[b.schluessel] = { ausgenommen: true, grund: String(b.grund || "").slice(0, 500), erfasstVon: String(b.erfasstVon || "").slice(0, 100), datum: new Date().toISOString() };
      else delete alle[b.schluessel];
      jsonSchreiben(ARTIKEL, alle);
      return json(res, 200, { ok: true });
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
  server.listen(PORT, () => console.log("Wareneingang-Dashboard: http://localhost:" + PORT + "  (Liste: " + LISTE + ")"));
}

module.exports = { server };

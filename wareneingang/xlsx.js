// Minimaler .xlsx-Leser ohne Abhängigkeiten: liest Zellwerte, Füllfarben und Kommentare.

const fs = require("fs");
const zlib = require("zlib");

function entpacken(puffer) {
  let eocd = -1;
  for (let i = puffer.length - 22; i >= Math.max(0, puffer.length - 65557); i--) {
    if (puffer.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("Keine gültige .xlsx-Datei");
  const anzahl = puffer.readUInt16LE(eocd + 10);
  let pos = puffer.readUInt32LE(eocd + 16);
  const dateien = {};
  for (let n = 0; n < anzahl; n++) {
    const methode = puffer.readUInt16LE(pos + 10);
    const groesse = puffer.readUInt32LE(pos + 20);
    const nameLen = puffer.readUInt16LE(pos + 28);
    const extraLen = puffer.readUInt16LE(pos + 30);
    const kommentarLen = puffer.readUInt16LE(pos + 32);
    const lokal = puffer.readUInt32LE(pos + 42);
    const name = puffer.toString("utf8", pos + 46, pos + 46 + nameLen);
    const start = lokal + 30 + puffer.readUInt16LE(lokal + 26) + puffer.readUInt16LE(lokal + 28);
    const daten = puffer.subarray(start, start + groesse);
    dateien[name] = methode === 8 ? zlib.inflateRawSync(daten) : daten;
    pos += 46 + nameLen + extraLen + kommentarLen;
  }
  return dateien;
}

function xmlText(s) {
  return s.replace(/&(lt|gt|quot|apos|amp|#\d+|#x[0-9a-f]+);/gi, (_, e) => {
    if (e[0] === "#") return String.fromCodePoint(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    return { lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" }[e.toLowerCase()];
  });
}

function texte(xml) {
  return xmlText((xml.match(/<t[^>]*>[\s\S]*?<\/t>/g) || []).map((t) => t.replace(/<[^>]+>/g, "")).join(""));
}

function attr(s, name) {
  const m = s.match(new RegExp("\\b" + name + '="([^"]*)"'));
  return m ? m[1] : null;
}

function spalteZuZahl(buchstaben) {
  let n = 0;
  for (const c of buchstaben) n = n * 26 + c.charCodeAt(0) - 64;
  return n;
}

function bereich(ref) {
  const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  return m && { c1: spalteZuZahl(m[1]), r1: +m[2], c2: spalteZuZahl(m[3]), r2: +m[4] };
}

// Liefert für jede Füllung "#RRGGBB" (oder null), indiziert nach Stil-Index der Zelle.
function stilFarben(stylesXml) {
  if (!stylesXml) return [];
  const fills = ((stylesXml.match(/<fills[^>]*>([\s\S]*?)<\/fills>/) || [])[1] || "").match(/<fill>[\s\S]*?<\/fill>|<fill\/>/g) || [];
  const farben = fills.map((f) => {
    if (!/patternType="solid"/.test(f)) return null;
    const rgb = attr((f.match(/<fgColor[^>]*>/) || [""])[0], "rgb");
    return rgb ? "#" + rgb.slice(-6).toUpperCase() : null;
  });
  const xfs = ((stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/) || [])[1] || "").match(/<xf\b[^>]*>/g) || [];
  return xfs.map((xf) => farben[Number(attr(xf, "fillId") || 0)] || null);
}

// Liest ein Blatt als Map "zeile:spalte" → { wert, farbe, kommentar }
function lesen(pfad) {
  const dateien = entpacken(fs.readFileSync(pfad));
  const str = (name) => (dateien[name] ? dateien[name].toString("utf8") : null);
  const geteilt = ((str("xl/sharedStrings.xml") || "").match(/<si>[\s\S]*?<\/si>/g) || []).map(texte);
  const farben = stilFarben(str("xl/styles.xml"));

  const blaetter = Object.keys(dateien).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();
  return blaetter.map((name) => {
    const xml = str(name);
    const zellen = new Map();
    const re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = re.exec(xml))) {
      const attrs = m[1], inhalt = m[2] || "";
      const ref = attr(attrs, "r").match(/^([A-Z]+)(\d+)$/);
      const typ = attr(attrs, "t");
      const v = (inhalt.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      let wert = null;
      if (typ === "s" && v != null) wert = geteilt[Number(v)];
      else if (typ === "inlineStr") wert = texte(inhalt);
      else if (v != null) wert = typ === "str" || typ === "e" ? xmlText(v) : (typ === "b" ? v === "1" : Number(v));
      const farbe = farben[Number(attr(attrs, "s") || 0)] || null;
      if (wert === null && !farbe) continue;
      zellen.set(+ref[2] + ":" + spalteZuZahl(ref[1]), { wert, farbe });
    }

    // Kommentare und Pivot-Bereiche über die Beziehungen des Blatts
    const rels = str(name.replace("worksheets/", "worksheets/_rels/") + ".rels") || "";
    const beziehungen = (rels.match(/<Relationship\b[^>]*>/g) || []).map((r) => ({
      typ: (attr(r, "Type") || "").split("/").pop(),
      ziel: "xl/" + attr(r, "Target").replace(/^\.\.\//, "").replace(/^\/xl\//, ""),
    }));
    const ziele = beziehungen.map((b) => b.ziel);
    for (const ziel of beziehungen.filter((b) => b.typ === "comments").map((b) => b.ziel)) {
      for (const k of (str(ziel) || "").match(/<comment\b[\s\S]*?<\/comment>/g) || []) {
        const ref = attr(k, "ref").match(/^([A-Z]+)(\d+)$/);
        const schluessel = +ref[2] + ":" + spalteZuZahl(ref[1]);
        const text = texte(k).split(/Kommentar:\s*/).pop().trim();
        zellen.set(schluessel, Object.assign({ wert: null, farbe: null }, zellen.get(schluessel), { kommentar: text }));
      }
    }
    const pivots = ziele.filter((z) => /pivotTable\d*\.xml$/.test(z))
      .map((z) => bereich(attr((str(z) || "").match(/<location\b[^>]*>/)?.[0] || "", "ref") || ""))
      .filter(Boolean);

    return { name, zellen, pivots, zelle: (r, c) => zellen.get(r + ":" + c) };
  });
}

function excelDatum(wert) {
  if (typeof wert !== "number") return String(wert || "").trim();
  return new Date(Math.round((wert - 25569) * 86400000)).toISOString().slice(0, 10);
}

module.exports = { lesen, excelDatum, spalteZuZahl };

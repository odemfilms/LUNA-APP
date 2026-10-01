(function () {
  const { auswerten, STATUS } = window.Regeln;
  const AKTUALISIEREN_MS = 60 * 1000;
  const FREIPASS_TAGE = 62;

  const STATUS_TEXT = {
    ZU_VOGT: "Zu Vogt bringen",
    BEI_VOGT: "Bei Vogt · Ergebnis offen",
    OK: "Gemessen i.O.",
    NOK: "Abweichung",
    FREIPASS: "Freipass",
    FREIPASS_NOK: "Freipass · Abweichung gemeldet",
    NICHT_GEMESSEN: "Nicht gemessen",
    AUSGENOMMEN: "Nicht messpflichtig",
  };
  const OFFEN = [STATUS.ZU_VOGT, STATUS.BEI_VOGT];
  const ABWEICHUNG = [STATUS.NOK, STATUS.FREIPASS_NOK];

  let daten = null;
  let auswertung = { lieferungen: [], artikel: {} };
  let filter = "ALLE";
  let suche = "";
  let offenerArtikel = null;

  const $ = (id) => document.getElementById(id);

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  // "2026-09" → "09.2026", "2026-09-14" → "14.09.2026"
  function zeitpunkt(wert) {
    if (!wert) return "";
    const [j, m, t] = wert.slice(0, 10).split("-");
    return t ? t + "." + m + "." + j : m + "." + j;
  }
  function zeit(iso) {
    return new Date(iso).toLocaleString("de-CH", { dateStyle: "short", timeStyle: "short" });
  }
  function tageSeit(wert) {
    return Math.floor((Date.now() - new Date(wert.slice(0, 10)).getTime()) / 86400000);
  }
  function badge(status) {
    return '<span class="badge s-' + status + '">' + esc(STATUS_TEXT[status]) + "</span>";
  }
  function name() {
    try { return localStorage.getItem("we-name") || ""; } catch (e) { return ""; }
  }
  function nameMerken(n) {
    try { localStorage.setItem("we-name", n); } catch (e) {}
  }
  function passt(l) {
    if (!suche) return true;
    const s = suche.toLowerCase();
    return [l.artikel, l.bezeichnung, l.lieferant, l.lieferantId, l.bestellung].some((f) => String(f || "").toLowerCase().includes(s));
  }

  async function laden() {
    try {
      const res = await fetch("api/daten", { cache: "no-store" });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).fehler || "Server antwortet mit " + res.status);
      daten = await res.json();
      auswertung = auswerten(daten.lieferungen, daten.rueckmeldungen, daten.artikel);
      $("fehler").hidden = true;
      $("stand").textContent = "Liste zuletzt aktualisiert " + zeit(daten.listeStand) + " · geprüft " + zeit(new Date().toISOString()) +
        " · " + Object.keys(auswertung.artikel).length + " Artikel/Lieferanten · " + daten.lieferungen.length + " Lieferungen";
      zeichnen();
    } catch (e) {
      $("fehler").hidden = false;
      $("fehler").textContent = "Daten konnten nicht geladen werden (" + e.message + "). Läuft der Server und ist der Pfad zur Excel-Liste richtig?";
    }
  }

  async function senden(url, nutzlast) {
    const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(nutzlast) });
    if (!res.ok) { alert("Speichern fehlgeschlagen: " + (await res.text())); return; }
    await laden();
  }

  function zeichnen() {
    const alle = auswertung.lieferungen;
    const offen = alle.filter((l) => OFFEN.includes(l.auswertung.status));
    const freipass = alle.filter((l) => l.auswertung.status === STATUS.FREIPASS && tageSeit(l.eingang) <= FREIPASS_TAGE);
    const artikel = Object.entries(auswertung.artikel);
    const naechsteVogt = artikel.filter(([, a]) => a.naechsteMussZuVogt).length;
    const naechsteFrei = artikel.filter(([, a]) => !a.naechsteMussZuVogt && !a.ausgenommen).length;
    const abw = alle.filter((l) => ABWEICHUNG.includes(l.auswertung.status) && tageSeit(l.eingang) <= 365);

    $("kpis").innerHTML = [
      ["var(--rot)", offen.filter((l) => l.auswertung.status === STATUS.ZU_VOGT).length, "Lieferungen jetzt zu Vogt bringen"],
      ["var(--gelb)", offen.filter((l) => l.auswertung.status === STATUS.BEI_VOGT).length, "bei Vogt, Ergebnis offen"],
      ["var(--rot)", naechsteVogt, "Artikel: nächste Lieferung muss zu Vogt"],
      ["var(--blau)", naechsteFrei, "Artikel: nächste Lieferung hat Freipass"],
      ["var(--rot)", abw.length, "Abweichungen (12 Monate)"],
    ].map(([farbe, zahl, text]) => '<div class="kpi" style="--farbe:' + farbe + '"><b>' + zahl + "</b><span>" + text + "</span></div>").join("");

    const offenSichtbar = offen.filter(passt).sort((a, b) => (a.eingang < b.eingang ? -1 : 1));
    $("liste-vogt").innerHTML = offenSichtbar.length
      ? offenSichtbar.map((l) => karte(l, true)).join("")
      : '<div class="leer">Keine gelieferte Ware wartet auf eine Messung.</div>';

    const fpSichtbar = freipass.filter(passt);
    $("liste-freipass").innerHTML = fpSichtbar.length
      ? fpSichtbar.map((l) => karte(l, false)).join("")
      : '<div class="leer">Keine Freipass-Lieferungen in den letzten 2 Monaten.</div>';

    // Artikelübersicht: was passiert mit der nächsten Lieferung?
    const artikelSichtbar = artikel.filter(([, a]) => passt(a)).sort(([, a], [, b]) =>
      (b.naechsteMussZuVogt - a.naechsteMussZuVogt) || (a.ausgenommen - b.ausgenommen) || a.artikel.localeCompare(b.artikel) || a.lieferant.localeCompare(b.lieferant));
    $("artikel").innerHTML = artikelSichtbar.map(([k, a]) => {
      const klasse = a.ausgenommen ? "s-AUSGENOMMEN" : a.naechsteMussZuVogt ? "s-ZU_VOGT" : "s-FREIPASS";
      return '<tr data-schluessel="' + esc(k) + '"><td><b>' + esc(a.artikel) + "</b><br><small>" + esc(a.bezeichnung) + "</small></td><td>" + esc(a.lieferant) + "</td>" +
        "<td>" + verlauf(k) + "</td><td>" + zeitpunkt(a.letzteLieferung) + "</td>" +
        '<td><span class="badge ' + klasse + '">' + esc(a.naechsteLieferung) + "</span></td></tr>";
    }).join("");

    const filterListe = [["ALLE", "Alle"], ["OFFEN", "Offen"], ["FREIPASS", "Freipass"], ["OK", "i.O."], ["ABW", "Abweichung"], ["NICHT", "Nicht gemessen"]];
    $("filter").innerHTML = filterListe.map(([k, t]) => '<button class="chip" data-filter="' + k + '" aria-pressed="' + (filter === k) + '">' + t + "</button>").join("");

    const tab = alle.filter(passt).filter((l) => {
      const s = l.auswertung.status;
      if (filter === "OFFEN") return OFFEN.includes(s);
      if (filter === "FREIPASS") return s === STATUS.FREIPASS;
      if (filter === "OK") return s === STATUS.OK;
      if (filter === "ABW") return ABWEICHUNG.includes(s);
      if (filter === "NICHT") return s === STATUS.NICHT_GEMESSEN;
      return true;
    });
    $("tabelle").innerHTML = tab.length ? tab.map((l) =>
      '<tr data-schluessel="' + esc(l.auswertung.schluessel) + '">' +
      "<td>" + zeitpunkt(l.eingang) + "</td><td><b>" + esc(l.artikel) + "</b><br><small>" + esc(l.bezeichnung) + "</small></td><td>" + esc(l.lieferant) + "</td>" +
      '<td><span class="badge s-phase">' + esc(l.auswertung.regelPhase) + "</span></td><td>" + badge(l.auswertung.status) + "</td></tr>"
    ).join("") : '<tr><td colspan="5" class="leer">Keine Lieferungen gefunden.</td></tr>';

    if (offenerArtikel && $("dialog").open) detail(offenerArtikel);
  }

  // Kleine Punktreihe der letzten Lieferungen (älteste links)
  function verlauf(schluessel) {
    const punkte = auswertung.lieferungen.filter((l) => l.auswertung.schluessel === schluessel).slice(0, 12).reverse();
    return '<span class="punkte">' + punkte.map((l) =>
      '<i class="p-' + l.auswertung.status + '" title="' + esc(zeitpunkt(l.eingang) + ": " + STATUS_TEXT[l.auswertung.status]) + '"></i>').join("") + "</span>";
  }

  function karte(l, vogt) {
    const a = l.auswertung;
    let aktionen = "";
    if (vogt) {
      if (a.status === STATUS.ZU_VOGT) aktionen += '<button data-aktion="BEI_VOGT" data-id="' + esc(l.id) + '">Bei Vogt abgegeben</button>';
      aktionen += '<button class="ok" data-aktion="OK" data-id="' + esc(l.id) + '">Messung i.O.</button>';
      aktionen += '<button class="nok" data-aktion="NOK" data-id="' + esc(l.id) + '">Ausserhalb Toleranz</button>';
    } else {
      aktionen += '<button class="nok" data-aktion="NOK" data-id="' + esc(l.id) + '">Abweichung melden</button>';
    }
    const info = (vogt ? a.grund : a.regelPhase) + " · geliefert " + zeitpunkt(l.eingang) + (a.beiVogtSeit ? " · bei Vogt seit " + zeitpunkt(a.beiVogtSeit) : "");
    return '<div class="karte" data-schluessel="' + esc(a.schluessel) + '">' +
      '<div class="titel">' + esc(l.artikel) + " · " + esc(l.bezeichnung) + "</div><div>" + badge(a.status) + "</div>" +
      '<div class="info">' + esc(l.lieferant) + (l.bestellung ? " · Best. " + esc(l.bestellung) : "") + "<br>" + esc(info) + "</div>" +
      '<div class="aktionen">' + aktionen + "</div></div>";
  }

  function rueckmeldungText(r) {
    if (!r) return "";
    const teile = [];
    if (r.massabweichung) teile.push("Massabweichung " + r.massabweichung);
    if (r.stueckzahl) teile.push(r.stueckzahl + " Stk.");
    if (r.bestellnummer) teile.push("Best. " + r.bestellnummer);
    if (r.bemerkung) teile.push(r.bemerkung);
    const wer = [r.erfasstVon, r.datum ? zeitpunkt(r.datum) : ""].filter(Boolean).join(", ");
    return teile.join(" · ") + (wer ? " (" + wer + ")" : "");
  }

  function detail(schluessel) {
    offenerArtikel = schluessel;
    const a = auswertung.artikel[schluessel];
    const liste = auswertung.lieferungen.filter((l) => l.auswertung.schluessel === schluessel);
    $("dialog-inhalt").innerHTML =
      "<h2>" + esc(a.artikel) + " · " + esc(a.bezeichnung) + "</h2>" +
      '<div class="stand">' + esc(a.lieferant) + (a.lieferantId ? " (" + esc(a.lieferantId) + ")" : "") + " · nächste Lieferung: <b>" + esc(a.naechsteLieferung) + "</b></div>" +
      '<ul class="verlauf">' + liste.map((l, i) => {
        const r = l.auswertung.rueckmeldung;
        let knoepfe = "";
        if (OFFEN.includes(l.auswertung.status)) knoepfe +=  ' <button class="klein" data-aktion="NICHT_GEMESSEN" data-id="' + esc(l.id) + '">Ging nicht zu Vogt</button>';
        if (r || l.auswertung.beiVogtSeit) knoepfe +=  ' <button class="klein" data-aktion="ZURUECKSETZEN" data-id="' + esc(l.id) + '">Rückmeldung löschen</button>';
        return '<li class="' + (i === 0 ? "aktiv" : "") + '">' + zeitpunkt(l.eingang) + (l.bestellung ? " · Best. " + esc(l.bestellung) : "") + " " +
          '<span class="badge s-phase">' + esc(l.auswertung.regelPhase) + "</span> " + badge(l.auswertung.status) + knoepfe +
          (r && rueckmeldungText(r) ? '<br><small class="stand">' + esc(rueckmeldungText(r)) + "</small>" : "") +
          "</li>";
      }).join("") + "</ul>" +
      '<label class="schalter"><input type="checkbox" id="f-ausgenommen"' + (a.ausgenommen ? " checked" : "") + "> Artikel ist nicht messpflichtig (geht nie zu Vogt)</label>" +
      '<div class="zeile"><button data-schliessen>Schliessen</button></div>';
    $("f-ausgenommen").onchange = (e) => senden("api/artikel", { schluessel, ausgenommen: e.target.checked, erfasstVon: name() });
    if (!$("dialog").open) $("dialog").showModal();
  }

  function rueckmeldungDialog(id, aktion) {
    const l = auswertung.lieferungen.find((x) => x.id === id);
    offenerArtikel = null;
    const nok = aktion === "NOK";
    $("dialog-inhalt").innerHTML =
      "<h2>" + (nok ? "Abweichung erfassen" : "Messung i.O. bestätigen") + "</h2>" +
      '<div class="stand">' + esc(l.artikel) + " " + esc(l.bezeichnung) + " · " + esc(l.lieferant) + " · geliefert " + zeitpunkt(l.eingang) + "</div>" +
      (nok ? '<div class="fehler">Der Artikel verliert seinen Freipass: die nächsten 3 Lieferungen müssen zu Vogt.</div>' +
        '<div class="felder"><label>Massabweichung<input type="text" id="f-mass" placeholder="z.B. 0.004mm"></label>' +
        '<label>Stückzahl n.i.O.<input type="text" id="f-stueck" placeholder="z.B. 2 von 15"></label>' +
        '<label>Bestellnummer<input type="text" id="f-best" value="' + esc(l.bestellung || "") + '"></label></div>' : "") +
      '<label>Bemerkung<textarea id="f-bemerkung" rows="2"></textarea></label>' +
      '<label>Erfasst von<input type="text" id="f-name" value="' + esc(name()) + '"></label>' +
      '<div class="zeile"><button data-schliessen>Abbrechen</button><button class="primaer" id="f-speichern">Speichern</button></div>';
    if (!$("dialog").open) $("dialog").showModal();
    $("f-speichern").onclick = async () => {
      const wert = (feld) => ($(feld) ? $(feld).value.trim() : "");
      nameMerken(wert("f-name"));
      $("dialog").close();
      await senden("api/rueckmeldung", {
        id, aktion, bemerkung: wert("f-bemerkung"), erfasstVon: wert("f-name"),
        massabweichung: wert("f-mass"), stueckzahl: wert("f-stueck"), bestellnummer: wert("f-best"),
      });
    };
  }

  document.addEventListener("click", (e) => {
    const knopf = e.target.closest("button");
    if (knopf && knopf.dataset.filter) { e.preventDefault(); filter = knopf.dataset.filter; zeichnen(); return; }
    if (knopf && knopf.hasAttribute("data-schliessen")) { $("dialog").close(); return; }
    if (knopf && knopf.dataset.aktion) {
      e.stopPropagation();
      const { id, aktion } = knopf.dataset;
      if (aktion === "OK" || aktion === "NOK") rueckmeldungDialog(id, aktion);
      else if (aktion === "ZURUECKSETZEN") { if (confirm("Rückmeldung wirklich löschen?")) senden("api/rueckmeldung", { id, aktion }); }
      else senden("api/rueckmeldung", { id, aktion, erfasstVon: name() });
      return;
    }
    const zeile = e.target.closest("[data-schluessel]");
    if (zeile) detail(zeile.dataset.schluessel);
  });
  $("dialog").addEventListener("close", () => { offenerArtikel = null; });
  $("suche").addEventListener("input", (e) => { suche = e.target.value.trim(); zeichnen(); });

  laden();
  setInterval(() => { if (!$("dialog").open) laden(); }, AKTUALISIEREN_MS);
})();

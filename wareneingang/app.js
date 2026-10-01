(function () {
  const { auswerten, STATUS } = window.Regeln;
  const AKTUALISIEREN_MS = 60 * 1000;

  const STATUS_TEXT = {
    ZU_VOGT: "Zu Vogt bringen",
    BEI_VOGT: "Bei Vogt · Ergebnis offen",
    OK: "Gemessen i.O.",
    NOK: "Abweichung",
    FREIPASS: "Freipass",
    FREIPASS_NOK: "Freipass · Abweichung gemeldet",
  };

  let daten = null;
  let auswertung = { lieferungen: [], artikel: {} };
  let filter = "ALLE";
  let suche = "";
  let offenerArtikel = null;

  const $ = (id) => document.getElementById(id);

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function datum(iso) {
    if (!iso) return "";
    const [j, m, t] = iso.slice(0, 10).split("-");
    return t + "." + m + "." + j;
  }
  function zeit(iso) {
    return new Date(iso).toLocaleString("de-CH", { dateStyle: "short", timeStyle: "short" });
  }
  function tageSeit(iso) {
    return Math.floor((Date.now() - new Date(iso.slice(0, 10)).getTime()) / 86400000);
  }
  function badge(status) {
    return '<span class="badge s-' + status + '">' + esc(STATUS_TEXT[status]) + "</span>";
  }
  function passt(l) {
    if (!suche) return true;
    const s = suche.toLowerCase();
    return [l.id, l.bestellung, l.artikel, l.bezeichnung, l.lieferant, l.lieferschein].some((f) => String(f || "").toLowerCase().includes(s));
  }

  async function laden() {
    try {
      const res = await fetch("api/daten", { cache: "no-store" });
      if (!res.ok) throw new Error("Server antwortet mit " + res.status);
      daten = await res.json();
      auswertung = auswerten(daten.lieferungen, daten.rueckmeldungen);
      $("fehler").hidden = true;
      $("stand").textContent = "ERP-Stand " + zeit(daten.erpStand) + " · aktualisiert " + zeit(new Date().toISOString()) + " · " + daten.lieferungen.length + " Lieferungen";
      zeichnen();
    } catch (e) {
      $("fehler").hidden = false;
      $("fehler").textContent = "Daten konnten nicht geladen werden (" + e.message + "). Läuft der Server (node server.js)?";
    }
  }

  async function melden(id, aktion, extra) {
    const res = await fetch("api/rueckmeldung", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(Object.assign({ id, aktion }, extra)),
    });
    if (!res.ok) { alert("Speichern fehlgeschlagen: " + (await res.text())); return; }
    await laden();
  }

  function zeichnen() {
    const alle = auswertung.lieferungen;
    const sichtbar = alle.filter(passt);
    const offen = alle.filter((l) => l.auswertung.status === STATUS.ZU_VOGT || l.auswertung.status === STATUS.BEI_VOGT);
    const freipass30 = alle.filter((l) => l.auswertung.status === STATUS.FREIPASS && tageSeit(l.eingang) <= 30);
    const abw90 = alle.filter((l) => (l.auswertung.status === STATUS.NOK || l.auswertung.status === STATUS.FREIPASS_NOK) && tageSeit(l.eingang) <= 90);

    $("kpis").innerHTML = [
      ["var(--rot)", offen.filter((l) => l.auswertung.status === STATUS.ZU_VOGT).length, "zu Vogt bringen"],
      ["var(--gelb)", offen.filter((l) => l.auswertung.status === STATUS.BEI_VOGT).length, "bei Vogt, Ergebnis offen"],
      ["var(--blau)", freipass30.length, "Freipässe (30 Tage)"],
      ["var(--rot)", abw90.length, "Abweichungen (90 Tage)"],
    ].map(([farbe, zahl, text]) => '<div class="kpi" style="--farbe:' + farbe + '"><b>' + zahl + "</b><span>" + text + "</span></div>").join("");

    const offenSichtbar = offen.filter(passt).sort((a, b) => (a.eingang < b.eingang ? -1 : 1));
    $("liste-vogt").innerHTML = offenSichtbar.length
      ? offenSichtbar.map((l) => karte(l, true)).join("")
      : '<div class="leer">Keine offenen Messungen – alles erledigt.</div>';

    const fpSichtbar = freipass30.filter(passt);
    $("liste-freipass").innerHTML = fpSichtbar.length
      ? fpSichtbar.map((l) => karte(l, false)).join("")
      : '<div class="leer">Keine Freipass-Lieferungen in den letzten 30 Tagen.</div>';

    const filterListe = [["ALLE", "Alle"], ["OFFEN", "Offen"], ["FREIPASS", "Freipass"], ["OK", "i.O."], ["ABW", "Abweichung"]];
    $("filter").innerHTML = filterListe.map(([k, t]) => '<button class="chip" data-filter="' + k + '" aria-pressed="' + (filter === k) + '">' + t + "</button>").join("");

    const tab = sichtbar.filter((l) => {
      const s = l.auswertung.status;
      if (filter === "OFFEN") return s === STATUS.ZU_VOGT || s === STATUS.BEI_VOGT;
      if (filter === "FREIPASS") return s === STATUS.FREIPASS;
      if (filter === "OK") return s === STATUS.OK;
      if (filter === "ABW") return s === STATUS.NOK || s === STATUS.FREIPASS_NOK;
      return true;
    });
    $("tabelle").innerHTML = tab.length ? tab.map((l) =>
      '<tr data-schluessel="' + esc(l.auswertung.schluessel) + '">' +
      "<td>" + datum(l.eingang) + "</td><td>" + esc(l.id) + "</td><td>" + esc(l.bestellung) + (l.position ? "/" + esc(l.position) : "") + "</td>" +
      "<td><b>" + esc(l.artikel) + "</b><br><small>" + esc(l.bezeichnung) + "</small></td><td>" + esc(l.lieferant) + "</td>" +
      '<td><span class="badge s-phase">' + esc(l.auswertung.regelPhase) + "</span></td><td>" + badge(l.auswertung.status) + "</td></tr>"
    ).join("") : '<tr><td colspan="7" class="leer">Keine Lieferungen gefunden.</td></tr>';

    const artikel = Object.entries(auswertung.artikel)
      .filter(([, a]) => passt(a))
      .sort(([, a], [, b]) => (b.naechsteMussZuVogt - a.naechsteMussZuVogt) || a.artikel.localeCompare(b.artikel));
    $("artikel").innerHTML = artikel.map(([k, a]) =>
      '<tr data-schluessel="' + esc(k) + '"><td><b>' + esc(a.artikel) + "</b></td><td>" + esc(a.bezeichnung) + "</td><td>" + esc(a.lieferant) + "</td><td>" + a.anzahl + "</td>" +
      '<td><span class="badge ' + (a.naechsteMussZuVogt ? "s-ZU_VOGT" : "s-FREIPASS") + '">' + esc(a.naechsteLieferung) + "</span></td></tr>"
    ).join("");

    if (offenerArtikel && $("dialog").open) detail(offenerArtikel);
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
    const info = vogt
      ? a.grund + " · Eingang " + datum(l.eingang) + " (vor " + tageSeit(l.eingang) + " Tagen)" + (a.beiVogtSeit ? " · bei Vogt seit " + datum(a.beiVogtSeit) : "")
      : a.regelPhase + " · Eingang " + datum(l.eingang);
    return '<div class="karte" data-schluessel="' + esc(a.schluessel) + '">' +
      '<div class="titel">Best. ' + esc(l.bestellung) + " · " + esc(l.artikel) + " " + esc(l.bezeichnung) + "</div><div>" + badge(a.status) + "</div>" +
      '<div class="info">' + esc(l.lieferant) + " · " + esc(l.id) + (l.menge ? " · " + esc(l.menge) + " Stk." : "") + "<br>" + esc(info) + "</div>" +
      '<div class="aktionen">' + aktionen + "</div></div>";
  }

  function detail(schluessel) {
    offenerArtikel = schluessel;
    const a = auswertung.artikel[schluessel];
    const liste = auswertung.lieferungen.filter((l) => l.auswertung.schluessel === schluessel);
    $("dialog-inhalt").innerHTML =
      "<h2>" + esc(a.artikel) + " · " + esc(a.bezeichnung) + "</h2>" +
      '<div class="stand">' + esc(a.lieferant) + " · nächste Lieferung: <b>" + esc(a.naechsteLieferung) + "</b></div>" +
      '<ul class="verlauf">' + liste.map((l, i) => {
        const r = l.auswertung.rueckmeldung;
        const knoepfe = l.auswertung.rueckmeldung || l.auswertung.beiVogtSeit
          ? ' <button data-aktion="ZURUECKSETZEN" data-id="' + esc(l.id) + '">Rückmeldung löschen</button>'
          : "";
        return '<li class="' + (i === 0 ? "aktiv" : "") + '">' + datum(l.eingang) + " · " + esc(l.id) + " · Best. " + esc(l.bestellung) + " " +
          '<span class="badge s-phase">' + esc(l.auswertung.regelPhase) + "</span> " + badge(l.auswertung.status) + knoepfe +
          (r ? '<br><small class="stand">' + (r.ergebnis === "NOK" ? "Abweichung" : "i.O.") + " erfasst " + zeit(r.datum) + (r.erfasstVon ? " von " + esc(r.erfasstVon) : "") + (r.bemerkung ? " · " + esc(r.bemerkung) : "") + "</small>" : "") +
          "</li>";
      }).join("") + "</ul>" +
      '<div class="zeile"><button data-schliessen>Schliessen</button></div>';
    if (!$("dialog").open) $("dialog").showModal();
  }

  function rueckmeldungDialog(id, aktion) {
    const l = auswertung.lieferungen.find((x) => x.id === id);
    offenerArtikel = null;
    let name = "";
    try { name = localStorage.getItem("we-name") || ""; } catch (e) {}
    const nok = aktion === "NOK";
    $("dialog-inhalt").innerHTML =
      "<h2>" + (nok ? "Abweichung erfassen" : "Messung i.O. bestätigen") + "</h2>" +
      '<div class="stand">' + esc(l.id) + " · Best. " + esc(l.bestellung) + " · " + esc(l.artikel) + " " + esc(l.bezeichnung) + " · " + esc(l.lieferant) + "</div>" +
      (nok ? '<div class="fehler">Der Artikel verliert seinen Freipass: die nächsten 3 Lieferungen müssen zu Vogt.</div>' : "") +
      '<label>Bemerkung' + (nok ? " (welches Mass, wie viel?)" : "") + '<textarea id="f-bemerkung" rows="3"></textarea></label>' +
      '<label>Erfasst von<input type="text" id="f-name" value="' + esc(name) + '"></label>' +
      '<div class="zeile"><button data-schliessen>Abbrechen</button><button class="primaer" id="f-speichern">Speichern</button></div>';
    if (!$("dialog").open) $("dialog").showModal();
    $("f-speichern").onclick = async () => {
      const erfasstVon = $("f-name").value.trim();
      try { localStorage.setItem("we-name", erfasstVon); } catch (e) {}
      $("dialog").close();
      await melden(id, aktion, { bemerkung: $("f-bemerkung").value.trim(), erfasstVon });
    };
  }

  document.addEventListener("click", (e) => {
    const knopf = e.target.closest("button");
    if (knopf && knopf.dataset.filter) { filter = knopf.dataset.filter; zeichnen(); return; }
    if (knopf && knopf.hasAttribute("data-schliessen")) { $("dialog").close(); return; }
    if (knopf && knopf.dataset.aktion) {
      e.stopPropagation();
      const { id, aktion } = knopf.dataset;
      if (aktion === "OK" || aktion === "NOK") rueckmeldungDialog(id, aktion);
      else if (aktion === "ZURUECKSETZEN") { if (confirm("Rückmeldung für " + id + " wirklich löschen?")) melden(id, aktion); }
      else melden(id, aktion);
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

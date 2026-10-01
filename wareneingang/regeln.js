// Freipass-Regeln für Stichprobenmessungen im Wareneingang.
//
// Ablauf pro Artikel + Lieferant:
//   1. Qualifizierung: jede Lieferung geht zu Vogt, bis 3 Messungen in Folge i.O. sind.
//   2. 3 Lieferungen mit Freipass (keine Messung).
//   3. Die 4. Lieferung geht zur Kontrollmessung zu Vogt.
//   4. Kontrollmessung i.O. → nochmals 3 Lieferungen mit Freipass.
//   5. Danach Requalifizierung (Messung bei Vogt). i.O. → wieder ab Schritt 2.
//   Jede Abweichung (bei Kontrolle, Requalifizierung oder gemeldet bei einer
//   Freipass-Lieferung) → zurück zu Schritt 1.
//
// Eine freiwillige Messung i.O. während der Freipass-Phase verbraucht keinen Freipass.
// Solange das Ergebnis einer Messung aussteht, gehen auch die folgenden Lieferungen
// desselben Artikels sicherheitshalber zu Vogt. Sobald das Ergebnis erfasst ist,
// wird alles neu berechnet.

(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Regeln = factory();
})(typeof self !== "undefined" ? self : this, function () {
  const QUALI_ANZAHL = 3;
  const FREIPASS_ANZAHL = 3;

  const PHASE = {
    QUALI: "QUALIFIZIERUNG",
    FREIPASS_1: "FREIPASS_1",
    KONTROLLE: "KONTROLLE",
    FREIPASS_2: "FREIPASS_2",
    REQUALI: "REQUALIFIZIERUNG",
  };

  const STATUS = {
    ZU_VOGT: "ZU_VOGT",       // muss zu Vogt, noch nicht abgegeben
    BEI_VOGT: "BEI_VOGT",     // bei Vogt abgegeben, Ergebnis ausstehend
    OK: "OK",                 // gemessen, in Toleranz
    NOK: "NOK",               // gemessen, ausserhalb Toleranz
    FREIPASS: "FREIPASS",     // keine Messung nötig
    FREIPASS_NOK: "FREIPASS_NOK", // Freipass, aber Abweichung nachträglich gemeldet
    NICHT_GEMESSEN: "NICHT_GEMESSEN", // hätte zu Vogt gemusst, ging aber nicht
    AUSGENOMMEN: "AUSGENOMMEN", // Artikel ist nicht messpflichtig
  };

  function artikelSchluessel(l) {
    return (l.artikel || "").trim() + "|" + (l.lieferantId || l.lieferant || "").trim();
  }

  function vergleiche(a, b) {
    if (a.eingang !== b.eingang) return a.eingang < b.eingang ? -1 : 1;
    return String(a.id).localeCompare(String(b.id), "de", { numeric: true });
  }

  function phaseText(zustand) {
    switch (zustand.phase) {
      case PHASE.QUALI: return "Qualifizierung " + (zustand.zaehler + 1) + "/" + QUALI_ANZAHL;
      case PHASE.FREIPASS_1: return "Freipass " + (zustand.zaehler + 1) + "/" + FREIPASS_ANZAHL;
      case PHASE.KONTROLLE: return "Kontrollmessung";
      case PHASE.FREIPASS_2: return "Freipass " + (zustand.zaehler + 1) + "/" + FREIPASS_ANZAHL + " (2. Runde)";
      case PHASE.REQUALI: return "Requalifizierung";
    }
    return zustand.phase;
  }

  function istMessphase(phase) {
    return phase === PHASE.QUALI || phase === PHASE.KONTROLLE || phase === PHASE.REQUALI;
  }

  function nachAbweichung() {
    return { phase: PHASE.QUALI, zaehler: 0 };
  }

  function nachOk(zustand) {
    switch (zustand.phase) {
      case PHASE.QUALI:
        return zustand.zaehler + 1 >= QUALI_ANZAHL
          ? { phase: PHASE.FREIPASS_1, zaehler: 0 }
          : { phase: PHASE.QUALI, zaehler: zustand.zaehler + 1 };
      case PHASE.KONTROLLE:
        return { phase: PHASE.FREIPASS_2, zaehler: 0 };
      case PHASE.REQUALI:
        return { phase: PHASE.FREIPASS_1, zaehler: 0 };
      default:
        // Zusätzliche Messung während der Freipass-Phase: Freipass bleibt, es wird kein Platz verbraucht
        return zustand;
    }
  }

  function nachFreipass(zustand) {
    if (zustand.zaehler + 1 < FREIPASS_ANZAHL) return { phase: zustand.phase, zaehler: zustand.zaehler + 1 };
    return zustand.phase === PHASE.FREIPASS_1
      ? { phase: PHASE.KONTROLLE, zaehler: 0 }
      : { phase: PHASE.REQUALI, zaehler: 0 };
  }

  // Berechnet für alle Lieferungen den Status.
  // lieferungen: [{ id, eingang: "YYYY-MM" oder "YYYY-MM-DD", artikel, lieferantId, lieferant, ... }]
  // rueckmeldungen: { [id]: { ergebnis: "OK" | "NOK" | "NICHT_GEMESSEN", beiVogtSeit, ... } }
  // einstellungen: { [artikelSchluessel]: { ausgenommen: true } }
  // Rückgabe: { lieferungen: [...mit .auswertung], artikel: { [schluessel]: Zusammenfassung } }
  function auswerten(lieferungen, rueckmeldungen, einstellungen) {
    rueckmeldungen = rueckmeldungen || {};
    einstellungen = einstellungen || {};
    const gruppen = new Map();
    for (const l of lieferungen) {
      const k = artikelSchluessel(l);
      if (!gruppen.has(k)) gruppen.set(k, []);
      gruppen.get(k).push(l);
    }

    const ergebnis = [];
    const artikel = {};

    for (const [schluessel, liste] of gruppen) {
      liste.sort(vergleiche);
      const ausgenommen = !!(einstellungen[schluessel] && einstellungen[schluessel].ausgenommen);
      let zustand = { phase: PHASE.QUALI, zaehler: 0 };
      let offeneMessung = null;

      for (const l of liste) {
        const r = rueckmeldungen[l.id] || {};
        const regelPhase = ausgenommen ? "Nicht messpflichtig" : phaseText(zustand);
        let mussZuVogt = !ausgenommen && istMessphase(zustand.phase);
        let grund = regelPhase;
        if (!mussZuVogt && !ausgenommen && offeneMessung) {
          mussZuVogt = true;
          grund = "Ergebnis der Lieferung " + offeneMessung + " ausstehend";
        }

        let status;
        if (ausgenommen && r.ergebnis !== "NOK" && r.ergebnis !== "OK") {
          status = STATUS.AUSGENOMMEN;
        } else if (r.ergebnis === "NOK") {
          status = mussZuVogt ? STATUS.NOK : STATUS.FREIPASS_NOK;
          zustand = nachAbweichung();
        } else if (r.ergebnis === "OK") {
          status = STATUS.OK;
          zustand = nachOk(zustand);
        } else if (r.ergebnis === "NICHT_GEMESSEN") {
          // Freipass-Platz wird verbraucht; in einer Messphase zählt es nicht als Messung
          status = mussZuVogt ? STATUS.NICHT_GEMESSEN : STATUS.FREIPASS;
          if (!istMessphase(zustand.phase)) zustand = nachFreipass(zustand);
        } else if (mussZuVogt) {
          // Ergebnis offen: für die Zählung wird i.O. angenommen, aber bis das
          // Ergebnis da ist, gehen alle weiteren Lieferungen ebenfalls zu Vogt.
          status = r.beiVogtSeit ? STATUS.BEI_VOGT : STATUS.ZU_VOGT;
          offeneMessung = l.eingang;
          zustand = nachOk(zustand);
        } else {
          status = STATUS.FREIPASS;
          zustand = nachFreipass(zustand);
        }

        ergebnis.push(Object.assign({}, l, {
          auswertung: { schluessel, regelPhase, mussZuVogt, grund, status, rueckmeldung: r.ergebnis ? r : null, beiVogtSeit: r.beiVogtSeit || null },
        }));
      }

      const letzte = liste[liste.length - 1];
      artikel[schluessel] = {
        artikel: letzte.artikel,
        bezeichnung: letzte.bezeichnung,
        lieferantId: letzte.lieferantId,
        lieferant: letzte.lieferant,
        anzahl: liste.length,
        letzteLieferung: letzte.eingang,
        ausgenommen,
        naechsteLieferung: ausgenommen
          ? "Nicht messpflichtig"
          : istMessphase(zustand.phase)
          ? "Zu Vogt – " + phaseText(zustand)
          : offeneMessung
            ? "Zu Vogt (Ergebnis Lieferung " + offeneMessung + " ausstehend)"
            : phaseText(zustand),
        naechsteMussZuVogt: !ausgenommen && (!!offeneMessung || istMessphase(zustand.phase)),
      };
    }

    ergebnis.sort((a, b) => -vergleiche(a, b));
    return { lieferungen: ergebnis, artikel };
  }

  return { auswerten, artikelSchluessel, PHASE, STATUS, QUALI_ANZAHL, FREIPASS_ANZAHL };
});

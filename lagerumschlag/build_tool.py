# -*- coding: utf-8 -*-
"""
build_tool.py - erzeugt / aktualisiert das Excel-Tool «Lagerumschlag_Tool.xlsx».

Aufruf:   python build_tool.py            (Datenbericht + Tool bauen; Excel rechnet beim Öffnen)
          python build_tool.py --bericht  (nur Datenbericht Schritt 1)
          python build_tool.py --berechnen (Werte mit LibreOffice berechnen und speichern; braucht RECALC_SCRIPT)
          python build_tool.py --pruefen  (zusätzlich Kopie mit LibreOffice berechnen, Formelfehler prüfen;
                                           braucht RECALC_SCRIPT=Pfad/zu/recalc.py)
Benötigt: Python 3, pandas, numpy, openpyxl.

Ablauf:
 1. Exporte lesen (Kennzahlen-Auszug + Artikelstamm mit Gebinden).
 2. Falls TOOL_FILE schon existiert: alle gelben Eingaben übernehmen
    (Parameter, Cockpit-Regler, Artikel-Eingaben je Artikelnummer, Neue Artikel, Gebinde-Kategorien).
    Die alte Datei wird vorher als Sicherung kopiert.
 3. Tool neu schreiben (nur Formeln, keine Makros) und mit LibreOffice berechnen.
"""
import os, re, sys, shutil, datetime, subprocess, json, glob, gzip, math
import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment

# ============================================================================
# Dateinamen (bei neuen Exporten hier anpassen)
# ============================================================================
HERE = os.path.dirname(os.path.abspath(__file__))
KENNZAHLEN_FILE = os.path.join(HERE, "input", "Kennzahlen_Lager_Auszug.xlsx")
ARTIKELSTAMM_FILE = os.path.join(HERE, "input", "Artikelstamm_mit_Verbrauch_und_Gebinde_Umschlag_2_20260520.xlsx")
# neuere Gebindeliste als CSV (Export «Artikelstamm Umschlag 2»); hat Vorrang vor dem Blatt in ARTIKELSTAMM_FILE
ARTIKELSTAMM_CSV = os.path.join(HERE, "input", "Artikelstamm_Umschlag_2_20260623.csv")
TOOL_FILE = os.path.join(HERE, "Lagerumschlag_Tool.xlsx")      # bestehendes Tool -> Eingaben übernehmen
OUTPUT_FILE = os.path.join(HERE, "Lagerumschlag_Tool.xlsx")
# Liftberichte (Modula «Artikelbestand für Maschine», .prnx) – alle Dateien im Ordner input
# Werte aus der Lagerplanung (Tabelle am Ende der Gebindeliste): Gebinde je Tablar und Höhe im Lift (cm)
PLAN_GPT = {"S21": 280, "S22": 140, "S32": 104, "S33": 52, "S41": 52, "S51": 26, "S52": 26, "S61": 12, "S62": 12, "S63": 12,
            "S71": 6, "S72": 6, "S73": 6, "S81": 4, "S82": 4, "S83": 4}
PLAN_HCM = {"S21": 12, "S22": 12, "S32": 12, "S33": 12, "S41": 12, "S51": 12, "S52": 21, "S61": 11, "S62": 21, "S63": 31,
            "S71": 20, "S72": 40, "S73": 60, "S81": 20, "S82": 40, "S83": 60}
LIFT_REPORTS = sorted(glob.glob(os.path.join(HERE, "input", "*.prnx")))
MASCHINE_ZU_LIFT = {"1": 1, "2": 2, "3": 3}
ERGAENZEN_HL = ("KTL", "PAL")   # Artikel dieser Hauptlager mit Bestand im Artikelstamm, aber nicht im Kennzahlen-Auszug -> ergänzen                  # Maschinen-Nr. im Bericht -> Lift
N_NEW = 100                                                  # reservierte Zeilen für neue Artikel
RECALC_SCRIPT = os.environ.get("RECALC_SCRIPT", "")          # optional für --pruefen

# ============================================================================
# Stil
# ============================================================================
FONT = "Arial"
YELLOW = PatternFill("solid", fgColor="FFF2A6")
GREY = PatternFill("solid", fgColor="D9D9D9")
HEAD = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E1F2")
GREEN_F = PatternFill("solid", fgColor="C6EFCE")
ORANGE_F = PatternFill("solid", fgColor="FFD966")
RED_F = PatternFill("solid", fgColor="F4B084")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def fnt(bold=False, size=10, color="000000", italic=False):
    return Font(name=FONT, bold=bold, size=size, color=color, italic=italic)


def put(ws, ref, value, bold=False, size=10, fill=None, fmt=None, color="000000", align=None, wrap=False, italic=False, border=False):
    c = ws[ref]
    c.value = value
    c.font = fnt(bold, size, color, italic)
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        c.border = BOX
    return c


# ============================================================================
# Hilfsfunktionen Daten
# ============================================================================
def norm_nr(x):
    """Artikelnummer vereinheitlichen: 1029 / '1029' / '001029' -> '001029'."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if s.isdigit():
        s = s.zfill(6)
    return s or None


def num(x, default=0.0):
    try:
        if x is None or x == "" or (isinstance(x, float) and np.isnan(x)):
            return default
        return float(x)
    except (TypeError, ValueError):
        return default


def txt(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    return str(x).strip()


GEB_RE = re.compile(r"[SP]\d\d")


def parse_liftbericht(path):
    """Modula-Bericht RPT_ART_GIAC_MACCHINA_MOD (.prnx = gzip-XML) -> (Maschine, [Zeilen])."""
    raw = open(path, "rb").read()
    x = gzip.decompress(raw) if raw[:2] == b"\x1f\x8b" else raw
    root = ET.fromstring(x.decode("utf-8-sig").encode("utf-8"))
    labels = [e for e in root.iter() if e.get("BrickType") == "Label"]
    texts = [e.get("Text") for e in labels]
    maschine = ""
    for i, t in enumerate(texts):
        if t == "Maschine" and i + 1 < len(texts) and re.fullmatch(r"\d+", texts[i + 1] or ""):
            maschine = texts[i + 1]
            break
    names = {"Code": "code", "Beschreibung": "bez", "Gesamtkapazität": "kap", "Gesamtbestand": "bestand"}
    hdr = []  # (x, zeile, key)
    head_y = None
    for e in labels:
        if e.get("Text") in names:
            x, y = [float(v) for v in e.get("Rect").split(",")[:2]]
            head_y = y if head_y is None else min(head_y, y)
            hdr.append((x, y, names[e.get("Text")]))
    hdr = [(x, 0 if y - head_y < 40 else 1, k) for x, y, k in hdr]
    rows = []
    for pnl in root.iter():
        if pnl.get("BrickType") != "Panel" or not pnl.get("Rect"):
            continue
        px = float(pnl.get("Rect").split(",")[0])
        d = {}
        for e in pnl.iter():
            if e.get("BrickType") != "Label" or e is pnl:
                continue
            x, y = [float(v) for v in e.get("Rect").split(",")[:2]]
            line = 0 if y < 40 else 1
            for hx, hl, k in hdr:
                if hl == line and abs(px + x - hx) < 60:
                    d[k] = e.get("Text")
        if d.get("code") and d.get("kap") is not None:
            rows.append(dict(code=norm_nr(d["code"]), bez=txt(d.get("bez")), kap=num(d.get("kap")),
                             bestand=num(d.get("bestand"))))
    return maschine, rows


def load_data():
    # --- Kennzahlen: Basis Bestand -----------------------------------------
    raw = pd.read_excel(KENNZAHLEN_FILE, "Auswertung Basis Bestand", header=None)
    h4 = num(raw.iat[3, 7])
    hdr = 5  # Zeile 6
    b = raw.iloc[hdr + 1:].reset_index(drop=True)
    b = b[b[0].notna()]
    # Spalten nach Buchstaben (0-basiert): A0 B1 C2 D3 E4 F5 G6 H7 K10 O14 T19 U20 W22 Y24 Z25 AK36 AP41 AQ42
    # --- Verbrauch Pivot ----------------------------------------------------
    vraw = pd.read_excel(KENNZAHLEN_FILE, "Auswertung Verbrauch", header=None)
    period = txt(vraw.iat[1, 1])
    verb = {}
    for _, r in vraw.iloc[6:].iterrows():
        n = norm_nr(r[0])
        if not n or n.lower().startswith("gesamtergebnis"):
            continue
        verb[n] = max(0.0, num(r[2]))
    # --- MB -----------------------------------------------------------------
    mbraw = pd.read_excel(KENNZAHLEN_FILE, "Auswertung MB", header=1)
    mbmap = {}
    for _, r in mbraw.iterrows():
        n = norm_nr(r.iloc[0])
        if n:
            mbmap.setdefault(n, {})[txt(r.iloc[2])] = num(r.iloc[3])
    # --- Artikelstamm (Gebinde) --------------------------------------------
    if ARTIKELSTAMM_CSV and os.path.exists(ARTIKELSTAMM_CSV):
        s2 = pd.read_csv(ARTIKELSTAMM_CSV, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        s2 = s2.apply(lambda c: c.str.strip())
        for c in ("Preis-GLD-Akt", "Losgröße", "Wbz.", "Bestand") + tuple(x for x in s2.columns if str(x).startswith("Menge pro")):
            if c in s2.columns:
                s2[c] = pd.to_numeric(s2[c].str.replace(",", ""), errors="coerce")
        s2 = s2.replace("", np.nan)
    else:
        s2 = pd.read_excel(ARTIKELSTAMM_FILE, "Artikelstamm Umschlag 2")
    s1 = pd.read_excel(ARTIKELSTAMM_FILE, "Artikelstamm mit Verbrauch ")
    col_geb2 = [c for c in s2.columns if str(c).replace("\n", "").replace("-", "").lower().startswith("gebindekategorie")][0]
    col_mpg2 = [c for c in s2.columns if str(c).replace("\n", " ").lower().startswith("menge pro")][0]
    stamm = {}
    for _, r in s1.iterrows():
        n = norm_nr(r["Artikel-Nr."])
        g = txt(r.get("Gebindekategorie")).upper()
        if n and GEB_RE.fullmatch(g):
            stamm[n] = dict(geb=g, mpg=num(r.get("Menge pro Gebinde"), 0), bem="", hl="")
    for _, r in s2.iterrows():
        n = norm_nr(r["Artikel-Nr."])
        if not n:
            continue
        g = txt(r[col_geb2]).upper()
        bem = "; ".join(x for x in [txt(r.get("Bemerkung")), txt(r.get("Bemerkungen"))] if x)
        d = stamm.get(n, dict(geb="", mpg=0, bem="", hl=""))
        if GEB_RE.fullmatch(g):
            d["geb"], d["mpg"] = g, num(r[col_mpg2], 0)
        d["bem"] = bem
        d["hl"] = txt(r.get("Hauptlager"))
        d["liz"] = txt(r.get("Lager in Zukunft"))
        d["bestand"] = num(r.get("Bestand"))
        d.update(bez=txt(r.get("Bezeichnung")), preis=num(r.get("Preis-GLD-Akt")), lg=num(r.get("Losgröße")),
                 status=txt(r.get("Artikelstat.")), wbz=num(r.get("Wbz."), None))
        stamm[n] = d
    gk = pd.read_excel(ARTIKELSTAMM_FILE, "Gebinde-Kategorie")
    gebkat = []
    for _, r in gk.iterrows():
        c = txt(r.iloc[0]).upper()
        if GEB_RE.fullmatch(c):
            gebkat.append([c, txt(r.iloc[1]), num(r.iloc[2]), num(r.iloc[3]), num(r.iloc[4])])

    arts = []
    for i, r in b.iterrows():
        n = norm_nr(r[0])
        st = stamm.get(n, {})
        hl = txt(r[4]) or "(leer)"
        mb = r[36]
        if num(mb, None) is None:
            m = mbmap.get(n, {})
            mb = m.get(hl, max(m.values()) if m else 0)
        v = verb.get(n)
        if v is None:
            v = max(0.0, num(r[24]))
        arts.append(dict(
            nr=n, bez=txt(r[1]), status=txt(r[2]), hl=hl, abc=txt(r[14]), preis=num(r[5]),
            bm=num(r[6]), ja=num(r[19]), bwx=num(r[7]), verbx=v, avg=num(r[22], None), lu_ist=num(r[25], None),
            mb=num(mb), lg=num(r[3]), wbz=num(r[10], None), geb=st.get("geb") or "LOSE", mpg=st.get("mpg") or None,
            bem=st.get("bem", ""), liz=st.get("liz", ""), ursache=txt(r[41]), aktion=txt(r[42]), xrow=hdr + 2 + i,
        ))
    # --- Artikel mit Bestand im Artikelstamm, die im Kennzahlen-Auszug fehlen (nur KTL/PAL) ---
    in_k = {a["nr"] for a in arts}
    for n, st in stamm.items():
        if n in in_k or st.get("hl") not in ERGAENZEN_HL or num(st.get("bestand")) <= 0:
            continue
        m = mbmap.get(n, {})
        arts.append(dict(nr=n, bez=st.get("bez", ""), status=st.get("status", ""), hl=st["hl"], abc="", preis=st.get("preis", 0.0),
                         bm=st["bestand"], ja=0.0, bwx=0.0, verbx=verb.get(n, 0.0), avg=None, lu_ist=None,
                         mb=m.get(st["hl"], max(m.values()) if m else 0.0), lg=st.get("lg", 0.0), wbz=st.get("wbz"),
                         geb=st.get("geb") or "LOSE", mpg=st.get("mpg") or None,
                         bem="; ".join(v for v in ["nur im Artikelstamm (Bestand 23.06.2026)", st.get("bem", "")] if v),
                         liz=st.get("liz", ""), ursache="", aktion="", xrow=None, ergaenzt=1))
    # --- Liftberichte -------------------------------------------------------
    lift_rows = []
    for f in LIFT_REPORTS:
        m, rows = parse_liftbericht(f)
        lift = MASCHINE_ZU_LIFT.get(m)
        if not lift:
            print(f"Warnung: {os.path.basename(f)} – Maschine '{m}' keinem Lift zugeordnet, übersprungen")
            continue
        for x in rows:
            x["lift"] = lift
            lift_rows.append(x)
    byn = {a["nr"]: a for a in arts}
    for x in lift_rows:
        a = byn.get(x["code"])
        if a is None:  # im Lift, aber nicht im Bestandsauszug -> ergänzen
            st = stamm.get(x["code"], {})
            m = mbmap.get(x["code"], {})
            a = dict(nr=x["code"], bez=st.get("bez") or x["bez"], status=st.get("status", ""), hl=st.get("hl") or "(leer)",
                     abc="", preis=st.get("preis", 0.0), bm=x["bestand"], ja=0.0, bwx=0.0,
                     verbx=verb.get(x["code"], 0.0), avg=None, lu_ist=None, mb=max(m.values()) if m else 0.0,
                     lg=st.get("lg", 0.0), wbz=st.get("wbz"), geb=st.get("geb") or "LOSE", mpg=st.get("mpg") or None,
                     bem="; ".join(v for v in ["nur im Liftbericht (nicht im Kennzahlen-Auszug)", st.get("bem", "")] if v),
                     liz=st.get("liz", ""),
                     ursache="", aktion="", xrow=None)
            arts.append(a)
            byn[a["nr"]] = a
        a.update(lift_rep=x["lift"], kap_rep=x["kap"], best_rep=x["bestand"])
    return dict(arts=arts, h4=h4, period=period, gebkat=gebkat, verb=verb, stamm=stamm, lift_rows=lift_rows)


# ============================================================================
# Grobe Schätzung loser Artikel (ohne Gebinde) über Stichworte in der Bezeichnung
# ============================================================================
KLASSEN_DEFAULT = [  # Klasse, Schätz-Gebinde, Menge pro Gebinde, Beschreibung
    ("Klein", "S51", 20, "Schrauben, Muttern, Kabel, Ventile, Sensoren, Etiketten … (Eurobox 400×300, 20 Stk)"),
    ("Mittel", "S61", 4, "Panels, Gehäuse, Module, Schläuche, Sets, Service-Kits … (Eurobox 600×400, 4 Stk)"),
    ("Gross", "S81", 2, "Holzkisten, Rahmen, Rohre, Profile, Wagen, Kartons, Füllmaterial … (Trennblech 1200×800, 2 Stk)"),
]
KW_KLEIN_STARK = ("schraube", "mutter", "scheibe", "stopfen", "etikett", "beiblatt", "checkliste", "handbuch", "dichtung",
                  "o-ring", "nippel", "klemme", "hülse", "stift", "feder", "litze", "draht", "bürste", "schlüssel", "sicherung")
KW_GROSS = ("holzkiste", "kiste", "faltrahmen", "palette", "gestell", "rohr", "profil", "kranbahn", "wagen", "karton",
            "füllmaterial", "luftpolster", "rohmaterial", "innenbox", "seitenfaltenbeutel", "lagerkasten", "kompressor",
            "flasche", "behälter", "tank", "ausgleichsvolumen")
KW_MITTEL = ("panel", "gehäuse", "modul", "block", "schlauch", "set", "kit", "service", "verpackung", "vorwärmer",
             "zylinder", "strecke", "düse", "schauglas", "blech", "heiz", "pumpe", "motor", "steuerung", "sps")


def schaetzklasse(bez, hl):
    b = (bez or "").lower()
    if any(k in b for k in KW_KLEIN_STARK) and not any(k in b for k in ("panel", "modul", "kiste")):
        return "Klein"
    if any(k in b for k in KW_GROSS) and "rohrbürste" not in b:
        return "Gross"
    if any(k in b for k in KW_MITTEL):
        return "Mittel"
    return "Klein" if hl == "KTL" else "Mittel"


# ============================================================================
# Schritt 1: Datenbericht
# ============================================================================
# ============================================================================
# Ausgeglichene Lift-Zuteilung (Vorschlag, Basis Ist-Bestand)
# ============================================================================
NUR_AUSFAHRBAR = ("S7", "S8", "P")   # Trennbleche / Paletten nur in Lift 1 + 2 (ausfahrbar)
AUSGLEICH_LU = 3                      # Basis der ausgeglichenen Zuteilung: max(Ist, Ziel-Ø bei diesem LU) = Variante C


def mengen(a, lu=AUSGLEICH_LU):
    """(Ist, Ø Variante B, Variante C) – wie im Excel (Platz = Ø-Bestand)."""
    ist = max(0.0, a["ja"] + a["bm"]) if a.get("xrow") is not None else max(0.0, a["bm"])
    if a.get("best_rep") is not None:
        ist = a["best_rep"]
    mb, lg, v = a.get("mb") or 0, a.get("lg") or 0, a.get("verbx") or 0   # Verbrauch = 12 Monate (Zeitraum 2025)
    mx = math.ceil(mb + max(lg, 2 * (v / lu - mb)))
    avg = (mb + mx) / 2
    return ist, (ist, math.ceil(avg), math.ceil(max(avg, ist)))


def hoehen_mm(a, zuschlag=35):
    """Python-Nachbildung der Excel-Rechnung: Höhe im Lift (mm) je Szenario (Ist, B, C)."""
    ist, qs = mengen(a)
    geb, mpg = a["geb"], a.get("mpg") or 0
    if geb == "LOSE":
        kl = schaetzklasse(a["bez"], a["hl"])
        geb, mpg = {k: (g, m) for k, g, m, _ in KLASSEN_DEFAULT}[kl]
        mpg = max(mpg, ist, a.get("lg") or 0)
    if mpg <= 0:
        mpg = max(1.0, ist)
    gpt, hcm = PLAN_GPT.get(geb), PLAN_HCM.get(geb)
    if not gpt:
        return geb, (0.0, 0.0, 0.0)
    return geb, tuple(math.ceil(q / mpg) / gpt * (hcm * 10 + zuschlag) if q > 0 else 0.0 for q in qs)


def regel_lift(a):
    if a.get("lift_rep") is not None:
        return a["lift_rep"]
    z = dict(LIZ_DEFAULT).get(a.get("liz")) if a.get("liz") else None
    if z is None:
        z = {k: g for k, g, _ in PARAM_HL_DEFAULT}.get(a["hl"], "Nein")
    if z == "Nein":
        return None
    if z == "2+3":
        return 2 if "grosser Artikel" in (a.get("bem") or "") else dict(GEB_LIFT_DEFAULT).get(a["geb"], 2)
    return z


def ausgleichen(arts, kap=(11900, 11900, 11900), lift1_heute=0.85):
    """Greedy (grösste zuerst): jeder Artikel in den erlaubten Lift, der danach den tiefsten höchsten Füllgrad
    über die Szenarien Ist / B / C (Ziel-LU AUSGLEICH_LU) hat. Liftbericht-Artikel bleiben in Lift 1."""
    last = {l: [0.0, 0.0, 0.0] for l in (1, 2, 3)}
    res, todo = {}, []
    fach = [a for a in arts if a.get("lift_rep") == 1 and a.get("kap_rep")]
    skap = sum(a["kap_rep"] for a in fach) or 1
    for a in fach:   # Fächer aus dem Liftbericht: gemessene Belegung, wächst wenn Menge > Fach-Kapazität
        _, qs = mengen(a)
        for i, q in enumerate(qs):
            last[1][i] += lift1_heute * kap[0] * a["kap_rep"] / skap * max(1.0, q / a["kap_rep"])
        res[a["nr"]] = 1
    for a in arts:
        if a["nr"] in res or regel_lift(a) is None:
            continue
        geb, hs = hoehen_mm(a)
        erlaubt = (1, 2) if geb.startswith(NUR_AUSFAHRBAR) else (1, 2, 3)
        todo.append((hs, a["nr"], erlaubt, regel_lift(a)))
    for hs, nr, erlaubt, l in sorted(todo, key=lambda x: -max(x[0])):
        def score(x):   # zuerst Ist und Zielzustand B, dann Übergang C ausgleichen
            f = [(last[x][i] + hs[i]) / kap[x - 1] for i in range(3)]
            return (round(max(f[0], f[1]), 3), round(f[2], 3), x != l)
        best = min(erlaubt, key=score)
        for i in range(3):
            last[best][i] += hs[i]
        res[nr] = best
    return res, {l: [round(v / kap[l - 1] * 100) for v in last[l]] for l in last}


def bericht(D):
    a = pd.DataFrame(D["arts"])
    out = []
    p = out.append
    p("=== DATENBERICHT ===")
    p(f"Artikel im Bestandsauszug: {len(a)}  (eindeutige Nummern: {a.nr.nunique()})")
    p(f"Verbrauchszeitraum (Pivot-Filter 'Auswertung Verbrauch'!B2): {D['period']}")
    p(f"H4 Total Bestandswert lt. Export: {D['h4']:,.2f}")
    p(f"Summe Bestandswert alle Zeilen (inkl. negative): {a.bwx.sum():,.2f}")
    p(f"Summe Bestandswert nur positive Bestände: {a[a.bm > 0].bwx.sum():,.2f}")
    p(f"Negative Bestandsmenge (Spalte G): {(a.bm < 0).sum()} Artikel; = 0: {(a.bm == 0).sum()}")
    rec = (a.ja + a.bm)
    p(f"Jahresanfang 2025 + Bestandsmenge < 0: {(rec < 0).sum()} Artikel  -> G sieht nach Netto-Bewegung seit 01.01.2025 aus")
    p("")
    a["lose"] = a.geb == "LOSE"
    a["verb0"] = a.verbx <= 0
    g = a.groupby("hl").agg(Artikel=("nr", "size"), Bestand_pos=("bm", lambda s: (s > 0).sum()),
                            ohne_Verbrauch=("verb0", "sum"), Gebinde=("lose", lambda s: (~s).sum()), lose=("lose", "sum"))
    g["lose %"] = (g.lose / g.Artikel * 100).round(0)
    p(g.sort_values("Artikel", ascending=False).to_string())
    p("")
    def gpt(r):  # Gebinde je Tablar (4060 x 857 mm), beste Anordnung
        m = re.search(r"(\d{3,4})\s*x\s*(\d{3,4})", r[1] or "")
        l, b = (float(m.group(1)), float(m.group(2))) if m and not r[1].upper().startswith("BITO") else (r[2] * 10, r[3] * 10)
        return max((4060 // l) * (857 // b), (4060 // b) * (857 // l)) if l and b else 0
    gk = {r[0]: gpt(r) for r in D["gebkat"]}
    lift_of = lambda r: 1 if r.hl == "LIFT1" else (None if r.hl not in ("KTL", "PAL") else (
        2 if ("grosser Artikel" in r.bem or r.geb == "LOSE" or r.geb >= "S6") else 3))
    a["lift"] = a.apply(lift_of, axis=1)
    for label, q in [("Export G (neg.=0)", a.bm.clip(lower=0)), ("Jahresanfang+G", rec.clip(lower=0))]:
        mpg = a.mpg.fillna(0).where(a.mpg.fillna(0) > 0, q.clip(lower=1))  # fehlt: ganzer Bestand = 1 Gebinde
        g = a.geb.map(gk).fillna(0)
        tab = np.where(a.lose | (g == 0), 0, np.ceil(q / mpg) / g.replace(0, 1))
        p(f"Ist-Belegung nur Gebinde-Artikel (ohne Liftbericht), Bestand = {label}:")
        for l in (1, 2, 3):
            m = tab[a.lift == l].sum()
            p(f"   Lift {l}: {m:5.1f} Tablare von 50")
    return "\n".join(out)


# ============================================================================
# Bestehende Eingaben lesen
# ============================================================================
ART_INPUT_HEADERS = ["Schätzklasse lose", "Tablare lose (ganzer Bestand)", "Gebindekategorie neu", "Menge pro Gebinde neu",
                     "Lift manuell", "Ziel-LU Artikel", "MAX final", "Kommentar Diskussion"]


def read_old(path):
    if not os.path.exists(path):
        return None
    wf = load_workbook(path)
    wv = load_workbook(path, data_only=True)
    old = {}
    old["version"] = wf["Parameter"]["Z1"].value if "Parameter" in wf.sheetnames else None
    # Parameter: alle Zellen mit gelber Füllung
    for sh in ("Parameter", "Cockpit", "Einkauf", "DREIER", "Ergebnis je Artikel"):
        if sh in wf.sheetnames:
            d = {}
            for row in wf[sh].iter_rows():
                for c in row:
                    if c.fill and c.fill.fgColor and c.fill.fgColor.rgb in ("00FFF2A6", "FFFFF2A6") and c.value is not None \
                            and not (isinstance(c.value, str) and c.value.startswith("=")):
                        d[c.coordinate] = c.value
            old[sh] = d
    # Artikel-Eingaben
    if "Artikel" in wf.sheetnames:
        ws, wsv = wf["Artikel"], wv["Artikel"]
        hdr = {str(c.value).strip(): c.column for c in ws[6] if c.value}
        cols = {h: hdr[h] for h in ART_INPUT_HEADERS if h in hdr}
        d = {}
        for r in range(7, ws.max_row + 1):
            nr = norm_nr(wsv.cell(r, 1).value)
            if not nr:
                continue
            vals = {}
            for h, c in cols.items():
                v = ws.cell(r, c).value
                if v is not None and not (isinstance(v, str) and v.startswith("=")):
                    vals[h] = v
            if vals:
                d[nr] = vals
        old["Artikel"] = d
    if "Neue Artikel" in wf.sheetnames:
        ws = wf["Neue Artikel"]
        old["Neue Artikel"] = [[ws.cell(r, c).value for c in range(1, 11)] for r in range(5, 5 + N_NEW)]
    if "Gebinde-Kategorie" in wf.sheetnames:
        ws = wf["Gebinde-Kategorie"]
        rows = []
        for r in range(5, ws.max_row + 1):
            v = [ws.cell(r, c).value for c in range(1, 6)]
            ext = [ws.cell(r, c).value for c in (8, 9, 10, 12)] if str(ws["H4"].value or "").startswith("Stellmass") else [None] * 4
            ext += [ws.cell(r, 14).value if str(ws["N4"].value or "").startswith("Höhe") else None]
            if v[0]:
                rows.append([str(v[0]).strip().upper()] + v[1:] + ext)
        old["Gebinde"] = rows
    return old


# ============================================================================
# Workbook bauen
# ============================================================================
R0 = 7  # erste Datenzeile im Blatt Artikel

PARAM_HL_DEFAULT = [  # Hauptlager, Liftgruppe, Tablare je loser Artikel (Startwert)
    ("LIFT1", 1, 1.0), ("KTL", "2+3", 0.03), ("PAL", "2+3", 0.15), ("NEU", "2+3", 0.15),
    ("A-BEZ", "Nein", 0.15), ("OCC", "Nein", 0.15), ("VERB-P", "Nein", 0.15), ("EK-TV", "Nein", 0.15),
    ("BEKLEI", "Nein", 0.15), ("GK", "Nein", 0.15), ("VERPAC", "Nein", 0.15), ("GERAET", "Nein", 0.15),
    ("ENTSORGEN!", "Nein", 0.15), ("DUMMY", "Nein", 0.15), ("(leer)", "Nein", 0.15),
]
TOOL_VERSION = "v4-oe"
LIZ_DEFAULT = [  # «Lager in Zukunft» (Artikelstamm) -> Lift
    ("Kardex Schwer", 1), ("Kardex Kleinteil", "2+3"), ("PAL", "Nein"), ("Nicht NLZ", "Nein"), ("SVC Verpackung", "Nein"),
    ("Verpackungsmaterial", "Nein"), ("Kisten Leer", "Nein"), ("Steuerschrank & Gestell", "Nein"), ("Kompressor", "Nein"),
    ("-", "Nein"),
]
LIZ_FIRST, LIZ_LAST = 35, 52
GEB_LIFT_DEFAULT = [(c, 3) for c in ("S21", "S22", "S32", "S33", "S41", "S51", "S52")] + \
                   [(c, 2) for c in ("S61", "S62", "S63", "S71", "S72", "S73", "S81", "S82", "S83",
                                     "P20", "P21", "P22", "P23", "P24", "P25", "LOSE")]
LU_LIST = [2, 2.5, 3, 3.5, 4, 5, 6, 8]
DREIER_SZ = ["Ist", "LU 2", "LU 3"]
EINKAUF_SORT = ["Kombiniert (Platz + Wert + tiefer LU)", "Grösste Lagerfläche", "Höchster Lagerwert", "Tiefster Umschlag",
                "Grösstes Abbaupotenzial"]
VARIANTS = ["A – Ist-Bestand", "B – Ziel-LU", "C – Ziel-LU oder Ist"]
IST_MODES = ["Export (Bestandsmenge)", "Jahresanfang 2025 + Export"]
HL_FIRST, HL_LAST = 35, 64
GL_FIRST, GL_LAST = 35, 64


def build(D, old):
    arts = D["arts"]
    n_art = len(arts)
    RN = R0 + n_art + N_NEW - 1  # letzte Datenzeile
    wb = Workbook()
    wsC = wb.active
    wsC.title = "Cockpit"
    wsE = wb.create_sheet("Einkauf")
    wsD = wb.create_sheet("DREIER")
    wsR = wb.create_sheet("Ergebnis je Artikel")
    wsS = wb.create_sheet("Szenarien")
    wsA = wb.create_sheet("Artikel")
    wsP = wb.create_sheet("Parameter")
    wsN = wb.create_sheet("Neue Artikel")
    wsG = wb.create_sheet("Gebinde-Kategorie")
    wsL = wb.create_sheet("Liftbericht")
    wsI = wb.create_sheet("Anleitung")
    oldP = dict((old or {}).get("Parameter", {}))
    if (old or {}).get("version") not in ("v2-tablare", "v3-bestand", TOOL_VERSION):  # Startwerte lose waren früher m² -> nicht übernehmen
        for r in range(HL_FIRST, HL_LAST + 1):
            oldP.pop(f"C{r}", None)
    if (old or {}).get("version") not in ("v3-bestand", TOOL_VERSION):  # ab v3: Ist-Bestand = Jahresanfang 2025 + Bewegung
        oldP.pop("B20", None)
    if (old or {}).get("version") != TOOL_VERSION:  # ab v4: neue LU-Liste und Standard-Ansichten
        for r in range(23, 31):
            oldP.pop(f"A{r}", None)
    oldC = dict((old or {}).get("Cockpit", {}))
    if (old or {}).get("version") != TOOL_VERSION:
        oldC = {}

    def pin(ws, ref, default, oldd, **kw):
        """gelbe Eingabezelle, alter Wert hat Vorrang"""
        v = oldd.get(ref, default)
        return put(ws, ref, v, fill=YELLOW, border=True, **kw)

    # ------------------------------------------------------------------ Parameter
    ws = wsP
    put(ws, "A1", "Parameter", bold=True, size=14)
    put(ws, "A2", "Gelbe Felder dürfen geändert werden. Alles andere rechnet automatisch.", italic=True)
    put(ws, "A3", "Lagerlifte", bold=True, size=11)
    heads = ["Lift", "Name", "Typ", "Hauptlager", "Anzahl Tablare", "Tablar Breite mm", "Tablar Tiefe mm",
             "Lifthöhe mm", "max. Ladehöhe je Tablar mm (optional)", "Tablarfläche m²", "Kapazität m²",
             "Belegung heute gemessen % (leer = schätzen)"]
    for j, h in enumerate(heads):
        put(ws, f"{L(j + 1)}4", h, bold=True, fill=SUB, wrap=True, border=True)
    lifts = [("Lift 1 – Schwer", "ausfahrbar, schwere Teile", "LIFT1"),
             ("Lift 2 – Mittel", "ausfahrbar, mittelschwere Teile", "KTL + PAL (mit Lift 3)"),
             ("Lift 3 – Klein", "nicht ausfahrbar, Kleinteile", "KTL + PAL (mit Lift 2)")]
    for i, (nm, typ, hl) in enumerate(lifts):
        r = 5 + i
        put(ws, f"A{r}", i + 1, bold=True, border=True)
        pin(ws, f"B{r}", nm, oldP)
        pin(ws, f"C{r}", typ, oldP)
        pin(ws, f"D{r}", hl, oldP)
        pin(ws, f"E{r}", 50, oldP, fmt="0")
        pin(ws, f"F{r}", 4060, oldP, fmt="#,##0")
        pin(ws, f"G{r}", 857, oldP, fmt="#,##0")
        pin(ws, f"H{r}", 11900, oldP, fmt="#,##0")
        pin(ws, f"I{r}", None, oldP, fmt="#,##0")
        put(ws, f"J{r}", f"=F{r}*G{r}/1000000", fmt="0.000", border=True)
        put(ws, f"K{r}", f"=E{r}*J{r}", fmt="#,##0.0", border=True)
        pin(ws, f"L{r}", 0.85 if i == 0 else None, oldP, fmt="0%")
    put(ws, "A8", "Total", bold=True, border=True)
    put(ws, "E8", "=SUM(E5:E7)", bold=True, fmt="0", border=True)
    put(ws, "K8", "=SUM(K5:K7)", bold=True, fmt="#,##0.0", border=True)
    put(ws, "H8", "=SUM(H5:H7)", bold=True, fmt="#,##0", border=True)
    dv = DataValidation(type="whole", operator="between", formula1="40", formula2="60", allow_blank=False)
    dv.error = "Anzahl Tablare 40–60"
    ws.add_data_validation(dv)
    dv.add("E5:E7")

    put(ws, "A10", "Ampel Auslastung", bold=True, size=11)
    put(ws, "A11", "grün unter")
    pin(ws, "B11", 0.85, oldP, fmt="0%")
    put(ws, "A12", "rot über")
    pin(ws, "B12", 1.0, oldP, fmt="0%")
    put(ws, "C11", "dazwischen orange", italic=True)

    put(ws, "A14", "Verbrauchszeitraum", bold=True, size=11)
    y = re.fullmatch(r"\d{4}", D["period"] or "")
    v_von = datetime.datetime(int(D["period"]), 1, 1) if y else oldP.get("B15", datetime.datetime(2025, 1, 1))
    v_bis = datetime.datetime(int(D["period"]), 12, 31) if y else oldP.get("B16", datetime.datetime(2025, 12, 31))
    put(ws, "A15", "Verbrauch von")
    put(ws, "B15", v_von, fill=YELLOW, fmt="DD.MM.YYYY", border=True)
    put(ws, "A16", "Verbrauch bis")
    put(ws, "B16", v_bis, fill=YELLOW, fmt="DD.MM.YYYY", border=True)
    put(ws, "A17", "Anzahl Monate")
    put(ws, "B17", "=MAX(1,(YEAR(B16)-YEAR(B15))*12+MONTH(B16)-MONTH(B15)+1)", fmt="0", border=True)
    put(ws, "A18", "Faktor auf 12 Monate")
    put(ws, "B18", "=12/B17", fmt="0.000", border=True)
    put(ws, "C15", f"Export-Filter «Auswertung Verbrauch»!B2 = {D['period']}", italic=True)

    put(ws, "A19", "Lager in Zukunft verwenden", bold=True, size=11)
    pin(ws, "B19", "Ja", oldP)
    dv = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B19")
    put(ws, "C19", "Ja = Lift nach Spalte «Lager in Zukunft» (Tabelle I34). Leer = Regel nach Hauptlager.", italic=True)
    put(ws, "I32", "Zuordnung «Lager in Zukunft» → Lift  (1 / 2 / 3 / 2+3 / Nein)", bold=True, size=11)
    put(ws, "I34", "Lager in Zukunft", bold=True, fill=SUB, border=True)
    put(ws, "J34", "Lift", bold=True, fill=SUB, border=True)
    old_liz = {str(oldP[f"I{r}"]): oldP.get(f"J{r}") for r in range(LIZ_FIRST, LIZ_LAST + 1) if oldP.get(f"I{r}")}
    liz_rows = [(k, old_liz.pop(k, v)) for k, v in LIZ_DEFAULT] + list(old_liz.items())
    known_liz = {k for k, _ in liz_rows}
    for a in arts:
        if a.get("liz") and a["liz"] not in known_liz:
            liz_rows.append((a["liz"], "Nein"))
            known_liz.add(a["liz"])
    for i in range(LIZ_LAST - LIZ_FIRST + 1):
        r = LIZ_FIRST + i
        k, v = liz_rows[i] if i < len(liz_rows) else (None, None)
        put(ws, f"I{r}", k, fill=YELLOW, border=True)
        put(ws, f"J{r}", v, fill=YELLOW, border=True, align="center")
    dv = DataValidation(type="list", formula1='"1,2,3,2+3,Nein"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"J{LIZ_FIRST}:J{LIZ_LAST}")
    put(ws, "A20", "Ist-Bestand aus", bold=True, size=11)
    pin(ws, "B20", IST_MODES[1], oldP)
    dv = DataValidation(type="list", formula1='"' + ",".join(IST_MODES) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B20")
    put(ws, "C20", "Export (Bestandsmenge) = Spalte G wie geliefert, negativ = 0.  "
                   "«Jahresanfang 2025 + Export» = Spalte T + Spalte G (falls G die Bewegung seit 01.01.2025 ist).",
        italic=True)
    put(ws, "A21", "Liftbericht verwenden", bold=True, size=11)
    pin(ws, "B21", "Ja", oldP)
    dv = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B21")
    put(ws, "C21", "Ja = für Artikel im Liftbericht gelten dessen Bestand und Lift. Die Fachflächen werden so verteilt, "
                   "dass «Ist» der gemessenen Belegung (Spalte L) entspricht.", italic=True)
    put(ws, "A22", "Ziel-LU Auswahl", bold=True, size=11)
    put(ws, "D22", "Varianten (nicht ändern)", bold=True, size=11)
    for i, v in enumerate(LU_LIST):
        pin(ws, f"A{23 + i}", v, oldP, fmt="0.0")
    for i, v in enumerate(VARIANTS):
        put(ws, f"D{23 + i}", v)
    put(ws, "H22", "DREIER (nicht ändern)", bold=True, size=11)
    for i, v in enumerate(DREIER_SZ):
        put(ws, f"H{23 + i}", v)
    put(ws, "G22", "Einkauf: Sortierung (nicht ändern)", bold=True, size=11)
    for i, v in enumerate(EINKAUF_SORT):
        put(ws, f"G{23 + i}", v)
    put(ws, "A32", "Zuordnung Hauptlager → Lift  (1 / 2 / 3 / 2+3 / Nein)", bold=True, size=11)
    put(ws, "F32", "Zuordnung Gebinde → Lift (nur Liftgruppe 2+3)", bold=True, size=11)
    for j, h in enumerate(["Hauptlager", "Lift", "Tablare je loser Artikel (Startwert)"]):
        put(ws, f"{L(j + 1)}34", h, bold=True, fill=SUB, wrap=True, border=True)
    for j, h in enumerate(["Gebinde", "Lift"]):
        put(ws, f"{L(j + 6)}34", h, bold=True, fill=SUB, border=True)
    # Hauptlager-Tabelle: alte Werte + Defaults + neue aus Daten
    hl_rows = []
    old_hl = {}
    for r in range(HL_FIRST, HL_LAST + 1):
        k = oldP.get(f"A{r}")
        if k:
            old_hl[str(k)] = (oldP.get(f"B{r}", "Nein"), oldP.get(f"C{r}", 0.5))
    for k, gl, la in PARAM_HL_DEFAULT:
        hl_rows.append((k,) + old_hl.pop(k, (gl, la)))
    hl_rows += [(k,) + v for k, v in old_hl.items()]
    known = {x[0] for x in hl_rows}
    for a in arts:
        if a["hl"] not in known:
            hl_rows.append((a["hl"], "Nein", 0.5))
            known.add(a["hl"])
    hl_rows = hl_rows[:HL_LAST - HL_FIRST + 1]
    for i in range(HL_LAST - HL_FIRST + 1):
        r = HL_FIRST + i
        k, gl, la = hl_rows[i] if i < len(hl_rows) else (None, None, None)
        put(ws, f"A{r}", k, fill=YELLOW, border=True)
        put(ws, f"B{r}", gl, fill=YELLOW, border=True, align="center")
        put(ws, f"C{r}", la, fill=YELLOW, border=True, fmt="0.00")
    dv = DataValidation(type="list", formula1='"1,2,3,2+3,Nein"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"B{HL_FIRST}:B{HL_LAST}")
    old_gl = {}
    for r in range(GL_FIRST, GL_LAST + 1):
        k = oldP.get(f"F{r}")
        if k:
            old_gl[str(k).upper()] = oldP.get(f"G{r}", 2)
    gl_rows = [(k, old_gl.pop(k, v)) for k, v in GEB_LIFT_DEFAULT] + list(old_gl.items())
    for i in range(GL_LAST - GL_FIRST + 1):
        r = GL_FIRST + i
        k, v = gl_rows[i] if i < len(gl_rows) else (None, None)
        put(ws, f"F{r}", k, fill=YELLOW, border=True)
        put(ws, f"G{r}", v, fill=YELLOW, border=True, align="center")
    dv = DataValidation(type="list", formula1='"2,3"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"G{GL_FIRST}:G{GL_LAST}")
    put(ws, "F33", "Bemerkung «grosser Artikel …» im Artikelstamm → immer Lift 2", italic=True)
    put(ws, "A33", "«Nein» = zählt nicht zu den Liften (Zeile «Nicht im Lift»). ENTSORGEN!/DUMMY standardmässig «Nein».", italic=True)

    put(ws, "Z1", TOOL_VERSION, color="FFFFFF")
    put(ws, "I20", "Optionen Tablar-Berechnung", bold=True, size=11)
    put(ws, "I21", "Tablar-Ausnutzung (Gebinde passen nie lückenlos; 90 % = 10 % Reserve)")
    pin(ws, "M21", 1.0, oldP, fmt="0%")
    put(ws, "I22", "Gebinde je Tablar berechnen mit Tablarmass von Lift")
    pin(ws, "M22", 2, oldP, fmt="0")
    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("M22")
    put(ws, "I23", "Zuschlag je Tablar (Tablarstärke + Zug) mm")
    pin(ws, "M23", 35, oldP, fmt="0")
    put(ws, "I24", "Höhe je Tablar für lose Artikel mm (inkl. Zuschlag)")
    pin(ws, "M24", 235, oldP, fmt="0")
    put(ws, "D27", "Alle 3 Lifte mischen (gemeinsam beurteilen)", bold=True)
    pin(ws, "E27", "Ja", oldP)
    dv = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("E27")
    put(ws, "D28", "Ja = Lifte ausgeglichen befüllt (Spalte «Lift ausgeglichen»: Liftbericht bleibt in Lift 1, Trennbleche nur Lift 1/2, "
                   "sonst so, dass Ist, B und C bei LU 3 möglichst gleichmässig verteilt sind); «Passt?» über die Summe aller 3 Lifte. Nein = Regel nach Gebinde.", italic=True)
    put(ws, "D30", "Platzbedarf rechnen mit", bold=True)
    pin(ws, "E30", "Ø-Bestand", oldP)
    dv = DataValidation(type="list", formula1='"Ø-Bestand,MAX-Bestand"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("E30")
    put(ws, "D31", "Ø-Bestand = Material rotiert, nicht alle Artikel gleichzeitig auf MAX (wie Lagerplanung «Max.-Bestand Ergebnis»). "
                   "MAX-Bestand = fixe Fächer für MB + ganze Bestellmenge.", italic=True)
    put(ws, "I25", "Grobe Schätzung lose Artikel (ohne Gebinde, KTL/PAL)", bold=True, size=11)
    for j, h in enumerate(["Klasse", "Schätz-Gebinde", "Menge pro Gebinde", "Typische Artikel"]):
        put(ws, f"{L(9 + j)}26", h, bold=True, fill=SUB, border=True, wrap=True)
    for i, (kl, gb, mp, txt_) in enumerate(KLASSEN_DEFAULT):
        r = 27 + i
        put(ws, f"I{r}", kl, bold=True, border=True)
        pin(ws, f"J{r}", gb, oldP)
        pin(ws, f"K{r}", mp, oldP, fmt="0")
        put(ws, f"L{r}", txt_, italic=True)
    put(ws, "I30", "Zuordnung je Artikel: Blatt «Artikel», Spalte «Schätzklasse lose» (Vorschlag aus der Bezeichnung, änderbar). "
                   "Eigene Werte in «Gebindekategorie neu» oder «Tablare lose» haben Vorrang.", italic=True)
    # Kontrolle
    put(ws, "I10", "Kontrolle Datenstand", bold=True, size=11)
    put(ws, "I11", "H4 «Total Bestandswert» im Export (= SUM(H7:H1440))")
    put(ws, "M11", D["h4"], fmt="#,##0.00")
    put(ws, "I12", "Summe Bestandswert Export, alle Zeilen (inkl. negative)")
    put(ws, "I13", "Summe Bestandswert Export, nur Artikel mit Bestand > 0")
    put(ws, "I14", "Ist-Wert im Tool (Ist-Bestand × Preis)")
    put(ws, "I15", "Artikel mit negativer Bestandsmenge im Export")
    put(ws, "I16", "Artikel im Tool (ohne neue)")
    put(ws, "I17", "Artikel im Liftbericht / davon nicht im Kennzahlen-Auszug")
    put(ws, "I18", "Ist-Belegung Lift 1 lt. Tool (Kontrolle gegen L5)")
    put(ws, "I19", "Ergänzt aus Artikelstamm (Bestand, nicht im Kennzahlen-Auszug)")
    # Formeln dafür weiter unten (Spaltenbuchstaben nötig)
    ws.column_dimensions["A"].width = 16
    for c, w in zip("BCDEFGHIJKL", (22, 28, 24, 10, 11, 11, 11, 14, 11, 11, 14)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A5"

    # ------------------------------------------------------------------ Gebinde-Kategorie
    ws = wsG
    put(ws, "A1", "Gebinde-Kategorie", bold=True, size=14)
    put(ws, "A2", "Masse in cm (gelb). Gebinde je Tablar (manuell) und Höhe im Lift = Werte aus der Lagerplanung. Stellmass = Aussenmass auf dem Tablar in mm. Gebinde je Tablar = beste Anordnung (längs/quer) "
                  "auf dem Tablar × Lagen. Manuelle Anzahl je Tablar überschreibt die Berechnung. Neue Kategorien unten ergänzen.", italic=True)
    put(ws, "A3", '="Tablar: "&INDEX(Parameter!$F$5:$F$7,Parameter!$M$22)&" × "&INDEX(Parameter!$G$5:$G$7,Parameter!$M$22)&" mm (Lift "&Parameter!$M$22&")"',
        bold=True, color="1F3864")
    for j, h in enumerate(["Gebinde", "Bezeichnung", "Länge cm", "Breite cm", "Höhe cm", "Grundfläche m²", "Volumen m³",
                           "Stellmass L mm", "Stellmass B mm", "Lagen (stapeln)", "Gebinde je Tablar (berechnet)",
                           "Gebinde je Tablar (manuell)", "Gebinde je Tablar (wirksam)", "Höhe im Lift cm (ohne Zuschlag)"]):
        put(ws, f"{L(j + 1)}4", h, bold=True, fill=SUB, border=True, wrap=True)
    ws.row_dimensions[4].height = 42

    def stellmass(code, bez, lcm, bcm):
        m = re.search(r"(\d{3,4})\s*x\s*(\d{3,4})", bez or "")
        if m and not str(bez).upper().startswith("BITO"):
            return float(m.group(1)), float(m.group(2))
        return (round(num(lcm) * 10), round(num(bcm) * 10)) if lcm else (None, None)

    geb = {}
    for r in D["gebkat"]:
        lm, bm_ = stellmass(*r[:4])
        geb[r[0]] = list(r[:5]) + [lm, bm_, 1, PLAN_GPT.get(r[0]), PLAN_HCM.get(r[0], r[4])]
    for r in (old or {}).get("Gebinde", []):
        base = geb.get(r[0], list(r[:5]) + [None, None, 1, None, r[4]])
        r = list(r) + [None] * (10 - len(r))
        new = list(r[:5]) + [r[5] if r[5] is not None else base[5], r[6] if r[6] is not None else base[6],
                             r[7] if r[7] is not None else base[7], r[8] if r[8] is not None else base[8],
                             r[9] if r[9] is not None else base[9]]
        if new[5] is None:
            new[5], new[6] = stellmass(*new[:4])
        geb[r[0]] = new
    order = [r[0] for r in D["gebkat"]] + [k for k in geb if k not in {r[0] for r in D["gebkat"]}]
    GEB_LAST = 5 + max(len(order) + 15, 40)
    TL = "INDEX(Parameter!$F$5:$F$7,Parameter!$M$22)"
    TT = "INDEX(Parameter!$G$5:$G$7,Parameter!$M$22)"
    for i in range(GEB_LAST - 4):
        r = 5 + i
        v = geb[order[i]] if i < len(order) else [None] * 10
        for j in range(5):
            put(ws, f"{L(j + 1)}{r}", v[j], fill=YELLOW, border=True, fmt="0.0" if j >= 2 else None)
        put(ws, f"F{r}", f'=IF(A{r}="","",C{r}*D{r}/10000)', fmt="0.0000", border=True)
        put(ws, f"G{r}", f'=IF(A{r}="","",C{r}*D{r}*E{r}/1000000)', fmt="0.0000", border=True)
        put(ws, f"H{r}", v[5], fill=YELLOW, border=True, fmt="0")
        put(ws, f"I{r}", v[6], fill=YELLOW, border=True, fmt="0")
        put(ws, f"J{r}", v[7], fill=YELLOW, border=True, fmt="0")
        put(ws, f"K{r}", (f'=IF(OR(A{r}="",N(H{r})<=0,N(I{r})<=0),"",MAX(INT({TL}/H{r})*INT({TT}/I{r}),'
                          f'INT({TL}/I{r})*INT({TT}/H{r}))*MAX(1,N(J{r})))'), fmt="0", border=True)
        put(ws, f"L{r}", v[8], fill=YELLOW, border=True, fmt="0")
        put(ws, f"M{r}", f'=IF(A{r}="","",IF(N(L{r})>0,L{r},N(K{r})))', fmt="0", border=True, bold=True)
        put(ws, f"N{r}", v[9], fill=YELLOW, border=True, fmt="0")
    for c, w in zip("ABCDEFGHIJKLMN", (10, 36, 9, 9, 9, 11, 10, 10, 10, 9, 11, 11, 11, 11)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "C5"
    GEB_RNG = f"'Gebinde-Kategorie'!$A$5:$N${GEB_LAST}"

    # ------------------------------------------------------------------ Neue Artikel
    ws = wsN
    put(ws, "A1", "Neue Artikel", bold=True, size=14)
    put(ws, "A2", f"Bis {N_NEW} Artikel erfassen. Sie werden im Blatt «Artikel» (unterste Zeilen) automatisch mitgerechnet und als «Neu» markiert.", italic=True)
    put(ws, "A3", "Beispiel: 112345 | Dichtung DN50 | 120 | 0 | 10 | 50 | 4.20 | S41 | 25 | 3   (Lift leer = Vorschlag nach Gebinde)", italic=True, color="595959")
    nh = ["Artikelnummer", "Bezeichnung", "erwarteter Jahresverbrauch", "Ist-Bestand", "Mindestbestand (MB)",
          "Losgrösse (LG)", "Preis CHF", "Gebindekategorie (leer = lose)", "Menge pro Gebinde", "Lift (1/2/3/Aussen)"]
    for j, h in enumerate(nh):
        put(ws, f"{L(j + 1)}4", h, bold=True, fill=SUB, wrap=True, border=True)
    oldN = (old or {}).get("Neue Artikel") or [[None] * 10 for _ in range(N_NEW)]
    for i in range(N_NEW):
        for j in range(10):
            put(ws, f"{L(j + 1)}{5 + i}", oldN[i][j] if i < len(oldN) else None, fill=YELLOW, border=True)
    dv = DataValidation(type="list", formula1='"1,2,3,Aussen"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"J5:J{4 + N_NEW}")
    ws.row_dimensions[4].height = 42
    for c, w in zip("ABCDEFGHIJ", (14, 32, 13, 11, 12, 11, 10, 14, 11, 12)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A5"

    # ------------------------------------------------------------------ Artikel
    ws = wsA
    COLS = []  # (key, header, width, kind, group)   kind: v=Wert, f=Formel, i=Eingabe gelb

    def col(key, header, width=10, kind="v", grp=0, fmt=None):
        COLS.append(dict(key=key, header=header, width=width, kind=kind, grp=grp, fmt=fmt))

    col("nr", "Artikelnummer", 12); col("bez", "Bezeichnung", 30); col("status", "Status", 6)
    col("hl", "Hauptlager", 9); col("liz", "Lager in Zukunft", 12); col("abc", "ABC", 5); col("preis", "Preis CHF", 9, fmt="#,##0.00")
    col("ist", "Ist-Bestand", 9, "f", fmt="#,##0"); col("istwert", "Ist-Wert CHF", 11, "f", fmt="#,##0")
    col("verb", "Verbrauch 12 Monate", 10, "f", fmt="#,##0"); col("avg", "Ø-Bestand", 9, fmt="#,##0.0")
    col("lu_ist", "IST-LU", 7, fmt="0.00")
    col("mb", "Mindestbestand (MB)", 9, fmt="#,##0"); col("lg", "Losgrösse (LG)", 9, fmt="#,##0"); col("wbz", "WBZ Tage", 7, fmt="0")
    col("geb", "Gebindekategorie", 9); col("mpg", "Menge pro Gebinde", 8, fmt="#,##0")
    col("gpt", "Gebinde je Tablar", 8, "f", fmt="0")
    col("lose_in", "Tablare lose (ganzer Bestand)", 10, "i", fmt="0.00")
    col("klasse", "Schätzklasse lose", 8, "i")
    col("geb_neu", "Gebindekategorie neu", 9, "i"); col("mpg_neu", "Menge pro Gebinde neu", 9, "i")
    col("liftr", "Lift lt. Liftbericht", 7); col("kapr", "Fach-Kapazität Stk (Liftbericht)", 8, fmt="#,##0")
    col("bestr", "Bestand lt. Liftbericht", 8, fmt="#,##0")
    col("lausg", "Lift ausgeglichen", 7)
    col("lift_vor", "Lift Vorschlag", 7, "f"); col("lift_in", "Lift manuell", 7, "i")
    col("lu_art", "Ziel-LU Artikel", 7, "i", fmt="0.0")
    col("maxb", "MAX Variante B (Ansicht 2)", 10, "f", fmt="#,##0"); col("maxc", "MAX Variante C (Ansicht 2)", 10, "f", fmt="#,##0")
    col("max_fin", "MAX final", 9, "i", fmt="#,##0")
    for v in (1, 2, 3):
        col(f"q{v}", f"A{v} Menge", 9, "f", 1, "#,##0"); col(f"n{v}", f"A{v} Anzahl Gebinde", 8, "f", 1, "#,##0")
        col(f"tb{v}", f"A{v} Tablare", 8, "f", 1, "0.00"); col(f"a{v}", f"A{v} Höhe mm", 8, "f", 1, "#,##0"); col(f"d{v}", f"A{v} Ø-Bestand", 9, "f", 1, "#,##0.0"); col(f"w{v}", f"A{v} Wert CHF (Ø)", 10, "f", 1, "#,##0")
        col(f"b{v}", f"A{v} MAX B (Hilfe)", 8, "f", 2, "#,##0"); col(f"pr{v}", f"A{v} Priorität Verbr./Tablar (Hilfe)", 9, "f", 2, "#,##0.0")
        col(f"cum{v}", f"A{v} Fläche kumuliert (Hilfe)", 9, "f", 2, "#,##0.0")
        col(f"ant{v}", f"A{v} Anteil im Lift (Hilfe)", 7, "f", 2, "0%")
    for v in (1, 2, 3):
        col(f"loc{v}", f"Lagerort Ansicht {v}", 11, "f")
    for k, nm in enumerate(DREIER_SZ):
        col(f"dq{k}", f"DREIER {nm} Stk", 9, "f", fmt="#,##0")
    col("hint", "Hinweise", 45, "f"); col("ursache", "Ursache (Kennzahlen)", 30); col("aktion", "Aktion (Kennzahlen)", 30)
    col("komm", "Kommentar Diskussion", 30, "i")
    # Hilfsspalten
    for k, h, f in [("bm", "Export Bestandsmenge (G)", "#,##0"), ("ja", "Export Jahresanfang 2025 (T)", "#,##0"),
                    ("bwx", "Export Bestandswert (H)", "#,##0.00"), ("verbx", "Verbrauch Export (Zeitraum)", "#,##0"),
                    ("bem", "Bemerkungen Artikelstamm", None), ("neu", "Neu (1/0)", "0"), ("idx", "Zeile", "0"),
                    ("geb_e", "Gebinde wirksam", None), ("mpg_e", "Menge/Gebinde wirksam", "#,##0"),
                    ("lose", "Lose (1/0)", "0"), ("est", "Gebinde geschätzt (1/0)", "0"), ("hpt", "Höhe je Tablar mm", "0"), ("zgrp", "Liftregel (LiZ/Hauptlager)", None), ("la", "Tablare lose wirksam", "0.00"), ("lift", "Lift wirksam", None),
                    ("grp", "Liftgruppe", None), ("luo", "Ziel-LU Artikel (0=keiner)", "0.0"), ("finf", "MAX final gesetzt", "0"),
                    ("finv", "MAX final Wert", "#,##0"), ("mblg", "MB + LG", "#,##0"), ("mbn", "MB (Zahl)", "#,##0"), ("lgn", "LG (Zahl)", "#,##0"), ("istn", "Gebinde bei Ist", "#,##0"),
                    ("finn", "Gebinde bei MAX final", "#,##0"), ("hgt", "Gebindehöhe cm", "0"), ("maxh", "max. Ladehöhe Lift mm", "0"),
                    ("fach", "Fach lt. Liftbericht (Tablare)", "0.00"),
                    ("hI", "Einkauf: Höhe Ist mm", "#,##0"), ("luI", "Einkauf: IST-LU", "0.00"), ("wI", "Einkauf: Ist-Wert", "#,##0"),
                    ("avgD", "Einkauf: Ziel-Ø", "#,##0.0"), ("abbau", "Einkauf: Abbaupotenzial CHF", "#,##0"),
                    ("frei", "Einkauf: Platz frei mm", "#,##0"), ("score", "Einkauf: Priorität (0–3)", "0.00"),
                    ("ekey", "Einkauf: Sortierschlüssel", "0.000000"), ("mx0", "DREIER 0: MAX", "#,##0"), ("ho0", "DREIER 0: Höhe Überbestand", "#,##0"), ("cuo0", "DREIER 0: Überbestand kumuliert", "#,##0"), ("he0", "DREIER 0: Überbestand raus", "#,##0"), ("r0", "DREIER 0: Höhe Rest", "#,##0"), ("cu0", "DREIER 0: kumuliert (tiefer LU zuerst)", "#,##0"), ("mv0", "DREIER 0: ganz raus (1/0)", "0"), ("dh0", "DREIER 0: Höhe raus", "#,##0"), ("dw0", "DREIER 0: Wert raus", "#,##0"), ("mx1", "DREIER 1: MAX", "#,##0"), ("ho1", "DREIER 1: Höhe Überbestand", "#,##0"), ("cuo1", "DREIER 1: Überbestand kumuliert", "#,##0"), ("he1", "DREIER 1: Überbestand raus", "#,##0"), ("r1", "DREIER 1: Höhe Rest", "#,##0"), ("cu1", "DREIER 1: kumuliert (tiefer LU zuerst)", "#,##0"), ("mv1", "DREIER 1: ganz raus (1/0)", "0"), ("dh1", "DREIER 1: Höhe raus", "#,##0"), ("dw1", "DREIER 1: Wert raus", "#,##0"), ("mx2", "DREIER 2: MAX", "#,##0"), ("ho2", "DREIER 2: Höhe Überbestand", "#,##0"), ("cuo2", "DREIER 2: Überbestand kumuliert", "#,##0"), ("he2", "DREIER 2: Überbestand raus", "#,##0"), ("r2", "DREIER 2: Höhe Rest", "#,##0"), ("cu2", "DREIER 2: kumuliert (tiefer LU zuerst)", "#,##0"), ("mv2", "DREIER 2: ganz raus (1/0)", "0"), ("dh2", "DREIER 2: Höhe raus", "#,##0"), ("dw2", "DREIER 2: Wert raus", "#,##0"), ("dkey", "DREIER: Sortierschlüssel LU", "0.000000000"), ("dsk", "DREIER: Liste Schlüssel", "0.000000")]:
        col(k, h, 9, "f", 3, f)
    C = {c["key"]: L(i + 1) for i, c in enumerate(COLS)}
    KIND = {c["key"]: c["kind"] for c in COLS}

    def F(t, r):
        """@key@ -> Zelle gleiche Zeile, #key# -> absoluter Bereich"""
        t = re.sub(r"@(\w+)@", lambda m: f"{C[m.group(1)]}{r}", t)
        return re.sub(r"#(\w+)#", lambda m: f"${C[m.group(1)]}${R0}:${C[m.group(1)]}${RN}", t)

    def RG(k, sheet="Artikel"):
        return f"{sheet}!${C[k]}${R0}:${C[k]}${RN}"

    # Cockpit-Referenzen
    VB = {v: 2 + (v - 1) * 6 for v in (1, 2, 3)}       # Startspalte Block
    VAR = {v: f"Cockpit!${L(VB[v] + 1)}$5" for v in (1, 2, 3)}
    LUC = {v: f"Cockpit!${L(VB[v] + 1)}$6" for v in (1, 2, 3)}
    CAP1, CAP2, CAP3 = "Parameter!$H$5", "Parameter!$H$6", "Parameter!$H$7"   # Kapazität = Lifthöhe mm
    HLT = f"Parameter!$A${HL_FIRST}:$C${HL_LAST}"
    GLT = f"Parameter!$F${GL_FIRST}:$G${GL_LAST}"
    LIZT = f"Parameter!$I${LIZ_FIRST}:$J${LIZ_LAST}"
    LIZT_K = f"Parameter!$I${LIZ_FIRST}:$I${LIZ_LAST}"

    FORM = {
        "ist": ('=MAX(0,IF(AND(Parameter!$B$21="Ja",ISNUMBER(@bestr@)),@bestr@,'
                'IF(Parameter!$B$20="Jahresanfang 2025 + Export",@ja@+@bm@,@bm@)))'),
        "istwert": "=@ist@*@preis@",
        "hI": ('=IF(AND(OR(@grp@="L1",@grp@="L23",@grp@="LIFT"),@ist@>0),@hpt@*IF(@fach@>0,@fach@*MAX(1,@ist@/@kapr@),'
               'IF(@lose@=1,@la@,IF(@gpt@>0,@istn@/@gpt@,0))/MAX(0.1,N(Parameter!$M$21))),"")'),
        "luI": '=IF(@hI@="","",@verb@/@ist@)',
        "wI": '=IF(@hI@="","",@ist@*@preis@)',
        "avgD": '=IF(@hI@="","",(@mbn@+ROUNDUP(@mbn@+MAX(@lgn@,2*(@verb@/MAX(0.1,N(Einkauf!$C$6))-@mbn@)),0))/2)',
        "abbau": '=IF(@hI@="","",MAX(0,@ist@-@avgD@)*@preis@)',
        "frei": '=IF(@hI@="","",@hI@*MAX(0,1-@avgD@/@ist@))',
        "score": '=IF(@hI@="","",PERCENTRANK(#hI#,@hI@)+PERCENTRANK(#wI#,@wI@)+1-PERCENTRANK(#luI#,@luI@))',
        "mx0": '=IF(@hI@="","",@ist@)',
        "ho0": '=IF(@hI@="","",@hI@*MAX(0,@ist@-@mx0@)/@ist@)',
        "cuo0": '=IF(@hI@="","",SUMIFS(#ho0#,#dkey#,"<"&@dkey@))',
        "he0": '=IF(@hI@="","",IF(DREIER!$C$9="Alles",@ho0@,IF(@cuo0@<DREIER!$F$5,@ho0@,0)))',
        "r0": '=IF(@hI@="","",@hI@-@he0@)',
        "cu0": '=IF(@hI@="","",SUMIFS(#r0#,#dkey#,"<"&@dkey@))',
        "mv0": '=IF(@hI@="","",IF(@cu0@<DREIER!$C$11,1,0))',
        "dh0": '=IF(@hI@="","",@he0@+@mv0@*@r0@)',
        "dq0": '=IF(@hI@="","",IF(@mv0@=1,@ist@,IF(@he0@>0,MAX(0,@ist@-@mx0@),0)))',
        "dw0": '=IF(@hI@="","",@dq0@*@preis@)',
        "mx1": '=IF(@hI@="","",ROUNDUP(@mbn@+MAX(@lgn@,2*(@verb@/MAX(0.1,N(DREIER!$C$6))-@mbn@)),0))',
        "ho1": '=IF(@hI@="","",@hI@*MAX(0,@ist@-@mx1@)/@ist@)',
        "cuo1": '=IF(@hI@="","",SUMIFS(#ho1#,#dkey#,"<"&@dkey@))',
        "he1": '=IF(@hI@="","",IF(DREIER!$C$9="Alles",@ho1@,IF(@cuo1@<DREIER!$F$5,@ho1@,0)))',
        "r1": '=IF(@hI@="","",@hI@-@he1@)',
        "cu1": '=IF(@hI@="","",SUMIFS(#r1#,#dkey#,"<"&@dkey@))',
        "mv1": '=IF(@hI@="","",IF(@cu1@<DREIER!$D$11,1,0))',
        "dh1": '=IF(@hI@="","",@he1@+@mv1@*@r1@)',
        "dq1": '=IF(@hI@="","",IF(@mv1@=1,@ist@,IF(@he1@>0,MAX(0,@ist@-@mx1@),0)))',
        "dw1": '=IF(@hI@="","",@dq1@*@preis@)',
        "mx2": '=IF(@hI@="","",ROUNDUP(@mbn@+MAX(@lgn@,2*(@verb@/MAX(0.1,N(DREIER!$C$7))-@mbn@)),0))',
        "ho2": '=IF(@hI@="","",@hI@*MAX(0,@ist@-@mx2@)/@ist@)',
        "cuo2": '=IF(@hI@="","",SUMIFS(#ho2#,#dkey#,"<"&@dkey@))',
        "he2": '=IF(@hI@="","",IF(DREIER!$C$9="Alles",@ho2@,IF(@cuo2@<DREIER!$F$5,@ho2@,0)))',
        "r2": '=IF(@hI@="","",@hI@-@he2@)',
        "cu2": '=IF(@hI@="","",SUMIFS(#r2#,#dkey#,"<"&@dkey@))',
        "mv2": '=IF(@hI@="","",IF(@cu2@<DREIER!$E$11,1,0))',
        "dh2": '=IF(@hI@="","",@he2@+@mv2@*@r2@)',
        "dq2": '=IF(@hI@="","",IF(@mv2@=1,@ist@,IF(@he2@>0,MAX(0,@ist@-@mx2@),0)))',
        "dw2": '=IF(@hI@="","",@dq2@*@preis@)',
        "dkey": '=IF(@hI@="","",@luI@+ROW()/1000000000)',
        "dsk": '=IF(@hI@="","",IF(CHOOSE(MATCH(DREIER!$C$8,Parameter!$H$23:$H$25,0),@dh0@,@dh1@,@dh2@)>0,CHOOSE(MATCH(DREIER!$C$8,Parameter!$H$23:$H$25,0),@dh0@,@dh1@,@dh2@)+ROW()/1000000000,""))',
        "ekey": ('=IF(@hI@="","",CHOOSE(MATCH(Einkauf!$C$5,Parameter!$G$23:$G$27,0),@score@,@hI@,@wI@,-@luI@,@abbau@)+ROW()/1000000000)'),
        "verb": "=@verbx@*Parameter!$B$18",
        "gpt": f'=IF(@lose@=1,0,IFERROR(VLOOKUP(@geb_e@,{GEB_RNG},13,0)+0,0))',
        "zgrp": (f'=IF(AND(Parameter!$B$19="Ja",@liz@<>"",ISNUMBER(MATCH(@liz@,{LIZT_K},0))),VLOOKUP(@liz@,{LIZT},2,0),'
                 f'IFERROR(VLOOKUP(@hl@,{HLT},2,0),"Nein"))'),
        "lift_vor": (f'=IF(@nr@="","",IF(AND(Parameter!$B$21="Ja",ISNUMBER(@liftr@)),@liftr@,'
                     f'IF(AND(Parameter!$E$27="Ja",ISNUMBER(@lausg@),@zgrp@<>"Nein"),@lausg@,IF(@zgrp@="2+3",'
                     f'IF(ISNUMBER(SEARCH("grosser Artikel",@bem@)),2,IFERROR(VLOOKUP(@geb_e@,{GLT},2,0),2)),'
                     f'IF(OR(@zgrp@="Nein",@zgrp@=""),"-",@zgrp@)))))'),
        "maxb": "=@b2@",
        "maxc": "=MAX(@b2@,@ist@)",
        "hint": ('=IF(@nr@="","",IF(@neu@=1,"Neu; ","")'
                 '&IF(AND(@neu@=0,@bm@<0),"negativer Bestand (Export "&@bm@&"); ","")'
                 '&IF(@verb@<=0,"Kein Verbrauch; ","")'
                 '&IF(@fach@>0,"Lift "&@liftr@&" lt. Liftbericht (Fach "&@kapr@&" Stk); ",IF(@est@=1,"Gebinde geschätzt ("&@klasse@&"); ",IF(@lose@=1,"Lose – Tablare geschätzt; ","")))'
                 '&IF(@ist@>@b2@,"Überbestand "&ROUND(@ist@-@b2@,0)&" Stk / CHF "&ROUND((@ist@-@b2@)*@preis@,0)&"; ","")'
                 '&IF(AND(@finf@=1,@finv@<@mblg@),"MAX < MB+LG; ","")'
                 '&IF(AND(@maxh@>0,@hgt@*10>@maxh@),"zu hoch für Tablar; ","")'
                 '&IF(AND(@neu@=0,OR(@finf@=1,@lift_in@<>"",@geb_neu@<>"",@mpg_neu@<>"",@lu_art@<>"")),"manuell übersteuert; ","")'
                 '&IF(AND(@lose@=0,@gpt@=0),"Gebinde unbekannt / Stellmass fehlt; ","")'
                 '&IF(AND(@lose@=0,N(@mpg_neu@)<=0,N(@mpg@)<=0),"Menge pro Gebinde fehlt (Ist = 1 Gebinde); ","")'
                 '&@bem@)'),
        "est": '=IF(AND(@geb_neu@="",OR(@geb@="",@geb@="LOSE"),@klasse@<>"",N(@lose_in@)<=0,ISNUMBER(MATCH(@klasse@,Parameter!$I$27:$I$29,0))),1,0)',
        "geb_e": '=IF(@geb_neu@<>"",UPPER(TRIM(@geb_neu@)),IF(@est@=1,UPPER(VLOOKUP(@klasse@,Parameter!$I$27:$K$29,2,0)),IF(@geb@="","LOSE",@geb@)))',
        "mpg_e": ('=IF(@geb_e@="LOSE",1,IF(N(@mpg_neu@)>0,@mpg_neu@,IF(@est@=1,MAX(1,N(VLOOKUP(@klasse@,Parameter!$I$27:$K$29,3,0)),@ist@,N(@lg@)),'
                  'IF(N(@mpg@)>0,@mpg@,MAX(1,@ist@)))))'),
        "lose": '=IF(@geb_e@="LOSE",1,0)',
        "la": f'=IF(@lose@=0,0,IF(N(@lose_in@)>0,@lose_in@,IFERROR(VLOOKUP(@hl@,{HLT},3,0)+0,0)))',
        "lift": '=IF(@lift_in@<>"",@lift_in@,@lift_vor@)',
        "grp": ('=IF(OR(@lift@=1,@lift@=2,@lift@=3),IF(Parameter!$E$27="Ja","LIFT",IF(@lift@=1,"L1","L23")),'
                'IF(@lift@="Aussen","Aussen","-"))'),
        "hpt": (f'=IF(@fach@>0,INDEX(Parameter!$H$5:$H$7,@liftr@)/INDEX(Parameter!$E$5:$E$7,@liftr@),'
                f'IF(AND(@lose@=0,N(IFERROR(VLOOKUP(@geb_e@,{GEB_RNG},14,0),0))>0),VLOOKUP(@geb_e@,{GEB_RNG},14,0)*10+Parameter!$M$23,Parameter!$M$24))'),
        "luo": "=IF(N(@lu_art@)>0,@lu_art@,0)",
        "finf": "=IF(ISNUMBER(@max_fin@),1,0)",
        "finv": "=N(@max_fin@)",
        "mblg": "=N(@mb@)+N(@lg@)",
        "mbn": "=N(@mb@)",
        "lgn": "=N(@lg@)",
        "istn": "=ROUNDUP(@ist@/@mpg_e@,0)",
        "finn": "=ROUNDUP(@finv@/@mpg_e@,0)",
        "hgt": f'=IF(@lose@=1,0,IFERROR(VLOOKUP(@geb_e@,{GEB_RNG},5,0)+0,0))',
        "fach": ('=IF(AND(Parameter!$B$21="Ja",ISNUMBER(@liftr@),N(@kapr@)>0),IFERROR(N(INDEX(Parameter!$L$5:$L$7,@liftr@))'
                 '*INDEX(Parameter!$E$5:$E$7,@liftr@)*@kapr@/SUMIFS(#kapr#,#liftr#,@liftr@),0),0)'),
        "maxh": "=IFERROR(IF(OR(@lift@=1,@lift@=2,@lift@=3),N(INDEX(Parameter!$I$5:$I$7,@lift@)),0),0)",
    }
    for v in (1, 2, 3):
        var, lu = VAR[v], LUC[v]
        cap = f'IF(@grp@="L1",{CAP1},IF(@grp@="L23",{CAP2}+{CAP3},{CAP1}+{CAP2}+{CAP3}))'
        FORM.update({
            f"b{v}": f"=ROUNDUP(@mbn@+MAX(@lgn@,2*(@verb@/IF(@luo@>0,@luo@,MAX(0.1,N({lu})))-@mbn@)),0)",
            f"q{v}": (f'=IF(LEFT({var},1)="A",@ist@,IF(Parameter!$E$30="MAX-Bestand",IF(@finf@=1,@finv@,IF(LEFT({var},1)="B",@b{v}@,MAX(@b{v}@,@ist@))),ROUNDUP(@d{v}@,0)))'),
            f"n{v}": f"=IF(@lose@=1,0,ROUNDUP(@q{v}@/@mpg_e@,0))",
            f"a{v}": f"=@tb{v}@*@hpt@",
            f"tb{v}": (f"=IF(@fach@>0,@fach@*MAX(1,@q{v}@/@kapr@),"
                      f"IF(@lose@=1,IF(@q{v}@<=0,0,@la@*MAX(1,IF(@ist@>0,@q{v}@/@ist@,1))),IF(@gpt@>0,@n{v}@/@gpt@,0))"
                      f"/MAX(0.1,N(Parameter!$M$21)))"),
            f"d{v}": (f'=IF(LEFT({var},1)="A",@ist@,IF(@finf@=1,(@mbn@+@finv@)/2,'
                      f'IF(LEFT({var},1)="B",(@mbn@+@b{v}@)/2,MAX((@mbn@+@b{v}@)/2,@ist@))))'),
            f"w{v}": f"=@d{v}@*@preis@",
            f"pr{v}": f"=IF(@a{v}@>0,@verb@/@a{v}@,-1)",
            f"cum{v}": (f'=IF(AND(OR(@grp@="L1",@grp@="L23",@grp@="LIFT"),@a{v}@>0),SUMIFS(#a{v}#,#grp#,@grp@,#pr{v}#,">"&@pr{v}@)'
                        f'+SUMIFS(#a{v}#,#grp#,@grp@,#pr{v}#,@pr{v}@,#idx#,"<="&@idx@),"")'),
            f"ant{v}": f'=IF(@cum{v}@="",0,MAX(0,MIN(1,({cap}-(@cum{v}@-@a{v}@))/@a{v}@)))',
            f"loc{v}": (f'=IF(@a{v}@<=0,"",IF(@grp@="Aussen","Aussenlager",IF(@grp@="-","Nicht im Lift",'
                        f'IF(@ant{v}@>=1,"Lift",IF(@ant{v}@<=0,"Aussenlager","teilweise")))))'),
        })

    # Kopf
    put(ws, "A1", "Artikel – eine Zeile pro Artikel", bold=True, size=14)
    put(ws, "A2", "Gelb = Eingabe (leer lassen = Vorschlag gilt). Weiss = Formel. Spalten A1–A3 = Cockpit-Ansichten 1–3 "
                  "(Gruppen mit + / – auf- und zuklappen). Ganz rechts: Hilfsspalten.", italic=True)
    put(ws, "A3", f"Datenstand: Kennzahlen-Auszug ({os.path.basename(KENNZAHLEN_FILE)}), Gebinde aus "
                  f"{os.path.basename(ARTIKELSTAMM_FILE)}. Erzeugt {datetime.date.today():%d.%m.%Y}.", italic=True, color="595959")
    groups_title = {1: "Ansichten (Cockpit)", 2: "Hilfsspalten Zuteilung", 3: "Hilfsspalten"}
    for i, c in enumerate(COLS):
        cl = L(i + 1)
        put(ws, f"{cl}6", c["header"], bold=True, fill=YELLOW if c["kind"] == "i" else SUB, wrap=True, border=True)
        ws.column_dimensions[cl].width = c["width"]
    ws.row_dimensions[6].height = 54
    # Gruppen-Überschrift Zeile 5
    for v in (1, 2, 3):
        put(ws, f"{C[f'q{v}']}5", f'="Ansicht {v}: "&{VAR[v]}&IF(LEFT({VAR[v]},1)="A",""," / LU "&{LUC[v]})', bold=True, color="1F3864")
    put(ws, f"{C['bm']}5", "Hilfsspalten (nicht ändern)", bold=True, color="595959")

    oldA = (old or {}).get("Artikel", {})
    hdr2key = {c["header"]: c["key"] for c in COLS}
    inkeys = [c["key"] for c in COLS if c["kind"] == "i"]

    rows = []
    for a in arts:
        rows.append(("art", a))
    for i in range(N_NEW):
        rows.append(("new", i))
    NA = "'Neue Artikel'!"
    for idx, (kind, a) in enumerate(rows):
        r = R0 + idx
        vals = {}
        if kind == "art":
            vals = dict(nr=a["nr"], bez=a["bez"], status=a["status"], hl=a["hl"], abc=a["abc"], preis=a["preis"],
                        avg=a["avg"], lu_ist=a["lu_ist"], mb=a["mb"], lg=a["lg"], wbz=a["wbz"], geb=a["geb"],
                        liz=a.get("liz") or None,
                        klasse=(schaetzklasse(a["bez"], a["hl"]) if a["geb"] == "LOSE" and a.get("lift_rep") is None
                                and (a["hl"] in ("KTL", "PAL") or str(a.get("liz", "")).startswith("Kardex")) else None),
                        mpg=a["mpg"], ursache=a["ursache"] or None, aktion=a["aktion"] or None,
                        bm=a["bm"], ja=a["ja"], bwx=a["bwx"], verbx=a["verbx"], bem=a["bem"], neu=0,
                        liftr=a.get("lift_rep"), kapr=a.get("kap_rep"), bestr=a.get("best_rep"),
                        lausg=D.get("ausg", {}).get(a["nr"]))
            prev = oldA.get(a["nr"], {})
        else:
            nr_ = f"{NA}$A${5 + a}"
            q = lambda c: f"{NA}${c}${5 + a}"
            vals = dict(nr=f'=IF({nr_}="","",{nr_}&"")', bez=f'={q("B")}&""', status="", hl="NEU", abc="",
                        preis=f'=N({q("G")})', mb=f'=N({q("E")})', lg=f'=N({q("F")})',
                        geb=f'=IF({q("H")}="","LOSE",UPPER(TRIM({q("H")})))', mpg=f'=N({q("I")})',
                        lift_in=f'=IF({q("J")}="","",{q("J")})', bm=f'=N({q("D")})', ja=0, bwx=0,
                        verbx=f'=N({q("C")})/Parameter!$B$18', bem="", neu=1)
            prev = {}
        vals["idx"] = r
        for i, c in enumerate(COLS):
            k = c["key"]
            cell = ws.cell(r, i + 1)
            if k in vals:
                cell.value = vals[k]
            elif k in FORM:
                cell.value = F(FORM[k], r)
            elif c["kind"] == "i":
                ov = prev.get(c["header"])
                cell.value = ov
            if c["kind"] == "i" and not (kind == "new" and k == "lift_in"):
                cell.fill = YELLOW
            if c["fmt"]:
                cell.number_format = c["fmt"]
            cell.font = Font(name=FONT, size=9)
    # Gruppen
    for v in (1, 2, 3):
        ws.column_dimensions.group(C[f"q{v}"], C[f"w{v}"], outline_level=1, hidden=False)
        ws.column_dimensions.group(C[f"b{v}"], C[f"ant{v}"], outline_level=2, hidden=True)
    ws.column_dimensions.group(C["bm"], C["fach"], outline_level=1, hidden=True)
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.freeze_panes = "C7"
    ws.auto_filter.ref = f"A6:{C['komm']}{RN}"
    dv = DataValidation(type="list", formula1='"1,2,3,Aussen"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{C['lift_in']}{R0}:{C['lift_in']}{R0 + n_art - 1}")
    dv = DataValidation(type="list", formula1="Parameter!$I$27:$I$29", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{C['klasse']}{R0}:{C['klasse']}{RN}")
    # bedingte Formate
    for v in (1, 2, 3):
        rng = f"{C[f'loc{v}']}{R0}:{C[f'loc{v}']}{RN}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Lift"'], fill=GREEN_F))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"teilweise"'], fill=ORANGE_F))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Aussenlager"'], fill=RED_F))
    hr = f"{C['hint']}{R0}:{C['hint']}{RN}"
    ws.conditional_formatting.add(hr, FormulaRule(formula=[f'ISNUMBER(SEARCH("zu hoch",{C["hint"]}{R0}))'], fill=RED_F))
    ws.conditional_formatting.add(hr, FormulaRule(formula=[f'ISNUMBER(SEARCH("Überbestand",{C["hint"]}{R0}))'], fill=ORANGE_F))

    # Defined Names (für Szenarien)
    for k in ("ist", "verb", "mblg", "mpg_e", "gpt", "lose", "la", "luo", "finf", "finv", "lift", "preis", "istn", "finn", "fach", "kapr", "hpt", "mbn", "lgn"):
        nm = "A_" + k.upper()
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Artikel!${C[k]}${R0}:${C[k]}${RN}")

    # Kontrolle-Formeln (Parameter)
    put(wsP, "M12", f"=SUM({RG('bwx')})", fmt="#,##0.00")
    put(wsP, "M13", f'=SUMIFS({RG("bwx")},{RG("bm")},">0")', fmt="#,##0.00")
    put(wsP, "M14", f"=SUM({RG('istwert')})", fmt="#,##0.00")
    put(wsP, "M15", f'=COUNTIFS({RG("bm")},"<0",{RG("neu")},0)', fmt="0")
    put(wsP, "M16", f'=COUNTIFS({RG("neu")},0)', fmt="0")
    put(wsP, "M17", f'=COUNT({RG("liftr")})&" / "&COUNTIFS({RG("liftr")},">0",{RG("bem")},"nur im Liftbericht*")')
    put(wsP, "M18", "=Szenarien!C5", fmt="0%")
    put(wsP, "M19", f'=COUNTIFS({RG("bem")},"nur im Artikelstamm*")', fmt="0")
    wsP.column_dimensions["M"].width = 14

    # ------------------------------------------------------------------ Cockpit
    ws = wsC
    put(ws, "B1", "Cockpit – Lagerlift-Belegung & MAX-Bestand nach Lagerumschlag", bold=True, size=16, color="1F3864")
    put(ws, "B2", '="Verbrauchsbasis: "&TEXT(MONTH(Parameter!B15),"00")&"."&YEAR(Parameter!B15)&" – "&TEXT(MONTH(Parameter!B16),"00")&"."&YEAR(Parameter!B16)'
                  '&"   ·   Ist-Bestand: "&Parameter!B20&"   ·   Kapazität "&TEXT(Parameter!H8,"#,##0")&" mm Lifthöhe (3 × 11 900)"&IF(Parameter!E27="Ja","   ·   Lifte gemischt","")&"   ·   IST-LU gesamt "&TEXT(SUMPRODUCT(A_VERB,A_PREIS)/MAX(1,SUMPRODUCT(A_IST,A_PREIS)),"0.0")',
        bold=True, size=11, color="C00000")
    put(ws, "B3", "Gelbe Felder = Regler. Auslastung: grün < 85 %, orange 85–100 %, rot > 100 %.", italic=True, color="595959")
    defaults = {1: (VARIANTS[0], 3), 2: (VARIANTS[2], 3), 3: (VARIANTS[1], 4)}
    dvv = DataValidation(type="list", formula1="Parameter!$D$23:$D$25", allow_blank=False)
    dvl = DataValidation(type="list", formula1="Parameter!$A$23:$A$30", allow_blank=False, showErrorMessage=False)
    ws.add_data_validation(dvv)
    ws.add_data_validation(dvl)
    lifts_lbl = [("Lift 1 – Schwer", "1"), ("Lift 2 – Mittel", "2"), ("Lift 3 – Klein", "3")]
    for v in (1, 2, 3):
        c0 = VB[v]
        c = [L(c0 + j) for j in range(5)]
        ws.merge_cells(f"{c[0]}4:{c[4]}4")
        put(ws, f"{c[0]}4", f"Ansicht {v}", bold=True, size=12, fill=HEAD, color="FFFFFF", align="center")
        put(ws, f"{c[0]}5", "Variante", bold=True)
        ws.merge_cells(f"{c[1]}5:{c[4]}5")
        pin(ws, f"{c[1]}5", defaults[v][0], oldC, bold=True)
        dvv.add(f"{c[1]}5")
        put(ws, f"{c[0]}6", "Ziel-LU", bold=True)
        pin(ws, f"{c[1]}6", defaults[v][1], oldC, bold=True, fmt="0.0", align="center")
        dvl.add(f"{c[1]}6")
        put(ws, f"{c[2]}6", "(bei A ohne Wirkung)", italic=True, color="808080")
        ws.conditional_formatting.add(f"{c[1]}6", FormulaRule(formula=[f'LEFT(${c[1]}$5,1)="A"'], fill=GREY, font=Font(color="808080")))
        for j, h in enumerate(["Bereich", "Höhe belegt mm", "Lifthöhe mm", "Tablare (Anzahl)", "Füllgrad"]):
            put(ws, f"{c[j]}8", h, bold=True, fill=SUB, border=True, align="center", wrap=True)
        A = lambda k: RG(k)
        for i, (lbl, n) in enumerate(lifts_lbl):
            r = 9 + i
            put(ws, f"{c[0]}{r}", lbl, border=True)
            put(ws, f"{c[1]}{r}", f"=SUMIFS({A(f'a{v}')},{A('lift')},{n})", fmt="#,##0", border=True)
            put(ws, f"{c[2]}{r}", f"=Parameter!$H${4 + int(n)}", fmt="#,##0", border=True)
            put(ws, f"{c[3]}{r}", f"=SUMIFS({A(f'tb{v}')},{A('lift')},{n})", fmt="#,##0.0", border=True)
            put(ws, f"{c[4]}{r}", f"=IF({c[2]}{r}>0,{c[1]}{r}/{c[2]}{r},0)", fmt="0%", border=True, bold=True)
        put(ws, f"{c[0]}12", "Lift 2 + 3 gemeinsam (KTL/PAL)", border=True, bold=True)
        put(ws, f"{c[1]}12", f"={c[1]}10+{c[1]}11", fmt="#,##0", border=True, bold=True)
        put(ws, f"{c[2]}12", f"={c[2]}10+{c[2]}11", fmt="#,##0", border=True, bold=True)
        put(ws, f"{c[3]}12", f"={c[3]}10+{c[3]}11", fmt="#,##0.0", border=True, bold=True)
        put(ws, f"{c[4]}12", f"=IF({c[2]}12>0,{c[1]}12/{c[2]}12,0)", fmt="0%", border=True, bold=True)
        put(ws, f"{c[0]}13", "Total 3 Lifte", border=True, bold=True)
        put(ws, f"{c[1]}13", f"={c[1]}9+{c[1]}12", fmt="#,##0", border=True, bold=True)
        put(ws, f"{c[2]}13", f"={c[2]}9+{c[2]}12", fmt="#,##0", border=True, bold=True)
        put(ws, f"{c[3]}13", f"={c[3]}9+{c[3]}12", fmt="#,##0.0", border=True, bold=True)
        put(ws, f"{c[4]}13", f"=IF({c[2]}13>0,{c[1]}13/{c[2]}13,0)", fmt="0%", border=True, bold=True)
        put(ws, f"{c[0]}14", "Nicht im Lift (andere Hauptlager)", italic=True, border=True)
        put(ws, f"{c[1]}14", f'=SUMIFS({A(f"a{v}")},{A("grp")},"-")', fmt="#,##0", border=True, italic=True)
        put(ws, f"{c[0]}15", "Manuell «Aussen» zugeteilt", italic=True, border=True)
        put(ws, f"{c[1]}15", f'=SUMIFS({A(f"a{v}")},{A("grp")},"Aussen")', fmt="#,##0", border=True, italic=True)
        pr = f"{c[4]}9:{c[4]}13"
        ws.conditional_formatting.add(pr, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.5, color="5B9BD5", showValue=True))
        ws.conditional_formatting.add(pr, CellIsRule(operator="greaterThan", formula=["Parameter!$B$12"], fill=RED_F))
        ws.conditional_formatting.add(pr, CellIsRule(operator="greaterThanOrEqual", formula=["Parameter!$B$11"], fill=ORANGE_F))
        ws.conditional_formatting.add(pr, CellIsRule(operator="lessThan", formula=["Parameter!$B$11"], fill=GREEN_F))
        # Passt?
        put(ws, f"{c[0]}17", "Passt alles in die 3 Lifte?", bold=True, size=12)
        ws.merge_cells(f"{c[3]}17:{c[4]}18")
        put(ws, f"{c[3]}17", f'=IF(Parameter!$E$27="Ja",IF({c[1]}13<={c[2]}13,"JA","NEIN"),IF(AND({c[1]}9<={c[2]}9,{c[1]}12<={c[2]}12),"JA","NEIN"))', bold=True, size=20, align="center")
        put(ws, f"{c[0]}18", "Lift 1 (schwer, separat)")
        put(ws, f"{c[2]}18", f'=IF({c[1]}9<={c[2]}9,"JA","NEIN")', bold=True, align="center")
        put(ws, f"{c[0]}19", "Lift 2 + 3 gemeinsam")
        put(ws, f"{c[2]}19", f'=IF({c[1]}12<={c[2]}12,"JA","NEIN")', bold=True, align="center")
        ws.merge_cells(f"{c[0]}20:{c[4]}20")
        put(ws, f"{c[0]}20", (f'=IF(Parameter!$E$27="Ja","Lifte gemischt: Summe aller 3 Lifte massgebend (frei: "&TEXT(MAX(0,{c[2]}13-{c[1]}13),"0")&" mm)",IF(OR(AND({c[1]}10>{c[2]}10,{c[1]}11<{c[2]}11),AND({c[1]}11>{c[2]}11,{c[1]}10<{c[2]}10)),'
                              f'"Ausgleich Lift 2 ↔ 3 möglich: "&ROUNDUP(MIN(MAX({c[1]}10-{c[2]}10,{c[1]}11-{c[2]}11),'
                              f'MAX({c[2]}10-{c[1]}10,{c[2]}11-{c[1]}11)),0)&" mm",""))'),
            bold=True, color="C55A11")
        for rr in (f"{c[3]}17", f"{c[2]}18", f"{c[2]}19"):
            ws.conditional_formatting.add(rr, CellIsRule(operator="equal", formula=['"JA"'], fill=GREEN_F, font=Font(color="006100", bold=True)))
            ws.conditional_formatting.add(rr, CellIsRule(operator="equal", formula=['"NEIN"'], fill=RED_F, font=Font(color="9C0006", bold=True)))
        # Aussenlager
        put(ws, f"{c[0]}22", "Aussenlager nötig (mm Lifthöhe)", bold=True)
        put(ws, f"{c[2]}22", (f'=IF(Parameter!$E$27="Ja",MAX(0,{c[1]}13-{c[2]}13),MAX(0,{c[1]}9-{c[2]}9)+MAX(0,{c[1]}12-{c[2]}12))+{c[1]}15'),
            fmt="#,##0", bold=True)
        put(ws, f"{c[3]}22", f'="≈ "&TEXT({c[2]}22/Parameter!$M$24,"0")&" Tablare"', italic=True, color="808080")
        put(ws, f"{c[0]}23", "Artikel im Aussenlager (ganz/teilweise)")
        put(ws, f"{c[2]}23", f'=COUNTIF({A(f"loc{v}")},"Aussenlager")+COUNTIF({A(f"loc{v}")},"teilweise")', fmt="0")
        # Werte
        put(ws, f"{c[0]}25", "Lagerwert CHF gesamt", bold=True)
        put(ws, f"{c[2]}25", f"=SUM({A(f'w{v}')})", fmt="#,##0", bold=True)
        put(ws, f"{c[0]}26", "   davon in den Liften")
        put(ws, f"{c[2]}26", f"=SUMPRODUCT({A(f'w{v}')},{A(f'ant{v}')})", fmt="#,##0")
        put(ws, f"{c[0]}27", "   davon im Aussenlager")
        put(ws, f"{c[2]}27", f'=SUMIFS({A(f"w{v}")},{A("grp")},"L1")+SUMIFS({A(f"w{v}")},{A("grp")},"L23")+SUMIFS({A(f"w{v}")},{A("grp")},"LIFT")-{c[2]}26+SUMIFS({A(f"w{v}")},{A("grp")},"Aussen")', fmt="#,##0")
        put(ws, f"{c[0]}28", "   davon nicht im Lift (andere Lager)")
        put(ws, f"{c[2]}28", f'=SUMIFS({A(f"w{v}")},{A("grp")},"-")', fmt="#,##0")
        put(ws, f"{c[0]}30", "Artikel mit Überbestand (Ist > MAX B)", bold=True)
        put(ws, f"{c[2]}30", f'=SUMPRODUCT(({A("ist")}>{A(f"b{v}")})*1)', fmt="0")
        put(ws, f"{c[0]}31", "Überbestandswert CHF")
        put(ws, f"{c[2]}31", f'=SUMPRODUCT(({A("ist")}>{A(f"b{v}")})*({A("ist")}-{A(f"b{v}")})*{A("preis")})', fmt="#,##0")
        put(ws, f"{c[3]}30", "(LU dieser Ansicht)", italic=True, color="808080")
        # Gebinde vs lose
        for j, h in enumerate(["Höhe mm je Lift", "aus Gebinden (Stammdaten)", "Fächer lt. Liftbericht", "lose (grob geschätzt)", "Anteil geschätzt"]):
            put(ws, f"{c[j]}33", h, bold=True, fill=SUB, border=True, wrap=True, align="center")
        for i, (lbl, n) in enumerate(lifts_lbl):
            r = 34 + i
            put(ws, f"{c[0]}{r}", lbl, border=True)
            put(ws, f"{c[1]}{r}", f"=SUMIFS({A(f'a{v}')},{A('lift')},{n},{A('lose')},0,{A('fach')},0,{A('est')},0)", fmt="#,##0.0", border=True)
            put(ws, f"{c[2]}{r}", f'=SUMIFS({A(f"a{v}")},{A("lift")},{n},{A("fach")},">0")', fmt="#,##0.0", border=True)
            put(ws, f"{c[3]}{r}", f"=SUMIFS({A(f'a{v}')},{A('lift')},{n},{A('lose')},1,{A('fach')},0)+SUMIFS({A(f'a{v}')},{A('lift')},{n},{A('est')},1,{A('fach')},0)", fmt="#,##0.0", border=True)
            put(ws, f"{c[4]}{r}", f"=IF(SUM({c[1]}{r}:{c[3]}{r})>0,{c[3]}{r}/SUM({c[1]}{r}:{c[3]}{r}),0)", fmt="0%", border=True)
        put(ws, f"{c[0]}37", "Total", bold=True, border=True)
        for j in (1, 2, 3):
            put(ws, f"{c[j]}37", f"=SUM({c[j]}34:{c[j]}36)", fmt="#,##0.0", bold=True, border=True)
        put(ws, f"{c[4]}37", f"=IF(SUM({c[1]}37:{c[3]}37)>0,{c[3]}37/SUM({c[1]}37:{c[3]}37),0)", fmt="0%", bold=True, border=True)
        ws.row_dimensions[33].height = 30
    for i in range(1, 20):
        ws.column_dimensions[L(i)].width = 3 if i in (1, 7, 13) else 12
    for v in (1, 2, 3):
        ws.column_dimensions[L(VB[v])].width = 30
    ws.row_dimensions[17].height = 22
    ws.row_dimensions[18].height = 22
    # Diagrammdaten
    put(ws, "B40", "Diagrammdaten (Auslastung %)", bold=True)
    put(ws, "C40", '="A1: "&LEFT(C5,1)&IF(LEFT(C5,1)="A",""," LU "&C6)', bold=True)
    put(ws, "D40", '="A2: "&LEFT(I5,1)&IF(LEFT(I5,1)="A",""," LU "&I6)', bold=True)
    put(ws, "E40", '="A3: "&LEFT(O5,1)&IF(LEFT(O5,1)="A",""," LU "&O6)', bold=True)
    put(ws, "F40", "100 %", bold=True)
    for i, (lbl, rr) in enumerate([("Lift 1", 9), ("Lift 2", 10), ("Lift 3", 11), ("Lift 2+3", 12), ("Total", 13)]):
        r = 41 + i
        put(ws, f"B{r}", lbl)
        for j, v in enumerate((1, 2, 3)):
            put(ws, f"{L(3 + j)}{r}", f"={L(VB[v] + 4)}{rr}", fmt="0%")
        put(ws, f"F{r}", 1, fmt="0%")
    bar = BarChart()
    bar.type = "col"
    bar.title = "Auslastung je Lift – 3 Ansichten"
    bar.y_axis.title = "Auslastung"
    bar.y_axis.numFmt = "0%"
    bar.add_data(Reference(ws, min_col=3, max_col=5, min_row=40, max_row=45), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=2, min_row=41, max_row=45))
    line = LineChart()
    line.add_data(Reference(ws, min_col=6, min_row=40, max_row=45), titles_from_data=True)
    line.series[0].graphicalProperties.line.solidFill = "C00000"
    line.series[0].graphicalProperties.line.width = 28000
    line.series[0].marker.symbol = "none"
    bar += line
    bar.height, bar.width = 9, 22
    ws.add_chart(bar, "H39")
    ws.freeze_panes = "A4"
    ws.sheet_view.zoomScale = 90

    # ------------------------------------------------------------------ Szenarien
    ws = wsS
    put(ws, "A1", "Szenarien – alle LU-Werte auf einen Blick", bold=True, size=14)
    put(ws, "A2", "Berechnung direkt über das Blatt «Artikel» (SUMPRODUCT). Berücksichtigt Lift-Zuordnung, Gebinde-Anpassungen, "
                  "Ziel-LU je Artikel und MAX final. LU-Werte aus Parameter A23:A30.", italic=True)
    sh = ["Variante", "Ziel-LU", "Lift 1 %", "Lift 2 %", "Lift 3 %", "Lift 2+3 %", "Total %", "Passt alles?",
          "Aussenlager mm", "Lagerwert CHF", "Δ zum Ist CHF", "Δ zum Ist mm",
          "Lift 1 mm", "Lift 2 mm", "Lift 3 mm", "Aussen manuell mm"]
    for j, h in enumerate(sh):
        put(ws, f"{L(j + 1)}4", h, bold=True, fill=SUB, wrap=True, border=True, align="center")

    def area_expr(G, Q, V):
        """Fläche je Artikel als Array-Ausdruck. G = Anzahl Gebinde, Q = Menge (für lose), V = Menge (Fach aus Liftbericht)."""
        return (f"((A_FACH>0)*A_FACH*(1+(A_KAPR>0)*({V}>A_KAPR)*({V}/(A_KAPR+(A_KAPR=0))-1))"
                f"+(A_FACH=0)*((1-A_LOSE)*(A_GPT>0)*{G}/(A_GPT+(A_GPT=0))+A_LOSE*({Q}>0)*A_LA*(1+(A_IST>0)*({Q}>A_IST)*({Q}/(A_IST+(A_IST=0))-1)))"
                f"/MAX(0.1,N(Parameter!$M$21)))")

    def scen_exprs(var, x, mode="AVG"):
        """liefert (Höhe-Ausdruck, Lagerwert-Menge Ø)"""
        if var == "Ist":
            return area_expr("A_ISTN", "A_IST", "A_IST"), "A_IST"
        LUE = f"(A_LUO+(A_LUO=0)*{x})"
        M2 = f"(2*(A_VERB/{LUE}-A_MBN))"
        QB = f"(A_LGN+({M2}>A_LGN)*({M2}-A_LGN))"            # Bestellmenge = max(LG, 2*(V/LU-MB))
        MX = f"ROUNDUP(A_MBN+{QB},0)"                            # MAX Stück
        NB = f"ROUNDUP({MX}/A_MPG_E,0)"
        AVG = f"((A_MBN+{MX})/2)"
        if var == "B":
            W = f"(A_FINF*(A_MBN+A_FINV)/2+(1-A_FINF)*{AVG})"
            if mode == "MAX":
                G = f"(A_FINF*A_FINN+(1-A_FINF)*{NB})"
                V = f"(A_FINF*A_FINV+(1-A_FINF)*{MX})"
        else:
            W = f"(A_FINF*(A_MBN+A_FINV)/2+(1-A_FINF)*({AVG}+(A_IST>{AVG})*(A_IST-{AVG})))"
            if mode == "MAX":
                G = f"(A_FINF*A_FINN+(1-A_FINF)*({NB}+(A_ISTN>{NB})*(A_ISTN-{NB})))"
                V = f"(A_FINF*A_FINV+(1-A_FINF)*({MX}+(A_IST>{MX})*(A_IST-{MX})))"
        if mode == "AVG":
            V = f"ROUNDUP({W},0)"
            G = f"ROUNDUP({V}/A_MPG_E,0)"
        return area_expr(G, G, V), W

    srows = [("Ist", None)] + [("B", i) for i in range(8)] + [("C", i) for i in range(8)]
    for k, (var, i) in enumerate(srows):
        r = 5 + k
        if var == "Ist":
            put(ws, f"A{r}", "Ist-Bestand", bold=True, border=True)
            put(ws, f"B{r}", "–", border=True, align="center")
            x = None
        else:
            put(ws, f"A{r}", f"{var} – " + ("Ziel-LU" if var == "B" else "Ziel-LU oder Ist"), border=True)
            put(ws, f"B{r}", f"=Parameter!$A${23 + i}", fmt="0.0", border=True, align="center")
            x = f"$B{r}"
        AR, V = scen_exprs(var, x, "AVG")
        AR_MAX, _ = scen_exprs(var, x, "MAX")
        put(ws, f"M{r}", f'=IF(Parameter!$E$30="MAX-Bestand",SUMPRODUCT((A_LIFT=1)*A_HPT*{AR_MAX}),SUMPRODUCT((A_LIFT=1)*A_HPT*{AR}))', fmt="#,##0", border=True)
        put(ws, f"N{r}", f'=IF(Parameter!$E$30="MAX-Bestand",SUMPRODUCT((A_LIFT=2)*A_HPT*{AR_MAX}),SUMPRODUCT((A_LIFT=2)*A_HPT*{AR}))', fmt="#,##0", border=True)
        put(ws, f"O{r}", f'=IF(Parameter!$E$30="MAX-Bestand",SUMPRODUCT((A_LIFT=3)*A_HPT*{AR_MAX}),SUMPRODUCT((A_LIFT=3)*A_HPT*{AR}))', fmt="#,##0", border=True)
        put(ws, f"P{r}", f'=IF(Parameter!$E$30="MAX-Bestand",SUMPRODUCT((A_LIFT="Aussen")*A_HPT*{AR_MAX}),SUMPRODUCT((A_LIFT="Aussen")*A_HPT*{AR}))', fmt="#,##0", border=True)
        put(ws, f"C{r}", f"=M{r}/{CAP1}", fmt="0%", border=True)
        put(ws, f"D{r}", f"=N{r}/{CAP2}", fmt="0%", border=True)
        put(ws, f"E{r}", f"=O{r}/{CAP3}", fmt="0%", border=True)
        put(ws, f"F{r}", f"=(N{r}+O{r})/({CAP2}+{CAP3})", fmt="0%", border=True, bold=True)
        put(ws, f"G{r}", f"=(M{r}+N{r}+O{r})/Parameter!$H$8", fmt="0%", border=True, bold=True)
        put(ws, f"H{r}", (f'=IF(Parameter!$E$27="Ja",IF(M{r}+N{r}+O{r}<=Parameter!$H$8,"JA","NEIN"),'
                          f'IF(AND(M{r}<={CAP1},N{r}+O{r}<={CAP2}+{CAP3}),"JA","NEIN"))'), border=True, align="center", bold=True)
        put(ws, f"I{r}", (f'=IF(Parameter!$E$27="Ja",MAX(0,M{r}+N{r}+O{r}-Parameter!$H$8),'
                          f'MAX(0,M{r}-{CAP1})+MAX(0,N{r}+O{r}-{CAP2}-{CAP3}))+P{r}'), fmt="#,##0", border=True)
        put(ws, f"J{r}", f"=SUMPRODUCT({V}*A_PREIS)", fmt="#,##0", border=True)
        put(ws, f"K{r}", f"=J{r}-$J$5", fmt="#,##0;-#,##0", border=True)
        put(ws, f"L{r}", f"=(M{r}+N{r}+O{r}+P{r})-($M$5+$N$5+$O$5+$P$5)", fmt="#,##0;-#,##0", border=True)
        for cc in "ABCDEFGHIJKL":
            if var == "Ist":
                ws[f"{cc}{r}"].fill = SUB
    pr = "C5:G21"
    ws.conditional_formatting.add(pr, CellIsRule(operator="greaterThan", formula=["Parameter!$B$12"], fill=RED_F))
    ws.conditional_formatting.add(pr, CellIsRule(operator="greaterThanOrEqual", formula=["Parameter!$B$11"], fill=ORANGE_F))
    ws.conditional_formatting.add(pr, CellIsRule(operator="lessThan", formula=["Parameter!$B$11"], fill=GREEN_F))
    ws.conditional_formatting.add("H5:H21", CellIsRule(operator="equal", formula=['"JA"'], fill=GREEN_F))
    ws.conditional_formatting.add("H5:H21", CellIsRule(operator="equal", formula=['"NEIN"'], fill=RED_F))
    put(ws, "A23", "Ab welchem LU passt alles in die Lifte?", bold=True, size=11)
    put(ws, "A24", "Variante B (nach Abbau Überbestand):")
    put(ws, "E24", '=IFERROR("ab LU "&INDEX(B6:B13,MATCH("JA",H6:H13,0)),"bei keinem LU der Liste")', bold=True, color="C00000")
    put(ws, "A25", "Variante C (Überbestand bleibt):")
    put(ws, "E25", '=IFERROR("ab LU "&INDEX(B14:B21,MATCH("JA",H14:H21,0)),"bei keinem LU der Liste")', bold=True, color="C00000")
    put(ws, "A26", "(Annahme: LU-Liste aufsteigend sortiert. Ist-Bestand: siehe Zeile 5.)", italic=True, color="808080")
    # Diagrammdaten
    put(ws, "A28", "Diagrammdaten", bold=True)
    for j, h in enumerate(["Ziel-LU", "B – Total %", "C – Total %", "100 %", "Ist – Total %"]):
        put(ws, f"{L(j + 1)}29", h, bold=True)
    for i in range(8):
        r = 30 + i
        put(ws, f"A{r}", f"=B{6 + i}", fmt="0.0")
        put(ws, f"B{r}", f"=G{6 + i}", fmt="0%")
        put(ws, f"C{r}", f"=G{14 + i}", fmt="0%")
        put(ws, f"D{r}", 1, fmt="0%")
        put(ws, f"E{r}", "=$G$5", fmt="0%")
    lc = LineChart()
    lc.title = "Auslastung Total % je Ziel-LU"
    lc.y_axis.title = "Auslastung 3 Lifte"
    lc.x_axis.title = "Ziel-LU"
    lc.y_axis.numFmt = "0%"
    lc.add_data(Reference(ws, min_col=2, max_col=5, min_row=29, max_row=37), titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=1, min_row=30, max_row=37))
    lc.series[2].graphicalProperties.line.solidFill = "C00000"
    lc.series[2].graphicalProperties.line.dashStyle = "dash"
    lc.series[3].graphicalProperties.line.solidFill = "808080"
    lc.series[3].graphicalProperties.line.dashStyle = "sysDot"
    lc.height, lc.width = 9, 20
    ws.add_chart(lc, "G23")
    for cc, w in zip("ABCDEFGHIJKLMNOP", (24, 8, 9, 9, 9, 10, 9, 9, 11, 13, 12, 11, 10, 10, 10, 11)):
        ws.column_dimensions[cc].width = w
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "C5"

    # ------------------------------------------------------------------ Einkauf (Dashboard)
    ws = wsE
    NTOP = 50
    oldE = {} if (old or {}).get("version") != TOOL_VERSION else (old or {}).get("Einkauf", {})
    put(ws, "B1", "Einkauf – Artikel zum Geradebiegen", bold=True, size=16, color="1F3864")
    put(ws, "B2", "Lift-Artikel mit Bestand, bewertet nach Platz im Lift, Lagerwert und Umschlag. Priorität = Perzentil Platz + Perzentil Wert "
                  "+ (1 − Perzentil LU), 0 bis 3 – je höher, desto dringender. Abbaupotenzial = (Ist − Ziel-Ø) × Preis.", italic=True, color="595959")
    put(ws, "B3", "Kommentare/Entscheide bitte im Blatt «Artikel» (MAX final, Ziel-LU Artikel, Kommentar Diskussion) erfassen.", italic=True, color="595959")
    put(ws, "B5", "Sortieren nach", bold=True)
    ws.merge_cells("C5:E5")
    pin(ws, "C5", EINKAUF_SORT[0], oldE, bold=True)
    dv = DataValidation(type="list", formula1="Parameter!$G$23:$G$27", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C5")
    put(ws, "B6", "Ziel-LU (für Ziel-Ø / Abbau)", bold=True)
    pin(ws, "C6", 4, oldE, fmt="0.0", align="center")
    dv = DataValidation(type="list", formula1="Parameter!$A$23:$A$30", allow_blank=False, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add("C6")
    # KPI-Kacheln
    elig = RG("hI")
    kpis = [
        ("Lift-Artikel mit Bestand", f"=COUNT({elig})", "#,##0"),
        ("davon Umschlag < 1", f'=COUNTIFS({RG("luI")},"<1")', "#,##0"),
        ("Lagerwert dieser Artikel", f'=SUMIFS({RG("wI")},{RG("luI")},"<1")', "#,##0"),
        ("Abbaupotenzial gesamt CHF", f"=SUM({RG('abbau')})", "#,##0"),
        ("Platz frei bei Ziel-LU ≈ mm", f"=SUM({RG('frei')})", "#,##0"),
        (f"Top {NTOP}: Anteil am Abbaupotenzial", f"=IF(SUM({RG('abbau')})>0,SUM(O11:O{10 + NTOP})/SUM({RG('abbau')}),0)", "0%"),
    ]
    for j, (lbl, f, fm) in enumerate(kpis):
        col_ = L(7 + j * 2)
        ws.merge_cells(f"{col_}5:{L(8 + j * 2)}5")
        ws.merge_cells(f"{col_}6:{L(8 + j * 2)}7")
        put(ws, f"{col_}5", lbl, bold=True, size=9, fill=SUB, align="center", wrap=True)
        put(ws, f"{col_}6", f, bold=True, size=16, fmt=fm, align="center", color="1F3864")
    ws.row_dimensions[5].height = 28
    hdr_e = ["Rang", "Artikel", "Bezeichnung", "Lift", "Gebinde", "Ist-Bestand", "Verbrauch 12 M", "IST-LU", "Ist-Wert CHF",
             "Platz Ist mm", "≈ Tablare", "Priorität (0–3)", "Ziel-Ø", "Abbaupotenzial CHF", "Platz frei ≈ mm", "Ursache / Aktion"]
    for j, h in enumerate(hdr_e):
        put(ws, f"{L(2 + j)}10", h, bold=True, fill=HEAD, color="FFFFFF", wrap=True, align="center", border=True)
    ws.row_dimensions[10].height = 30
    src = {"Artikel": "nr", "Bezeichnung": "bez", "Lift": "lift", "Gebinde": "geb_e", "Ist-Bestand": "ist", "Verbrauch 12 M": "verb",
           "IST-LU": "luI", "Ist-Wert CHF": "wI", "Platz Ist mm": "hI", "Priorität (0–3)": "score", "Ziel-Ø": "avgD",
           "Abbaupotenzial CHF": "abbau", "Platz frei ≈ mm": "frei"}
    fmts = {"Ist-Bestand": "#,##0", "Verbrauch 12 M": "#,##0", "IST-LU": "0.00", "Ist-Wert CHF": "#,##0", "Platz Ist mm": "#,##0",
            "≈ Tablare": "0.0", "Priorität (0–3)": "0.00", "Ziel-Ø": "#,##0", "Abbaupotenzial CHF": "#,##0", "Platz frei ≈ mm": "#,##0"}
    for k in range(1, NTOP + 1):
        r = 10 + k
        put(ws, f"B{r}", k, border=True, align="center")
        # Zeile im Blatt Artikel (Hilfsspalte R)
        put(ws, f"R{r}", f'=IFERROR(MATCH(LARGE({RG("ekey")},B{r}),{RG("ekey")},0),"")', color="FFFFFF")
        for j, h in enumerate(hdr_e[1:], start=1):
            cl = L(2 + j)
            if h in src:
                f = f'=IF($R{r}="","",INDEX({RG(src[h])},$R{r}))'
            elif h == "≈ Tablare":
                f = f'=IF($R{r}="","",K{r}/Parameter!$M$24)'
            else:
                f = (f'=IF($R{r}="","",TRIM(INDEX({RG("ursache")},$R{r})&IF(INDEX({RG("aktion")},$R{r})<>""," → "&'
                     f'INDEX({RG("aktion")},$R{r}),"")))')
            put(ws, f"{cl}{r}", f, border=True, fmt=fmts.get(h), size=9, wrap=(h == "Ursache / Aktion"))
    last_r = 10 + NTOP
    ws.conditional_formatting.add(f"I11:I{last_r}", CellIsRule(operator="lessThan", formula=["1"], fill=RED_F))
    ws.conditional_formatting.add(f"I11:I{last_r}", CellIsRule(operator="between", formula=["1", "2"], fill=ORANGE_F))
    ws.conditional_formatting.add(f"J11:J{last_r}", DataBarRule(start_type="min", end_type="max", color="5B9BD5", showValue=True))
    ws.conditional_formatting.add(f"K11:K{last_r}", DataBarRule(start_type="min", end_type="max", color="A9D18E", showValue=True))
    ws.conditional_formatting.add(f"O11:O{last_r}", DataBarRule(start_type="min", end_type="max", color="F4B084", showValue=True))
    ws.conditional_formatting.add(f"M11:M{last_r}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=3, color="C00000", showValue=True))
    for cl, w in zip("ABCDEFGHIJKLMNOPQR", (2, 6, 11, 30, 5, 8, 9, 9, 7, 11, 9, 8, 9, 9, 12, 10, 40, 4)):
        ws.column_dimensions[cl].width = w
    ws.column_dimensions["R"].hidden = True
    ws.freeze_panes = "D11"
    # Diagramm: Top 15 Abbaupotenzial
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Top 15 – Abbaupotenzial CHF (aktuelle Sortierung)"
    ch.add_data(Reference(ws, min_col=15, min_row=10, max_row=25), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=3, min_row=11, max_row=25))
    ch.y_axis.numFmt = "#,##0"
    ch.x_axis.scaling.orientation = "maxMin"
    ch.legend = None
    ch.height, ch.width = 10, 16
    ws.add_chart(ch, f"T10")
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.tabColor = "C00000"
    ws.sheet_view.zoomScale = 90

    # ------------------------------------------------------------------ DREIER (Aussenlager)
    ws = wsD
    NLIST = 300
    oldD = {} if (old or {}).get("version") != TOOL_VERSION else (old or {}).get("DREIER", {})
    put(ws, "B1", "DREIER – was muss ins Aussenlager, damit die Lifte höchstens X % voll sind?", bold=True, size=16, color="1F3864")
    put(ws, "B2", "Basis Ist-Bestand der Lift-Artikel. Schritt 1: Überbestand über dem MAX bei LU 2 / LU 3 geht ins DREIER. Schritt 2: reicht das nicht, "
                  "gehen ganze Artikel mit dem tiefsten Umschlag (zuerst ohne Verbrauch) ins DREIER, bis der Ziel-Füllgrad erreicht ist. Lift-Artikel mit Bestand (ohne leere Fächer).", italic=True, color="595959")
    put(ws, "B3", "Fläche: ≈ Tablare = Lifthöhe / Höhe je Tablar; ≈ m² = Tablare × Tablarfläche (4'060 × 857 mm = 3.48 m²). Artikel-Spalten «DREIER … Stk» im Blatt Artikel filterbar.",
        italic=True, color="595959")
    put(ws, "B5", "Ziel-Füllgrad Lifte (max.)", bold=True)
    pin(ws, "C5", 0.8, oldD, fmt="0%", align="center")
    put(ws, "B6", "Szenario LU 2 – Ziel-LU", bold=True)
    pin(ws, "C6", 2, oldD, fmt="0.0", align="center")
    put(ws, "B7", "Szenario LU 3 – Ziel-LU", bold=True)
    pin(ws, "C7", 3, oldD, fmt="0.0", align="center")
    put(ws, "B9", "Überbestand ins DREIER", bold=True)
    pin(ws, "C9", "Nur bis Ziel-Füllgrad", oldD, align="center")
    dv = DataValidation(type="list", formula1='"Nur bis Ziel-Füllgrad,Alles"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C9")
    put(ws, "E5", "Überschuss heute über Ziel (mm)", bold=True)
    put(ws, "F5", f"=MAX(0,SUM({RG('hI')})-DREIER!$C$5*Parameter!$H$8)", fmt="#,##0", bold=True)
    put(ws, "E6", "«Nur bis Ziel»: Überbestand der Artikel mit tiefstem Umschlag zuerst, bis der Ziel-Füllgrad erreicht ist.", italic=True, color="595959")
    put(ws, "B8", "Artikelliste unten für", bold=True)
    pin(ws, "C8", "LU 3", oldD, align="center")
    dv = DataValidation(type="list", formula1="Parameter!$H$23:$H$25", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C8")
    # Übersicht
    heads = ["Kennzahl", "Ist", "LU 2", "LU 3"]
    for j, h in enumerate(heads):
        put(ws, f"{L(2 + j)}10", h, bold=True, fill=HEAD, color="FFFFFF", align="center", border=True)
    T = "(DREIER!$C$5*Parameter!$H$8)"
    rows_ = [
        ("Überschuss über Ziel nach Schritt 1 (mm)", lambda k: f"=MAX(0,SUM({RG(f'r{k}')})-{T})", "#,##0"),
        ("Lifte heute (Ist) mm", lambda k: f"=SUM({RG('hI')})", "#,##0"),
        ("Lifte heute Füllgrad", lambda k: f"={L(3 + k)}12/Parameter!$H$8", "0%"),
        ("Schritt 1 – Überbestand ins DREIER (mm)", lambda k: f"=SUM({RG(f'he{k}')})", "#,##0"),
        ("Schritt 2 – tiefer LU ins DREIER (mm)", lambda k: f"=SUMIFS({RG(f'r{k}')},{RG(f'mv{k}')},1)", "#,##0"),
        ("DREIER total Lifthöhe mm", lambda k: f"={L(3 + k)}14+{L(3 + k)}15", "#,##0"),
        ("DREIER ≈ Tablare", lambda k: f"={L(3 + k)}16/Parameter!$M$24", "#,##0.0"),
        ("DREIER ≈ Fläche m² (Tablarfläche)", lambda k: f"={L(3 + k)}17*Parameter!$J$6", "#,##0.0"),
        ("DREIER Lagerwert CHF", lambda k: f"=SUM({RG(f'dw{k}')})", "#,##0"),
        ("Artikel ganz ins DREIER", lambda k: f"=COUNTIFS({RG(f'mv{k}')},1)", "#,##0"),
        ("Artikel mit Teilmenge (nur Überbestand)", lambda k: f'=COUNTIFS({RG(f"he{k}")},">0",{RG(f"mv{k}")},0)', "#,##0"),
        ("Lifte nachher mm", lambda k: f"={L(3 + k)}12-{L(3 + k)}16", "#,##0"),
        ("Lifte nachher Füllgrad", lambda k: f"={L(3 + k)}22/Parameter!$H$8", "0%"),
        ("höchster IST-LU, der (teilweise) ins DREIER geht", lambda k: f'=IFERROR(_xlfn.MAXIFS({RG("luI")},{RG(f"dh{k}")},">0"),0)', "0.00"),
    ]
    for i, (lbl, fn, fm) in enumerate(rows_):
        r = 11 + i
        put(ws, f"B{r}", lbl, border=True, bold=i in (5, 7, 8, 12))
        for k in range(3):
            put(ws, f"{L(3 + k)}{r}", fn(k), fmt=fm, border=True, bold=i in (5, 7, 8, 12), align="right")
    for r in (16, 17, 18, 19):
        for k in range(3):
            ws[f"{L(3 + k)}{r}"].fill = ORANGE_F
    for k in range(3):
        ws[f"{L(3 + k)}23"].fill = GREEN_F
    # Artikelliste
    put(ws, "B27", '="Artikel ins DREIER – Szenario "&C8&" (sortiert nach Lifthöhe, max. ' + str(NLIST) + ' Zeilen)"', bold=True, size=12)
    hl_ = ["Rang", "Artikel", "Bezeichnung", "Lift", "Gebinde", "IST-LU", "Ist Stk", "MAX Stk", "DREIER Stk", "Rest im Lift Stk",
           "Grund", "DREIER mm", "≈ Tablare", "≈ m²", "DREIER Wert CHF"]
    for j, h in enumerate(hl_):
        put(ws, f"{L(2 + j)}28", h, bold=True, fill=HEAD, color="FFFFFF", wrap=True, align="center", border=True)
    ws.row_dimensions[28].height = 30
    KSEL = 'MATCH($C$8,Parameter!$H$23:$H$25,0)'
    for i in range(NLIST):
        r = 29 + i
        put(ws, f"B{r}", i + 1, border=True, align="center", size=9)
        put(ws, f"R{r}", f'=IFERROR(MATCH(LARGE({RG("dsk")},B{r}),{RG("dsk")},0),"")', color="FFFFFF")
        ix = lambda k_: f'INDEX({RG(k_)},$R{r})'
        cells = {
            "C": f'=IF($R{r}="","",{ix("nr")})', "D": f'=IF($R{r}="","",{ix("bez")})', "E": f'=IF($R{r}="","",{ix("lift")})',
            "F": f'=IF($R{r}="","",{ix("geb_e")})', "G": f'=IF($R{r}="","",{ix("luI")})', "H": f'=IF($R{r}="","",{ix("ist")})',
            "I": f'=IF($R{r}="","",CHOOSE({KSEL},"–",{ix("mx1")},{ix("mx2")}))',
            "J": f'=IF($R{r}="","",CHOOSE({KSEL},{ix("dq0")},{ix("dq1")},{ix("dq2")}))',
            "K": f'=IF($R{r}="","",H{r}-J{r})',
            "L": f'=IF($R{r}="","",IF(CHOOSE({KSEL},{ix("mv0")},{ix("mv1")},{ix("mv2")})=1,"tiefer LU – ganz","Überbestand"))',
            "M": f'=IF($R{r}="","",CHOOSE({KSEL},{ix("dh0")},{ix("dh1")},{ix("dh2")}))',
            "N": f'=IF($R{r}="","",M{r}/{ix("hpt")})', "O": f'=IF($R{r}="","",N{r}*Parameter!$J$6)',
            "P": f'=IF($R{r}="","",CHOOSE({KSEL},{ix("dw0")},{ix("dw1")},{ix("dw2")}))',
        }
        fm_ = {"G": "0.00", "H": "#,##0", "I": "#,##0", "J": "#,##0", "K": "#,##0", "M": "#,##0", "N": "0.0", "O": "0.0", "P": "#,##0"}
        for cl, f in cells.items():
            put(ws, f"{cl}{r}", f, border=True, size=9, fmt=fm_.get(cl))
    lr = 28 + NLIST
    ws.conditional_formatting.add(f"L29:L{lr}", CellIsRule(operator="equal", formula=['"tiefer LU – ganz"'], fill=RED_F))
    ws.conditional_formatting.add(f"L29:L{lr}", CellIsRule(operator="equal", formula=['"Überbestand"'], fill=ORANGE_F))
    ws.conditional_formatting.add(f"M29:M{lr}", DataBarRule(start_type="min", end_type="max", color="5B9BD5", showValue=True))
    put(ws, f"B{lr + 2}", f'="Angezeigt: "&COUNT(M29:M{lr})&" von "&COUNT({RG("dsk")})&" Artikeln. Vollständig: Blatt Artikel, Filter auf Spalte «DREIER … Stk» > 0."',
        italic=True, color="595959")
    for cl, w in zip("ABCDEFGHIJKLMNOPQR", (2, 34, 11, 11, 11, 8, 8, 8, 9, 9, 15, 10, 8, 8, 12, 2, 2, 2)):
        ws.column_dimensions[cl].width = w
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["R"].hidden = True
    ws.freeze_panes = "A29"
    ws.sheet_properties.tabColor = "7030A0"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ------------------------------------------------------------------ Ergebnis je Artikel
    ws = wsR
    oldR = {} if (old or {}).get("version") != TOOL_VERSION else (old or {}).get("Ergebnis je Artikel", {})
    put(ws, "A1", "Ergebnis je Artikel – MAX, Ø-Bestand, vor Ort und DREIER", bold=True, size=14, color="1F3864")
    put(ws, "A2", "Gelb = wählbar. MAX = MB + Bestellmenge (mind. Losgrösse), Ø = (MB + MAX) / 2. «MAX gültig» = MAX final (Blatt Artikel) oder MAX beim "
                  "Entscheid-LU → Wert fürs ERP. Vor Ort / DREIER gemäss gewähltem DREIER-Szenario. Filter in Zeile 5.", italic=True, color="595959")
    RR0 = 6
    A_ = lambda k, r: f"Artikel!{C[k]}{r}"
    rcols = [("Artikel", "nr", None, 12), ("Bezeichnung", "bez", None, 30), ("Hauptlager", "hl", None, 8), ("Lager in Zukunft", "liz", None, 13),
             ("Lift", "lift", None, 5), ("Gebinde", "geb_e", None, 7), ("Menge/ Gebinde", "mpg_e", "#,##0", 7), ("MB", "mbn", "#,##0", 6),
             ("LG", "lgn", "#,##0", 6), ("Verbrauch 12 M", "verb", "#,##0", 8), ("Ist", "ist", "#,##0", 7), ("IST-LU", None, "0.00", 6),
             ("Ist-Wert CHF", "istwert", "#,##0", 9)]
    for j, (h, k, fm, w) in enumerate(rcols):
        put(ws, f"{L(j + 1)}5", h, bold=True, fill=SUB, wrap=True, border=True, align="center")
        ws.column_dimensions[L(j + 1)].width = w
    c0 = len(rcols) + 1   # erste LU-Spalte
    lu_defaults = (2, 3, 4)
    for b in range(3):
        cl = L(c0 + b * 3)
        ws.merge_cells(f"{cl}3:{L(c0 + b * 3 + 2)}3")
        put(ws, f"{L(c0 + b * 3)}4", "Ziel-LU", italic=True, size=8)
        pin(ws, f"{L(c0 + b * 3 + 1)}4", lu_defaults[b], oldR, fmt="0.0", align="center", bold=True)
        put(ws, f"{cl}3", f'="LU "&TEXT({L(c0 + b * 3 + 1)}4,"0.0")', bold=True, fill=HEAD, color="FFFFFF", align="center")
        for j, h in enumerate(("MAX", "Ø-Bestand", "Gebinde bei MAX")):
            put(ws, f"{L(c0 + b * 3 + j)}5", h, bold=True, fill=SUB, wrap=True, border=True, align="center")
            ws.column_dimensions[L(c0 + b * 3 + j)].width = 8
    ce = c0 + 9   # Entscheid-Block
    ws.merge_cells(f"{L(ce)}3:{L(ce + 5)}3")
    put(ws, f"{L(ce)}3", "Entscheid / ERP", bold=True, fill=HEAD, color="FFFFFF", align="center")
    put(ws, f"{L(ce)}4", "Entscheid-LU", italic=True, size=8)
    pin(ws, f"{L(ce + 1)}4", 3, oldR, fmt="0.0", align="center", bold=True)
    for j, (h, w) in enumerate((("MAX final (Artikel)", 9), ("MAX gültig", 9), ("Ø gültig", 8), ("Überbestand Stk (Ist − MAX)", 10),
                                ("Überbestand CHF", 10), ("Quelle", 9))):
        put(ws, f"{L(ce + j)}5", h, bold=True, fill=SUB, wrap=True, border=True, align="center")
        ws.column_dimensions[L(ce + j)].width = w
    cd = ce + 6   # DREIER-Block
    ws.merge_cells(f"{L(cd)}3:{L(cd + 2)}3")
    put(ws, f"{L(cd)}3", "Vor Ort / DREIER", bold=True, fill=HEAD, color="FFFFFF", align="center")
    put(ws, f"{L(cd)}4", "Szenario", italic=True, size=8)
    pin(ws, f"{L(cd + 1)}4", "LU 3", oldR, align="center", bold=True)
    dv = DataValidation(type="list", formula1="Parameter!$H$23:$H$25", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"{L(cd + 1)}4")
    for j, (h, w) in enumerate((("ins DREIER Stk", 9), ("vor Ort (Lift) Stk", 9), ("Kommentar Diskussion", 30))):
        put(ws, f"{L(cd + j)}5", h, bold=True, fill=SUB, wrap=True, border=True, align="center")
        ws.column_dimensions[L(cd + j)].width = w
    dvl2 = DataValidation(type="list", formula1="Parameter!$A$23:$A$30", allow_blank=False, showErrorMessage=False)
    ws.add_data_validation(dvl2)
    for b in range(3):
        dvl2.add(f"{L(c0 + b * 3 + 1)}4")
    dvl2.add(f"{L(ce + 1)}4")
    ws.row_dimensions[5].height = 42
    LUD = f"${L(ce + 1)}$4"
    SZ = f"${L(cd + 1)}$4"
    for i in range(RN - R0 + 1):
        ar, r = R0 + i, RR0 + i
        for j, (h, k, fm, w) in enumerate(rcols):
            if k:
                f = f'=IF({A_("nr", ar)}="","",{A_(k, ar)}{"&" + chr(34) * 2 if fm is None and k != "lift" else ""})'
            else:
                f = f'=IF(OR({A_("nr", ar)}="",N({A_("ist", ar)})<=0),"",{A_("verb", ar)}/{A_("ist", ar)})'
            c = ws.cell(r, j + 1, f)
            c.font = Font(name=FONT, size=9)
            if fm:
                c.number_format = fm
        mb, lg, v, luo, mpg = A_("mbn", ar), A_("lgn", ar), A_("verb", ar), A_("luo", ar), A_("mpg_e", ar)
        for b in range(3):
            lu = f"{L(c0 + b * 3 + 1)}$4"
            cm, ca, cg = (L(c0 + b * 3 + j) for j in range(3))
            fs = [f'=IF($A{r}="","",ROUNDUP({mb}+MAX({lg},2*({v}/IF({luo}>0,{luo},MAX(0.1,{lu}))-{mb})),0))',
                  f'=IF($A{r}="","",({mb}+{cm}{r})/2)',
                  f'=IF($A{r}="","",IF({A_("lose", ar)}=1,"lose",ROUNDUP({cm}{r}/{mpg},0)))']
            for j, f in enumerate(fs):
                c = ws.cell(r, c0 + b * 3 + j, f)
                c.font = Font(name=FONT, size=9)
                c.number_format = "#,##0" if j != 1 else "#,##0.0"
        maxd = f'ROUNDUP({mb}+MAX({lg},2*({v}/IF({luo}>0,{luo},MAX(0.1,{LUD}))-{mb})),0)'
        fe = [f'=IF($A{r}="","",IF({A_("finf", ar)}=1,{A_("finv", ar)},""))',
              f'=IF($A{r}="","",IF({A_("finf", ar)}=1,{A_("finv", ar)},{maxd}))',
              f'=IF($A{r}="","",({mb}+{L(ce + 1)}{r})/2)',
              f'=IF($A{r}="","",MAX(0,$K{r}-{L(ce + 1)}{r}))',
              f'=IF($A{r}="","",{L(ce + 3)}{r}*{A_("preis", ar)})',
              f'=IF($A{r}="","",IF({A_("finf", ar)}=1,"MAX final",IF({luo}>0,"Ziel-LU Artikel","berechnet")))']
        for j, f in enumerate(fe):
            c = ws.cell(r, ce + j, f)
            c.font = Font(name=FONT, size=9, bold=(j == 1))
            c.number_format = "#,##0" if j != 2 else "#,##0.0"
        fd = [f'=IF($A{r}="","",IF({A_("hI", ar)}="",0,CHOOSE(MATCH({SZ},Parameter!$H$23:$H$25,0),{A_("dq0", ar)},{A_("dq1", ar)},{A_("dq2", ar)})))',
              f'=IF($A{r}="","",IF(OR($E{r}="-",$E{r}="Aussen"),0,$K{r}-{L(cd)}{r}))',
              f'=IF($A{r}="","",IF({A_("komm", ar)}="","",{A_("komm", ar)}))']
        for j, f in enumerate(fd):
            c = ws.cell(r, cd + j, f)
            c.font = Font(name=FONT, size=9)
            if j < 2:
                c.number_format = "#,##0"
    rlast = RR0 + RN - R0
    for j in range(ce + 1, ce + 2):
        for rr in range(RR0, rlast + 1):
            ws.cell(rr, j).fill = PatternFill("solid", fgColor="E2EFDA")
    ws.conditional_formatting.add(f"L{RR0}:L{rlast}", CellIsRule(operator="lessThan", formula=["1"], fill=RED_F))
    ws.conditional_formatting.add(f"{L(ce + 3)}{RR0}:{L(ce + 3)}{rlast}", CellIsRule(operator="greaterThan", formula=["0"], fill=ORANGE_F))
    ws.conditional_formatting.add(f"{L(cd)}{RR0}:{L(cd)}{rlast}", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_F))
    ws.auto_filter.ref = f"A5:{L(cd + 2)}{rlast}"
    ws.freeze_panes = "C6"
    ws.sheet_properties.tabColor = "548235"
    ws.sheet_view.zoomScale = 90

    # ------------------------------------------------------------------ Liftbericht
    ws = wsL
    put(ws, "A1", "Liftbericht (Modula «Artikelbestand für Maschine»)", bold=True, size=14)
    put(ws, "A2", "Nur zur Kontrolle – wird von build_tool.py aus input/*.prnx eingelesen. Fach-Kapazität = reservierter Platz im Lift.",
        italic=True)
    for j, h in enumerate(["Lift", "Artikel", "Bezeichnung", "Fach-Kapazität Stk", "Bestand Lift", "Hauptlager Export",
                           "Bestandsmenge Export (G)", "Jahresanfang 2025 + G", "Fach (Tablare, Tool)"]):
        put(ws, f"{L(j + 1)}4", h, bold=True, fill=SUB, wrap=True, border=True)
    byn = {a["nr"]: a for a in arts}
    for i, x in enumerate(D.get("lift_rows", [])):
        r = 5 + i
        a = byn.get(x["code"], {})
        nx = a.get("xrow") is not None
        for j, v in enumerate([x["lift"], x["code"], x["bez"], x["kap"], x["bestand"],
                               a.get("hl") if nx else "– fehlt –", a.get("bm") if nx else None,
                               (a.get("ja", 0) + a.get("bm", 0)) if nx else None]):
            put(ws, f"{L(j + 1)}{r}", v, border=True)
        put(ws, f"I{r}", f'=SUMIFS({RG("fach")},{RG("nr")},B{r})', fmt="0.000", border=True)
    n = len(D.get("lift_rows", []))
    if n:
        put(ws, f"C{5 + n}", "Total", bold=True)
        put(ws, f"D{5 + n}", f"=SUM(D5:D{4 + n})", bold=True)
        put(ws, f"E{5 + n}", f"=SUM(E5:E{4 + n})", bold=True)
        put(ws, f"I{5 + n}", f"=SUM(I5:I{4 + n})", bold=True, fmt="0.0")
    for cc, w in zip("ABCDEFGHI", (6, 10, 30, 11, 10, 12, 12, 12, 12)):
        ws.column_dimensions[cc].width = w
    ws.row_dimensions[4].height = 30
    ws.freeze_panes = "A5"

    # ------------------------------------------------------------------ Anleitung
    ws = wsI
    ws.column_dimensions["A"].width = 130
    lines = [
        ("Anleitung – Lagerumschlag-Tool", "h"),
        ("Was ändere ich? Nur gelbe Felder.", "b"),
        ("• Cockpit: je Ansicht «Variante» und «Ziel-LU» (Dropdown, freie Eingabe erlaubt). Alles rechnet sofort neu.", ""),
        ("• Artikel: Fläche lose, Gebinde neu / Menge pro Gebinde neu, Lift manuell (1/2/3/Aussen), Ziel-LU Artikel, MAX final, Kommentar. Leer = Vorschlag gilt.", ""),
        ("• Parameter: Lifte (Tablare, Tablarmass, max. Ladehöhe), Ampel, Verbrauchszeitraum, Ist-Bestand-Quelle, LU-Liste, Zuordnung Hauptlager → Lift, Gebinde → Lift, Startwerte lose.", ""),
        ("• Neue Artikel: werden automatisch unten im Blatt «Artikel» mitgerechnet («Neu»). • Gebinde-Kategorie: Masse, neue Kategorien unten ergänzen.", ""),
        ("Liftbericht (Modula)", "b"),
        ("Artikel im Liftbericht (input/*.prnx) bekommen dessen Lift und Bestand. Ihr Platz = reserviertes Fach: Die gemessene Belegung (Parameter L5:L7, z. B. Lift 1 = 85 % "
         "der Tablare) wird im Verhältnis der Fach-Kapazität (Stk) auf die Artikel verteilt. In B/C wächst der Bedarf erst, wenn MAX > Fach-Kapazität.", ""),
        ("Lifte mischen / ausgleichen", "b"),
        ("Parameter E27 = Ja: Die Spalte «Lift ausgeglichen» verteilt die Artikel so, dass alle 3 Lifte möglichst gleich voll sind – gleichzeitig für Ist, B und C bei LU 3 (Parameter im Skript: AUSGLEICH_LU): Liftbericht-Artikel "
         "bleiben in Lift 1, Trennbleche/Paletten nur in die ausfahrbaren Lifte 1 und 2, alles andere in den Lift mit dem tiefsten Füllgrad. Wird bei jedem Neuaufbau neu berechnet. "
         "«Lift manuell» hat Vorrang.", ""),
        ("Ergebnis je Artikel", "b"),
        ("Pro Artikel nebeneinander: MAX, Ø-Bestand und Gebinde bei 3 wählbaren Ziel-LUs; «MAX gültig» (MAX final oder MAX beim Entscheid-LU) = Wert fürs ERP; "
         "Überbestand; Menge ins DREIER und vor Ort gemäss DREIER-Szenario. Mit Filter in Zeile 5 exportierbar.", ""),
        ("DREIER (Aussenlager)", "b"),
        ("Blatt «DREIER»: Damit die Lifte beim Ist-Bestand höchstens den Ziel-Füllgrad (Standard 80 %) haben, geht zuerst der Überbestand über dem MAX "
         "(bei LU 2 bzw. LU 3) ins DREIER, danach ganze Artikel mit dem tiefsten Umschlag. Ergebnis als Lifthöhe, ≈ Tablare, ≈ m², Wert und Artikelliste.", ""),
        ("Einkauf (Dashboard)", "b"),
        ("Zeigt die 50 Lift-Artikel mit Bestand, die zuerst angegangen werden sollten. Sortierung wählbar: kombiniert (Platz + Wert + tiefer Umschlag), Fläche, Wert, "
         "tiefster Umschlag oder Abbaupotenzial. Abbaupotenzial = (Ist − Ziel-Ø bei gewähltem LU) × Preis; Platz frei ≈ anteilig.", ""),
        ("Lose Artikel grob geschätzt", "b"),
        ("Artikel ohne Gebinde (KTL/PAL) bekommen aus der Bezeichnung eine Schätzklasse: Klein = Eurobox S51 à 20 Stk, Mittel = S61 à 4 Stk, Gross = Trennblech S81 à 2 Stk "
         "(Parameter I27:K29); mindestens der Ist-Bestand bzw. eine Losgrösse passt in 1 Gebinde. Damit rechnen sie wie Gebinde-Artikel. Klasse im Blatt Artikel änderbar; «Gebindekategorie neu» oder «Tablare lose» haben Vorrang. Kein Stapeln.", ""),
        ("Varianten", "b"),
        ("A – Ist-Bestand: heutiger Bestand (unabhängig vom LU).", ""),
        ("B – Ziel-LU: Ziel-Ø-Bestand = Jahresverbrauch / Ziel-LU (= Spalte «Max.Bestand bei vorg. Umschlag» im Kennzahlen-Export). Bestellmenge = 2 × (Ziel-Ø − MB), "
         "mindestens Losgrösse (Mindestbestands-Strategie fix). MAX = MB + Bestellmenge → bestimmt den Platz (Gebinde). Ø-Bestand = MB + Bestellmenge/2 → bestimmt den Lagerwert.", ""),
        ("Platzbedarf (Parameter E30): «Ø-Bestand» (Standard, wie Lagerplanung: Material rotiert) oder «MAX-Bestand» (fixe Fächer für MB + Bestellmenge).", ""),
        ("C – Ziel-LU oder Ist (Hauptvariante): Platz = höherer Wert aus MAX (B) und Ist; Wert = höherer Wert aus Ø (B) und Ist – solange der Überbestand nicht abgebaut ist. "
         "Heutiger Gesamt-LU ≈ 2.4: erst ein Ziel-LU darüber senkt Bestand und Platz.", ""),
        ("MAX final (gelb) ersetzt in B und C den berechneten Wert. Ziel-LU Artikel ersetzt den LU der Ansicht für diesen Artikel. Jahresverbrauch = Verbrauch Export × 12 / Anzahl Monate.", ""),
        ("Lift-Auslastung", "b"),
        ("Anzahl Gebinde = AUFRUNDEN(Menge / Menge pro Gebinde). Stapeln nur, wenn im Blatt «Gebinde-Kategorie» Lagen > 1 eingetragen sind.", ""),
        ("Lose Artikel: Tablare = Startwert je Hauptlager (Parameter) oder eigener Wert; in B/C skaliert mit Menge / Ist-Menge, mindestens Startwert.", ""),
        ("Füllgrad über die Lifthöhe (wie Lagerplanung): Tablare je Gebindetyp = Anzahl Gebinde / Gebinde je Tablar; jedes Tablar braucht Gebindehöhe + 35 mm "
         "(Parameter M23). Füllgrad = Summe Höhe / Lifthöhe 11'900 mm. Mit «Lifte mischen» (Parameter E27) zählt die Summe aller 3 Lifte.", ""),
        ("Gebinde je Tablar = Vorgabe Lagerplanung bzw. beste Anordnung des Gebindes (Stellmass, längs/quer) auf dem Tablar × Lagen (Blatt «Gebinde-Kategorie», "
         "manuell überschreibbar). Tablare je Artikel = Anzahl Gebinde / Gebinde je Tablar. "
         "Tablar-Ausnutzung (Parameter M21) gibt Reserve für Lücken.", ""),
        ("Zuordnung zuerst nach «Lager in Zukunft» (Artikelstamm, Parameter I34): Kardex Schwer → Lift 1, Kardex Kleinteil → Lift 2/3 nach Gebinde, "
         "Verpackung/Leerkisten/PAL/Nicht NLZ → nicht im Lift. Leer → Regel nach Hauptlager: ", ""),
        ("LIFT1 → Lift 1. KTL/PAL → Lift 3 (BITO S21–S33, Eurobox S41–S52) oder Lift 2 (S61–S63, Trennbleche S71–S83, lose, «grosser Artikel»). Andere Hauptlager → «Nicht im Lift».", ""),
        ("«Passt?» wird für Lift 1 und Lift 2 + 3 gemeinsam beurteilt. Aussenlager: je Liftgruppe Artikel nach Verbrauch pro Tablar sortiert (Schnelldreher zuerst), "
         "Fläche kumuliert; was die Kapazität überschreitet = «Aussenlager», der Artikel an der Grenze = «teilweise».", ""),
        ("Hinweise", "b"),
        ("Kein Verbrauch · Lose – Fläche geschätzt · Überbestand (Ist > MAX B, Menge und CHF, LU Ansicht 2) · MAX < MB+LG (MAX final zu tief) · zu hoch für Tablar (Gebindehöhe > max. Ladehöhe) · "
         "manuell übersteuert · Neu · negativer Bestand (Export, als 0 gerechnet) · Gebinde unbekannt · + Bemerkungen aus dem Artikelstamm.", ""),
        ("Annahmen", "b"),
        ("1. Ist-Bestand = Jahresanfangsbestand 2025 (Spalte T) + Spalte G «Bestandsmenge» (= Bewegung seit 01.01.2025). KTL/PAL-Artikel mit Bestand, die im "
         "Kennzahlen-Auszug fehlen, sind aus dem Artikelstamm ergänzt (Bestand 23.06.2026). Lagerwert ≈ 3.55 Mio. CHF. Lift 1: Bestand aus Liftbericht. Umschaltbar in Parameter B20.", ""),
        ("2. Verbrauch = Pivot «Auswertung Verbrauch» Spalte C (Kalenderjahr 2025), negativ = 0; fehlt ein Artikel dort, Spalte Y «Wareneinsatz».", ""),
        ("3. MB = Spalte AK; wenn leer, Blatt «Auswertung MB» (Lager = Hauptlager). Losgrösse = Spalte D.", ""),
        ("4. Gebinde aus «Artikelstamm Umschlag 2», sonst «Artikelstamm mit Verbrauch»; ohne Gebinde = LOSE. Menge pro Gebinde fehlt → ganzer Ist-Bestand = 1 Gebinde.", ""),
        ("5. Startwerte lose (Tablare je Artikel): KTL 0.03, PAL 0.15, LIFT1 1.0, andere 0.15 – unabhängig von der Menge.", ""),
        ("6. Lagerwert = Menge × Preis GLD Aktuell. Hauptlager aus dem Kennzahlen-Auszug (nicht aus dem Artikelstamm).", ""),
        ("7. Artikel ohne Verbrauch: B = MB + LG (sonst 0). Aussenlager-Priorität: Verbrauch pro Tablar absteigend.", ""),
        ("Schieberegler statt Dropdown (optional)", "b"),
        ("Entwicklertools → Einfügen → Formularsteuerelemente → Bildlaufleiste. Rechtsklick → Steuerelement formatieren: Min 10, Max 60, Schrittweite 5, "
         "Zellverknüpfung z. B. Cockpit!Z6. In Cockpit!C6 dann =Z6/10 eintragen (LU 1.0 – 6.0).", ""),
        ("Aktualisieren", "b"),
        ("Neue Exporte in den Ordner «input» legen, Dateinamen oben in build_tool.py anpassen, «python build_tool.py» ausführen. Alle gelben Eingaben werden übernommen (Zuordnung über Artikelnummer).", ""),
    ]
    for i, (t, s) in enumerate(lines):
        c = put(ws, f"A{i + 1}", t, bold=s in ("h", "b"), size=14 if s == "h" else 10, wrap=True)
        if s == "b":
            c.fill = SUB
    ws.sheet_view.showGridLines = False

    for w in wb.worksheets:
        w.sheet_view.zoomScale = w.sheet_view.zoomScale or 100
    for w, area in ((wsC, "A1:S70"), (wsS, "A1:P45")):
        w.page_setup.orientation = "landscape"
        w.page_setup.paperSize = w.PAPERSIZE_A4
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 1
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.print_area = area
    wsI.page_setup.fitToWidth = 1
    wsI.page_setup.fitToHeight = 1
    wsI.sheet_properties.pageSetUpPr.fitToPage = True
    wsC.sheet_properties.tabColor = "1F3864"
    wsS.sheet_properties.tabColor = "2E75B6"
    wsP.sheet_properties.tabColor = "FFC000"
    wsN.sheet_properties.tabColor = "FFC000"
    return wb, C


def pruefen(path):
    """Kopie mit LibreOffice neu berechnen und auf Formelfehler prüfen (die Tool-Datei selbst bleibt unverändert;
    Excel rechnet beim Öffnen automatisch)."""
    script = RECALC_SCRIPT
    if not script or not os.path.exists(script):
        print("Prüfung übersprungen: Umgebungsvariable RECALC_SCRIPT (recalc.py mit LibreOffice) nicht gesetzt.")
        return
    tmp = path.replace(".xlsx", "_Pruefung.xlsx")
    shutil.copy2(path, tmp)
    out = subprocess.run([sys.executable, script, tmp, "600"], capture_output=True, text=True,
                         cwd=os.path.dirname(script))
    print(out.stdout[-2000:])
    os.remove(tmp)


def main():
    D = load_data()
    rep = bericht(D)
    print(rep)
    if "--bericht" in sys.argv:
        return
    old = read_old(TOOL_FILE)
    if old and os.path.abspath(TOOL_FILE) == os.path.abspath(OUTPUT_FILE):
        bak = OUTPUT_FILE.replace(".xlsx", f"_Sicherung_{datetime.datetime.now():%Y%m%d_%H%M}.xlsx")
        shutil.copy2(TOOL_FILE, bak)
        print(f"Sicherung: {bak}")
    print("Eingaben übernommen:" if old else "Kein bestehendes Tool gefunden – neu erstellt.",
          {k: len(v) for k, v in (old or {}).items()})
    D["ausg"], last = ausgleichen(D["arts"])
    print(f"Ausgeglichene Zuteilung, Füllgrad % je Lift [Ist, B, C] bei LU {AUSGLEICH_LU}:", last)
    wb, C = build(D, old)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTPUT_FILE)
    print(f"Gespeichert: {OUTPUT_FILE}")
    if "--pruefen" in sys.argv:
        pruefen(OUTPUT_FILE)
    if "--berechnen" in sys.argv:  # Werte mit LibreOffice berechnen und in der Datei speichern
        if RECALC_SCRIPT and os.path.exists(RECALC_SCRIPT):
            out = subprocess.run([sys.executable, RECALC_SCRIPT, OUTPUT_FILE, "600"], capture_output=True, text=True,
                                 cwd=os.path.dirname(RECALC_SCRIPT))
            print(out.stdout[-1500:])
        else:
            print("--berechnen: RECALC_SCRIPT nicht gesetzt – Excel berechnet beim Öffnen (ggf. F9 drücken).")


if __name__ == "__main__":
    main()

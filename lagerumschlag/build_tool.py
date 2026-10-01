# -*- coding: utf-8 -*-
"""
build_tool.py - erzeugt / aktualisiert das Excel-Tool «Lagerumschlag_Tool.xlsx».

Aufruf:   python build_tool.py            (Datenbericht + Tool bauen; Excel rechnet beim Öffnen)
          python build_tool.py --bericht  (nur Datenbericht Schritt 1)
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
import os, re, sys, shutil, datetime, subprocess, json
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
TOOL_FILE = os.path.join(HERE, "Lagerumschlag_Tool.xlsx")      # bestehendes Tool -> Eingaben übernehmen
OUTPUT_FILE = os.path.join(HERE, "Lagerumschlag_Tool.xlsx")
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
            bem=st.get("bem", ""), ursache=txt(r[41]), aktion=txt(r[42]), xrow=hdr + 2 + i,
        ))
    return dict(arts=arts, h4=h4, period=period, gebkat=gebkat, verb=verb, stamm=stamm)


# ============================================================================
# Schritt 1: Datenbericht
# ============================================================================
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
    gk = {r[0]: r[2] * r[3] / 10000 for r in D["gebkat"]}
    lift_of = lambda r: 1 if r.hl == "LIFT1" else (None if r.hl not in ("KTL", "PAL") else (
        2 if ("grosser Artikel" in r.bem or r.geb == "LOSE" or r.geb >= "S6") else 3))
    a["lift"] = a.apply(lift_of, axis=1)
    for label, q in [("Export G (neg.=0)", a.bm.clip(lower=0)), ("Jahresanfang+G", rec.clip(lower=0))]:
        mpg = a.mpg.fillna(0).where(a.mpg.fillna(0) > 0, q.clip(lower=1))  # fehlt: ganzer Bestand = 1 Gebinde
        area = np.where(a.lose, 0, np.ceil(q / mpg) * a.geb.map(gk).fillna(0))
        p(f"Ist-Belegung nur Gebinde-Artikel, Bestand = {label}:")
        for l in (1, 2, 3):
            m = area[a.lift == l].sum()
            p(f"   Lift {l}: {m:7.1f} m2 = {np.ceil(m / 3.47942):4.0f} Tablare  (Kapazität 174.0 m2 / 50 Tablare)")
    return "\n".join(out)


# ============================================================================
# Bestehende Eingaben lesen
# ============================================================================
ART_INPUT_HEADERS = ["Fläche lose m² (ganzer Bestand)", "Gebindekategorie neu", "Menge pro Gebinde neu",
                     "Lift manuell", "Ziel-LU Artikel", "MAX final", "Kommentar Diskussion"]


def read_old(path):
    if not os.path.exists(path):
        return None
    wf = load_workbook(path)
    wv = load_workbook(path, data_only=True)
    old = {}
    # Parameter: alle Zellen mit gelber Füllung
    for sh in ("Parameter", "Cockpit"):
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
            if v[0]:
                rows.append([str(v[0]).strip().upper()] + v[1:])
        old["Gebinde"] = rows
    return old


# ============================================================================
# Workbook bauen
# ============================================================================
R0 = 7  # erste Datenzeile im Blatt Artikel

PARAM_HL_DEFAULT = [  # Hauptlager, Liftgruppe, Fläche lose Startwert m²
    ("LIFT1", 1, 3.479), ("KTL", "2+3", 0.1), ("PAL", "2+3", 0.5), ("NEU", "2+3", 0.5),
    ("A-BEZ", "Nein", 0.5), ("OCC", "Nein", 0.5), ("VERB-P", "Nein", 0.5), ("EK-TV", "Nein", 0.5),
    ("BEKLEI", "Nein", 0.5), ("GK", "Nein", 0.5), ("VERPAC", "Nein", 0.5), ("GERAET", "Nein", 0.5),
    ("ENTSORGEN!", "Nein", 0.5), ("DUMMY", "Nein", 0.5), ("(leer)", "Nein", 0.5),
]
GEB_LIFT_DEFAULT = [(c, 3) for c in ("S21", "S22", "S32", "S33", "S41", "S51", "S52")] + \
                   [(c, 2) for c in ("S61", "S62", "S63", "S71", "S72", "S73", "S81", "S82", "S83",
                                     "P20", "P21", "P22", "P23", "P24", "P25", "LOSE")]
LU_LIST = [1, 1.5, 2, 2.5, 3, 4, 5, 6]
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
    wsS = wb.create_sheet("Szenarien")
    wsA = wb.create_sheet("Artikel")
    wsP = wb.create_sheet("Parameter")
    wsN = wb.create_sheet("Neue Artikel")
    wsG = wb.create_sheet("Gebinde-Kategorie")
    wsI = wb.create_sheet("Anleitung")
    oldP = (old or {}).get("Parameter", {})
    oldC = (old or {}).get("Cockpit", {})

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
             "Lifthöhe mm", "max. Ladehöhe je Tablar mm (optional)", "Tablarfläche m²", "Kapazität m²"]
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
    put(ws, "A8", "Total", bold=True, border=True)
    put(ws, "E8", "=SUM(E5:E7)", bold=True, fmt="0", border=True)
    put(ws, "K8", "=SUM(K5:K7)", bold=True, fmt="#,##0.0", border=True)
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

    put(ws, "A20", "Ist-Bestand aus", bold=True, size=11)
    pin(ws, "B20", IST_MODES[0], oldP)
    dv = DataValidation(type="list", formula1='"' + ",".join(IST_MODES) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B20")
    put(ws, "C20", "Export (Bestandsmenge) = Spalte G wie geliefert, negativ = 0.  "
                   "«Jahresanfang 2025 + Export» = Spalte T + Spalte G (falls G die Bewegung seit 01.01.2025 ist).",
        italic=True)
    put(ws, "A22", "Ziel-LU Auswahl", bold=True, size=11)
    put(ws, "D22", "Varianten (nicht ändern)", bold=True, size=11)
    for i, v in enumerate(LU_LIST):
        pin(ws, f"A{23 + i}", v, oldP, fmt="0.0")
    for i, v in enumerate(VARIANTS):
        put(ws, f"D{23 + i}", v)
    put(ws, "A32", "Zuordnung Hauptlager → Lift  (1 / 2 / 3 / 2+3 / Nein)", bold=True, size=11)
    put(ws, "F32", "Zuordnung Gebinde → Lift (nur Liftgruppe 2+3)", bold=True, size=11)
    for j, h in enumerate(["Hauptlager", "Lift", "Fläche lose Startwert m² je Artikel"]):
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
        put(ws, f"C{r}", la, fill=YELLOW, border=True, fmt="0.000")
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

    # Kontrolle
    put(ws, "I10", "Kontrolle Datenstand", bold=True, size=11)
    put(ws, "I11", "H4 «Total Bestandswert» im Export (= SUM(H7:H1440))")
    put(ws, "M11", D["h4"], fmt="#,##0.00")
    put(ws, "I12", "Summe Bestandswert Export, alle Zeilen (inkl. negative)")
    put(ws, "I13", "Summe Bestandswert Export, nur Artikel mit Bestand > 0")
    put(ws, "I14", "Ist-Wert im Tool (Ist-Bestand × Preis)")
    put(ws, "I15", "Artikel mit negativer Bestandsmenge im Export")
    put(ws, "I16", "Artikel im Tool (ohne neue)")
    # Formeln dafür weiter unten (Spaltenbuchstaben nötig)
    ws.column_dimensions["A"].width = 16
    for c, w in zip("BCDEFGHIJK", (22, 28, 24, 10, 11, 11, 11, 14, 11, 11)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A5"

    # ------------------------------------------------------------------ Gebinde-Kategorie
    ws = wsG
    put(ws, "A1", "Gebinde-Kategorie", bold=True, size=14)
    put(ws, "A2", "Länge/Breite/Höhe in cm (gelb). Grundfläche = L × B, Volumen = L × B × H. Neue Kategorien unten ergänzen.", italic=True)
    for j, h in enumerate(["Gebinde", "Bezeichnung", "Länge cm", "Breite cm", "Höhe cm", "Grundfläche m²", "Volumen m³"]):
        put(ws, f"{L(j + 1)}4", h, bold=True, fill=SUB, border=True)
    geb = {r[0]: r for r in D["gebkat"]}
    for r in (old or {}).get("Gebinde", []):
        geb[r[0]] = r
    order = [r[0] for r in D["gebkat"]] + [k for k in geb if k not in {r[0] for r in D["gebkat"]}]
    GEB_LAST = 5 + max(len(order) + 15, 40)
    for i in range(GEB_LAST - 4):
        r = 5 + i
        v = geb[order[i]] if i < len(order) else [None] * 5
        for j in range(5):
            put(ws, f"{L(j + 1)}{r}", v[j], fill=YELLOW, border=True, fmt="0.0" if j >= 2 else None)
        put(ws, f"F{r}", f'=IF(A{r}="","",C{r}*D{r}/10000)', fmt="0.0000", border=True)
        put(ws, f"G{r}", f'=IF(A{r}="","",C{r}*D{r}*E{r}/1000000)', fmt="0.0000", border=True)
    for c, w in zip("ABCDEFG", (10, 36, 10, 10, 10, 14, 12)):
        ws.column_dimensions[c].width = w
    GEB_RNG = f"'Gebinde-Kategorie'!$A$5:$G${GEB_LAST}"

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
    col("hl", "Hauptlager", 9); col("abc", "ABC", 5); col("preis", "Preis CHF", 9, fmt="#,##0.00")
    col("ist", "Ist-Bestand", 9, "f", fmt="#,##0"); col("istwert", "Ist-Wert CHF", 11, "f", fmt="#,##0")
    col("verb", "Verbrauch 12 Monate", 10, "f", fmt="#,##0"); col("avg", "Ø-Bestand", 9, fmt="#,##0.0")
    col("lu_ist", "IST-LU", 7, fmt="0.00")
    col("mb", "Mindestbestand (MB)", 9, fmt="#,##0"); col("lg", "Losgrösse (LG)", 9, fmt="#,##0"); col("wbz", "WBZ Tage", 7, fmt="0")
    col("geb", "Gebindekategorie", 9); col("mpg", "Menge pro Gebinde", 8, fmt="#,##0")
    col("garea", "Grundfläche Gebinde m²", 9, "f", fmt="0.0000")
    col("lose_in", "Fläche lose m² (ganzer Bestand)", 10, "i", fmt="0.000")
    col("geb_neu", "Gebindekategorie neu", 9, "i"); col("mpg_neu", "Menge pro Gebinde neu", 9, "i")
    col("lift_vor", "Lift Vorschlag", 7, "f"); col("lift_in", "Lift manuell", 7, "i")
    col("lu_art", "Ziel-LU Artikel", 7, "i", fmt="0.0")
    col("maxb", "MAX Variante B (Ansicht 2)", 10, "f", fmt="#,##0"); col("maxc", "MAX Variante C (Ansicht 2)", 10, "f", fmt="#,##0")
    col("max_fin", "MAX final", 9, "i", fmt="#,##0")
    for v in (1, 2, 3):
        col(f"q{v}", f"A{v} Menge", 9, "f", 1, "#,##0"); col(f"n{v}", f"A{v} Anzahl Gebinde", 8, "f", 1, "#,##0")
        col(f"a{v}", f"A{v} Fläche m²", 8, "f", 1, "0.000"); col(f"w{v}", f"A{v} Wert CHF", 10, "f", 1, "#,##0")
        col(f"b{v}", f"A{v} MAX B (Hilfe)", 8, "f", 2, "#,##0"); col(f"pr{v}", f"A{v} Priorität Verbr./m² (Hilfe)", 9, "f", 2, "#,##0.0")
        col(f"cum{v}", f"A{v} Fläche kumuliert (Hilfe)", 9, "f", 2, "#,##0.0")
        col(f"ant{v}", f"A{v} Anteil im Lift (Hilfe)", 7, "f", 2, "0%")
    for v in (1, 2, 3):
        col(f"loc{v}", f"Lagerort Ansicht {v}", 11, "f")
    col("hint", "Hinweise", 45, "f"); col("ursache", "Ursache (Kennzahlen)", 30); col("aktion", "Aktion (Kennzahlen)", 30)
    col("komm", "Kommentar Diskussion", 30, "i")
    # Hilfsspalten
    for k, h, f in [("bm", "Export Bestandsmenge (G)", "#,##0"), ("ja", "Export Jahresanfang 2025 (T)", "#,##0"),
                    ("bwx", "Export Bestandswert (H)", "#,##0.00"), ("verbx", "Verbrauch Export (Zeitraum)", "#,##0"),
                    ("bem", "Bemerkungen Artikelstamm", None), ("neu", "Neu (1/0)", "0"), ("idx", "Zeile", "0"),
                    ("geb_e", "Gebinde wirksam", None), ("mpg_e", "Menge/Gebinde wirksam", "#,##0"),
                    ("lose", "Lose (1/0)", "0"), ("la", "Fläche lose wirksam m²", "0.000"), ("lift", "Lift wirksam", None),
                    ("grp", "Liftgruppe", None), ("luo", "Ziel-LU Artikel (0=keiner)", "0.0"), ("finf", "MAX final gesetzt", "0"),
                    ("finv", "MAX final Wert", "#,##0"), ("mblg", "MB + LG", "#,##0"), ("istn", "Gebinde bei Ist", "#,##0"),
                    ("finn", "Gebinde bei MAX final", "#,##0"), ("hgt", "Gebindehöhe cm", "0"), ("maxh", "max. Ladehöhe Lift mm", "0")]:
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
    CAP1, CAP2, CAP3 = "Parameter!$K$5", "Parameter!$K$6", "Parameter!$K$7"
    HLT = f"Parameter!$A${HL_FIRST}:$C${HL_LAST}"
    GLT = f"Parameter!$F${GL_FIRST}:$G${GL_LAST}"

    FORM = {
        "ist": '=MAX(0,IF(Parameter!$B$20="Jahresanfang 2025 + Export",@ja@+@bm@,@bm@))',
        "istwert": "=@ist@*@preis@",
        "verb": "=@verbx@*Parameter!$B$18",
        "garea": f'=IF(@lose@=1,0,IFERROR(VLOOKUP(@geb_e@,{GEB_RNG},6,0)+0,0))',
        "lift_vor": (f'=IF(@nr@="","",IF(IFERROR(VLOOKUP(@hl@,{HLT},2,0),"Nein")="2+3",'
                     f'IF(ISNUMBER(SEARCH("grosser Artikel",@bem@)),2,IFERROR(VLOOKUP(@geb_e@,{GLT},2,0),2)),'
                     f'IF(OR(IFERROR(VLOOKUP(@hl@,{HLT},2,0),"Nein")="Nein",IFERROR(VLOOKUP(@hl@,{HLT},2,0),"Nein")=""),"-",'
                     f'VLOOKUP(@hl@,{HLT},2,0))))'),
        "maxb": "=@b2@",
        "maxc": "=MAX(@b2@,@ist@)",
        "hint": ('=IF(@nr@="","",IF(@neu@=1,"Neu; ","")'
                 '&IF(AND(@neu@=0,@bm@<0),"negativer Bestand (Export "&@bm@&"); ","")'
                 '&IF(@verb@<=0,"Kein Verbrauch; ","")'
                 '&IF(@lose@=1,"Lose – Fläche geschätzt; ","")'
                 '&IF(@ist@>@b2@,"Überbestand "&ROUND(@ist@-@b2@,0)&" Stk / CHF "&ROUND((@ist@-@b2@)*@preis@,0)&"; ","")'
                 '&IF(AND(@finf@=1,@finv@<@mblg@),"MAX < MB+LG; ","")'
                 '&IF(AND(@maxh@>0,@hgt@*10>@maxh@),"zu hoch für Tablar; ","")'
                 '&IF(AND(@neu@=0,OR(@finf@=1,@lift_in@<>"",@geb_neu@<>"",@mpg_neu@<>"",@lu_art@<>"")),"manuell übersteuert; ","")'
                 '&IF(AND(@lose@=0,@garea@=0),"Gebinde unbekannt; ","")'
                 '&IF(AND(@lose@=0,N(@mpg_neu@)<=0,N(@mpg@)<=0),"Menge pro Gebinde fehlt (Ist = 1 Gebinde); ","")'
                 '&@bem@)'),
        "geb_e": '=IF(@geb_neu@<>"",UPPER(TRIM(@geb_neu@)),IF(@geb@="","LOSE",@geb@))',
        "mpg_e": '=IF(@geb_e@="LOSE",1,IF(N(@mpg_neu@)>0,@mpg_neu@,IF(N(@mpg@)>0,@mpg@,MAX(1,@ist@))))',
        "lose": '=IF(@geb_e@="LOSE",1,0)',
        "la": f'=IF(@lose@=0,0,IF(N(@lose_in@)>0,@lose_in@,IFERROR(VLOOKUP(@hl@,{HLT},3,0)+0,0)))',
        "lift": '=IF(@lift_in@<>"",@lift_in@,@lift_vor@)',
        "grp": '=IF(@lift@=1,"L1",IF(OR(@lift@=2,@lift@=3),"L23",IF(@lift@="Aussen","Aussen","-")))',
        "luo": "=IF(N(@lu_art@)>0,@lu_art@,0)",
        "finf": "=IF(ISNUMBER(@max_fin@),1,0)",
        "finv": "=N(@max_fin@)",
        "mblg": "=N(@mb@)+N(@lg@)",
        "istn": "=ROUNDUP(@ist@/@mpg_e@,0)",
        "finn": "=ROUNDUP(@finv@/@mpg_e@,0)",
        "hgt": f'=IF(@lose@=1,0,IFERROR(VLOOKUP(@geb_e@,{GEB_RNG},5,0)+0,0))',
        "maxh": "=IFERROR(IF(OR(@lift@=1,@lift@=2,@lift@=3),N(INDEX(Parameter!$I$5:$I$7,@lift@)),0),0)",
    }
    for v in (1, 2, 3):
        var, lu = VAR[v], LUC[v]
        cap = f'IF(@grp@="L1",{CAP1},{CAP2}+{CAP3})'
        FORM.update({
            f"b{v}": f"=ROUNDUP(MAX(@verb@/IF(@luo@>0,@luo@,MAX(0.1,N({lu}))),@mblg@)/@mpg_e@,0)*@mpg_e@",
            f"q{v}": (f'=IF(LEFT({var},1)="A",@ist@,IF(@finf@=1,@finv@,IF(LEFT({var},1)="B",@b{v}@,MAX(@b{v}@,@ist@))))'),
            f"n{v}": f"=IF(@lose@=1,0,ROUNDUP(@q{v}@/@mpg_e@,0))",
            f"a{v}": f"=IF(@lose@=1,IF(@q{v}@<=0,0,@la@*MAX(1,IF(@ist@>0,@q{v}@/@ist@,1))),@n{v}@*@garea@)",
            f"w{v}": f"=@q{v}@*@preis@",
            f"pr{v}": f"=IF(@a{v}@>0,@verb@/@a{v}@,-1)",
            f"cum{v}": (f'=IF(AND(OR(@grp@="L1",@grp@="L23"),@a{v}@>0),SUMIFS(#a{v}#,#grp#,@grp@,#pr{v}#,">"&@pr{v}@)'
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
                        mpg=a["mpg"], ursache=a["ursache"] or None, aktion=a["aktion"] or None,
                        bm=a["bm"], ja=a["ja"], bwx=a["bwx"], verbx=a["verbx"], bem=a["bem"], neu=0)
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
    ws.column_dimensions.group(C["bm"], C["maxh"], outline_level=1, hidden=True)
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.freeze_panes = "C7"
    ws.auto_filter.ref = f"A6:{C['komm']}{RN}"
    dv = DataValidation(type="list", formula1='"1,2,3,Aussen"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{C['lift_in']}{R0}:{C['lift_in']}{R0 + n_art - 1}")
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
    for k in ("ist", "verb", "mblg", "mpg_e", "garea", "lose", "la", "luo", "finf", "finv", "lift", "preis", "istn", "finn"):
        nm = "A_" + k.upper()
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Artikel!${C[k]}${R0}:${C[k]}${RN}")

    # Kontrolle-Formeln (Parameter)
    put(wsP, "M12", f"=SUM({RG('bwx')})", fmt="#,##0.00")
    put(wsP, "M13", f'=SUMIFS({RG("bwx")},{RG("bm")},">0")', fmt="#,##0.00")
    put(wsP, "M14", f"=SUM({RG('istwert')})", fmt="#,##0.00")
    put(wsP, "M15", f'=COUNTIFS({RG("bm")},"<0",{RG("neu")},0)', fmt="0")
    put(wsP, "M16", f'=COUNTIFS({RG("neu")},0)', fmt="0")
    wsP.column_dimensions["M"].width = 14

    # ------------------------------------------------------------------ Cockpit
    ws = wsC
    put(ws, "B1", "Cockpit – Lagerlift-Belegung & MAX-Bestand nach Lagerumschlag", bold=True, size=16, color="1F3864")
    put(ws, "B2", '="Verbrauchsbasis: "&TEXT(MONTH(Parameter!B15),"00")&"."&YEAR(Parameter!B15)&" – "&TEXT(MONTH(Parameter!B16),"00")&"."&YEAR(Parameter!B16)'
                  '&"   ·   Ist-Bestand: "&Parameter!B20&"   ·   Kapazität "&TEXT(Parameter!K8,"0")&" m² / "&Parameter!E8&" Tablare"',
        bold=True, size=11, color="C00000")
    put(ws, "B3", "Gelbe Felder = Regler. Auslastung: grün < 85 %, orange 85–100 %, rot > 100 %.", italic=True, color="595959")
    defaults = {1: (VARIANTS[0], 2), 2: (VARIANTS[2], 2), 3: (VARIANTS[2], 3)}
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
        for j, h in enumerate(["Bereich", "belegt m²", "belegte Tablare", "Kapazität m²", "Auslastung"]):
            put(ws, f"{c[j]}8", h, bold=True, fill=SUB, border=True, align="center", wrap=True)
        A = lambda k: RG(k)
        for i, (lbl, n) in enumerate(lifts_lbl):
            r = 9 + i
            put(ws, f"{c[0]}{r}", lbl, border=True)
            put(ws, f"{c[1]}{r}", f"=SUMIFS({A(f'a{v}')},{A('lift')},{n})", fmt="#,##0.0", border=True)
            put(ws, f"{c[2]}{r}", f"=ROUNDUP({c[1]}{r}/Parameter!$J${4 + int(n)},0)", fmt="0", border=True)
            put(ws, f"{c[3]}{r}", f"=Parameter!$K${4 + int(n)}", fmt="#,##0.0", border=True)
            put(ws, f"{c[4]}{r}", f"=IF({c[3]}{r}>0,{c[1]}{r}/{c[3]}{r},0)", fmt="0%", border=True, bold=True)
        put(ws, f"{c[0]}12", "Lift 2 + 3 gemeinsam (KTL/PAL)", border=True, bold=True)
        put(ws, f"{c[1]}12", f"={c[1]}10+{c[1]}11", fmt="#,##0.0", border=True, bold=True)
        put(ws, f"{c[2]}12", f"=ROUNDUP({c[1]}12/Parameter!$J$6,0)", fmt="0", border=True, bold=True)
        put(ws, f"{c[3]}12", f"={c[3]}10+{c[3]}11", fmt="#,##0.0", border=True, bold=True)
        put(ws, f"{c[4]}12", f"=IF({c[3]}12>0,{c[1]}12/{c[3]}12,0)", fmt="0%", border=True, bold=True)
        put(ws, f"{c[0]}13", "Total 3 Lifte", border=True, bold=True)
        put(ws, f"{c[1]}13", f"={c[1]}9+{c[1]}12", fmt="#,##0.0", border=True, bold=True)
        put(ws, f"{c[2]}13", f"={c[2]}9+{c[2]}10+{c[2]}11", fmt="0", border=True, bold=True)
        put(ws, f"{c[3]}13", f"={c[3]}9+{c[3]}12", fmt="#,##0.0", border=True, bold=True)
        put(ws, f"{c[4]}13", f"=IF({c[3]}13>0,{c[1]}13/{c[3]}13,0)", fmt="0%", border=True, bold=True)
        put(ws, f"{c[0]}14", "Nicht im Lift (andere Hauptlager)", italic=True, border=True)
        put(ws, f"{c[1]}14", f'=SUMIFS({A(f"a{v}")},{A("grp")},"-")', fmt="#,##0.0", border=True, italic=True)
        put(ws, f"{c[0]}15", "Manuell «Aussen» zugeteilt", italic=True, border=True)
        put(ws, f"{c[1]}15", f'=SUMIFS({A(f"a{v}")},{A("grp")},"Aussen")', fmt="#,##0.0", border=True, italic=True)
        pr = f"{c[4]}9:{c[4]}13"
        ws.conditional_formatting.add(pr, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1.5, color="5B9BD5", showValue=True))
        ws.conditional_formatting.add(pr, CellIsRule(operator="greaterThan", formula=["Parameter!$B$12"], fill=RED_F))
        ws.conditional_formatting.add(pr, CellIsRule(operator="greaterThanOrEqual", formula=["Parameter!$B$11"], fill=ORANGE_F))
        ws.conditional_formatting.add(pr, CellIsRule(operator="lessThan", formula=["Parameter!$B$11"], fill=GREEN_F))
        # Passt?
        put(ws, f"{c[0]}17", "Passt alles in die 3 Lifte?", bold=True, size=12)
        ws.merge_cells(f"{c[3]}17:{c[4]}18")
        put(ws, f"{c[3]}17", f'=IF(AND({c[1]}9<={c[3]}9,{c[1]}12<={c[3]}12),"JA","NEIN")', bold=True, size=20, align="center")
        put(ws, f"{c[0]}18", "Lift 1 (schwer, separat)")
        put(ws, f"{c[2]}18", f'=IF({c[1]}9<={c[3]}9,"JA","NEIN")', bold=True, align="center")
        put(ws, f"{c[0]}19", "Lift 2 + 3 gemeinsam")
        put(ws, f"{c[2]}19", f'=IF({c[1]}12<={c[3]}12,"JA","NEIN")', bold=True, align="center")
        ws.merge_cells(f"{c[0]}20:{c[4]}20")
        put(ws, f"{c[0]}20", (f'=IF(OR(AND({c[1]}10>{c[3]}10,{c[1]}11<{c[3]}11),AND({c[1]}11>{c[3]}11,{c[1]}10<{c[3]}10)),'
                              f'"Ausgleich Lift 2 ↔ 3 möglich: "&ROUNDUP(MIN(MAX({c[1]}10-{c[3]}10,{c[1]}11-{c[3]}11),'
                              f'MAX({c[3]}10-{c[1]}10,{c[3]}11-{c[1]}11))/Parameter!$J$6,0)&" Tablare","")'),
            bold=True, color="C55A11")
        for rr in (f"{c[3]}17", f"{c[2]}18", f"{c[2]}19"):
            ws.conditional_formatting.add(rr, CellIsRule(operator="equal", formula=['"JA"'], fill=GREEN_F, font=Font(color="006100", bold=True)))
            ws.conditional_formatting.add(rr, CellIsRule(operator="equal", formula=['"NEIN"'], fill=RED_F, font=Font(color="9C0006", bold=True)))
        # Aussenlager
        put(ws, f"{c[0]}22", "Aussenlager nötig m²", bold=True)
        put(ws, f"{c[2]}22", f"=MAX(0,{c[1]}9-{c[3]}9)+MAX(0,{c[1]}12-{c[3]}12)+{c[1]}15", fmt="#,##0.0", bold=True)
        put(ws, f"{c[0]}23", "Artikel im Aussenlager (ganz/teilweise)")
        put(ws, f"{c[2]}23", f'=COUNTIF({A(f"loc{v}")},"Aussenlager")+COUNTIF({A(f"loc{v}")},"teilweise")', fmt="0")
        # Werte
        put(ws, f"{c[0]}25", "Lagerwert CHF gesamt", bold=True)
        put(ws, f"{c[2]}25", f"=SUM({A(f'w{v}')})", fmt="#,##0", bold=True)
        put(ws, f"{c[0]}26", "   davon in den Liften")
        put(ws, f"{c[2]}26", f"=SUMPRODUCT({A(f'w{v}')},{A(f'ant{v}')})", fmt="#,##0")
        put(ws, f"{c[0]}27", "   davon im Aussenlager")
        put(ws, f"{c[2]}27", f'=SUMIFS({A(f"w{v}")},{A("grp")},"L1")+SUMIFS({A(f"w{v}")},{A("grp")},"L23")-{c[2]}26+SUMIFS({A(f"w{v}")},{A("grp")},"Aussen")', fmt="#,##0")
        put(ws, f"{c[0]}28", "   davon nicht im Lift (andere Lager)")
        put(ws, f"{c[2]}28", f'=SUMIFS({A(f"w{v}")},{A("grp")},"-")', fmt="#,##0")
        put(ws, f"{c[0]}30", "Artikel mit Überbestand (Ist > MAX B)", bold=True)
        put(ws, f"{c[2]}30", f'=SUMPRODUCT(({A("ist")}>{A(f"b{v}")})*1)', fmt="0")
        put(ws, f"{c[0]}31", "Überbestandswert CHF")
        put(ws, f"{c[2]}31", f'=SUMPRODUCT(({A("ist")}>{A(f"b{v}")})*({A("ist")}-{A(f"b{v}")})*{A("preis")})', fmt="#,##0")
        put(ws, f"{c[3]}30", "(LU dieser Ansicht)", italic=True, color="808080")
        # Gebinde vs lose
        for j, h in enumerate(["Fläche je Lift", "Gebinde m²", "lose m² (geschätzt)", "Anteil lose"]):
            put(ws, f"{c[j]}33", h, bold=True, fill=SUB, border=True, wrap=True, align="center")
        for i, (lbl, n) in enumerate(lifts_lbl):
            r = 34 + i
            put(ws, f"{c[0]}{r}", lbl, border=True)
            put(ws, f"{c[1]}{r}", f"=SUMIFS({A(f'a{v}')},{A('lift')},{n},{A('lose')},0)", fmt="#,##0.0", border=True)
            put(ws, f"{c[2]}{r}", f"=SUMIFS({A(f'a{v}')},{A('lift')},{n},{A('lose')},1)", fmt="#,##0.0", border=True)
            put(ws, f"{c[3]}{r}", f"=IF({c[1]}{r}+{c[2]}{r}>0,{c[2]}{r}/({c[1]}{r}+{c[2]}{r}),0)", fmt="0%", border=True)
        put(ws, f"{c[0]}37", "Total", bold=True, border=True)
        put(ws, f"{c[1]}37", f"=SUM({c[1]}34:{c[1]}36)", fmt="#,##0.0", bold=True, border=True)
        put(ws, f"{c[2]}37", f"=SUM({c[2]}34:{c[2]}36)", fmt="#,##0.0", bold=True, border=True)
        put(ws, f"{c[3]}37", f"=IF({c[1]}37+{c[2]}37>0,{c[2]}37/({c[1]}37+{c[2]}37),0)", fmt="0%", bold=True, border=True)
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
          "Aussenlager m²", "Lagerwert CHF", "Δ zum Ist CHF", "Δ zum Ist m²",
          "Lift 1 m²", "Lift 2 m²", "Lift 3 m²", "Aussen manuell m²"]
    for j, h in enumerate(sh):
        put(ws, f"{L(j + 1)}4", h, bold=True, fill=SUB, wrap=True, border=True, align="center")

    def area_expr(G, Q):
        """Fläche je Artikel als Array-Ausdruck. G = Anzahl Gebinde, Q = Menge (für lose)."""
        return f"((1-A_LOSE)*{G}*A_GAREA+A_LOSE*({Q}>0)*A_LA*(1+(A_IST>0)*({Q}>A_IST)*({Q}/(A_IST+(A_IST=0))-1)))"

    def scen_exprs(var, x):
        if var == "Ist":
            G, Q = "A_ISTN", "A_IST"
            V = "A_IST"
        else:
            RAW = f"A_VERB/(A_LUO+(A_LUO=0)*{x})"
            M = f"({RAW}+(A_MBLG>{RAW})*(A_MBLG-{RAW}))"
            NB = f"ROUNDUP({M}/A_MPG_E,0)"
            if var == "B":
                G = f"(A_FINF*A_FINN+(1-A_FINF)*{NB})"
                V = f"(A_FINF*A_FINV+(1-A_FINF)*{NB}*A_MPG_E)"
            else:
                G = f"(A_FINF*A_FINN+(1-A_FINF)*({NB}+(A_ISTN>{NB})*(A_ISTN-{NB})))"
                V = f"(A_FINF*A_FINV+(1-A_FINF)*({NB}*A_MPG_E+(A_IST>{NB}*A_MPG_E)*(A_IST-{NB}*A_MPG_E)))"
            Q = G  # lose: Menge pro Gebinde = 1
        return area_expr(G, Q), V

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
        AR, V = scen_exprs(var, x)
        put(ws, f"M{r}", f"=SUMPRODUCT((A_LIFT=1)*{AR})", fmt="#,##0.0", border=True)
        put(ws, f"N{r}", f"=SUMPRODUCT((A_LIFT=2)*{AR})", fmt="#,##0.0", border=True)
        put(ws, f"O{r}", f"=SUMPRODUCT((A_LIFT=3)*{AR})", fmt="#,##0.0", border=True)
        put(ws, f"P{r}", f'=SUMPRODUCT((A_LIFT="Aussen")*{AR})', fmt="#,##0.0", border=True)
        put(ws, f"C{r}", f"=M{r}/{CAP1}", fmt="0%", border=True)
        put(ws, f"D{r}", f"=N{r}/{CAP2}", fmt="0%", border=True)
        put(ws, f"E{r}", f"=O{r}/{CAP3}", fmt="0%", border=True)
        put(ws, f"F{r}", f"=(N{r}+O{r})/({CAP2}+{CAP3})", fmt="0%", border=True, bold=True)
        put(ws, f"G{r}", f"=(M{r}+N{r}+O{r})/Parameter!$K$8", fmt="0%", border=True, bold=True)
        put(ws, f"H{r}", f'=IF(AND(M{r}<={CAP1},N{r}+O{r}<={CAP2}+{CAP3}),"JA","NEIN")', border=True, align="center", bold=True)
        put(ws, f"I{r}", f"=MAX(0,M{r}-{CAP1})+MAX(0,N{r}+O{r}-{CAP2}-{CAP3})+P{r}", fmt="#,##0.0", border=True)
        put(ws, f"J{r}", f"=SUMPRODUCT({V}*A_PREIS)", fmt="#,##0", border=True)
        put(ws, f"K{r}", f"=J{r}-$J$5", fmt="#,##0;-#,##0", border=True)
        put(ws, f"L{r}", f"=(M{r}+N{r}+O{r}+P{r})-($M$5+$N$5+$O$5+$P$5)", fmt="#,##0.0;-#,##0.0", border=True)
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
        ("Varianten", "b"),
        ("A – Ist-Bestand: heutiger Bestand (unabhängig vom LU).", ""),
        ("B – Ziel-LU: MAX = Jahresverbrauch / Ziel-LU, mindestens MB + Losgrösse, aufgerundet auf ganze Gebinde. Zielzustand nach Abbau des Überbestands.", ""),
        ("C – Ziel-LU oder Ist (Hauptvariante): der höhere Wert aus B und Ist-Bestand. Platzbedarf, solange der Überbestand noch nicht abgebaut ist.", ""),
        ("MAX final (gelb) ersetzt in B und C den berechneten Wert. Ziel-LU Artikel ersetzt den LU der Ansicht für diesen Artikel. Jahresverbrauch = Verbrauch Export × 12 / Anzahl Monate.", ""),
        ("Lift-Auslastung", "b"),
        ("Fläche je Artikel = Anzahl Gebinde × Grundfläche (L × B) – Gebinde nicht gestapelt. Anzahl Gebinde = AUFRUNDEN(Menge / Menge pro Gebinde).", ""),
        ("Lose Artikel: Fläche = Startwert je Hauptlager (Parameter) oder eigener Wert; in B/C skaliert mit Menge / Ist-Menge, mindestens Startwert.", ""),
        ("Kapazität je Lift = Tablare × Tablarbreite × Tablartiefe. Belegte Tablare = AUFRUNDEN(m² / Tablarfläche). Auslastung = m² / Kapazität.", ""),
        ("Zuordnung: LIFT1 → Lift 1. KTL/PAL → Lift 3 (BITO S21–S33, Eurobox S41–S52) oder Lift 2 (S61–S63, Trennbleche S71–S83, lose, «grosser Artikel»). Andere Hauptlager → «Nicht im Lift».", ""),
        ("«Passt?» wird für Lift 1 und Lift 2 + 3 gemeinsam beurteilt. Aussenlager: je Liftgruppe Artikel nach Verbrauch pro m² sortiert (Schnelldreher zuerst), "
         "Fläche kumuliert; was die Kapazität überschreitet = «Aussenlager», der Artikel an der Grenze = «teilweise».", ""),
        ("Hinweise", "b"),
        ("Kein Verbrauch · Lose – Fläche geschätzt · Überbestand (Ist > MAX B, Menge und CHF, LU Ansicht 2) · MAX < MB+LG (MAX final zu tief) · zu hoch für Tablar (Gebindehöhe > max. Ladehöhe) · "
         "manuell übersteuert · Neu · negativer Bestand (Export, als 0 gerechnet) · Gebinde unbekannt · + Bemerkungen aus dem Artikelstamm.", ""),
        ("Annahmen", "b"),
        ("1. Ist-Bestand = Spalte G «Bestandsmenge» des Kennzahlen-Auszugs, negativ = 0. ACHTUNG: G sieht nach Netto-Bewegung seit 01.01.2025 aus (fast die Hälfte negativ; "
         "Jahresanfang 2025 + G ist nie negativ). Umschaltbar in Parameter B20.", ""),
        ("2. Verbrauch = Pivot «Auswertung Verbrauch» Spalte C (Kalenderjahr 2025), negativ = 0; fehlt ein Artikel dort, Spalte Y «Wareneinsatz».", ""),
        ("3. MB = Spalte AK; wenn leer, Blatt «Auswertung MB» (Lager = Hauptlager). Losgrösse = Spalte D.", ""),
        ("4. Gebinde aus «Artikelstamm Umschlag 2», sonst «Artikelstamm mit Verbrauch»; ohne Gebinde = LOSE. Menge pro Gebinde fehlt → ganzer Ist-Bestand = 1 Gebinde.", ""),
        ("5. Startwerte lose: KTL 0.1 m², PAL 0.5 m², LIFT1 3.479 m² (1 Tablar), andere 0.5 m² – je Artikel, unabhängig von der Menge.", ""),
        ("6. Lagerwert = Menge × Preis GLD Aktuell. Hauptlager aus dem Kennzahlen-Auszug (nicht aus dem Artikelstamm).", ""),
        ("7. Artikel ohne Verbrauch: B = MB + LG (sonst 0). Aussenlager-Priorität: Verbrauch pro m² absteigend.", ""),
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
    wb, C = build(D, old)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTPUT_FILE)
    print(f"Gespeichert: {OUTPUT_FILE}")
    if "--pruefen" in sys.argv:
        pruefen(OUTPUT_FILE)


if __name__ == "__main__":
    main()

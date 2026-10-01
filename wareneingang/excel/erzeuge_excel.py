"""Ergänzt die Excel-Liste "Lieferungen zu Vogt" um ein Dashboard aus Formeln.

    python3 erzeuge_excel.py Lieferungen_zu_Vogt.xlsx Lieferungen_zu_Vogt_Dashboard.xlsx

Die BI-Pivot "Pivot Alle Artikel" bleibt unverändert (Aktualisierung beim Öffnen).
Neue Blätter:
  Übersicht    – was jetzt zu Vogt muss, was die nächste Lieferung braucht
  Ergebnisse   – Eingabe der Messergebnisse pro Artikel/Lieferant und Monat (Auswahlliste)
  Status       – lesbarer Verlauf (Zu Vogt / Freipass / i.O. / Abweichung)
  Abweichungen – Details zu Abweichungen (Massabweichung, Stückzahl, Bestellnummer)
  Anleitung    – Regeln und Bedienung
  Berechnung   – ausgeblendet, Regel-Logik

Die bisherigen Zellfarben (grün/rot/gelb) werden einmalig als Ergebnisse übernommen.
Benötigt: openpyxl (nur zum Lesen der Farben).
"""

import re
import sys
import zipfile
from xml.sax.saxutils import escape

import openpyxl

PIVOT = "Pivot Alle Artikel"
P = "'" + PIVOT + "'"
MONATE = [f"{j}-{m:02d}" for j in (2025, 2026, 2027) for m in range(1, 13)]
NM = len(MONATE)
ZEILEN = 150          # Kapazität Artikel/Lieferant-Zeilen
PIVOT_ZEILEN = 600    # so viele Zeilen der Pivot werden nach neuen Artikeln durchsucht
R0 = 3                # erste Datenzeile
R1 = R0 + ZEILEN - 1
ERGEBNISSE = ["i.O.", "Abweichung", "nicht gemessen"]


def spalte(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- Spalten der Berechnung
B_ART, B_LIEF, B_BEZ, B_NAME = 1, 2, 3, 4
B_D = 5                      # geliefert je Monat (NM Spalten)
B_S = B_D + NM               # Zustand vor Monat j (NM + 1 Spalten)
B_O = B_S + NM + 1           # offen je Monat (NM Spalten)
B_FIN = B_O + NM
B_OFFEN, B_IDX, B_MONAT, B_REGEL, B_NAECHSTE, B_MESS = range(B_FIN + 1, B_FIN + 7)
B_F_OFFEN, B_C_OFFEN, B_F_VOGT, B_C_VOGT, B_F_FREI, B_C_FREI = range(B_MESS + 1, B_MESS + 7)
B_F_NEU, B_C_NEU = B_C_FREI + 2, B_C_FREI + 3
E_M = 6                      # erste Monatsspalte in Ergebnisse/Status (F)


def messphase(x):
    return f"OR({x}<3,{x}=20,{x}=40)"


def regel(x):
    return (f'IF({x}<3,"Qualifizierung "&({x}+1)&"/3",IF({x}=20,"Kontrollmessung",IF({x}=40,"Requalifizierung",'
            f'IF({x}<20,"Freipass "&({x}-9)&"/3","Freipass "&({x}-29)&"/3 (2. Runde)"))))')


# Zustände: 0-2 Qualifizierung (n i.O.), 10-12 Freipass Runde 1, 20 Kontrolle, 30-32 Freipass Runde 2, 40 Requali
def naechster(x, d, r):
    return (f'IF({d}=0,{x},IF({r}="Abweichung",0,IF({r}="i.O.",IF({x}<2,{x}+1,IF({x}=2,10,IF({x}=20,30,IF({x}=40,10,{x})))),'
            f'IF({messphase(x)},{x},IF({x}=12,20,IF({x}=32,40,{x}+1))))))')


# ---------------------------------------------------------------- Zellen-Helfer
VIEW0 = '<sheetView workbookViewId="0"/>'


class Blatt:
    def __init__(self):
        self.zeilen = {}
        self.extra = []
        self.cols = []
        self.views = ""

    def setze(self, r, c, inhalt, stil=0, formel=False):
        self.zeilen.setdefault(r, {})[c] = (inhalt, stil, formel)

    def xml(self, sheet_pr=""):
        teile = []
        for r in sorted(self.zeilen):
            zellen = []
            for c in sorted(self.zeilen[r]):
                inhalt, stil, formel = self.zeilen[r][c]
                ref = f"{spalte(c)}{r}"
                s = f' s="{stil}"' if stil else ""
                if inhalt is None:
                    zellen.append(f'<c r="{ref}"{s}/>')
                elif formel:
                    zellen.append(f'<c r="{ref}"{s}><f>{escape(inhalt)}</f></c>')
                elif isinstance(inhalt, (int, float)):
                    zellen.append(f'<c r="{ref}"{s}><v>{inhalt}</v></c>')
                else:
                    zellen.append(f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(inhalt))}</t></is></c>')
            teile.append(f'<row r="{r}">{"".join(zellen)}</row>')
        cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"%s/>' % (a, b, w, ' hidden="1"' if h else "") for a, b, w, h in self.cols)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f'{sheet_pr}<sheetViews>{self.views or VIEW0}</sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                f'{"<cols>" + cols + "</cols>" if cols else ""}<sheetData>{"".join(teile)}</sheetData>'
                f'{"".join(self.extra)}'
                '<pageMargins left="0.7" right="0.7" top="0.78" bottom="0.78" header="0.3" footer="0.3"/></worksheet>')


def ansicht(zelle=None, gewaehlt=False):
    sel = ' tabSelected="1"' if gewaehlt else ""
    if not zelle:
        return f'<sheetView{sel} workbookViewId="0"/>'
    m = re.match(r"([A-Z]+)(\d+)", zelle)
    x = sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(m.group(1)))) - 1
    y = int(m.group(2)) - 1
    pane = "bottomRight" if x and y else ("bottomLeft" if y else "topRight")
    attrs = (f'xSplit="{x}" ' if x else "") + (f'ySplit="{y}" ' if y else "")
    return (f'<sheetView{sel} workbookViewId="0"><pane {attrs}topLeftCell="{zelle}" activePane="{pane}" state="frozen"/>'
            f'<selection pane="{pane}" activeCell="{zelle}" sqref="{zelle}"/></sheetView>')


# ---------------------------------------------------------------- Altbestand aus der Liste lesen
def lese_liste(pfad):
    wb = openpyxl.load_workbook(pfad)
    ws = wb[PIVOT]
    kopf = None
    for row in ws.iter_rows(min_row=1, max_row=20):
        for c in row:
            if c.value == "Artikel ID":
                kopf = c.row
    sp = {str(c.value).strip(): c.column for c in ws[kopf] if c.value}
    monate = {k: v for k, v in sp.items() if re.fullmatch(r"\d{4}-\d{2}", k)}

    def farbe(c):
        f = c.fill
        if not f or f.fill_type != "solid" or f.fgColor.type != "rgb":
            return None
        rgb = f.fgColor.rgb[-6:]
        r, g, b = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
        if r > 200 and g > 200 and b < 120:
            return "nicht gemessen"
        if g > r and g > b:
            return "i.O."
        if r > g and r > b:
            return "Abweichung"
        return None

    artikel, ergebnisse = [], {}
    for r in range(kopf + 1, ws.max_row + 1):
        art = ws.cell(r, sp["Artikel ID"]).value
        lief = ws.cell(r, sp["Lieferant ID"]).value
        if art in (None, "Artikel") or not isinstance(lief, str):
            break
        art, lief = str(art).strip(), lief.strip()
        artikel.append((art, lief, ws.cell(r, sp["Artikelbezeichnung"]).value or "", ws.cell(r, sp["Lieferant Bez."]).value or ""))
        for m, c in monate.items():
            zelle = ws.cell(r, c)
            if zelle.value == "✓" and farbe(zelle):
                ergebnisse[(art, lief, m)] = farbe(zelle)

    # Tabelle "Massabweichung" unter der Pivot
    abweichungen = []
    for row in ws.iter_rows():
        for c in row:
            if c.value == "Massabweichung":
                hdr = {str(x.value).strip(): x.column for x in ws[c.row] if x.value}
                r = c.row + 1
                while ws.cell(r, hdr["Artikel"]).value not in (None, ""):
                    g = lambda n: ws.cell(r, hdr[n]).value
                    datum = g("Datum")
                    abweichungen.append({
                        "artikel": str(g("Artikel")).strip(), "lieferant": str(g("Lieferant") or "").strip(),
                        "monat": datum.strftime("%Y-%m") if hasattr(datum, "strftime") else "",
                        "mass": str(g("Massabweichung") or "").strip(), "stueck": str(g("Stückzahl") or "").strip(),
                        "bestellung": str(g("Bestellnummer") or "").strip(),
                        "bemerkung": "",
                    })
                    r += 1
                break
    return artikel, ergebnisse, abweichungen


# ---------------------------------------------------------------- Stile (werden an styles.xml angehängt)
class Stile:
    def __init__(self, styles_xml):
        self.xml = styles_xml
        self.n = {t: int(re.search(rf"<{t} count=\"(\d+)\"", styles_xml).group(1)) for t in ("fonts", "fills", "borders", "cellXfs", "dxfs")}
        self.neu = {t: [] for t in self.n}

    def _add(self, typ, xml):
        self.neu[typ].append(xml)
        return self.n[typ] + len(self.neu[typ]) - 1

    def font(self, b=False, sz=11, farbe="FF000000", i=False):
        return self._add("fonts", f'<font>{"<b/>" if b else ""}{"<i/>" if i else ""}<sz val="{sz}"/><color rgb="{farbe}"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>')

    def fill(self, rgb):
        return self._add("fills", f'<fill><patternFill patternType="solid"><fgColor rgb="{rgb}"/><bgColor indexed="64"/></patternFill></fill>')

    def border(self):
        s = '<{0} style="thin"><color rgb="FFBFBFBF"/></{0}>'
        return self._add("borders", "<border>" + "".join(s.format(x) for x in ("left", "right", "top", "bottom")) + "<diagonal/></border>")

    def xf(self, font=0, fill=0, border=0, num=0, h=None, v="center", wrap=False):
        hz = ' horizontal="%s"' % h if h else ""
        wr = ' wrapText="1"' if wrap else ""
        al = f'<alignment{hz} vertical="{v}"{wr}/>'
        return self._add("cellXfs", f'<xf numFmtId="{num}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"' +
                                    ("".join(t for t, an in ((' applyFont="1"', font), (' applyFill="1"', fill), (' applyBorder="1"', border), (' applyNumberFormat="1"', num)) if an)) +
                                    f' applyAlignment="1">{al}</xf>')

    def dxf(self, bg, farbe, b=False):
        return self._add("dxfs", f'<dxf><font>{"<b/>" if b else ""}<color rgb="{farbe}"/></font>'
                                 f'<fill><patternFill patternType="solid"><bgColor rgb="{bg}"/></patternFill></fill></dxf>')

    def fertig(self):
        x = self.xml
        for t in ("fonts", "fills", "borders", "cellXfs", "dxfs"):
            if not self.neu[t]:
                continue
            gesamt = self.n[t] + len(self.neu[t])
            x = re.sub(rf'<{t} count="\d+"', f'<{t} count="{gesamt}"', x, count=1)
            x = x.replace(f"</{t}>", "".join(self.neu[t]) + f"</{t}>", 1)
        return x


# ---------------------------------------------------------------- Blätter bauen
def baue(artikel, ergebnisse, abweichungen, st):
    rand = st.border()
    f_weiss = st.font(b=True, farbe="FFFFFFFF")
    f_titel = st.font(b=True, sz=18, farbe="FF1F3864")
    f_fett = st.font(b=True)
    f_grau = st.font(i=True, sz=10, farbe="FF595959")
    f_kpi = st.font(b=True, sz=24, farbe="FF1F3864")
    f_h2 = st.font(b=True, sz=13, farbe="FF1F3864")
    fl_kopf = st.fill("FF1F3864")
    fl_kpi = st.fill("FFF2F2F2")
    fl_rot = st.fill("FFC00000")
    fl_gruen = st.fill("FF00B050")
    fl_blau = st.fill("FF2F5597")

    X = {
        "kopf": st.xf(f_weiss, fl_kopf, rand, h="center", wrap=True),
        "kopf_rot": st.xf(f_weiss, fl_rot, rand, h="left"),
        "kopf_blau": st.xf(f_weiss, fl_blau, rand, h="left"),
        "kopf_gruen": st.xf(f_weiss, fl_gruen, rand, h="left"),
        "titel": st.xf(f_titel),
        "h2": st.xf(f_h2),
        "notiz": st.xf(f_grau, wrap=False),
        "notiz_wrap": st.xf(f_grau, v="top", wrap=True),
        "text": st.xf(0, 0, rand, num=49),
        "zelle": st.xf(0, 0, rand),
        "mitte": st.xf(0, 0, rand, h="center"),
        "fett": st.xf(f_fett, 0, rand),
        "kpi_zahl": st.xf(f_kpi, fl_kpi, rand, h="center"),
        "kpi_text": st.xf(f_fett, fl_kpi, rand, h="center", wrap=True),
        "wrap": st.xf(0, 0, 0, v="top", wrap=True),
    }
    D = {
        "gruen": st.dxf("FF00B050", "FFFFFFFF"),
        "rot": st.dxf("FFC00000", "FFFFFFFF", b=True),
        "gelb": st.dxf("FFFFFF00", "FF000000"),
        "offen": st.dxf("FFF8CBAD", "FFC00000", b=True),
        "frei": st.dxf("FFDDEBF7", "FF1F4E79"),
        "gruen_hell": st.dxf("FFC6EFCE", "FF006100"),
    }
    BC = lambda c, r: f"Berechnung!${spalte(c)}${r}"

    # ---- Ergebnisse (Eingabe)
    erg = Blatt()
    erg.setze(1, 1, "Messergebnisse eintragen: Zelle des Monats wählen → i.O. / Abweichung / nicht gemessen", X["h2"])
    erg.setze(1, E_M, "Orange = geliefert, muss zu Vogt, Ergebnis fehlt · Hellblau = geliefert mit Freipass · Grün/Rot/Gelb = eingetragenes Ergebnis", X["notiz"])
    for c, t in enumerate(["Artikel ID", "Lieferant ID", "Artikelbezeichnung", "Lieferant", "Nächste Lieferung"], start=1):
        erg.setze(2, c, t, X["kopf"])
    for j, m in enumerate(MONATE):
        erg.setze(2, E_M + j, m, X["kopf"])
    for i in range(ZEILEN):
        r = R0 + i
        if i < len(artikel):
            art, lief, bez, name = artikel[i]
            erg.setze(r, 1, art, X["text"]); erg.setze(r, 2, lief, X["text"])
            erg.setze(r, 3, bez, X["zelle"]); erg.setze(r, 4, name, X["zelle"])
        else:
            for c in range(1, 5):
                erg.setze(r, c, None, X["text"])
        erg.setze(r, 5, f'IF($A{r}="","",{BC(B_NAECHSTE, r)})', X["fett"], True)
        for j, m in enumerate(MONATE):
            wert = ergebnisse.get((artikel[i][0], artikel[i][1], m)) if i < len(artikel) else None
            erg.setze(r, E_M + j, wert, X["mitte"])
    bereich = f"{spalte(E_M)}{R0}:{spalte(E_M + NM - 1)}{R1}"
    tl = f"{spalte(E_M)}{R0}"
    d1 = f"Berechnung!{spalte(B_D)}{R0}"
    s1 = f"Berechnung!{spalte(B_S)}{R0}"
    f_offen = escape('AND(%s="",%s>0,%s)' % (tl, d1, messphase(s1)))
    f_frei = escape('AND(%s="",%s>0)' % (tl, d1))
    cf = [
        f'<cfRule type="cellIs" dxfId="{D["gruen"]}" priority="1" operator="equal"><formula>"i.O."</formula></cfRule>',
        f'<cfRule type="cellIs" dxfId="{D["rot"]}" priority="2" operator="equal"><formula>"Abweichung"</formula></cfRule>',
        f'<cfRule type="cellIs" dxfId="{D["gelb"]}" priority="3" operator="equal"><formula>"nicht gemessen"</formula></cfRule>',
        f'<cfRule type="expression" dxfId="{D["offen"]}" priority="4"><formula>{f_offen}</formula></cfRule>',
        f'<cfRule type="expression" dxfId="{D["frei"]}" priority="5"><formula>{f_frei}</formula></cfRule>',
    ]
    erg.extra.append(f'<conditionalFormatting sqref="{bereich}">{"".join(cf)}</conditionalFormatting>')
    liste = ",".join(ERGEBNISSE)
    erg.extra.append(f'<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" '
                     f'errorTitle="Ungültig" error="Bitte i.O., Abweichung oder nicht gemessen wählen." sqref="{bereich}">'
                     f'<formula1>"{liste}"</formula1></dataValidation></dataValidations>')
    erg.cols = [(1, 2, 11, False), (3, 3, 30, False), (4, 4, 24, False), (5, 5, 30, False), (E_M, E_M + NM - 1, 13, False)]
    erg.views = ansicht(f"{spalte(E_M)}{R0}")

    # ---- Berechnung (ausgeblendet)
    ber = Blatt()
    ber.setze(1, 1, "Hilfsblatt für die Regel-Logik – bitte nicht ändern")
    for j, m in enumerate(MONATE):
        ber.setze(2, B_D + j, m)
    for i in range(ZEILEN):
        r = R0 + i
        ber.setze(r, B_ART, f'IF(Ergebnisse!$A{r}="","",TRIM(Ergebnisse!$A{r}&""))', 0, True)
        ber.setze(r, B_LIEF, f'IF(Ergebnisse!$B{r}="","",TRIM(Ergebnisse!$B{r}&""))', 0, True)
        ber.setze(r, B_BEZ, f'IF(Ergebnisse!$C{r}="","",Ergebnisse!$C{r})', 0, True)
        ber.setze(r, B_NAME, f'IF(Ergebnisse!$D{r}="","",Ergebnisse!$D{r})', 0, True)
        a = f"${spalte(B_ART)}{r}"
        for j in range(NM):
            hdr = f"{spalte(B_D + j)}$2"
            ber.setze(r, B_D + j,
                      f'IF({a}="",0,IFERROR(COUNTIFS({P}!$B$3:$B$2000,{a},{P}!$E$3:$E$2000,${spalte(B_LIEF)}{r},'
                      f'INDEX({P}!$A$3:$ZZ$2000,0,MATCH({hdr},{P}!$A$2:$ZZ$2,0)),"✓"),0))', 0, True)
        ber.setze(r, B_S, 0)
        for j in range(NM):
            x = f"{spalte(B_S + j)}{r}"
            d = f"{spalte(B_D + j)}{r}"
            e = f"Ergebnisse!{spalte(E_M + j)}{r}"
            ber.setze(r, B_S + j + 1, naechster(x, d, e), 0, True)
            ber.setze(r, B_O + j, f'IF(AND({d}>0,{e}="",{messphase(x)}),1,0)', 0, True)
        fin = f"{spalte(B_FIN)}{r}"
        ober = f"{spalte(B_O)}{r}:{spalte(B_O + NM - 1)}{r}"
        ber.setze(r, B_FIN, f"{spalte(B_S + NM)}{r}", 0, True)
        ber.setze(r, B_OFFEN, f"SUM({ober})", 0, True)
        ber.setze(r, B_IDX, f'IFERROR(MATCH(1,{ober},0),"")', 0, True)
        idx = f"{spalte(B_IDX)}{r}"
        ber.setze(r, B_MONAT, f'IF({idx}="","",INDEX(${spalte(B_D)}$2:${spalte(B_D + NM - 1)}$2,{idx}))', 0, True)
        sx = f"INDEX({spalte(B_S)}{r}:{spalte(B_S + NM - 1)}{r},{idx})"
        ber.setze(r, B_REGEL, f'IF({idx}="","",{regel(sx)})', 0, True)
        ber.setze(r, B_NAECHSTE, f'IF({a}="","",IF({messphase(fin)},"Zu Vogt – ","")&{regel(fin)})', 0, True)
        ber.setze(r, B_MESS, f'AND({a}<>"",{messphase(fin)})', 0, True)
        for flag, cnt, bed in ((B_F_OFFEN, B_C_OFFEN, f'AND({a}<>"",{spalte(B_OFFEN)}{r}>0)'),
                               (B_F_VOGT, B_C_VOGT, f"{spalte(B_MESS)}{r}"),
                               (B_F_FREI, B_C_FREI, f'AND({a}<>"",NOT({spalte(B_MESS)}{r}))')):
            ber.setze(r, flag, bed, 0, True)
            ber.setze(r, cnt, f'IF({spalte(flag)}{r},COUNTIF(${spalte(flag)}${R0}:{spalte(flag)}{r},TRUE),"")', 0, True)
    for r in range(3, PIVOT_ZEILEN + 1):
        bed = (f'IF(ISTEXT({P}!$E{r}),IF(ISTEXT({P}!$B{r}),IF({P}!$B{r}<>"Artikel",'
               f'COUNTIFS(Ergebnisse!$A${R0}:$A${R1},{P}!$B{r},Ergebnisse!$B${R0}:$B${R1},{P}!$E{r})=0,FALSE),FALSE),FALSE)')
        ber.setze(r, B_F_NEU, bed, 0, True)
        ber.setze(r, B_C_NEU, f'IF({spalte(B_F_NEU)}{r},COUNTIF(${spalte(B_F_NEU)}$3:{spalte(B_F_NEU)}{r},TRUE),"")', 0, True)

    # ---- Status (Verlauf lesbar)
    sta = Blatt()
    sta.setze(1, 1, "Verlauf pro Artikel und Lieferant (wird berechnet – Ergebnisse im Blatt «Ergebnisse» eintragen)", X["h2"])
    for c, t in enumerate(["Artikel ID", "Lieferant ID", "Artikelbezeichnung", "Lieferant", "Nächste Lieferung"], start=1):
        sta.setze(2, c, t, X["kopf"])
    for j, m in enumerate(MONATE):
        sta.setze(2, E_M + j, m, X["kopf"])
    for i in range(ZEILEN):
        r = R0 + i
        for c in range(1, 5):
            sta.setze(r, c, f'IF(Ergebnisse!{spalte(c)}{r}="","",Ergebnisse!{spalte(c)}{r})', X["zelle"], True)
        sta.setze(r, 5, f"Ergebnisse!E{r}", X["fett"], True)
        for j in range(NM):
            d = f"Berechnung!{spalte(B_D + j)}{r}"
            x = f"Berechnung!{spalte(B_S + j)}{r}"
            e = f"Ergebnisse!{spalte(E_M + j)}{r}"
            sta.setze(r, E_M + j,
                      f'IF({d}=0,"",IF({e}="Abweichung","Abweichung",IF({messphase(x)},'
                      f'IF({e}="i.O.","i.O.",IF({e}="nicht gemessen","nicht gemessen","ZU VOGT")),'
                      f'IF({e}="i.O.","i.O. (zusätzl.)","Freipass"))))', X["mitte"], True)
    sb = f"{spalte(E_M)}{R0}:{spalte(E_M + NM - 1)}{R1}"
    tl = f"{spalte(E_M)}{R0}"
    f_io = escape('LEFT(%s,4)="i.O."' % tl)
    cf = [
        f'<cfRule type="cellIs" dxfId="{D["offen"]}" priority="1" operator="equal"><formula>"ZU VOGT"</formula></cfRule>',
        f'<cfRule type="cellIs" dxfId="{D["rot"]}" priority="2" operator="equal"><formula>"Abweichung"</formula></cfRule>',
        f'<cfRule type="cellIs" dxfId="{D["gelb"]}" priority="3" operator="equal"><formula>"nicht gemessen"</formula></cfRule>',
        f'<cfRule type="cellIs" dxfId="{D["frei"]}" priority="4" operator="equal"><formula>"Freipass"</formula></cfRule>',
        f'<cfRule type="beginsWith" dxfId="{D["gruen"]}" priority="5" operator="beginsWith" text="i.O."><formula>{f_io}</formula></cfRule>',
    ]
    sta.extra.append(f'<conditionalFormatting sqref="{sb}">{"".join(cf)}</conditionalFormatting>')
    sta.cols = [(1, 2, 11, False), (3, 3, 30, False), (4, 4, 24, False), (5, 5, 30, False), (E_M, E_M + NM - 1, 13, False)]
    sta.views = ansicht(f"{spalte(E_M)}{R0}")

    # ---- Übersicht
    ueb = Blatt()
    ueb.setze(1, 1, "Wareneingang · Stichprobenmessung Vogt", X["titel"])
    ueb.setze(2, 1, "Die BI-Liste aktualisiert sich beim Öffnen. Ergebnisse von Vogt im Blatt «Ergebnisse» eintragen – diese Übersicht rechnet sich selbst.", X["notiz"])
    kpis = [
        (1, "Jetzt zu Vogt bringen\n(geliefert, Ergebnis fehlt)", f"COUNTIF(Berechnung!${spalte(B_F_OFFEN)}${R0}:${spalte(B_F_OFFEN)}${R1},TRUE)"),
        (6, "Artikel: nächste Lieferung\nmuss zu Vogt", f"COUNTIF(Berechnung!${spalte(B_F_VOGT)}${R0}:${spalte(B_F_VOGT)}${R1},TRUE)"),
        (11, "Artikel: nächste Lieferung\nhat Freipass", f"COUNTIF(Berechnung!${spalte(B_F_FREI)}${R0}:${spalte(B_F_FREI)}${R1},TRUE)"),
        (16, "Neue Artikel in der BI-Liste\n(fehlen in «Ergebnisse»)", f"COUNTIF(Berechnung!${spalte(B_F_NEU)}$3:${spalte(B_F_NEU)}${PIVOT_ZEILEN},TRUE)"),
    ]
    merges = ["A1:T1", "A2:T2"]
    for c, text, formel in kpis:
        ueb.setze(4, c, formel, X["kpi_zahl"], True)
        ueb.setze(5, c, text, X["kpi_text"])
        for cc in range(c + 1, c + 4):
            ueb.setze(4, cc, None, X["kpi_zahl"]); ueb.setze(5, cc, None, X["kpi_text"])
        merges += [f"{spalte(c)}4:{spalte(c + 3)}4", f"{spalte(c)}5:{spalte(c + 3)}5"]

    abschnitte = [
        (1, "kopf_rot", "JETZT ZU VOGT BRINGEN", ["Artikel", "Bezeichnung", "Lieferant", "Geliefert", "Regel"], B_C_OFFEN,
         [B_ART, B_BEZ, B_NAME, B_MONAT, B_REGEL]),
        (7, "kopf_rot", "NÄCHSTE LIEFERUNG → ZU VOGT", ["Artikel", "Bezeichnung", "Lieferant", "Grund"], B_C_VOGT,
         [B_ART, B_BEZ, B_NAME, B_NAECHSTE]),
        (12, "kopf_blau", "NÄCHSTE LIEFERUNG → FREIPASS", ["Artikel", "Bezeichnung", "Lieferant", "Freipass"], B_C_FREI,
         [B_ART, B_BEZ, B_NAME, B_NAECHSTE]),
    ]
    LR = 9
    for c0, stil, titel, koepfe, cnt, quellen in abschnitte:
        ueb.setze(7, c0, titel, X[stil])
        for k in range(1, len(koepfe)):
            ueb.setze(7, c0 + k, None, X[stil])
        merges.append(f"{spalte(c0)}7:{spalte(c0 + len(koepfe) - 1)}7")
        for k, t in enumerate(koepfe):
            ueb.setze(8, c0 + k, t, X["kopf"])
        for n in range(ZEILEN):
            r = LR + n
            for k, q in enumerate(quellen):
                ueb.setze(r, c0 + k,
                          f'IFERROR(INDEX(Berechnung!${spalte(q)}${R0}:${spalte(q)}${R1},MATCH({n + 1},Berechnung!${spalte(cnt)}${R0}:${spalte(cnt)}${R1},0)),"")',
                          X["zelle"], True)
    # Neue Artikel
    c0 = 17
    ueb.setze(7, c0, "NEU IN DER BI-LISTE → in «Ergebnisse» ergänzen", X["kopf_gruen"])
    for k in range(1, 4):
        ueb.setze(7, c0 + k, None, X["kopf_gruen"])
    merges.append(f"{spalte(c0)}7:{spalte(c0 + 3)}7")
    for k, (t, pc) in enumerate([("Artikel ID", "B"), ("Bezeichnung", "C"), ("Lieferant ID", "E"), ("Lieferant", "F")]):
        ueb.setze(8, c0 + k, t, X["kopf"])
        for n in range(40):
            r = LR + n
            ueb.setze(r, c0 + k,
                      f'IFERROR(INDEX({P}!${pc}$3:${pc}${PIVOT_ZEILEN},MATCH({n + 1},Berechnung!${spalte(B_C_NEU)}$3:${spalte(B_C_NEU)}${PIVOT_ZEILEN},0)),"")',
                      X["zelle"], True)
    ueb.extra.append(f'<mergeCells count="{len(merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + "</mergeCells>")
    ueb.cols = [(1, 1, 10, False), (2, 2, 28, False), (3, 3, 22, False), (4, 4, 10, False), (5, 5, 18, False), (6, 6, 2, False),
                (7, 7, 10, False), (8, 8, 28, False), (9, 9, 22, False), (10, 10, 28, False), (11, 11, 2, False),
                (12, 12, 10, False), (13, 13, 28, False), (14, 14, 22, False), (15, 15, 26, False), (16, 16, 2, False),
                (17, 17, 11, False), (18, 18, 28, False), (19, 19, 11, False), (20, 20, 22, False)]
    ueb.views = ansicht("A9", gewaehlt=True)

    # ---- Abweichungen
    abw = Blatt()
    abw.setze(1, 1, "Abweichungen (Details zum Messprotokoll) – bei jeder Abweichung hier eine Zeile ergänzen", X["h2"])
    kopf = ["Artikel", "Lieferant", "Geliefert (JJJJ-MM)", "Massabweichung", "Stückzahl n.i.O.", "Bestellnummer", "Bemerkung"]
    for c, t in enumerate(kopf, start=1):
        abw.setze(2, c, t, X["kopf"])
    for n in range(200):
        a = abweichungen[n] if n < len(abweichungen) else None
        werte = [a["artikel"], a["lieferant"], a["monat"], a["mass"], a["stueck"], a["bestellung"], a["bemerkung"]] if a else [None] * 7
        for c, w in enumerate(werte, start=1):
            abw.setze(3 + n, c, w or None, X["text"] if c in (1, 3, 6) else X["zelle"])
    abw.cols = [(1, 1, 11, False), (2, 2, 24, False), (3, 3, 18, False), (4, 4, 18, False), (5, 5, 16, False), (6, 6, 15, False), (7, 7, 60, False)]
    abw.views = ansicht("A3")

    # ---- Anleitung
    anl = Blatt()
    texte = [
        ("titel", "Anleitung"),
        ("h2", "Was zeigt diese Datei?"),
        ("wrap", "Die Pivot «Pivot Alle Artikel» kommt aus dem BI und aktualisiert sich beim Öffnen. Jedes ✓ ist eine Lieferung (Artikel × Lieferant × Monat). "
                 "Aus den eingetragenen Messergebnissen berechnet die Datei für jede Lieferung, ob sie zu Vogt muss oder einen Freipass hat."),
        ("h2", "Täglich"),
        ("wrap", "1. Blatt «Übersicht» öffnen: links steht, was jetzt zu Vogt muss, rechts was die nächste Lieferung jedes Artikels braucht."),
        ("wrap", "2. Kommt das Ergebnis von Vogt: im Blatt «Ergebnisse» in der Zeile des Artikels/Lieferanten die orange Monatszelle anklicken und i.O. oder Abweichung wählen."),
        ("wrap", "3. Bei einer Abweichung zusätzlich im Blatt «Abweichungen» eine Zeile mit Massabweichung, Stückzahl und Bestellnummer ergänzen."),
        ("wrap", "4. Ging eine Lieferung, die zu Vogt gemusst hätte, nicht zu Vogt: «nicht gemessen» wählen."),
        ("wrap", "5. Erscheint unter «Neu in der BI-Liste» ein Artikel: im Blatt «Ergebnisse» unten Artikel ID, Lieferant ID, Bezeichnung und Lieferant eintragen (genau wie in der Pivot)."),
        ("h2", "Regeln (pro Artikel + Lieferant)"),
        ("wrap", "• Qualifizierung: jede Lieferung geht zu Vogt, bis 3 Messungen in Folge i.O. sind. Wurde ein Artikel noch nie bei Vogt gemessen, muss die nächste Lieferung zu Vogt."),
        ("wrap", "• Danach 3 Lieferungen mit Freipass."),
        ("wrap", "• Die 4. Lieferung geht zur Kontrollmessung zu Vogt. i.O. → nochmals 3 Freipässe."),
        ("wrap", "• Danach Requalifizierung bei Vogt. i.O. → wieder 3 Freipässe, Kontrollmessung usw."),
        ("wrap", "• Jede Abweichung (über 0.002 mm im Massprotokoll) – auch bei einer Freipass-Lieferung – führt zurück zur Qualifizierung (3× zu Vogt)."),
        ("wrap", "• Eine zusätzliche Messung i.O. während der Freipass-Phase verbraucht keinen Freipass."),
        ("wrap", "• Fehlt das Ergebnis einer Messung noch, bleiben auch die folgenden Lieferungen dieses Artikels «zu Vogt»."),
        ("h2", "Farben im Blatt «Ergebnisse»"),
        ("wrap", "Orange = geliefert, muss zu Vogt, Ergebnis fehlt · Hellblau = geliefert mit Freipass · Grün = i.O. · Rot = Abweichung · Gelb = nicht gemessen"),
        ("h2", "Hinweise"),
        ("wrap", "Die bisherigen Zellfarben der Pivot wurden einmalig als Ergebnisse übernommen. Farben in der Pivot werden nicht mehr ausgewertet. "
                 "Das Blatt «Berechnung» ist ausgeblendet und enthält die Regel-Logik."),
    ]
    for i, (stil, t) in enumerate(texte, start=1):
        anl.setze(i, 1, t, X[stil])
    anl.cols = [(1, 1, 130, False)]

    return {"Übersicht": ueb, "Ergebnisse": erg, "Status": sta, "Abweichungen": abw, "Anleitung": anl, "Berechnung": ber}


# ---------------------------------------------------------------- In die Original-Datei einfügen
def einfuegen(quelle, ziel):
    artikel, ergebnisse, abweichungen = lese_liste(quelle)
    zin = zipfile.ZipFile(quelle)
    teile = {n: zin.read(n) for n in zin.namelist()}
    st = Stile(teile["xl/styles.xml"].decode("utf8"))
    blaetter = baue(artikel, ergebnisse, abweichungen, st)
    teile["xl/styles.xml"] = st.fertig().encode("utf8")

    wb = teile["xl/workbook.xml"].decode("utf8")
    rels = teile["xl/_rels/workbook.xml.rels"].decode("utf8")
    ct = teile["[Content_Types].xml"].decode("utf8")
    if "Übersicht" in wb:
        raise SystemExit("Die Datei enthält bereits ein Dashboard.")
    sheet_ids = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)]
    rel_ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    nr = max(int(x) for x in re.findall(r"worksheets/sheet(\d+)\.xml", rels))
    neue_sheets = []
    tab_farben = {"Übersicht": "FF1F3864", "Ergebnisse": "FF00B050", "Status": "FF2F5597", "Abweichungen": "FFC00000", "Anleitung": "FF7F7F7F"}
    for i, (name, blatt) in enumerate(blaetter.items(), start=1):
        nr += 1
        rid = f"rId{max(rel_ids) + i}"
        pfad = f"xl/worksheets/sheet{nr}.xml"
        pr = f'<sheetPr><tabColor rgb="{tab_farben[name]}"/></sheetPr>' if name in tab_farben else ""
        teile[pfad] = blatt.xml(pr).encode("utf8")
        rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{nr}.xml"/></Relationships>')
        ct = ct.replace("</Types>", f'<Override PartName="/{pfad}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        state = ' state="hidden"' if name == "Berechnung" else ""
        neue_sheets.append(f'<sheet name="{name}" sheetId="{max(sheet_ids) + i}"{state} r:id="{rid}"/>')
    # Übersicht & Co. vor die Pivot, Berechnung ans Ende
    vorne = "".join(s for s in neue_sheets if 'name="Berechnung"' not in s)
    hinten = "".join(s for s in neue_sheets if 'name="Berechnung"' in s)
    wb = wb.replace("<sheets>", "<sheets>" + vorne, 1).replace("</sheets>", hinten + "</sheets>", 1)
    wb = re.sub(r"<calcPr([^/]*)/>", lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="\d"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wb, count=1)
    teile["xl/workbook.xml"] = wb.encode("utf8")
    teile["xl/_rels/workbook.xml.rels"] = rels.encode("utf8")
    teile["[Content_Types].xml"] = ct.encode("utf8")
    # Die Pivot ist nicht mehr das ausgewählte Blatt, die Übersicht öffnet zuerst
    for n in zin.namelist():
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n):
            teile[n] = teile[n].decode("utf8").replace(' tabSelected="1"', "").encode("utf8")

    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, teile[n])
        for n in teile:
            if n not in zin.namelist():
                zout.writestr(n, teile[n])
    print(f"{len(artikel)} Artikel/Lieferanten, {len(ergebnisse)} Ergebnisse aus Farben, {len(abweichungen)} Abweichungen → {ziel}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    einfuegen(sys.argv[1], sys.argv[2])

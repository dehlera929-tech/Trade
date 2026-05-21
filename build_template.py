"""Erstellt die Excel-Bewertungstabelle fuer Jobinterviews (Vertrieb MSR / Gebaeudeautomation)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()

# ----------------------------- Styles -----------------------------
THIN = Side(border_style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_TITLE = PatternFill("solid", fgColor="1F4E78")
FILL_CAT = PatternFill("solid", fgColor="2E75B6")
FILL_HEAD = PatternFill("solid", fgColor="BDD7EE")
FILL_INFO = PatternFill("solid", fgColor="FFF2CC")
FILL_NOTE = PatternFill("solid", fgColor="E2EFDA")
FILL_TOTAL = PatternFill("solid", fgColor="FFE699")

FONT_TITLE = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
FONT_CAT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
FONT_HEAD = Font(name="Calibri", size=11, bold=True, color="1F4E78")
FONT_LABEL = Font(name="Calibri", size=11, bold=True)
FONT_BODY = Font(name="Calibri", size=11)

WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")

# ------------------------- Sheet 1: Bewertung -------------------------
ws = wb.active
ws.title = "Bewertungsbogen"

# Spaltenbreiten
widths = {"A": 5, "B": 48, "C": 42, "D": 42, "E": 14, "F": 28}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

# Titel
ws.merge_cells("A1:F1")
c = ws["A1"]
c.value = "Bewertungsbogen Jobinterview - Vertrieb MSR / Gebaeudeautomation"
c.font = FONT_TITLE
c.fill = FILL_TITLE
c.alignment = CENTER
ws.row_dimensions[1].height = 30

# Kopfdaten
ws.merge_cells("A2:F2")
ws["A2"].value = "Stammdaten Kandidat / Interview"
ws["A2"].font = FONT_CAT
ws["A2"].fill = FILL_CAT
ws["A2"].alignment = CENTER

stamm = [
    ("Name Kandidat", ""),
    ("Standort / Niederlassung", ""),
    ("Position / Rolle", "Vertriebsingenieur Gebaeudeautomation (MSR)"),
    ("Datum Interview", ""),
    ("Dauer (ca. 60 Min.)", ""),
    ("Interviewer 1", ""),
    ("Interviewer 2", ""),
    ("Lebenslauf gepruegt am", ""),
]
row = 3
for label, val in stamm:
    ws.cell(row=row, column=1, value="").fill = FILL_HEAD
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=2)
    a = ws.cell(row=row, column=2, value=label); a.font = FONT_LABEL; a.fill = FILL_HEAD; a.alignment = WRAP; a.border = BORDER
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
    b = ws.cell(row=row, column=3, value=val); b.font = FONT_BODY; b.alignment = WRAP; b.border = BORDER
    row += 1

# Hinweis Bewertungsskala
row += 1
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
c = ws.cell(row=row, column=1, value="Bewertungsskala: 1 = sehr schwach  |  2 = schwach  |  3 = mittel  |  4 = gut  |  5 = sehr gut   (n.b. = nicht bewertbar)")
c.font = Font(italic=True, size=10); c.fill = FILL_INFO; c.alignment = CENTER
row += 2

# Kategorien-Header
def cat_header(title):
    global row
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = ws.cell(row=row, column=1, value=title)
    c.font = FONT_CAT; c.fill = FILL_CAT; c.alignment = CENTER
    ws.row_dimensions[row].height = 22
    row += 1
    # Spaltenkoepfe
    heads = ["Nr.", "Frage", "Erwartete / Musterantwort (Anhaltspunkt)", "Antwort Kandidat", "Bewertung 1-5", "Bemerkung"]
    for i, h in enumerate(heads, start=1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = FONT_HEAD; cell.fill = FILL_HEAD; cell.alignment = CENTER; cell.border = BORDER
    ws.row_dimensions[row].height = 24
    row += 1

def add_q(nr, frage, muster, hoehe=70):
    global row
    vals = [nr, frage, muster, "", "", ""]
    for i, v in enumerate(vals, start=1):
        cell = ws.cell(row=row, column=i, value=v)
        cell.font = FONT_BODY
        cell.alignment = WRAP if i != 5 else CENTER
        cell.border = BORDER
    ws.row_dimensions[row].height = hoehe
    row += 1

def add_note_row(text="Sonstige Bemerkungen zur Kategorie:"):
    global row
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    c = ws.cell(row=row, column=1, value=text); c.font = FONT_LABEL; c.fill = FILL_NOTE; c.alignment = WRAP; c.border = BORDER
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
    c2 = ws.cell(row=row, column=3, value=""); c2.fill = FILL_NOTE; c2.alignment = WRAP; c2.border = BORDER
    ws.row_dimensions[row].height = 50
    row += 2

# --------- Kategorie 1: Markt- und Branchenkenntnis ---------
cat_header("1. Markt- und Branchenkenntnis / Netzwerk")
add_q(1, "Wie gut kennen Sie den Markt der Gebaeudeautomation in unserer Region? Welche Wettbewerber sehen Sie als die staerksten?",
      "Nennt z.B. Siemens, Johnson Controls, Honeywell, ABB, Schneider Electric, Sauter, Kieback&Peter, WAGO, Beckhoff, Priva, Loytec. Kann regionale Player benennen.")
add_q(2, "Welche Fachplaner (TGA-Planer, Ingenieurbueros) kennen Sie persoenlich? Zu welchen haben Sie aktive Kontakte?",
      "Konkrete Bueronamen + Ansprechpartner. Hinweis auf laufende Projekte / Ausschreibungen / Pflichtenhefte.")
add_q(3, "Welche Generalunternehmer (GU) und Anlagenbauer kennen Sie? Wo bestehen belastbare Kontakte?",
      "Nennt GUs aus HKLS / Elektro / Hochbau. Idealerweise Kontakte zu Einkauf, PL und Obermonteur.")
add_q(4, "Welche Endkunden / Betreiber haben Sie aktuell im Portfolio? Bei welchen koennten Sie sofort 'Tueren oeffnen'?",
      "Konkrete Kundennamen, Branche, Volumen, letzter Kontakt. Mind. 5-10 aktive Kunden.")
add_q(5, "Wie viele Kundenbesuche / Termine machen Sie aktuell pro Woche? Verhaeltnis Neukunde / Bestandskunde?",
      "Realistisch: 8-15 Aussentermine/Woche bei aktivem Aussendienst; Mix ca. 30% Neu / 70% Bestand.")
add_q(6, "Was waren Ihre 3 groessten Vertriebserfolge der letzten 24 Monate (Projekt, Auftragswert in EUR, Ihre Rolle)?",
      "Konkrete Projekte mit Volumen (>100 TEUR aufwaerts wuenschenswert), klarer Eigenanteil, kein Team-Schmuecken.")
add_q(7, "Wie erfolgreich haben Sie bisher Cross-Selling zu angrenzenden Gewerken betrieben (z.B. HLK, Elektro, Sicherheit, Beleuchtung, Energiemonitoring)?",
      "Beispiele: GA + Beleuchtungssteuerung (DALI), GA + Zutritt, GA + Submetering, GA + PV/Lastmanagement. Quote / Umsatzanteil.")
add_q(8, "Wie ist Ihre Pipeline heute aufgestellt? Volumen, Anzahl, Wahrscheinlichkeiten?",
      "Sollte Pipeline-Wert min. 3-5x Jahresziel kennen, Forecast-Genauigkeit erlaeutern koennen.")
add_q(9, "Welche Verbaende / Netzwerke nutzen Sie aktiv (BACnet Interest Group, KNX Association, VDI-GA, BTGA, BACI, DENEFF, etc.)?",
      "Mitgliedschaft + tatsaechliche Aktivitaet (Veranstaltungen, Arbeitskreise).")
add_note_row()

# --------- Kategorie 2: Kenntnis Johnson Controls / Wettbewerb ---------
cat_header("2. Kenntnis Johnson Controls & Wettbewerbsumfeld")
add_q(10, "Was wissen Sie ueber Johnson Controls? Welche Produkte/Loesungen verbinden Sie mit uns?",
      "Metasys (GMS), Facility Explorer, FX/FEC Controller, NAE/SNE, OpenBlue, York (Kaelte), Sabroe, Brandmeldung (TYCO/Autocall), Sicherheit. Globale Aufstellung.")
add_q(11, "Wo sehen Sie die Staerken und wo die Schwaechen von Johnson Controls im Vergleich zu z.B. Siemens Desigo oder Honeywell?",
      "Differenziert: Staerken (Komplettanbieter HVAC+GA+Security+Service, OpenBlue-Plattform), Schwaechen (Preis, Marktanteil DACH ggue. Siemens).")
add_q(12, "Warum bewerben Sie sich gerade bei Johnson Controls? Was reizt Sie?",
      "Konkrete, recherchierte Antwort - nicht 'grosser Konzern'. Bezug zu Produktportfolio, Karriereperspektive, Kultur.")
add_note_row()

# --------- Kategorie 3: Technische Kompetenz GA ---------
cat_header("3. Technische Kompetenz Gebaeudeautomation")
add_q(13, "Wie wuerden Sie einem Kunden ohne Vorkenntnisse in 3 Minuten erklaeren, was Gebaeudeautomation ist und welchen Nutzen sie bringt?",
      "Ebenen-Modell (Feld-/Automations-/Managementebene), Nutzen: Energieeffizienz, Komfort, Betrieb/Service, ESG-Reporting, vorausschauende Wartung. Bildhafte Sprache.")
add_q(14, "Welche Gewerke grenzen an die GA an und muessen integriert werden?",
      "HLK (Heizung/Lueftung/Klima/Kaelte), Elektro/Beleuchtung (DALI/KNX), Sanitaer, Sicherheit (BMA/EMA/ZuKo/Video), Aufzug, PV / Speicher / Ladeinfrastruktur, Submetering.")
add_q(15, "Sind Sie in der Lage, beim Kunden eigenstaendig eine GA-Loesung zu entwerfen (Funktionsumfang, Topologie, grobe Mengen, Schnittstellen)?",
      "Bejaht mit Beispiel: Aufnahme Bedarf -> Funktionsliste -> Topologie (Feld/Auto/Mgmt) -> Schnittstellen -> grobe Kalkulation. Kennt GA-Funktionslisten nach VDI 3814.")
add_q(16, "Welche Anlagenarten haben Sie bisher automatisiert? (RLT, Heizzentrale, Kaelte, Einzelraum, Beleuchtung, Verschattung ...)",
      "Konkrete Anlagen, Groessenklassen, eigene Verantwortung bei der Projektierung / im Verkauf.")
add_q(17, "Welche Trends sehen Sie aktuell in der GA (z.B. IP-Technik, Cloud/Edge, KI, Cybersecurity, ESG, GEG, Smart Readiness Indicator)?",
      "Nennt: BACnet/SC, IP-Migration, OT-Security (IEC 62443), Cloud-Analytics, KI fuer Optimierung, ESG/CSRD-Reporting, GEG, EU-Taxonomie, SRI, Digitaler Zwilling.")
add_q(18, "Welche Besonderheiten / Schwierigkeiten gibt es typischerweise in GA-Projekten?",
      "Schnittstellen, Gewerkegrenzen GA/Elektro, unklare Pflichtenhefte, IBN unter Zeitdruck, Inbetriebnahme-Reihenfolge, Cybersecurity, Datenpunktlisten, Abnahme nach VDI 6041.")
add_note_row()

# --------- Kategorie 4: Kommunikationsprotokolle ---------
cat_header("4. Kommunikationsprotokolle (Kurzerklaerung im Glossar-Blatt)")
add_q(19, "BACnet IP / BACnet MS/TP - wo wuerden Sie was einsetzen? Was sind die Unterschiede?",
      "BACnet/IP: Backbone, Automationsebene, hoehere Bandbreite. MS/TP: Feldbus, RS-485, guenstig, langsamer (bis ~76,8 kBit/s), Token-Passing. Heute Trend zu BACnet/IP und BACnet/SC.")
add_q(20, "KNX - wann ist es das Mittel der Wahl, wann nicht?",
      "Mittel der Wahl: Raumautomation, Bueros, Wohnbau, Beleuchtung/Verschattung/Einzelraumregelung. Schwaechen: Primaeranlagen / RLT-Steuerung, grosse Datenmengen.")
add_q(21, "Modbus RTU / Modbus TCP - typische Anwendungen?",
      "Anbindung Drittgeraete: Zaehler, Frequenzumrichter, Kaeltemaschinen, PV-Wechselrichter, USV. RTU seriell, TCP ueber Ethernet. Kein objektorientiertes Modell wie BACnet.")
add_q(22, "M-Bus / wireless M-Bus - wofuer?",
      "Verbrauchszaehler (Waerme, Wasser, Gas, Strom-Subzaehler), Submetering, Heizkostenabrechnung (HKVO/EED).")
add_q(23, "LoRaWAN - wo macht es Sinn?",
      "Funkbasiertes LPWAN, batteriebetriebene Sensorik, grosse Reichweite, geringer Datendurchsatz. Sinnvoll fuer Submetering, Raumklima, Belegung, Liegenschaftsmonitoring ohne Verkabelung.")
add_q(24, "DALI / DALI-2 - Anwendungsfeld?",
      "Beleuchtungssteuerung (Dimmen, Szenen, Tageslicht), DALI-2 mit erweitertem Geraetemodell inkl. Sensoren/Taster, Diagnose, HCL.")
add_q(25, "EnOcean - was ist die Idee?",
      "Funk-Sensorik mit Energy Harvesting (batterielos), z.B. Taster, Fensterkontakte, Praesenz. Nachruest- und sanierungsfreundlich.")
add_q(26, "OPC UA - welche Rolle spielt es in der GA?",
      "Herstellerneutrale Kommunikation, Industrie 4.0, semantische Informationsmodelle, sichere (zertifikatsbasierte) Anbindung Mgmt-Ebene / IT / Cloud / MES.")
add_q(27, "BACnet/SC, IEC 62443 - Stichwort Cybersecurity in der GA - was bedeutet das fuer den Vertrieb?",
      "BACnet Secure Connect (TLS-basiert, IT-freundlich, keine UDP/Broadcasts). IEC 62443 als Security-Norm fuer OT. Vertrieblich relevant: IT-Abteilung des Kunden frueh einbinden.")
add_note_row()

# --------- Kategorie 5: Normen / Regelwerk ---------
cat_header("5. Normen und Regelwerk (siehe Glossar)")
add_q(28, "VDI 3814 - worum geht es und wofuer brauchen wir das im Verkauf?",
      "GA-Norm: Planung, Funktionslisten, BACS, AMEV-Bezug. Vertrieblich Basis fuer LV-Erstellung, Datenpunktlisten, Mengenermittlung.")
add_q(29, "VDI 6026 - was steht da drin?",
      "Dokumentation in der TGA: Inhalt und Umfang der Planungs- und Ausfuehrungsunterlagen. Wichtig fuer Abnahme und Uebergabe.")
add_q(30, "Weitere relevante Normen / Richtlinien?",
      "VDI 3813 (Raumautomation), VDI 6041 (Inspektion GA), DIN EN ISO 16484 (GA-Reihe), DIN EN ISO 52120 (Energieeffizienz GA, frueher EN 15232), GEG, AMEV BACS, DIN V 18599.")
add_q(31, "DIN EN ISO 52120 / EN 15232 - was sind die GA-Effizienzklassen und wofuer nutzen Sie sie im Verkauf?",
      "Klassen D (nicht energieeffizient) bis A (hochenergieeffizient). Argument fuer Investitionen, Foerdermittel (BEG/BAFA), ESG/CSRD, GEG-Nachweise.")
add_note_row()

# --------- Kategorie 6: Vertrieb & Kalkulation ---------
cat_header("6. Vertrieb, Kalkulation und Verhandlung")
add_q(32, "Wie kalkulieren Sie ein GA-Angebot? Welche Bestandteile / Stundensaetze / Materialaufschlaege?",
      "Material (Controller, I/O, Feld, Schaltschrank) + Engineering (Hardware-/Software-Planung) + IBN + Schulung + Service. Kennt Kalkulationsschema, Deckungsbeitrag, Riskozuschlag.")
add_q(33, "Wie gehen Sie mit Nachtraegen und Aenderungen waehrend der Bauphase um?",
      "Klares Nachtragsmanagement, Anzeige nach VOB/B 2 Nr. 5/6, Beauftragung VOR Ausfuehrung, Dokumentation.")
add_q(34, "Wie ueberzeugen Sie einen preisgetriebenen GU davon, nicht den billigsten Anbieter zu nehmen?",
      "TCO-Argumentation, Energieeinsparung, Servicekosten, Folgekosten, Termin- und Inbetriebnahme-Sicherheit, Referenzen, Schnittstellenrisiko.")
add_q(35, "Wie sieht Ihr typischer Sales-Prozess aus (vom Lead bis zur Auftragsbestaetigung)?",
      "Strukturiertes Vorgehen: Lead-Qualifizierung -> Bedarfsaufnahme -> Loesungsentwurf -> Angebot -> Verhandlung -> Auftrag. Tools: CRM (Salesforce/SAP/HubSpot), Forecast.")
add_q(36, "Wie gehen Sie mit langen Entscheidungszyklen und mehreren Stakeholdern (Investor, Planer, GU, Betreiber, IT, Einkauf) um?",
      "Stakeholder-Mapping, mehrgleisige Bearbeitung, individuelle Nutzenargumentation, Multi-Level-Selling.")
add_note_row()

# --------- Kategorie 7: Persoenliches / Konditionen ---------
cat_header("7. Persoenliches, Konditionen, Rahmen")
add_q(37, "Aktuelles Gehalt (fix / variabel / Sonderzahlungen)?", "Konkrete Angabe.", 40)
add_q(38, "Gehaltsvorstellung (fix / variabel)?", "Konkrete Angabe; Bandbreite akzeptabel.", 40)
add_q(39, "Aktuelle Kuendigungsfrist / frueheste Verfuegbarkeit?", "Datum / Monatsfrist.", 40)
add_q(40, "Urlaubsanspruch heute / Erwartung?", "Tage.", 40)
add_q(41, "Firmenwagen: aktuelle Regelung, private Nutzung, Wunschmodell / -klasse?", "Klasse, private Nutzung ja/nein, 1%-Regel.", 40)
add_q(42, "Wohnort und Reisebereitschaft (Tage/Woche, Uebernachtungen)?", "PLZ, Pendelbereitschaft, Home-Office-Anteil.", 40)
add_q(43, "Sprachkenntnisse (Deutsch, Englisch, weitere)?", "Niveau A1-C2 / Muttersprache.", 40)
add_q(44, "Sonstige Kenntnisse / Zertifikate (BACnet, KNX-Partner, ITK, IT, Projektmanagement)?", "Konkrete Zertifikate mit Datum.", 50)
add_q(45, "Was sind Ihre persoenlichen Ziele in den naechsten 3-5 Jahren?", "Realistische Karriere- und Entwicklungsperspektive.", 50)
add_q(46, "Warum verlassen Sie Ihren aktuellen Arbeitgeber?", "Plausible, nicht-negative Begruendung.", 50)
add_note_row()

# --------- Gesamtbewertung ---------
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
c = ws.cell(row=row, column=1, value="Gesamtbewertung & Empfehlung")
c.font = FONT_CAT; c.fill = FILL_CAT; c.alignment = CENTER
ws.row_dimensions[row].height = 22
row += 1

def kv_row(label, value="", height=30, fill=None):
    global row
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    a = ws.cell(row=row, column=1, value=label); a.font = FONT_LABEL; a.alignment = WRAP; a.border = BORDER
    if fill: a.fill = fill
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=6)
    b = ws.cell(row=row, column=3, value=value); b.font = FONT_BODY; b.alignment = WRAP; b.border = BORDER
    if fill: b.fill = fill
    ws.row_dimensions[row].height = height
    row += 1

kv_row("Staerken", "", 60, FILL_NOTE)
kv_row("Schwaechen / Risiken", "", 60, FILL_NOTE)
kv_row("Offene Fragen / naechste Schritte", "", 50, FILL_NOTE)
kv_row("Durchschnittsbewertung (1-5)", "", 30, FILL_TOTAL)
kv_row("Empfehlung", "Einstellen / 2. Gespraech / Ablehnen", 30, FILL_TOTAL)
kv_row("Unterschrift Interviewer", "", 30)

# Datenvalidierung 1-5 in Spalte E
dv = DataValidation(type="list", formula1='"1,2,3,4,5,n.b."', allow_blank=True)
dv.error = "Bitte 1-5 oder n.b. eingeben"
dv.prompt = "Bewertung 1 (sehr schwach) bis 5 (sehr gut)"
ws.add_data_validation(dv)
# Bereich: alle Zeilen in Spalte E ab Zeile 3
dv.add(f"E3:E{row}")

# Freeze
ws.freeze_panes = "A3"

# ------------------------- Sheet 2: Glossar -------------------------
ws2 = wb.create_sheet("Glossar Protokolle & Normen")
for col, w in {"A": 22, "B": 70, "C": 50}.items():
    ws2.column_dimensions[col].width = w

ws2.merge_cells("A1:C1")
c = ws2["A1"]; c.value = "Glossar - Kommunikationsprotokolle, Normen, Begriffe"
c.font = FONT_TITLE; c.fill = FILL_TITLE; c.alignment = CENTER
ws2.row_dimensions[1].height = 28

def gloss_head(title):
    r = ws2.max_row + 2
    ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws2.cell(row=r, column=1, value=title); c.font = FONT_CAT; c.fill = FILL_CAT; c.alignment = CENTER
    r += 1
    for i, h in enumerate(["Begriff", "Kurzbeschreibung", "Typische Anwendung"], start=1):
        cc = ws2.cell(row=r, column=i, value=h); cc.font = FONT_HEAD; cc.fill = FILL_HEAD; cc.alignment = CENTER; cc.border = BORDER

def gloss_row(begriff, beschr, anwendung):
    r = ws2.max_row + 1
    for i, v in enumerate([begriff, beschr, anwendung], start=1):
        c = ws2.cell(row=r, column=i, value=v); c.font = FONT_BODY; c.alignment = WRAP; c.border = BORDER
    ws2.row_dimensions[r].height = 55

gloss_head("Kommunikationsprotokolle")
gloss_row("BACnet/IP",
          "Standardprotokoll der Gebaeudeautomation ueber Ethernet/IP. Objektorientiert, herstellerneutral (ISO 16484-5).",
          "Backbone-Kommunikation zwischen Automationsstationen, GLT/Managementebene, Integration Drittgewerke.")
gloss_row("BACnet MS/TP",
          "BACnet ueber RS-485 mit Master-Slave/Token-Passing. Bis ~76,8 kBit/s. Kostenguenstig, aber langsamer als IP.",
          "Feldbus zu Reglern, VAV-Boxen, Raumcontrollern, einfache Sensoren/Aktoren.")
gloss_row("BACnet/SC",
          "BACnet Secure Connect: TLS-verschluesselt, hub-basiert ueber TCP, IT-konform (kein UDP-Broadcast).",
          "Zukunftsfaehige, cybersichere Variante - bevorzugt in Neubauten mit IT-Anforderungen.")
gloss_row("KNX",
          "Internationaler Standard (ISO/IEC 14543) fuer Heim- und Gebaeudeautomation. TP (Twisted Pair), IP, RF, PL.",
          "Raumautomation: Beleuchtung, Verschattung, Einzelraumregelung, Taster, Praesenz - vor allem Buero/Wohnen.")
gloss_row("Modbus RTU / TCP",
          "Einfaches Industrieprotokoll, Register-basiert. RTU seriell (RS-485), TCP ueber Ethernet. Kein objektorientiertes Modell.",
          "Anbindung Zaehler, FU, Kaeltemaschinen, PV-Wechselrichter, USV - haeufig als Drittsystem-Schnittstelle.")
gloss_row("M-Bus / wM-Bus",
          "Meter-Bus nach EN 13757 (kabelgebunden) bzw. wireless M-Bus (Funk).",
          "Verbrauchszaehler (Waerme, Wasser, Gas, Strom-Sub), Submetering, Heizkostenabrechnung.")
gloss_row("LoRaWAN",
          "Low Power Wide Area Network. Funk, sehr stromsparend, grosse Reichweite, geringe Bandbreite.",
          "Batteriebetriebene Sensorik, Submetering, Raumklima, Belegungserkennung, Liegenschaftsmonitoring ohne Verkabelung.")
gloss_row("DALI / DALI-2",
          "Digital Addressable Lighting Interface (IEC 62386). DALI-2 erweitert um Sensoren, Eingabegeraete, Diagnose, HCL.",
          "Beleuchtungssteuerung: Dimmen, Szenen, Tageslicht-/Praesenzregelung, Human Centric Lighting.")
gloss_row("EnOcean",
          "Funkstandard mit Energy Harvesting (batterielos). Sub-GHz, ISO/IEC 14543-3-1x.",
          "Nachruestbare Taster, Fensterkontakte, Praesenz - ideal fuer Bestand/Sanierung.")
gloss_row("OPC UA",
          "Plattformunabhaengiges Industrie-4.0-Protokoll mit semantischen Informationsmodellen, zertifikatsbasierte Sicherheit.",
          "IT/OT-Bruecke, Anbindung an MES/ERP, Cloud, Digitaler Zwilling, herstellerneutrale Mgmt-Ebene.")
gloss_row("MQTT",
          "Leichtgewichtiges Publish/Subscribe-Protokoll ueber TCP. Sehr verbreitet im IoT.",
          "Cloud-Anbindung, IoT-Sensorik, Edge-zu-Cloud-Telemetrie.")
gloss_row("LonWorks / LON",
          "Aelteres dezentrales Feldbussystem (ISO/IEC 14908). Verliert Marktanteile an BACnet/KNX.",
          "Bestandsanlagen, einzelne Hersteller (z.B. Sauter, Siemens-Bestand).")

gloss_head("Normen und Richtlinien")
gloss_row("VDI 3814",
          "Leitnorm Gebaeudeautomation: Begriffe, Planung, Funktionslisten, Datenpunkte. Bezug zu DIN EN ISO 16484.",
          "Grundlage fuer LV-Erstellung, Datenpunktlisten, Mengenermittlung und Pflichtenheft.")
gloss_row("VDI 3813",
          "Raumautomation: Funktionen, Anwendungsregeln, Bedienung.",
          "Planung und Beschreibung der Raumautomation (Beleuchtung, Verschattung, Klima, Praesenz).")
gloss_row("VDI 6041",
          "Inspektion und Funktionspruefung von Gebaeudeautomations- und -managementsystemen.",
          "Abnahme und wiederkehrende Pruefung von GA-Anlagen.")
gloss_row("VDI 6026",
          "Dokumentation in der TGA - Anforderungen an Inhalt und Form der Planungs- und Ausfuehrungsdokumentation.",
          "Uebergabedokumentation, Bestandsunterlagen, Abnahmegrundlage.")
gloss_row("DIN EN ISO 16484",
          "Internationale Normenreihe Gebaeudeautomation (u.a. Teil 5: BACnet, Teil 3: Funktionen).",
          "Technische Grundlage des BACnet-Standards und der GA-Funktionsbeschreibungen.")
gloss_row("DIN EN ISO 52120 (frueher EN 15232)",
          "Energetische Bewertung von Gebaeuden - Einfluss der GA. Effizienzklassen A (hoch) bis D (gering).",
          "Argumentation Energieeinsparung, GEG-Nachweis, Foerderung (BEG/BAFA), ESG/CSRD-Reporting.")
gloss_row("DIN V 18599",
          "Energetische Bewertung von Gebaeuden (Heizung, Kuehlung, Lueftung, Beleuchtung, Warmwasser).",
          "Grundlage Energieausweis, GEG-Berechnung; Schnittstelle zu GA-Effizienzklassen.")
gloss_row("GEG",
          "Gebaeudeenergiegesetz - vereint EnEV, EnEG, EEWaermeG. Effizienzanforderungen an Neubau / Bestand.",
          "Argumentationsbasis fuer GA-Investitionen, Effizienzklasse-Nachweis.")
gloss_row("AMEV BACS / AMEV-Empfehlungen",
          "Empfehlungen der Arbeitsgemeinschaft fuer technische Anlagen oeffentlicher Bauherren.",
          "Verbindliche Grundlage bei oeffentlichen Auftraggebern (Bund, Land, Kommune).")
gloss_row("IEC 62443",
          "Normenreihe Cybersecurity fuer industrielle Automation (OT).",
          "Sicherheits-Anforderungen an GA-Systeme, Argumentation gegenueber IT-Abteilung des Kunden.")
gloss_row("SRI",
          "Smart Readiness Indicator der EU - bewertet 'Smartness' von Gebaeuden.",
          "Vertriebs-Argument fuer GA-Ausstattung und ESG-/Taxonomie-Bezug.")
gloss_row("VOB/B",
          "Vergabe- und Vertragsordnung fuer Bauleistungen Teil B - Vertragsrecht im Bau.",
          "Nachtragsmanagement, Behinderungsanzeige, Abnahme.")

ws2.freeze_panes = "A2"

# ------------------------- Speichern -------------------------
out = "/home/user/Trade/Bewertungsbogen_Interview_Vertrieb_GA.xlsx"
wb.save(out)
print("OK:", out)

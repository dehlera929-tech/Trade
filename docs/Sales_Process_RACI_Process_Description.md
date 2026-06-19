# Sales & Calculation Process — Prozessbeschreibung / Process Description
### RACI / Stakeholder-Matrix (Current State)

---

## Dokumentenlenkung / Document Control

| Feld / Field | Wert / Value |
|---|---|
| Titel / Title | Sales & Calculation Process — Prozessbeschreibung (RACI / Stakeholder-Matrix) |
| Mitgeltende Unterlage / Reference document | `Sales_Process_RACI_Matrix_1.xlsx` (Blatt / sheet: *RACI Matrix*) |
| Sprache / Language | Deutsch / English (zweisprachig / bilingual) |
| Status | **TBD — in Abstimmung / to be discussed** |
| Stand / Version | Entwurf / Draft — 2026-06-19 |
| Geltungsbereich / Scope | Vertriebs- und Kalkulationsprozess, Abteilung Installation Systems |

> **Hinweis zum Status / Note on status:** Der Freigabestatus dieses Dokuments ist **TBD** (to be discussed). Das Attribut *Standardisiert (Y/N)* je Prozessschritt zeigt an, ob der jeweilige Schritt bereits standardisiert ist (**Y**) oder noch zu standardisieren ist (**N**).
> The approval status of this document is **TBD** (to be discussed). The *Standardized (Y/N)* attribute per step indicates whether the step is already standardized (**Y**) or still to be standardized (**N**).

---

## 1. Zweck und Geltungsbereich / Purpose and Scope

**DE:** Dieses Dokument beschreibt den Vertriebs- und Kalkulationsprozess (Sales & Calculation Process) von der Identifikation einer Opportunity bis zum Projektabschluss und zur Nachbetrachtung. Es überführt die beiliegende RACI-Matrix in eine verständliche Prozess- und Arbeitsanweisung und ist so gefasst, dass auch eine mit der Thematik nicht vorbefasste dritte Person die Aufgaben, Verantwortlichkeiten und Schnittstellen je Prozessschritt nachvollziehen kann. Maßgeblich sind die in der Excel-Tabelle verwendeten Begriffe, Rollen und Prozessschritte.

**EN:** This document describes the Sales & Calculation Process from identifying an opportunity through project closure and post-mortem review. It translates the attached RACI matrix into a comprehensible process and work instruction, written so that a third party not previously familiar with the subject can follow the tasks, responsibilities and interfaces of each process step. The terms, roles and process steps used in the Excel table are authoritative.

---

## 2. Begriffe und Abkürzungen / Definitions and Abbreviations

### 2.1 RACI-Legende / RACI legend

| Code | Deutsch | English |
|---|---|---|
| **R** | Responsible — führt die Aufgabe durch | Responsible — performs the task |
| **A** | Accountable — trägt die (End-)Verantwortung | Accountable — holds final accountability |
| **C** | Consulted — wird konsultiert / wirkt mit | Consulted — is consulted / contributes |
| **I** | Informed — wird informiert | Informed — is kept informed |

### 2.2 Rollen / Roles (Begriffe der Excel-Tabelle / terms from the Excel table)

| Rolle / Role | Erläuterung / Explanation |
|---|---|
| **Seller** | Vertriebsmitarbeiter; führt die Opportunity durch den Prozess. / Sales representative; drives the opportunity through the process. |
| **Sales Manager** | Vertriebsleitung; gesamtverantwortlich für den Prozess. / Sales management; overall accountable for the process. |
| **Estimator** | Kalkulator; verantwortet Kalkulation und Angebotseinholung. / Estimator; owns calculation and quote sourcing. |
| **Estimate Manager** | Leitung Kalkulation; finale Verantwortung für Kalkulationsergebnisse. / Head of estimation; final accountability for calculation results. |
| **Offshore Engineer** | Technische Zuarbeit aus dem Offshore-Team. / Technical input from the offshore team. |
| **Post-Sales Engineer** (POST Sales Engineer) | Technische Beratung mit Blick auf die spätere Ausführung. / Technical advisory with a view to later execution. |
| **Procurement** | Einkauf. / Procurement. |
| **Project Manager / Ops Manager** | Projekt-/Betriebsleitung. / Project / operations management. |
| **Subcontractors** | Subunternehmen. / Subcontractors. |
| **Service Seller** | Vertrieb Service; etablierte Beziehungen zu Service-Subunternehmen. / Service sales; established relationships with service subcontractors. |
| **Service TL** | Service Team Lead. / Service team lead. |
| **JCI Global Products** | Internes Produktteam für JCI-Equipment und Konditionen. / Internal product team for JCI equipment and conditions. |
| **3rd Party Products** | Drittanbieter-Produkte. / Third-party products. |
| **Legal** | Rechtsabteilung. / Legal department. |
| **Finance** | Finanzabteilung. / Finance department. |
| **EHS** | Arbeits-, Gesundheits- und Umweltschutz. / Environment, Health & Safety. |
| **Customer** | Kunde; trägt die externe Vergabeentscheidung. / Customer; holds the external award decision. |

### 2.3 Werkzeuge / Tools

| Tool | Erläuterung / Explanation |
|---|---|
| **SFDC** | Salesforce.com (CRM / Case-Management). |
| **SharePoint** | Gemeinsame Ablage für Opportunity-Ordner und Dateien. / Shared storage for the opportunity folder and files. |
| **THF** | Kalkulations-/Vorlagenwerkzeug der Kalkulation. / Estimation calculation/template tool. |
| **Excel** | Tabellenkalkulation für Kalkulation und Prüfformulare. / Spreadsheet for calculation and review forms. |
| **Email / Phone / Teams** | Kommunikationsmittel. / Communication means. |
| **Site Visit** | Vor-Ort-Begehung. / On-site visit. |
| **C. Portal** | Vergabeportal des Kunden / öffentliche Ausschreibung. / Customer / public tender portal. |
| **Oracle Webtools** | Buchung im EIC-System. / Booking in the EIC system. |

### 2.4 Weitere Begriffe / Further terms

| Begriff / Term | Erläuterung / Explanation |
|---|---|
| **LV / Gap-Datei** | Leistungsverzeichnis (Bill of Quantities) bzw. Gap-Datei. / Bill of quantities (LV) or gap file. |
| **Handover-Liste** | Übergabeliste/-checkliste mit den vom Seller benötigten technischen Angaben (Hardware, Technik). / Handover list/checklist with the technical inputs required by the Seller (hardware, technology). |
| **QARF** | Standardisiertes Dokument zu Margenerwartung und Lebenszyklus des Gesamtprojekts. / Standardized document covering margin expectation and lifecycle of the overall project. |
| **DoA** | Delegation of Authority — Freigabeprozess gemäß Wertgrenzen. / Delegation of Authority — approval process based on value thresholds. |
| **Brief-it** | DoA-Freigabe-Briefing für Projekte > 1 Mio. USD. / DoA approval briefing for projects > USD 1 million. |
| **4I-Prinzip / 4-eyes** | Vier-Augen-Prüfung des Angebots vor Abgabe. / Four-eyes review of the quote before submission. |
| **MBE** | Minority Business Enterprise — erforderlich, wenn nicht mind. zwei Subunternehmen vorliegen. / Minority Business Enterprise — required if at least two subcontractors are not available. |
| **Prime Revenue** | Lieferanten-/Zahlungsprogramm, dem alle eingesetzten Subunternehmen folgen müssen. / Supplier/payment program that all engaged subcontractors must follow. |
| **EIC-System** | Buchungssystem für den Projektabschluss. / Booking system for project closure. |

### 2.5 Zeitkonvention / Time convention

**DE:** Die Attribute *Time (W)* und *Time (C)* stammen aus der Matrix. *W* bezeichnet die geplante **Durchlaufzeit** (Lead Time), *C* den geplanten **Bearbeitungsaufwand** (Effort). Mehrwertige Angaben (z. B. „350k 3days / 1m 10days") differenzieren nach Projektwert; „service" und „install" trennen Service- und Installationsanteil.

**EN:** The attributes *Time (W)* and *Time (C)* are taken from the matrix. *W* denotes the planned **lead time**, *C* the planned **effort**. Multi-value entries (e.g. "350k 3days / 1m 10days") differentiate by project value; "service" and "install" separate the service and installation portions.

---

## 3. Prozessübersicht / Process Overview

| # | Prozessschritt / Process Step | Standardisiert / Standardized |
|---|---|---|
| 1 | Start | N |
| 2 | Go / No-Go | Y |
| 3 | Create | Y |
| 4 | Kickoff | Y |
| 5 | Technical Solution | N |
| 6 | Commercial | N |
| 7 | Labor Calculation | N |
| 8 | Subcontracting | Y |
| 9 | JCI Equipment | N |
| 10 | 3rd Party Equipment | N |
| 11 | Service | Y |
| 12 | OPS Review | Y |
| 13 | QARF | Y |
| 14 | DoA (Delegation of Authority) | Y |
| 15 | Present Quote | Y |
| 16 | Negotiate | Y |
| 17 | Close (Booking) | Y |
| 18 | Post Mortem | N |

---

## 4. Prozessschritte im Detail / Detailed Process Steps

---

### 1. Start

**Standardisiert / Standardized:** N (noch zu standardisieren / to be standardized) · **Durchlaufzeit (W):** — · **Aufwand (C):** 2hrs · **Tools:** Email, Phone, Site Visit

**DE:** Der Seller recherchiert laufend Vergabeportale und identifiziert Opportunities; dieser Schritt bildet den Einstieg in den Vertriebsprozess. Die Durchführung verantwortet der Seller (R), die Gesamtverantwortung für den Prozess trägt der Sales Manager (A). Der Estimator wird über identifizierte Opportunities informiert (I).
*Zielzustand:* Künftig identifiziert und qualifiziert ein dediziertes Team zunächst Leads bzw. Opportunities und übergibt diese an den Seller.

**EN:** The Seller continuously screens tender portals and identifies opportunities; this step is the entry point of the sales process. The Seller is responsible for execution (R); the Sales Manager is accountable for the process as a whole (A). The Estimator is informed of identified opportunities (I).
*Future state:* A dedicated team will first identify and qualify leads/opportunities and then hand them over to the Seller.

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | — | Estimator |

---

### 2. Go / No-Go

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** 5 days · **Aufwand (C):** 6hrs · **Tools:** Email, Teams

**DE:** Der Seller startet das Quality Gate gemäß den Vorgaben der Vertriebsleitung und holt die Freigabe zur Weiterverfolgung der Opportunity ein (R). Der Sales Manager bleibt federführend verantwortlich (A). Der Estimator unterstützt die Bewertung; bei Bedarf wird zusätzlich Legal hinzugezogen (C). Bezugspunkt der Entscheidung ist die spätere Vergabe durch den Kunden.

**EN:** The Seller initiates the quality gate in line with sales-management guidelines and obtains approval to pursue the opportunity (R). The Sales Manager remains accountable (A). The Estimator supports the assessment; where required, Legal is additionally consulted (C). The decision relates to the customer's eventual award.

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | Estimator, Legal | — |

---

### 3. Create

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** — · **Aufwand (C):** 2hrs · **Tools:** SharePoint

**DE:** Nach positivem Go/No-Go erstellt der Seller den Case in Salesforce.com (SFDC) und übergibt ihn an das Presales-Team (R/A). Zeitgleich legt er den Opportunity-Ordner im gemeinsamen SharePoint an und hinterlegt dort sämtliche erforderlichen Informationen und Dateien — insbesondere das Leistungsverzeichnis (LV) bzw. die Gap-Datei, die Go/No-Go-Präsentation und, maßgeblich, die Handover-Liste. Legal wird informiert (I).

**EN:** After a positive Go/No-Go, the Seller creates the case in Salesforce.com (SFDC) and hands it to the pre-sales team (R/A). In parallel, the Seller sets up the opportunity folder in the shared SharePoint and stores all required information and files there — in particular the bill of quantities (LV) / gap file, the Go/No-Go presentation and, critically, the handover list. Legal is informed (I).

| R/A | C | I |
|---|---|---|
| Seller (R + A) | — | Legal |

---

### 4. Kickoff

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** 48hrs · **Aufwand (C):** 1hr · **Tools:** SFDC, Phone, Teams

**DE:** Innerhalb von 48 Stunden nach Eingang des Cases lädt der Estimator zum Kickoff-Call ein und führt diesen durch (R). Seller (A) und Estimator nehmen verbindlich teil. Procurement und der Project/Ops Manager wirken mit, werden jedoch nur optional eingeladen (C); der Offshore Engineer wird informiert (I).

**EN:** Within 48 hours of receiving the case, the Estimator convenes and runs the kickoff call (R). Seller (A) and Estimator attend on a mandatory basis. Procurement and the Project/Ops Manager contribute but are invited on an optional basis (C); the Offshore Engineer is informed (I).

| R | A | C | I |
|---|---|---|---|
| Estimator | Seller | Procurement, Project/Ops Manager | Offshore Engineer |

---

### 5. Technical Solution

**Standardisiert / Standardized:** N (noch zu standardisieren / to be standardized) · **Durchlaufzeit (W):** 3 days · **Aufwand (C):** 12hrs · **Tools:** THF, Excel

**DE:** Der Estimator empfiehlt die preislich günstigste und technisch am besten geeignete Lösung bzw. Lösungsvarianten (R). Die finale Entscheidung über das Vorgehen trifft der Seller (A). Post-Sales Engineer und Offshore Engineer werden konsultiert (C). Grundlage ist die Handover-Liste/-Checkliste aus dem Schritt „Create", in der der Seller die gewünschte Hardware und Technik festgelegt hat.

**EN:** The Estimator recommends the most cost-effective and technically most suitable solution(s) (R). The Seller makes the final decision on the approach (A). The Post-Sales Engineer and Offshore Engineer are consulted (C). The basis is the handover list/checklist from the "Create" step, in which the Seller specified the desired hardware and technology.

| R | A | C | I |
|---|---|---|---|
| Estimator | Seller | Post-Sales Engineer, Offshore Engineer | — |

---

### 6. Commercial

**Standardisiert / Standardized:** N (noch zu standardisieren / to be standardized) · **Durchlaufzeit (W):** — · **Aufwand (C):** — · **Tools:** —

**DE:** Der Seller prüft die Vertragsunterlagen — kommerzielle und technische Anhänge — auf valide Punkte hinsichtlich Risiken und Chancen (R). Bei den kommerziellen Unterlagen sind insbesondere zu prüfen: Vertrag und Vertragsparteien, Vertragspreis, Rangfolge der Dokumente, Scope of Work und Abgrenzungen, Haftung, Gewährleistung, Ersatzteile, Reaktionszeiten, Vertragsstrafen, Ausführungsfristen und Termine sowie sonstige Vereinbarungen. Das Prüfergebnis ist in einem standardisierten Formular zu dokumentieren und über den Salesforce-Case mit allen erforderlichen Dokumenten zur weiteren Prüfung und Bestätigung zu übermitteln. Der Sales Manager verantwortet die ordnungsgemäße Durchführung des Prozesses (A); Legal wird über Salesforce informiert (I).

**EN:** The Seller reviews the contract documents — commercial and technical annexes — for valid points regarding risks and opportunities (R). For the commercial documents, the following in particular are reviewed: contract and contracting parties, contract price, order of precedence of documents, scope of work and delineations, liability, warranty, spare parts, response times, liquidated damages, completion deadlines and dates, and any other agreements. The review result is documented in a standardized form and submitted via the Salesforce case, together with all required documents, for further review and confirmation. The Sales Manager is accountable for the proper conduct of the process (A); Legal is informed via Salesforce (I).

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | — | Legal |

---

### 7. Labor Calculation

**Standardisiert / Standardized:** N (noch zu standardisieren / to be standardized) · **Durchlaufzeit (W):** 1 week · **Aufwand (C):** 16hrs · **Tools:** THF

**DE:** Die Arbeitskalkulation liegt in der Verantwortung des Estimators (R), mit möglicher Zuarbeit des Offshore Engineers (C). Der Project/Ops Manager wird konsultiert (C). Die finale Verantwortung trägt der Estimate Manager (A).

**EN:** The labor calculation is owned by the Estimator (R), with possible input from the Offshore Engineer (C). The Project/Ops Manager is consulted (C). Final accountability rests with the Estimate Manager (A).

| R | A | C | I |
|---|---|---|---|
| Estimator | Estimate Manager | Offshore Engineer, Project/Ops Manager | — |

---

### 8. Subcontracting

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** 10 days · **Aufwand (C):** 8hrs · **Tools:** Email, Phone, THF, Excel

**DE:** Der Seller benennt dem Estimator die anzufragenden Subcontractor zur Angebotseinholung (C). Der Estimator steuert und verantwortet die Subcontractor-Anfrage (A) und informiert zugleich das Procurement-Team. Die Anfrage und Angebotseinholung erfolgt operativ über Procurement sowie die Subcontractor selbst (R). Reminder, Klärungen und die Beziehungspflege laufen aktuell über den Seller. Der Service Seller ist einzubeziehen (C), da über ihn weitere, ihm bekannte und etablierte Subcontractor verfügbar sein können; die Abstimmung mit dem Service Seller ist eng zu führen.

**EN:** The Seller specifies to the Estimator which subcontractors are to be approached for a quote (C). The Estimator directs and is accountable for the subcontractor request (A) and concurrently informs the Procurement team. The request and quote sourcing are carried out operationally by Procurement and the subcontractors themselves (R). Reminders, clarifications and relationship management currently run through the Seller. The Service Seller is to be involved (C), as additional established subcontractors known to him may be available; close coordination with the Service Seller is to be ensured.

| R | A | C | I |
|---|---|---|---|
| Procurement, Subcontractors | Estimator | Seller, Service Seller | — |

---

### 9. JCI Equipment

**Standardisiert / Standardized:** N (noch zu standardisieren / to be standardized) · **Durchlaufzeit (W):** 3 days · **Aufwand (C):** 1,5hrs · **Tools:** Email, THF, Excel

**DE:** Die Anfrage des JCI-Equipments liegt in der Verantwortung des Estimators (R); ist die Aufgabe an das Offshore-Team ausgelagert, fragt dieses das JCI Global Products Team an (C) und erfragt den bestmöglichen Discount. Das JCI Global Products Team stellt das bestmögliche Discount-Angebot und das bestmögliche Material gemäß Vorgabe bereit (A). Das Material wird nach unserer Vorgabe angefragt. Der Seller ist über den Vorgang zu informieren (I).

**EN:** Sourcing the JCI equipment is owned by the Estimator (R); where the task is delegated to the Offshore team, the latter requests the JCI Global Products team (C) and obtains the best possible discount. The JCI Global Products team provides the best possible discount offer and material per specification (A). The material is requested according to our specification. The Seller is to be informed of the process (I).

| R | A | C | I |
|---|---|---|---|
| Estimator | JCI Global Products | Offshore Engineer | Seller |

---

### 10. 3rd Party Equipment

**Standardisiert / Standardized:** N (noch zu standardisieren / to be standardized) · **Durchlaufzeit (W):** 3 days · **Aufwand (C):** 20hrs · **Tools:** Email, THF, Excel

**DE:** Die Verantwortung liegt beim Offshore Engineer (R), der hierfür die Handover-Liste aus dem Schritt „Create" erhält. Fehlen Preise, kontaktiert der Offshore Engineer unverzüglich den Estimator, der die Preise anfragt (C). Informiert werden Seller, Procurement und Project/Ops Manager (I). Die finale Verantwortung trägt der Estimate Manager (A).

**EN:** Ownership rests with the Offshore Engineer (R), who receives the handover list from the "Create" step. If prices are missing, the Offshore Engineer immediately contacts the Estimator, who requests the prices (C). Seller, Procurement and Project/Ops Manager are informed (I). Final accountability rests with the Estimate Manager (A).

| R | A | C | I |
|---|---|---|---|
| Offshore Engineer | Estimate Manager | Estimator | Seller, Procurement, Project/Ops Manager |

---

### 11. Service

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** 3 days · **Aufwand (C):** 2hrs · **Tools:** Email, Phone, THF, Teams

**DE:** Der Estimator holt das bestmögliche Angebot beim Service Seller ein und informiert zugleich den Service Team Lead (Service TL) über die Anfrage (A). Der Service Seller verantwortet die Ermittlung des bestmöglichen Preises (R), damit die Installation-Systems-Abteilung wettbewerbsfähig am Markt anbieten kann. Der Service TL wird informiert (I).

**EN:** The Estimator obtains the best possible offer from the Service Seller and concurrently informs the Service Team Lead (Service TL) of the request (A). The Service Seller is responsible for determining the best possible price (R) so that the Installation Systems department can offer competitively in the market. The Service TL is informed (I).

| R | A | C | I |
|---|---|---|---|
| Service Seller | Estimator | — | Service TL |

---

### 12. OPS Review

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** 3 days · **Aufwand (C):** 4hrs · **Tools:** SFDC, Excel, Teams

**DE:** Der Seller lädt den Project/Ops Manager zum OPS Review ein und bereitet die Dokumente vor (R). Konsultierend hinzuzuziehen sind Post-Sales Engineer, Estimator und EHS (C). Verantwortlich für den Prozess ist der Sales Manager (A); der Estimate Manager wird informiert (I).

**EN:** The Seller invites the Project/Ops Manager to the OPS Review and prepares the documents (R). The Post-Sales Engineer, Estimator and EHS are consulted (C). The Sales Manager is accountable for the process (A); the Estimate Manager is informed (I).

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | Post-Sales Engineer, Estimator, EHS | Estimate Manager |

---

### 13. QARF

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** — · **Aufwand (C):** 3hrs · **Tools:** Excel

**DE:** Das QARF-Dokument ist vom Seller zu erstellen und dem Sales Manager vorzulegen (R), der die finale Verantwortung für das Dokument trägt (A). Gegenstand sind insbesondere die Margenerwartung und der Lebenszyklus für das Gesamtprojekt.

**EN:** The QARF document is prepared by the Seller and submitted to the Sales Manager (R), who holds final accountability for the document (A). It addresses in particular the margin expectation and the lifecycle of the overall project.

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | — | — |

---

### 14. DoA (Delegation of Authority)

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** 350k → 3 days / 1m → 10 days · **Aufwand (C):** 350k → 1hr / 1m → 40hrs · **Tools:** Email (ppt), SFDC, Excel, Teams

**DE:** Der DoA-Prozess wird durch den Seller eingesteuert (R); gesamtverantwortlich ist der Sales Manager (A).
- **Projekte < 100.000 USD:** Die Freigabeanforderung ist über Salesforce an Legal, Finance und den Sales Manager zu stellen.
- **Projekte > 1.000.000 USD:** Es ist ein Brief-it (DoA-Freigabe-Briefing) durchzuführen. Drei Tage vor dem Brief-it-Termin führt der Seller einen vorbereitenden Brief-it-Call durch. Einzuladen sind als Teilnehmende: Estimator, Sales Manager, Procurement, Projektmanager, Ops Manager, Legal und Finance; optional: Service Seller und PoC-Engineering. Zum eigentlichen Brief-it-Termin werden dieselben Parteien geladen.

Ziel ist die Freigabe der Projektdurchführung. Darzustellen sind die Projektdurchführung im Ausführungszeitraum, Chancen und Risiken, weitere Wachstumsszenarien sowie der finale Preis und die Marge; die rechtlichen Rahmenbedingungen bedürfen einer Freigabe gemäß DOA-Richtlinie. Es sind mindestens zwei Subunternehmen vorzulegen; wird diese Anforderung nicht erfüllt, ist ein MBE-Prozess erforderlich. Alle eingesetzten Subunternehmen folgen dem Prime-Revenue-Programm. Konsultiert werden Legal, Finance, Procurement, Project/Ops Manager, Service Seller und Post-Sales Engineer (C).

**EN:** The DoA process is initiated by the Seller (R); the Sales Manager is overall accountable (A).
- **Projects < USD 100,000:** The approval request is submitted via Salesforce to Legal, Finance and the Sales Manager.
- **Projects > USD 1,000,000:** A Brief-it (DoA approval briefing) is held. Three days before the Brief-it, the Seller runs a preparatory Brief-it call. The invitees are: Estimator, Sales Manager, Procurement, Project Manager, Ops Manager, Legal and Finance; optional: Service Seller and PoC engineering. The same parties are invited to the Brief-it itself.

The objective is approval to execute the project. The execution over the delivery period, opportunities and risks, further growth scenarios as well as the final price and margin are presented; the legal framework requires approval in accordance with the DOA policy. At least two subcontractors must be presented; if this requirement is not met, an MBE process is required. All engaged subcontractors follow the Prime Revenue program. Legal, Finance, Procurement, Project/Ops Manager, Service Seller and Post-Sales Engineer are consulted (C).

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | Legal, Finance, Procurement, Project/Ops Manager, Service Seller, Post-Sales Engineer | — |

---

### 15. Present Quote

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** — · **Aufwand (C):** 4hrs · **Tools:** Email, Phone, Site Visit

**DE:** Der Estimator stellt sämtliche Dokumente zusammen. Anschließend erfolgt eine Prüfung nach dem 4I-Prinzip (Vier-Augen-Prüfung) gemeinsam mit dem Seller. Der Seller zeichnet das Angebot und lädt es bei öffentlichen Ausschreibungen im Portal hoch (R). Bei privaten Vergaben wird das Angebot im besten Fall in einem persönlichen Gespräch überreicht, andernfalls per E-Mail. Verantwortlich für den Prozess ist der Sales Manager (A); der Project/Ops Manager wird konsultiert (C).

**EN:** The Estimator compiles all documents. The quote is then reviewed under the 4I principle (four-eyes review) together with the Seller. The Seller signs the quote and, for public tenders, uploads it to the portal (R). For private awards, the quote is preferably handed over in a personal meeting, otherwise by email. The Sales Manager is accountable for the process (A); the Project/Ops Manager is consulted (C).

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | Project/Ops Manager | — |

---

### 16. Negotiate

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** Service 3hrs / Install 2 weeks · **Aufwand (C):** Service 3hrs / Install 1 week · **Tools:** Email, Phone, Site Visit, Teams

**DE:** Die Verhandlung wird federführend vom Seller organisiert (R); gesamtverantwortlich bleibt der Sales Manager (A). Unterstützend wirken Estimator, Project/Ops Manager, Legal und Finance (C). Ziel ist das bestmögliche Verhandlungsergebnis. Die Verhandlung ist vorzugsweise in Präsenz zu führen und nicht über E-Mail, Telefon oder Teams; abweichend ist dies im Einzelfall zulässig, im Sinne des Kundenkontakts und der persönlichen Vorstellung ist die Verhandlung vor Ort vorzuziehen. Sämtliche freigabepflichtigen Änderungen sind anzumelden und durch die zuständigen Stellen und Vorgesetzten gemäß DOA freizugeben. Der Vertragsabschluss erfolgt stets im Vier-Augen-Prinzip und mit Unterschrift auf den entsprechenden Dokumenten, jedoch erst nach erfolgter DOA-Freigabe.

**EN:** The negotiation is organized and led by the Seller (R); the Sales Manager remains overall accountable (A). Estimator, Project/Ops Manager, Legal and Finance support (C). The objective is the best possible negotiation outcome. The negotiation is preferably conducted in person and not via email, phone or Teams; deviations are permissible in individual cases, but for customer contact and personal representation the on-site negotiation is preferred. All changes requiring approval must be raised and approved by the responsible functions and superiors in accordance with DOA. Contract conclusion always follows the four-eyes principle and is signed on the relevant documents, but only after DOA approval has been granted.

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | Estimator, Project/Ops Manager, Legal, Finance | — |

---

### 17. Close (Booking)

**Standardisiert / Standardized:** Y · **Durchlaufzeit (W):** Service 1 day / Install 2 days · **Aufwand (C):** Service 3hrs / Install 3hrs · **Tools:** Email, Phone, SFDC, Oracle Webtools

**DE:** Der Close erfolgt über den Seller, der das Projekt in Salesforce mit dem entsprechenden Status (gewonnen oder verloren) abmeldet (R) und Finance mit den entsprechenden Unterlagen versorgt, damit das Projekt im EIC-System gebucht werden kann (C: Finance). Zu beachten ist der Monatsabschluss: Projekte < 100.000 USD können nicht abgegrenzt werden; der Monatsabschluss findet in der Regel zwei Tage vor Monatsende statt. Zu informieren sind Estimator, Procurement (Einkauf) und Project/Ops Manager (I). Den Subcontractors ist mitzuteilen, ob das Projekt gewonnen oder verloren wurde (I). Ebenso zu informieren sind Legal, Estimate Manager, Service Seller und Service TL (I). Verantwortlich für den Prozess ist der Sales Manager (A).

**EN:** The Close is performed by the Seller, who closes the project in Salesforce with the corresponding status (won or lost) (R) and provides Finance with the relevant documents so that the project can be booked in the EIC system (consulted: Finance). The month-end close must be observed: projects < USD 100,000 cannot be accrued/deferred; the month-end close usually takes place two days before month-end. Estimator, Procurement and Project/Ops Manager are informed (I). The Subcontractors are informed whether the project was won or lost (I). Legal, Estimate Manager, Service Seller and Service TL are likewise informed (I). The Sales Manager is accountable for the process (A).

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | Finance | Estimator, Procurement, Project/Ops Manager, Subcontractors, Legal, Estimate Manager, Service Seller, Service TL |

---

### 18. Post Mortem

**Standardisiert / Standardized:** N (noch zu standardisieren / to be standardized) · **Durchlaufzeit (W):** — · **Aufwand (C):** — · **Tools:** —

**DE:** Die Auswertung ausgewählter Projekte erfolgt durch Sales Manager und Seller. Die Durchführung verantwortet der Seller (R); der Sales Manager ist verantwortlich (A) und bespricht mit dem Seller, welche Projekte ausgewertet werden. Ausgewertet werden verlorene und gewonnene Projekte; entsprechend wird ein Lessons-Learned erstellt. Unterstützend wirken Estimator, Offshore Engineer, Procurement, Project/Ops Manager, Legal, Finance und Service Seller mit (C). Informiert werden Post-Sales Engineer, Subcontractors, JCI Global Products, Estimate Manager, EHS und Service TL (I).
Ziel sind langfristige Verbesserungen: Erfolge werden festgehalten und auf andere Projekte übertragen; bei verlorenen Projekten werden die Ursachen herausgearbeitet — interne Kalkulation, zu hohe Subunternehmerkosten oder nicht erkannte Vertragsrisiken — und gemeinsam mit den Fachabteilungen sowie gegebenenfalls den Subunternehmen Verbesserungen erarbeitet.

**EN:** The review of selected projects is carried out by the Sales Manager and the Seller. The Seller is responsible for execution (R); the Sales Manager is accountable (A) and agrees with the Seller which projects are reviewed. Both lost and won projects are reviewed; a lessons-learned document is produced accordingly. Estimator, Offshore Engineer, Procurement, Project/Ops Manager, Legal, Finance and Service Seller contribute (C). Post-Sales Engineer, Subcontractors, JCI Global Products, Estimate Manager, EHS and Service TL are informed (I).
The objective is long-term improvement: successes are captured and transferred to other projects; for lost projects the causes are identified — internal calculation, excessive subcontractor costs or unrecognized contract risks — and improvements are developed together with the specialist functions and, where applicable, the subcontractors.

| R | A | C | I |
|---|---|---|---|
| Seller | Sales Manager | Estimator, Offshore Engineer, Procurement, Project/Ops Manager, Legal, Finance, Service Seller | Post-Sales Engineer, Subcontractors, JCI Global Products, Estimate Manager, EHS, Service TL |

---

## 5. Governance und Regeln / Governance and Rules

| Thema / Topic | Regel / Rule |
|---|---|
| **DoA-Schwellen / DoA thresholds** | < 100.000 USD: Freigabe über Salesforce an Legal, Finance, Sales Manager. > 1.000.000 USD: Brief-it inkl. vorbereitendem Call (T-3 Tage). / < USD 100k: approval via Salesforce to Legal, Finance, Sales Manager. > USD 1m: Brief-it incl. preparatory call (T-3 days). |
| **Subunternehmen / Subcontractors** | Mindestens zwei Subunternehmen vorzulegen; andernfalls MBE-Prozess. Alle Subunternehmen folgen dem Prime-Revenue-Programm. / At least two subcontractors required; otherwise MBE process. All subcontractors follow the Prime Revenue program. |
| **Angebotsprüfung / Quote review** | 4I-Prinzip (Vier-Augen-Prüfung) vor Abgabe. / 4I principle (four-eyes review) before submission. |
| **Vertragsabschluss / Contract conclusion** | Vier-Augen-Prinzip; Unterschrift erst nach DOA-Freigabe. / Four-eyes principle; signature only after DOA approval. |
| **Monatsabschluss / Month-end close** | In der Regel zwei Tage vor Monatsende; Projekte < 100.000 USD nicht abgrenzbar. / Usually two days before month-end; projects < USD 100k cannot be accrued/deferred. |

---

## 6. Mitgeltende Unterlagen / Referenced Documents

- `Sales_Process_RACI_Matrix_1.xlsx` — Blatt / sheet *RACI Matrix* (Quelle der RACI-Zuordnungen, Attribute und Werkzeuge / source of RACI assignments, attributes and tools).

---

*Ende des Dokuments / End of document — Status: TBD (to be discussed).*

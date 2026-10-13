"""Fictional bilingual tender questions and approved Contoso knowledge."""

VERSION = "2026.10.13-v1"


def bi(en, de):
    return {"en": en, "de": de}


TOPICS = {
    "QUAL": bi("Company and qualification", "Unternehmen und Eignung"),
    "ARCH": bi("Architecture and integration", "Architektur und Integration"),
    "NET": bi("Network engineering", "Netzwerktechnik"),
    "SEC": bi("Security and identity", "Sicherheit und Identität"),
    "DATA": bi("Data protection", "Datenschutz"),
    "EVID": bi("Evidence and assurance", "Nachweise und Prüfung"),
    "OPS": bi("Operations and service levels", "Betrieb und Service Levels"),
    "SUPPORT": bi("Support and continuity", "Support und Kontinuität"),
    "IMPL": bi("Implementation and acceptance", "Einführung und Abnahme"),
    "SUPPLY": bi("Supply and lifecycle", "Lieferung und Lebenszyklus"),
    "LEGAL": bi("Contract and exceptions", "Vertrag und Ausnahmen"),
    "COMM": bi("Pricing and five-year TCO", "Preise und Fünfjahres-TCO"),
}

# Each chapter has eleven independently answered requirements, not eleven copies
# of a parent prompt. The compact tender retains original IDs and dependencies.
QUESTION_TEXT = {
    "QUAL": """
Identify the bidding legal entity, delivery scope and countries covered; distinguish subcontractors from Contoso. || Benenne die anbietende Rechtseinheit, den Leistungsumfang und die abgedeckten Länder; trenne Unterauftragnehmer von Contoso.
Provide audited financial statements for the last three years and explain any missing year. || Lege geprüfte Abschlüsse der letzten drei Jahre vor und erkläre fehlende Jahre.
Name two reference projects with at least twelve banking sites; provide permission to contact each reference. || Benenne zwei Referenzprojekte mit mindestens zwölf Bankstandorten und die Erlaubnis zur Kontaktaufnahme.
Describe the engineering team and evidence of current network qualifications; do not list unverified individuals. || Beschreibe das Engineering-Team und belege aktuelle Netzwerkqualifikationen; führe keine ungeprüften Personen auf.
Disclose sanctions-screening and anti-bribery controls, owners and the frequency of review. || Beschreibe Sanktionsprüfung und Antikorruptionskontrollen mit Zuständigkeit und Prüfintervall.
State which subcontractors can access bank systems and how changes require bank approval. || Benenne Unterauftragnehmer mit Zugriff auf Banksysteme und den Freigabeprozess für Änderungen.
Provide professional indemnity and cyber-insurance limits, expiry dates and actual policy evidence. || Lege Deckung, Ablaufdaten und Policennachweise für Berufshaftpflicht und Cyberversicherung vor.
Explain how conflicts of interest are checked before assigning staff to this bank. || Erkläre die Prüfung von Interessenkonflikten vor dem Einsatz bei dieser Bank.
Provide an accessibility policy and evidence that operational documentation meets WCAG 2.2 AA. || Lege Barrierefreiheitsrichtlinie und Nachweise für Betriebsdokumentation nach WCAG 2.2 AA vor.
Describe environmental and disposal controls with evidence, without claiming an unheld certification. || Beschreibe Umwelt- und Entsorgungsmaßnahmen mit Belegen, ohne nicht vorhandene Zertifizierungen zu behaupten.
Identify the authority needed to sign a binding response and confirm whether this draft has it. || Benenne die nötige Zeichnungsberechtigung für eine bindende Antwort und ob dieser Entwurf sie besitzt.
""",
    "ARCH": """
Propose a segmented network for twelve sites with separate staff, guest, payment and management zones; show trust boundaries. || Entwirf für zwölf Standorte ein segmentiertes Netz mit Mitarbeiter-, Gast-, Zahlungs- und Managementzonen und Vertrauensgrenzen.
Explain how the design integrates with existing identity, DNS, monitoring and incident tools without assuming connectors. || Erkläre die Einbindung bestehender Identitäts-, DNS-, Monitoring- und Incident-Systeme ohne unbelegte Connectoren.
Provide a topology showing redundant uplinks, failure domains and the points that remain single points of failure. || Lege eine Topologie mit redundanten Uplinks, Ausfalldomänen und verbleibenden Single Points of Failure vor.
Describe the migration path from the bank's current VLANs and the information needed before final design. || Beschreibe den Weg von den bestehenden VLANs der Bank zum Zielbild und fehlende Planungsangaben.
Explain IP address management and overlapping ranges during coexistence; identify required bank approvals. || Erkläre IP-Adressverwaltung und überlappende Netze im Parallelbetrieb sowie nötige Bankfreigaben.
State the boundary between hardware supply, managed operations and the bank's own responsibilities. || Trenne Hardwarelieferung, Managed Operations und Eigenverantwortung der Bank.
Describe integration authentication, secret storage and rotation; provide evidence for any claimed automation. || Beschreibe Integrationsauthentifizierung, Geheimnisspeicherung und Rotation; belege behauptete Automatisierung.
Explain how configuration export avoids vendor lock-in and what cannot be exported. || Erkläre, wie Konfigurationsexport Herstellerbindung verringert und was nicht exportierbar ist.
Provide capacity assumptions for 400 endpoints per site and identify the measurements required to validate them. || Lege Kapazitätsannahmen für 400 Endgeräte je Standort und die nötigen Validierungsmessungen dar.
Describe a lab validation plan including identity outage and upstream network failure. || Beschreibe eine Laborprüfung mit Identitätsausfall und Ausfall des vorgelagerten Netzes.
Identify architecture decisions still requiring the bank's security and network teams. || Benenne Architekturentscheidungen, die noch die Sicherheits- und Netzwerkteams der Bank benötigen.
""",
    "NET": """
Confirm 10G uplinks on every offered access switch; reconcile the requested CMS-48P-POE with documented uplink capability. || Bestätige 10G-Uplinks an jedem angebotenen Access-Switch; gleiche CMS-48P-POE mit der belegten Uplink-Fähigkeit ab.
Demonstrate the PoE power budget for 32 access points per site at 30 W each and explain any required extra switches. || Belege das PoE-Budget für 32 Access Points je Standort mit je 30 W und erkläre nötige zusätzliche Switches.
Specify optical modules, cable types, distances and compatibility evidence for each uplink. || Benenne Optikmodule, Kabeltypen, Entfernungen und Kompatibilitätsnachweise je Uplink.
Explain 802.1X and fallback behavior for devices without supplicants, including payment-terminal risk. || Erkläre 802.1X und das Verhalten bei Geräten ohne Supplicant samt Risiko für Zahlungsterminals.
Confirm whether Wi-Fi 7 and legacy clients can coexist; identify survey and regulatory prerequisites. || Bestätige das Zusammenspiel von Wi-Fi 7 und Altgeräten sowie Funkmessungs- und regulatorische Voraussetzungen.
Provide roaming latency evidence for voice clients, not just a marketing capability statement. || Lege Roaming-Latenznachweise für Sprachclients statt einer Marketingaussage vor.
Explain guest isolation and how guest traffic is prevented from reaching payment systems. || Erkläre Gastisolation und die Sperre von Gastverkehr zu Zahlungssystemen.
State throughput with security inspection enabled and supply independently measured results. || Benenne den Durchsatz mit Sicherheitsprüfung und liefere unabhängig gemessene Ergebnisse.
Describe multicast and QoS configuration for voice/video, including validation under congestion. || Beschreibe Multicast und QoS für Sprache/Video samt Prüfung bei Überlastung.
Confirm rack, power and cooling requirements and list site-survey gaps. || Bestätige Rack-, Strom- und Kühlungsanforderungen und offene Standortmessungen.
Describe tested configuration backup and restore including firmware-version compatibility. || Beschreibe getestete Konfigurationssicherung und Wiederherstellung samt Firmware-Kompatibilität.
""",
    "SEC": """
Does the management console support SAML 2.0 SSO and bank-controlled MFA? State the supported boundary. || Unterstützt die Managementkonsole SAML-2.0-SSO und bankgesteuerte MFA? Benenne die unterstützte Grenze.
If SSO is supported, prove SCIM deprovisioning within five minutes, including a measured test. || Falls SSO unterstützt wird, belege SCIM-Deprovisionierung binnen fünf Minuten einschließlich Messung.
Describe least-privilege roles and segregation between bank operators and supplier administrators. || Beschreibe Minimalrechte und die Trennung von Bankbetrieb und Lieferantenadministration.
State encryption protocols and key ownership for management traffic and stored configuration. || Benenne Verschlüsselungsprotokolle und Schlüsselverantwortung für Managementverkehr und gespeicherte Konfiguration.
Provide independent penetration-test evidence no older than twelve months and remediation status. || Lege einen unabhängigen Penetrationstest unter zwölf Monaten und den Behebungsstand vor.
Explain critical vulnerability handling with a 48-hour mitigation requirement and escalation. || Erkläre kritische Schwachstellen mit einer Minderungsanforderung binnen 48 Stunden und Eskalation.
Show immutable audit retention for 365 days, export options and evidence of tamper resistance. || Belege unveränderliche Auditaufbewahrung über 365 Tage, Export und Manipulationsschutz.
Describe emergency administrator access, dual control and the audit trail. || Beschreibe administrativen Notzugang, Vieraugenkontrolle und Prüfspur.
Explain secure firmware provenance, update verification and rollback constraints. || Erkläre sichere Firmwareherkunft, Updateprüfung und Rollback-Grenzen.
State whether a software bill of materials is supplied for each firmware release and provide one. || Benenne, ob je Firmwareversion eine Software Bill of Materials vorliegt, und lege eine vor.
Identify residual security risks that require bank acceptance before production. || Benenne verbleibende Sicherheitsrisiken, die vor Produktion Bankakzeptanz benötigen.
""",
    "DATA": """
Can all management telemetry and backups remain exclusively in the EU with no non-EU administrator access? || Können Managementtelemetrie und Sicherungen ausschließlich in der EU bleiben, ohne administrativen Zugriff außerhalb der EU?
If any non-EU processing is required, identify countries, subprocessors, transfer safeguards and missing approvals. || Falls Verarbeitung außerhalb der EU nötig ist, benenne Länder, Unterauftragnehmer, Transfermaßnahmen und fehlende Freigaben.
List personal data categories, purposes and legal roles separately for telemetry and support. || Trenne personenbezogene Datenkategorien, Zwecke und rechtliche Rollen für Telemetrie und Support.
Provide the signed data processing agreement for this bank, not an unsigned template. || Lege die unterzeichnete Auftragsverarbeitungsvereinbarung dieser Bank statt einer unsignierten Vorlage vor.
State retention periods by data type and explain how deletion is evidenced. || Benenne Aufbewahrungsfristen je Datentyp und den Nachweis der Löschung.
Describe subject-access and deletion support, response times and bank responsibilities. || Beschreibe Auskunfts- und Löschunterstützung, Fristen und Bankverantwortung.
Confirm whether production data is used for analytics or AI training and cite an approved policy. || Bestätige die Nutzung produktiver Daten für Analytik oder KI-Training mit freigegebener Richtlinie.
Explain minimization and masking in support bundles before external access. || Erkläre Datenminimierung und Maskierung in Supportpaketen vor externem Zugriff.
Describe how privacy incidents are reported within the bank's required twelve-hour window. || Beschreibe Meldungen von Datenschutzvorfällen innerhalb des verlangten Zwölfstundenfensters.
Provide a data flow diagram with processing locations and a current subprocessor register. || Lege ein Datenflussdiagramm mit Verarbeitungsorten und ein aktuelles Unterauftragnehmerregister vor.
Describe exit deletion including replicas, backups and the certification of completion. || Beschreibe Löschung beim Exit einschließlich Replikaten, Backups und Abschlussbescheinigung.
""",
    "EVID": """
Is the complete proposed service currently ISO 27001 certified? Provide evidence, not self-assessment. || Ist der gesamte vorgeschlagene Service aktuell ISO-27001-zertifiziert? Liefere Nachweise statt Selbsteinschätzung.
If certified, supply certificate issuer, scope, number and expiry and map all exclusions to this service. || Falls zertifiziert, benenne Aussteller, Geltungsbereich, Nummer und Ablauf und ordne alle Ausnahmen diesem Service zu.
Provide a current SOC 2 Type II report and identify its observation period and service coverage. || Lege einen aktuellen SOC-2-Type-II-Bericht samt Prüfzeitraum und Leistungsabdeckung vor.
Map proposed payment-network controls to the bank's PCI requirements without claiming certification. || Ordne Zahlungsnetzmaßnahmen den PCI-Anforderungen der Bank zu, ohne Zertifizierung zu behaupten.
Provide independent evidence for claimed redundancy and failover times. || Lege unabhängige Belege für behauptete Redundanz und Failover-Zeiten vor.
Supply the most recent disaster-recovery test record and list unresolved findings. || Liefere das jüngste Wiederanlauf-Testprotokoll und offene Feststellungen.
Show document ownership, approval date and version for every response-library source. || Benenne Zuständigkeit, Freigabedatum und Version jeder Antwortbibliotheksquelle.
Explain how superseded documents are detected and excluded from current commitments. || Erkläre das Erkennen und Ausschließen überholter Dokumente aus aktuellen Zusagen.
List missing mandatory evidence and the accountable review function, without fabricated attachments. || Liste fehlende Pflichtnachweise und zuständige Prüffunktionen ohne erfundene Anlagen.
Confirm the bank's right to inspect evidence under NDA and identify redacted items. || Bestätige Einsicht unter NDA und benenne geschwärzte Unterlagen.
Provide a traceability register linking each requirement to an actual section and cloud file. || Liefere ein Nachweisregister von jeder Anforderung zu einem tatsächlichen Abschnitt und einer Cloud-Datei.
""",
    "OPS": """
Commit to 99.95% 24/7 availability; reconcile any different SLA versions in the knowledge base. || Sage 99,95 % Verfügbarkeit rund um die Uhr zu; gleiche abweichende SLA-Versionen der Wissensbasis ab.
Explain the availability denominator, maintenance exclusions and maximum service credits under the current SLA. || Erkläre Bezugszeitraum, Wartungsausnahmen und maximale Servicegutschriften des aktuellen SLA.
Describe monitoring coverage, alert ownership and the limits of the proposed managed service. || Beschreibe Monitoringabdeckung, Alarmzuständigkeit und Grenzen des Managed Service.
Define severity levels and response versus resolution targets; do not conflate the two. || Definiere Schweregrade und Reaktions- gegenüber Lösungszielen; setze sie nicht gleich.
Provide an escalation matrix for out-of-hours outages and state uncovered periods. || Lege Eskalation für Störungen außerhalb der Betriebszeit und ungedeckte Zeiträume dar.
Describe change windows, bank sign-off and rollback ownership. || Beschreibe Änderungsfenster, Bankfreigabe und Rollback-Verantwortung.
State patch cadence and emergency-change handling with evidence. || Benenne Patchintervalle und belegten Umgang mit Notfalländerungen.
Specify operational dashboards, report frequency and which inputs the bank must provide. || Benenne Betriebsübersichten, Berichtstakt und nötige Bankeingaben.
Describe configuration drift detection and the approval needed to remediate it. || Beschreibe Konfigurationsabweichungen und die nötige Freigabe zur Korrektur.
Explain how major incidents produce root-cause reports and corrective actions. || Erkläre Ursachenberichte und Korrekturmaßnahmen nach Großstörungen.
State service transition criteria before the supplier takes operational responsibility. || Benenne Übergabekriterien vor Übernahme der Betriebsverantwortung.
""",
    "SUPPORT": """
Can three-year next-business-day hardware care support the entire five-year term? Identify the renewal evidence gap. || Deckt dreijährige Next-Business-Day-Hardwarebetreuung die gesamten fünf Jahre? Benenne die Nachweislücke für Verlängerungen.
Define business-day cutoffs, replacement logistics and geographic exclusions without promising four-hour replacement. || Definiere Geschäftstagsfristen, Ersatzlogistik und Gebietsausnahmen ohne Vierstundenersatz zu versprechen.
State support channels, languages and covered hours. || Benenne Supportkanäle, Sprachen und abgedeckte Zeiten.
Describe responsibility for diagnosis versus hardware replacement. || Trenne Verantwortung für Diagnose und Hardwareersatz.
Provide recovery-time and recovery-point objectives with supporting test evidence. || Benenne Wiederanlauf- und Datenverlustziele mit Testbelegen.
Explain resilience when the supplier's own service desk is unavailable. || Erkläre Resilienz beim Ausfall des eigenen Service Desks.
Specify on-site spare holdings and who pays for unused or obsolete stock. || Benenne Vor-Ort-Ersatzbestände und die Kosten ungenutzter oder veralteter Ware.
Describe end-user communications during an incident without assuming external-send permission. || Beschreibe Nutzerkommunikation im Störfall ohne Versandfreigabe vorauszusetzen.
Explain defect triage across hardware, firmware, identity and upstream networks. || Erkläre Fehlerabgrenzung zwischen Hardware, Firmware, Identität und vorgelagertem Netz.
Provide a verified continuity-test schedule and identify missing records. || Lege belegte Kontinuitätsprüftermine vor und benenne fehlende Protokolle.
Describe handover to a replacement supplier at contract end. || Beschreibe Übergabe an einen Nachfolgelieferanten am Vertragsende.
""",
    "IMPL": """
Deliver twelve sites within six weeks; provide a gated plan and conditions rather than an unsupported date promise. || Liefere zwölf Standorte in sechs Wochen; gib einen gestuften Plan und Bedingungen statt einer unbelegten Terminzusage.
Define objective acceptance tests for segmentation, failover and throughput and who signs the results. || Definiere objektive Abnahmetests für Segmentierung, Failover und Durchsatz samt Unterschrift.
Specify the pilot site, go/no-go criteria and what happens after pilot failure. || Benenne Pilotstandort, Go/No-Go-Kriterien und das Vorgehen bei Pilotfehlern.
List bank dependencies including access, cabling, power and identity configuration. || Liste Bankabhängigkeiten wie Zugang, Verkabelung, Strom und Identitätskonfiguration.
Provide training scope and evidence of operator readiness. || Benenne Schulungsumfang und Nachweise der Betriebsbereitschaft.
Describe the cutover window, outage notice and tested rollback plan. || Beschreibe Umschaltfenster, Störungsankündigung und getesteten Rollback.
Show project governance, decision rights and change control. || Benenne Projektsteuerung, Entscheidungsrechte und Änderungskontrolle.
Define configuration documentation and asset-register handover deliverables. || Definiere Konfigurationsdokumentation und Übergabe des Anlagenregisters.
Describe the method for validating RF coverage before and after installation. || Beschreibe die Prüfung der Funkabdeckung vor und nach Installation.
Estimate on-site effort and separate included work from additional day rates. || Schätze Vor-Ort-Aufwand und trenne enthaltene Arbeit von zusätzlichen Tagessätzen.
State commercial consequences of bank-caused delays without automatically accepting penalties. || Benenne Folgen bankverursachter Verzögerungen ohne automatische Vertragsstrafenakzeptanz.
""",
    "SUPPLY": """
The bill of materials requests 48 CAP-6E devices; verify lifecycle and propose CAP-7 only as a disclosed alternative. || Die Stückliste fordert 48 CAP-6E; prüfe den Lebenszyklus und schlage CAP-7 nur als offengelegte Alternative vor.
If available stock is insufficient, show dated allocation and delivery options without reserving or guaranteeing them. || Bei unzureichendem Bestand zeige datierte Zuteilungs- und Lieferoptionen ohne Reservierung oder Garantie.
For each line show on-hand, reserved and available quantities and distinguish a plan from a commitment. || Zeige je Position vorhandene, reservierte und verfügbare Mengen; trenne Plan und Zusage.
Identify which alternatives need technical or bank approval and quantify price differences from live previews. || Benenne technisch oder von der Bank freizugebende Alternativen und Preisdifferenzen aus Live-Vorschauen.
Explain split shipments and their impact on site sequencing and costs. || Erkläre Teillieferungen und Auswirkungen auf Standortreihenfolge und Kosten.
Provide lifecycle dates only where currently evidenced; do not turn an unknown date into a five-year warranty. || Nenne nur aktuell belegte Lebenszyklusdaten; mache aus einem unbekannten Datum keine Fünfjahresgarantie.
State incoterms, delivery locations and who bears transit risk. || Benenne Incoterms, Lieferorte und Trägerschaft des Transportrisikos.
Describe counterfeit prevention and serial-number traceability with actual evidence. || Beschreibe Fälschungsprävention und Seriennummern-Nachverfolgung mit Belegen.
Show dependencies between optics, switches, gateways and care subscriptions before promising completeness. || Zeige Abhängigkeiten zwischen Optiken, Switches, Gateways und Betreuung vor einer Vollständigkeitszusage.
Describe cancellation and substitution rules for delayed items. || Beschreibe Stornierungs- und Austauschregeln für verspätete Artikel.
Explain replenishment and replacement sourcing after product end of life. || Erkläre Nachbeschaffung und Ersatzbeschaffung nach Produktabkündigung.
""",
    "LEGAL": """
Accept unlimited liability, including indirect loss, or explicitly record a contractual exception. || Akzeptiere unbegrenzte Haftung einschließlich Folgeschäden oder erfasse ausdrücklich eine Vertragsausnahme.
If unlimited liability is not accepted, propose the approved counterposition and identify legal review requirements. || Falls unbegrenzte Haftung nicht akzeptiert wird, nenne die freigegebene Gegenposition und nötige Rechtsprüfung.
State the governing law, jurisdiction and which differences from bank terms remain unresolved. || Benenne Recht, Gerichtsstand und offene Abweichungen von Bankbedingungen.
Explain ownership of configurations, documentation and pre-existing intellectual property. || Erkläre Eigentum an Konfiguration, Dokumentation und bestehendem geistigem Eigentum.
Describe termination assistance, duration and charges without assuming free unlimited transition. || Beschreibe Exit-Unterstützung, Dauer und Kosten ohne unbegrenzten kostenfreien Übergang anzunehmen.
Specify price-adjustment rights and whether any annual indexation has been agreed. || Benenne Preisanpassungsrechte und ob jährliche Indexierung vereinbart wurde.
Provide audit rights and restrictions for third-party supplier facilities. || Benenne Prüfrechte und Grenzen für Einrichtungen Dritter.
State warranty scope, remedies and exclusions using approved evidence rather than hidden API fields. || Benenne Garantieumfang, Abhilfe und Ausnahmen aus freigegebenen Belegen statt verborgener API-Felder.
Describe delay penalties and limits that require commercial review. || Beschreibe Verzögerungsstrafen und kaufmännisch zu prüfende Grenzen.
Explain confidentiality and permitted use of bank information during bidding. || Erkläre Vertraulichkeit und zulässige Bankdatennutzung während des Verfahrens.
State explicitly that this draft does not constitute signature, contract acceptance or order creation. || Stelle klar, dass dieser Entwurf keine Unterschrift, Vertragsannahme oder Auftragserstellung ist.
""",
    "COMM": """
Provide a five-year TCO in the customer's verified currency, separating hardware, implementation, operations and care renewals. || Lege Fünfjahres-TCO in geprüfter Kundenwährung vor; trenne Hardware, Einführung, Betrieb und Betreuungsverlängerung.
State years four and five renewal assumptions and mark TCO incomplete where no approved renewal price exists. || Benenne Verlängerungsannahmen für Jahre vier und fünf und markiere TCO ohne freigegebenen Preis als unvollständig.
Show customer-specific unit net prices, quantities, discounts and source timestamps from the unsaved Atlas preview. || Zeige kundenspezifische Nettopreise, Mengen, Rabatte und Quellenzeitpunkte aus der ungespeicherten Atlas-Vorschau.
Reconcile line rounding, subtotal, tax and total against the preview, not an EUR list-price conversion. || Gleiche Positionsrundung, Zwischensumme, Steuer und Gesamtbetrag mit der Vorschau statt EUR-Listenpreisumrechnung ab.
Separate one-time, monthly and annual costs and state the unit and period of each rate. || Trenne einmalige, monatliche und jährliche Kosten mit Einheit und Zeitraum jeder Rate.
Provide optional and mandatory costs separately without silently removing mandatory items from the total. || Trenne optionale und verpflichtende Kosten ohne Pflichtpositionen still aus der Summe zu entfernen.
Explain payment terms, credit warnings and financing dependencies observed on the actual account. || Erkläre beobachtete Zahlungskonditionen, Kreditwarnungen und Finanzierungsabhängigkeiten des tatsächlichen Kontos.
State any price validity period as a draft assumption unless explicitly supported by current evidence. || Benenne Preisgültigkeit nur als Entwurfsannahme, sofern kein aktueller Beleg vorliegt.
Show how a switch or access-point substitution changes TCO without inventing a price. || Zeige TCO-Auswirkungen eines Switch- oder Access-Point-Austauschs ohne erfundene Preise.
Identify internal approval needs separately; do not expose supplier costs, margin or approval mechanics to the customer. || Erfasse interne Freigaben getrennt; zeige dem Kunden keine Lieferantenkosten, Marge oder Freigabelogik.
Provide an open-cost register so missing shipping, tax or renewal values cannot appear as confirmed zero. || Erstelle ein Register offener Kosten, damit fehlende Versand-, Steuer- oder Verlängerungswerte nicht als bestätigte Null erscheinen.
""",
}

KB = {
    "QUAL": [
        bi("Scope: Contoso Network Services is the fictional bidder for hardware supply, design and controlled deployment at twelve Woodgrove sites in the US. Bank systems and payment processing remain the bank's responsibility.",
           "Umfang: Contoso Network Services ist der fiktive Anbieter für Hardware, Planung und geregelte Einführung an zwölf Woodgrove-Standorten in den USA. Banksysteme und Zahlungsabwicklung bleiben Bankverantwortung."),
        bi("Governance: engineering, service operations, security, privacy, legal and commercial review are separate functions. Only an authorized signatory may bind the company. This package has no signature or approval.",
           "Steuerung: Engineering, Betrieb, Sicherheit, Datenschutz, Recht und kaufmännische Prüfung sind getrennte Funktionen. Nur Zeichnungsberechtigte binden das Unternehmen. Dieses Paket besitzt keine Unterschrift oder Freigabe."),
        bi("Evidence limits: no audited accounts, permitted bank references, named staff credentials, insurance policies, sanctions audit, accessibility audit or environmental certificate are included. Request evidence; do not invent names, limits or certifications.",
           "Nachweisgrenzen: Keine geprüften Abschlüsse, freigegebenen Bankreferenzen, Personalqualifikationen, Versicherungspolicen, Sanktions-, Barrierefreiheits- oder Umweltaudits liegen bei. Nachweise anfordern; keine Namen, Deckungen oder Zertifikate erfinden."),
    ],
    "ARCH": [
        bi("Design basis: separate staff, guest, payment and management zones; deny inter-zone traffic unless explicitly approved. Redundant uplinks require independently powered paths. Configuration and identity changes require bank sign-off.",
           "Planung: Mitarbeiter-, Gast-, Zahlungs- und Managementzonen trennen; zonenübergreifenden Verkehr ohne Freigabe sperren. Redundante Uplinks benötigen unabhängig versorgte Pfade. Konfigurations- und Identitätsänderungen benötigen Bankfreigabe."),
        bi("Integration: the bank supplies address ranges, VLANs, DNS, identity metadata and monitoring endpoints. No preinstalled connector, topology drawing, lab result or 400-endpoint capacity benchmark is approved in this pack.",
           "Integration: Die Bank liefert Adressräume, VLANs, DNS, Identitätsmetadaten und Monitoringziele. Kein vorinstallierter Connector, Topologieplan, Laborergebnis oder Benchmark für 400 Endgeräte ist in diesem Paket freigegeben."),
        bi("Responsibilities: supply includes hardware and design; managed operations is a separately priced service. Bank owns applications, payment controls, cabling and site access. Configuration exports are required at handover; vendor-specific settings require review.",
           "Verantwortung: Lieferung umfasst Hardware und Planung; Managed Operations wird getrennt bepreist. Anwendungen, Zahlungskontrollen, Kabel und Zugang liegen bei der Bank. Konfigurationsexport gehört zur Übergabe; herstellerspezifische Einstellungen benötigen Prüfung."),
    ],
    "NET": [
        bi("Fictional engineering specification: CMS-48P-POE has 48 PoE ports, 1G uplinks and a 740 W budget. CMS-48X-POE has 48 PoE ports, 10G uplinks and a 960 W budget. A 32 x 30 W load consumes the full 960 W before headroom; final design needs a measured diversity/headroom decision.",
           "Fiktive Engineering-Spezifikation: CMS-48P-POE besitzt 48 PoE-Ports, 1G-Uplinks und 740 W Budget. CMS-48X-POE besitzt 48 PoE-Ports, 10G-Uplinks und 960 W Budget. 32 x 30 W verbrauchen 960 W ohne Reserve; die endgültige Planung benötigt eine geprüfte Gleichzeitigkeits-/Reserveentscheidung."),
        bi("CAP-7 is the proposed Wi-Fi 7 alternative to CAP-6E. Compatibility, radio regulations, coverage and client roaming must be validated in a survey and pilot. CSFP-10G needs fiber-distance and switch compatibility confirmation. These are demo KB specifications, not visible Atlas measurements.",
           "CAP-7 ist die vorgeschlagene Wi-Fi-7-Alternative zu CAP-6E. Kompatibilität, Funkregulierung, Abdeckung und Roaming benötigen Messung und Pilot. CSFP-10G benötigt Bestätigung von Faserdistanz und Switch-Kompatibilität. Dies sind Demo-KB-Spezifikationen, keine sichtbaren Atlas-Messungen."),
        bi("Gaps: no independent throughput, roaming, PoE load, optics interoperability, rack/cooling or restore test is supplied. 802.1X, guest ACLs and QoS are design requirements, not proven as deployed.",
           "Lücken: Keine unabhängigen Durchsatz-, Roaming-, PoE-, Optik-, Rack/Kühlungs- oder Restoretests liegen bei. 802.1X, Gast-ACLs und QoS sind Planungsanforderungen, keine nachgewiesene Installation."),
    ],
    "SEC": [
        bi("Approved demo capability: management-console SAML 2.0 SSO with MFA enforced by the bank's identity provider. No SCIM provisioning or deprovisioning is supported. Use operator/admin role separation and logged break-glass access with bank approval.",
           "Freigegebene Demo-Fähigkeit: SAML-2.0-SSO der Managementkonsole mit MFA des Bank-Identitätsdiensts. SCIM-Provisionierung und -Deprovisionierung werden nicht unterstützt. Betrieb/Administration trennen und bankfreigegebenen Notzugang protokollieren."),
        bi("Baseline: TLS 1.2 or later for management traffic and encrypted configuration storage under supplier-managed keys. Customer-managed keys and immutable 365-day retention are not approved capabilities; audit export retention is 90 days.",
           "Basis: TLS 1.2 oder neuer für Managementverkehr und verschlüsselte Konfigurationsspeicherung mit Lieferantenschlüsseln. Kundeneigene Schlüssel und unveränderliche 365-Tage-Aufbewahrung sind nicht freigegeben; Auditexport wird 90 Tage aufbewahrt."),
        bi("Evidence gaps: independent penetration test, release-level SBOM, signed firmware-provenance record and 48-hour vulnerability remediation commitment are absent. Do not turn a design intention into evidence.",
           "Nachweislücken: Unabhängiger Penetrationstest, versionsbezogene SBOM, signierter Firmwareherkunftsnachweis und 48-Stunden-Behebungszusage fehlen. Planung ist kein Nachweis."),
    ],
    "DATA": [
        bi("Current processing: management telemetry in US East; backup in US West; supplier support administrators in the US and UK. EU-only processing and EU-only administrator access are not currently supported. Bank approval and legal transfer assessment are required.",
           "Aktuelle Verarbeitung: Managementtelemetrie in US East, Sicherung in US West, Supportadministration in USA und UK. Ausschließliche EU-Verarbeitung und EU-Administration werden derzeit nicht unterstützt. Bankfreigabe und rechtliche Transferprüfung sind nötig."),
        bi("Minimization: telemetry contains device identifiers, operator account identifiers and event timestamps; no payment payload is intended. Telemetry retention is 90 days; backups expire after 35 days. Mask identifiers before support-bundle sharing. Production bank data must not train AI.",
           "Minimierung: Telemetrie enthält Gerätekennungen, Betriebskonten und Ereigniszeiten; Zahlungsinhalte sind nicht vorgesehen. Telemetrie: 90 Tage; Sicherungen: 35 Tage. Kennungen vor Supportweitergabe maskieren. Produktive Bankdaten dürfen KI nicht trainieren."),
        bi("Only an unsigned DPA template exists. No signed Woodgrove DPA, transfer assessment, current subprocessor register, deletion certificate or twelve-hour breach commitment is in the evidence pack. Legal/privacy review is mandatory.",
           "Nur eine unsignierte AVV-Vorlage existiert. Unterzeichnete Woodgrove-AVV, Transferprüfung, aktuelles Unterauftragnehmerregister, Löschbescheinigung und Zwölfstunden-Meldezusage fehlen. Rechts-/Datenschutzprüfung ist Pflicht."),
    ],
    "EVID": [
        bi("Evidence inventory: ISO 27001 certification status for the complete service is unverified. No certificate number, issuer, scope, expiry or SOC 2 Type II report is supplied. Absence of evidence does not prove certification and does not prove non-certification.",
           "Nachweisbestand: ISO-27001-Zertifizierung des gesamten Service ist ungeprüft. Zertifikatsnummer, Aussteller, Umfang, Ablauf und SOC-2-Type-II-Bericht fehlen. Fehlende Nachweise beweisen weder Zertifizierung noch Nichtzertifizierung."),
        bi("Approved sources: use the current Source_Index and document section numbers. Self-assessments and design policies do not replace independently issued reports. Evidence may be reviewed under NDA only when a real approved document exists.",
           "Freigegebene Quellen: Aktuellen Source_Index und Abschnittsnummern nutzen. Selbsteinschätzungen und Planungsrichtlinien ersetzen keine unabhängigen Berichte. NDA-Einsicht erst anbieten, wenn echte freigegebene Unterlagen existieren."),
        bi("Missing test reports include failover, disaster recovery and payment-control audit. Keep the certification follow-up unresolved while its parent is unverified; never mark it not applicable merely because the certificate is missing.",
           "Failover-, Wiederanlauf- und Zahlungsprüfung fehlen. Zertifizierungsfolgefragen bleiben bei ungeprüfter Hauptfrage offen; fehlendes Zertifikat macht sie nicht automatisch irrelevant."),
    ],
    "OPS": [
        bi("CURRENT SLA v2, effective 2026-10-01: 99.5% managed-service availability within Monday-Friday 08:00-18:00 site-local time excluding holidays. It does not promise 99.95% 24/7. SLA v1 is withdrawn and cannot be used for commitments.",
           "AKTUELLES SLA v2 ab 01.10.2026: 99,5 % Managed-Service-Verfügbarkeit Montag-Freitag 08:00-18:00 Ortszeit ohne Feiertage. Keine Zusage von 99,95 % rund um die Uhr. SLA v1 ist zurückgezogen und darf keine Zusage begründen."),
        bi("Measurement: covered minutes minus approved maintenance; bank power/ISP outages and bank-caused changes are excluded. Credits capped at 10% of that month's managed-service fee. P1 acknowledgement target is one covered hour; there is no guaranteed resolution time or overnight response.",
           "Messung: Abgedeckte Minuten abzüglich freigegebener Wartung; Bankstrom/ISP-Ausfälle und bankverursachte Änderungen ausgeschlossen. Gutschriften maximal 10 % der monatlichen Servicegebühr. P1-Reaktionsziel: eine abgedeckte Stunde; keine garantierte Lösung oder Nachtreaktion."),
        bi("Monthly service report; standard patch window agreed with bank; emergency changes require approval and rollback record. Transition requires signed pilot acceptance and monitoring contacts. No historical incident, drift or patch audit is included.",
           "Monatlicher Servicebericht; Patchfenster mit Bank abstimmen; Notfalländerungen benötigen Freigabe und Rollback-Nachweis. Übergabe verlangt unterschriebene Pilotabnahme und Monitoringkontakte. Keine historischen Incident-, Drift- oder Patchaudits liegen bei."),
    ],
    "SUPPORT": [
        bi("CARE-NET-3Y-NBD covers three years only. Next business day is a service planning target after diagnosis, subject to cutoff, location and availability; not four-hour replacement. No approved year-four/year-five renewal price or future stock guarantee exists.",
           "CARE-NET-3Y-NBD deckt nur drei Jahre. Next Business Day ist ein Planungsziel nach Diagnose, abhängig von Frist, Ort und Bestand; kein Vierstundenersatz. Freigegebene Preise für Jahre vier/fünf und künftige Bestandsgarantien fehlen."),
        bi("Support: English service desk in the current SLA window; bank identifies a contact and distinguishes endpoint, identity, ISP and supplier hardware incidents. Local spares are optional and separately scoped.",
           "Support: Englischer Service Desk im aktuellen SLA-Fenster; Bank benennt Kontakt und grenzt Endgerät, Identität, ISP und Lieferantenhardware ab. Lokale Ersatzteile sind optional und getrennt zu planen."),
        bi("No measured RTO/RPO, continuity test or service-desk failover record exists in this pack. Exit includes configuration exports and asset list; transition effort needs an agreed scope and rate.",
           "Gemessene RTO/RPO, Kontinuitätstest und Service-Desk-Failover-Nachweis fehlen. Exit umfasst Konfigurationsexport und Anlagenliste; Übergangsaufwand benötigt vereinbarten Umfang und Satz."),
    ],
    "IMPL": [
        bi("Gated rollout: survey, approved design, single-site pilot, bank acceptance, then remaining sites in waves. Six weeks is a buyer target, conditional on stock, site readiness and bank sign-off, not an unconditional commitment.",
           "Gestufter Rollout: Messung, freigegebene Planung, Einstandortpilot, Bankabnahme, anschließend Standortwellen. Sechs Wochen sind ein Käuferziel, abhängig von Bestand, Bereitschaft und Bankfreigabe, keine unbedingte Zusage."),
        bi("Acceptance: test denied guest-to-payment traffic, uplink loss, agreed throughput and restore. Record configuration, timestamp, observed result and bank signature. Failed pilot prevents the next wave until review.",
           "Abnahme: Gast-Zahlungs-Sperre, Uplinkverlust, vereinbarten Durchsatz und Restore prüfen. Konfiguration, Zeit, Ergebnis und Bankunterschrift festhalten. Fehlgeschlagener Pilot stoppt die nächste Welle bis zur Prüfung."),
        bi("Planning allowance: two on-site days per site, twelve sites, one assessment. Bank supplies racks, power, cabling, access and identity settings. Training and asset handover are required; no completed survey or readiness sign-off exists.",
           "Planungsansatz: Zwei Vor-Ort-Tage je Standort, zwölf Standorte, ein Assessment. Bank liefert Racks, Strom, Kabel, Zugang und Identitätseinstellungen. Schulung und Anlagenübergabe nötig; abgeschlossene Messung und Bereitschaftsfreigabe fehlen."),
    ],
    "SUPPLY": [
        bi("Lifecycle baseline: CAP-6E is a legacy requested SKU; verify its current Atlas status and CAP-7 substitution live. CMS-48X-POE is the engineering candidate for 10G instead of requested CMS-48P-POE. Bank must approve substitutions; original demand must remain visible.",
           "Lebenszyklus: CAP-6E ist die verlangte Alt-SKU; aktuellen Atlas-Status und CAP-7-Austausch live prüfen. CMS-48X-POE ist Engineering-Kandidat für 10G statt CMS-48P-POE. Bank muss Austausch freigeben; Originalbedarf bleibt sichtbar."),
        bi("No stock or incoming date is frozen in the KB. Read visible On hand, Reserved and Available, then Check availability in the unsaved New quote. Plans are not reservations or commitments. Never Save quote or create an order.",
           "Kein Bestand oder Zugangstermin ist in der KB fixiert. Sichtbare On hand, Reserved und Available lesen; dann Check availability im ungespeicherten New quote. Pläne sind keine Reservierungen oder Zusagen. Niemals Save quote oder Auftrag erstellen."),
        bi("Delivery terms, transport charges, counterfeit/serial evidence and lifecycle dates require review when absent. Do not infer warranty or delivery guarantees from stock, API schema or a product name.",
           "Lieferbedingungen, Transportkosten, Fälschungs-/Seriennachweise und Lebenszyklusdaten benötigen bei Fehlen Prüfung. Bestand, API-Schema oder Produktname beweisen keine Garantie oder Lieferzusage."),
    ],
    "LEGAL": [
        bi("Approved negotiating position, not bank acceptance: ordinary liability capped at twelve months of service fees; privacy/confidentiality cap twenty-four months. Indirect loss excluded except where applicable mandatory law prevents exclusion. Unlimited liability is not accepted.",
           "Freigegebene Verhandlungsposition, keine Bankannahme: Allgemeine Haftung auf zwölf Monatsservicegebühren begrenzt; Datenschutz/Vertraulichkeit auf 24 Monate. Folgeschäden ausgeschlossen, soweit zwingendes Recht es erlaubt. Unbegrenzte Haftung wird nicht akzeptiert."),
        bi("Bank owns its data and configurations; Contoso retains pre-existing tooling. Transition assistance is scoped and chargeable. No governing-law agreement, delay penalty acceptance, indexation, free unlimited transition or signed warranty schedule is supplied.",
           "Bank besitzt Daten und Konfiguration; Contoso behält bestehende Werkzeuge. Übergangsunterstützung ist abgegrenzt und kostenpflichtig. Rechtswahl, Vertragsstrafenakzeptanz, Indexierung, unbegrenzter Gratisübergang und unterschriebene Garantieanlage fehlen."),
        bi("All exceptions require legal/commercial review. This package and response are fictional DRAFTs, not signature, contract acceptance, legal advice or an order. Confidential material must not be sent outside the authorized process.",
           "Alle Ausnahmen benötigen rechtliche/kaufmännische Prüfung. Paket und Antwort sind fiktive DRAFTs, keine Unterschrift, Vertragsannahme, Rechtsberatung oder Bestellung. Vertrauliches Material nicht außerhalb des freigegebenen Prozesses versenden."),
    ],
    "COMM": [
        bi("Buyer model: twelve sites; 24 switches, 48 access points, 48 optics, twelve gateways, 24 three-year care units, one assessment and 24 on-site days. The bank requires five years of ownership cost. Validate technical completeness; these quantities do not establish a complete 400-endpoint site design.",
           "Käufermodell: Zwölf Standorte; 24 Switches, 48 Access Points, 48 Optiken, zwölf Gateways, 24 Dreijahresbetreuungen, ein Assessment und 24 Vor-Ort-Tage. Die Bank verlangt fünf Jahre TCO. Technische Vollständigkeit prüfen; Mengen beweisen kein vollständiges 400-Endgeräte-Design."),
        bi("Approved demo operations rate: USD 240 per site per month for the current limited managed service, fixed for the five-year scenario; 12 months/year. No automatic annual increase. This KB service is not a Managed Workplace/Intune SKU in Atlas. Its tax treatment is unconfirmed.",
           "Freigegebener Demo-Betriebssatz: 240 USD je Standort und Monat für den begrenzten aktuellen Managed Service, im Fünfjahresszenario fix; zwölf Monate/Jahr. Kein automatischer Jahresaufschlag. Dies ist keine Atlas-Managed-Workplace/Intune-SKU. Steuerbehandlung ungeklärt."),
        bi("Hardware, care and implementation net prices come only from the actual customer's Atlas preview. Verify currency, discounts, unit rounding, tax and terms live. EUR list prices are not USD customer prices. Year-four/year-five care renewals, freight and non-Atlas tax remain open if unpriced.",
           "Hardware-, Betreuungs- und Einführungsnettopreise nur aus echter Atlas-Kundenvorschau. Währung, Rabatte, Positionsrundung, Steuer und Konditionen live prüfen. EUR-Listenpreise sind keine USD-Kundenpreise. Verlängerung Jahre vier/fünf, Fracht und Nicht-Atlas-Steuer bleiben ohne Preis offen."),
        bi("Internal review only: record actual visible margin and approval status without sending an approval request or exposing margin/cost to the customer. An approval-needed display is not approval. TCO formulas must show missing required amounts, not silently price them as zero.",
           "Nur interne Prüfung: Sichtbare Marge und Freigabestatus erfassen, ohne Antrag oder Offenlegung von Marge/Kosten an Kunden. Freigabebedarf ist keine Freigabe. TCO-Formeln müssen fehlende Pflichtbeträge statt stiller Nullpreise zeigen."),
    ],
}

SPECIAL = {
    "NET-01": ("alternative", "GAP"),
    "NET-02": ("capacity", "REVIEW"),
    "SEC-02": ("conditional", "GAP"),
    "DATA-01": ("privacy", "GAP"),
    "DATA-02": ("conditional", "REVIEW"),
    "DATA-04": ("missing-evidence", "OPEN"),
    "EVID-01": ("missing-evidence", "OPEN"),
    "EVID-02": ("conditional", "BLOCKED"),
    "OPS-01": ("conflict", "GAP"),
    "SUPPORT-01": ("renewal", "OPEN"),
    "SUPPLY-01": ("alternative", "REVIEW"),
    "SUPPLY-02": ("availability", "LIVE"),
    "LEGAL-01": ("contract", "GAP"),
    "LEGAL-02": ("conditional", "REVIEW"),
    "COMM-01": ("tco", "LIVE"),
    "COMM-02": ("tco", "OPEN"),
    "COMM-03": ("live-price", "LIVE"),
    "COMM-04": ("rounding", "LIVE"),
}

CONDITIONS = {
    "SEC-02": ("SEC-01", bi("If SSO is supported", "Wenn SSO unterstützt wird")),
    "DATA-02": ("DATA-01", bi("If processing is not exclusively EU", "Wenn Verarbeitung nicht ausschließlich in der EU erfolgt")),
    "EVID-02": ("EVID-01", bi("If certification is confirmed; unknown parent remains unresolved", "Wenn Zertifizierung bestätigt; unbekannte Hauptfrage bleibt offen")),
    "SUPPLY-02": ("SUPPLY-01", bi("If the requested or proposed line has insufficient stock", "Wenn angeforderte oder vorgeschlagene Position zu wenig Bestand hat")),
    "LEGAL-02": ("LEGAL-01", bi("If unlimited liability is not accepted", "Wenn unbegrenzte Haftung nicht akzeptiert wird")),
}

FORMATS = {
    "SEC-01": bi("Yes/No/Partial + scope", "Ja/Nein/Teilweise + Umfang"),
    "SEC-02": bi("Yes/No/Partial + measured minutes", "Ja/Nein/Teilweise + gemessene Minuten"),
    "EVID-01": bi("Yes/No/Unverified + evidence", "Ja/Nein/Ungeprüft + Nachweis"),
    "EVID-02": bi("Issuer, scope, number, expiry or unresolved dependency", "Aussteller, Umfang, Nummer, Ablauf oder offene Abhängigkeit"),
    "NET-02": bi("Watts, headroom and design exception", "Watt, Reserve und Planungsabweichung"),
    "DATA-01": bi("Yes/No/Partial + locations", "Ja/Nein/Teilweise + Orte"),
    "LEGAL-01": bi("Accept/Exception + reason", "Akzeptiert/Ausnahme + Begründung"),
    "COMM-01": bi("Amount/currency or INCOMPLETE; pricing-sheet references", "Betrag/Währung oder INCOMPLETE; Kalkulationsverweise"),
    "COMM-04": bi("Reconciliation by line; amount/currency", "Abgleich je Position; Betrag/Währung"),
}


def questions():
    result = []
    for topic, text in QUESTION_TEXT.items():
        lines = [line.strip() for line in text.strip().splitlines()]
        if len(lines) != 11:
            raise ValueError(f"{topic}: expected eleven questions")
        for n, line in enumerate(lines, 1):
            en, de = [s.strip() for s in line.split("||")]
            qid = f"{topic}-{n:02}"
            parent, condition = CONDITIONS.get(qid, ("", bi("Always", "Immer")))
            tag, expectation = SPECIAL.get(qid, ("standard", "EVIDENCE_OR_OPEN"))
            result.append({
                "id": qid, "topic": topic, "text": bi(en, de),
                "parent": parent, "condition": condition,
                "priority": "MUST" if n <= 8 else "SHOULD",
                "format": FORMATS.get(qid, bi("Concise text; evidence in source columns", "Kurzer Text; Nachweis in Quellenspalten")),
                "max_chars": 160 if qid in ("SEC-01", "EVID-01", "LEGAL-01") else 450,
                "evidence": bi(f"Approved KB-{topic} section and actual link; external test/certificate when requested",
                               f"Freigegebener KB-{topic}-Abschnitt und tatsächlicher Link; geforderter externer Test/Zertifikat"),
                "tag": tag, "expectation": expectation,
            })
    return result


COMPACT_IDS = {
    f"{topic}-{n:02}" for topic in TOPICS for n in (1, 2)
} | {"DATA-04", "COMM-03", "COMM-04", "EVID-08"}

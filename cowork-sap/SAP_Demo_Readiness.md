# SAP with arnold / SAP mit arnold

## English

This folder supports the four sections on **Cowork > SAP with arnold**.
The global `-Tenant` setup places five business requests from Teresa in the inbox and this reference
in **OneDrive > Demo-Kit > Cowork > cowork-sap**. It does not configure SAP, create
SAP master data, create sales orders, or install the existing `customer-email-html` skill.

| Scenario | Preparation | Evidence |
|---|---|---|
| Order quantity 1 to 50 | Prepared BLAZE order, customer 100231, quantity 1; keep its number and line item | Read the same SAP record back after updating it to 50 |
| Email to quote | Teresa's demo request; connected arnold plugin; existing customer-email-html skill | Intended request, SAP values and any quote number; actual skill output and email status |
| ARIA 2 orders | Intended SAP demo customer in the conversation; September 2025 data | Matching order numbers, material, quantities and the date field actually filtered |
| Stock | Intended customer in the conversation; ARIA 2, 10 EA, WSLL, next month's period | Supported stock/availability response; no implied reservation or delivery guarantee |

The order, quantity amendment and quote request explicitly use customer reference 100231. That does **not**
assign customer 100231 to every other scenario. Preserve each workflow's intended context.
The public order-search prompt uses the customer from the conversation instead of a real
company name. Test this variant separately.

Use **+ > Customize** to add the plugin and **Sources & Skills** to enable it.
An active enosix subscription, connected SAP ECC or S/4HANA and a Copilot license are required.
Finish connection sign-in before presenting.

For the separately tested order-creation opening and exact original prompts, use the private
presenter cue sheet. Microsoft lists quantity changes, order search, stock lookup and quote
creation; this reference does not claim that order creation is generally available.

Use the tested prompts without inventing a new skill. "Headless" is the wording of the demo
request, not a promise that consent, authentication or required approvals disappear.
If a workflow prepares or sends mail, check the recipient and actual state. A new run can
create another SAP document; do not rerun a write blindly or reset documents automatically.
Keep prepared results separate from files Cowork can search.

## Deutsch

Dieser Ordner gehört zu den vier Abschnitten auf **Cowork > SAP mit arnold**.
Das globale `-Tenant`-Setup legt fünf Geschäftsanfragen von Teresa im Posteingang ab und diese Referenz
unter **OneDrive > Demo-Kit > Cowork > cowork-sap**. Es konfiguriert kein SAP, erstellt
keine SAP-Stammdaten oder Aufträge und installiert nicht den vorhandenen Skill
`customer-email-html`.

| Szenario | Vorbereitung | Nachweis |
|---|---|---|
| Auftragsmenge 1 auf 50 | Vorbereiteter BLAZE-Auftrag, Kunde 100231, Menge 1; Nummer und Position festhalten | Denselben SAP-Beleg nach der Änderung auf 50 erneut lesen |
| Mail-to-Quote | Teresas Demo-Anfrage; verbundenes arnold-Plugin; vorhandener Skill customer-email-html | Richtige Anfrage, SAP-Werte und gegebenenfalls Angebotsnummer; tatsächliche Skill-Ausgabe und Mailstatus |
| ARIA-2-Aufträge | Beabsichtigter SAP-Demokunde im Verlauf; Daten aus September 2025 | Passende Belegnummern, Material, Mengen und tatsächlich gefiltertes Datumsfeld |
| Bestand | Beabsichtigter Kunde im Verlauf; ARIA 2, 10 EA, WSLL, Zeitraum des nächsten Monats | Belegte Bestands-/Verfügbarkeitsauskunft; keine implizite Reservierung oder Liefergarantie |

Auftrag, Mengennachtrag und Angebotsanfrage nennen ausdrücklich Kundenreferenz 100231. Damit ist **nicht**
für jedes andere Szenario Kunde 100231 gesetzt. Den jeweils beabsichtigten Gesprächskontext
beibehalten. Der öffentliche Rechercheprompt verwendet den Kunden aus dem Verlauf statt
eines echten Unternehmensnamens. Diese Variante separat prüfen.

Über **+ > Customize** das Plugin hinzufügen und in **Sources & Skills** aktivieren.
Aktive enosix-Subscription, verbundenes SAP ECC oder S/4HANA und Copilot-Lizenz sind erforderlich.
Die Verbindungsanmeldung vor der Vorführung abschließen.

Für den gesondert getesteten Einstieg zur Auftragsanlage und die exakten Originalprompts
das private Presenter-Cue-Sheet verwenden. Microsoft nennt Mengenänderungen, Auftragsrecherche,
Bestandsauskunft und Angebotserstellung; diese Referenz verspricht keine allgemeine
Verfügbarkeit der Auftragsneuanlage.

Die getesteten Prompts verwenden, keinen Ersatz-Skill erfinden. „Headless“ ist die Formulierung
des Demo-Auftrags, kein Versprechen, dass Anmeldung, Zustimmung oder notwendige Freigaben
entfallen. Bei vorbereiteten oder gesendeten Mails Empfänger und tatsächlichen Status prüfen.
Ein neuer Durchlauf kann einen weiteren SAP-Beleg erstellen; Schreibvorgänge nicht blind
wiederholen und Belege nicht automatisch zurücksetzen. Vorbereitete Ergebnisse von Dateien
trennen, die Cowork durchsuchen kann.

## Business context / Geschäftskontext

EN: The five inbox messages cover the BLAZE order, its later amendment, the ARIA 2 quote,
the September 2025 order review and the WSLL supply question. Use the amendment only after
the order exists. The messages contain no invented order numbers, prices or stock results.
The presenter supplies SAP; the setup verifies Teresa's currently enabled Exchange plan.

DE: Die fünf Posteingangsmails behandeln den BLAZE-Auftrag, den späteren Mengennachtrag,
das ARIA-2-Angebot, den Auftragsreview September 2025 und die WSLL-Bestandsfrage. Den Nachtrag
erst nach Auftragsanlage verwenden. Keine erfundenen Auftragsnummern, Preise oder Bestände.
SAP wird vom Presenter bereitgestellt; das Setup prüft Teresas aktuell aktivierten Exchange-Plan.

## Official sources / Offizielle Quellen

- https://learn.microsoft.com/en-us/connectors/enosixarnoldcopilotc/
- https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-available-plugins
- https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-plugins
- https://learn.microsoft.com/microsoft-365/copilot/cowork/use-cowork
- https://learn.microsoft.com/graph/api/user-get?view=graph-rest-1.0
- https://learn.microsoft.com/graph/api/resources/assignedplan?view=graph-rest-1.0

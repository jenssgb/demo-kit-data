# Browser: Lieferverzug bis Rettungsmaßnahme

- Dateien: OneDrive > Demo-Kit > Cowork > cowork-atlas-browser-supply. Teresa hat die Mail **Proseware | Rotterdam-Switches 25 Tage verspätet** im Posteingang.
- Admin, einmalig im Microsoft 365 Admin Center: **Copilot > Settings > View All > Cowork settings > Allow browser access**. **Allow Cowork to use the Microsoft Edge browser to perform tasks on behalf of users** wählen; Website- und Edge-Richtlinien prüfen. Nicht per Skript ändern.
- **Cowork** im Web in **Microsoft Edge** mit demselben Arbeitskonto nutzen; Microsoft empfiehlt Edge 152.0.4191.53 oder neuer. https://contoso-atlas.swedencentral.cloudapp.azure.com in diesem Profil öffnen und **Sign in with Microsoft** vor der Vorführung abschließen.
- Vor der Vorführung https://contoso-atlas.swedencentral.cloudapp.azure.com/guide auf der VM öffnen und Erreichbarkeit, Rechte deines Kontos, aktuelle Schaltflächennamen und benötigte Datensätze prüfen. Atlas ist eine separate Demo-App, keine Microsoft-Produktfunktion. MCP-Setup, Acting-User-Schalter und API-Anweisungen des Guides nicht verwenden.
- **skills/contoso-atlas-browser-supply/SKILL.md** über **Customize > Skills > Pfeil neben Add > Upload skill** hochladen. Falls bereits vorhanden, den Skill öffnen und seine Anweisungen bearbeiten, keine weitere Kopie hochladen. In **Sources** mit Mail und **inputs/de/request.md** auswählen. Keinen Atlas-CRM-Connector auswählen.
- Beim ersten Browserlauf **I understand** prüfen. Mit **Switch to tab** die Arbeit zeigen; folgenreiche Aktionen auf Nachfrage freigeben.
- In der angemeldeten VM proben. Diese Paketdateien befüllen Atlas nicht und setzen es nicht zurück. Kein Reset der gesamten gemeinsamen Datenbank; aktuelle IDs/Termine festhalten und geprüfte Zielzustände bei Wiederholungen wiederverwenden.

Prosewares 600 PoE-Switches für Rotterdam kommen 25 Tage später. Prüfe im Atlas-Browser die betroffenen Kundenaufträge und Rettungsoptionen.

Trag die Verspätung ein und wende die belegte Rettungsoption für Blue Yonder an. Bereite die Kundeninfos als Entwürfe vor.

- In Atlas **Purchase Orders** wählen und die Proseware-Bestellung für Rotterdam mit **CMS-48P-POE** öffnen. **600** Stück, ursprüngliches Lieferdatum und aktuellen Status prüfen.
- Unter **Lines and dates** bei **Expected** das ursprüngliche Lieferdatum plus **25 Kalendertage** eintragen. Vor dem Speichern **Impact analysis** lesen.
- Betroffene Auftragslinks, Risiko und angezeigte Rettungsoption vergleichen. Genaue Bestellung/Datum vor **Mark PO delayed** prüfen; den freigegebenen **Delay reason** eingeben.
- **Apply recovery** nur für die geprüfte Zeile **Blue Yonder Airlines** mit ausreichender Rettungsoption nutzen. Bestellung und Kundenauftrag erneut öffnen und gespeicherte Termine prüfen. Kundeninfos bleiben Entwürfe.

Ein gespeicherter Bestellverzug, eine geprüfte Blue-Yonder-Rettungsmaßnahme, erneut gelesene Liefertermine und ungesendete Kundenentwürfe.

Microsoft documentation: https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-local-browser; https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-customize; https://learn.microsoft.com/microsoft-365/copilot/responsible-ai/cowork-responsible-ai-faq
Demo application guide (verify on the VM): https://contoso-atlas.swedencentral.cloudapp.azure.com/guide

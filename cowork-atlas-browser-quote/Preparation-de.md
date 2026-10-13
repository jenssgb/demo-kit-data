# Browser: Angebot bis Freigabe und PDF

- Dateien: OneDrive > Demo-Kit > Cowork > cowork-atlas-browser-quote. Teresa hat die Mail **Woodgrove Bank | Angebot über 10 PoE-Switches** im Posteingang.
- Admin, einmalig im Microsoft 365 Admin Center: **Copilot > Settings > View All > Cowork settings > Allow browser access**. **Allow Cowork to use the Microsoft Edge browser to perform tasks on behalf of users** wählen; Website- und Edge-Richtlinien prüfen. Nicht per Skript ändern.
- **Cowork** im Web in **Microsoft Edge** mit demselben Arbeitskonto nutzen; Microsoft empfiehlt Edge 152.0.4191.53 oder neuer. https://contoso-atlas.swedencentral.cloudapp.azure.com in diesem Profil öffnen und **Sign in with Microsoft** vor der Vorführung abschließen.
- Vor der Vorführung https://contoso-atlas.swedencentral.cloudapp.azure.com/guide auf der VM öffnen und Erreichbarkeit, Rechte deines Kontos, aktuelle Schaltflächennamen und benötigte Datensätze prüfen. Atlas ist eine separate Demo-App, keine Microsoft-Produktfunktion. MCP-Setup, Acting-User-Schalter und API-Anweisungen des Guides nicht verwenden.
- **skills/contoso-atlas-browser-quote/SKILL.md** über **Customize > Skills > Pfeil neben Add > Upload skill** hochladen. Falls bereits vorhanden, den Skill öffnen und seine Anweisungen bearbeiten, keine weitere Kopie hochladen. In **Sources** auswählen. Vollständigen Prompt oder Teresas eindeutig benannte Posteingangsmail nutzen; request.md muss nicht angehängt werden. Keinen Atlas-CRM-Connector auswählen.
- Beim ersten Browserlauf **I understand** prüfen. Mit **Switch to tab** die Arbeit zeigen; folgenreiche Aktionen auf Nachfrage freigeben.
- In der angemeldeten VM proben. Diese Paketdateien befüllen Atlas nicht und setzen es nicht zurück. Kein Reset der gesamten gemeinsamen Datenbank; aktuelle IDs/Termine festhalten und geprüfte Zielzustände bei Wiederholungen wiederverwenden.

Öffne Contoso Atlas im Browser und erstelle für Woodgrove Bank ein Angebot über 10 Contoso Mesh Switch 48-port PoE+. Prüfe Preis und Verfügbarkeit.

Reiche dieses Angebot zur Freigabe ein und lade sein PDF herunter. Zeig mir den gespeicherten Status.

- In Atlas **Quotes › New quote** wählen. Unter **Account search** nach **Woodgrove Bank** suchen und das genaue Konto auswählen.
- **CMS-48P-POE** suchen, Beschreibung prüfen und **Qty** auf **10** setzen. **Check availability** nutzen und **Live preview** prüfen.
- Konto, Positionen, Preis, Steuer, Marge und Freigaberegel vor **Save quote** prüfen. Die erzeugte Angebots-ID festhalten und erneut öffnen.
- Am selben Angebot **Submit for approval**, dann **Preview PDF** / **Download PDF** nutzen. Gespeicherten Status prüfen; ein ausstehendes Angebot bleibt im PDF **DRAFT**.

Ein gespeichertes Atlas-Angebot mit tatsächlicher ID, geprüftem Konto/SKU/Menge und einem PDF passend zum echten Freigabestatus.

Microsoft documentation: https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-local-browser; https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-customize; https://learn.microsoft.com/microsoft-365/copilot/responsible-ai/cowork-responsible-ai-faq
Demo application guide (verify on the VM): https://contoso-atlas.swedencentral.cloudapp.azure.com/guide

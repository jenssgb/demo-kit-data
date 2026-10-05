# Governance Copilot demo setup / Demo-Setup Governance Copilot

All files are fictional Contoso data for the Fargo distribution center expansion. Do not replace them with real customer data.

## English setup

1. Install the demo data with the Demo Kit one-liner. The folder should appear as **OneDrive > Demo-governance-copilot**.
2. In Excel, open **Fargo_Labor_Costs_Confidential.xlsx**. On **Home > Sensitivity**, apply your tenant's **Confidential** label. Save the file to OneDrive or SharePoint.
3. In Word, open **Contoso_Fargo_Public_Site_Fact_Sheet.docx**. On **Home > Sensitivity**, apply **Public**. Save it.
4. In Word, open **Contoso_Fargo_Highly_Confidential_MA_Note.docx**. On **Home > Sensitivity**, apply **Highly Confidential**. Save it.
5. Ask your Purview admin to create or confirm a DLP policy for the **Microsoft 365 Copilot and Copilot Chat** location that prevents Copilot from processing content with the **Highly Confidential** label.
6. For the admin-view part, prepare accounts/permissions for Microsoft 365 admin center, SharePoint admin center, Microsoft Purview portal, and Viva Insights/Copilot Dashboard as needed.

## Deutsches Setup

1. Demodaten mit dem One-Liner aus dem Demo Kit installieren. Der Ordner sollte als **OneDrive > Demo-governance-copilot** erscheinen.
2. In Excel **Fargo_Labor_Costs_Confidential.xlsx** öffnen. Unter **Start > Vertraulichkeit** (oder **Sensitivity**, je nach UI-Sprache) das Tenant-Label **Confidential/Vertraulich** anwenden. In OneDrive oder SharePoint speichern.
3. In Word **Contoso_Fargo_Public_Site_Fact_Sheet.docx** öffnen. Unter **Start > Vertraulichkeit/Sensitivity** das Label **Public/Öffentlich** anwenden. Speichern.
4. In Word **Contoso_Fargo_Highly_Confidential_MA_Note.docx** öffnen. Unter **Start > Vertraulichkeit/Sensitivity** das Label **Highly Confidential/Hoch vertraulich** anwenden. Speichern.
5. Die Purview-Administration soll eine DLP-Richtlinie für den Ort **Microsoft 365 Copilot and Copilot Chat** erstellen oder bestätigen, die Copilot das Verarbeiten von Inhalten mit dem Label **Highly Confidential/Hoch vertraulich** verbietet.
6. Für den Admin-Teil Konten/Berechtigungen für Microsoft 365 Admin Center, SharePoint Admin Center, Microsoft Purview Portal und Viva Insights/Copilot Dashboard vorbereiten.

## Files / Dateien

- **Fargo_Labor_Costs_Confidential.xlsx** — cost forecast, apply Confidential.
- **Contoso_Fargo_Public_Site_Fact_Sheet.docx** — harmless public grounding file, apply Public.
- **Contoso_Fargo_Highly_Confidential_MA_Note.docx** — blocked-file demo, apply Highly Confidential.

# Browser: quote to approval and PDF

- Files: OneDrive > Demo-Kit > Cowork > cowork-atlas-browser-quote. The inbox message is **Woodgrove Bank | Quote for 10 PoE switches** from Teresa.
- Admin, once in the Microsoft 365 admin center: **Copilot > Settings > View All > Cowork settings > Allow browser access**. Select **Allow Cowork to use the Microsoft Edge browser to perform tasks on behalf of users**; check site and Edge policies. Do not change these by script.
- Use **Cowork** on the web in **Microsoft Edge** with the same work account; Microsoft recommends Edge 152.0.4191.53 or later. Open https://contoso-atlas.swedencentral.cloudapp.azure.com in that profile and complete **Sign in with Microsoft** before presenting.
- Before presenting, open https://contoso-atlas.swedencentral.cloudapp.azure.com/guide on the VM and verify website availability, your account's permissions, current button labels and this workflow's records. Atlas is a separate demo app, not a Microsoft product capability. Do not use the guide's MCP setup, acting-user switch or API instructions.
- Upload **skills/contoso-atlas-browser-quote/SKILL.md** via **Customize > Skills > arrow next to Add > Upload skill**. If it already exists, open that skill and edit its instructions instead of uploading another copy. Select it in **Sources** with the mail and **inputs/en/request.md**. Do not select an Atlas CRM connector.
- On first browser use, review **I understand**. Use **Switch to tab** to show the work; approve consequential actions when requested.
- Rehearse in the signed-in VM. These package files do not seed/reset Atlas. Do not reset the entire shared database; record current IDs/dates and reuse verified target states on repeats.

Open Contoso Atlas in the browser and create a quote for Woodgrove Bank for 10 Contoso Mesh Switch 48-port PoE+. Check pricing and availability.

Submit this quote for approval and download its PDF. Show me the saved status.

- In Atlas, select **Quotes › New quote**. Search for **Woodgrove Bank** under **Account search** and select the exact account.
- Search **CMS-48P-POE**, check the description and set **Qty** to **10**. Use **Check availability** and review **Live preview**.
- Review the account, lines, price, tax, margin and approval policy before **Save quote**. Record the resulting quote ID and reopen it.
- On the same quote, use **Submit for approval**, then **Preview PDF** / **Download PDF**. Check the saved status; a pending quote's PDF stays **DRAFT**.

A saved Atlas quote with its actual ID, checked account/SKU/quantity and a PDF reflecting the real approval status.

Microsoft documentation: https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-local-browser; https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-customize; https://learn.microsoft.com/microsoft-365/copilot/responsible-ai/cowork-responsible-ai-faq
Demo application guide (verify on the VM): https://contoso-atlas.swedencentral.cloudapp.azure.com/guide

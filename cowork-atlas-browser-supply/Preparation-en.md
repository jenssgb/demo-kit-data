# Browser: supply delay to recovery

- Files: OneDrive > Demo-Kit > Cowork > cowork-atlas-browser-supply. The inbox message is **Proseware | Rotterdam switches 25 days late** from Teresa.
- Admin, once in the Microsoft 365 admin center: **Copilot > Settings > View All > Cowork settings > Allow browser access**. Select **Allow Cowork to use the Microsoft Edge browser to perform tasks on behalf of users**; check site and Edge policies. Do not change these by script.
- Use **Cowork** on the web in **Microsoft Edge** with the same work account; Microsoft recommends Edge 152.0.4191.53 or later. Open https://contoso-atlas.swedencentral.cloudapp.azure.com in that profile and complete **Sign in with Microsoft** before presenting.
- Before presenting, open https://contoso-atlas.swedencentral.cloudapp.azure.com/guide on the VM and verify website availability, your account's permissions, current button labels and this workflow's records. Atlas is a separate demo app, not a Microsoft product capability. Do not use the guide's MCP setup, acting-user switch or API instructions.
- Upload **skills/contoso-atlas-browser-supply/SKILL.md** via **Customize > Skills > arrow next to Add > Upload skill**. If it already exists, open that skill and edit its instructions instead of uploading another copy. Select it in **Sources** with the mail and **inputs/en/request.md**. Do not select an Atlas CRM connector.
- On first browser use, review **I understand**. Use **Switch to tab** to show the work; approve consequential actions when requested.
- Rehearse in the signed-in VM. These package files do not seed/reset Atlas. Do not reset the entire shared database; record current IDs/dates and reuse verified target states on repeats.

Proseware's 600 PoE switches for Rotterdam are 25 days late. Check the affected customer orders and recovery options in the Atlas browser.

Record the delay and apply the verified recovery for Blue Yonder. Prepare the customer updates as drafts.

- In Atlas, select **Purchase Orders** and open the Proseware PO for Rotterdam and **CMS-48P-POE**. Check the **600** units, original expected date and current status.
- Under **Lines and dates**, set **Expected** to the original expected date plus **25 calendar days**. Read **Impact analysis** before saving.
- Compare the affected order links, risk and displayed recovery. Review the exact PO/date before **Mark PO delayed**; enter the approved **Delay reason**.
- Use **Apply recovery** only for the verified **Blue Yonder Airlines** row with sufficient recovery. Reopen both PO and customer order to check the stored dates. Keep notifications as drafts.

A saved PO delay, a verified Blue Yonder recovery, reopened delivery dates and unsent customer-update drafts.

Microsoft documentation: https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-local-browser; https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-customize; https://learn.microsoft.com/microsoft-365/copilot/responsible-ai/cowork-responsible-ai-faq
Demo application guide (verify on the VM): https://contoso-atlas.swedencentral.cloudapp.azure.com/guide

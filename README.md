# Demo Kit – Demo data

Public sample files for the demos in the [Demo Kit](https://github.com/jenssgb) – one folder per demo.
All content is fictional (**Contoso**). No customer data, no credentials.

## Get the files onto a demo machine

Open **PowerShell** on the demo VM and run (replace the demo id):

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Demo onenote-copilot
```

The files land in `OneDrive\Demo-<id>` (work/school OneDrive first, then personal OneDrive, otherwise the Desktop)
and the folder opens. Works with Windows PowerShell 5.1 and PowerShell 7, no admin rights, no modules.
Optional: `-Target <folder>`, `-Branch <branch>`.

## Demos

| Demo id | Content | Source / license |
| --- | --- | --- |
| `onenote-copilot` | Contoso Fargo distribution center expansion: OneNote notebook (.onepkg), Word and Excel files | Microsoft Learn course MS-4004, MIT (see `onenote-copilot/LICENSE-MS-4004.txt`) |
| `powerpoint-copilot` | Fictional Contoso brand guidelines (.docx + .pdf for brand kits), logos, brand background, a branded deck, an off-brand draft deck (Fargo scenario) and a sample PowerPoint skill (`skills/contoso-site-update/SKILL.md`) | Created for the Demo Kit, rebuild with `python tools/build-powerpoint-copilot.py` |
| `cowork-copilot` | Fargo go-live for Copilot Cowork: four emails as .eml (EN + DE, incl. a Tailspin Toys quote request for the SAP/arnold part), readiness tracker (.xlsx), construction weekly week 16 (.docx), Cowork skill (`skills/contoso-golive-readiness/SKILL.md`), SAP placeholder list | Created for the Demo Kit, rebuild with `python tools/build-cowork-copilot.py` |
| `word-excel-copilot` | Edit with Copilot in Word and Excel (Fargo): runbook source pack (.docx), inbound shipment workbook (~400 rows) for the Dock 2 capacity analysis, four Fargo emails (.eml) | Created for the Demo Kit, rebuild with `python tools/build-word-excel-copilot.py` |
| `researcher-copilot` | Researcher + Vision: Contoso Fargo floor plan (Dock 2/3, Gate 4, new wing) as PNG; the rest reuses `cowork-copilot` | Created for the Demo Kit, rebuild with `python tools/build-researcher-copilot.py` |
| `notebooks-copilot` | Copilot Notebooks: Dock 3 go-live sync transcript (.docx/.vtt) and a night-shift handover template | Created for the Demo Kit, rebuild with `python tools/build-notebooks-copilot.py` |
| `prompts-copilot` | Contoso Fargo prompt pack (.docx + .md) for Prompt Gallery, scheduled prompts, memory and Pages | Created for the Demo Kit, rebuild with `python tools/build-prompts-copilot.py` |
| `agent-builder-copilot` | Agent Builder "Fargo Site Assistant": safety handbook (.docx), dock schedule (.csv/.xlsx for a SharePoint list), paste-ready agent instructions | Created for the Demo Kit, rebuild with `python tools/build-agent-builder-copilot.py` |
| `governance-copilot` | Governance & trust: Confidential labor cost workbook, public site fact sheet, Highly Confidential M&A note, label/DLP setup notes (`README_setup.md`) | Created for the Demo Kit, rebuild with `python tools/build-governance-copilot.py` |

## Add a demo

1. New folder `<demo-id>/` with the files.
2. Create `manifest.json` with `title` (en/de) and `next` (en/de) – copy one from an existing demo.
3. Run `.\tools\Update-Manifest.ps1 -Demo <demo-id>` to fill the file list.
4. Commit and push. In the Demo Kit deck set `"files": { "demoData": { "demo": "<demo-id>" } }`.

---

# Deutsch

Öffentliche Beispieldateien für die Demos im Demo Kit, ein Ordner pro Demo. Alle Inhalte sind fiktiv (**Contoso**).

**Dateien auf die Demo-VM holen:** PowerShell öffnen, den One-Liner oben einfügen (Demo-ID anpassen), **Enter**.
Die Dateien landen in `OneDrive\Demo-<id>` (ohne OneDrive auf dem Desktop), danach öffnet sich der Ordner.

**Neue Demo:** Ordner anlegen, `manifest.json` (Titel und nächste Schritte auf en/de) anlegen,
`.\tools\Update-Manifest.ps1 -Demo <id>` ausführen, pushen und im Deck `files.demoData.demo` setzen.

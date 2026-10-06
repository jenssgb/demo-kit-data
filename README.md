# Demo Kit – Demo data

Public sample files for the demos in the [Demo Kit](https://github.com/jenssgb) – one folder per demo.
All content is fictional (**Contoso**). No customer data, no credentials.

## Get the files onto a demo machine

Open **PowerShell** on the demo VM and run **one command for the whole customer deck** (bundle):

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Bundle bpw
```

All demos of the deck land in `OneDrive\Demo-BPW\<demo-id>` (work/school OneDrive first, then personal OneDrive,
otherwise the Desktop). Each demo folder is self-contained: files that several demos use are copied into each of them.
**Run it again any time** – new demos and changed files are downloaded, unchanged files are skipped (SHA-256).
The command never changes; new demos are added to the bundle.

One demo only, into the same bundle folder: `... -Demo cowork-copilot -Bundle bpw`.
One demo on its own (folder `Demo-<id>`): `... -Demo onenote-copilot`.
Works with Windows PowerShell 5.1 and PowerShell 7, no admin rights, no modules.
Optional: `-Target <folder>`, `-Branch <branch>`, `-NoExplorer`.

## Straight into the tenant (`-Tenant`)

On a demo VM signed in as the demo admin (e.g. **MOD Administrator** in a CDX tenant) the same command with `-Tenant`
also creates what the demos need **in the tenant** – no manual mailing, no .onepkg import:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Bundle bpw -Tenant
```

| What | How (official APIs) |
| --- | --- |
| Files | uploaded to the signed-in user's OneDrive (`Demo-BPW\<demo-id>`) with Microsoft Graph; unchanged files are skipped by `quickXorHash` |
| Mails | real received mails from the right senders: Graph `sendMail` with `from` (delegated `Mail.Send.Shared` + Exchange **Send As**) |
| Senders | existing mailboxes by display name (Teresa Sac, Billie Vester); people not in the tenant become **shared mailboxes** (Exchange Online PowerShell). A tenant user **without a mailbox** (e.g. Teams license only) is an error, never a look-alike mailbox |
| Meetings | Graph `POST /me/events` with attendees (next matching weekday) |
| OneNote | notebook, section and pages via Graph (delegated) |
| Not possible by API | sensitivity labels and DLP, .onepkg import, brand kit, skills upload, cost policy, Cowork browser access, Teams meeting with transcript – printed as a checklist at the end |

- Two sign-ins: Microsoft Graph (consent to the listed permissions) and Exchange Online. Needs Exchange admin rights (Global Admin is fine).
- `-Language de` sends the German mails (default `en`). `-WhatIf` shows everything without changing the tenant.
- Modules `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` are installed for the current user if missing.
- "Send As" on new shared mailboxes can take a while; the script waits up to 20 minutes, otherwise it says "run again later".
- Run it again any time: nothing is created twice. Every run (with or without `-Tenant`) writes a log to `Desktop\DemoKit-Logs\<bundle>-<time>.log` (last 20 kept) – send that file when something goes wrong.
- Set-up senders are cached locally and in the admin's OneDrive (`DemoKit/senders.json`): the Exchange sign-in (second login) is needed only once per tenant, and then prepares every person of every demo (`people.json`).
- After each run: `Desktop\Demo-<Customer> - Links.html` with the OneDrive links of all files, OneNote, extra links (e.g. Contoso Atlas) and the manual checklist. In a terminal started **as administrator**, the same links also appear as Edge favorites (folder **Demo Kit**, locked; the next run updates it).
- Each demo describes its tenant data in `<demo-id>/tenant.json` (mails, events, onenote, manual); extra files live in `<demo-id>/tenant/`.
  Neither is part of the file download.

## Demo people = real users of the CDX demo tenant

All mails, prompts, files and meetings use the **real users of the CDX demo tenant M365CPI98544940** directly
(Teresa Sac, Vance DeLeon, Sonia Rees, Billie Vester, Sydney Mattos – addresses `<alias>@M365CPI98544940.OnMicrosoft.com`).
Only use users **with an Exchange mailbox** (license with Exchange Online, e.g. Microsoft 365 E5) – Teams-only users can't send mails.
There is no mapping layer: what you see in the files is what's in the tenant. `-Tenant` finds them by display name and
grants "Send As"; only the external Tailspin Toys customer (Jordan Mitchell) becomes a shared mailbox.
New demo tenant? Rewrite the names once in the files and `tools/build-*.py`, then `Update-Manifest.ps1 -All`.

## Bundles

| Bundle id | Deck | Folder |
| --- | --- | --- |
| `bpw` | `bpw-ai-multiplikatoren` (BPW – AI Champions) | `Demo-BPW` |
| `dhl` | `dhl-innovation-briefing` (DHL Innovation Briefing – Copilot Cowork live) | `Demo-DHL` |

## Demos

| Demo id | Content | Source / license |
| --- | --- | --- |
| `onenote-copilot` | Contoso Fargo distribution center expansion: OneNote notebook (.onepkg), Word and Excel files | Microsoft Learn course MS-4004, MIT (see `onenote-copilot/LICENSE-MS-4004.txt`) |
| `powerpoint-copilot` | Fictional Contoso brand guidelines (.docx + .pdf for brand kits), logos, brand background, a branded deck, an off-brand draft deck (Fargo scenario) and a sample PowerPoint skill (`skills/contoso-site-update/SKILL.md`) | Created for the Demo Kit, rebuild with `python tools/build-powerpoint-copilot.py` |
| `cowork-copilot` | Fargo go-live for Copilot Cowork: four emails as .eml (EN + DE, incl. a Tailspin Toys quote request for the SAP/arnold part), readiness tracker (.xlsx), construction weekly week 16 (.docx), Cowork skill (`skills/contoso-golive-readiness/SKILL.md`), SAP placeholder list | Created for the Demo Kit, rebuild with `python tools/build-cowork-copilot.py` |
| `word-excel-copilot` | Edit with Copilot in Word and Excel (Fargo): runbook source pack (.docx), inbound shipment workbook (~400 rows) for the Dock 2 capacity analysis, four Fargo emails (.eml) | Created for the Demo Kit, rebuild with `python tools/build-word-excel-copilot.py` |
| `researcher-copilot` | Researcher + Vision: Contoso Fargo floor plan (Dock 2/3, Gate 4, new wing) as PNG; the cowork-copilot tracker, weekly notes and emails are copied in via `shared.json` | Created for the Demo Kit, rebuild with `python tools/build-researcher-copilot.py` |
| `notebooks-copilot` | Copilot Notebooks: Dock 3 go-live sync transcript (.docx/.vtt) and a night-shift handover template; cowork-copilot and brand-guideline files copied in via `shared.json` | Created for the Demo Kit, rebuild with `python tools/build-notebooks-copilot.py` |
| `prompts-copilot` | Contoso Fargo prompt pack (.docx + .md) for Prompt Gallery, scheduled prompts, memory and Pages | Created for the Demo Kit, rebuild with `python tools/build-prompts-copilot.py` |
| `agent-builder-copilot` | Agent Builder "Fargo Site Assistant": safety handbook (.docx), dock schedule (.csv/.xlsx for a SharePoint list), paste-ready agent instructions | Created for the Demo Kit, rebuild with `python tools/build-agent-builder-copilot.py` |
| `governance-copilot` | Governance & trust: Confidential labor cost workbook, public site fact sheet, Highly Confidential M&A note, label/DLP setup notes (`README_setup.md`) | Created for the Demo Kit, rebuild with `python tools/build-governance-copilot.py` |

## Add a demo

1. New folder `<demo-id>/` with the files.
2. Create `manifest.json` with `title` (en/de) and `next` (en/de) – copy one from an existing demo.
3. Optional: `shared.json` – list of `"<other-demo>/<path>"` files this demo also needs (copied into its folder).
4. Run `.\tools\Update-Manifest.ps1 -Demo <demo-id>` (or `-All`) – fills the file list with SHA-256 hashes.
   Run it again after every `tools/build-*.py` run.
5. Add the demo to the deck (`"demoData": { "demo": "<demo-id>" }`) and update the bundle:
   `.\tools\Update-Bundle.ps1 -Bundle bpw -Deck <path>\decks\bpw-ai-multiplikatoren.json`
6. Needs data in the tenant (mails, meetings, OneNote, manual steps)? Add `<demo-id>/tenant.json` – copy `onenote-copilot/tenant.json`.
7. Commit and push. The bundle one-liner stays the same.

New customer deck: `Update-Bundle.ps1 -Bundle <id> -Deck <deck.json> -Folder Demo-<Customer>`, then set
`"files": { "demoData": { "demo": "<first-demo>", "bundle": "<id>" } }` in the deck.

The repo uses `.gitattributes` `* -text` so files are stored byte for byte and the hashes match the raw downloads.

---

# Deutsch

Öffentliche Beispieldateien für die Demos im Demo Kit, ein Ordner pro Demo. Alle Inhalte sind fiktiv (**Contoso**).

**Dateien auf die Demo-VM holen:** PowerShell öffnen, den Bundle-Befehl oben einfügen (z. B. `-Bundle bpw`), **Enter**.
Alle Demos des Kundendecks landen in `OneDrive\Demo-BPW\<demo-id>` (ohne OneDrive auf dem Desktop), danach öffnet sich der Ordner.
Jederzeit erneut ausführen: neue Demos und geänderte Dateien kommen dazu, unveränderte werden übersprungen. Der Befehl bleibt immer gleich.

**Direkt in den Tenant:** Auf der Demo-VM als Admin (z. B. MOD Administrator im CDX-Tenant) denselben Befehl mit `-Tenant` ausführen.
Das Skript lädt die Dateien per Graph in den OneDrive des angemeldeten Benutzers. Es legt die Mails mit den echten Absendern an
(fehlende Personen als freigegebene Postfächer mit „Senden als“), dazu den Termin und das OneNote-Notizbuch. Am Ende listet es, was noch
von Hand zu tun ist (Bezeichnungen/DLP, Brand Kit, Skills, Cowork-Browserzugriff). `-Language de` für deutsche Mails, `-WhatIf` zum Ausprobieren.

**Demo-Personen = echte Benutzer des CDX-Tenants** (Teresa Sac, Vance DeLeon, Sonia Rees, Billie Vester, Sydney Mattos) – direkt
in allen Mails, Prompts und Dateien, ohne Zuordnung. Nur der externe Kunde (Jordan Mitchell, Tailspin Toys) wird ein freigegebenes Postfach. Nur Benutzer **mit Postfach** nehmen
(Lizenz mit Exchange Online, z. B. E5). Jeder Lauf schreibt ein Log nach `Desktop\DemoKit-Logs` – bei Fehlern diese Datei schicken.
Die Exchange-Anmeldung (zweites Login) ist nur einmal pro Tenant nötig (Absender-Cache auch im Admin-OneDrive, `DemoKit/senders.json`).
Nach jedem Lauf liegt `Desktop\Demo-<Kunde> - Links.html` mit allen OneDrive-Links, OneNote, Zusatzlinks (z. B. Contoso Atlas) und der
Checkliste bereit. Im Terminal **als Administrator** kommen dieselben Links als Edge-Favoriten (Ordner **Demo Kit**, gesperrt; nächster Lauf aktualisiert).

**Neue Demo:** Ordner anlegen, `manifest.json` (Titel und nächste Schritte auf en/de) anlegen, optional `shared.json`,
`.\tools\Update-Manifest.ps1 -Demo <id>` ausführen, im Deck `demoData.demo` setzen,
`.\tools\Update-Bundle.ps1 -Bundle <bundle> -Deck <deck.json>` ausführen, bei Tenant-Daten `tenant.json` pflegen und pushen.

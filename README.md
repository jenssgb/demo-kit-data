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
| Senders | existing users by display name (Megan Bowen, Alex Wilber); missing people become **shared mailboxes** (Exchange Online PowerShell) |
| Meetings | Graph `POST /me/events` with attendees (next matching weekday) |
| OneNote | notebook, section and pages via Graph (delegated) |
| Not possible by API | sensitivity labels and DLP, .onepkg import, brand kit, skills upload, cost policy, Cowork browser access, Teams meeting with transcript – printed as a checklist at the end |

- Two sign-ins: Microsoft Graph (consent to the listed permissions) and Exchange Online. Needs Exchange admin rights (Global Admin is fine).
- `-Language de` sends the German mails (default `en`). `-WhatIf` shows everything without changing the tenant.
- Modules `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` are installed for the current user if missing.
- "Send As" on new shared mailboxes can take a while; the script waits up to 20 minutes, otherwise it says "run again later".
- Run it again any time: nothing is created twice. Log: `%LOCALAPPDATA%\DemoKit\logs\tenant-<time>.log`.
- Each demo describes its tenant data in `<demo-id>/tenant.json` (mails, events, onenote, manual); extra files live in `<demo-id>/tenant/`.
  Neither is part of the file download.

## Tenant casts (profiles)

Demo people (Megan Bowen, Riley Johnson, …) can be played by real users of a demo tenant, so mails come from real
mailboxes and Copilot finds the people. The deck defines `personas` (default names) and `profiles.<id>.people`
(real users); `files.demoData.profile` is the cast the bundle installs by default.

- `profiles/<id>.json` – written by `Update-Bundle.ps1` from the deck.
- `profiles/<id>/<demo-id>/` – copy of every demo with the names and e-mail addresses swapped (text, .eml, Word/Excel/PowerPoint).
  Rebuild after every change to demo files or the cast:
  `python tools\build_profile.py <id>; .\tools\Update-Manifest.ps1 -All -Root profiles\<id>`.
  The build fails if a default name is left over (e.g. split across Word runs). File names and binaries (.onepkg, .pdf) stay unchanged.
- `install.ps1` uses the bundle's profile; `-Profile <id>` picks another one, `-Profile default` the Contoso names.
- `-Tenant` resolves senders by display name: real users get "Send As", only people missing in the tenant become shared mailboxes.

## Bundles

| Bundle id | Deck | Folder | Profile |
| --- | --- | --- | --- |
| `bpw` | `bpw-ai-multiplikatoren` (BPW – AI Champions) | `Demo-BPW` | `cdx` (CDX tenant M365CPI98544940) |

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
8. Deck has tenant casts? Rebuild them: `python tools\build_profile.py <profile>; .\tools\Update-Manifest.ps1 -All -Root profiles\<profile>`.

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

**Echte Tenant-Benutzer (Profile):** Das Deck legt fest, welche echten Benutzer die Demo-Personen spielen (`personas` + `profiles`).
Das Bundle installiert dann `profiles/<profil>/` mit ausgetauschten Namen und Adressen; `-Profile default` holt die Contoso-Namen.
Nach Änderungen: `python tools\build_profile.py <profil>; .\tools\Update-Manifest.ps1 -All -Root profiles\<profil>`.

**Neue Demo:** Ordner anlegen, `manifest.json` (Titel und nächste Schritte auf en/de) anlegen, optional `shared.json`,
`.\tools\Update-Manifest.ps1 -Demo <id>` ausführen, im Deck `demoData.demo` setzen,
`.\tools\Update-Bundle.ps1 -Bundle <bundle> -Deck <deck.json>` ausführen, bei Tenant-Daten `tenant.json` pflegen und pushen.

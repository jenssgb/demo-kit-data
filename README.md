# Demo Kit – Demo data

Public sample files for the demos in the [Demo Kit](https://github.com/jenssgb). There is one folder per demo.
All content is fictional (**Contoso**). The repo holds no customer data and no credentials.

## One command for everything

Open **PowerShell** on the demo VM, signed in as the demo admin (e.g. **MOD Administrator** in a CDX tenant), and run:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Tenant
```

- It installs **all** demos of the kit. `catalog.json` lists every app and its demos. Each demo lands in
  `OneDrive\Demo-Kit\<App>\<demo-id>`, for example `Demo-Kit\Cowork\cowork-copilot`.
- New demos in the catalog are picked up automatically. **Run it once per tenant** (ideally the day before a session,
  because Copilot only finds files once they are indexed) and run it again after kit updates. A rerun adds new and
  changed files and skips unchanged ones (SHA-256 / `quickXorHash`). Nothing is created twice.
- The command never changes. There are no customer bundles.
- Options:
  - `-WhatIf` shows everything without changing the tenant. Try this first.
  - `-Language de` sends the German mails, meetings and OneNote pages (default `en`).
  - `-RemoveLegacy` deletes the folders of older kit versions (`Demo-BPW`, `Demo-DHL`, listed under `legacy` in
    `catalog.json`). It removes them from OneDrive, the local cache and the old Desktop link pages, so Copilot doesn't
    find duplicate files. It works together with `-WhatIf`.
- Without `-Tenant`, the files are only copied to `OneDrive\Demo-Kit` on this PC (fallback: Desktop).
- Works with Windows PowerShell 5.1 and PowerShell 7.
- Troubleshooting options: `-Demo <id>[,<id>]` installs only these demos. Also available: `-Target <folder>`,
  `-Repo`, `-Branch`, `-NoExplorer`.

### What `-Tenant` creates (official APIs)

| What | How |
| --- | --- |
| Files | Uploaded with Microsoft Graph to the signed-in user's OneDrive (`Demo-Kit/<App>/<demo-id>`). Unchanged files are skipped by `quickXorHash`. |
| Mails | Real received mails from the right senders: Graph `sendMail` with `from` (delegated `Mail.Send.Shared` + Exchange **Send As**). |
| Senders | Existing mailboxes are found by display name (Teresa Sac, Billie Vester). People not in the tenant become **shared mailboxes** (Exchange Online PowerShell). A tenant user **without a mailbox** (e.g. Teams license only) is an error, never a look-alike mailbox. |
| Meetings | Graph `POST /me/events` with attendees, on the next matching weekday. |
| OneNote | Notebook, section and pages via Graph (delegated). |
| Not possible by API | Sensitivity labels and DLP, brand kit, skills upload, cost policy, Cowork browser access, Teams meeting with transcript. These are printed as a checklist at the end. |

- **Sign-ins:**
  - One Microsoft Graph sign-in in a browser window, consenting to all listed permissions.
  - Exchange Online, once per tenant. Set-up senders are cached locally and in the admin's OneDrive
    (`DemoKit/senders.json`), and that one sign-in prepares every person of every demo (`people.json`).
  - Needs Exchange admin rights (Global Admin is fine).
- The modules `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` are installed for the current user if
  missing.
- "Send As" on new shared mailboxes can take a while. The script waits up to 20 minutes, otherwise it says "run again
  later".
- **After each run:**
  - `Desktop\Demo-Kit - Links.html` holds, grouped by app: the OneDrive links of all files, OneNote, extra links
    (e.g. Contoso Atlas) and the manual checklist.
  - In a terminal started **as administrator**, the same links also appear as Edge favorites: folder **Demo Kit**,
    with one subfolder per app. The folder is locked; the next run updates it.
- **Log:** every run writes a log to `Desktop\DemoKit-Logs\demo-kit-<time>.log` (last 20 kept). Send that file when
  something goes wrong.
- Each demo describes its tenant data in `<demo-id>/tenant.json` (mails, events, onenote, manual, links). Extra files
  live in `<demo-id>/tenant/`. Neither is part of the file download.

## Demo people = real users of the CDX demo tenant

All mails, prompts, files and meetings use the **real users of the CDX demo tenant M365CPI98544940** directly: Teresa
Sac, Vance DeLeon, Sonia Rees, Billie Vester and Sydney Mattos, with addresses
`<alias>@M365CPI98544940.OnMicrosoft.com`.

- Only use users **with an Exchange mailbox**, i.e. a license with Exchange Online such as Microsoft 365 E5.
  Teams-only users can't send mails.
- There is no mapping layer: what you see in the files is what's in the tenant. `-Tenant` finds the people by display
  name and grants "Send As".
- Only the external Tailspin Toys customer (Jordan Mitchell) becomes a shared mailbox.
- New demo tenant? Rewrite the names once in the files and in `tools/build-*.py`, then run `Update-Manifest.ps1 -All`.

## Apps and demos (`catalog.json`)

`catalog.json` is generated from the app decks of the kit (`decks/*.json` with `files.demoData.app`, pages with
`data`) by `tools\Update-Catalog.ps1`. `Update-Manifest.ps1 -All` runs it automatically when the kit repo sits next to
this one.

| App folder | Demo id | Content | Source / license |
| --- | --- | --- | --- |
| PowerPoint | `powerpoint-copilot` | Fictional Contoso brand guidelines (.docx + .pdf for brand kits), logos, brand background, a branded deck, an off-brand draft deck (Fargo scenario) and a sample PowerPoint skill (`skills/contoso-site-update/SKILL.md`) | Created for the Demo Kit, rebuild with `python tools/build-powerpoint-copilot.py` |
| Word-Excel | `word-excel-copilot` | Edit with Copilot in Word and Excel (Fargo): runbook source pack (.docx), inbound shipment workbook (~400 rows) for the Dock 2 capacity analysis, four Fargo emails (.eml) | Created for the Demo Kit, `python tools/build-word-excel-copilot.py` |
| OneNote | `onenote-copilot` | Contoso Fargo distribution center expansion: OneNote notebook (.onepkg), Word and Excel files | Microsoft Learn course MS-4004, MIT (see `onenote-copilot/LICENSE-MS-4004.txt`) |
| Cowork | `cowork-copilot` | Fargo go-live: four emails as .eml (EN + DE, incl. a Tailspin Toys quote request), readiness tracker (.xlsx), construction weekly week 16 (.docx), Cowork skill (`skills/contoso-golive-readiness/SKILL.md`), SAP placeholder list | Created for the Demo Kit, `python tools/build-cowork-copilot.py` |
| Cowork | `cowork-executive-event` | Eight connected assistant jobs: planning reply, calendar conflicts, conversation preparation, business travel, venue/briefing, event pack, controlled change and follow-up; shared Word inputs, synthetic travel requirements, threaded inbox evidence, fixed-date calendar conflicts, truthful Teams discussions and isolated action workbooks | Created for the Demo Kit, `python tools/build-cowork-executive-event.py`; no Cowork results or recordings supplied |
| Cowork | `cowork-freight-invoice-check` | Freight invoice check: carrier invoices (.pdf), transport orders, carrier master, POD receipts, audit policy | Created for the Demo Kit, `python tools/build-cowork-usecases.py` |
| Cowork | `cowork-carrier-contracts` | Carrier contract comparison: three carrier contracts (.pdf) | `python tools/build-cowork-usecases.py` |
| Cowork | `cowork-lane-margin` | Lane and service margin: raw margin workbook (.xlsx) | `python tools/build-cowork-usecases.py` |
| Cowork | `cowork-customer-qbr` | Customer business review: QBR prep workbook (.xlsx) | `python tools/build-cowork-usecases.py` |
| Cowork | `cowork-disruption-response` | Disruption response: affected shipments (.xlsx), outbound disruption playbook (.docx), trigger message | `python tools/build-cowork-usecases.py` |
| Cowork | `cowork-steering-to-board` | Steering workshop to board pack: workshop transcript (.docx) | `python tools/build-cowork-usecases.py` |
| Copilot-Chat | `researcher-copilot` | Researcher + Vision: Contoso Fargo floor plan (PNG). The cowork-copilot tracker, weekly notes and emails are copied in via `shared.json`. | Created for the Demo Kit, `python tools/build-researcher-copilot.py` |
| Copilot-Chat | `notebooks-copilot` | Copilot Notebooks: Dock 3 go-live sync transcript (.docx/.vtt) and a night-shift handover template. cowork-copilot and brand-guideline files are copied in via `shared.json`. | Created for the Demo Kit, `python tools/build-notebooks-copilot.py` |
| Copilot-Chat | `prompts-copilot` | Contoso Fargo prompt pack (.docx + .md) for Prompt Gallery, scheduled prompts, memory and Pages | Created for the Demo Kit, `python tools/build-prompts-copilot.py` |
| Agents | `agent-builder-copilot` | Agent Builder "Fargo Site Assistant": safety handbook (.docx), dock schedule (.csv/.xlsx for a SharePoint list), paste-ready agent instructions | Created for the Demo Kit, `python tools/build-agent-builder-copilot.py` |
| Agents | `governance-copilot` | Governance & trust: Confidential labor cost workbook, public site fact sheet, Highly Confidential M&A note, label/DLP setup notes (`README_setup.md`) | Created for the Demo Kit, `python tools/build-governance-copilot.py` |

## Add a demo

### Executive-event scenario (LEF26)

Use the same global setup command, with `-Tenant -Language de` for the German presentation.
This scenario additionally needs a deliberately approved **existing demo-only Teams channel**.
It never creates a team, site, public host or external booking.

An authorized operator first identifies the correct demo tenant and channel, then records the actual IDs in
`%LOCALAPPDATA%\DemoKit\executive-event-settings.json`:

```json
{
  "tenantId": "<actual authorized demo tenant GUID>",
  "account": "admin@M365CPI98544940.OnMicrosoft.com",
  "teamId": "<approved existing demo team GUID>",
  "channelId": "<actual 19:...@thread.tacv2 channel ID>"
}
```

These are placeholders describing the **private** file format, not deployable fixture values. Do not commit the
filled file. Missing/invalid settings block only this scenario, visibly; its tenant files/mails/events are not
written. Other demos continue. A valid configuration still requires matching Graph tenant/account, actual channel
and SharePoint access, and approved additional delegated scopes (`Files.ReadWrite.All`, `Team.ReadBasic.All`,
`Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `Mail.ReadWrite`) in the one normal
browser/window sign-in. No device-code login. Do not grant policies, billing or permissions merely to pass a test.

The three BEFORE Word documents are uploaded only to the selected channel's SharePoint files folder.
Other inputs/workbooks live in the presenter's OneDrive. Setup creates four fixed-date Frankfurt calendar entries
and actual Teams threads/replies **as the presenter**; quoted workstream updates are explicitly synthetic.
Two internal setup replies use Graph `createReply` followed by `send`, not forged message headers. They are
sent only to the approved existing demo users. **All live workflow prompts are draft-only.**
Inspect actual conversation membership and mailbox delivery in the tenant before presenting.

The private ownership ledger and exact links live under
`%LOCALAPPDATA%\DemoKit\executive-event\<tenant-id>\`; actual links also reach the global Desktop link page.
An untracked existing shared folder or a changed tracked baseline blocks setup rather than overwriting unknown data.
Keep the ledger when changing PCs; do not delete it to force a rerun.
The ledger's SHA-256 identifies the local upload source; its cTag protects the tracked server version.
SharePoint enriches these Word packages with document metadata, so downloaded ZIP bytes can differ from the
original upload. The live fixture audit verified unchanged Word business parts and non-metadata relationships,
as well as unchanged tracked cTags. Do not treat an enriched package's whole-file hash alone as evidence of a
business-content change, or discard the ownership guard to force an overwrite.

The eight demonstrations retain seven stable exercise codes: CO-26 through CO-29 introduce the smaller jobs,
then CO-23, CO-24 (initial pack and a separate controlled-change stage) and CO-25 form the event chain.
Travel research is a proposal, never an actual booking; EUR 650 for Billie's synthetic trip is inside, not
in addition to, the EUR 9,000 event ceiling. No fare, hotel availability or travel-policy approval is supplied.
Networking preparation uses fictional scenario responsibilities, not scraped biographies or LinkedIn actions.

**Live requests stay short.** Three reusable Cowork skills contain the source boundaries, safety rules,
output contracts and formatting: **contoso-executive-briefing** (the four introductory jobs and CO-23),
**contoso-event-pack** (CO-24 and its change), and **contoso-event-follow-up** (CO-25).
The generator writes bilingual procedures to `cowork-executive-event\skills\<skill>\SKILL.md` and a ZIP
with `SKILL.md` at its root. Import the three trusted ZIPs through **Customize › Add › Upload**
(the verified German interface is **Anpassen › Hinzufügen › Hochladen**), keeping them private.
Start new conversations with the appropriate skill selected. Do not duplicate existing uploads or assume
that a short request works without its skill. The skill derives the run identifier from the selected
original workbook; technical run identifiers are not part of the copy-ready business request.
Skill registration is an explicit one-time manual checklist item, not an undocumented installer API.
Skill names use at most 64 letters, digits, dots, hyphens or underscores, as required by the verified tenant upload dialog.
Earlier long-prompt results do not establish acceptance of this revised skill-based workflow.

Presenter runbooks are `tools\expected\cowork-executive-event-runbook-de.md` and `-en.md`, intentionally outside
the installed/indexed package. `Preparation-de.md`/`-en.md` explain source selection. Start each act from its correct
BEFORE/CHANGE/AFTER evidence. This is a preparation package, not acceptance of the live chain. The presenter
owns the live rehearsal, actual tenant/device checks and stage decision. Use the eight-row checklist in the
runbook to record PASS / FIX / BACKUP ONLY and actual results privately. Local checks do not prove Cowork
execution; delivery does not require further autonomous UI tests or two consecutive complete rehearsals.
Existing genuine outputs, screenshots and native Office renders can be disclosed backups, but repaired files
or files from different runs must not be presented as one successful chain. Missing evidence stays open.
Keep all tenant-bearing results and captures outside both public repositories.

`Reset-CoworkExecutiveEvent.ps1` is installed beside `ExecutiveEvent.ps1`. Run it from the installed package folder
(including `baseline\hashes.json`) or from this repository's `tools` folder, with `-WhatIf` first.
The scenario setup retains the authorized Graph context for reset in the **same PowerShell window**; other
demo setup behavior is unchanged. A newly launched PowerShell process must initialize its Graph context from
the existing token cache first; do not rerun the installer just to reconnect, because it can overwrite numbered
working copies. In a fresh shell, load the private settings and reconnect normally with only reset's scopes:

```powershell
$settings = Get-Content "$env:LOCALAPPDATA\DemoKit\executive-event-settings.json" -Raw | ConvertFrom-Json
Connect-MgGraph -TenantId $settings.tenantId -ContextScope CurrentUser -Scopes 'User.Read','Files.ReadWrite' -NoWelcome
```

This uses the existing browser-authenticated cache when valid. If a new consent or MFA step is requested,
complete the approved normal browser/window sign-in before rehearsal; never use device code.
It hash-checks the untouched `baseline/<language>/Action_Register.xlsx`
and creates a new isolated run and output folder, preserving all earlier results, mails, calendar and Teams posts.
The private `active-<language>.json` report gives the exact new links. Use a new Cowork conversation and the new
run name in draft subjects. Installer reruns are not resets and can overwrite the numbered working copies.
If an original source was changed, stop and review the exact object before restoring; this reset deliberately
does not delete or silently repair objects.

Rebuild and check locally:

```powershell
python tools\build-cowork-executive-event.py
.\tools\Update-Manifest.ps1 -All
python tools\check-cowork-executive-event.py
.\tools\Test-CoworkExecutiveEvent.ps1
```

### Adding another demo

1. Create a new folder `<demo-id>/` with the files.
2. Create `manifest.json` with `title` (en/de) and `next` (en/de). Copy one from an existing demo.
3. Optional: add `shared.json`, a list of `"<other-demo>/<path>"` files this demo also needs. They are copied into
   its folder.
4. Optional: if the demo needs data in the tenant (mails, meetings, OneNote, links, manual steps), add
   `<demo-id>/tenant.json`. Copy `onenote-copilot/tenant.json`.
5. In the kit, add a page to the matching app deck with `"data": "<demo-id>"`. A new app is a new deck with
   `files.demoData.app`.
6. Run `.\tools\Update-Manifest.ps1 -All`. It fills the file lists with SHA-256 hashes and regenerates `people.json`
   and `catalog.json`. Run it again after every `tools/build-*.py` run.
7. Commit and push. The presenter reruns the same `-Tenant` command.

The repo uses `.gitattributes` `* -text`, so files are stored byte for byte and the hashes match the raw downloads.

---

# Deutsch

Öffentliche Beispieldateien für die Demos im Demo Kit, ein Ordner pro Demo. Alle Inhalte sind fiktiv (**Contoso**).

**Ein Befehl für alles:** Auf der Demo-VM als Admin anmelden (z. B. MOD Administrator im CDX-Tenant), PowerShell
öffnen, den Befehl oben einfügen und **Enter** drücken.

- Das Skript installiert **alle** Demos nach `OneDrive\Demo-Kit\<App>\<Demo-ID>`. Es lädt die Dateien per Graph in den
  OneDrive und legt die Mails mit den echten Absendern an (fehlende Personen als freigegebene Postfächer mit „Senden
  als“), dazu Termine und OneNote-Notizbücher.
- Am Ende listet es, was noch von Hand zu tun ist (Vertraulichkeitsbezeichnungen/DLP, Brand Kit, Skills,
  Cowork-Browserzugriff).
- Neue Demos kommen automatisch dazu (`catalog.json`). Einmal pro Tenant ausführen, am besten am Vortag (Indexierung),
  und nach Updates erneut ausführen. Unveränderte Dateien werden übersprungen.
- `-WhatIf` zum Ausprobieren, `-Language de` für deutsche Mails.
- `-RemoveLegacy` löscht einmalig die alten Ordner `Demo-BPW` / `Demo-DHL`, damit Copilot keine doppelten Dateien
  findet.

**Demo-Personen = echte Benutzer des CDX-Tenants** (Teresa Sac, Vance DeLeon, Sonia Rees, Billie Vester, Sydney
Mattos). Sie stehen direkt in allen Mails, Prompts und Dateien, ohne Zuordnung. Nur der externe Kunde (Jordan
Mitchell, Tailspin Toys) wird ein freigegebenes Postfach.

**Nach jedem Lauf:**
- `Desktop\Demo-Kit - Links.html` (nach App gruppiert).
- Im Terminal **als Administrator** zusätzlich Edge-Favoriten: Ordner **Demo Kit** mit einem Unterordner pro App.
- Log unter `Desktop\DemoKit-Logs`. Bei Fehlern diese Datei schicken.
- Die Exchange-Anmeldung (zweites Login) ist nur einmal pro Tenant nötig.

**Neue Demo:**
1. Ordner und `manifest.json` anlegen, optional `shared.json` / `tenant.json`.
2. Im Kit eine Seite mit `"data": "<Demo-ID>"` im passenden App-Deck anlegen.
3. `.\tools\Update-Manifest.ps1 -All` ausführen und pushen.
4. Auf der VM denselben Befehl erneut ausführen.

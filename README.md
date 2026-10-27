# Demo Kit – Demo data

Public sample files for the demos in the [Demo Kit](https://github.com/jenssgb). There is one folder per demo.
All content is fictional (**Contoso**). The repo holds no customer data and no credentials.

## One command for everything

Open **PowerShell** on the demo VM and sign in to Graph as the **demo administrator**. The script lists enabled
internal tenant users with Exchange mailboxes; **choose the demo user by number**. Empty input or `Q` cancels
before tenant writes. An Exchange administrator authenticates separately only when sender grants need setup:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/jenssgb/demo-kit-data/main/install.ps1))) -Tenant
```

- It installs **all** eligible demos of the kit. `catalog.json` lists every app and its demos. When choosing oneself,
  files land in `OneDrive\Demo-Kit\<App>\<demo-id>`, for example `Demo-Kit\Cowork\cowork-copilot`.
  For another selected user, ordinary files remain in the admin's OneDrive under
  `DemoKit-Presenters/<user-id>/Demo-Kit/<App>/<demo-id>`. Only that user's inner **Demo-Kit** folder is shared
  with **editing access**, with a native OneDrive invitation on the first grant. The existing general
  `Demo-Kit` folder, other users' trees and private delivery state are never shared as the start folder.
  Shared **Work IQ / WIQ26** evidence stays in its existing SharePoint location (see below).
  The selected user receives demo mail, meeting invitations and a **Demo Kit - links and preparation** email.
  Open the **Demo-Kit** start link as the demo user and select **Add shortcut to My files** once. It then appears
  under **My files › Demo-Kit**. Sign the OneDrive sync client in as that user to sync it to File Explorer.
  In Copilot, select the cloud original through **+ › Attach cloud files**, rather than uploading another copy.
  A shortcut does not automatically select sources in a prompt or establish Copilot indexing.
  Shortcut/sync availability still depends on the tenant's OneDrive policies.
- New demos in the catalog are picked up automatically. **Run it once per presenter account** (ideally the day before a session,
  because Copilot only finds files once they are indexed) and run it again after kit updates. A rerun adds new and
  changed files and skips unchanged ones (SHA-256 / `quickXorHash`). Use the **same publishing admin** for reruns.
  Choosing the signed-in user retains personal setup. Shared WIQ26 does not require a presenter install;
  use the same publishing admin for its reruns.
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
| Files | Self setup retains `Demo-Kit/<App>/<demo-id>`. Central setup uploads only eligible ordinary demos under `DemoKit-Presenters/<user-id>/Demo-Kit/<App>/<demo-id>` and shares that one inner root with the selected authenticated user. Confirmed write permission is required before publishing its links; unchanged files are skipped by `quickXorHash`. No anonymous or organization-wide links. |
| Mails | Submitted to the chosen user's mailbox with the right `from` (delegated `Mail.Send.Shared` + publisher's Exchange **Send As**). Submission is not proof of delivery. No access to the target inbox is requested or assumed. |
| Senders | Existing mailboxes are found by display name (Teresa Sac, Billie Vester). People not in the tenant become **shared mailboxes** (Exchange Online PowerShell). A tenant user **without a mailbox** (e.g. Teams license only) is an error, never a look-alike mailbox. |
| Meetings | Publisher organizes the meeting and invites the selected user along with the existing cast. Reruns add a missing selected attendee without removing previous attendees. `onlineMeeting: true` requests a Teams meeting. Private fixed-date scenarios are excluded from central sharing. Invitation processing still needs checking. |
| OneNote | Notebook, section and pages are created under the publisher via delegated Graph. An exact OneNote drive package is shared for editing; ambiguous or not-yet-discoverable notebooks explicitly fail sharing and need a rerun. The presenter opens the shared notebook once in their app. |
| Rehearsal / approved tenant setup | Transcript-enabled meeting rehearsal, attaching the video's original transcript, OneNote desktop opening, brand kit, skills, policies and approved SharePoint/Planner destinations. These are printed as a checklist; the installer does not fabricate meeting history, change billing, or create arbitrary sites/groups. |

- **Sign-ins:**
  - One Microsoft Graph sign-in in a browser window, consenting to all listed permissions.
  - Exchange Online, once per **sending Graph account**. Choose an Exchange administrator of the same tenant;
    Send As targets that Graph account, not the selected recipient. The recipient needs no administrator role.
  - Grants are cached per tenant and sending account, locally as `DemoKit\senders-v3-<tenant-id>-<presenter-hash>.json`
    and in that publisher's OneDrive (`DemoKit/senders-v3.json`). The envelope must match both identities.
    Changing the selected recipient does not invalidate the publisher's sender grants. The sign-in prepares `people.json`.
  - Account-locked and private executive scenarios are **BLOCKED** before their writes when preparing another user.
    Choosing oneself retains their existing account and ownership checks. WIQ26 shared publishing is unchanged.
- Central mail submissions are recorded per tenant, publisher, selected recipient and language in the publisher's
  private `DemoKit-Presenter-State-<tenant>-<recipient>-<language>` OneDrive folder, outside the shared demo tree.
  A cloud `installer-lock` prevents simultaneous sends for the same recipient. A pending/unknown submission is
  **never automatically resent**: check the publisher's Sent Items or Exchange trace, then repair only the
  affected ledger entry. After a crashed run, remove only `installer-lock` after confirming no installer is active.
  Explicitly rejected requests can be retried. Recorded submissions do not establish delivery.
  Offline regressions: `.\tools\Test-CentralSetup.ps1` and `.\tools\Test-PresenterFolders.ps1`
  (Windows PowerShell 5.1 and PowerShell 7; no sign-in).
- Existing installations are not moved or deleted and earlier individual demo grants are not revoked.
  The first central run with the new layout uploads a separate recipient-scoped set; later runs reuse it.
  Use the new start link and select sources from this folder, not older per-demo links.
  Previously shared copies can remain discoverable to Copilot; this change does not claim to remove duplicates.
  Private/account-locked demos are filtered before choosing upload paths and never copied into this tree.
  The user's one-time shortcut/sync step is not impersonated through Graph and requires no new admin sign-in.
  References: [Graph sharing](https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0),
  [folder creation](https://learn.microsoft.com/graph/api/driveitem-post-children?view=graph-rest-1.0),
  [OneDrive shared shortcuts](https://support.microsoft.com/en-us/onedrive/add-shortcuts-to-shared-folders-in-onedrive).
- Required PowerShell modules are checked before sign-in and installed automatically for the current user.
  The installer uses tested versions of `Microsoft.Graph.Authentication`, `ExchangeOnlineManagement` and its
  `PowerShellGet` / `PackageManagement` dependencies. Existing matching versions are reused.
  If `Install-Module` fails (including Gallery download 404s), it downloads the same package from the official
  `cdn.powershellgallery.com` endpoint, verifies its pinned SHA-256 checksum and manifest, and installs it under
  `%LOCALAPPDATA%\DemoKit\Modules`. Downloads retry up to three times; an unverifiable package stops setup.
  This also works during `-WhatIf`: local prerequisites are installed, but tenant changes remain preview-only.
  No separate manual package installation or extra sign-in is needed. The fixed command above is unchanged.
  When GitHub revision lookup is available, all files in a run use that immutable commit, avoiding mixed
  old/new scripts from branch-URL caches immediately after an update.
- "Send As" on new shared mailboxes can take a while. The script waits up to 20 minutes, otherwise it says "run again
  later".
- **After each run:**
  - For central setup the chosen demo user receives one shared **Demo-Kit** start link, notebook links and the
    one-time **Add shortcut to My files** instruction by mail. The same verified start link is in the
    publisher's Desktop page and Edge favorites; failed sharing does not publish inaccessible file links.
    Mail, invitation processing, Copilot licensing, indexing and live retrieval must still be rehearsed as that user.
  - `Desktop\Demo-Kit - Links.html` holds, grouped by app: the OneDrive links of all files, OneNote, extra links
    (e.g. Contoso Atlas) and the manual checklist.
  - In a terminal started **as administrator**, the same links also appear as Edge favorites: folder **Demo Kit**,
    with one subfolder per app. The folder is locked; the next run updates it.
- **Log:** every run writes a log to `Desktop\DemoKit-Logs\demo-kit-<time>.log` (last 20 kept). Send that file when
  something goes wrong.
- Each demo describes its tenant data in `<demo-id>/tenant.json` (mails, events, onenote, manual, links). Extra files
  live in `<demo-id>/tenant/`. Neither is part of the file download.

### First central run: progress and completion

The first run for a selected user is **not just a permission change**. It uploads a complete new set of eligible
demo files into that user's isolated tree, one file at a time. This also happens when switching from the older
individual-folder sharing layout. Hundreds of file checks/uploads can make this run lengthy; do not promise
a fixed duration. In a real run without `-WhatIf`, green `new` rows identify uploaded files. New rows mean progress, but an old
screenshot alone cannot establish that the current process is still advancing.

The shared root is published after ordinary file uploads and permission verification. Notebook/calendar/mail
setup continues afterward; the links email is submitted near the end. The final area table, Desktop/Edge/log
output and a returned PowerShell prompt identify completion. A later run with the **same publishing admin and
selected user** reuses the tree and skips unchanged uploads and recorded mail submissions, but still checks
files and other setup state.

As the selected user, the correct start folder has the breadcrumb
**publisher › DemoKit-Presenters › user ID › Demo-Kit**, with app folders such as **PowerPoint**, **Word-Excel**
and **Copilot-Chat** underneath. Do not try to open the publisher's general **Demo-Kit** folder or add every app
folder separately. Add a shortcut to the **whole inner Demo-Kit folder** once using **Add shortcut to My files**.
Microsoft's [shared-folder instructions](https://support.microsoft.com/en-us/onedrive/add-shortcuts-to-shared-folders-in-onedrive)
describe selecting the shared folder and choosing that command. Opening the folder successfully is not evidence
that the shortcut has been added or that the sync client is configured.

**Observed check, 2026-10-27:** the submitted console output reported 494 new OneDrive files, one root sharing
grant, two unchanged notebook grants, 41 mail submissions and one links-message submission, with no failures in
those areas. Two `AccountLocked` entries were deliberately excluded private scenarios, not failed ordinary
uploads. A screenshot from the selected recipient showed the new shared root and app folders opening successfully.
These are observations from one run, not fixed totals for the growing catalog. Shortcut creation, desktop sync,
mail delivery and Copilot indexing/retrieval were **not** confirmed by this evidence.

That run also printed `Failed to clear the persisted MSAL token cache during Disconnect-MgGraph` after the
summary and then returned to the prompt. Record this as an authentication-cleanup warning separately from the
area results; do not claim the persisted cache was cleared. It is not, on its own, a reason to restart installation
or resubmit mail. Inspect any actual `Failed`, `Later` or `BLOCKED` entries independently.

### Shipping rehearsal corpus

The 27 shipping/DG training folders map to 23 workflow groups in the kit's existing app collections,
not to a customer bundle. `tools\build-shipping-playbook.py` creates original synthetic Contoso
DOCX/XLSX/PDF/CSV sources, mailbox seeds, scheduled events and administrative-only DG skill instructions.
`word-shipping-brief` owns the canonical files; `shared.json` makes each other page self-contained.
English source files are intentional (Excel workbook rules are fully supported in English); the
kit's instructions, prompts, setup and manual checklists are bilingual.
`QBR_authored_reference.pptx` is an explicitly authored six-slide fallback with an editable chart
and speaker notes, not a generated Copilot result or an approved Brand Kit. It is shared by both
PowerPoint QBR workflows; the live demo still saves and reuses its own reviewed Copilot output.

Run `python tools\build-shipping-playbook.py`, then `.\tools\Update-Manifest.ps1 -All`.
Validate offline with `python tools\test_shipping.py` and `.\tools\Test-DemoMail.ps1`.
The data checks cover 216 volume rows, exactly three D&D variances above 5%, five invoice/accrual
exceptions, the 300-row CSV join/ranking, Office structure and skill package bytes.

The same global `install.ps1 -Tenant` command installs this content. It does **not** publish SharePoint
sites, create Teams threads/transcripts, install Copilot Chat skills, build agents, publish Brand Kits
or generate Copilot results. The folder's `tenant.json` manual checklist and the kit page explicitly
identify those rehearsals and availability gates. A script, storyboard or expected-result CSV is
never represented as a real transcript, model output or cloud edit history. Actual generated outputs
must be saved during rehearsal; don't publish private tenant links in this public repository.

Optional `tenant.json` `mailOptions` entries are matched to an exact (optionally localized) mail subject:

```json
{
  "mailOptions": [{
    "subject": "Contoso Star delay - carrier attachments",
    "attachments": ["Revised_ETA_ContosoStar.pdf", "Affected_bookings.xlsx"],
    "html": true
  }]
}
```

Each attachment must be listed in that demo's manifest (own or shared file). The combined raw size
limit is 2 MB; larger evidence belongs in cloud links. `html: true` safely converts the plain EML body
to HTML, preserving quoted context without claiming that independent emails form a genuine thread.
Unconfigured mails retain their existing text payload. Existing subject/sender matches are skipped,
not modified; use a genuinely new synthetic subject when changing an already-installed seed.
Graph `sendMail` acceptance is not proof of delivery or indexing; verify both before demonstrating
attachment search or Work IQ. `-WhatIf` validates new-mail payloads without sending them.

### Work IQ: publish once, demonstrate with other users

The same global `-Tenant` command now includes `workiq-port-operations` (**WIQ26**). Run this setup as the
demo administrator, ideally ahead of the session; adding `-WhatIf` first previews without tenant writes.
It publishes **24 distinct DOCX documents in EN/DE (48 files)** into the **existing tenant root SharePoint
document library**, folder `Demo-Kit-Work-IQ-WIQ26`. It grants folder-only **read** access to all enabled,
internal `Member` users with an enabled Exchange service plan and a mail address. Guests, disabled accounts
and Teams-only users are excluded; existing broader inherited permissions are not removed.

Four real evidence messages are submitted individually to each eligible user's **primary inbox**, with the
shared folder link, and one real rehearsal invitation includes the audience. Other users **do not rerun the
installer**, need no admin role and do not rely on the admin's private files/chat history. Each still needs
the appropriate Copilot license, Work IQ availability and tenant policy access. The admin needs existing
write access to the root library and consent to delegated `Sites.Read.All`, `Files.ReadWrite.All` and
`User.Read.All` in the existing bundled Graph sign-in. Missing access is explicitly **BLOCKED**; no site,
admin role, search-policy change, custom vector index or API billing is created.

Canonical documents are not duplicated into the admin's regular demo OneDrive folder. The Desktop link page
and Edge favorites contain the actual shared folder URL. Mails use the selected setup language; both document
languages are shared. `tools\build-workiq.py` reuses the established port builder's Office styling/packages;
run `tools\Update-Manifest.ps1 -All` afterward. Test gold under `tenant\expected.json` is not uploaded/indexed.

Rerun **as the same admin** to update documents and include newly eligible users. The private admin OneDrive
folder `DemoKit-WIQ26-State-<tenant-id>` keeps publisher/tenant-checked submission state across PCs. A
non-expiring `installer-lock` prevents concurrent runs; only remove that specific lock after a crashed run
and after confirming no installer remains active. Pending/uncertain mail sends are **never automatically
resent**: inspect Sent Items/Exchange message trace before repairing the matching private state entry.
Do not delete the ledger to retry blindly. This prevents a timeout from silently creating duplicate mail.
A future calendar invitation is preserved. If it has become past, rerunning as the same admin moves that
same event to the next Thursday at 10:00 local installer time, without creating another event. New eligible
attendees are added without removing manually added ones.

Uploading is **not** an indexing proof, HTTP 202 is **not** delivery confirmation, and a calendar invitation
is **not** a transcript. As the actual presenter, open `[WIQ26-M02]`, its shared folder and
`en\WIQ-D03_status_v2.docx` (or `de`), verify all four mails and the real invitation, and then rehearse the
kit's **WI-01** prompts without attachments. Require both real mail and document citations and all four
holds to remain unresolved. Microsoft manages semantic indexing; prepare ahead and verify actual discovery,
not an assumed timing SLA. Direct Work IQ API/CLI/MCP consent and billing are a separate, unconfigured path.

Offline integration checks: `tools\Test-WorkIqShared.ps1`; corpus checks: `python tools\check-workiq.py`.

### Other port-operations app rehearsals

The per-app `*-port-operations` folders share one original synthetic storyline: 12 shipments, 115 containers,
four held shipments / 43 containers, and an intentionally important distinction between document receipt and
release approval. The Word/Excel folder owns the source files; `shared.json` makes the other app folders
self-contained. `tools\build-port-operations.py` rebuilds the DOCX/XLSX/CSV, localized EML, OneNote HTML and an
original English-narrated MP4 with timed WebVTT. It uses the existing `python-docx`, `openpyxl`, `Pillow` and
`imageio-ffmpeg` packages plus an installed English Windows speech voice. Run `tools\Update-Manifest.ps1 -All`
after rebuilding. No Microsoft screenshots, real shipment identifiers or customer data are included.

For a read-only account/connectivity check, use `tools\Test-DemoAccount.ps1 -ExpectedAccount <presenter-UPN>`.
Add `-Connect` only when ready for the normal browser sign-in; `-Report` can write a **private**, uncommitted JSON
report. It never assigns licenses/roles or enables billing. A missing context fails explicitly. License evidence,
app access and a completed live Copilot rehearsal are different checks: a locally generated package or a successful
Graph read is **not** proof that the demo account is ready.

Local checks: `tools\Test-PresenterSetup.ps1`, `tools\Test-Cowork-Mailbox.ps1` and
`python tools\check-port-operations.py`. The kit holds the bilingual click paths, prompts, limits and official sources.

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
| Cowork | `cowork-assistant-day` | Eight independent everyday cases CO-30–CO-37: inbox, morning/meeting briefs, commitments, calendar, meeting-to-action, internal visit and day closure; 18 separate emails, nine own-calendar appointments, bilingual source documents, preserved run workbooks and one reusable skill | Original synthetic Contoso inputs, `python tools/build-cowork-assistant-day.py`; live acceptance remains with the presenter |
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

### Assistant-day scenarios (EAD26)

In Demo Kit, open **Cowork end to end › Executive assistant: A day with Cowork**
(DE: **Cowork Ende zu Ende › Assistenzalltag: Ein Tag mit Cowork**).
This adds eight independent cases alongside the existing eight leadership-event moments; it does not replace that story.
Use the same global `-Tenant` command above, with `-Language de` for German, and preview with `-WhatIf`.
The agent prepares and validates this package locally; the presenter installs and performs all live tests.

Files land in `Demo-Kit\Cowork\cowork-assistant-day`. Mails appear in the approved admin's Inbox
under `[EAD26 nn-nn]`; these are separate emails, not claimed conversation threads.
Nine `[EAD26 key]` appointments appear in the same admin's **own** calendar on **19–23 October 2026**.
There is no delegated executive mailbox, private-chat dependency or invitation. Scenario timestamps in sources,
not actual delivery time or the current clock, determine the business chronology.
`presenterAccount` checks cached Graph identity and mailbox before writes; `requiredUsers` verifies real Exchange mailboxes.
Ordinary `weekday` configurations keep their rolling behavior.

Fixed-date entries use an owned Graph single-value extended property and deterministic transaction ID.
Repeat setup finds that identity even if an appointment was moved or renamed. Multiple or changed matches are
reported, never overwritten. The fixed-date path rejects attendees and ambiguous daylight-saving times.
There are no new scopes beyond the ordinary installer requirements.

Read `Preparation-en.md` or `Preparation-de.md`. Register **contoso-assistant-day.zip** once through
**Customize › Skills › arrow next to Add › Upload skill** and keep it private. A duplicate name gets numbered,
not replaced; inspect an existing skill first. For each new conversation select that skill, the active run's
`Run_Context.md`, the case's `Case.md` and its exact listed cloud originals. CO-35 edits the same active-run
`Action_Register.xlsx`, not a baseline or replacement copy. Do not send/post/invite/move/book/pay or activate automation.

Three initial runs per language are supplied. Installer reruns preserve their existing workbooks and contexts
(`preserveExistingFiles`), rather than resetting edited originals. A fresh run is created with
`Reset-CoworkAssistantDay.ps1 -Language en -WhatIf` (or `de`), then without `-WhatIf` after review.
The script reuses the authorized cached Graph session, checks baseline workbook **and context** SHA-256,
and creates a unique run/output folder without deleting or editing older results. It does not alter mails or calendars.
The new cases do not require the event story's private Teams settings.

Presenter-only selection, output contracts and eight **OPEN** live checks:
`tools\expected\cowork-assistant-day-runbook-en.md` and `-de.md`. CO-37 includes a clearly labeled synthetic
checkpoint for independent starts, not evidence that earlier Cowork tasks succeeded.
No new Cowork outputs, screenshots or fabricated backups are included. Save genuine accepted results privately
after your rehearsal. Local regression commands:

```powershell
python tools\check-cowork-assistant-day.py
.\tools\Test-CoworkAssistantDay.ps1
python tools\check-cowork-executive-event.py
.\tools\Test-CoworkExecutiveEvent.ps1
.\tools\Test-ExecutiveEventState.ps1
```

### Executive-event scenario (LEF26)

Use the same global setup command, with `-Tenant -Language de` for the German presentation.
This scenario additionally needs a deliberately approved **existing demo-only Teams channel**.
It never creates a team, site, public host or external booking.

Once per tenant, an authorized operator identifies the correct demo tenant and channel, then records the actual IDs in
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
filled file. The global installer backs up approved settings and each ownership-ledger update to the same demo
admin's private OneDrive at `DemoKit/executive-event-state-<language>-v1.json`. On another PC or VM it restores
missing local state automatically after the **same** Graph sign-in; no ZIP, extra setup command or new sign-in
is required. `-WhatIf` reads and validates this state without restoring files or updating the cache. Run only
one tenant installer at a time. Conflicting local/cloud state is never silently overwritten.
The approved destination can be recovered from either language's cache; ownership ledgers are language-specific
and are never borrowed from a different language.

Missing/invalid state in both locations blocks only this scenario, visibly; other demos continue.
A valid configuration still requires matching Graph tenant/account, actual channel
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
The installer recovers this ledger together with the destination on a new PC; do not delete it to force a rerun.
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
- Die Exchange-Anmeldung (mit einem Administrator desselben Tenants) ist einmal pro Präsentatorkonto nötig.

**Neue Demo:**
1. Ordner und `manifest.json` anlegen, optional `shared.json` / `tenant.json`.
2. Im Kit eine Seite mit `"data": "<Demo-ID>"` im passenden App-Deck anlegen.
3. `.\tools\Update-Manifest.ps1 -All` ausführen und pushen.
4. Auf der VM denselben Befehl erneut ausführen.

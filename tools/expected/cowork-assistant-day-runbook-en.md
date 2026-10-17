# CSMAD presenter selection

# Assistant day: start here

Demo Kit: Cowork end to end > Executive assistant: A day with Cowork. The separate leadership-event page remains available.

Use the standard global tenant setup, first with -WhatIf. It installs every demo; do not run a per-case installer. Confirm the signed-in approved demo admin and current Exchange licenses. Actual mail appears in this admin's Inbox with [EAD26 nn-nn]; appointments appear in this admin's own calendar on 19-23 October 2026, under [EAD26 key]. Files: Demo-Kit > Cowork > cowork-assistant-day.

One-time skill setup: download skills/contoso-assistant-day.zip. In Cowork select Customize > Skills > arrow next to Add > Upload skill. Choose the ZIP and keep Only you. If that name already exists, inspect it instead of creating a numbered duplicate. Start a new conversation and select the skill in Sources; verify it is active in the side panel. UI translations/rollout require your tenant check.

For each case, attach runs/en/01/Run_Context.md, inputs/en/<code>/Case.md and exactly the files named there through Add attachments > Attach cloud files. For CO-35 also select that run's original Action_Register.xlsx. Installing data does not select it in Cowork. Check the listed inbox messages/calendar before starting. Use the short request in Demo Kit; inspect persistent files and Outlook Drafts outside chat.

Runs 01-03 are independent clean starting slots. Use a new slot after any edits. If used up, the local Reset-CoworkAssistantDay.ps1 -Language en -WhatIf previews a non-destructive fresh-run reset; remove -WhatIf only after review. It uses the existing cached authorized Graph context, never starts sign-in, downloads and SHA-256 checks the original baseline, and creates a unique run. No calendar or mail reset is needed because live tasks may not change them. If they did, stop; do not silently reinstall.

The installer updates baselines/fixtures and is NOT a reset; existing run workbooks and contexts are preserved even when edited. New reset runs are not manifest-managed and earlier results are retained. For a reset, the ZIP-uploaded skill is unchanged; re-register only an intentional skill revision, avoiding duplicates.

All eight live checks remain OPEN. No new Cowork backups are supplied. Prerequisites: active Copilot license, Cowork enabled, usage billing, reachable cloud files, inbox discovery and your own calendar. No delegation, private executive Teams chats, browser-control add-on or Frontier App skill is required. Stop on missing sources or send/post/invite/move/automation approvals. Presenter expectations are outside this folder in tools/expected.

## Select, do not race through sixteen jobs

Pool: eight new independent everyday cases plus the unchanged eight event moments. For the demo portion choose CO-30, CO-31, CO-34, CO-35 and CO-36 as the core, with CO-24 presentation/change as a contrasting saved artifact and CO-37 as the closing thought. CO-32 and CO-33 are reserves, not mandatory extra steps. Start longer jobs before explaining another saved result. Suggested 40-minute speaking allocation (not Cowork runtime): 4 minutes framing, 6 inbox, 5 morning, 6 calendar, 8 meeting-to-action, 6 visit or event-pack contrast, 5 closure/judgment. Do not fill the slot with generation waits. Use only your own genuinely saved and verified backups; if no accepted backup exists, omit that live case instead of presenting a fabricated result.

Existing event pool, unchanged: CO-26 reply; CO-27 calendar; CO-28 conversations; CO-29 travel; CO-23 executive briefing/venue comparison; CO-24 event pack; CO-24 speaker change; CO-25 follow-up. These remain a separate story with their separate three skills and private historical evidence.

## CO-30 — Inbox triage

The executive office has six unrelated messages. Two concern the same request; one needs no response. Prepare decisions without making them.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/en/CO-30/Case.md`, `inputs/en/CO-30/Inbox_Rules.docx`, `runs/en/<active>/Run_Context.md`.

Please sort these inbox requests for me and prepare the two replies that need my attention. Save the overview so I can review it with Teresa.

Read only the six CO-30 inbox messages listed in Case.md. Save a compact table with all six IDs, five categories (decision required, meeting request, information, delegable, no action), reason, proposed next step and source. Flag 30-05 as the duplicate of 30-01, not a second approval. Save exactly two actual unsent Outlook drafts: Sydney (30-01) asking for the missing purchase justification without approving EUR 2,200; Teresa (30-02) asking which existing commitment takes priority, without promising a slot. Do not file, flag, delete or move mail.

Human judgment: Choose the escalation and tone personally.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## CO-31 — Morning brief

Before the day starts, a supplier decision has moved forward, IT has an update and the office calendar contains a clash.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/en/CO-31/Case.md`, `inputs/en/CO-31/Morning_Papers.docx`, `inputs/en/CO-31/Calendar_Snapshot.docx`, `runs/en/<active>/Run_Context.md`.

Please prepare my morning brief for the scenario day. Bring together the calendar, Teresa's update and the current papers, with the decisions I should tackle first.

Use only CO-31 mail, Morning_Papers and the listed EAD26 own-calendar entries. Treat the supplied snapshot as 19 October 2026 08:30 Europe/Berlin, irrespective of the actual clock. Verify the actual calendar against the selected snapshot; disclose discrepancies. Save a readable one-page Word briefing with top three priorities, changed facts, decision deadlines, source links, missing evidence and proposed work order. Replace EUR 72,000 / five days with EUR 88,000 / eight days from the dated current paper. Teresa's decision deadline is 09:30 today, not tomorrow. Billie's test is reported complete in synthetic evidence, not verified by you. Exclude CO-35 afternoon notes and CO-37 day-close sources.

Human judgment: Decide what to raise with Teresa first.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## CO-32 — Executive meeting brief

A supplier service review needs a decision. The old dashboard looks positive; a newer quality note changes the picture.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/en/CO-32/Case.md`, `inputs/en/CO-32/Review_Previous.docx`, `inputs/en/CO-32/Review_Current.docx`, `runs/en/<active>/Run_Context.md`.

Please get my supplier review briefing ready. Compare the old dashboard with the latest quality update and give me the questions Teresa needs to decide.

Use the two CO-32 messages and the two selected review papers. Save a one-page Word briefing with meeting goal, earlier/current comparison, decisions, three talking points, unknowns and real source links. Keep the periods separate: September 96% on-time and six issues; October-to-date 88% and fourteen issues. Sonia's training capacity remains unconfirmed; do not designate her as approved owner. The supplier review is 20 October at 11:00, not the CO-35 internal process review. No access to anyone else's Teams chats is assumed.

Human judgment: Choose the negotiating position, not just the numbers.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## CO-33 — Commitments radar

Three promises have different evidence: one is closed, one has no completion evidence, and one response is ambiguous.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/en/CO-33/Case.md`, `inputs/en/CO-33/Handover_List.docx`, `runs/en/<active>/Run_Context.md`.

Please check what became of these three commitments and prepare the follow-ups. Keep uncertain cases separate so I don't chase someone who has already delivered.

Read the five CO-33 messages chronologically using their explicit synthetic timestamps, not delivery time. Save a three-row evidence register: Vance's handover list confirmed delivered by 33-02 and selected Handover_List; Sonia's listening summary has no completion evidence; Billie's 'looks fine' is not explicit accessibility-test completion. Missing evidence is not proof of failure. Save two unsent drafts: a gentle request to Sonia for the summary/status and a clarification to Billie about the test evidence. Do not remind Vance, mark promises complete or invent deadlines.

Human judgment: Choose whom to call rather than send a reminder.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## CO-34 — Calendar week

The office week has an overlap, a building change without travel time and a workshop without an agenda. Other people's availability is unknown.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/en/CO-34/Case.md`, `inputs/en/CO-34/Calendar_Snapshot.docx`, `inputs/en/CO-34/Calendar_Rules.docx`, `runs/en/<active>/Run_Context.md`.

Please review my scenario week and prepare sensible options for the calendar problems. Draft the coordination, but don't change any appointments.

Inspect only the EAD26 events for 19-23 October in the presenter's own calendar; compare with the selected snapshot. Save a decision paper separating hard overlap, missing travel/preparation buffer and missing agenda. Monday site review 10:00-10:45 overlaps budget 10:30-11:00. Building B visit at 11:00 follows Building A budget without the required 30-minute journey. Monday workshop 13:00 has no agenda. Tuesday supplier review 11:00 follows a protected focus block ending 11:00 with no 15-minute preparation. Offer alternatives as proposals only; do not assert other attendees are available. Save two unsent drafts: Sydney about the budget conflict and Sonia requesting the workshop agenda. Keep calendar IDs, times and attendees unchanged.

Human judgment: Protect the commitments that matter politically.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## CO-35 — Meeting to action

A separate internal process review has released notes, one existing workbook row, and new actions with one missing deadline.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T14:00:00+02:00

Sources: `inputs/en/CO-35/Case.md`, `inputs/en/CO-35/Released_Notes.docx`, `runs/en/<active>/Action_Register.xlsx`, `runs/en/<active>/Run_Context.md`.

Please turn these released review notes into decisions and next steps. Update our action register and prepare the individual follow-ups.

Use only the selected released CO-35 notes and active-run original workbook. Save Decisions.docx and edit that exact cloud workbook in place. Preserve sheets, column order, original P01 cell values/types/formats, Readme and table name. Append A01 Sydney Mattos, reconcile handover exceptions, 2026-10-22; A02 Billie Vester, accessibility test evidence, deadline not decided (blank); A03 Sonia Rees, onboarding checklist draft, 2026-10-23. Expand Actions table to A1:F5. D01 is a two-week pilot, D02 no software purchase; neither is approval to send. Save three individualized unsent drafts to the named owners. Billie's draft asks for a due date, not an invented commitment. Do not overwrite the baseline or make a replacement workbook.

Human judgment: Confirm the missing due date and have the difficult conversation.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## CO-36 — Short-notice executive visit

Teresa plans an internal site visit on Thursday, with a tight arrival/departure window, a safety constraint and uncertain technical availability.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T14:30:00+02:00

Sources: `inputs/en/CO-36/Case.md`, `inputs/en/CO-36/Visit_Brief.docx`, `runs/en/<active>/Run_Context.md`.

Please get Teresa's Thursday visit ready using these messages and the site notes. Save a workable visit pack and draft the coordination for my review.

Use only CO-36 evidence. Produce three matching saved documents: one-page agenda for Thursday 22 October 10:00-12:00; one-page participant briefing grounded in synthetic responsibilities; decision list separating approved constraints from open choices. Include Sonia's mandatory safety introduction before the 30-minute walk, 10-minute transitions, and at least 20 minutes of personal discussion. Teresa leaves at 12:00; Billie is available only from 11:15 and his technical demo is not confirmed. Never invent attendee biographies, a room booking or acceptance. Save three unsent drafts: Teresa agenda proposal, Sonia safety/logistics questions, Billie feasibility of the 11:15 demo. No invitations or calendar writes.

Human judgment: Decide the visit's real purpose and speak to the hosts.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## CO-37 — Day closure

A deliberately synthetic end-of-day checkpoint records one completed decision, an unsent draft and an unresolved test. It is not a record of earlier Cowork runs.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T17:00:00+02:00

Sources: `inputs/en/CO-37/Case.md`, `inputs/en/CO-37/Day_Checkpoint.docx`, `inputs/en/CO-37/Tomorrow_Calendar.docx`, `runs/en/<active>/Run_Context.md`.

Please prepare my day closure from this checkpoint and get tomorrow's morning brief started. Separate completed work, open approvals and anything the evidence doesn't settle.

Standalone route: use only CO-37 synthetic checkpoint and tomorrow calendar, as of 19 October 17:00. These are staged source records, not Cowork output evidence. Save a one-page closure and one-page draft briefing for 20 October. C01 purchase declined by Teresa in the supplied source; C02 draft prepared but NOT sent/approved; C03 accessibility test completion unknown. The 11:00 review needs 15 minutes preparation; financial reconciliation remains due 22 October. Give top three proposed priorities and source links; no claimed real task completion. Optional real-results route only when the user explicitly selects actual saved results AND separate human approval records: label it as a different route, do not merge staged and real statuses or infer that a draft was sent. No automation.

Human judgment: Set tomorrow's priorities and protect relationship time.

Approval boundary: create documents/drafts; only CO-35 edits its run workbook. One-time approval only. Stop at missing sources, unexpected recipients, send/post/invite/calendar/move/automation actions or inconsistent saved files. Record a limitation and switch to an accepted result; no repair marathon.

## Your live acceptance — all OPEN

| Code | Check outside chat / Außerhalb des Chats prüfen | Status |
|---|---|---|

| CO-30 | Six rows; five categories; duplicate linked; two saved unsent drafts; no approval of EUR 2,200. | OPEN |

| CO-31 | One page; 09:30 deadline; EUR 88,000/eight days; calendar overlap; no afternoon information. | OPEN |

| CO-32 | One page; distinct periods; 96 to 88%, six to fourteen; training decision open. | OPEN |

| CO-33 | Closed / no completion evidence / uncertain separated; no Vance reminder; two unsent drafts. | OPEN |

| CO-34 | Four problem types; two actual drafts; all original calendar entries unchanged. | OPEN |

| CO-35 | Original workbook identity; exact P01/Readme preservation; A01-A03 and blank A02 date; three drafts. | OPEN |

| CO-36 | Three consistent saved documents; safety before walk; 10:00-12:00, 20 conversation minutes; three drafts. | OPEN |

| CO-37 | Two saved one-page documents; declined / draft only / unknown separated; no inferred sends. | OPEN |

Keep actual task links, cloud file identities, draft IDs, screenshots and downloaded accepted outputs privately after your rehearsal. Check saved documents in their native apps, not only chat previews. No agent UI rehearsal was performed. Separate three states: documented capability, locally validated inputs, presenter-accepted live output. Use the explicit synthetic date on stage; 'today' is not a claim that these appointments are current.

## Official references

- https://learn.microsoft.com/microsoft-365/copilot/cowork/
- https://learn.microsoft.com/microsoft-365/copilot/cowork/use-cowork
- https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-faq
- https://learn.microsoft.com/microsoft-365/copilot/cowork/get-started
- https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-customize
- https://learn.microsoft.com/training/modules/empower-workforce-copilot-executives/
- https://support.microsoft.com/en-us/microsoft-365-copilot/frequently-asked-questions-about-retired-copilot-features

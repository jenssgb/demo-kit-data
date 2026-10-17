# Assistant day: start here

Demo Kit: Cowork end to end > Executive assistant: A day with Cowork. The separate leadership-event page remains available.

Use the standard global tenant setup, first with -WhatIf. It installs every demo; do not run a per-case installer. Confirm the signed-in approved demo admin and current Exchange licenses. Actual mail appears in this admin's Inbox with [EAD26 nn-nn]; appointments appear in this admin's own calendar on 19-23 October 2026, under [EAD26 key]. Files: Demo-Kit > Cowork > cowork-assistant-day.

One-time skill setup: download skills/contoso-assistant-day.zip. In Cowork select Customize > Skills > arrow next to Add > Upload skill. Choose the ZIP and keep Only you. If that name already exists, inspect it instead of creating a numbered duplicate. Start a new conversation and select the skill in Sources; verify it is active in the side panel. UI translations/rollout require your tenant check.

For each case, attach runs/en/01/Run_Context.md, inputs/en/<code>/Case.md and exactly the files named there through Add attachments > Attach cloud files. For CO-35 also select that run's original Action_Register.xlsx. Installing data does not select it in Cowork. Check the listed inbox messages/calendar before starting. Use the short request in Demo Kit; inspect persistent files and Outlook Drafts outside chat.

Runs 01-03 are independent clean starting slots. Use a new slot after any edits. If used up, the local Reset-CoworkAssistantDay.ps1 -Language en -WhatIf previews a non-destructive fresh-run reset; remove -WhatIf only after review. It uses the existing cached authorized Graph context, never starts sign-in, downloads and SHA-256 checks the original baseline, and creates a unique run. No calendar or mail reset is needed because live tasks may not change them. If they did, stop; do not silently reinstall.

The installer updates baselines/fixtures and is NOT a reset; existing run workbooks and contexts are preserved even when edited. New reset runs are not manifest-managed and earlier results are retained. For a reset, the ZIP-uploaded skill is unchanged; re-register only an intentional skill revision, avoiding duplicates.

All eight live checks remain OPEN. No new Cowork backups are supplied. Prerequisites: active Copilot license, Cowork enabled, usage billing, reachable cloud files, inbox discovery and your own calendar. No delegation, private executive Teams chats, browser-control add-on or Frontier App skill is required. Stop on missing sources or send/post/invite/move/automation approvals. Presenter expectations are outside this folder in tools/expected.

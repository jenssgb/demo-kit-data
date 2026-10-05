---
name: Contoso go-live readiness
description: Builds a go/no-go readiness summary for a Contoso site go-live. Use when someone asks for a readiness check, go-live status, go/no-go summary or "are we ready for go-live" for a Contoso site such as Fargo.
---

# Contoso go-live readiness

Use this skill when someone asks whether a Contoso site is ready for a go-live, wants a go/no-go summary, or asks for the readiness status of a site project.

## Inputs

1. The readiness tracker in OneDrive (an Excel file whose name contains "Readiness_Tracker"). Use the Readiness table and the Milestones sheet.
2. Recent emails and meeting notes about the site from the last 14 days.
3. If something in an email is newer than the tracker, the email wins. Mention the difference.

## Output

Create one Word document, at most two pages, named "<Site> go-live readiness <date>.docx":

1. **Verdict** in one sentence: Go, Go with conditions, or No-go, plus the main reason.
2. **Status per workstream** as a table: Workstream, RAG, What is open, Owner, Due.
3. **Top three risks** with impact on the go-live date and a concrete mitigation.
4. **Decisions needed** from management, each with a deadline.
5. **Sources**: list the files, emails and meetings used.

Then offer to draft a short email to the stakeholders with the verdict and the document attached. Never send the email without asking first.

## Rules

- Red means the item can move the go-live date. Amber means it needs attention this week. Green means on track.
- Use dates without weekdays, for example "May 22".
- Do not invent numbers. If something is missing, write "tbd".
- Write in the language of the request.

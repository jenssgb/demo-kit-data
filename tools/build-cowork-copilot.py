"""Builds the fictional Contoso demo files for the cowork-copilot demo (Copilot Cowork).

Run:  python tools/build-cowork-copilot.py
Needs: pip install python-docx openpyxl
All content is fictional (Contoso, Tailspin Toys and Fabrikam are Microsoft demo companies).
Scenario: Contoso distribution center Fargo, Dock 3 closes on Apr 27, go-live of the new wing on Jun 22.
"""
from email.message import EmailMessage
from email.utils import format_datetime
from datetime import datetime, timezone, date
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule

OUT = Path(__file__).resolve().parent.parent / "cowork-copilot"
OUT.mkdir(exist_ok=True)
YEAR = 2027  # future year so due dates never look overdue in the demo

NAVY = "0B3D91"

# ---------------------------------------------------------------- emails
MAILS = {
    "1_Megan_Carrier_Update": {
        "from": "Megan Bowen <megan@contoso.com>",
        "en": (
            "Dock 3 – carrier update and Saturday shifts",
            """Hi,

quick update on the carriers before Dock 3 closes on April 27:

- Fabrikam Logistics now agrees to use Gate 4 from April 30 (not May 4). For the three days in between we need a slot at Dock 2 from 7 to 9 am.
- Northwind Freight still arrives at 5:30 am. That is before the vendor window (6 am–8 pm). We need an exception approved by Safety, otherwise they move to 6:15 am and lose their slot in Minneapolis.
- Saturday shifts for wave 1 (A items) are approved for May 2 and May 9. I still need 4 forklift drivers for May 9.

I've added all of this to the readiness tracker.

Thanks,
Megan""",
        ),
        "de": (
            "Dock 3 – Update Speditionen und Samstagsschichten",
            """Hallo,

kurzes Update zu den Speditionen, bevor Dock 3 am 27. April schließt:

- Fabrikam Logistics nutzt Tor 4 jetzt schon ab dem 30. April (nicht erst ab 4. Mai). Für die drei Tage dazwischen brauchen wir einen Slot an Dock 2 von 7 bis 9 Uhr.
- Northwind Freight kommt weiterhin um 5:30 Uhr. Das ist vor dem Dienstleisterfenster (6–20 Uhr). Wir brauchen eine Ausnahme von Safety, sonst müssen sie auf 6:15 Uhr gehen und verlieren ihren Slot in Minneapolis.
- Die Samstagsschichten für Welle 1 (A-Artikel) sind für den 2. und 9. Mai genehmigt. Für den 9. Mai fehlen mir noch 4 Staplerfahrer.

Ich hab alles im Readiness-Tracker eingetragen.

Danke,
Megan""",
        ),
    },
    "2_Riley_Construction_Delay": {
        "from": "Riley Johnson <riley@contoso.com>",
        "en": (
            "New wing – electrical one week late, sprinkler test moved",
            """Hi,

as feared, the cable trays arrived late. Electrical in the new wing will be done on May 22 instead of May 15. Go-live on June 22 is not at risk yet, but the buffer is down to one week.

Two more things:
- The sprinkler test moves from April 13 to April 20. We still need the approval from Facility and the insurance. Who signs this on our side?
- Floor markings for Gate 4 are planned for April 24–25 (Friday/Saturday). Gate 4 is closed during that time.

Riley""",
        ),
        "de": (
            "Neuer Flügel – Elektrik eine Woche später, Sprinklertest verschoben",
            """Hallo,

wie befürchtet kamen die Kabeltrassen zu spät. Die Elektrik im neuen Flügel ist am 22. Mai statt am 15. Mai fertig. Das Go-live am 22. Juni ist noch nicht gefährdet, aber der Puffer ist nur noch eine Woche.

Zwei Dinge noch:
- Der Sprinklertest wird vom 13. auf den 20. April verschoben. Wir brauchen weiterhin die Freigabe von Facility und Versicherung. Wer unterschreibt das bei uns?
- Die Bodenmarkierungen an Tor 4 sind für den 24.–25. April (Freitag/Samstag) geplant. Tor 4 ist in der Zeit zu.

Riley""",
        ),
    },
    "3_Anne_Safety_Training": {
        "from": "Anne Weiler <anne@contoso.com>",
        "en": (
            "Shift lead training before go-live",
            """Hi,

before go-live all 12 shift leads need the training on the new safety rules (escape routes, forklift routes, PPE at Dock 2). I suggest three 90-minute sessions in the week of June 7, four people each, in the training room next to Dock 1.

The evacuation drill for the new wing is planned for June 15. That should happen before the go/no-go decision.

Can you send out the invites or should I?

Anne""",
        ),
        "de": (
            "Schulung Schichtleiter vor dem Go-live",
            """Hallo,

vor dem Go-live brauchen alle 12 Schichtleiter die Schulung zu den neuen Sicherheitsregeln (Fluchtwege, Staplerrouten, PSA an Dock 2). Mein Vorschlag: drei Termine à 90 Minuten in der Woche ab 7. Juni, je vier Leute, im Schulungsraum neben Dock 1.

Die Räumungsübung für den neuen Flügel ist für den 15. Juni geplant. Die sollte vor der Go/No-Go-Entscheidung stattfinden.

Schickst du die Einladungen raus oder soll ich?

Anne""",
        ),
    },
    "4_TailspinToys_Quote_Request": {
        "from": "Jordan Mitchell <jordan.mitchell@tailspintoys.com>",
        "en": (
            "Quote request – 1,200 units T300, delivery Minneapolis",
            """Hello,

we would like a quote for 1,200 units of the Contoso Smart Thermostat T300, delivered to our warehouse in Minneapolis by May 8.

We heard that your Fargo site is under construction. Will that affect the delivery date? If a partial delivery is easier, 600 units by May 8 and the rest by May 22 would also work for us.

Please send the quote as a PDF.

Best regards,
Jordan Mitchell
Purchasing, Tailspin Toys""",
        ),
        "de": (
            "Angebotsanfrage – 1.200 Stück T300, Lieferung Minneapolis",
            """Hallo,

wir hätten gern ein Angebot über 1.200 Stück Contoso Smart Thermostat T300, Lieferung bis 8. Mai an unser Lager in Minneapolis.

Wir haben gehört, dass bei Ihnen in Fargo gebaut wird. Hat das Auswirkungen auf den Liefertermin? Falls eine Teillieferung einfacher ist: 600 Stück bis 8. Mai und den Rest bis 22. Mai wäre für uns auch in Ordnung.

Bitte schicken Sie uns das Angebot als PDF.

Viele Grüße
Jordan Mitchell
Einkauf, Tailspin Toys""",
        ),
    },
}


def build_mails():
    for lang in ("en", "de"):
        folder = OUT / f"mails-{lang}"
        folder.mkdir(exist_ok=True)
        for name, m in MAILS.items():
            subject, body = m[lang]
            msg = EmailMessage()
            msg["From"] = m["from"]
            msg["To"] = ""
            msg["Subject"] = subject
            msg["Date"] = format_datetime(datetime(2026, 10, 1, 9, 0, tzinfo=timezone.utc))
            msg["X-Unsent"] = "1"  # Outlook opens it as a draft you can send to yourself
            msg.set_content(body)
            (folder / f"{name}.eml").write_bytes(bytes(msg))


# ---------------------------------------------------------------- tracker
ROWS = [
    ("R01", "Logistics", "Move inbound to Dock 2 / Gate 4", "Megan", (4, 27), "In progress", "Amber", "Northwind 5:30 am exception still open"),
    ("R02", "Logistics", "Carrier confirmations (Northwind, Fabrikam Logistics)", "Megan", (4, 24), "In progress", "Amber", "Fabrikam Gate 4 from Apr 30; needs Dock 2 slot Apr 27–29"),
    ("R03", "Logistics", "Saturday shifts wave 1", "Megan", (5, 2), "Open", "Amber", "4 forklift drivers missing for May 9"),
    ("R04", "Construction", "Electrical new wing", "Riley", (5, 15), "In progress", "Red", "Cable trays late, new date May 22"),
    ("R05", "Construction", "Floor markings Gate 4", "Riley", (4, 25), "Open", "Amber", "Gate 4 closed Apr 24–25"),
    ("R06", "Construction", "Sprinkler test new wing", "Riley", (4, 13), "Open", "Red", "Facility + insurance approval missing"),
    ("R07", "Safety", "Detour signage east corridor escape route", "Anne", (4, 10), "Done", "Green", ""),
    ("R08", "Safety", "Shift lead training new safety rules", "Anne", (6, 12), "Open", "Amber", "12 shift leads, dates tbd"),
    ("R09", "Safety", "Evacuation drill new wing", "Anne", (6, 15), "Open", "Green", ""),
    ("R10", "Inventory", "Wave 1 move (A items)", "Megan", (5, 9), "Open", "Amber", "Depends on Saturday shifts"),
    ("R11", "Inventory", "Wave 2 move (B/C items)", "Megan", (6, 5), "Open", "Green", ""),
    ("R12", "IT", "Scanners and Wi-Fi new wing", "Alex", (6, 1), "Open", "Green", "Needs electrical done"),
    ("R13", "Customers", "Inform key customers about reduced dock capacity", "Ops lead", (4, 20), "Open", "Amber", ""),
    ("R14", "Go-live", "Go/no-go decision", "Ops lead", (6, 15), "Open", "Green", "After evacuation drill"),
]
MILESTONES = [
    ("Sprinkler test", (4, 13)),
    ("Dock 3 closes, inbound via Dock 2 / Gate 4", (4, 27)),
    ("Wave 1 inventory move done", (5, 9)),
    ("Electrical new wing done", (5, 15)),
    ("Evacuation drill", (6, 15)),
    ("Go/no-go decision", (6, 15)),
    ("Go-live new wing", (6, 22)),
]
FILL = {"Red": "F8D7DA", "Amber": "FFF1CC", "Green": "D9F2E3"}


def build_tracker():
    wb = Workbook()
    ws = wb.active
    ws.title = "Readiness"
    head = ["ID", "Workstream", "Item", "Owner", "Due", "Status", "RAG", "Notes"]
    ws.append(head)
    for r in ROWS:
        ws.append([r[0], r[1], r[2], r[3], date(YEAR, *r[4]), r[5], r[6], r[7]])
    last = len(ROWS) + 1
    tab = Table(displayName="Readiness", ref=f"A1:H{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    for i, w in enumerate([6, 14, 46, 10, 12, 13, 8, 46], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[f"E2:E{last}"]:
        c[0].number_format = "mmm d, yyyy"
    for rag, color in FILL.items():
        ws.conditional_formatting.add(
            f"G2:G{last}", CellIsRule(operator="equal", formula=[f'"{rag}"'], fill=PatternFill("solid", fgColor=color))
        )
    ws.freeze_panes = "A2"

    ms = wb.create_sheet("Milestones")
    ms.append(["Milestone", "Date"])
    for name, d in MILESTONES:
        ms.append([name, date(YEAR, *d)])
    ms.column_dimensions["A"].width = 44
    ms.column_dimensions["B"].width = 14
    for c in ms[f"B2:B{len(MILESTONES) + 1}"]:
        c[0].number_format = "mmm d, yyyy"
    for c in ms[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(vertical="center")
    wb.save(OUT / "Contoso_Fargo_GoLive_Readiness_Tracker.xlsx")


# ---------------------------------------------------------------- weekly notes
def build_weekly():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Segoe UI"
    st.font.size = Pt(11)
    h = doc.add_heading("Construction weekly – week 16 – Fargo expansion", level=1)
    h.runs[0].font.color.rgb = RGBColor.from_string(NAVY)
    doc.add_paragraph("Attending: Anne (Safety), Riley (Construction), Megan (Logistics), operations lead")
    doc.add_heading("Decisions", level=2)
    for t in [
        "Dock 3 closes as planned on April 27. Inbound moves to Dock 2 and Gate 4.",
        "Go-live of the new wing stays on June 22. Go/no-go decision on June 15, after the evacuation drill.",
        "Saturday shifts for wave 1 approved for May 2 and May 9.",
    ]:
        doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("Open points", level=2)
    for t in [
        "Electrical is one week late (May 22 instead of May 15). Riley sends a revised plan.",
        "Sprinkler test moved to April 20. Approval from Facility and insurance still missing. Owner unclear.",
        "Northwind Freight arrives at 5:30 am, before the vendor window. Safety exception needed (Anne).",
        "Shift lead training: 12 people, dates in the week of June 7 still to be set.",
        "Key customers have not been told about the reduced dock capacity yet (operations lead).",
    ]:
        doc.add_paragraph(t, style="List Bullet")
    doc.add_heading("Next meeting", level=2)
    doc.add_paragraph("Thursday, 10:00 am, Dock 1 meeting room and Teams.")
    doc.save(OUT / "Contoso_Construction_Weekly_Week16.docx")


# ---------------------------------------------------------------- Cowork skill
SKILL = """---
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
"""

PLACEHOLDERS = """# SAP demo data for arnold (enosix) – what to swap

The Cowork page uses these fictional values. Replace them with real master data from your SAP demo system
(and edit the Tailspin Toys email accordingly) so arnold finds them.

| In the demo                         | Replace with (your SAP system)          |
|-------------------------------------|-----------------------------------------|
| Customer: Tailspin Toys             | [SAP customer number / sold-to party]   |
| Material: Contoso Smart Thermostat T300 | [SAP material number]               |
| Quantity: 1,200 units (or 600 + 600) | keep or adapt                           |
| Plant / shipping point: Fargo       | [SAP plant / shipping point]            |
| Ship-to: Minneapolis warehouse      | [SAP ship-to party]                     |
| Requested delivery: May 8 / May 22  | keep or adapt                           |

Tip: run the prompt "Show me the open sales orders and quotes for Tailspin Toys" in Cowork once before the demo,
so the arnold sign-in is done and you can see which numbers come back.
"""


def build_skill():
    d = OUT / "skills" / "contoso-golive-readiness"
    d.mkdir(parents=True, exist_ok=True)
    (d / "SKILL.md").write_text(SKILL, encoding="utf-8")
    (OUT / "SAP_Demo_Placeholders.md").write_text(PLACEHOLDERS, encoding="utf-8")


if __name__ == "__main__":
    build_mails()
    build_tracker()
    build_weekly()
    build_skill()
    print("built", sorted(p.relative_to(OUT).as_posix() for p in OUT.rglob("*") if p.is_file()))

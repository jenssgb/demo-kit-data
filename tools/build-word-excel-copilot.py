from pathlib import Path
from datetime import date, timedelta
import random, json
from email.message import EmailMessage
from email.utils import format_datetime
from datetime import datetime, timezone
from docx import Document
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(r"D:\code\demo-kit-data")
OUT = ROOT / "word-excel-copilot"
OUT.mkdir(exist_ok=True)
YEAR = 2027
random.seed(427)

NAVY = "0B3D91"
TEAL = "0F6CBD"
AMBER = "FFF4CE"
RED = "FDE7E9"
GREEN = "DFF6DD"

PEOPLE = {
    "Megan Bowen": "Logistics",
    "Riley Johnson": "Construction",
    "Anne Weiler": "Safety",
    "Alex Wilber": "IT",
    "Ops lead": "Operations",
}

EMAILS = [
    ("Megan Bowen", "Dock 3 cutover: carrier plan for Apr 27", "Dock 3 closes after Friday Apr 26. Northwind keeps the 5:30 am inbound slot, Fabrikam Logistics needs 7-9 am coverage at Dock 2 for Apr 27-29, and Gate 4 opens for Fabrikam from Apr 30. I am worried about Dock 2 on Mondays once Dock 3 traffic moves over."),
    ("Riley Johnson", "New wing construction update", "Electrical completion moves to May 22. Sprinkler test is now Apr 20 and still needs Facility and insurance sign-off. Gate 4 floor markings are Apr 24-25, so do not schedule inbound there then."),
    ("Anne Weiler", "Safety prep before go-live", "All 12 shift leads need training in the week of Jun 7. Evacuation drill for the new wing is planned for Jun 15, before the go/no-go meeting."),
    ("Alex Wilber", "Scanner and Wi-Fi readiness", "Wi-Fi install depends on electrical completion. If electrical holds at May 22, scanner testing can start Jun 1 and still finish before go/no-go."),
]

# ---------------------------------------------------------------- docx source pack

def style_doc(doc):
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10.5)
    for name in ["Heading 1", "Heading 2"]:
        styles[name].font.name = "Arial"
        styles[name].font.color.rgb = RGBColor(11, 61, 145)

def build_docx():
    doc = Document()
    style_doc(doc)
    doc.add_heading("Contoso Fargo go-live source pack", 0)
    p = doc.add_paragraph()
    p.add_run("Scenario: ").bold = True
    p.add_run("New warehouse wing in Fargo. Dock 3 closes on Apr 27; inbound moves to Dock 2 / Gate 4. Go/no-go Jun 15, go-live Jun 22.")
    doc.add_heading("Construction weekly – week 16", level=1)
    rows = [
        ("Construction", "Electrical", "May 22", "Riley", "Late cable trays; buffer reduced to one week."),
        ("Construction", "Sprinkler test", "Apr 20", "Riley", "Facility and insurance approval still missing."),
        ("Logistics", "Gate 4 markings", "Apr 24-25", "Megan", "Gate closed while floor markings are applied."),
        ("Safety", "Shift lead training", "Week of Jun 7", "Anne", "Three 90-minute sessions, four leads each."),
        ("IT", "Scanner / Wi-Fi test", "Jun 1", "Alex", "Depends on electrical completion."),
    ]
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    for i, h in enumerate(["Workstream", "Item", "Date", "Owner", "Note"]):
        table.rows[0].cells[i].text = h
    for r in rows:
        cells = table.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = v
    doc.add_heading("Email snippets", level=1)
    for sender, subject, body in EMAILS:
        doc.add_heading(subject, level=2)
        doc.add_paragraph(f"From: {sender}")
        doc.add_paragraph(body)
    doc.add_heading("Runbook draft notes", level=1)
    for txt in [
        "Purpose: coordinate the Dock 3 closure, new wing readiness and go-live decision.",
        "Audience: Logistics, Construction, Safety, IT and Operations leads.",
        "The runbook should be concise, action-oriented and written for a live operations call.",
        "Include a RACI table and a short daily rhythm from Apr 24 through Jun 22.",
    ]:
        doc.add_paragraph(txt, style="List Bullet")
    path = OUT / "Contoso_Fargo_Runbook_Source_Pack.docx"
    doc.save(path)
    fix_docx_settings(path)


def fix_docx_settings(path):
    import zipfile, shutil
    tmp = path.with_suffix('.tmp.docx')
    with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/settings.xml':
                text = data.decode('utf-8')
                text = text.replace('<w:zoom w:val="bestFit"/>', '<w:zoom w:val="bestFit" w:percent="100"/>')
                data = text.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(str(tmp), str(path))

# ---------------------------------------------------------------- emails

def build_mails():
    folder = OUT / "mails-en"
    folder.mkdir(exist_ok=True)
    for idx, (sender, subject, body) in enumerate(EMAILS, 1):
        msg = EmailMessage()
        msg["From"] = f"{sender} <{sender.split()[0].lower()}@contoso.com>"
        msg["To"] = ""
        msg["Subject"] = subject
        msg["Date"] = format_datetime(datetime(YEAR, 4, 18 + idx, 9, 0, tzinfo=timezone.utc))
        msg["X-Unsent"] = "1"
        msg.set_content(body + "\n\nThanks,\n" + sender.split()[0])
        (folder / f"{idx}_{subject.replace(':','').replace(' ','_')}.eml").write_bytes(bytes(msg))

# ---------------------------------------------------------------- Excel

def daterange(start, end):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)

CARRIERS = ["Fabrikam Logistics", "Northwind Freight", "Contoso Fleet", "Blue Yonder Express", "Adventure Works Transport"]
FAMILIES = ["A items", "B/C items", "T300 Thermostats", "Safety stock", "Fixtures", "Returns"]
CUSTOMERS = ["Fargo DC", "Tailspin Toys", "Contoso Stores", "Fourth Coffee", "City Power & Light"]

def base_shipments_for_day(d):
    if d.weekday() == 6:
        return 0
    base = 2 if d.weekday() == 5 else 3
    if d.weekday() == 0:
        base += 2
    if d.weekday() == 2:
        base += 1
    if d.month == 5 and d.weekday() == 3:
        base += 1
    if date(YEAR,4,28) <= d <= date(YEAR,4,30):
        base += 4
    if d >= date(YEAR,4,27) and d.weekday() in (0,1,2):
        base += 1
    return base

def pick_dock(d, i):
    if d < date(YEAR,4,27):
        return random.choices(["Dock 1", "Dock 2", "Dock 3", "Gate 4"], [20, 35, 35, 10])[0]
    if date(YEAR,4,24) <= d <= date(YEAR,4,25):
        return random.choices(["Dock 1", "Dock 2"], [35, 65])[0]
    return random.choices(["Dock 1", "Dock 2", "Gate 4"], [18, 62, 20])[0]

def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Inbound Shipments"
    headers = ["Shipment ID", "Date", "Weekday", "Dock", "Carrier", "Customer", "Product family", "Pallets", "Cases", "Arrival window", "Unload minutes", "Priority", "Notes"]
    ws.append(headers)
    sid = 1001
    for d in daterange(date(YEAR,3,1), date(YEAR,6,30)):
        n = base_shipments_for_day(d)
        for i in range(n):
            dock = pick_dock(d, i)
            carrier = random.choice(CARRIERS)
            if date(YEAR,4,27) <= d <= date(YEAR,4,30) and i < 3:
                dock = "Dock 2"
                carrier = "Fabrikam Logistics" if i % 2 else "Northwind Freight"
            pallets = random.randint(18, 42)
            if dock == "Dock 2" and (d.weekday() == 0 or date(YEAR,4,28) <= d <= date(YEAR,4,30)):
                pallets += random.randint(12, 30)
            family = random.choice(FAMILIES)
            customer = random.choice(CUSTOMERS)
            if family == "T300 Thermostats": customer = "Tailspin Toys"
            hour = random.choice([5,6,7,8,9,10,13,14,15])
            notes = ""
            priority = random.choice(["Standard", "Standard", "High"])
            if dock == "Dock 2" and d >= date(YEAR,4,27):
                notes = "Rerouted from Dock 3 closure"
            if carrier == "Northwind Freight" and hour == 5:
                notes = (notes + "; " if notes else "") + "Before vendor window; Safety exception needed"
                priority = "High"
            ws.append([f"IN-{sid}", d, d.strftime("%A"), dock, carrier, customer, family, pallets, pallets * random.randint(18, 32), f"{hour:02d}:00-{hour+1:02d}:00", 45 + pallets * 2, priority, notes])
            sid += 1
    last = ws.max_row
    tab = Table(displayName="InboundShipments", ref=f"A1:M{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = "A2"
    widths = [14, 12, 12, 10, 24, 20, 18, 10, 10, 16, 15, 12, 42]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center")
    for c in ws["B"][1:]:
        c.number_format = "mmm d, yyyy"
    for row in ws.iter_rows(min_row=2, max_row=last):
        d, dock = row[1].value, row[3].value
        if dock == "Dock 2" and (d.weekday() == 0 or date(YEAR,4,28) <= d <= date(YEAR,4,30)):
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=AMBER)
    cap = wb.create_sheet("Dock Capacity")
    cap.append(["Dock", "Daily pallet capacity", "Status from Apr 27", "Note"])
    for row in [["Dock 1", 120, "Open", "Outbound priority"], ["Dock 2", 160, "Open", "Receives most Dock 3 inbound"], ["Dock 3", 140, "Closed", "Closed from Apr 27 for construction"], ["Gate 4", 90, "Open from Apr 30", "Unavailable Apr 24-25 for markings"]]:
        cap.append(row)
    for cell in cap[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=TEAL)
    for col in range(1,5):
        cap.column_dimensions[get_column_letter(col)].width = [12,22,18,45][col-1]
    brief = wb.create_sheet("Analysis Brief")
    brief.append(["Question", "Expected finding"])
    rows = [
        ["Which days overload Dock 2 after Dock 3 closes?", "Apr 28-30 and most Mondays exceed the 160-pallet daily capacity."],
        ["Which carrier drives the first overload?", "Fabrikam Logistics and Northwind Freight concentrate inbound on Dock 2 during Apr 27-30."],
        ["What should the chart show?", "Daily Dock 2 pallets vs. 160-pallet capacity line, grouped by weekday."],
        ["What should Python forecast?", "A simple weekday trend forecast through Jun 22 showing Monday risk unless loads move to Gate 4 or Saturday."],
    ]
    for r in rows: brief.append(r)
    for cell in brief[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
    brief.column_dimensions["A"].width = 42
    brief.column_dimensions["B"].width = 90
    wb.save(OUT / "Contoso_Fargo_Inbound_Shipments_Mar_Jun.xlsx")
    return last-1

# ---------------------------------------------------------------- manifest

def build_manifest():
    manifest = {
        "title": {"en": "Edit with Copilot in Word and Excel – Contoso Fargo", "de": "Edit with Copilot in Word und Excel – Contoso Fargo"},
        "next": {
            "en": [
                "Upload the files to OneDrive or SharePoint and make sure AutoSave is on before using Edit with Copilot in Word or Excel.",
                "Word demo: open Contoso_Fargo_Runbook_Source_Pack.docx in Word, then use Home > Copilot with Allow editing on.",
                "Excel demo: open Contoso_Fargo_Inbound_Shipments_Mar_Jun.xlsx in Excel, then use the Copilot button in the lower-right corner and keep Allow editing on.",
                "Emails: optional context. Open each .eml in mails-en, address it to yourself and send it before the demo so Copilot can reference recent email context.",
                "Check that Anthropic/Claude models are enabled for the presenter if you want to show model choice."
            ],
            "de": [
                "Dateien in OneDrive oder SharePoint hochladen und AutoSave einschalten, bevor Edit with Copilot in Word oder Excel genutzt wird.",
                "Word-Demo: Contoso_Fargo_Runbook_Source_Pack.docx in Word öffnen, dann Start > Copilot mit aktivem Bearbeiten erlauben nutzen.",
                "Excel-Demo: Contoso_Fargo_Inbound_Shipments_Mar_Jun.xlsx in Excel öffnen, dann den Copilot-Button unten rechts nutzen und Bearbeiten erlauben eingeschaltet lassen.",
                "Mails: optionaler Kontext. Jede .eml in mails-en öffnen, an dich selbst adressieren und vor der Demo senden, damit Copilot aktuelle Mails referenzieren kann.",
                "Prüfen, ob Anthropic-/Claude-Modelle für den Presenter freigeschaltet sind, falls die Modellwahl gezeigt werden soll."
            ],
        },
    }
    (OUT / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

build_docx()
build_mails()
rows = build_workbook()
build_manifest()
print(f"Built {OUT} with {rows} shipment rows")


"""Builds small fictional Contoso files for the notebooks-copilot demo.

Run:  python tools/build-notebooks-copilot.py
Needs: pip install python-docx
All content is fictional (Contoso, Tailspin Toys, Fabrikam Logistics, Northwind are Microsoft demo companies).
Scenario: Contoso distribution center Fargo, Dock 3 closes on Apr 27, go-live of the new wing on Jun 22.
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

OUT = Path(__file__).resolve().parent.parent / "notebooks-copilot"
OUT.mkdir(exist_ok=True)
YEAR = 2027
NAVY = RGBColor(11, 61, 145)


def style_doc(doc: Document):
    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)
    for name in ("Title", "Heading 1", "Heading 2"):
        styles[name].font.name = "Aptos Display"
        styles[name].font.color.rgb = NAVY


def add_title(doc: Document, title: str, subtitle: str):
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    r = p.add_run(title)
    r.bold = True
    r.font.name = "Aptos Display"
    r.font.size = Pt(20)
    r.font.color.rgb = NAVY
    s = doc.add_paragraph(subtitle)
    s.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    s.runs[0].italic = True


def build_transcript_docx():
    doc = Document()
    style_doc(doc)
    add_title(doc, "Dock 3 go-live sync", f"Teams meeting recap and transcript · Contoso Fargo · Apr 24, {YEAR}")
    doc.add_heading("Meeting recap", level=1)
    rows = [
        ("Decision", "Inbound moves from Dock 3 to Dock 2 on Apr 27. Gate 4 is the preferred staging area after floor markings finish."),
        ("Risk", "Fabrikam Logistics can use Gate 4 from Apr 30, but Apr 27–29 need a temporary Dock 2 slot from 7:00–9:00."),
        ("Risk", "Northwind Freight still arrives at 5:30, outside the 6:00–20:00 vendor window. Safety exception needed."),
        ("Dependency", "Sprinkler test moved to Apr 20; Facility and insurance approval still open."),
        ("Dependency", "Electrical completion moved to May 22. Wi-Fi and scanner setup cannot finish before electrical is signed off."),
        ("Action", "Anne to confirm evacuation drill timing for Jun 15 before the go/no-go meeting."),
        ("Action", "Megan to confirm four forklift drivers for the May 9 Saturday shift."),
        ("Action", "Ops lead to prepare a night-shift handover page by Apr 26."),
    ]
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Type"
    table.rows[0].cells[1].text = "Note"
    for typ, note in rows:
        cells = table.add_row().cells
        cells[0].text = typ
        cells[1].text = note
    doc.add_heading("Transcript excerpt", level=1)
    lines = [
        ("Megan Bowen", "The main thing for logistics is that Dock 3 is effectively out from Monday morning. Fabrikam can switch to Gate 4 on Apr 30, but the first three days are tight."),
        ("Riley Johnson", "Gate 4 markings happen Apr 24 and 25. Please do not route carriers there until I confirm the paint is cured and the cones are removed."),
        ("Anne Weiler", "Northwind at 5:30 is before the approved vendor window. I can support an exception, but I need the new Dock 2 marshal plan attached."),
        ("Alex Wilber", "Scanner setup depends on electrical sign-off. If May 22 slips, Wi-Fi in the new wing becomes a red risk."),
        ("Ops Lead", "I'll put all of this in the Fargo notebook and create a night-shift handover so supervisors can run the same checklist."),
    ]
    for speaker, text in lines:
        p = doc.add_paragraph()
        p.add_run(f"{speaker}: ").bold = True
        p.add_run(text)
    doc.save(OUT / "Dock_3_GoLive_Sync_Transcript.docx")


def build_vtt():
    text = f"""WEBVTT

00:00:00.000 --> 00:00:05.000
Megan Bowen: Dock 3 is effectively out from Monday morning, April 27.

00:00:05.500 --> 00:00:12.000
Megan Bowen: Fabrikam can switch to Gate 4 on April 30, but April 27 to 29 need a Dock 2 slot from 7 to 9.

00:00:12.500 --> 00:00:19.500
Riley Johnson: Gate 4 markings happen April 24 and 25. Do not route carriers there until the cones are removed.

00:00:20.000 --> 00:00:28.000
Anne Weiler: Northwind at 5:30 is before the approved vendor window. I need the Dock 2 marshal plan before approving the exception.

00:00:28.500 --> 00:00:36.500
Alex Wilber: Scanner setup depends on electrical sign-off. If May 22 slips, Wi-Fi in the new wing becomes a red risk.

00:00:37.000 --> 00:00:44.000
Ops Lead: I'll put this in the Fargo notebook and create a night-shift handover page for supervisors.
"""
    (OUT / "Dock_3_GoLive_Sync.vtt").write_text(text, encoding="utf-8")


def build_handover_template():
    doc = Document()
    style_doc(doc)
    add_title(doc, "Fargo go-live · night shift handover", "Template for supervisors · Contoso distribution center")
    doc.add_heading("Shift context", level=1)
    for item in [
        "Dock 3 closed; inbound uses Dock 2 unless a specific Gate 4 slot is confirmed.",
        "Gate 4 floor markings must be clear before carriers enter the area.",
        "Northwind before-hours arrival needs a Safety exception and a Dock 2 marshal.",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_heading("Checklist", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, h in enumerate(["Area", "Check", "Owner", "Done"]):
        table.rows[0].cells[i].text = h
    for row in [
        ("Inbound", "Confirm Dock 2 slots for Apr 27–29", "Megan", "☐"),
        ("Safety", "Attach marshal plan to Northwind exception", "Anne", "☐"),
        ("Construction", "Confirm Gate 4 markings are open", "Riley", "☐"),
        ("IT", "Escalate if electrical sign-off moves after May 22", "Alex", "☐"),
        ("Ops", "Update readiness tracker before 06:00", "Ops lead", "☐"),
    ]:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = val
    doc.add_heading("Open handoff notes", level=1)
    doc.add_paragraph("Use this section for exceptions, carrier delays, and supervisor decisions before the morning stand-up.")
    doc.save(OUT / "Fargo_Night_Shift_Handover_Template.docx")


def build_manifest():
    manifest = '''{
  "title": {
    "en": "Copilot Notebooks – Contoso Fargo go-live",
    "de": "Copilot Notebooks – Go-live Contoso Fargo"
  },
  "next": {
    "en": [
      "Keep these files in the OneDrive folder Demo-notebooks-copilot.",
      "Also run the existing one-liners for Demo-cowork-copilot and Demo-powerpoint-copilot; the notebook references the tracker, weekly notes, emails, and brand guidelines from those folders.",
      "In Teams, use Dock_3_GoLive_Sync_Transcript.docx or Dock_3_GoLive_Sync.vtt as the transcript source for a demo meeting recap if a real Teams recap is not available.",
      "Use Fargo_Night_Shift_Handover_Template.docx as the source/template when Copilot creates the handover page."
    ],
    "de": [
      "Diese Dateien im OneDrive-Ordner Demo-notebooks-copilot lassen.",
      "Zusätzlich die vorhandenen One-Liner für Demo-cowork-copilot und Demo-powerpoint-copilot ausführen; das Notebook referenziert Tracker, Jour-fixe-Notizen, Mails und Brand Guidelines aus diesen Ordnern.",
      "In Teams Dock_3_GoLive_Sync_Transcript.docx oder Dock_3_GoLive_Sync.vtt als Transcript-Quelle für einen Demo-Meeting-Recap nutzen, falls kein echter Teams-Recap vorhanden ist.",
      "Fargo_Night_Shift_Handover_Template.docx als Quelle/Vorlage verwenden, wenn Copilot die Handover Page erstellt."
    ]
  }
}
'''
    (OUT / "manifest.json").write_text(manifest, encoding="utf-8")


if __name__ == "__main__":
    build_transcript_docx()
    build_vtt()
    build_handover_template()
    build_manifest()
    print(f"Built {OUT}")

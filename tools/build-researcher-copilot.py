"""Builds the fictional Contoso demo files for the researcher-copilot demo.

Run: python tools/build-researcher-copilot.py
Needs: pillow
All content is fictional (Contoso). Scenario: Fargo distribution center, Dock 3 closes Apr 27,
inbound moves to Dock 2 / Gate 4, go-live of the new wing on Jun 22.
"""
from pathlib import Path
import json
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent.parent / "researcher-copilot"
OUT.mkdir(exist_ok=True)


def font(name="arial.ttf", size=22):
    try:
        return ImageFont.truetype(name, size)
    except Exception:
        return None


def build_floor_plan():
    W, H = 1600, 1000
    img = Image.new("RGB", (W, H), "#F8FAFC")
    d = ImageDraw.Draw(img)
    font_title = font(size=42)
    font_h = font(size=30)
    font_m = font(size=22)
    font_s = font(size=18)

    def rect(x1, y1, x2, y2, fill, outline="#334155", width=3, label=None, lf=None):
        d.rounded_rectangle([x1, y1, x2, y2], radius=16, fill=fill, outline=outline, width=width)
        if label:
            d.multiline_text((x1 + 18, y1 + 16), label, fill="#0F172A", font=lf or font_m, spacing=6)

    d.rounded_rectangle([70, 100, 1530, 900], radius=28, fill="#FFFFFF", outline="#0F172A", width=5)
    d.text((70, 35), "Contoso Fargo Distribution Center – Go-live visual (fictional)", fill="#0F172A", font=font_title)
    d.text((70, 925), "Scenario notes: Dock 3 closes Apr 27; inbound moves to Dock 2 / Gate 4; new wing go-live Jun 22.", fill="#334155", font=font_m)

    rect(120, 150, 570, 520, "#E0F2FE", label="Existing warehouse\nHall 1")
    rect(600, 150, 1030, 520, "#E2E8F0", label="Packing + staging\nForklift route changes")
    rect(1060, 150, 1480, 520, "#DCFCE7", label="NEW WING\nGo-live Jun 22\nElectrical due May 22")
    rect(120, 560, 760, 850, "#FEF3C7", label="Dock operations\nInbound temporary flow")
    rect(790, 560, 1480, 850, "#FEE2E2", label="Construction buffer / safety zone\nPPE required")
    rect(150, 600, 330, 820, "#DBEAFE", outline="#2563EB", width=5, label="DOCK 2\nINBOUND\nfrom Apr 27", lf=font_h)
    rect(370, 600, 550, 820, "#FECACA", outline="#DC2626", width=6, label="DOCK 3\nCLOSED\nApr 27", lf=font_h)
    rect(590, 600, 730, 820, "#DBEAFE", outline="#2563EB", width=5, label="Gate 4\nTruck entry", lf=font_h)

    for pts in [((720, 700), (600, 700)), ((600, 700), (340, 700)), ((340, 700), (240, 700))]:
        d.line(pts, fill="#2563EB", width=10)
    d.polygon([(250, 700), (275, 685), (275, 715)], fill="#2563EB")
    d.text((575, 850), "Temporary truck path: Gate 4 → Dock 2", fill="#2563EB", font=font_m)
    d.line([(390, 620), (530, 800)], fill="#DC2626", width=10)
    d.line([(530, 620), (390, 800)], fill="#DC2626", width=10)
    points = [(1300, 520), (1300, 450), (1020, 450), (1020, 560), (900, 560)]
    d.line(points, fill="#16A34A", width=8)
    d.text((1080, 425), "Evacuation detour: Mezzanine A → North stairs → Assembly C", fill="#166534", font=font_s)

    callouts = [
        (1085, 250, "New wing", "Sprinkler test Apr 20;\nFacility + insurance approval open"),
        (830, 620, "Safety", "Forklift/pedestrian separation\nneeds shift-lead briefing"),
        (170, 230, "Comparable DC question", "How do others handle go-live\nwhen one dock is closed?"),
    ]
    for x, y, title, body in callouts:
        d.rounded_rectangle([x, y, x + 350, y + 110], radius=12, fill="#FFFFFF", outline="#64748B", width=2)
        d.text((x + 14, y + 12), title, fill="#0F172A", font=font_m)
        d.multiline_text((x + 14, y + 46), body, fill="#475569", font=font_s, spacing=4)
    img.save(OUT / "Contoso_Fargo_Floor_Plan_Dock3_Gate4.png")


def build_manifest():
    manifest = {
        "title": {"en": "Researcher + Vision – Contoso Fargo floor plan", "de": "Researcher + Vision – Contoso Fargo Grundriss"},
        "next": {
            "en": [
                "Use existing cowork-copilot files for the Researcher briefing: the readiness tracker, week 16 construction notes and sent demo emails.",
                "Use Contoso_Fargo_Floor_Plan_Dock3_Gate4.png for the Vision part: open it, start a voice chat in Copilot and share the screen, or show it to the mobile camera.",
                "Check that Researcher, model choice and Vision are enabled in the tenant before the session. Some features require Frontier or Anthropic admin settings.",
            ],
            "de": [
                "Für das Researcher-Briefing die bestehenden cowork-copilot-Dateien nutzen: Readiness-Tracker, Bau-Jour-fixe KW 16 und gesendete Demo-Mails.",
                "Contoso_Fargo_Floor_Plan_Dock3_Gate4.png für den Vision-Teil nutzen: öffnen, Sprachchat in Copilot starten und Bildschirm teilen oder mit der mobilen Kamera zeigen.",
                "Vor dem Termin prüfen, ob Researcher, Modellauswahl und Vision im Tenant aktiv sind. Einige Funktionen brauchen Frontier oder die Anthropic-Adminfreigabe.",
            ],
        },
    }
    (OUT / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    build_floor_plan()
    build_manifest()
    print("built", sorted(p.relative_to(OUT).as_posix() for p in OUT.rglob("*") if p.is_file()))

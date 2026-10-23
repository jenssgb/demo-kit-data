"""Build bilingual WIQ26 evidence using the existing port builder's DOCX conventions.

Original Contoso data, not copies of Microsoft documents. Reuses python-docx to
keep typography and Office packages consistent with the existing corpus.
"""
import json
import runpy
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = runpy.run_path(str(ROOT / "tools" / "build-port-operations.py"))
DOC = PORT["write_doc"]
FOLDER = ROOT / "workiq-port-operations"
DOMAIN = PORT["DOMAIN"]


def write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def build():
    documents = []

    def document(number, slug, title, en, de):
        for lang, text in (("en", en), ("de", de)):
            name = f"{lang}/WIQ-D{number:02}_{slug}.docx"
            path = FOLDER / name
            path.parent.mkdir(parents=True, exist_ok=True)
            sections = [
                ("WIQ26 | Evidence" if lang == "en" else "WIQ26 | Nachweis", text),
                ("Exercise clock" if lang == "en" else "Übungszeit",
                 "Monday and times in this document describe fictional exercise chronology, not today's date. "
                 "Use the actual WIQ26 calendar invitation for the rehearsal. This is not a meeting transcript."
                 if lang == "en" else
                 "Montag und Uhrzeiten beschreiben die fiktive Reihenfolge der Übung, nicht das heutige Datum. "
                 "Für die Generalprobe gilt die wirkliche WIQ26-Kalendereinladung. Dies ist kein Besprechungstranskript."),
            ]
            DOC(path, f"WIQ26 / WIQ-D{number:02} / {title[lang]}", sections)
            documents.append(name)

    document(1, "brief", {"en": "Contoso port handover", "de": "Contoso-Hafenübergabe"},
             "Contoso coordinates twelve shipments and 115 containers: Rotterdam 60, Hamburg 55. Four held shipments cover "
             "43 containers; eight ready shipments cover 72. Teresa Sac follows up CT-1002 (12, corrected manifest) and "
             "CT-1005 (14, release approval). Sonia Rees follows up CT-1003 (6, inspection outcome). Billie Vester follows "
             "up CT-1008 (11, written carrier cut-off). At exercise Monday 15:00 the CT-1005 packing list has arrived, "
             "but release approval is still pending. The CT-1008 provisional slot is not written confirmation. "
             "No release or delivery date is confirmed. Current status: WIQ-D03; ownership: WIQ-D16 to WIQ-D18.",
             "Contoso koordiniert zwölf Sendungen mit 115 Containern: Rotterdam 60, Hamburg 55. Vier gesperrte Sendungen "
             "umfassen 43 Container, acht bereite Sendungen 72. Teresa Sac klärt CT-1002 (12, korrigiertes Manifest) und "
             "CT-1005 (14, Freigabe). Sonia Rees klärt CT-1003 (6, Prüfergebnis). Billie Vester klärt CT-1008 "
             "(11, schriftlicher Annahmeschluss). Am Übungsmontag um 15 Uhr ist die Packliste für CT-1005 eingegangen; "
             "die Freigabe bleibt offen. Der vorläufige Slot für CT-1008 ist keine schriftliche Bestätigung. "
             "Keine Freigabe und kein Liefertermin sind bestätigt. Aktueller Stand: WIQ-D03; Zuständigkeiten: WIQ-D16 bis WIQ-D18.")
    document(2, "status_v1", {"en": "14:00 snapshot - superseded for CT-1005", "de": "Stand 14 Uhr - für CT-1005 überholt"},
             "Exercise Monday 14:00, version 1. CT-1005: 14 containers, packing list missing, release approval pending. "
             "CT-1002 corrected manifest missing; CT-1003 inspection outcome missing; CT-1008 written carrier cut-off missing. "
             "Four holds / 43 containers. WIQ-D03 and Teresa's [WIQ26-M02] supersede only CT-1005 document receipt at 15:00. "
             "This snapshot does not establish a later release.",
             "Übungsmontag 14 Uhr, Version 1. CT-1005: 14 Container, Packliste fehlt, Freigabe offen. "
             "CT-1002: korrigiertes Manifest fehlt; CT-1003: Prüfergebnis fehlt; CT-1008: schriftlicher Annahmeschluss fehlt. "
             "Vier Sperren / 43 Container. WIQ-D03 und Teresas [WIQ26-M02] ersetzen um 15 Uhr nur den Dokumenteneingang "
             "für CT-1005. Dieser Stand belegt keine spätere Freigabe.")
    document(3, "status_v2", {"en": "15:00 update - current decision evidence", "de": "Update 15 Uhr - aktueller Entscheidungsstand"},
             "Exercise Monday 15:00, version 2. CT-1005 packing list RECEIVED; release approval NOT received. "
             "The shipment's 14 containers remain on hold. Teresa's [WIQ26-M02] confirms the same change. "
             "CT-1002: corrected manifest still missing (12). CT-1003: authorized inspection outcome still missing (6). "
             "CT-1008: provisional slot only; written carrier cut-off confirmation missing (11). "
             "Four holds remain open, 43 containers in total. No delivery commitment is authorized.",
             "Übungsmontag 15 Uhr, Version 2. Packliste für CT-1005 EINGEGANGEN; Freigabe NICHT eingegangen. "
             "Die 14 Container bleiben gesperrt. Teresas [WIQ26-M02] bestätigt dieselbe Änderung. "
             "CT-1002: korrigiertes Manifest fehlt weiter (12). CT-1003: autorisiertes Prüfergebnis fehlt weiter (6). "
             "CT-1008: nur vorläufiger Slot; schriftlicher Annahmeschluss fehlt (11). "
             "Vier Sperren bleiben offen, insgesamt 43 Container. Keine Lieferzusage ist autorisiert.")
    for index, shipment in enumerate(PORT["SHIPMENTS"], 4):
        sid, port, consignee, containers, delay, status, owner, action, due = shipment
        current = (
            " At 15:00 the packing list arrived; release approval is still pending. See WIQ-D03 and [WIQ26-M02]."
            if sid == "CT-1005" else
            " The slot is provisional; written cut-off confirmation is still missing. See [WIQ26-M04]."
            if sid == "CT-1008" else ""
        )
        current_de = (
            " Um 15 Uhr ist die Packliste eingegangen; die Freigabe bleibt offen. Siehe WIQ-D03 und [WIQ26-M02]."
            if sid == "CT-1005" else
            " Der Slot ist vorläufig; der schriftliche Annahmeschluss fehlt weiter. Siehe [WIQ26-M04]."
            if sid == "CT-1008" else ""
        )
        document(index, sid.lower().replace("-", "_"), {"en": f"{sid} shipment record", "de": f"{sid} Sendungsnachweis"},
                 f"Contoso WIQ26, exercise Monday 15:00. {sid}, {port}, consignee {consignee}, {containers} containers. "
                 f"Original status: {status}; recorded delay {delay} hours. Follow-up owner: {owner}. "
                 f"Original next action: {action}; exercise due: {due}.{current} "
                 "Ready means ready for coordination, not delivered. No release or delivery promise is established.",
                 f"Contoso WIQ26, Übungsmontag 15 Uhr. {sid}, {port}, Empfänger {consignee}, {containers} Container. "
                 f"Ausgangsstatus im Quellsystem: {status}; erfasste Verzögerung {delay} Stunden. Zuständig: {owner}."
                 f"{current_de} 'Ready' bedeutet bereit zur Koordination, nicht zugestellt. "
                 "Es liegt keine Freigabe oder Lieferzusage vor.")
    extras = [
        ("teresa", "Teresa's follow-up", "Teresas Aufgaben",
         "Teresa Sac owns CT-1002: corrected manifest for 12 containers; CT-1005: release approval for 14. "
         "CT-1005's packing list arrived at 15:00. Ask for the authorized release decision, not another copy of the packing list. "
         "The two document holds cover 26 containers. Cite WIQ-D03 and [WIQ26-M02].",
         "Teresa Sac klärt CT-1002: korrigiertes Manifest für 12 Container; CT-1005: Freigabe für 14. "
         "Die Packliste für CT-1005 ist um 15 Uhr eingegangen. Nach der autorisierten Freigabe fragen, nicht erneut nach der "
         "Packliste. Die beiden Dokumentensperren betreffen 26 Container. WIQ-D03 und [WIQ26-M02] belegen dies."),
        ("sonia", "Sonia's follow-up", "Sonias Aufgaben",
         "Sonia Rees owns CT-1003, six containers. At 15:10 the authorized inspection outcome is not available. "
         "No pass result or clearance can be inferred. Ask which authorized evidence is still due; see [WIQ26-M03].",
         "Sonia Rees klärt CT-1003, sechs Container. Um 15:10 Uhr liegt das autorisierte Prüfergebnis nicht vor. "
         "Eine bestandene Prüfung oder Freigabe lässt sich nicht ableiten. Nach dem noch ausstehenden autorisierten "
         "Nachweis fragen; siehe [WIQ26-M03]."),
        ("billie", "Billie's follow-up", "Billies Aufgaben",
         "Billie Vester owns CT-1008, eleven containers. A provisional slot exists, but no written carrier cut-off "
         "confirmation. The carrier hold remains open. Do not infer a delivery date; see [WIQ26-M04].",
         "Billie Vester klärt CT-1008, elf Container. Ein vorläufiger Slot besteht, aber keine schriftliche Bestätigung "
         "des Annahmeschlusses. Die Speditionssperre bleibt offen. Keinen Liefertermin ableiten; siehe [WIQ26-M04]."),
        ("approval_control", "Receipt is not approval", "Eingang ist keine Freigabe",
         "WIQ26 decision control. Receiving a packing list, booking a provisional slot or drafting an email does not "
         "clear a hold. An authorized decision must be separately evidenced. This training provides no customs, "
         "safety, legal or live routing advice. An operations lead reviews all stakeholder drafts.",
         "WIQ26-Entscheidungskontrolle. Der Eingang einer Packliste, ein vorläufiger Slot oder ein Mailentwurf heben "
         "keine Sperre auf. Eine autorisierte Entscheidung muss getrennt belegt sein. Die Übung liefert keine Zoll-, "
         "Sicherheits-, Rechts- oder echte Routenberatung. Die operative Leitung prüft alle Entwürfe."),
        ("exercise_clock", "Chronology and real rehearsal", "Übungszeit und echte Generalprobe",
         "WIQ26 Monday 14:00 precedes 15:00. Emails are actually delivered during installation; their received timestamps "
         "are not fabricated historical times. Document version and stated exercise time establish source precedence. "
         "Use the actual next WIQ26 invitation date, not a guessed date or a transcript.",
         "WIQ26: Montag 14 Uhr liegt vor 15 Uhr. Mails werden beim Setup wirklich verschickt; ihre Empfangszeitstempel "
         "sind keine nachgebauten historischen Zeiten. Dokumentversion und angegebene Übungszeit bestimmen den "
         "Quellenstand. Das wirkliche Datum der nächsten WIQ26-Einladung verwenden, kein geratenes Datum oder Transkript."),
        ("comparison", "Different shipment - not WIQ26", "Andere Sendung - nicht WIQ26",
         "Comparison material only: fictional Contoso Rotterdam shipment CP-2005, three containers, completed in an "
         "unrelated exercise. This is NOT CT-1005 and provides no approval or delivery evidence for WIQ26.",
         "Nur Vergleichsmaterial: fiktive Contoso-Sendung CP-2005, Rotterdam, drei Container, in einer anderen Übung "
         "abgeschlossen. Dies ist NICHT CT-1005 und belegt keine Freigabe oder Lieferung für WIQ26."),
        ("obsolete_draft", "Rejected draft - do not use as a decision", "Verworfener Entwurf - kein Entscheidungsnachweis",
         "Rejected WIQ26 draft, before review: 'The packing list will mean CT-1005 can be delivered.' "
         "Teresa rejected that assumption. It is not an operational approval. WIQ-D03 and [WIQ26-M02] are the current "
         "evidence: packing list received, release pending, no delivery commitment.",
         "Verworfener WIQ26-Entwurf vor der Prüfung: 'Mit der Packliste kann CT-1005 geliefert werden.' "
         "Teresa hat diese Annahme verworfen. Sie ist keine operative Freigabe. WIQ-D03 und [WIQ26-M02] sind aktuell: "
         "Packliste eingegangen, Freigabe offen, keine Lieferzusage."),
        ("handover_notes", "Handover planning notes", "Planungsnotizen zur Übergabe",
         "Synthetic planning notes, not a transcript. The next WIQ26 rehearsal asks Teresa about the CT-1002 manifest "
         "and CT-1005 decision, Sonia about CT-1003 inspection, and Billie about CT-1008 written confirmation. "
         "Review WIQ-D03, WIQ-D16, WIQ-D17 and WIQ-D18. No decisions have been made in a recorded meeting.",
         "Synthetische Planungsnotizen, kein Transkript. In der nächsten WIQ26-Generalprobe Teresa zum Manifest für "
         "CT-1002 und zur Entscheidung für CT-1005 fragen, Sonia zur Prüfung für CT-1003 und Billie zur schriftlichen "
         "Bestätigung für CT-1008. WIQ-D03, WIQ-D16, WIQ-D17 und WIQ-D18 prüfen. Es wurden keine Entscheidungen "
         "in einer aufgezeichneten Besprechung getroffen."),
        ("source_map", "Source map and review", "Quellenübersicht und Prüfung",
         "WIQ26 source map. WIQ-D01: scope; WIQ-D02: 14:00 snapshot; WIQ-D03: current 15:00 status; "
         "WIQ-D04 to WIQ-D15: twelve shipment records; WIQ-D16 to WIQ-D18: owners; WIQ-D19: approval control. "
         "Inbox [WIQ26-M01]: Teresa at 14:00; M02: Teresa at 15:00; M03: Sonia; M04: Billie. "
         "Both language folders describe the same evidence, not 48 different business events.",
         "WIQ26-Quellenübersicht. WIQ-D01: Umfang; WIQ-D02: Stand 14 Uhr; WIQ-D03: aktueller Stand 15 Uhr; "
         "WIQ-D04 bis WIQ-D15: zwölf Sendungen; WIQ-D16 bis WIQ-D18: Zuständigkeiten; WIQ-D19: Freigabekontrolle. "
         "Posteingang [WIQ26-M01]: Teresa um 14 Uhr; M02: Teresa um 15 Uhr; M03: Sonia; M04: Billie. "
         "Beide Sprachordner beschreiben dieselben Nachweise, nicht 48 verschiedene Geschäftsvorgänge."),
    ]
    for number, (slug, en_title, de_title, en, de) in enumerate(extras, 16):
        document(number, slug, {"en": en_title, "de": de_title}, en, de)
    mails = [
        ("Teresa Sac", "TeresaS", "14:00 document snapshot", "Dokumentenstand 14 Uhr",
         "Exercise Monday 14:00: CT-1002 needs a corrected manifest (12 containers). CT-1005 needs its packing list "
         "(14 containers); release approval is pending. Please check the later 15:00 update before briefing the team.",
         "Übungsmontag 14 Uhr: Für CT-1002 fehlt das korrigierte Manifest (12 Container). Für CT-1005 fehlt die Packliste "
         "(14 Container); die Freigabe ist offen. Vor dem Briefing bitte das spätere Update von 15 Uhr prüfen."),
        ("Teresa Sac", "TeresaS", "15:00 packing list received - approval pending", "15 Uhr: Packliste da - Freigabe offen",
         "The CT-1005 packing list arrived at exercise Monday 15:00. Release approval has NOT arrived. All 14 containers "
         "remain on hold. This supersedes document receipt in M01, not the release decision. WIQ-D03 agrees. "
         "CT-1002's corrected manifest remains missing. Please do not promise delivery.",
         "Die Packliste für CT-1005 ist am Übungsmontag um 15 Uhr eingegangen. Die Freigabe liegt NICHT vor. "
         "Alle 14 Container bleiben gesperrt. Dies ersetzt in M01 den Dokumentenstand, nicht die Freigabeentscheidung. "
         "WIQ-D03 bestätigt das. Für CT-1002 fehlt weiter das korrigierte Manifest. Bitte keine Lieferung zusagen."),
        ("Sonia Rees", "SoniaR", "15:10 inspection outcome still missing", "15:10 Uhr: Prüfergebnis weiter offen",
         "Exercise Monday 15:10: CT-1003 has six containers. I am following up the authorized inspection outcome. "
         "It has not been received. The inspection hold stays open. There is no pass result or release approval.",
         "Übungsmontag 15:10 Uhr: CT-1003 umfasst sechs Container. Ich kläre das autorisierte Prüfergebnis. "
         "Es liegt noch nicht vor. Die Prüfsperre bleibt offen. Weder ein positives Prüfergebnis noch eine Freigabe liegen vor."),
        ("Billie Vester", "BillieV", "15:05 provisional slot is not confirmation", "15:05 Uhr: Vorläufiger Slot ist keine Bestätigung",
         "Exercise Monday 15:05: CT-1008 has eleven containers. The carrier reserved a provisional slot but has not "
         "confirmed the cut-off in writing. I will follow up. Keep Carrier hold open; there is no delivery commitment.",
         "Übungsmontag 15:05 Uhr: CT-1008 umfasst elf Container. Die Spedition hat einen vorläufigen Slot reserviert, "
         "den Annahmeschluss aber nicht schriftlich bestätigt. Ich hake nach. Carrier hold bleibt offen; es gibt keine Lieferzusage."),
    ]
    templates = []
    for number, (person, alias, en_subject, de_subject, en_body, de_body) in enumerate(mails, 1):
        template = f"tenant/{{lang}}/WIQ-M{number:02}.eml"
        templates.append(template)
        for lang, subject, body in (("en", en_subject, en_body), ("de", de_subject, de_body)):
            message = EmailMessage()
            message["From"] = f"{person} <{alias}@{DOMAIN}>"
            message["Subject"] = f"[WIQ26-M{number:02}] Contoso - {subject}"
            message["Date"] = datetime(2026, 10, 23, 9, number, tzinfo=timezone.utc)
            message.set_content(body + ("\n\nSynthetic Contoso training data.\n" if lang == "en" else "\n\nSynthetische Contoso-Übungsdaten.\n"))
            path = FOLDER / template.replace("{lang}", lang)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(message.as_bytes())
    write_json(FOLDER / "tenant.json", {
        "requiredUsers": [f"{alias}@{DOMAIN}" for alias in ("TeresaS", "SoniaR", "BillieV")],
        "mails": templates,
        "workIqShared": {
            "scenario": "WIQ26", "folder": "Demo-Kit-Work-IQ-WIQ26", "documents": documents,
            "event": {
                "subject": {"en": "[WIQ26] Contoso port handover - rehearsal",
                            "de": "[WIQ26] Contoso-Hafenuebergabe - Generalprobe"},
                "body": {
                    "en": "Synthetic WIQ26 rehearsal. Teresa: CT-1002 manifest and CT-1005 release approval; Sonia: "
                          "CT-1003 inspection; Billie: CT-1008 written cut-off. Four holds remain open. "
                          "This invitation does not create a transcript.",
                    "de": "Synthetische WIQ26-Generalprobe. Teresa: CT-1002 Manifest und CT-1005 Freigabe; Sonia: "
                          "CT-1003 Prüfung; Billie: CT-1008 schriftlicher Annahmeschluss. Vier Sperren bleiben offen. "
                          "Diese Einladung erstellt kein Transkript.",
                },
                "weekday": "Thursday", "start": "10:00", "minutes": 30,
            },
        },
        "manual": {
            "en": [
                "Admin publishes this shared demo once with the global -Tenant command; other presenters do not rerun setup or need admin roles.",
                "As each presenter, open the WIQ26 SharePoint link in a [WIQ26-M02] inbox message; confirm WIQ-D03 opens, the four messages arrived and the WIQ26 rehearsal invitation is in Outlook. Accept the invitation if needed.",
                "Use Microsoft 365 Copilot (Premium), verify Work IQ is available and on, and rehearse unassisted discovery. Check the site's/library's search visibility. Install ahead of the session; uploading is not proof of indexing.",
            ],
            "de": [
                "Admin veröffentlicht diese gemeinsame Demo einmal mit dem globalen -Tenant-Befehl; andere Vorführende führen das Setup nicht erneut aus und brauchen keine Adminrolle.",
                "Mit jedem Vorführkonto den WIQ26-SharePoint-Link in einer [WIQ26-M02]-Mail öffnen; WIQ-D03, Eingang der vier Mails und die WIQ26-Einladung in Outlook prüfen. Einladung bei Bedarf annehmen.",
                "Microsoft 365 Copilot (Premium) nutzen, verfügbares und eingeschaltetes Work IQ prüfen und die Suche ohne Anhänge proben. Suchbarkeit von Website/Bibliothek prüfen. Vor dem Termin installieren; Upload ist kein Indexierungsnachweis.",
            ],
        },
        "links": [{"name": {"en": "Microsoft Copilot Chat", "de": "Microsoft Copilot Chat"}, "url": "https://copilot.cloud.microsoft/chat"}],
    })
    write_json(FOLDER / "manifest.json", {
        "title": {"en": "Work IQ: shared Contoso port evidence", "de": "Work IQ: gemeinsame Contoso-Hafennachweise"},
        "next": {"en": ["Admin: global -Tenant setup. Presenter: open the shared link from the WIQ26 inbox message."],
                 "de": ["Admin: globales -Tenant-Setup. Vorführkonto: gemeinsamen Link aus der WIQ26-Mail öffnen."]},
        "files": [],
    })
    write_json(FOLDER / "tenant" / "expected.json", {
        "scenario": "WIQ26", "logicalDocuments": 24, "physicalDocuments": 48, "messagesPerLanguage": 4,
        "totalContainers": 115, "heldContainers": 43, "readyContainers": 72,
        "holds": {"CT-1002": [12, "Teresa Sac", "corrected manifest"],
                  "CT-1003": [6, "Sonia Rees", "authorized inspection outcome"],
                  "CT-1005": [14, "Teresa Sac", "release approval"],
                  "CT-1008": [11, "Billie Vester", "written carrier cut-off"]},
        "checks": ["CT-1005 packing list received; release pending", "CT-1008 slot provisional; confirmation missing",
                   "No delivery promise", "M02 and WIQ-D03 both cited and opened as presenter",
                   "Actual next calendar date, no fabricated transcript"],
    })
    print("WIQ26: 24 distinct documents x 2 languages; 4 mails x 2 languages. Run Update-Manifest.ps1 -All.")


if __name__ == "__main__":
    build()

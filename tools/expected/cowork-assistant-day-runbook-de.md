# CSMAD Presenter-Auswahl

# Assistenzalltag: hier starten

Demo Kit: Cowork Ende zu Ende > Assistenzalltag: Ein Tag mit Cowork. Die separate Seite Executive Office: Leadership-Event bleibt erhalten.

Globales Tenant-Setup verwenden, zuerst mit -WhatIf. Es installiert alle Demos; kein Installer pro Fall. Angemeldeten freigegebenen Demo-Admin und aktuelle Exchange-Lizenzen prüfen. Mails liegen im Posteingang dieses Admins mit [EAD26 nn-nn]; Termine im eigenen Kalender dieses Admins am 19.-23. Oktober 2026 unter [EAD26 key]. Dateien: Demo-Kit > Cowork > cowork-assistant-day.

Skill einmalig: skills/contoso-assistant-day.zip herunterladen. In Cowork Anpassen > Skills > Pfeil neben Hinzufügen > Skill hochladen wählen. ZIP auswählen, Nur du beibehalten. Existiert der Name schon, vorhandenen Skill prüfen statt nummeriertes Duplikat anzulegen. Neuen Verlauf starten und Skill in Quellen auswählen; Aktivierung rechts prüfen. UI-Übersetzungen und Rollout im eigenen Tenant prüfen.

Pro Fall über Anlagen hinzufügen > Clouddateien anfügen runs/de/01/Run_Context.md, inputs/de/<code>/Case.md und genau deren Dateien auswählen. Für CO-35 zusätzlich Action_Register.xlsx desselben Laufs als Cloud-Original auswählen. Installation wählt keine Quellen in Cowork aus. Genannte Mails/Kalender vor dem Start prüfen. Kurzen Auftrag aus dem Demo Kit verwenden, gespeicherte Dateien und Outlook-Entwürfe außerhalb des Chats öffnen.

Läufe 01-03 sind unabhängige saubere Startplätze. Nach Änderungen neuen Platz nutzen. Sind alle benutzt, zeigt Reset-CoworkAssistantDay.ps1 -Language de -WhatIf lokal einen nicht löschenden Fresh-Run-Reset; -WhatIf erst nach Prüfung weglassen. Bestehende autorisierte Graph-Sitzung wird wiederverwendet, keine Anmeldung gestartet. Originalbaseline wird heruntergeladen und per SHA-256 geprüft; eindeutiger Lauf wird neu angelegt. Mail-/Kalenderreset unnötig, weil Live-Aufträge nichts daran ändern dürfen. Falls doch: stoppen, nicht still neu installieren.

Installer aktualisiert Baselines/Fixtures und ist KEIN Reset; vorhandene Laufmappen und Kontexte bleiben auch nach Bearbeitung erhalten. Neue Reset-Läufe gehören nicht zum Manifest; ältere Ergebnisse bleiben erhalten. Skill beim Reset nicht erneut registrieren; nur bewusste Skill-Revisionen austauschen, keine Duplikate.

Alle acht Live-Abnahmen OFFEN. Keine neuen Cowork-Backups mitgeliefert. Voraussetzungen: aktive Copilot-Lizenz, Cowork freigeschaltet, Nutzungsabrechnung, erreichbare Clouddateien, Mailfindung und eigener Kalender. Keine Delegation, privaten Chef-Chats, Browser-Erweiterung oder Frontier-App-Skill nötig. Bei fehlenden Quellen oder Senden/Posten/Einladen/Verschieben/Automation-Freigaben stoppen. Presenter-Erwartungen außerhalb dieses Ordners unter tools/expected.

## Auswählen, nicht durch sechzehn Jobs hetzen

Pool: acht neue unabhängige Alltagsfälle plus acht unveränderte Event-Momente. Für den Demoanteil CO-30, CO-31, CO-34, CO-35 und CO-36 als Kern auswählen, CO-24 Präsentation/Änderung als kontrastierendes gespeichertes Ergebnis und CO-37 als Schlussgedanken. CO-32 und CO-33 sind Reserven, keine Pflichtzusatzschritte. Lange Jobs starten, bevor ein anderes gespeichertes Ergebnis erklärt wird. Vorschlag für 40 Minuten Redeanteil (keine Cowork-Laufzeit): 4 Minuten Einordnung, 6 Posteingang, 5 Morgen, 6 Kalender, 8 Meeting-to-Action, 6 Besuch oder Eventpaket-Kontrast, 5 Abschluss/Urteil. Keine Generierungswartezeiten als Inhalt einplanen. Nur eigene tatsächlich gespeicherte und geprüfte Backups zeigen. Ohne abgenommenes Backup den Live-Fall weglassen statt ein Ergebnis vorzutäuschen.

Unveränderter Event-Pool: CO-26 Antwort; CO-27 Kalender; CO-28 Gespräche; CO-29 Reise; CO-23 Executive-Briefing/Ortsvergleich; CO-24 Eventpaket; CO-24 Speaker-Änderung; CO-25 Nachbereitung. Separate Geschichte mit ihren drei Skills und privaten historischen Belegen.

## CO-30 — Posteingang sortieren

Im Executive Office liegen sechs unterschiedliche Nachrichten. Zwei betreffen dieselbe Anfrage; eine braucht keine Antwort. Entscheidungen vorbereiten, nicht treffen.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/de/CO-30/Case.md`, `inputs/de/CO-30/Inbox_Rules.docx`, `runs/de/<active>/Run_Context.md`.

Sortiere bitte diese Anfragen im Posteingang und bereite die zwei Antworten vor, um die ich mich kümmern muss. Speichere die Übersicht für meine Abstimmung mit Teresa.

Nur die sechs in Case.md genannten CO-30-Mails lesen. Kompakte Tabelle mit allen sechs IDs, fünf Kategorien (Entscheidung nötig, Terminwunsch, Information, delegierbar, keine Aktion), Begründung, nächstem Schritt und Quelle speichern. 30-05 als Duplikat von 30-01 markieren, nicht als zweite Freigabe. Genau zwei echte ungesendete Outlook-Entwürfe speichern: an Sydney (30-01), fehlende Begründung erfragen ohne 2.200 EUR freizugeben; an Teresa (30-02), Priorität bestehender Verpflichtungen erfragen ohne einen Termin zuzusagen. Keine Mail verschieben, markieren oder löschen.

Menschliches Urteil: Eskalation und Ton persönlich entscheiden.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## CO-31 — Morning Brief

Vor Tagesbeginn wurde eine Lieferantenentscheidung vorgezogen, IT meldet einen neuen Stand und im eigenen Kalender kollidieren Termine.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/de/CO-31/Case.md`, `inputs/de/CO-31/Morning_Papers.docx`, `inputs/de/CO-31/Calendar_Snapshot.docx`, `runs/de/<active>/Run_Context.md`.

Bereite bitte meinen Morning Brief für unseren Szenariotag vor. Bring Kalender, Teresas Update und die aktuellen Unterlagen zusammen und zeig mir, welche Entscheidungen zuerst anstehen.

Nur CO-31-Mails, Morning_Papers und die genannten eigenen EAD26-Kalendereinträge nutzen. Stand ist 19. Oktober 2026, 08:30 Europe/Berlin, unabhängig von der echten Uhrzeit. Tatsächlichen Kalender mit der ausgewählten Momentaufnahme abgleichen; Abweichungen benennen. Lesbaren einseitigen Word-Brief mit Top 3, Änderungen, Entscheidungsfristen, Quellenlinks, Lücken und Arbeitsreihenfolge speichern. 72.000 EUR / fünf Tage durch 88.000 EUR / acht Tage aus dem aktuellen Papier ersetzen. Teresas Entscheidungsfrist ist heute 09:30, nicht morgen. Billies Test ist laut synthetischem Beleg erledigt, nicht von dir geprüft. CO-35-Nachmittagsnotizen und CO-37-Abschlussbelege ausschließen.

Menschliches Urteil: Entscheiden, was zuerst persönlich zu Teresa gehört.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## CO-32 — Meeting Brief

Ein Lieferantenreview braucht eine Entscheidung. Das alte Dashboard sieht positiv aus; eine neuere Qualitätsnotiz verändert das Bild.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/de/CO-32/Case.md`, `inputs/de/CO-32/Review_Previous.docx`, `inputs/de/CO-32/Review_Current.docx`, `runs/de/<active>/Run_Context.md`.

Bereite bitte mein Briefing fürs Lieferantenreview vor. Vergleiche das alte Dashboard mit dem aktuellen Qualitätsstand und zeig mir die Fragen, die Teresa entscheiden muss.

Die zwei CO-32-Mails und beide ausgewählten Review-Papiere nutzen. Einseitigen Word-Brief mit Terminziel, Alt-/Neu-Vergleich, Entscheidungen, drei Talking Points, offenen Fragen und echten Quellenlinks speichern. Zeiträume trennen: September 96 Prozent pünktlich/sechs Fälle; Oktober bisher 88 Prozent/vierzehn Fälle. Sonias Trainingskapazität ist unbestätigt, sie ist keine freigegebene Verantwortliche. Lieferantenreview am 20. Oktober um 11:00, nicht das interne CO-35-Prozessreview. Kein Zugriff auf fremde persönliche Teams-Chats unterstellen.

Menschliches Urteil: Die Gesprächsposition bestimmen, nicht nur Zahlen übernehmen.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## CO-33 — Wiedervorlage

Drei Zusagen mit unterschiedlicher Beleglage: eine abgeschlossen, eine ohne Abschlussbeleg und eine mehrdeutige Rückmeldung.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/de/CO-33/Case.md`, `inputs/de/CO-33/Handover_List.docx`, `runs/de/<active>/Run_Context.md`.

Schau bitte nach, was aus diesen drei Zusagen geworden ist, und bereite das Nachfassen vor. Trenne unklare Fälle, damit ich niemanden unnötig erinnere.

Die fünf CO-33-Mails anhand ihrer ausdrücklichen synthetischen Zeitstempel lesen, nicht nach Zustellzeit. Drei Zeilen mit Belegen speichern: Vances Übergabeliste durch 33-02 und Handover_List als geliefert bestätigt; für Sonias Gesprächsauswertung fehlt ein Abschlussbeleg; Billies 'sieht gut aus' bestätigt keinen abgeschlossenen Barrierefreiheitstest. Fehlende Belege sind kein Beweis für Versäumnis. Zwei ungesendete Entwürfe: Sonia freundlich nach Auswertung/Stand fragen, Billie um Testbeleg bitten. Vance nicht erinnern, keine Zusage als erledigt markieren und keine Frist erfinden.

Menschliches Urteil: Entscheiden, wen man lieber persönlich anspricht.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## CO-34 — Kalenderwoche

Die eigene Woche enthält eine Kollision, einen Ortswechsel ohne Wegezeit und einen Workshop ohne Agenda. Fremde Verfügbarkeit ist unbekannt.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T08:30:00+02:00

Sources: `inputs/de/CO-34/Case.md`, `inputs/de/CO-34/Calendar_Snapshot.docx`, `inputs/de/CO-34/Calendar_Rules.docx`, `runs/de/<active>/Run_Context.md`.

Prüfe bitte meine Szenariowoche und bereite sinnvolle Optionen für die Kalenderprobleme vor. Entwirf die Abstimmung, ändere aber noch keine Termine.

Nur EAD26-Termine vom 19. bis 23. Oktober im eigenen Presenter-Kalender prüfen und mit der ausgewählten Momentaufnahme vergleichen. Entscheidungsvorlage mit echter Überschneidung, fehlender Wege-/Vorbereitungszeit und fehlender Agenda speichern. Montag Standortreview 10:00-10:45 kollidiert mit Budget 10:30-11:00. Besuch in Gebäude B um 11:00 folgt auf Budget in Gebäude A ohne nötige 30 Minuten Weg. Montagsworkshop um 13:00 ohne Agenda. Lieferantenreview Dienstag 11:00 folgt auf geschützten Fokusblock bis 11:00 ohne 15 Minuten Vorbereitung. Alternativen nur vorschlagen, keine fremde Verfügbarkeit behaupten. Zwei ungesendete Entwürfe: Sydney zur Budgetkollision, Sonia zur Workshopagenda. Kalender-IDs, Zeiten und Teilnehmende unverändert lassen.

Menschliches Urteil: Politisch und persönlich wichtige Termine schützen.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## CO-35 — Vom Meeting zur Umsetzung

Ein unabhängiges internes Prozessreview hat freigegebene Notizen, eine vorhandene Aufgabenzeile und neue Aufgaben mit einer fehlenden Frist.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T14:00:00+02:00

Sources: `inputs/de/CO-35/Case.md`, `inputs/de/CO-35/Released_Notes.docx`, `runs/de/<active>/Action_Register.xlsx`, `runs/de/<active>/Run_Context.md`.

Mach bitte aus diesen freigegebenen Review-Notizen die Beschlüsse und nächsten Schritte. Aktualisiere unser Aufgabenregister und bereite das individuelle Nachfassen vor.

Nur ausgewählte freigegebene CO-35-Notizen und Originalmappe des aktiven Laufs nutzen. Decisions.docx speichern und genau diese Cloud-Mappe am bestehenden Ort bearbeiten. Blätter, Spaltenreihenfolge, Werte/Typen/Formate von P01, Readme und Tabellennamen erhalten. Ergänzen: A01 Sydney Mattos, Übergabeausnahmen abgleichen, 2026-10-22; A02 Billie Vester, Barrierefreiheitstest belegen, Frist nicht beschlossen (leer); A03 Sonia Rees, Einarbeitungscheckliste entwerfen, 2026-10-23. Actions-Tabelle auf A1:F5 erweitern. D01 zweiwöchiger Pilot, D02 kein Softwarekauf; keine Sendefreigabe. Drei individuelle ungesendete Entwürfe an die Verantwortlichen speichern. Billie nach Frist fragen statt Zusage zu erfinden. Keine Baseline überschreiben oder Ersatzmappe anlegen.

Menschliches Urteil: Fehlende Frist klären und das schwierige Gespräch führen.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## CO-36 — Kurzfristiger Executive-Besuch

Teresa plant kurzfristig einen internen Standortbesuch am Donnerstag: enges Zeitfenster, Sicherheitsvorgabe und unbestätigte IT-Verfügbarkeit.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T14:30:00+02:00

Sources: `inputs/de/CO-36/Case.md`, `inputs/de/CO-36/Visit_Brief.docx`, `runs/de/<active>/Run_Context.md`.

Bereite bitte Teresas Besuch am Donnerstag mit diesen Nachrichten und Standortnotizen vor. Speichere ein stimmiges Besuchspaket und entwirf die Abstimmung für meine Prüfung.

Nur CO-36-Belege nutzen. Drei passende gespeicherte Dokumente erstellen: einseitiger Ablauf Donnerstag 22. Oktober 10:00-12:00; einseitiges Teilnehmerbriefing aus synthetischen Zuständigkeiten; Entscheidungsliste mit festen Vorgaben und offenen Punkten. Sonias verpflichtende Sicherheitseinweisung vor dem 30-minütigen Rundgang, zehnminütige Übergänge und mindestens 20 Minuten persönliches Gespräch einplanen. Teresa geht um 12:00; Billie erst ab 11:15 verfügbar, seine Technikdemo ist unbestätigt. Keine Biografien, Raumbuchung oder Zusage erfinden. Drei ungesendete Entwürfe: Teresa Agendavorschlag, Sonia Sicherheits-/Logistikfragen, Billie Machbarkeit der Demo ab 11:15. Keine Einladungen oder Kalendereingriffe.

Menschliches Urteil: Eigentliches Besuchsziel entscheiden und Gastgeber persönlich einbinden.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## CO-37 — Tagesabschluss

Ein ausdrücklich synthetischer Tagesstand belegt eine erledigte Entscheidung, einen ungesendeten Entwurf und einen offenen Test. Kein Protokoll vorheriger Cowork-Läufe.

Skill: `contoso-assistant-day` | Snapshot: 2026-10-19T17:00:00+02:00

Sources: `inputs/de/CO-37/Case.md`, `inputs/de/CO-37/Day_Checkpoint.docx`, `inputs/de/CO-37/Tomorrow_Calendar.docx`, `runs/de/<active>/Run_Context.md`.

Bereite bitte meinen Tagesabschluss aus diesem Zwischenstand vor und fang den Morning Brief für morgen an. Trenne Erledigtes, offene Freigaben und alles, was die Belege nicht klären.

Eigenständiger Start: nur synthetischen CO-37-Checkpoint und morgigen Kalender mit Stand 19. Oktober 17:00 nutzen. Das sind vorbereitete Quellenbelege, keine Cowork-Ergebnisse. Einseitigen Abschluss und einseitigen Briefentwurf für 20. Oktober speichern. C01 laut Quelle Kauf von Teresa abgelehnt; C02 Entwurf vorbereitet, NICHT gesendet/freigegeben; C03 Abschluss Barrierefreiheitstest unbekannt. Review 11:00 braucht 15 Minuten Vorbereitung; Finanzabgleich bleibt zum 22. Oktober fällig. Drei vorgeschlagene Prioritäten und Quellenlinks nennen; keine tatsächliche Cowork-Erledigung behaupten. Optionaler echter Folgepfad nur bei ausdrücklich ausgewählten tatsächlichen Ergebnissen UND separaten menschlichen Freigabebelegen: getrennt kennzeichnen, synthetische/echte Zustände nicht mischen und Versand nicht aus Entwurf ableiten. Keine Automation.

Menschliches Urteil: Prioritäten für morgen setzen und Zeit für Beziehungen schützen.

Freigabegrenze: Dateien/Entwürfe anlegen; nur CO-35 bearbeitet seine Laufmappe. Nur einzelne Freigaben. Bei fehlenden Quellen, unerwarteten Empfängern, Senden/Posten/Einladen/Kalender/Verschieben/Automation oder widersprüchlichen Dateien stoppen. Grenze festhalten und zu abgenommenem Ergebnis wechseln; kein Reparaturmarathon.

## Deine Live-Abnahme — alles OFFEN

| Code | Check outside chat / Außerhalb des Chats prüfen | Status |
|---|---|---|

| CO-30 | Sechs Zeilen; fünf Kategorien; Duplikat zugeordnet; zwei gespeicherte ungesendete Entwürfe; keine Freigabe von 2.200 EUR. | OFFEN |

| CO-31 | Eine Seite; Frist 09:30; 88.000 EUR/acht Tage; Kalenderkollision; keine Nachmittagsinformationen. | OFFEN |

| CO-32 | Eine Seite; getrennte Zeiträume; 96 auf 88 Prozent, sechs auf vierzehn; Trainingsentscheidung offen. | OFFEN |

| CO-33 | Abgeschlossen / kein Abschlussbeleg / unklar getrennt; keine Erinnerung an Vance; zwei ungesendete Entwürfe. | OFFEN |

| CO-34 | Vier Problemarten; zwei echte Entwürfe; alle Originaltermine unverändert. | OFFEN |

| CO-35 | Identität der Originalmappe; P01/Readme exakt erhalten; A01-A03, A02 ohne Datum; drei Entwürfe. | OFFEN |

| CO-36 | Drei stimmige gespeicherte Dokumente; Sicherheit vor Rundgang; 10:00-12:00, 20 Gesprächsminuten; drei Entwürfe. | OFFEN |

| CO-37 | Zwei gespeicherte Einseiter; abgelehnt / nur Entwurf / unklar getrennt; kein behaupteter Versand. | OFFEN |

Nach deiner Probe tatsächliche Verläufe, Cloud-Dateiidentitäten, Entwurf-IDs, Screenshots und heruntergeladene abgenommene Ergebnisse privat sichern. Gespeicherte Dokumente in ihren Apps öffnen, nicht nur Chatvorschau. Keine Agenten-UI-Probe durchgeführt. Drei Zustände trennen: dokumentierte Funktion, lokal geprüfte Eingaben, vom Presenter abgenommenes Live-Ergebnis. Synthetisches Datum auf der Bühne nennen; 'heute' behauptet keine aktuellen Termine.

## Official references

- https://learn.microsoft.com/microsoft-365/copilot/cowork/
- https://learn.microsoft.com/microsoft-365/copilot/cowork/use-cowork
- https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-faq
- https://learn.microsoft.com/microsoft-365/copilot/cowork/get-started
- https://learn.microsoft.com/microsoft-365/copilot/cowork/cowork-customize
- https://learn.microsoft.com/training/modules/empower-workforce-copilot-executives/
- https://support.microsoft.com/en-us/microsoft-365-copilot/frequently-asked-questions-about-retired-copilot-features

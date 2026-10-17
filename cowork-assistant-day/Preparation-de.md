# Assistenzalltag: hier starten

Demo Kit: Cowork Ende zu Ende > Assistenzalltag: Ein Tag mit Cowork. Die separate Seite Executive Office: Leadership-Event bleibt erhalten.

Globales Tenant-Setup verwenden, zuerst mit -WhatIf. Es installiert alle Demos; kein Installer pro Fall. Angemeldeten freigegebenen Demo-Admin und aktuelle Exchange-Lizenzen prüfen. Mails liegen im Posteingang dieses Admins mit [EAD26 nn-nn]; Termine im eigenen Kalender dieses Admins am 19.-23. Oktober 2026 unter [EAD26 key]. Dateien: Demo-Kit > Cowork > cowork-assistant-day.

Skill einmalig: skills/contoso-assistant-day.zip herunterladen. In Cowork Anpassen > Skills > Pfeil neben Hinzufügen > Skill hochladen wählen. ZIP auswählen, Nur du beibehalten. Existiert der Name schon, vorhandenen Skill prüfen statt nummeriertes Duplikat anzulegen. Neuen Verlauf starten und Skill in Quellen auswählen; Aktivierung rechts prüfen. UI-Übersetzungen und Rollout im eigenen Tenant prüfen.

Pro Fall über Anlagen hinzufügen > Clouddateien anfügen runs/de/01/Run_Context.md, inputs/de/<code>/Case.md und genau deren Dateien auswählen. Für CO-35 zusätzlich Action_Register.xlsx desselben Laufs als Cloud-Original auswählen. Installation wählt keine Quellen in Cowork aus. Genannte Mails/Kalender vor dem Start prüfen. Kurzen Auftrag aus dem Demo Kit verwenden, gespeicherte Dateien und Outlook-Entwürfe außerhalb des Chats öffnen.

Läufe 01-03 sind unabhängige saubere Startplätze. Nach Änderungen neuen Platz nutzen. Sind alle benutzt, zeigt Reset-CoworkAssistantDay.ps1 -Language de -WhatIf lokal einen nicht löschenden Fresh-Run-Reset; -WhatIf erst nach Prüfung weglassen. Bestehende autorisierte Graph-Sitzung wird wiederverwendet, keine Anmeldung gestartet. Originalbaseline wird heruntergeladen und per SHA-256 geprüft; eindeutiger Lauf wird neu angelegt. Mail-/Kalenderreset unnötig, weil Live-Aufträge nichts daran ändern dürfen. Falls doch: stoppen, nicht still neu installieren.

Installer aktualisiert Baselines/Fixtures und ist KEIN Reset; vorhandene Laufmappen und Kontexte bleiben auch nach Bearbeitung erhalten. Neue Reset-Läufe gehören nicht zum Manifest; ältere Ergebnisse bleiben erhalten. Skill beim Reset nicht erneut registrieren; nur bewusste Skill-Revisionen austauschen, keine Duplikate.

Alle acht Live-Abnahmen OFFEN. Keine neuen Cowork-Backups mitgeliefert. Voraussetzungen: aktive Copilot-Lizenz, Cowork freigeschaltet, Nutzungsabrechnung, erreichbare Clouddateien, Mailfindung und eigener Kalender. Keine Delegation, privaten Chef-Chats, Browser-Erweiterung oder Frontier-App-Skill nötig. Bei fehlenden Quellen oder Senden/Posten/Einladen/Verschieben/Automation-Freigaben stoppen. Presenter-Erwartungen außerhalb dieses Ordners unter tools/expected.

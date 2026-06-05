# SFK Buchhaltung

Lokale Web-App zum Verarbeiten von Ausgabenbelegen fuer die Vereinsbuchhaltung.

## Start

```powershell
& "C:\Users\Lars\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\buchhaltung_app.py
```

Danach im Browser oeffnen:

```text
http://127.0.0.1:8501
```

## Ablauf

1. Jahr auswaehlen.
2. Pfad zur Buchhaltungs-Excel pruefen. Die App merkt sich diesen Pfad in `config.json`.
3. Ueberweisungsbeleg hochladen.
4. Rechnung oder Nachweis hochladen.
5. Vorschlag pruefen und korrigieren.
6. Beleg erzeugen.

Wenn die Checkbox aktiv ist, aktualisiert die App die konfigurierte Excel-Datei direkt. Vorher wird eine Backup-Kopie im Ordner `backups` angelegt. Zusaetzlich erzeugt sie immer eine aktualisierte `.xlsx` im Ordner `outputs`.

## Aktueller Umfang

- Ausgaben mit Belegnummern wie `A-2026-41`
- Jahresblaetter wie `Abrechnung 2026`
- PDF-Zusammenfuehrung in der Reihenfolge Ueberweisung, danach Nachweis
- Rote, eingekreiste Belegnummer nur auf den Seiten des Ueberweisungsbelegs
- Eintrag in die Ausgangsspalten der Excel

Eingaenge (`E-2026-xx`) sind noch nicht automatisiert.

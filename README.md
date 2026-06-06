# SFK Buchhaltung

Lokale Web-App zum Verarbeiten von Ausgabenbelegen fuer die Vereinsbuchhaltung.

## Installation auf einem neuen Rechner

### 1. Voraussetzungen installieren

- Python 3.12 installieren: https://www.python.org/downloads/
- Beim Python-Setup die Option `Add python.exe to PATH` aktivieren.
- Optional Git installieren: https://git-scm.com/download/win

Python 3.12 ist empfohlen, weil die App aktuell noch das Python-Modul `cgi` nutzt. Python 3.13 kann damit Probleme machen.

### 2. Projekt herunterladen

Mit Git:

```powershell
git clone https://github.com/lala1408/buchhaltung-tool.git
cd buchhaltung-tool
```

Ohne Git:

1. Auf GitHub `Code` -> `Download ZIP` auswaehlen.
2. ZIP entpacken.
3. PowerShell im entpackten Ordner oeffnen.

### 3. Virtuelle Python-Umgebung erstellen

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Falls PowerShell das Aktivieren blockiert:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
.\.venv\Scripts\Activate.ps1
```

### 4. App starten

```powershell
python .\buchhaltung_app.py
```

Danach im Browser oeffnen:

```text
http://127.0.0.1:8501
```

### 4a. Komfort-Start mit Fenster

Alternativ startet der Launcher die App und oeffnet den Browser automatisch:

```powershell
python .\buchhaltung_launcher.py
```

### 5. Erste Einrichtung

1. Im Feld `Buchhaltungs-Excel` den Pfad zur lokalen Excel-Datei eintragen.
   Beispiel: `C:\Users\Lars\Downloads\SFK-Buchhaltung ab 2022.xlsx`
2. Die App speichert diesen Pfad in `config.json` und nutzt ihn nach einem Neustart wieder.
3. `config.json`, `backups`, `outputs` und `work` werden nicht zu GitHub hochgeladen.

## Windows-EXE bauen

Auf einem eingerichteten Rechner:

```powershell
.\build_exe.ps1
```

Danach liegt die Doppelklick-Anwendung hier:

```text
dist\SFK-Buchhaltung.exe
```

Die `.exe` startet einen lokalen Server, oeffnet den Browser automatisch und zeigt ein Hinweisfenster. Dieses Fenster offen lassen, solange du arbeitest; `OK` beendet die App.

Wenn die `.exe` in einen anderen Ordner kopiert wird, legt sie `config.json`, `backups`, `outputs` und `work` neben der `.exe` an.

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
- Datum aus dem Ueberweisungsbeleg: zuerst Wertstellung/Valuta, sonst Buchungstag

Bei bildbasierten oder eingescannten Bank-PDFs kann das Datum eventuell nicht automatisch gelesen werden. Dann bleibt das Datumsfeld in der Vorschau leer und muss manuell gesetzt werden.

Eingaenge (`E-2026-xx`) sind noch nicht automatisiert.

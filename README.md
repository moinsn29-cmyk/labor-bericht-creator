# Laborbericht-Generator (LBS 4 Salzburg, IT-Labor)

Aus deinen Screenshots, der Aufgabenstellung und der Word-Vorlage wird ein fertiger Laborbericht als Word- und PDF-Datei. Ein KI-Agent schreibt den Text, das Programm baut den Bericht in der Schulvorlage, und zum Schluss liest ein zweiter Agent den Bericht so durch, wie es deine Lehrperson tun würde.

## Was du brauchst

- **LibreOffice** (auf den Schulrechnern installiert, sonst gratis unter libreoffice.org). Läuft unter Windows, Linux und macOS.
- **Claude Code** (oder einen anderen KI-Coding-Agenten, der Dateien lesen und Befehle ausführen kann)
- Python musst du **nicht** installieren: Auf Rechnern ohne Python (Schulrechner) läuft das Programm mit dem Python, das bei LibreOffice dabei ist; alles andere liegt in `app/libs`. Hast du ein eigenes Python, wird das verwendet (siehe „Zwei Betriebsarten").
- **Deine eigenen Dateien:** die Aufgabenstellung (PDF), die Word-Vorlage deiner Schule (DOCX) und deine Screenshots. Das Repository enthält keine Vorlage und keine Beispielberichte.

## So gehts

1. Repository klonen (`git clone https://github.com/moinsn29-cmyk/labor-bericht-creator.git`) oder als ZIP herunterladen und entpacken.
2. In diesem Ordner Claude Code starten (Terminal im Ordner öffnen, `claude` eingeben).
3. Sagen, was du willst, zum Beispiel:

   > Hier ist meine Aufgabenstellung `C:\…\Uebung3.pdf`, meine Vorlage `C:\…\Vorlage.docx` und meine Screenshots in `C:\…\Bilder`. Mach den Laborbericht.

   Oder leg die Dateien in den Ordner `eingabe` und sag nur „Mach den Laborbericht aus eingabe".

4. Der Agent fragt dich beim ersten Mal nach Name, Klasse, Katalog-Nr., Jahrgang und Ort und merkt sich das in `profil.json`. Dann fragt er nach allem, was man auf den Screenshots nicht sehen kann (zum Beispiel welches Image du verwendet oder ob du neu gestartet hast). Antworte ehrlich: Im Bericht steht nur, was belegt ist oder was du bestätigt hast.
5. Fertig ist der Bericht in `Dokumente\<Titel>_<Datum>\final\` als `.docx` und `.pdf`.

## Wichtig

- **Screenshots mit dem Snipping Tool machen und nicht umbenennen.** Aus dem Dateinamen (`Screenshot 2026-11-23 131924.png`) wird die Uhrzeit gelesen. Jede Überschrift eines Arbeitsschritts bekommt die Uhrzeit des ersten Screenshots dieses Schritts. **Mach zu Beginn jedes Arbeitsschritts einen Screenshot**, sonst hat der Schritt keine Uhrzeit. Uhrzeiten werden nie erfunden.
- **Lies den Bericht vor der Abgabe.** Er wird unter deinem Namen abgegeben, die Reflexion ist in Ich-Form geschrieben, und die Eidesstattliche Erklärung bestätigt, dass du alle Hilfsmittel angegeben hast. Der Agent fragt dich, ob der KI-Assistent unter „Hilfsmittel" stehen soll. Kläre mit deiner Lehrperson, was erlaubt ist.
- Ist eine Datei in `final\` noch in Word oder im PDF-Viewer geöffnet, kann sie nicht ersetzt werden. Schließen und den Agenten weitermachen lassen.
- Die Vorlage muss die Schulvorlage `Vorlage_Laborbericht_PG_V2` sein (oder ihr sehr ähnlich). Bei einer anderen Vorlage sagt dir der Agent, was fehlt.

## Zwei Betriebsarten

Der Agent erkennt selbst, welche passt. Der Bericht sieht in beiden gleich aus.

- **Normal** (eigener Rechner mit Python 3, auch Linux und macOS): `python3 app/check_env.py --setup` legt einmalig den Ordner `.venv` an und installiert dort die Pakete aus `requirements.txt`. Systemweit wird nichts installiert. Unter Debian/Ubuntu braucht LibreOffice zusätzlich das Paket `python3-uno`.
- **Gesperrt** (Schulrechner ohne Python, pip und Adminrechte): Es läuft alles mit dem Python von LibreOffice und den mitgelieferten Bibliotheken in `app/libs`. Es wird nichts installiert.

## MCP-Server (optional)

`app\mcp_server.py` stellt den Generator als MCP-Werkzeuge bereit (Bericht bauen, prüfen, Dateien auflisten). Claude Code bietet ihn über `.mcp.json` beim Start im Ordner an. Für den normalen Ablauf über den Agenten ist er nicht nötig. `.mcp.json` ist für die Schulrechner eingestellt (LibreOffice-Python unter Windows); auf einem eigenen Rechner `requirements-mcp.txt` in `.venv` installieren und in `.mcp.json` das Python aus `.venv` eintragen.

## Lizenz

MIT, siehe `LICENSE`. Gilt für den Code und die Anleitungen in diesem Repository, nicht für Vorlagen oder Unterlagen deiner Schule.

## Ordner eines Berichts

```
Dokumente\<Titel>_<Datum>\
  screenshots\   deine Screenshots (Kopien)
  work\          Aufgabenstellung, Vorlage, content.json (der Text), report.json (Prüfprotokoll)
  final\         <Abgabename>.docx und <Abgabename>.pdf
```

Willst du etwas ändern, sag es dem Agenten („In Abschnitt 3.4 hab ich eigentlich …"). Er ändert `content.json` und baut den Bericht neu. Die Dateien in `final\` nicht von Hand bearbeiten, der nächste Lauf überschreibt sie.

## Ohne Agent neu bauen

```powershell
& "C:\Program Files\LibreOffice\program\python.exe" app\headless_core.py "C:\…\Dokumente\<Titel>_<Datum>"
```

Mit eigenem Python (normale Betriebsart):

```
.venv/bin/python app/headless_core.py "<Berichtsordner>"
```

Unter Windows heißt das Python `.venv\Scripts\python.exe`.

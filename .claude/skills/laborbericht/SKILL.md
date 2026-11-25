---
name: laborbericht
description: Write a complete, hand-in-ready Laborbericht (lab report, LBS 4 Salzburg IT-Labor) from the student's screenshots, the Aufgabenstellung PDF and the Word Vorlage. Use when the user hands over these files and says things like "write the report", "mach den Laborbericht", "schreib den Bericht", or asks to rebuild, fix or review an existing Laborbericht.
argument-hint: "[where the screenshots, the Aufgabenstellung and the Vorlage are]"
---

# Laborbericht

Turn screenshots, an Aufgabenstellung PDF and the Word Vorlage into a finished report: a DOCX and a PDF built from the template, checked by the generator and reviewed by the `laborbericht-lehrer` agent in the role of the teacher. The student wants to hand the PDF in without touching it, so the job is finished only when the teacher review says `ABGABEBEREIT`.

All paths below are relative to the bundle folder (the folder that contains this `.claude` folder and `app/`).

## 0. Check the environment (once per session)

The generator runs in two modes and the report comes out the same in both. Find out which one applies; do not ask the user.

**Normal mode** — the machine has its own Python 3 (`python3 --version` or `python --version` works; any Linux or macOS machine, a private Windows PC):

```
python3 app/check_env.py
```

It prints `MODE=normal`, `PYTHON=…` (the interpreter to build with, normally the bundle's `.venv`), `DOCUMENTS=…`, `SOFFICE=…` and `UNO_PYTHON=…`. Remember `PYTHON` and `DOCUMENTS`.

- `IMPORTS=FEHLER`: the packages are not installed yet. Run `python3 app/check_env.py --setup` once; it creates `.venv` in the bundle folder and installs `requirements.txt` there. Nothing is installed system-wide.
- `SOFFICE_FOUND=False`: LibreOffice is required in both modes (it refreshes the indexes and writes the PDF). Tell the user to install it (or set `LIBREOFFICE_PROGRAM`) and stop.
- `UNO_PYTHON=NICHT GEFUNDEN`: LibreOffice's Python bridge is missing. On Debian/Ubuntu the user installs the package `python3-uno`; otherwise set `LIBREOFFICE_PYTHON` to LibreOffice's own Python.
- If the setup cannot work (no internet, pip blocked, no rights), use the locked-down mode.

**Locked-down mode** — no Python, no pip, no admin rights (the school machines):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File app\check_env.ps1
```

It prints `LIBREOFFICE_PYTHON=…` (the interpreter to build with), `MODE=locked-down`, `DOCUMENTS=…` (which may be a OneDrive folder) and `IMPORTS=OK`. Remember both paths.

- `LIBREOFFICE_PYTHON=NICHT GEFUNDEN`: LibreOffice is required. Tell the user to install it (or set `LIBREOFFICE_PROGRAM`) and stop.
- If PowerShell scripts are blocked, run the same test directly: `& "C:\Program Files\LibreOffice\program\python.exe" -c "import sys; sys.path.insert(0, 'app'); import headless_core as h; print(h.documents_folder())"`. School machines run PowerShell in Constrained Language Mode, so never use `[Environment]::…` or other .NET calls in PowerShell.
- In this mode only LibreOffice's own Python works: it brings the UNO bridge for the PDF export, and `app/libs` provides python-docx, pypdf and Pillow. Do not install anything, do not use pip, do not use another Python.

Below, `<PYTHON>` is the interpreter the check printed (`PYTHON=` in normal mode, `LIBREOFFICE_PYTHON=` in locked-down mode). Paths are written the Windows way; on Linux and macOS use `/`.

## 1. Gather what you need

### Student profile

Read `profil.json` in the bundle folder. If it does not exist or a field is empty, ask the user for the missing fields in **one** message, then write the file:

```json
{
  "name": "Vorname Nachname",
  "klasse": "2IB",
  "katalog_nr": "2IB",
  "jahrgang": "25/26 LG-B",
  "ort": "Salzburg",
  "ki_hinweis": true
}
```

- `klasse` is used in the file name the Aufgabenstellung prescribes. `katalog_nr` goes on the title page after "Katalog - Nr." (some students write their class there; take what they say).
- `jahrgang` is school year plus Lehrgang as on the title page, for example `25/26 LG-B`. If it does not fit the lab date (a school year runs September to August), ask once whether it is right, then accept the answer.
- `ki_hinweis`: ask whether the use of this AI assistant should be listed under "Quellen und verwendete Hilfsmittel". Explain in one sentence why it matters: the template's Eidesstattliche Erklärung confirms that all aids are listed. Recommend yes, but it is the student's decision.

Values in `profil.json` are confirmed by the student. Do not question them again; ask only if the user says something changed.

### Input files

The user names the files or puts them into `eingabe/`. You need:

- the **Aufgabenstellung** (PDF),
- the **screenshots** (PNG or JPG),
- the **Vorlage** (DOCX), the Word template the school requires. This repository ships no template and no example report: the user always provides them.
- **extra information** the user wants in the report (what went wrong, what they did differently, anything the screenshots do not show).

Ask for everything that is missing, profile fields included, in **one** first message, so the user can answer once and then wait for the finished report.

Screenshot names must contain date and time as the Windows Snipping Tool writes them: `Screenshot 2026-11-23 131924.png`. The times in the report come only from these names. If the names have no time, tell the user. Offer to rename the files from their modification time only after showing the list of old and new names and getting a yes; otherwise the steps get no times.

### Report folder

- Lab date: from the screenshot names. If they span several days, ask which day counts.
- Folder: `<DOCUMENTS>\<Titel>_<YYYY-MM-DD>`, with the exercise title from the sheet, spaces replaced by `_`, and characters invalid in Windows file names removed (on every system, so the folder can be moved to a school machine). Example: `…\Documents\Raspberry_Pi_einrichten_2026-11-23`.
- If the folder already exists, ask whether to continue with it or to start over.
- Create `screenshots\`, `work\` and `final\`. **Copy** (never move) the screenshots into `screenshots\` without renaming them, and the Aufgabenstellung and the Vorlage into `work\`.

```
<Titel>_<Datum>\
  screenshots\   the screenshots, unchanged
  work\          Aufgabenstellung, Vorlage, content.json, report.json, README.md, build files
  final\         the finished DOCX and PDF, nothing else
```

## 2. Read the Aufgabenstellung completely

Read the PDF with the Read tool **without** the `pages` parameter (page ranges fail when pdftoppm is missing, as on the school machines). Write down, before anything else:

- every numbered step of the Übungsanweisung, in order,
- every explicit question or documentation duty ("Dokumentieren Sie", "Notieren Sie", "Aufgabe:", "Erklären Sie"),
- the Funktionstest or acceptance list,
- the Erweiterung tasks,
- the required file name in the Abgabe section and any other hand-in rule.

Each of these must be visibly answered in the report. This list is your checklist in step 7.

## 3. Look at every screenshot

Open every image with the Read tool and note what it proves: commands, prompts, numbers, warnings, host names, IP addresses, clocks visible in the image. Also note defects: lines cut off at an edge, large empty areas, near-duplicates.

Assign each screenshot to the step it documents. A screenshot that adds nothing (for example a cut-off duplicate) goes into `unused_screenshots` with a reason. Every image must be either a figure or listed there, otherwise the build stops.

## 4. Ask about what the screenshots cannot show

Before writing, collect everything the report would otherwise have to guess and ask the user in **one** message, as short concrete questions. Typical gaps:

- which operating system image, version or tool was really used (for example the Lite image instead of the desktop one, the preinstalled `vi` instead of installing `vim`),
- options used that explain the output (for example `ssh -vv` when the screenshot shows debug lines),
- steps without a screenshot: were they done, and how (reboot after an update, packages installed, which of two suggested commands was used),
- anything in a screenshot that contradicts the sheet (default password, default host name, errors).

Write only what the user confirms. If they do not know, describe the step neutrally as the sheet prescribes, without claiming details.

## 5. Write `work/content.json`

Start from `app/beispiel_content.json` (structure only). Rules for the content:

- **Only facts you can back up.** Numbers, names and outputs come from the screenshots; everything else from the user's answers. Never invent output, versions, times or problems.
- **Deviations are stated where they happen.** If a requirement of the sheet was not met (a warning, a default password, another OS image, another editor), say so in the step concerned, not only in the Reflexion, and say what should be done about it. A deviation that is disclosed is fine; one that is glossed over costs marks.
- **Timestamps and order.** Every step heading (level 2) shows its start time. The generator takes it automatically from the earliest screenshot in that step and appends "(13:06 Uhr)"; never write a time into a heading yourself. A step without a screenshot has no time; never estimate one. Within each chapter, order steps and figures the way the screenshots were taken (the build checks this to the second and stops otherwise). If that differs from the sheet's order, say so in one sentence at the start of the chapter. A screenshot taken later as a re-run of an earlier step gets `"heading_time": false` and one explaining sentence. The introduction says that heading times are the times of the first screenshot and that steps without a screenshot have none.
- **The report must agree with its own figures.** A command shown as code must be able to produce the output in the figure next to it. A clock visible in a screenshot that differs from the caption time needs one explaining sentence.
- **Do not claim more evidence than there is.** In the Funktionstest, a row without a screenshot says so.
- **Reproducible.** A classmate must be able to repeat the exercise from the report alone: include installs, directory creation and `cd` commands.
- **Own words, level of a second-year Berufsschule student.** Short, clear German sentences in Präteritum passive ("wurde installiert") for the protocol. No copying of sentences from the Aufgabenstellung.
- **Sections in this order:** Einleitung und Zusammenfassung, Inventarliste (Hardware, Software), Übungsdurchführung (Schrittprotokoll) with one level-2 heading per step, Erweiterung / Zusatzaufgabe, Einsatzgebiet, Persönliche Reflexion (Erkenntnisse), Quellen und verwendete Hilfsmittel. Do not add the Eidesstattliche Erklärung: the generator appends the template's own text with place and date. The template's signature block is left out by design.
- **Commands as code**, each with its prompt (`pi@raspberrypi:~ $ …`, `C:\Users\…>…`); short outputs (up to about three lines) as code too. Code lines under 75 characters.
- **Answer the sheet's questions explicitly**, findable at a glance ("Vergebener Hostname: …").
- **Funktionstest as a table** with the columns Prüfung, Ergebnis, Nachweis.
- **Reflexion in first person**, based on what actually happened.
- **Quellen:** the Aufgabenstellung, the Vorlage, official documentation of the tools used (only URLs you are sure exist). If `ki_hinweis` is true, add "KI-Assistent (Claude) zur Formulierung und Gestaltung des Berichts".

Fields:

- `title`: exercise title from the sheet. `exercise_number`: from the sheet or its file name (`01_…` means 1); ask if unclear.
- `output_basename`: the file name the Abgabe section prescribes, with the placeholders replaced (for example `Klasse_IhrNachname_ILA_KPT_RPIeinrichten.pdf` becomes `2IB_Mustermann_ILA_KPT_RPIeinrichten`). Without a rule: `<klasse>_<Nachname>_<Titel>`.
- `template`: file name of the Vorlage in `work\`.
- `metadata`: `author` = name, `catalog` = katalog_nr, `school_year` = jahrgang, `place` = ort, `date` = lab date, `signing_date` = today, both `TT.MM.JJJJ`.
- `{abb:<id>}` becomes "Abbildung N". Always reference figures this way.
- Captions describe the figure without a time (the generator appends it). Keep them under about 75 characters including the time.
- `crop` is `[left, top, right, bottom]` in pixels: remove half-cut lines and large empty areas, never hide content; say "Ausschnitt" in the caption when much is cut away. Screenshots are scaled at 50 px per cm and capped at the 16 cm text width, so anything wider than about 1200 px must be cropped to the part the text talks about. One screenshot may appear as two figures with different crops.
- `max_height_cm` defaults to 11; raise it (up to 15) for tall screenshots.
- Table `widths_cm` must add up to 16.0.

## 6. Build

```powershell
& "<PYTHON>" app\headless_core.py "<report folder>"
```

(Linux/macOS: `"<PYTHON>" app/headless_core.py "<report folder>"`.)

It prints `passed`, `pages`, `final` and `failed_checks`; details are in `work\report.json`. A build takes from a few seconds up to about a minute on a slow machine. Fix every failed check in `content.json` and rebuild. Never change `headless_core.py` to get a pass.

The PDF is rendered in `work\` and then copied with the DOCX to `final\`. If a file in `final\` is open in Word or a PDF viewer, Windows blocks the copy: the only failed check is `final_files_written` and the new PDF stays in `work\_build.pdf`. Review `_build.pdf`, and before finishing ask the user to close the file and build again.

If the build stops with "Vorlage geändert", the student's Vorlage differs from the template the generator knows (`Vorlage_Laborbericht_PG_V2`: title page with `<…>` placeholders, a table of contents, `TOC \c "Abbildung"` and `TOC \c "Tabelle"` indexes, Heading 1/2 sections, the styles `Abbildung (user)`, `Tabelle (user)`, `Code`, and an "Eidesstattliche Erklärung" with an "Ort, Datum" line). Tell the user exactly which part is missing and ask for the right Vorlage. Do not change `headless_core.py` to fit a different template without the user's agreement.

## 7. Look at the result yourself

Read the PDF in `final\` (or `work\_build.pdf`) with the Read tool, whole file, no `pages`. Check against your list from step 2 and on every page image:

- every step, question, Funktionstest point and Erweiterung task is covered,
- every number in the text matches its screenshot,
- no screenshot has cut-off lines or unreadable text, no page is half empty without reason, no lead-in sentence is stranded at a page end, table columns fit,
- title page, Inhalt, Abbildungs- and Tabellenverzeichnis are complete.

Fix and rebuild.

## 8. Teacher review

Start the `laborbericht-lehrer` agent with the report folder path. If your environment cannot start sub-agents, read `.claude\agents\laborbericht-lehrer.md` and do the review yourself in a separate pass, strictly in that role.

- Fix every `MUSS` finding and every `SOLLTE` finding you agree with, rebuild, and review again.
- Stop at `ABGABEBEREIT` or after three rounds; then tell the user what is left and why.
- `RÜCKFRAGE` findings are for the user. Ask them; do not guess.

## 9. Hand over

Tell the user briefly:

- the folder with the DOCX and PDF in `final\`, the page count and the teacher verdict,
- which screenshots were not used and why, and which step headings have no time (and that a screenshot at the start of every task fixes this next time),
- every statement that rests only on their answers,
- that the Reflexion is written in their name and the Eidesstattliche Erklärung confirms they wrote the report themselves and listed all aids, so they must read the report before handing it in.

## Changing an existing report

Edit `work\content.json` in the report folder, rebuild (step 6), check (step 7) and review again (step 8). Never edit the DOCX or PDF in `final\` by hand: the next build overwrites them.

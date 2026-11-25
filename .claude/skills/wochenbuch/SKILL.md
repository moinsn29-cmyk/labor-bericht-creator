---
name: wochenbuch
description: Write a Wochenbuch (weekly summary, LBS 4 Salzburg DTS) as DOCX and PDF from an eduvidual course section and the Wochenbuch Vorlage. Use when the user gives an eduvidual section link or says "Wochenbuch", "mach mein Wochenbuch", "Zusammenfassung fuer die Woche".
argument-hint: "[eduvidual section URL, e.g. https://www.eduvidual.at/course/section.php?id=NNN]"
---

# Wochenbuch

A Wochenbuch is a written summary of one week's topic, handed in as PDF. The Vorlage (`Wochenbuch DTS2.docx`, found in the same eduvidual course) has a title page, a table of contents and the headings "Thema der Woche" (1) with "Zusammenfassung", "Aufgabe" and "Quellen" (1.x). The Laborbericht generator does not fit it, so this skill uses `app/wochenbuch_core.py`.

All paths are relative to the bundle folder. Step 0 of the `laborbericht` skill (environment check, `<PYTHON>`, `DOCUMENTS`) applies unchanged.

## 1. Fetch the material from eduvidual

Moodle has a web service API. `app/fetch_eduvidual.py` logs in with the student's eduvidual user and password, finds the section in one of their courses and downloads its files, pages, links, assignment text and the Vorlage from the same course:

```
EDUVIDUAL_USER=<login> <PYTHON> app/fetch_eduvidual.py "<section URL>" --out eingabe/<name> --password-file <file>
```

- The password comes from a file the student names (or `EDUVIDUAL_PASSWORD`). **Never read, print or copy that file, never pass the password on a command line.** Do not store a token.
- The login name comes from `--user`, `EDUVIDUAL_USER` or `eduvidual_user` in `profil.json`. Ask for it if missing.
- If the login fails ("service not available", invalid login), tell the student; some accounts only have SSO. Fall back to files they download themselves into `eingabe/`.
- The script prints `KURS`, `ABSCHNITT`, `OUT`, `DATEI`, `SEITE`, `LINKS`, `VORLAGE`. Read `abschnitt.txt`, `aufgaben.txt`, `links.txt`, the page `.txt` files, and the PDFs (`pdftotext -layout` works on Linux; otherwise read the PDF whole with the Read tool). Formulas in slides are often images: render the page with `pdftoppm` and look at it.
- Open the linked pages too (curl, strip HTML) so every number in the report is one you have seen. Sites that answer 503 are not cited.

## 2. Ask once

Ask in **one** message: what the student actually did that week (the report must not invent activities), name/Klasse/Schuljahr, the calendar week, whether the use of the AI assistant is listed under Quellen (recommend yes, their decision), and anything the Vorlage title says that looks wrong (the Vorlage may say "DTS 1" for a DTS 2 course).

## 3. Write `work/wochenbuch.json`

Create the report folder `<DOCUMENTS>/<Titel>_<YYYY-MM-DD>/` with `work/` and `final/`, copy the Vorlage into `work/`. Structure: see `app/beispiel_wochenbuch.json`. Rules:

- **Only facts you can back up**: the course material, the pages you read, the student's answers. Numbers in examples are computed and checked. A "case from practice" is a real, cited report (internet) or a clearly labelled calculation, never an invented personal experience.
- **Answer every question of the assignment** (the PDF usually ends with "Arbeitsauftrag Wochenbuch") visibly, in the Aufgabe section.
- **Voice of a second-year Berufsschule student**: plain short sentences, own words, no copied sentences, no personal feelings or past states you cannot know.
- **ASCII only** apart from German letters (ae oe ue ss as real umlauts are fine). No dashes, arrows, curly quotes, Omega, degree sign, superscripts: write "Ohm", "Grad Celsius", "*", "/". The build fails otherwise.
- **No stock phrases** (the list is `BANNED_PHRASES` in `wochenbuch_core.py`; the build fails on them).
- **Length**: 15 to 20 pages including title page and table of contents; the build checks it. The Vorlage uses 12 pt with 1.5 line spacing, about 170 words per page. Do not pad: tables, small diagrams drawn with Pillow (`work/fig_*.png`, text at least about 8 pt at print size) and calculated examples are legitimate content.
- Sub-headings (`h3`) are bold lines; only `h1` and `h2` appear in the table of contents, as in the Vorlage.
- Quellen: the course documents, the pages you really read with URL and retrieval date, and, if the student agreed, "KI-Assistent (Claude) ...".

## 4. Build and check

```
<PYTHON> app/wochenbuch_core.py "<report folder>"
```

It prints `passed`, `pages`, `failed_checks`; details are in `work/report.json`. Fix `wochenbuch.json` and rebuild; never weaken a check in the code. The PDF is rendered by LibreOffice (table of contents refreshed), DOCX and PDF are copied to `final/`. The DOCX asks Word to update its fields once when opened.

Then read the PDF page by page (`Read` with `pages`, at most 8 at a time) for cut-off figures, gaps, justified lists and wrong numbers.

## 5. Hand over

Tell the student: folder, page count, which statements rest on their answers, that nothing about their own experience was invented, that the report is handed in under their name so they must read it, the title/KW choices to confirm, and that the AI line under Quellen is their decision. Never touch git: the student commits.

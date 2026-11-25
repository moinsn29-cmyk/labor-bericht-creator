# Test prompt: Wochenbericht (weekly report) for DTS2

Used to test this toolkit with a different, cheaper model. Fill in the three paths, then paste everything below the line into an agent session.

---

# Task: write my Wochenbericht (weekly report) for DTS2

You are a coding agent on a Windows school laptop. I am a second-year student at the Landesberufsschule 4 Salzburg (IT). I need my Wochenbericht ("Wochenbuch") for the subject DTS2: a short report about what we learned last week, written in my own voice, in German, filled into the school's Word template, delivered as DOCX and PDF.

## Where everything is

- Template: `<path to the Wochenbuch template .docx>`
- What we covered last week (the Skriptum): `<path to the Skriptum .pdf>`
- Toolkit: `<path to this repository>\`. It was built for lab reports, but its environment is what you must use. **Read its `CLAUDE.md` first**, especially "Constraints of the school machines". In short:
  - The only working Python is LibreOffice's: `C:\Program Files\LibreOffice\program\python.exe` (3.11). No system Python, no pip, no admin rights. Do not install anything.
  - python-docx, pypdf, Pillow and lxml are in `<path to this repository>\app\libs`. Put that folder first on `sys.path` in any script you write.
  - Run `powershell -NoProfile -ExecutionPolicy Bypass -File app\check_env.ps1` inside the bundle folder. It prints my Documents folder and confirms the imports work.
  - PowerShell runs in Constrained Language Mode, so no `[Something]::Method()` calls. Put helper scripts in `%TEMP%`, and never name them after standard modules (`inspect.py`, `types.py`, …).
  - Convert DOCX to PDF with `"C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf --outdir <folder> <file.docx>`.
- `<path to this repository>\profil.json` may already hold my name, class and so on. Use it if it exists.

**Do not use `app\headless_core.py` to build this.** It only knows the lab report template and will stop on this one. Fill the Wochenbuch template yourself with python-docx, keeping its styles, headers, footers and layout.

## How to work

1. **Read the template.** Write a small script that lists every paragraph (with style), every table cell, the headers and footers, and every placeholder or instruction text. Work out what the template wants: which fields, which sections, how long. The template's own instructions win over anything in this prompt.
2. **Read the Skriptum completely.** If your file-reading tool cannot handle the PDF, extract the text with pypdf in LibreOffice's Python. Note the topics, definitions, formulas, units and examples that were covered.
3. **Ask me before writing**, in one single message, for everything you cannot know:
   - name, class, Katalog-Nr. (if not in `profil.json`),
   - the week: dates from–to, or the calendar week,
   - which topics of the Skriptum we actually covered that week, if it covers more than one week,
   - what I did myself (exercises, measurements, calculations) and what I found easy, hard or interesting. Give me two or three concrete options to pick from, so I can answer fast,
   - whether the use of an AI assistant should be named in the report, if the template has a field for sources or aids.
4. **Write the text.** Rules:
   - **Sound like me:** first person, plain German as a 16–17-year-old apprentice writes it. Short sentences, normal everyday words, technical terms used correctly but explained simply, now and then a personal remark taken from my answers.
   - **Avoid what students don't write:** no marketing tone, no "zusammenfassend lässt sich sagen", no lists of five adjectives, no perfect textbook phrasing.
   - **Own words:** do not copy sentences from the Skriptum. Formulas and units must be exactly right (for example U = R · I, units V, A, Ω).
   - **Only what is true:** content comes from the Skriptum and my answers. Do not invent exercises, results, measurements or feelings I did not tell you.
   - **Length and structure:** as the template asks. If it says nothing, about one page of text.
5. **Fill the template** with python-docx, save as DOCX, and convert it to PDF with soffice.
   - Output folder: `<Documents>\Wochenbericht_DTS2_<YYYY-MM-DD>\`, with the Monday of that week as the date. Use the `DOCUMENTS=` path from `check_env.ps1`, since it may point into OneDrive.
   - Put your scripts and notes in `work\`, and only the finished `.docx` and `.pdf` in `final\`.
   - File name: whatever the template or Skriptum prescribes. Otherwise `<Klasse>_<Nachname>_Wochenbericht_DTS2_KW<nn>`.
6. **Check your own result.**
   - Open the finished DOCX again in a script: no placeholder or instruction text from the template may be left, every field must be filled, and my name and the dates must be right.
   - Read the PDF and look at the pages: nothing cut off, no empty template boxes, no broken formatting.
   - Fix what you find.
7. **Review it as my teacher would.** Read the PDF once more as the DTS2 teacher. Would you accept this as a student's Wochenbericht? Check that it is correct, complete, in the student's own words, and that the template is followed. Fix every real problem.

## When you are done, tell me

- where the DOCX and PDF are,
- which topics the report covers and which statements rest only on my answers,
- **harness notes:** everything in the toolkit or its instructions that was missing, wrong, confusing or broken (paths, commands, tools that failed, workarounds you needed). Also write this list to `work\harness_notes.md`. This run is a test of the toolkit, so this part matters as much as the report.

Do not change any file inside `<path to this repository>\`. If something there is in your way, work around it and write it down in the harness notes.

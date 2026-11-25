# Laborbericht-Bundle

This folder turns a student's screenshots, Aufgabenstellung (PDF) and Word Vorlage into a finished Laborbericht for the IT-Labor at LBS 4 Salzburg, as DOCX and PDF.

When the user wants a report ("mach meinen Laborbericht", "here are my screenshots, write the report", …), follow the skill in `.claude/skills/laborbericht/SKILL.md` step by step. If your tool does not load skills automatically, read that file and follow it. The teacher review is defined in `.claude/agents/laborbericht-lehrer.md`.

Talk to the user in the language they use (usually German). Ask for missing information in one message, not one question at a time.

This repository contains only code and instructions. The user always provides the Aufgabenstellung, the Word Vorlage, the screenshots and their personal data; none of it is ever committed (see `.gitignore`).

## What is in here

| Path | Purpose |
|---|---|
| `app/headless_core.py` | The generator. Builds DOCX and PDF from `work/content.json`, checks the result, writes `work/report.json`. |
| `app/libs/` | python-docx, pypdf, Pillow, lxml for LibreOffice's Python 3.11 on Windows. |
| `app/mcp_server.py`, `app/mcp_libs/`, `.mcp.json` | Optional MCP server exposing the generator as tools (build, check, inspect). The skill does not need it. |
| `app/check_env.ps1` | Finds LibreOffice's Python and the Documents folder, tests the imports. |
| `app/beispiel_content.json` | Structure example for `content.json`. |
| `eingabe/` | Optional drop folder for the student's files. |
| `profil.json` | Student name, class etc. Created on first use. |

Reports are written to `<Documents>\<Titel>_<YYYY-MM-DD>\` with `screenshots\`, `work\` and `final\` (DOCX + PDF).

## Constraints of the school machines

- **No system Python, no pip, no admin rights.** The only Python is LibreOffice's (`…\LibreOffice\program\python.exe`, 3.11). It has the UNO bridge needed for the PDF export; `app/libs` adds the rest. Never install packages or use another interpreter.
- **LibreOffice is required.** Without it no report can be built; say so and stop.
- **No tkinter, no Pandoc.** Nothing here needs them.
- **The Read tool cannot read page ranges of PDFs** (no pdftoppm). Always read a whole PDF, without the `pages` parameter.
- **Glob and Grep may fail with `spawn UNKNOWN`.** Fall back to reading known paths or listing folders with the shell.
- **Do not name helper scripts after standard modules** (`inspect.py`, `types.py`, …): LibreOffice's Python then imports your file instead and lxml fails with a confusing error. Put helper scripts in `%TEMP%` with distinct names.
- **Files open in Word or a PDF viewer are locked.** The build then leaves the new PDF in `work\_build.pdf`; ask the user to close the file and build again.
- **PowerShell runs in Constrained Language Mode**: .NET calls like `[Environment]::GetFolderPath()` fail. Use cmdlets only, or do the work in LibreOffice's Python.
- **Documents may be redirected to OneDrive.** Use the `DOCUMENTS=` path from `check_env.ps1`, not a guessed `~/Documents`.

## Rules that are not negotiable

- Write only what the screenshots show or the student confirmed. No invented output, times, versions or problems.
- Times in headings come only from screenshot file names. No estimated or made-up times.
- Never edit the generated DOCX or PDF by hand, and never weaken a check in `headless_core.py` to make a build pass.
- The student hands the report in under their own name with an Eidesstattliche Erklärung. Tell them to read it before handing it in.

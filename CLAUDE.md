# Laborbericht-Bundle

This folder turns a student's screenshots, Aufgabenstellung (PDF) and Word Vorlage into a finished Laborbericht for the IT-Labor at LBS 4 Salzburg, as DOCX and PDF.

When the user wants a report ("mach meinen Laborbericht", "here are my screenshots, write the report", …), follow the skill in `.claude/skills/laborbericht/SKILL.md` step by step. If your tool does not load skills automatically, read that file and follow it. The teacher review is defined in `.claude/agents/laborbericht-lehrer.md`.

When the user wants a Wochenbuch (weekly summary, usually given as an eduvidual section link), follow `.claude/skills/wochenbuch/SKILL.md` instead. It uses its own builder, `app/wochenbuch_core.py`, because the Wochenbuch Vorlage is not the Laborbericht Vorlage.

Talk to the user in the language they use (usually German). Ask for missing information in one message, not one question at a time.

This repository contains only code and instructions. The user always provides the Aufgabenstellung, the Word Vorlage, the screenshots and their personal data; none of it is ever committed (see `.gitignore`).

## What is in here

| Path | Purpose |
|---|---|
| `app/headless_core.py` | The generator. Builds DOCX and PDF from `work/content.json`, checks the result, writes `work/report.json`. Detects its mode itself (`MODE`). |
| `app/uno_export.py` | The PDF export through LibreOffice's UNO bridge, in-process or as a helper process. |
| `app/check_env.py`, `requirements.txt` | Normal mode: checks the environment; `--setup` creates `.venv` and installs the packages. |
| `app/check_env.ps1`, `app/libs/` | Locked-down mode: finds LibreOffice's Python; `app/libs` holds python-docx, pypdf, Pillow, lxml for LibreOffice's Python 3.11 on Windows. |
| `app/mcp_server.py`, `app/mcp_libs/`, `.mcp.json` | Optional MCP server exposing the generator as tools (build, check, inspect). The skill does not need it. `.mcp.json` is set up for the locked-down mode; in normal mode install `requirements-mcp.txt` and point it at `.venv`. |
| `app/beispiel_content.json` | Structure example for `content.json`. |
| `app/fetch_eduvidual.py` | Downloads an eduvidual (Moodle) course section through the Moodle web service: files, pages, links, assignment text and the Vorlage. Standard library only. Password from a file the user names, never printed or stored. |
| `app/wochenbuch_core.py`, `app/beispiel_wochenbuch.json` | The Wochenbuch builder (DOCX + PDF from `work/wochenbuch.json`) with its own checks (page range, ASCII, stock phrases, template leftovers); structure example. |
| `eingabe/` | Optional drop folder for the student's files. |
| `profil.json` | Student name, class etc. Created on first use. |

Reports are written to `<Documents>\<Titel>_<YYYY-MM-DD>\` with `screenshots\`, `work\` and `final\` (DOCX + PDF).

## Two modes

The generator produces the same report in both modes; step 0 of the skill decides which one applies.

- **Normal mode**: the machine has its own Python (Linux, macOS, a private Windows PC). `python3 app/check_env.py --setup` creates `.venv` with the packages from `requirements.txt`; build with that Python. The constraints below do not apply, but nothing is installed outside `.venv`. If the generator's Python has no `uno` module, the PDF export runs in a helper process with a Python that has it (LibreOffice's own, or the system Python with `python3-uno`).
- **Locked-down mode**: the school machines. Everything runs with LibreOffice's Python and the bundled `app/libs`. Use it whenever normal mode is not possible.
- **LibreOffice is required in both modes.** Without it no report can be built; say so and stop.

Keep both modes working when changing code: `app/libs` is Python 3.11 on Windows, so the generator must stay compatible with Python 3.11 and with the bundled library versions.

## Constraints of the school machines (locked-down mode)

- **No system Python, no pip, no admin rights.** The only Python is LibreOffice's (`…\LibreOffice\program\python.exe`, 3.11). It has the UNO bridge needed for the PDF export; `app/libs` adds the rest. Never install packages or use another interpreter.
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
- If a rule here conflicts with an explicit request, tell the user what the rule says and why. If they then insist, it is their decision: comply. Work must stay real (real installs, real runs, real screenshots), and the AI help is listed under "Quellen und verwendete Hilfsmittel" (`ki_hinweis`).

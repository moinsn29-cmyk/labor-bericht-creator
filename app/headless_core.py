r"""Tkinter-free report generation core for the Laborbericht app.

A report lives in its own folder:

    <report>/screenshots/            screenshots taken during the lab (PNG or JPG)
    <report>/work/content.json       text, figures and tables of the report
    <report>/work/*.pdf              exercise sheet
    <report>/work/<Vorlage>.docx     the Word template provided by the user
    <report>/work/report.json        manifest written by build_report()
    <report>/final/                  the finished DOCX and PDF, nothing else

It runs in two modes and picks the right one by itself (see ``MODE``):

* normal: an ordinary Python (usually the bundle's ``.venv``) with python-docx, pypdf and
  Pillow installed from ``requirements.txt``:
      python headless_core.py <report folder>
* locked-down: machines without Python, pip or admin rights. LibreOffice's own Python runs
  the generator and the libraries come from ``app/libs``:
      "C:\Program Files\LibreOffice\program\python.exe" headless_core.py <report folder>

LibreOffice is needed in both modes: it refreshes the indexes and writes the PDF.
"""

from __future__ import annotations

import io
import os
import re
import sys
import json
import shutil
import subprocess
import importlib.util
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple


HERE = Path(__file__).resolve().parent
REQUIRED_MODULES = ("docx", "pypdf", "PIL", "lxml")


def _add_local_dependencies() -> str:
    """Make python-docx, pypdf, Pillow and lxml importable and return the mode this implies.

    ``normal``: the running Python already has them (pip, venv); nothing is changed.
    ``locked-down``: it does not, as with LibreOffice's Python on the school machines; the
    bundled copies in ``app/libs`` (Windows, Python 3.11) are put first on the path.
    """
    if str(HERE) not in sys.path:
        sys.path.insert(0, str(HERE))  # uno_export.py
    if all(importlib.util.find_spec(name) is not None for name in REQUIRED_MODULES):
        return "normal"

    candidates = [HERE / "_libs", HERE / "libs"]
    if os.environ.get("LOCALAPPDATA"):
        # Without the variable the path would be relative and pick up a folder in the current directory.
        candidates.insert(0, Path(os.environ["LOCALAPPDATA"]) / "laborbericht-mcp" / "libs")
    for path in candidates:
        if path.exists():
            sys.path.insert(0, str(path))
    return "locked-down"


MODE = _add_local_dependencies()

try:
    from PIL import Image
    from pypdf import PdfReader
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.text.paragraph import Paragraph
except ImportError as exc:
    raise ImportError(
        f"{exc}. Dieses Python ({sys.executable}) hat die benötigten Pakete nicht, und die "
        "mitgelieferten in app/libs passen nur zu LibreOffices Python 3.11 unter Windows. "
        "Abhilfe: 'python app/check_env.py --setup' ausführen (legt .venv an und installiert "
        "requirements.txt) oder den Generator mit LibreOffices Python starten.") from exc

from uno_export import UnoSession, WorkerSession, UnavailableSession


def _find_soffice() -> Path:
    exe = "soffice.exe" if os.name == "nt" else "soffice"
    candidates = []
    if os.environ.get("LIBREOFFICE_PROGRAM"):
        candidates.append(Path(os.environ["LIBREOFFICE_PROGRAM"]) / exe)
    candidates.append(Path(sys.executable).resolve().parent / exe)
    if os.name == "nt":
        if os.environ.get("LOCALAPPDATA"):
            candidates.append(Path(os.environ["LOCALAPPDATA"]) / "Programs" / "LibreOffice" / "program" / exe)
        for root in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            candidates.append(Path(root) / "LibreOffice" / "program" / exe)
    else:
        for name in ("soffice", "libreoffice"):
            found = shutil.which(name)
            if found:
                candidates.append(Path(found))
        candidates += [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
                       Path("/usr/lib/libreoffice/program/soffice"),
                       Path("/opt/libreoffice/program/soffice"),
                       Path("/usr/bin/soffice")]
    return next((c for c in candidates if c.is_file()), candidates[-1])


SOFFICE = _find_soffice()

_uno_python_cache: List[Optional[Path]] = []


def find_uno_python() -> Optional[Path]:
    """A Python that can ``import uno``: this one, LibreOffice's own, or the system's."""
    if importlib.util.find_spec("uno") is not None:
        return Path(sys.executable)
    if _uno_python_cache:
        return _uno_python_cache[0]

    program = SOFFICE.resolve().parent
    candidates = []
    if os.environ.get("LIBREOFFICE_PYTHON"):
        candidates.append(Path(os.environ["LIBREOFFICE_PYTHON"]))
    if os.name == "nt":
        candidates.append(program / "python.exe")
    else:
        candidates += [program / "python",
                       program.parent / "Resources" / "python",  # macOS app bundle
                       Path("/usr/bin/python3")]
        found = shutil.which("python3")
        if found:
            candidates.append(Path(found))

    result = None
    for candidate in candidates:
        if not candidate.is_file():
            continue
        try:
            if subprocess.run([str(candidate), "-c", "import uno"], capture_output=True,
                              stdin=subprocess.DEVNULL, timeout=30).returncode == 0:
                result = candidate
                break
        except (OSError, subprocess.SubprocessError):
            continue
    _uno_python_cache.append(result)
    return result


def _export_session(timeout: float = 90.0):
    """The PDF exporter for this environment (see ``uno_export``)."""
    if importlib.util.find_spec("uno") is not None:
        return UnoSession(SOFFICE, timeout)
    if not SOFFICE.exists():
        return UnavailableSession(f"LibreOffice nicht gefunden: {SOFFICE}")
    uno_python = find_uno_python()
    if uno_python is None:
        return UnavailableSession(
            "UNO nicht verfügbar: kein Python mit dem Modul 'uno' gefunden. Es gehört zu LibreOffice "
            "(Linux: Paket python3-uno installieren) oder LIBREOFFICE_PYTHON auf LibreOffices Python setzen.")
    return WorkerSession(uno_python, SOFFICE, timeout)


def documents_folder() -> str:
    """The user's Documents folder, also when Windows redirects it to OneDrive."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as key:
            value, _ = winreg.QueryValueEx(key, "Personal")
        return os.path.expandvars(value)
    except (ImportError, OSError):
        pass
    try:  # Linux desktops may localise the folder ("Dokumente")
        result = subprocess.run(["xdg-user-dir", "DOCUMENTS"], capture_output=True, text=True, timeout=10)
        folder = result.stdout.strip()
        if result.returncode == 0 and folder and Path(folder) != Path.home():
            return folder
    except (OSError, subprocess.SubprocessError):
        pass
    return str(Path.home() / "Documents")


def environment() -> Dict:
    """What the generator found on this machine (used by check_env and the MCP server)."""
    uno_python = find_uno_python() if SOFFICE.is_file() else None
    return {"mode": MODE, "python": sys.executable, "documents": documents_folder(),
            "soffice": str(SOFFICE), "soffice_found": SOFFICE.is_file(),
            "uno_python": str(uno_python) if uno_python else None}


IMAGE_PATTERNS = ("*.png", "*.jpg", "*.jpeg")


def _images(folder: Path) -> List[Path]:
    return sorted({p for pattern in IMAGE_PATTERNS for p in folder.glob(pattern)})


CONTENT_FILE = "content.json"
SCREENSHOT_DIR = "screenshots"
WORK_DIR = "work"
FINAL_DIR = "final"
BUILD_PDF = "_build.pdf"

# Style names that actually exist in Vorlage_Laborbericht_PG_V2.docx.
STYLE_HEADING_1 = "Heading 1"
STYLE_CAPTION_FIGURE = "Abbildung (user)"
STYLE_CAPTION_TABLE = "Tabelle (user)"
STYLE_TABLE_HEADER = "Tabellenüberschrift (user)"
STYLE_TABLE_BODY = "Tabelleninhalt (user)"
STYLE_CODE = "Code"
STYLE_LIST = "List Paragraph"
STYLE_TOC_ENTRY = "table of figures"

DECLARATION_HEADING = "Eidesstattliche Erklärung"
SIGNATURE_PLACEHOLDER = "Ort, Datum"

# Screenshots are scaled by a fixed pixel density so terminal text has the same size everywhere.
PIXELS_PER_CM = 50.0
DEFAULT_MAX_IMAGE_HEIGHT_CM = 11.0
SPARSE_PAGE_FILL = 0.35


# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------

def extract_pdf(pdf_path: str) -> Dict:
    path = Path(pdf_path)
    if not path.is_file():
        raise FileNotFoundError(pdf_path)
    reader = PdfReader(str(path))
    pages = []
    all_text = []
    for page_number, page in enumerate(reader.pages, 1):
        text = (page.extract_text() or "").strip()
        pages.append({"page": page_number, "text": text})
        all_text.append(text)
    return {"path": str(path), "pages": len(reader.pages), "text": "\n\n".join(all_text), "pages_data": pages}


def ocr_screenshots(folder: str, run_ocr: bool = False) -> List[Dict]:
    """Collect screenshot metadata from ``<report folder>/screenshots``.

    OCR is very slow (minutes for a handful of images) and is only a help for writing
    ``content.json``; the build itself never needs it.
    """
    ocr = None
    if run_ocr:
        try:
            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
        except Exception:
            ocr = None

    image_paths = _images(Path(folder) / SCREENSHOT_DIR)

    results = []
    for image_path in image_paths:
        text = ""
        if ocr is not None:
            try:
                result, _ = ocr(str(image_path))
                if result:
                    text = "\n".join(line[1].strip() for line in result if len(line) > 1)
            except Exception:
                text = ""
        with Image.open(image_path) as image:
            width, height = image.size
        results.append({"path": str(image_path), "name": image_path.name,
                        "width": width, "height": height, "ocr": text})
    return results


def list_assets(folder: str, run_ocr: bool = False) -> Dict:
    """List the source PDFs, content.json and screenshots of a report folder."""
    folder_path = Path(folder)
    content_path = folder_path / WORK_DIR / CONTENT_FILE
    return {
        "folder": str(folder_path),
        "pdfs": sorted(str(p) for p in (folder_path / WORK_DIR).glob("*.pdf") if p.name != BUILD_PDF),
        "content": str(content_path) if content_path.is_file() else None,
        "screenshots": ocr_screenshots(str(folder_path), run_ocr=run_ocr),
    }


def default_template(report_folder: str) -> str:
    """Find the Word template for a report.

    Order: ``template`` in content.json (relative to ``work/``), then any other DOCX in
    ``work/``. The user always provides the template; none is shipped.
    """
    report = Path(report_folder).resolve()
    work = report / WORK_DIR
    content_path = work / CONTENT_FILE
    content = json.loads(content_path.read_text(encoding="utf-8")) if content_path.is_file() else {}

    if content.get("template"):
        named = work / content["template"]
        if not named.is_file():
            raise FileNotFoundError(f"Vorlage aus content.json nicht gefunden: {named}")
        return str(named)

    output = content.get("output_basename")
    in_work = [p for p in sorted(work.glob("*.docx")) if p.stem != output and not p.name.startswith("~$")]
    if in_work:
        return str(in_work[0])
    raise FileNotFoundError(f"Keine Word-Vorlage in {work} gefunden")


def load_content(report_folder: str) -> Dict:
    """Read and validate ``work/content.json`` against the screenshots on disk."""
    content_path = Path(report_folder) / WORK_DIR / CONTENT_FILE
    if not content_path.is_file():
        raise FileNotFoundError(f"{content_path} fehlt")

    content = json.loads(content_path.read_text(encoding="utf-8"))
    for key in ("title", "output_basename", "metadata", "blocks"):
        if key not in content:
            raise ValueError(f"content.json: Schlüssel '{key}' fehlt")
    for key in ("author", "catalog", "school_year", "exercise_number", "date", "place", "signing_date"):
        if not content["metadata"].get(key):
            raise ValueError(f"content.json: metadata.{key} fehlt")

    screenshots = Path(report_folder) / SCREENSHOT_DIR
    on_disk = {p.name for p in _images(screenshots)}
    used = [block["file"] for block in content["blocks"] if block.get("type") == "figure"]
    unused = [entry["file"] for entry in content.get("unused_screenshots", [])]

    # The same screenshot may appear twice only as two different crops.
    views = [(block["file"], tuple(block.get("crop") or ())) for block in content["blocks"]
             if block.get("type") == "figure"]
    duplicates = sorted({name for name, crop in views if views.count((name, crop)) > 1})
    if duplicates:
        raise ValueError(f"content.json: Screenshots mehrfach mit gleichem Ausschnitt verwendet: {duplicates}")
    missing = sorted(set(used + unused) - on_disk)
    if missing:
        raise FileNotFoundError(f"Screenshots fehlen in {screenshots}: {missing}")
    unlisted = sorted(on_disk - set(used) - set(unused))
    if unlisted:
        raise ValueError(
            f"Screenshots weder als Abbildung noch unter unused_screenshots eingetragen: {unlisted}")

    ids = [block["id"] for block in content["blocks"] if block.get("type") == "figure" and block.get("id")]
    if len(ids) != len(set(ids)):
        raise ValueError("content.json: Abbildungs-IDs sind nicht eindeutig")

    problems = chronology_problems(content)
    if problems:
        raise ValueError("content.json: Abbildungen nicht in zeitlicher Reihenfolge:\n  "
                         + "\n  ".join(problems))
    return content


def _screenshot_clock(file_name: str) -> Optional[str]:
    """``Screenshot 2026-11-23 131924.png`` -> ``13:19``."""
    match = re.search(r"\d{4}-\d{2}-\d{2} (\d{2})(\d{2})\d{2}", file_name)
    return f"{match.group(1)}:{match.group(2)}" if match else None


def _screenshot_seconds(file_name: str) -> Optional[str]:
    match = re.search(r"\d{4}-\d{2}-\d{2} (\d{6})", file_name)
    return match.group(1) if match else None


def _caption_time_suffix(file_name: str) -> str:
    clock = _screenshot_clock(file_name)
    return f" ({clock} Uhr)" if clock else ""


def _times_heading(block: Dict) -> bool:
    """A figure dates its step unless it is marked as a later re-run (``"heading_time": false``)."""
    return block.get("type") == "figure" and block.get("heading_time", True)


def heading_times(content: Dict) -> Dict[int, str]:
    """Start time of every level-2 step: the time of its earliest screenshot.

    Steps without a screenshot get no time: the report never states a time it cannot prove.
    """
    times: Dict[int, str] = {}
    current: Optional[int] = None
    for index, block in enumerate(content["blocks"]):
        if block.get("type") == "heading":
            current = index if block["level"] == 2 else None
        elif current is not None and _times_heading(block):
            clock = _screenshot_clock(block["file"])
            if clock and (current not in times or clock < times[current]):
                times[current] = clock
    return times


def heading_display(content: Dict, index: int, times: Optional[Dict[int, str]] = None) -> str:
    """Heading text as printed. Pass ``times`` (from ``heading_times``) when calling this in a loop."""
    block = content["blocks"][index]
    clock = (heading_times(content) if times is None else times).get(index)
    return f"{block['text']} ({clock} Uhr)" if clock else block["text"]


def chronology_problems(content: Dict) -> List[str]:
    """Within each chapter, the screenshots must appear in the order they were taken."""
    problems: List[str] = []
    chapter, last, last_file = "", None, ""
    for block in content["blocks"]:
        if block.get("type") == "heading" and block["level"] == 1:
            chapter, last, last_file = block["text"], None, ""
        elif _times_heading(block):
            stamp = _screenshot_seconds(block["file"])
            if not stamp:
                continue
            if last and stamp < last:
                problems.append(f"{chapter}: {block['file']} steht nach {last_file}")
            last, last_file = stamp, block["file"]
    return problems


def expected_figures(content: Dict) -> List[Dict]:
    """Figures the report must contain, derived from the input only."""
    figures = []
    for block in content["blocks"]:
        if block.get("type") != "figure":
            continue
        number = len(figures) + 1
        caption = block["caption"] + _caption_time_suffix(block["file"])
        figures.append({"number": number, "id": block.get("id"), "file": block["file"],
                        "caption": caption, "entry": f"Abbildung {number}: {caption}"})
    return figures


def expected_tables(content: Dict) -> List[Dict]:
    tables = []
    for block in content["blocks"]:
        if block.get("type") != "table":
            continue
        number = len(tables) + 1
        tables.append({"number": number, "header": block["header"], "rows": block["rows"],
                       "entry": f"Tabelle {number}: {block['caption']}"})
    return tables


def _resolve_refs(text: str, figure_numbers: Dict[str, int]) -> str:
    """Replace ``{abb:<id>}`` with ``Abbildung <n>`` and glue numbers to their units."""
    def replace(match):
        key = match.group(1)
        if key not in figure_numbers:
            raise ValueError(f"content.json: unbekannter Abbildungsverweis {{abb:{key}}}")
        return f"Abbildung\u00a0{figure_numbers[key]}"
    text = re.sub(r"\{abb:([^}]+)\}", replace, text)
    # A non-breaking space keeps "7 %" or "266 MiB" from being split at a line end.
    return re.sub(r"(?<=\d) (?=(?:%|°C|[A-Za-z]{1,3}\b))", "\u00a0", text)


# ---------------------------------------------------------------------------
# Template handling
# ---------------------------------------------------------------------------

def _replace_text_in_paragraph(paragraph, replacements: Dict[str, str]) -> None:
    """Replace placeholders that Word has split across several w:r runs.

    Word frequently breaks a placeholder such as ``< Schuljahr Lehgang z. B.: 25/26 LG-B>``
    into five separate runs.  Matching per run therefore only ever patches a fragment and
    leaves the rest of the placeholder behind.  Instead we join the paragraph text, apply
    every replacement to the joined string, and then write the result back into a single
    run while keeping the formatting of the first run.
    """
    full_text = paragraph.text
    if not full_text:
        return

    new_text = full_text
    for old, value in replacements.items():
        if old in new_text:
            new_text = new_text.replace(old, value)

    if new_text == full_text:
        return

    runs = paragraph.runs
    if not runs:
        paragraph.add_run(new_text)
        return

    for run in runs[1:]:
        run._element.getparent().remove(run._element)
    paragraph.runs[0].text = new_text


def _replace_in_all_paragraphs(document: Document, replacements: Dict[str, str]) -> None:
    """Apply replacements to body paragraphs, table cells, headers and footers."""
    for paragraph in document.paragraphs:
        _replace_text_in_paragraph(paragraph, replacements)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_text_in_paragraph(paragraph, replacements)

    for section in document.sections:
        for container in (section.header, section.footer,
                          section.first_page_header, section.first_page_footer,
                          section.even_page_header, section.even_page_footer):
            try:
                # A header or footer the template does not define is inherited; merely reading
                # its paragraphs would make python-docx add an empty one to the document.
                if container.is_linked_to_previous:
                    continue
                for paragraph in container.paragraphs:
                    _replace_text_in_paragraph(paragraph, replacements)
            except Exception:
                continue


def _align_title_fields(document: Document) -> None:
    """Line up the values of the title page (``Name :<tabs>value``) at one tab stop.

    The template separates label and value with one to three default tabs, which puts the
    values at different horizontal positions depending on the label length.
    """
    for paragraph in document.paragraphs:
        if paragraph.style is not None and paragraph.style.name == STYLE_HEADING_1:
            break
        match = re.match(r"^([^\t]*:)[\t ]*\t[\t ]*(.+)$", paragraph.text)
        if not match or not paragraph.runs:
            continue
        for run in paragraph.runs[1:]:
            run._element.getparent().remove(run._element)
        paragraph.runs[0].text = f"{match.group(1)}\t{match.group(2).strip()}"
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(4.5))


def _start_toc_on_new_page(document: Document) -> None:
    """Keep the title page clean: the table of contents starts on page 2."""
    body = document.element.body
    for child in body:
        if child.tag == qn("w:sdt"):
            previous = child.getprevious()
            if previous is None or previous.tag != qn("w:p"):
                raise RuntimeError("Vorlage geändert: kein Absatz vor dem Inhaltsverzeichnis")
            Paragraph(previous, document).add_run().add_break(_page_break())
            return
    raise RuntimeError("Vorlage geändert: Inhaltsverzeichnis (w:sdt) nicht gefunden")


def _page_break():
    from docx.enum.text import WD_BREAK
    return WD_BREAK.PAGE


def _find_index_paragraph(document: Document, seq_name: str):
    """Return the paragraph holding the ``TOC \\c "<seq_name>"`` field, if present."""
    needle = f'TOC \\c "{seq_name}"'
    for paragraph in document.paragraphs:
        for instr in paragraph._p.iter(qn("w:instrText")):
            if instr.text and needle in instr.text:
                return paragraph
    return None


def _take_template_tail(document: Document) -> List:
    """Remove the template's demo body and detach the declaration block.

    The demo body is everything from the first ``Heading 1`` after the table index up to the
    ``Eidesstattliche Erklärung`` heading.  The declaration with its signature line is the
    template's own wording and is returned so it can be re-attached after the generated body.
    """
    body = document.element.body
    children = list(body)
    if not children or children[-1].tag != qn("w:sectPr"):
        raise RuntimeError("Vorlage geändert: letztes Element ist kein Abschnittsende (sectPr)")

    table_index = _find_index_paragraph(document, "Tabelle")
    if table_index is None:
        raise RuntimeError("Vorlage geändert: Tabellenverzeichnis nicht gefunden")
    if table_index._p not in children:
        raise RuntimeError("Vorlage geändert: Tabellenverzeichnis steht nicht direkt im Dokumenttext")
    index_position = children.index(table_index._p)

    def heading_text(element) -> Optional[str]:
        if element.tag != qn("w:p"):
            return None
        paragraph = Paragraph(element, document)
        if paragraph.style is None or paragraph.style.name != STYLE_HEADING_1:
            return None
        return paragraph.text.strip()

    start = next((i for i in range(index_position + 1, len(children))
                  if heading_text(children[i]) is not None), None)
    end = next((i for i in range(index_position + 1, len(children))
                if heading_text(children[i]) == DECLARATION_HEADING), None)
    if start is None or end is None or start >= end:
        raise RuntimeError(
            "Vorlage geändert: Beispielinhalt zwischen Tabellenverzeichnis und "
            f"'{DECLARATION_HEADING}' nicht gefunden")

    for child in children[start:end]:
        body.remove(child)
    tail = children[end:-1]
    for child in tail:
        body.remove(child)
    return tail


def _attach_declaration(document: Document, tail: List, place: str, signing_date: str) -> None:
    """Re-attach the declaration text, followed by place and date.

    The template's signature block (``Ort, Datum`` / ``Unterschrift``) is dropped by design:
    the report is handed in as a PDF.
    """
    signature_start = next(
        (i for i, element in enumerate(tail)
         if SIGNATURE_PLACEHOLDER in "".join(element.itertext())), None)
    if signature_start is None or signature_start < 2:
        raise RuntimeError(f"Vorlage geändert: Signaturblock ('{SIGNATURE_PLACEHOLDER}') nicht gefunden")

    sect_pr = document.element.body[-1]
    for element in tail[:signature_start]:
        sect_pr.addprevious(element)
    Paragraph(tail[0], document).paragraph_format.keep_with_next = True
    document.add_paragraph(f"{place}, {signing_date}")


def _find_placeholder_leftovers(document: Document) -> List[str]:
    """Return every paragraph text that still contains an unfilled ``<...>`` placeholder."""
    leftovers: List[str] = []
    pattern = re.compile(r"<[^<>]{2,}>")
    for paragraph in document.paragraphs:
        if pattern.search(paragraph.text):
            leftovers.append(paragraph.text.strip())
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if pattern.search(paragraph.text):
                        leftovers.append(paragraph.text.strip())
    return leftovers


def _populate_index(document: Document, seq_name: str,
                    entries: List[Tuple[str, Optional[int]]]) -> List[Paragraph]:
    """Replace a ``TOC \\c`` field with plain index entries, optionally with page numbers.

    LibreOffice does not treat these ``TOC \\c`` fields as real index objects, so
    ``getDocumentIndexes()`` never refreshes them and they would stay empty.  Writing the
    entries directly is deterministic and always matches the captions in the body.

    ``entries`` is a list of ``(text, page_number)``.  When the page number is ``None`` the
    entry is written without a page reference (used for the first, unmeasured pass).
    """
    paragraph = _find_index_paragraph(document, seq_name)
    if paragraph is None:
        raise RuntimeError(f"Vorlage geändert: Verzeichnis für '{seq_name}' nicht gefunden")

    # Drop every child (field chars, instrText, cached demo hyperlink) but keep the paragraph
    # properties so the index keeps its tab stops and dot leader.
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)

    if not entries:
        return []

    def fill(target, text: str, page: Optional[int]) -> None:
        target.style = document.styles[STYLE_TOC_ENTRY]
        target.add_run(text)
        if page is not None:
            target.add_run("\t")
            target.add_run(str(page))

    fill(paragraph, entries[0][0], entries[0][1])
    created = [paragraph]

    anchor = paragraph._p
    for text, page in entries[1:]:
        new_paragraph = document.add_paragraph()
        anchor.addnext(new_paragraph._p)
        fill(new_paragraph, text, page)
        created.append(new_paragraph)
        anchor = new_paragraph._p
    return created


# ---------------------------------------------------------------------------
# Building blocks
# ---------------------------------------------------------------------------

def _add_heading(document: Document, text: str, level: int):
    heading = document.add_heading(text, level=level)
    heading.paragraph_format.keep_with_next = True
    return heading


def _add_code(document: Document, lines: List[str]) -> None:
    for index, line in enumerate(lines):
        paragraph = document.add_paragraph(style=STYLE_CODE)
        run = paragraph.add_run(line)
        # Column-aligned terminal output relies on runs of spaces surviving.
        for text_node in run._r.iter(qn("w:t")):
            text_node.set(qn("xml:space"), "preserve")
        if index < len(lines) - 1:
            paragraph.paragraph_format.keep_with_next = True


def _add_bullets(document: Document, items: List[str]) -> None:
    for item in items:
        paragraph = document.add_paragraph(style=STYLE_LIST)
        # Hanging indent: wrapped lines line up with the text, not with the bullet.
        paragraph.paragraph_format.first_line_indent = Cm(-0.45)
        paragraph.add_run("•\t").bold = True
        paragraph.add_run(str(item))


def _picture_size_cm(width_px: int, height_px: int, max_width_cm: float,
                     max_height_cm: float) -> float:
    """Width for a screenshot: fixed pixel density, capped by text width and a maximum height."""
    width_cm = min(max_width_cm, width_px / PIXELS_PER_CM)
    if width_cm * height_px / width_px > max_height_cm:
        width_cm = max_height_cm * width_px / height_px
    return round(width_cm, 2)


PNG_MODES = ("1", "L", "LA", "I", "I;16", "P", "RGB", "RGBA")


def _picture_png(image_path: Path, crop: Optional[List[int]]) -> Tuple[bytes, int, int]:
    """The (cropped) screenshot as PNG bytes with its pixel size."""
    with Image.open(image_path) as image:
        if crop:
            left, top, right, bottom = crop
            if not (0 <= left < right <= image.width and 0 <= top < bottom <= image.height):
                raise ValueError(f"crop {crop} liegt außerhalb von {image_path.name} {image.size}")
            image = image.crop((left, top, right, bottom))
        if image.mode not in PNG_MODES:  # e.g. a CMYK JPEG
            image = image.convert("RGB")
        width_px, height_px = image.size
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
    return buffer.getvalue(), width_px, height_px


def _add_picture(document: Document, image_path: Path, crop: Optional[List[int]],
                 max_width_cm: float, max_height_cm: float, cache: Optional[Dict] = None):
    """Insert a screenshot. ``cache`` saves re-encoding the same image in every layout pass."""
    key = (str(image_path), tuple(crop or ()))
    if cache is not None and key in cache:
        data, width_px, height_px = cache[key]
    else:
        data, width_px, height_px = _picture_png(image_path, crop)
        if cache is not None:
            cache[key] = (data, width_px, height_px)

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.keep_with_next = True
    paragraph.add_run().add_picture(
        io.BytesIO(data), width=Cm(_picture_size_cm(width_px, height_px, max_width_cm, max_height_cm)))
    return paragraph


def _add_seq_field(paragraph, seq_name: str, number: int) -> None:
    """Insert a ``SEQ`` field whose cached result already shows the right number."""
    begin_run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin_run._r.append(begin)

    instr_run = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" SEQ {seq_name} \\* ARABIC "
    instr_run._r.append(instr)

    sep_run = paragraph.add_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    sep_run._r.append(sep)

    paragraph.add_run(str(number))

    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)


def _add_caption(document: Document, label: str, number: int, text: str, style: str,
                 keep_with_next: bool = False):
    paragraph = document.add_paragraph(style=style)
    if keep_with_next:
        paragraph.paragraph_format.keep_with_next = True
    paragraph.add_run(f"{label} ")
    _add_seq_field(paragraph, label, number)
    paragraph.add_run(f": {text}")
    return paragraph


def _set_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")
        borders.append(element)
    # Word expects the children of w:tblPr in schema order.
    tbl_pr.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
                                 "w:tblCaption", "w:tblDescription", "w:tblPrChange")


def _keep_table_together(table) -> None:
    """Prevent the table from being split across a page break."""
    rows = table.rows
    for row in rows:
        cant_split = OxmlElement("w:cantSplit")
        cant_split.set(qn("w:val"), "true")
        row._tr.get_or_add_trPr().append(cant_split)

    # keep_with_next on every row except the last glues the rows to one another.
    for row in rows[:-1]:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.keep_with_next = True


def _add_table(document: Document, header: List[str], rows: List[List[str]],
               widths_cm: List[float]) -> None:
    if len(widths_cm) != len(header) or any(len(row) != len(header) for row in rows):
        raise ValueError("content.json: Tabelle hat uneinheitliche Spaltenzahl")

    table = document.add_table(rows=1, cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_borders(table)

    header_row = table.rows[0]
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    header_row._tr.get_or_add_trPr().append(repeat)
    for cell, text, width in zip(header_row.cells, header, widths_cm):
        cell.width = Cm(width)
        paragraph = cell.paragraphs[0]
        paragraph.style = document.styles[STYLE_TABLE_HEADER]
        paragraph.add_run(text).bold = True

    for values in rows:
        row = table.add_row()
        for cell, text, width in zip(row.cells, values, widths_cm):
            cell.width = Cm(width)
            paragraph = cell.paragraphs[0]
            paragraph.style = document.styles[STYLE_TABLE_BODY]
            paragraph.add_run(text)

    # LibreOffice takes the column widths from the table grid, not from the cells.
    for column, width in zip(table.columns, widths_cm):
        column.width = Cm(width)

    _keep_table_together(table)
    # A table directly followed by a heading needs a paragraph in between to keep its spacing.
    document.add_paragraph().paragraph_format.space_after = Pt(0)


def _compact_index_paragraphs(paragraphs: List[Paragraph]) -> None:
    for paragraph in paragraphs:
        fmt = paragraph.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0


def _text_width_twips(document: Document) -> int:
    section = document.sections[0]
    usable = section.page_width - section.left_margin - section.right_margin
    return int(usable / 635)  # 1 twip = 635 EMU


def _compact_toc_styles(document: Document) -> None:
    """Tighten the TOC entry styles and pull their right tab stop inside the text width.

    LibreOffice regenerates the TOC during the UNO export and takes both the spacing and the
    tab stops from the ``toc 1``..``toc 3`` styles.  The template's right tab stop sits exactly
    at the text width, so every entry wraps onto a second line and the table of contents
    doubles in length.  Moving the stop a little inside fixes that.
    """
    right_pos = max(1000, _text_width_twips(document) - 120)

    for name in ("toc 1", "toc 2", "toc 3", "toc 4", STYLE_TOC_ENTRY):
        try:
            style = document.styles[name]
        except KeyError:
            continue

        fmt = style.paragraph_format
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0

        ppr = style.element.get_or_add_pPr()
        for existing in ppr.findall(qn("w:tabs")):
            ppr.remove(existing)

        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "right")
        tab.set(qn("w:leader"), "dot")
        tab.set(qn("w:pos"), str(right_pos))
        # get_or_add_tabs() puts w:tabs where the schema wants it (before w:spacing).
        ppr.get_or_add_tabs().append(tab)


def _build_document(template_path: str, content: Dict, screenshots: Path,
                    page_numbers: Optional[Dict[str, int]] = None,
                    image_cache: Optional[Dict] = None) -> Document:
    """Build the report from the template and ``content``.

    ``page_numbers`` maps an index entry text to the page it was measured on during a previous
    layout pass.  When it is ``None`` the indexes are written without page references.
    """
    meta = content["metadata"]
    figures = expected_figures(content)
    tables = expected_tables(content)
    figure_numbers = {figure["id"]: figure["number"] for figure in figures if figure["id"]}

    document = Document(template_path)
    _replace_in_all_paragraphs(document, {
        "< Title der Übung >": content["title"],
        "<Fortlaufende Nummer >": meta["exercise_number"],
        "<ihre Klassen Nummer>": meta["catalog"],
        "<Ihr Vor- und Nachname>": meta["author"],
        "< Schuljahr Lehgang z. B.: 25/26 LG-B>": meta["school_year"],
        "<Datum  TT.MM.JJJJ>": meta["date"],
    })
    _align_title_fields(document)
    _start_toc_on_new_page(document)
    tail = _take_template_tail(document)

    section = document.sections[0]
    text_width_cm = (section.page_width - section.left_margin - section.right_margin) / 360000

    blocks = content["blocks"]
    times = heading_times(content)
    figure_iter = iter(figures)
    table_iter = iter(tables)
    for index, block in enumerate(blocks):
        kind = block.get("type")
        next_kind = blocks[index + 1].get("type") if index + 1 < len(blocks) else None

        if kind == "heading":
            _add_heading(document, heading_display(content, index, times), block["level"])
        elif kind == "text":
            paragraph = document.add_paragraph(_resolve_refs(block["text"], figure_numbers))
            if next_kind in ("code", "bullets", "table"):
                paragraph.paragraph_format.keep_with_next = True
        elif kind == "code":
            _add_code(document, block["lines"])
        elif kind == "bullets":
            _add_bullets(document, [_resolve_refs(item, figure_numbers) for item in block["items"]])
        elif kind == "figure":
            figure = next(figure_iter)
            _add_picture(document, screenshots / figure["file"], block.get("crop"),
                         text_width_cm, block.get("max_height_cm", DEFAULT_MAX_IMAGE_HEIGHT_CM),
                         image_cache)
            _add_caption(document, "Abbildung", figure["number"], figure["caption"],
                         STYLE_CAPTION_FIGURE)
        elif kind == "table":
            table = next(table_iter)
            _add_caption(document, "Tabelle", table["number"], block["caption"],
                         STYLE_CAPTION_TABLE, keep_with_next=True)
            _add_table(document, block["header"],
                       [[_resolve_refs(cell, figure_numbers) for cell in row] for row in block["rows"]],
                       block["widths_cm"])
        else:
            raise ValueError(f"content.json: unbekannter Blocktyp '{kind}'")

    _attach_declaration(document, tail, meta["place"], meta["signing_date"])

    def with_pages(entries: List[Dict]) -> List[Tuple[str, Optional[int]]]:
        return [(entry["entry"], (page_numbers or {}).get(entry["entry"])) for entry in entries]

    index_paragraphs = (_populate_index(document, "Abbildung", with_pages(figures))
                        + _populate_index(document, "Tabelle", with_pages(tables)))
    _compact_index_paragraphs(index_paragraphs)
    _compact_toc_styles(document)
    return document


# ---------------------------------------------------------------------------
# PDF measurement
# ---------------------------------------------------------------------------

def _normalise(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _page_texts(reader: PdfReader) -> List[str]:
    return [_normalise(page.extract_text() or "") for page in reader.pages]


def _pdf_page_texts(pdf_path: str) -> List[str]:
    return _page_texts(PdfReader(str(pdf_path)))


def _index_page_set(page_texts: List[str]) -> set:
    return {number for number, text in enumerate(page_texts, 1)
            if "Abbildungsverzeichnis" in text or "Tabellenverzeichnis" in text}


def _measure_index_pages(page_texts: List[str], entry_texts: List[str]) -> Dict[str, int]:
    """Find the page each caption appears on in the rendered PDF, ignoring the index pages.

    The caption is the last hit: an index long enough to run over onto a page without its
    heading comes before the body and must not be mistaken for the caption.
    """
    index_pages = _index_page_set(page_texts)

    pages: Dict[str, int] = {}
    for text in entry_texts:
        needle = _normalise(text)
        for page_number in range(len(page_texts), 0, -1):
            if page_number not in index_pages and needle in page_texts[page_number - 1]:
                pages[text] = page_number
                break
    return pages


def _count_pdf_images(reader: PdfReader) -> int:
    """Count distinct image objects placed on the pages of the PDF."""
    seen = set()
    anonymous = 0
    for page in reader.pages:
        resources = page.get("/Resources")
        xobjects = resources.get_object().get("/XObject") if resources is not None else None
        if xobjects is None:
            continue
        xobjects = xobjects.get_object()
        for name in xobjects:
            if xobjects[name].get("/Subtype") != "/Image":
                continue
            reference = xobjects.raw_get(name)
            if hasattr(reference, "idnum"):
                seen.add(reference.idnum)
            else:
                anonymous += 1
    return len(seen) + anonymous


def _measure_page_fill(reader: PdfReader) -> List[Dict]:
    """Measure how much of each page's height is used by text and images.

    ``fill`` is the vertical extent of the content as a fraction of the body area.  The footer
    ("Seite x von y") sits inside the template's 3.4 cm bottom margin, so everything below
    11 % of the page height is ignored; otherwise every page would look full.
    """
    results: List[Dict] = []
    for page_no, page in enumerate(reader.pages, 1):
        page_height = float(page.mediabox.height or 842)
        low, high = page_height * 0.11, page_height * 0.92
        ys: List[float] = []

        def on_text(text, cm, tm, font_dict, font_size):
            if text.strip():
                # ``tm`` is relative to the current transformation matrix ``cm``.
                y = cm[1] * tm[4] + cm[3] * tm[5] + cm[5]
                if low <= y <= high:
                    ys.append(y)

        def on_operand(operator, operands, cm, tm):
            # ``Do`` paints an image into the unit square transformed by ``cm``.
            if operator == b"Do":
                for y in (cm[5], cm[5] + cm[3]):
                    ys.append(min(max(y, low), high))

        try:
            page.extract_text(visitor_text=on_text, visitor_operand_before=on_operand)
        except Exception:
            pass

        span = (max(ys) - min(ys)) if len(ys) > 1 else 0.0
        results.append({"page": page_no, "fill": round(min(span / (high - low), 1.0), 3)})
    return results


# ---------------------------------------------------------------------------
# LibreOffice export
# ---------------------------------------------------------------------------

def export_pdf_with_uno(docx_path: str, pdf_path: str, timeout: float = 90.0) -> Tuple[bool, str]:
    """Refresh all indexes and export the PDF via the LibreOffice UNO bridge.

    ``python-docx`` cannot evaluate fields, so the table of contents would otherwise keep the
    template's entries.  We start a headless soffice with a UNO socket, load the document,
    update every index, and export straight to PDF (see ``uno_export``).

    The DOCX itself is deliberately **not** written back: a LibreOffice round-trip renames the
    template's custom styles (``Abbildung (user)`` and friends), which would break the
    verification step and any later editing in Word.
    """
    session = _export_session(timeout)
    try:
        return session.export(docx_path, pdf_path)
    finally:
        session.close()


def export_with_libreoffice(input_path: str, output_dir: str, target_format: str = "pdf") -> Tuple[bool, str]:
    input_path = str(Path(input_path).resolve())
    output_dir = str(Path(output_dir).resolve())
    if not SOFFICE.exists():
        return False, f"LibreOffice nicht gefunden: {SOFFICE}"

    cmd = [str(SOFFICE), "--headless", "--norestore", "--convert-to", target_format,
           "--outdir", output_dir, input_path]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode == 0:
            return True, "LibreOffice-Konvertierung erfolgreich"
        return False, (result.stderr or result.stdout or f"Exit code {result.returncode}").strip()
    except subprocess.TimeoutExpired:
        return False, "LibreOffice-Konvertierung hat das Zeitlimit überschritten"


# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------

FORBIDDEN_TOKENS = [
    "Beispieltabelle", "Spaltenüberschrift", "Zelleninhalt", "Firmwareversion",
    "Text & Bilder & Zeitstempel", "Quellen- und Hilfsmittelangaben", "pritz@testpc",
    "Screenshot zur Übung", "<Schritt", "<Teilschritt", SIGNATURE_PLACEHOLDER,
]


def _verify_index_pages(page_texts: List[str], entry_texts: List[str],
                        index_entries: Dict[str, int]) -> Dict:
    """Every expected index entry needs a page number, and it must be the caption's real page."""
    actual_pages = _measure_index_pages(page_texts, entry_texts)
    missing = [text for text in entry_texts if index_entries.get(text) is None]
    mismatches = [{"entry": text, "claimed": index_entries.get(text), "actual": actual_pages.get(text)}
                  for text in entry_texts
                  if index_entries.get(text) is not None and index_entries[text] != actual_pages.get(text)]

    # The printed index line must carry the number as well: "<entry> ... <page>".  The indexes
    # are the pages with an index heading plus whatever else precedes the first caption.
    index_pages = _index_page_set(page_texts)
    first_caption_page = min(actual_pages.values(), default=0)
    index_text = " ".join(text for number, text in enumerate(page_texts, 1)
                          if number in index_pages or number < first_caption_page)
    unprinted = [text for text in entry_texts
                 if index_entries.get(text) is not None
                 and not re.search(re.escape(_normalise(text)) + r"[\s.]*" + str(index_entries[text]) + r"(?!\d)",
                                   index_text)]
    return {"entries": len(entry_texts), "missing_page_numbers": missing,
            "mismatches": mismatches, "not_printed_in_index": unprinted}


def verify_report(report_folder: str, docx_path: Optional[str] = None,
                  pdf_path: Optional[str] = None,
                  index_entries: Optional[Dict[str, int]] = None) -> Dict:
    """Check the generated files against the report's *input*.

    The expectation (which figures, captions and tables must exist) is derived from
    ``work/content.json`` and the screenshot folder, never from the build's own output.
    """
    content = load_content(report_folder)
    figures = expected_figures(content)
    tables = expected_tables(content)
    docx_path = docx_path or str(Path(report_folder) / WORK_DIR / f"{content['output_basename']}.docx")
    default_pdf = Path(report_folder) / FINAL_DIR / f"{content['output_basename']}.pdf"
    if pdf_path is None and default_pdf.is_file():
        pdf_path = str(default_pdf)

    document = Document(docx_path)
    paragraphs = document.paragraphs
    all_text = "\n".join(p.text for p in paragraphs)

    found_forbidden = [token for token in FORBIDDEN_TOKENS if token in all_text]
    leftovers = _find_placeholder_leftovers(document)

    def styled(style_name: str) -> List[str]:
        return [p.text for p in paragraphs if p.style is not None and p.style.name == style_name]

    captions = styled(STYLE_CAPTION_FIGURE)
    table_captions = styled(STYLE_CAPTION_TABLE)
    expected_captions = [figure["entry"] for figure in figures]
    expected_table_captions = [table["entry"] for table in tables]

    # Every figure caption must be directly preceded by a paragraph containing an image.
    orphan_captions = []
    for index, paragraph in enumerate(paragraphs):
        if paragraph.style is None or paragraph.style.name != STYLE_CAPTION_FIGURE:
            continue
        previous = paragraphs[index - 1] if index > 0 else None
        if previous is None or not previous._p.findall(".//" + qn("w:drawing")):
            orphan_captions.append(paragraph.text)
    drawings = len(document.element.body.findall(".//" + qn("w:drawing")))

    tables_ok = len(document.tables) == len(tables) and all(
        [cell.text.strip() for cell in built.rows[0].cells] == expected["header"]
        and len(built.rows) == len(expected["rows"]) + 1
        for built, expected in zip(document.tables, tables))

    meta = content["metadata"]
    times = heading_times(content)
    title_values = [content["title"], meta["author"], meta["catalog"], meta["school_year"], meta["date"]]

    checks = {
        "no_placeholder_leftovers": not leftovers,
        "no_template_artifacts": not found_forbidden,
        "title_page_filled": all(value in all_text for value in title_values),
        "figure_captions_match_input": captions == expected_captions,
        "figure_captions_distinct": len(set(captions)) == len(captions),
        "all_captions_have_image": not orphan_captions,
        "image_count_matches_input": drawings == len(figures),
        "table_captions_match_input": table_captions == expected_table_captions,
        "tables_match_input": tables_ok,
        "declaration_with_place_and_date":
            DECLARATION_HEADING in all_text and f"{meta['place']}, {meta['signing_date']}" in all_text,
        "signature_block_removed": not any(p.text.strip() == "Unterschrift" for p in paragraphs),
        "step_headings_timed_from_screenshots": styled("Heading 2") == [
            heading_display(content, i, times) for i, block in enumerate(content["blocks"])
            if block.get("type") == "heading" and block["level"] == 2],
        "screenshots_in_chronological_order": not chronology_problems(content),
    }

    result = {
        "placeholder_leftovers": leftovers,
        "forbidden_tokens_found": found_forbidden,
        "expected_figures": len(figures),
        "figure_captions": captions,
        "missing_figures": [entry for entry in expected_captions if entry not in captions],
        "table_captions": table_captions,
        "orphan_captions": orphan_captions,
        "images_in_docx": drawings,
    }

    checks["pdf_exists"] = bool(pdf_path) and Path(pdf_path).is_file()
    if checks["pdf_exists"]:
        reader = PdfReader(str(pdf_path))  # parsed once for every check below
        page_texts = _page_texts(reader)
        pdf_text = " ".join(page_texts)
        pdf_images = _count_pdf_images(reader)
        headings = [heading_display(content, i, times) for i, block in enumerate(content["blocks"])
                    if block.get("type") == "heading"]
        headings.append(DECLARATION_HEADING)
        missing_toc = [h for h in headings if pdf_text.count(_normalise(h)) < 2]
        page_fill = _measure_page_fill(reader)
        sparse_pages = [p for p in page_fill[:-1] if p["fill"] < SPARSE_PAGE_FILL]

        checks["pdf_image_count_matches_input"] = pdf_images == len(figures)
        checks["toc_lists_all_headings"] = not missing_toc
        checks["no_sparse_pages"] = not sparse_pages
        result.update({"pages": len(page_texts), "images_in_pdf": pdf_images,
                       "headings_missing_in_toc": missing_toc, "page_fill": page_fill,
                       "sparse_pages": sparse_pages})

        if index_entries is None:
            # Stand-alone check: read the numbers the indexes were built with from the DOCX.
            index_entries = {}
            for text in styled(STYLE_TOC_ENTRY):
                entry, _, page = text.partition("\t")
                index_entries[entry] = int(page) if page.strip().isdigit() else None
        index_check = _verify_index_pages(
            page_texts, expected_captions + expected_table_captions, index_entries)
        checks["all_index_entries_have_pages"] = not index_check["missing_page_numbers"]
        checks["index_page_numbers_correct"] = (
            not index_check["mismatches"] and not index_check["not_printed_in_index"])
        result["index_page_check"] = index_check
        result["index_page_numbers"] = index_entries

    result["checks"] = checks
    result["passed"] = all(checks.values())
    return result


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def _write_readme(report_folder: Path, content: Dict, template_path: str) -> None:
    meta = content["metadata"]
    lines = [
        f"# {content['title']}",
        "",
        f"Laborbericht zur IT-Laborübung Nr. {meta['exercise_number']} vom {meta['date']}, "
        f"verfasst von {meta['author']} ({meta['catalog']}, {meta['school_year']}).",
        "",
        f"- `final/{content['output_basename']}.pdf` ist die fertige Abgabedatei, daneben die Word-Fassung.",
        "- `screenshots/` enthält die Bildschirmaufnahmen der Übung.",
        "- `work/` enthält die Arbeitsdateien: Übungsunterlage, Vorlage, `content.json` (Text des Berichts)",
        "  und `report.json` (Protokoll des letzten Laufs samt Prüfergebnis).",
        "",
        "Neu erzeugen:",
        "",
        "```",
        f'"{Path(sys.executable).resolve()}" "{Path(__file__).resolve()}" "{report_folder}"',
        "```",
        "",
    ]
    (report_folder / WORK_DIR / "README.md").write_text("\n".join(lines), encoding="utf-8")


def _locked(path: Path) -> bool:
    """True when an existing file cannot be replaced (open in Word or a PDF viewer on Windows)."""
    if not path.exists():
        return False
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True


def build_report(report_folder: str, template_path: Optional[str] = None,
                 max_layout_passes: int = 4) -> Dict:
    """Build the DOCX (``work/``) and the hand-in PDF (``final/``) and verify the result.

    Index page numbers can only be measured on a rendered PDF, so at least two layout passes
    are needed: one to measure, one to print the numbers.  The loop stops as soon as a pass
    reproduces the numbers it was built with.
    """
    if max_layout_passes < 2:
        raise ValueError("max_layout_passes muss mindestens 2 sein (messen, dann drucken)")

    report_folder = Path(report_folder).resolve()
    template_path = str(Path(template_path).resolve()) if template_path else default_template(str(report_folder))
    if not Path(template_path).is_file():
        raise FileNotFoundError(template_path)

    content = load_content(str(report_folder))
    screenshots = report_folder / SCREENSHOT_DIR
    work_folder = report_folder / WORK_DIR
    final_folder = report_folder / FINAL_DIR
    final_folder.mkdir(parents=True, exist_ok=True)
    docx_path = work_folder / f"{content['output_basename']}.docx"
    final_pdf = final_folder / f"{content['output_basename']}.pdf"
    # Rendered in work/ first: final/ only ever receives a finished build, and a hand-in PDF
    # that is open in a viewer (locked on Windows) cannot break the build half-way.
    pdf_path = work_folder / BUILD_PDF

    entry_texts = ([figure["entry"] for figure in expected_figures(content)]
                   + [table["entry"] for table in expected_tables(content)])

    page_numbers: Optional[Dict[str, int]] = None
    export_message = "nicht ausgeführt"
    passes_used = 0
    converged = False

    image_cache: Dict = {}
    session = _export_session()
    try:
        for attempt in range(1, max_layout_passes + 1):
            passes_used = attempt
            _build_document(template_path, content, screenshots, page_numbers, image_cache).save(str(docx_path))

            if pdf_path.exists():
                try:
                    pdf_path.unlink()
                except PermissionError:
                    raise RuntimeError(f"{pdf_path} ist in einem anderen Programm geöffnet. "
                                       "Datei schließen und neu erzeugen.") from None
            exported, export_message = session.export(str(docx_path), str(pdf_path))
            if not exported or not pdf_path.exists():
                break

            measured = _measure_index_pages(_pdf_page_texts(str(pdf_path)), entry_texts)
            if page_numbers is not None and measured == page_numbers:
                converged = True
                break
            page_numbers = measured
    finally:
        session.close()

    verification = verify_report(
        str(report_folder), str(docx_path),
        str(pdf_path) if pdf_path.exists() else None,
        index_entries=page_numbers or {})
    verification["checks"]["layout_converged"] = converged

    final_docx = final_folder / docx_path.name
    final_message = "nicht geschrieben, weil kein PDF erzeugt wurde"
    final_written = False
    if pdf_path.exists():
        try:
            # Checked up front so final/ never ends up with a new DOCX next to an old PDF.
            if _locked(final_docx) or _locked(final_pdf):
                raise PermissionError
            shutil.copyfile(docx_path, final_docx)
            shutil.copyfile(pdf_path, final_pdf)
            pdf_path.unlink()
            final_written, final_message = True, "geschrieben"
        except PermissionError:
            final_message = (f"Eine Datei in {FINAL_DIR}/ ist in einem anderen Programm geöffnet und konnte "
                             f"nicht ersetzt werden. Datei schließen und neu erzeugen. Der neue Stand liegt "
                             f"in {WORK_DIR}/.")
    verification["checks"]["final_files_written"] = final_written
    verification["passed"] = all(verification["checks"].values())

    source_pdfs = sorted(p.name for p in work_folder.glob("*.pdf") if p.name != BUILD_PDF)
    manifest = {
        "title": content["title"],
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "mode": MODE,
        "metadata": content["metadata"],
        "template": template_path,
        "source_pdf": f"{WORK_DIR}/{source_pdfs[0]}" if source_pdfs else None,
        "docx": f"{WORK_DIR}/{docx_path.name}",
        "pdf": f"{FINAL_DIR}/{final_pdf.name}" if final_written else None,
        "final_docx": f"{FINAL_DIR}/{final_docx.name}" if final_written else None,
        "final_message": final_message,
        "screenshots": len(_images(screenshots)),
        "screenshots_unused": content.get("unused_screenshots", []),
        "figures_embedded": verification["images_in_docx"],
        "pages": verification.get("pages"),
        "layout_passes": passes_used,
        "pdf_export_message": export_message,
        "verification": verification,
    }
    (work_folder / "report.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    _write_readme(report_folder, content, template_path)
    return manifest


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("Aufruf: headless_core.py <report_folder> [template.docx]")
    manifest = build_report(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    checks = manifest["verification"]["checks"]
    summary = {"passed": manifest["verification"]["passed"], "pages": manifest["pages"],
               "layout_passes": manifest["layout_passes"],
               "final": manifest["final_message"],
               "failed_checks": [name for name, ok in checks.items() if not ok]}
    sys.stdout.buffer.write(json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8"))
    sys.exit(0 if manifest["verification"]["passed"] else 1)

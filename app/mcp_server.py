"""MCP server for the Laborbericht generator (optional; the skill works without it).

Start with LibreOffice's Python (locked-down mode):
    "C:\\Program Files\\LibreOffice\\program\\python.exe" app\\mcp_server.py
or, in normal mode, with a Python that has requirements-mcp.txt installed.
"""

from __future__ import annotations

import sys
import importlib.util
from pathlib import Path

APP = Path(__file__).resolve().parent
if importlib.util.find_spec("mcp") is None:
    # Locked-down mode: LibreOffice's Python has no mcp package, use the bundled Windows copies.
    # app/libs is needed here already (typing_extensions); headless_core decides the same way
    # for its own libraries.
    for path in (APP / "mcp_libs", APP / "mcp_libs" / "win32", APP / "mcp_libs" / "win32" / "lib", APP / "libs"):
        if path.exists():
            sys.path.insert(0, str(path))
sys.path.insert(0, str(APP))

from mcp.server.fastmcp import FastMCP

from headless_core import (
    build_report,
    default_template,
    environment,
    export_with_libreoffice,
    extract_pdf,
    list_assets,
    verify_report,
)


mcp = FastMCP(
    "Laborbericht Generator",
    instructions="Erstellt Laborberichte aus work/content.json, Screenshots und der Word-Vorlage des Schülers.",
)


@mcp.tool()
def get_status() -> dict:
    """Return the mode, the Python in use, LibreOffice and the Documents folder for reports."""
    return {"status": "ready", **environment()}


@mcp.tool()
def inspect_assets(report_folder: str, run_ocr: bool = False) -> dict:
    """List the PDFs, content.json and screenshots of a report folder."""
    return list_assets(report_folder, run_ocr=run_ocr)


@mcp.tool()
def read_source_pdf(pdf_path: str) -> dict:
    """Extract text from a source PDF, including per-page text."""
    return extract_pdf(pdf_path)


@mcp.tool()
def create_report(report_folder: str, template_path: str = "") -> dict:
    """Build final/*.docx and final/*.pdf for a report folder and write work/report.json.

    Text, metadata, figures and tables come from work/content.json. Without template_path the
    Vorlage in work/ (or the one named in content.json) is used.
    """
    return build_report(report_folder, template_path or default_template(report_folder))


@mcp.tool()
def export_docx(input_path: str, output_folder: str, target_format: str = "pdf") -> dict:
    """Export a DOCX file with LibreOffice as pdf, docx, html, odt, or txt."""
    success, message = export_with_libreoffice(input_path, output_folder, target_format)
    return {"success": success, "message": message, "output_folder": str(Path(output_folder).resolve())}


@mcp.tool()
def check_report(report_folder: str) -> dict:
    """Verify a report's output against its input: figures, captions, tables, index pages."""
    return verify_report(report_folder)


if __name__ == "__main__":
    mcp.run(transport="stdio")

r"""Wochenbuch generator: fills the school's Wochenbuch Vorlage and writes DOCX and PDF.

Separate from ``headless_core.py`` on purpose: the Laborbericht generator is tied to the
Laborbericht Vorlage (indexes, Eidesstattliche Erklaerung, screenshots) and its checks must stay
as they are. This one handles the much simpler Wochenbuch Vorlage ("Wochenbuch DTS2.docx"):
a title page with Name/Klasse/Schuljahr, a table of contents and the headings
"Thema der Woche" (1) with "Zusammenfassung", "Aufgabe" and "Quellen" (1.x).

A Wochenbuch lives in a report folder like a Laborbericht:

    <report>/work/wochenbuch.json    content (see below)
    <report>/work/<Vorlage>.docx     the Vorlage from eduvidual
    <report>/work/*.png              figures used by the content
    <report>/final/                  DOCX and PDF

    python wochenbuch_core.py <report folder>

wochenbuch.json:
    template, output_basename, titel (title page, optional), name, klasse, schuljahr, kw,
    pages_min, pages_max (optional, default 15 / 20),
    blocks: list of
      {"type": "h1"|"h2"|"h3", "text": ...}
      {"type": "p", "text": ...}                      text may contain {abb:id} and {tab:id}
      {"type": "bullets"|"numbered", "items": [...]}
      {"type": "formula", "text": ...}                centered, bold
      {"type": "table", "id": ..., "caption": ..., "header": [...], "rows": [[...]], "widths_cm": [...]}
      {"type": "figure", "id": ..., "file": ..., "caption": ..., "width_cm": 12}

Works in the normal mode and in the locked-down mode (see headless_core.MODE).
"""

from __future__ import annotations

import copy
import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import headless_core as core
from headless_core import (Cm, Document, OxmlElement, PdfReader, Pt, WD_ALIGN_PARAGRAPH,
                           WD_TABLE_ALIGNMENT, qn)

WORK_DIR, FINAL_DIR, BUILD_PDF = "work", "final", "_build.pdf"

# Characters allowed besides plain ASCII: the German letters.
ALLOWED_NON_ASCII = set("äöüÄÖÜß")

# Phrases that make a text sound machine-written. Their presence fails the build.
BANNED_PHRASES = [
    "zusammenfassend", "abschließend", "es ist wichtig zu beachten", "es ist wichtig zu betonen",
    "in der heutigen", "darüber hinaus", "des Weiteren", "nicht zuletzt", "unerlässlich",
    "essenziell", "von entscheidender Bedeutung", "spielt eine entscheidende Rolle",
    "ein breites Spektrum", "tauchen wir", "im Folgenden werden wir", "lass uns", "Fazit",
    "ganzheitlich", "nahtlos", "beleuchten", "Zusammenspiel",
]

TEMPLATE_LEFTOVERS = ["diesen Text ersetzen", "Nicht vergessen", "kurze Einleitung",
                      "Text, ....", "<Zusammenfassung", "<Aufgabenstellung", "DTS 1"]


def _w(el, tag):
    return el.iter(qn(tag))


def _para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _set_label(body, label: str, value: str) -> int:
    """Append ``value`` to every paragraph that consists only of ``label`` (title page text box)."""
    done = 0
    for p in list(body.iter(qn("w:p"))):
        if _para_text(p).strip() != label:
            continue
        runs = [r for r in p.iter(qn("w:r")) if r.find(qn("w:t")) is not None]
        if not runs:
            continue
        new = copy.deepcopy(runs[-1])
        for t in list(new.findall(qn("w:t")))[1:]:
            new.remove(t)
        t = new.find(qn("w:t"))
        t.text = " " + value
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        runs[-1].addnext(new)
        done += 1
    return done


def _replace_paragraph_text(p, new_text: str) -> None:
    ts = list(p.iter(qn("w:t")))
    if not ts:
        return
    ts[0].text = new_text
    for t in ts[1:]:
        t.text = ""


def _prepare_template(doc: Document, cfg: Dict) -> None:
    """Fill the title page, drop the template's hint and example text, keep the page break."""
    body = doc.element.body
    first = body[0]
    count = 0
    for key, label in (("name", "Name:"), ("klasse", "Klasse:"), ("schuljahr", "Schuljahr:")):
        count += _set_label(first, label, cfg[key])
    if count < 3:
        raise RuntimeError("Vorlage geaendert: Name/Klasse/Schuljahr auf der Titelseite nicht gefunden")
    if cfg.get("titel"):
        for p in first.iter(qn("w:p")):
            if _para_text(p).strip() == "Wochenbuch DTS 1":
                _replace_paragraph_text(p, cfg["titel"])

    # The yellow hint paragraph also carries the page break after the table of contents:
    # remove only its text run.
    note = None
    for p in body.findall(qn("w:p")):
        if "Nicht vergessen" in _para_text(p):
            note = p
            break
    if note is None:
        raise RuntimeError("Vorlage geaendert: Hinweisabsatz nach dem Inhaltsverzeichnis nicht gefunden")
    for r in list(note.findall(qn("w:r"))):
        if r.find(qn("w:t")) is not None:
            note.remove(r)
    if not any(br.get(qn("w:type")) == "page" for br in note.iter(qn("w:br"))):
        raise RuntimeError("Vorlage geaendert: Seitenumbruch nach dem Inhaltsverzeichnis fehlt")

    # Everything after the hint paragraph (the example text) is replaced by the content.
    after = False
    for el in list(body):
        if el is note:
            after = True
            continue
        if after and el.tag != qn("w:sectPr"):
            body.remove(el)


def _shade(cell, fill: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _table_borders(table) -> None:
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "808080")
        borders.append(el)
    tblPr.append(borders)


def _keep_with_next(paragraph) -> None:
    paragraph.paragraph_format.keep_with_next = True


def _number_refs(blocks: List[Dict]) -> Dict[str, str]:
    refs, abb, tab = {}, 0, 0
    for b in blocks:
        if b["type"] == "figure":
            abb += 1
            refs[f"abb:{b['id']}"] = f"Abbildung {abb}"
        elif b["type"] == "table":
            tab += 1
            refs[f"tab:{b['id']}"] = f"Tabelle {tab}"
    return refs


def _resolve(text: str, refs: Dict[str, str]) -> str:
    def sub(m):
        key = m.group(1)
        if key not in refs:
            raise KeyError(f"Verweis {{{key}}} hat kein Ziel")
        return refs[key]
    return re.sub(r"\{((?:abb|tab):[A-Za-z0-9_]+)\}", sub, text)


def heading_entries(blocks: List[Dict]) -> List[Dict]:
    """Numbered headings in document order: block index, level, number ("1.1.2"), text, anchor."""
    out, counters = [], [0, 0, 0]
    for i, b in enumerate(blocks):
        if b["type"] in ("h1", "h2", "h3"):
            level = int(b["type"][1])
            counters[level - 1] += 1
            for k in range(level, 3):
                counters[k] = 0
            out.append({"index": i, "level": level, "number": ".".join(str(c) for c in counters[:level]),
                        "text": b["text"], "anchor": f"_TocWb{len(out) + 1:03d}"})
    return out


def _add_bookmark(paragraph, anchor: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), anchor)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    p = paragraph._p
    first_run = p.find(qn("w:r"))
    if first_run is not None:
        first_run.addprevious(start)
    else:
        p.append(start)
    p.append(end)


def _toc_entry_xml(style: str, entry: Dict, page: str, first: bool) -> str:
    from xml.sax.saxutils import escape
    ns = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')
    fld = lambda t: f'<w:r><w:fldChar w:fldCharType="{t}"/></w:r>'
    head = ""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/>' + ('<w:tabs><w:tab w:val="left" w:pos="851"/></w:tabs>' if first else "") + "</w:pPr>"
    if first:
        head = (fld("begin") + '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
                + fld("separate"))
    link = (f'<w:hyperlink w:anchor="{entry["anchor"]}" w:history="1">'
            f'<w:r><w:t>{escape(entry["number"])}</w:t></w:r><w:r><w:tab/></w:r>'
            f'<w:r><w:t xml:space="preserve">{escape(entry["text"])}</w:t></w:r><w:r><w:tab/></w:r>'
            + fld("begin") + f'<w:r><w:instrText xml:space="preserve"> PAGEREF {entry["anchor"]} \\h </w:instrText></w:r>'
            + fld("separate") + f'<w:r><w:t>{escape(page)}</w:t></w:r>' + fld("end") + "</w:hyperlink>")
    return f"<w:p {ns}>{ppr}{head}{link}</w:p>"


def _populate_toc(doc: Document, entries: List[Dict], pages: Dict[str, int]) -> None:
    """Replace the Vorlage's example entries with the real ones (cached field result)."""
    from docx.oxml import parse_xml
    body = doc.element.body
    toc = None
    for sdt in body.iter(qn("w:sdt")):
        gallery = sdt.find(".//" + qn("w:docPartGallery"))
        if gallery is not None and gallery.get(qn("w:val")) == "Table of Contents":
            toc = sdt
            break
    if toc is None:
        raise RuntimeError("Vorlage geaendert: Inhaltsverzeichnis nicht gefunden")
    content = toc.find(qn("w:sdtContent"))
    paras = content.findall(qn("w:p"))
    if len(paras) < 3:
        raise RuntimeError("Vorlage geaendert: Inhaltsverzeichnis hat keine Eintraege")
    title, closing = paras[0], paras[-1]
    for p in paras[1:-1]:
        content.remove(p)
    anchor_after = title
    for n, e in enumerate(entries):
        el = parse_xml(_toc_entry_xml(f"Verzeichnis{e['level']}", e, str(pages.get(e["anchor"], 1)), n == 0))
        anchor_after.addnext(el)
        anchor_after = el


def _compact_toc_styles(doc: Document) -> None:
    for name in ("toc 1", "toc 2", "toc 3"):
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        st.paragraph_format.line_spacing = 1.0
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.space_after = Pt(2)


def _build_document(folder: Path, cfg: Dict, pages: Dict[str, int] | None = None) -> Document:
    work = folder / WORK_DIR
    doc = Document(str(work / cfg["template"]))
    _prepare_template(doc, cfg)
    refs = _number_refs(cfg["blocks"])
    fig_n = tab_n = 0
    entries = heading_entries(cfg["blocks"])
    by_index = {e["index"]: e for e in entries}
    _populate_toc(doc, entries, pages or {})
    _compact_toc_styles(doc)

    for idx, b in enumerate(cfg["blocks"]):
        kind = b["type"]
        if kind in ("h1", "h2", "h3"):
            para = doc.add_paragraph(_resolve(b["text"], refs), style=f"Heading {kind[1]}")
            _add_bookmark(para, by_index[idx]["anchor"], 1000 + idx)
        elif kind == "p":
            doc.add_paragraph(_resolve(b["text"], refs))
        elif kind in ("bullets", "numbered"):
            style = "List Bullet" if kind == "bullets" else "List Number"
            for n_item, item in enumerate(b["items"]):
                para = doc.add_paragraph(_resolve(item, refs), style=style)
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                para.paragraph_format.line_spacing = 1.15
                last = n_item == len(b["items"]) - 1
                para.paragraph_format.space_after = Pt(10 if last else 2)
        elif kind == "formula":
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(_resolve(b["text"], refs)).bold = True
        elif kind == "table":
            tab_n += 1
            cap = doc.add_paragraph(f"Tabelle {tab_n}: {_resolve(b['caption'], refs)}", style="Caption")
            _keep_with_next(cap)
            header, rows, widths = b["header"], b["rows"], b["widths_cm"]
            table = doc.add_table(rows=1 + len(rows), cols=len(header))
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            _table_borders(table)
            # Fixed layout with an explicit grid, otherwise LibreOffice shrinks the table.
            tblPr = table._tbl.tblPr
            for el in tblPr.findall(qn("w:tblW")):
                tblPr.remove(el)
            tblW = OxmlElement("w:tblW")
            tblW.set(qn("w:w"), str(int(round(sum(widths) * 567))))
            tblW.set(qn("w:type"), "dxa")
            tblPr.append(tblW)
            layout = OxmlElement("w:tblLayout")
            layout.set(qn("w:type"), "fixed")
            tblPr.append(layout)
            for c_i, w_cm in enumerate(widths):
                table.columns[c_i].width = Cm(w_cm)
            for r_i, row in enumerate([header] + rows):
                if len(row) != len(header):
                    raise ValueError(f"Tabelle {b['id']}: Zeile {r_i} hat {len(row)} statt {len(header)} Spalten")
                for c_i, text in enumerate(row):
                    cell = table.cell(r_i, c_i)
                    cell.width = Cm(widths[c_i])
                    para = cell.paragraphs[0]
                    para.paragraph_format.space_after = Pt(2)
                    para.paragraph_format.space_before = Pt(2)
                    para.paragraph_format.line_spacing = 1.0
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = para.add_run(_resolve(text, refs))
                    run.font.size = Pt(10.5)
                    if r_i == 0:
                        run.bold = True
                        _shade(cell, "D9E2EA")
                # A table may break between rows, but never inside a row, and the header row
                # always stays with the first data row.
                trPr = table.rows[r_i]._tr.get_or_add_trPr()
                trPr.append(OxmlElement("w:cantSplit"))
                if r_i == 0:
                    for cell in table.rows[0].cells:
                        for para in cell.paragraphs:
                            para.paragraph_format.keep_with_next = True
            doc.add_paragraph().paragraph_format.space_after = Pt(4)
        elif kind == "figure":
            fig_n += 1
            pic = doc.add_paragraph()
            pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _keep_with_next(pic)
            pic.add_run().add_picture(str(work / b["file"]), width=Cm(b.get("width_cm", 12)))
            cap = doc.add_paragraph(f"Abbildung {fig_n}: {_resolve(b['caption'], refs)}", style="Caption")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            raise ValueError(f"unbekannter Blocktyp: {kind}")

    props = doc.core_properties
    props.author = cfg["name"]
    props.title = cfg.get("titel") or "Wochenbuch"
    return doc


def _all_text(cfg: Dict) -> str:
    parts = []
    for b in cfg["blocks"]:
        for key in ("text", "caption"):
            if key in b:
                parts.append(b[key])
        parts += b.get("items", []) + b.get("header", [])
        for row in b.get("rows", []):
            parts += row
    return "\n".join(parts)


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def first_body_page(pages: List[str]) -> int:
    """Index of the first page after the table of contents (no dot leaders in its text)."""
    for i in range(1, len(pages)):
        if "....." not in pages[i] and "Inhalt" not in pages[i][:40]:
            return i
    return len(pages)


def measure_pages(pdf_path: Path, entries: List[Dict]) -> Dict[str, int]:
    """Page number (1-based) of every numbered heading, searched in document order."""
    texts = [_norm(pg.extract_text() or "") for pg in PdfReader(str(pdf_path)).pages]
    page, result = first_body_page(texts), {}
    for e in entries:
        needle = _norm(f"{e['number']} {e['text']}")
        for i in range(page, len(texts)):
            if needle in texts[i]:
                page = i
                result[e["anchor"]] = i + 1
                break
    return result


def verify(folder: Path, cfg: Dict, docx_path: Path, pdf_path: Path, toc_pages: Dict[str, int] | None = None) -> Dict:
    checks: Dict[str, bool] = {}
    notes: Dict[str, object] = {}
    text = _all_text(cfg)

    bad = sorted({c for c in text if ord(c) > 127 and c not in ALLOWED_NON_ASCII})
    checks["only_ascii_plus_umlauts"] = not bad
    if bad:
        notes["non_ascii"] = [f"U+{ord(c):04X} {c}" for c in bad]
    low = text.lower()
    hits = [p for p in BANNED_PHRASES if p.lower() in low]
    checks["no_stock_phrases"] = not hits
    if hits:
        notes["stock_phrases"] = hits
    checks["no_double_dash_or_arrow"] = not re.search(r"--|->|=>", text)

    used = {b["file"] for b in cfg["blocks"] if b["type"] == "figure"}
    work = folder / WORK_DIR
    checks["figure_files_exist"] = all((work / f).is_file() for f in used)
    checks["all_figures_used"] = not (set(p.name for p in work.glob("fig_*.png")) - used)

    if not pdf_path.exists():
        checks["pdf_exists"] = False
        return {"checks": checks, "notes": notes, "pages": None}
    reader = PdfReader(str(pdf_path))
    pages = [(pg.extract_text() or "") for pg in reader.pages]
    n = len(pages)
    lo, hi = cfg.get("pages_min", 15), cfg.get("pages_max", 20)
    checks["page_count_in_range"] = lo <= n <= hi
    notes["page_range"] = f"{lo}-{hi}"

    joined = "\n".join(pages)
    checks["no_template_leftovers"] = not any(s in joined for s in TEMPLATE_LEFTOVERS)
    first = pages[0] if pages else ""
    checks["title_page_filled"] = all(v in first for v in (cfg["name"], cfg["klasse"], cfg["schuljahr"]))
    if cfg.get("titel"):  # the Vorlage prints the title in capitals
        checks["title_page_filled"] = checks["title_page_filled"] and cfg["titel"].lower() in first.lower()

    entries = heading_entries(cfg["blocks"])
    body_start = first_body_page(pages)
    body_text = _norm("\n".join(pages[body_start:]))
    toc = _norm("\n".join(pages[1:body_start]))
    missing = [f"{e['number']} {e['text']}" for e in entries if _norm(f"{e['number']} {e['text']}") not in body_text]
    checks["headings_numbered_in_pdf"] = not missing
    if missing:
        notes["headings_missing"] = missing
    checks["toc_lists_all_headings"] = all(_norm(e["text"]) in toc for e in entries)
    measured = measure_pages(pdf_path, entries)
    checks["toc_page_numbers_correct"] = bool(toc_pages) and measured == toc_pages and len(measured) == len(entries)
    notes["toc_pages"] = f"{body_start - 1} Seite(n) Inhaltsverzeichnis"

    thin = [i + 1 for i, t in enumerate(pages) if i >= 2 and len(t.strip()) < 150]
    checks["no_nearly_empty_pages"] = not thin
    if thin:
        notes["thin_pages"] = thin
    return {"checks": checks, "notes": notes, "pages": n}


def build(folder_arg: str) -> Dict:
    folder = Path(folder_arg).resolve()
    work, final = folder / WORK_DIR, folder / FINAL_DIR
    cfg = json.loads((work / "wochenbuch.json").read_text(encoding="utf-8"))
    final.mkdir(exist_ok=True)
    docx_path = work / f"{cfg['output_basename']}.docx"
    pdf_path = work / BUILD_PDF

    entries = heading_entries(cfg["blocks"])
    pages_used: Dict[str, int] = {}
    exported, message = False, "nicht ausgefuehrt"
    session = core._export_session()
    try:
        # Page numbers can only be measured on a rendered PDF: build, measure, build again
        # with the numbers, until a pass reproduces the numbers it was built with.
        for _ in range(4):
            _build_document(folder, cfg, pages_used).save(str(docx_path))
            if pdf_path.exists():
                pdf_path.unlink()
            exported, message = session.export(str(docx_path), str(pdf_path))
            if not exported or not pdf_path.exists():
                break
            measured = measure_pages(pdf_path, entries)
            if measured == pages_used:
                break
            pages_used = measured
    finally:
        session.close()

    result = verify(folder, cfg, docx_path, pdf_path, pages_used)
    checks = result["checks"]
    checks["pdf_exported"] = exported and pdf_path.exists()

    written, final_message = False, "nicht geschrieben, weil kein PDF erzeugt wurde"
    if pdf_path.exists():
        f_docx, f_pdf = final / docx_path.name, final / f"{cfg['output_basename']}.pdf"
        try:
            if core._locked(f_docx) or core._locked(f_pdf):
                raise PermissionError
            shutil.copyfile(docx_path, f_docx)
            shutil.copyfile(pdf_path, f_pdf)
            pdf_path.unlink()
            written, final_message = True, "geschrieben"
        except PermissionError:
            final_message = "Datei in final/ ist geoeffnet; neuer Stand liegt in work/"
    checks["final_files_written"] = written

    manifest = {"generated_at": datetime.now().isoformat(timespec="seconds"), "mode": core.MODE,
                "passed": all(checks.values()), "pages": result["pages"], "final": final_message,
                "export_message": message, "checks": checks, "notes": result["notes"]}
    (work / "report.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Aufruf: wochenbuch_core.py <report_folder>")
    m = build(sys.argv[1])
    summary = {"passed": m["passed"], "pages": m["pages"], "final": m["final"],
               "failed_checks": [k for k, v in m["checks"].items() if not v], "notes": m["notes"]}
    sys.stdout.buffer.write(json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8"))
    sys.exit(0 if m["passed"] else 1)

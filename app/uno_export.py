r"""PDF export through LibreOffice's UNO bridge. Standard library and ``uno`` only.

``python-docx`` cannot evaluate fields, so the table of contents has to be refreshed by
LibreOffice itself before the PDF is written. That needs the ``uno`` module, which only exists
in the Python that belongs to LibreOffice:

* Locked-down mode (school machines): the generator already runs in LibreOffice's Python, so
  ``UnoSession`` is used in the same process.
* Normal mode: the generator runs in an ordinary Python or venv without ``uno``. It then starts
  this file as a helper process with a Python that has it (``WorkerSession``):

      <uno python> uno_export.py --worker <soffice> [timeout]

  The helper reads one JSON object per line on stdin (``{"docx": ..., "pdf": ...}``) and answers
  each with one JSON line (``{"ok": ..., "message": ...}``) until stdin is closed.

Both ways drive the same LibreOffice with the same calls, so the PDF is the same.
"""

from __future__ import annotations

import os
import sys
import json
import time
import tempfile
import subprocess
from pathlib import Path
from typing import Tuple


class UnoSession:
    """A headless soffice with a UNO socket, started on first use and reused until ``close()``.

    Starting LibreOffice takes far longer than exporting a document, so one build keeps a
    single instance for all of its layout passes.
    """

    PORT = 2002

    def __init__(self, soffice: Path, timeout: float = 90.0) -> None:
        self.soffice = Path(soffice)
        self.timeout = timeout
        self._process = None
        self._desktop = None
        self._uno = None
        self._property_value = None

    def _connect(self) -> Tuple[bool, str]:
        if self._desktop is not None:
            return True, ""
        try:
            import uno
            from com.sun.star.beans import PropertyValue
        except Exception as exc:
            return False, f"UNO nicht verfügbar: {exc}"
        if not self.soffice.exists():
            return False, f"LibreOffice nicht gefunden: {self.soffice}"

        profile = Path(os.environ.get("TEMP") or tempfile.gettempdir()).absolute() / "loprofile_indexrefresh"
        profile.mkdir(parents=True, exist_ok=True)
        cmd = [
            str(self.soffice),
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless", "--norestore", "--nologo", "--nodefault", "--nofirststartwizard",
            f"--accept=socket,host=localhost,port={self.PORT};urp;StarOffice.ComponentContext",
        ]
        self._process = subprocess.Popen(cmd, stdin=subprocess.DEVNULL,
                                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        try:
            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        except Exception as exc:
            self.close()
            return False, f"UNO nicht verfügbar: {exc}"

        ctx = None
        deadline = time.time() + self.timeout
        while time.time() < deadline:
            try:
                ctx = resolver.resolve(
                    f"uno:socket,host=localhost,port={self.PORT};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.25)

        if ctx is None:
            self._process.terminate()
            self._process = None
            return False, "Keine Verbindung zum UNO-Bridge (Zeitlimit erreicht)"

        self._desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        self._uno = uno
        self._property_value = PropertyValue
        return True, ""

    def export(self, docx_path: str, pdf_path: str) -> Tuple[bool, str]:
        """Load the DOCX, refresh every index and store it as PDF. The DOCX is not written back."""
        connected, message = self._connect()
        if not connected:
            return False, message

        def prop(name: str, value):
            item = self._property_value()
            item.Name = name
            item.Value = value
            return item

        uno = self._uno
        document = None
        try:
            document = self._desktop.loadComponentFromURL(
                uno.systemPathToFileUrl(str(Path(docx_path).resolve())), "_blank", 0, (prop("Hidden", True),))

            try:
                document.getTextFields().refresh()
            except Exception:
                pass

            indexes = document.getDocumentIndexes()
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()

            try:
                document.refresh()
            except Exception:
                pass

            document.storeToURL(
                uno.systemPathToFileUrl(str(Path(pdf_path).resolve())),
                (prop("FilterName", "writer_pdf_Export"),))
            return True, f"{indexes.getCount()} Verzeichnisse aktualisiert, PDF exportiert"
        except Exception as exc:
            return False, f"UNO-Export fehlgeschlagen: {type(exc).__name__}: {exc}"
        finally:
            if document is not None:
                try:
                    document.close(True)
                except Exception:
                    pass

    def close(self) -> None:
        desktop, process = self._desktop, self._process
        self._desktop = self._process = None
        if desktop is not None:
            try:
                desktop.terminate()
            except Exception:
                pass
        if process is not None:
            try:
                process.wait(timeout=20)
            except Exception:
                process.terminate()


class WorkerSession:
    """Same interface as ``UnoSession``, but the UNO calls run in a helper process.

    Used when the generator's own Python has no ``uno`` module (normal mode).
    """

    def __init__(self, uno_python: Path, soffice: Path, timeout: float = 90.0) -> None:
        self.uno_python = Path(uno_python)
        self.soffice = Path(soffice)
        self.timeout = timeout
        self._worker = None

    def export(self, docx_path: str, pdf_path: str) -> Tuple[bool, str]:
        if self._worker is None:
            self._worker = subprocess.Popen(
                [str(self.uno_python), str(Path(__file__).resolve()), "--worker",
                 str(self.soffice), str(self.timeout)],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
        request = {"docx": str(Path(docx_path).resolve()), "pdf": str(Path(pdf_path).resolve())}
        try:
            self._worker.stdin.write(json.dumps(request).encode("ascii") + b"\n")
            self._worker.stdin.flush()
            answer = json.loads(self._worker.stdout.readline().decode("utf-8"))
            return bool(answer["ok"]), str(answer["message"])
        except Exception as exc:
            return False, f"UNO-Hilfsprozess ({self.uno_python}) antwortet nicht: {type(exc).__name__}: {exc}"

    def close(self) -> None:
        worker, self._worker = self._worker, None
        if worker is None:
            return
        try:
            worker.stdin.close()  # ends the helper's loop; it then shuts LibreOffice down
        except Exception:
            pass
        try:
            worker.wait(timeout=40)
        except Exception:
            worker.kill()
        try:
            worker.stdout.close()
        except Exception:
            pass


class UnavailableSession:
    """Stands in when no Python with ``uno`` exists: every export fails with the reason."""

    def __init__(self, message: str) -> None:
        self.message = message

    def export(self, docx_path: str, pdf_path: str) -> Tuple[bool, str]:
        return False, self.message

    def close(self) -> None:
        pass


def _serve(soffice: str, timeout: float) -> None:
    session = UnoSession(Path(soffice), timeout)
    out = sys.stdout.buffer
    try:
        for line in sys.stdin.buffer:
            if not line.strip():
                continue
            try:
                request = json.loads(line.decode("utf-8"))
                ok, message = session.export(request["docx"], request["pdf"])
            except Exception as exc:
                ok, message = False, f"UNO-Export fehlgeschlagen: {type(exc).__name__}: {exc}"
            out.write(json.dumps({"ok": ok, "message": message}).encode("ascii") + b"\n")
            out.flush()
    finally:
        session.close()


if __name__ == "__main__":
    if len(sys.argv) < 3 or sys.argv[1] != "--worker":
        sys.exit("Aufruf: uno_export.py --worker <soffice> [timeout]")
    _serve(sys.argv[2], float(sys.argv[3]) if len(sys.argv) > 3 else 90.0)

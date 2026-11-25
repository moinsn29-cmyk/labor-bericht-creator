"""Check the environment for normal mode and print the paths the agent needs.

    python app/check_env.py            check only
    python app/check_env.py --setup    create .venv in the bundle folder and install requirements.txt

Normal mode means: an ordinary Python with the packages from requirements.txt, normally the
bundle's own ``.venv``. On machines without Python (school machines) use ``check_env.ps1``
instead, which runs everything with LibreOffice's Python (locked-down mode).

Standard library only, so it runs before anything is installed.
"""

from __future__ import annotations

import os
import sys
import subprocess
from pathlib import Path

APP = Path(__file__).resolve().parent
BUNDLE = APP.parent
VENV = BUNDLE / ".venv"
VENV_PYTHON = VENV / ("Scripts/python.exe" if os.name == "nt" else "bin/python")

PROBE = (
    "import sys; sys.path.insert(0, sys.argv[1]); import headless_core as h; e = h.environment(); "
    "print('MODE=' + e['mode']); print('DOCUMENTS=' + e['documents']); print('IMPORTS=OK'); "
    "print('SOFFICE=' + e['soffice']); print('SOFFICE_FOUND=' + str(e['soffice_found'])); "
    "print('UNO_PYTHON=' + str(e['uno_python'] or 'NICHT GEFUNDEN'))"
)


def probe(python: Path) -> subprocess.CompletedProcess:
    return subprocess.run([str(python), "-c", PROBE, str(APP)], capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def setup() -> None:
    if not VENV_PYTHON.is_file():
        print(f"Lege {VENV} an ...", flush=True)
        subprocess.run([sys.executable, "-m", "venv", str(VENV)], check=True)
    print("Installiere requirements.txt ...", flush=True)
    subprocess.run([str(VENV_PYTHON), "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
                    "-r", str(BUNDLE / "requirements.txt")], check=True)


def main() -> int:
    if "--setup" in sys.argv[1:]:
        try:
            setup()
        except (OSError, subprocess.CalledProcessError) as exc:
            print(f"SETUP=FEHLER ({exc})")
            print("Abhilfe: Internetverbindung prüfen; unter Debian/Ubuntu fehlt eventuell das Paket python3-venv.")
            return 1

    # Prefer the bundle's venv; otherwise the Python running this script.
    candidates = [p for p in (VENV_PYTHON, Path(sys.executable)) if p.is_file()]
    result = None
    for python in candidates:
        result = probe(python)
        if result.returncode == 0:
            print(f"PYTHON={python}")
            print(result.stdout.strip())
            if "SOFFICE_FOUND=False" in result.stdout:
                print("LibreOffice wurde nicht gefunden. Installieren oder LIBREOFFICE_PROGRAM auf den "
                      "Ordner mit soffice setzen.")
                return 1
            if "UNO_PYTHON=NICHT GEFUNDEN" in result.stdout:
                print("Kein Python mit dem Modul 'uno' gefunden (gehört zu LibreOffice). Linux: Paket "
                      "python3-uno installieren. Sonst LIBREOFFICE_PYTHON auf LibreOffices Python setzen.")
                return 1
            return 0

    print("IMPORTS=FEHLER")
    if result is not None and result.stderr.strip():
        print(result.stderr.strip().splitlines()[-1])
    print(f"Abhilfe: \"{sys.executable}\" \"{Path(__file__).resolve()}\" --setup")
    return 1


if __name__ == "__main__":
    sys.exit(main())

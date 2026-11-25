"""Fetch the material of one eduvidual (Moodle) course section into a folder.

Standard library only, so it runs with the normal Python and with LibreOffice's
Python 3.11 of the locked-down mode.

    python3 app/fetch_eduvidual.py <section-url-or-id> [--out eingabe/<name>]
        [--user <login>] [--password-file <file>] [--vorlage <regex>]

The password is read from a file (default ``~/eduvidualpw``) or, if no file is
given and the file does not exist, from the environment variable
``EDUVIDUAL_PASSWORD``. It is never printed, never passed on a command line and
never stored. The login name comes from ``--user``, ``EDUVIDUAL_USER`` or
``profil.json`` (field ``eduvidual_user``). Nothing here is ever committed:
``eingabe/*`` is ignored by git.

Written into the output folder:
  <section files>                 PDFs, documents and so on of the section
  seite_<name>.html / .txt        Moodle pages of the section (HTML and text)
  links.txt                       external links of the section
  aufgaben.txt                    text and attachments of the section's assignments
  abschnitt.txt                   section title and summary
  vorlage/<file>                  files elsewhere in the course whose module or
                                  file name matches --vorlage (default
                                  "vorlage|wochenbuch")
"""
from __future__ import annotations

import argparse
import html
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

BASE = "https://www.eduvidual.at"
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(APP_DIR)


class FetchError(Exception):
    pass


def _read_json(url: str, data: bytes | None = None):
    try:
        with urllib.request.urlopen(urllib.request.Request(url, data=data), timeout=60) as r:
            return json.load(r)
    except urllib.error.URLError as e:
        raise FetchError(f"Keine Verbindung zu eduvidual: {e.reason}") from None


def get_password(password_file: str | None) -> str:
    candidates = [password_file] if password_file else [os.path.expanduser("~/eduvidualpw")]
    for p in candidates:
        if p and os.path.isfile(p):
            with open(p, encoding="utf-8") as f:
                return f.read().strip()
    env = os.environ.get("EDUVIDUAL_PASSWORD")
    if env:
        return env
    raise FetchError("Kein Passwort: --password-file angeben oder EDUVIDUAL_PASSWORD setzen.")


def get_user(cli_user: str | None) -> str:
    if cli_user:
        return cli_user
    if os.environ.get("EDUVIDUAL_USER"):
        return os.environ["EDUVIDUAL_USER"]
    try:
        with open(os.path.join(ROOT, "profil.json"), encoding="utf-8") as f:
            u = json.load(f).get("eduvidual_user")
        if u:
            return u
    except (OSError, ValueError):
        pass
    raise FetchError("Kein Benutzername: --user angeben, EDUVIDUAL_USER setzen oder "
                     "\"eduvidual_user\" in profil.json eintragen.")


def get_token(user: str, password: str) -> str:
    data = urllib.parse.urlencode(
        {"username": user, "password": password, "service": "moodle_mobile_app"}).encode()
    r = _read_json(BASE + "/login/token.php", data)
    if "token" not in r:
        raise FetchError(f"Anmeldung fehlgeschlagen: {r.get('error') or r.get('errorcode') or r}")
    return r["token"]


class Moodle:
    def __init__(self, token: str):
        self.token = token

    def call(self, fn: str, **params):
        q = urllib.parse.urlencode(
            {"wstoken": self.token, "wsfunction": fn, "moodlewsrestformat": "json", **params})
        r = _read_json(BASE + "/webservice/rest/server.php?" + q)
        if isinstance(r, dict) and r.get("exception"):
            raise FetchError(f"{fn}: {r.get('message')}")
        return r

    def download(self, fileurl: str, dest: str) -> None:
        sep = "&" if "?" in fileurl else "?"
        url = fileurl + sep + "token=" + urllib.parse.quote(self.token)
        try:
            with urllib.request.urlopen(url, timeout=120) as r, open(dest, "wb") as f:
                while True:
                    chunk = r.read(1 << 16)
                    if not chunk:
                        break
                    f.write(chunk)
        except urllib.error.URLError as e:
            raise FetchError(f"Download fehlgeschlagen ({os.path.basename(dest)}): {e}") from None


def safe_name(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or "datei"


def html_to_text(src: str) -> str:
    src = re.sub(r"(?is)<(script|style).*?</\1>", "", src)
    src = re.sub(r"(?i)<br\s*/?>|</(p|div|li|h[1-6]|tr)>", "\n", src)
    src = re.sub(r"(?i)<li[^>]*>", "- ", src)
    src = re.sub(r"<[^>]+>", "", src)
    src = html.unescape(src)
    return re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t]+", " ", src)).strip() + "\n"


def section_id_from(arg: str) -> int:
    m = re.search(r"[?&]id=(\d+)", arg) or re.fullmatch(r"(\d+)", arg.strip())
    if not m:
        raise FetchError("Abschnitt-URL oder ID nicht erkannt (erwartet: …/course/section.php?id=NNN).")
    return int(m.group(1))


def find_section(moodle: Moodle, section_id: int):
    uid = moodle.call("core_webservice_get_site_info")["userid"]
    for c in moodle.call("core_enrol_get_users_courses", userid=uid):
        try:
            secs = moodle.call("core_course_get_contents", courseid=c["id"])
        except FetchError:
            continue
        for s in secs:
            if s.get("id") == section_id:
                return c, s, secs
    raise FetchError("Abschnitt in keinem deiner Kurse gefunden (eingeschrieben? Link richtig?).")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("section", help="Abschnitt-URL (…/course/section.php?id=NNN) oder ID")
    ap.add_argument("--out", help="Zielordner (Standard: eingabe/<Abschnittsname>)")
    ap.add_argument("--user")
    ap.add_argument("--password-file")
    ap.add_argument("--vorlage", default=r"vorlage|wochenbuch",
                    help="Regex: Dateien im restlichen Kurs, die als Vorlage geholt werden")
    args = ap.parse_args()

    try:
        token = get_token(get_user(args.user), get_password(args.password_file))
        moodle = Moodle(token)
        course, section, all_sections = find_section(moodle, section_id_from(args.section))
    except FetchError as e:
        print("FEHLER:", e, file=sys.stderr)
        return 1

    out = args.out or os.path.join(ROOT, "eingabe", safe_name(section.get("name") or "abschnitt"))
    os.makedirs(out, exist_ok=True)
    print(f"KURS={course['fullname']}")
    print(f"ABSCHNITT={section.get('name')}")
    print(f"OUT={out}")

    with open(os.path.join(out, "abschnitt.txt"), "w", encoding="utf-8") as f:
        f.write(f"Kurs: {course['fullname']}\nAbschnitt: {section.get('name')}\n\n")
        f.write(html_to_text(section.get("summary") or ""))

    links, tasks = [], []
    assigns = {}
    try:
        for c in moodle.call("mod_assign_get_assignments", **{"courseids[0]": course["id"]})["courses"]:
            for a in c["assignments"]:
                assigns[a["cmid"]] = a
    except FetchError:
        pass

    for m in section.get("modules", []):
        if not m.get("uservisible", True):
            continue
        kind, name = m["modname"], m["name"]
        if kind in ("resource", "folder"):
            for c in m.get("contents") or []:
                if c.get("type") == "file":
                    dest = os.path.join(out, safe_name(c["filename"]))
                    moodle.download(c["fileurl"], dest)
                    print(f"DATEI={dest}")
        elif kind == "page":
            for c in m.get("contents") or []:
                if c.get("filename") == "index.html":
                    base = os.path.join(out, "seite_" + safe_name(name))
                    moodle.download(c["fileurl"], base + ".html")
                    with open(base + ".html", encoding="utf-8", errors="replace") as f:
                        text = html_to_text(f.read())
                    with open(base + ".txt", "w", encoding="utf-8") as f:
                        f.write(name + "\n\n" + text)
                    print(f"SEITE={base}.txt")
        elif kind == "url":
            for c in m.get("contents") or []:
                links.append(f"{name}\n  {c.get('fileurl')}")
        elif kind == "assign":
            a = assigns.get(m["id"])
            block = [f"Aufgabe: {name}"]
            if a:
                block.append(html_to_text(a.get("intro") or ""))
                for att in a.get("introattachments", []) or []:
                    dest = os.path.join(out, safe_name(att["filename"]))
                    moodle.download(att["fileurl"], dest)
                    block.append(f"Anhang: {att['filename']}")
                    print(f"DATEI={dest}")
                if a.get("duedate"):
                    block.append(f"Abgabe (Unix-Zeit): {a['duedate']}")
            tasks.append("\n".join(block))

    if links:
        with open(os.path.join(out, "links.txt"), "w", encoding="utf-8") as f:
            f.write("\n\n".join(links) + "\n")
        print(f"LINKS={len(links)}")
    if tasks:
        with open(os.path.join(out, "aufgaben.txt"), "w", encoding="utf-8") as f:
            f.write("\n\n".join(tasks) + "\n")

    pat = re.compile(args.vorlage, re.I)
    vdir = os.path.join(out, "vorlage")
    for s in all_sections:
        for m in s.get("modules", []):
            if m["modname"] not in ("resource", "folder") or not m.get("uservisible", True):
                continue
            for c in m.get("contents") or []:
                if c.get("type") == "file" and (pat.search(m["name"]) or pat.search(c["filename"])):
                    os.makedirs(vdir, exist_ok=True)
                    dest = os.path.join(vdir, safe_name(c["filename"]))
                    moodle.download(c["fileurl"], dest)
                    print(f"VORLAGE={dest}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

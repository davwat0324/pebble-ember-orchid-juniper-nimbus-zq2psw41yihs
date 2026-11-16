"""Local files, kept per Windows user.

%APPDATA%\\Remind\\
    settings.json     message templates and which is the default, Outlook signature on/off, window size,
                      Spectrum server / database / email view
    spectrum.json     Spectrum SQL login; the password is encrypted with Windows DPAPI (only this
                      Windows user on this PC can read it back)
    history.jsonl     one line per send: when, file, subject, recipients and the result for each
"""
from __future__ import annotations

import base64
import json
import os
from datetime import datetime
from pathlib import Path

DATA_DIR = Path(os.environ.get("APPDATA") or Path.home()) / "Remind"
SETTINGS_FILE = DATA_DIR / "settings.json"
CRED_FILE = DATA_DIR / "spectrum.json"
HISTORY_FILE = DATA_DIR / "history.jsonl"

DEFAULT_SUBJECT = "Reminder: action needed"
DEFAULT_BODY = (
    "<p>Hi {FirstName},</p>"
    "<p>This is a friendly reminder to <b>[describe what needs to be done]</b> by <b>[date]</b>.</p>"
    "<p>If you have already taken care of this, please disregard this message. "
    "Let me know if you have any questions.</p>"
    "<p>Thank you,</p>"
)


def _write(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, path)  # atomic: a crash never leaves a half-written file


def _read(path: Path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return default


# --- settings ---
def blank_template() -> dict:
    return {"subject": DEFAULT_SUBJECT, "body": DEFAULT_BODY, "cc": "", "bcc": ""}


def load_settings() -> dict:
    """Settings with defaults filled in.

    Message templates live in s["templates"] ({name: {subject, body, cc, bcc}}) and s["default_template"]
    names the one a new message starts from. Settings from version 1.0 had a single default message in
    the top-level subject/body/cc/bcc keys; those become a template called "Default".
    """
    s = _read(SETTINGS_FILE, {})
    s.setdefault("signature", True)
    if not s.get("templates"):
        old = {k: s.pop(k) for k in ("subject", "body", "cc", "bcc") if k in s}
        s["templates"] = {"Default": {**blank_template(), **old}}
    for k in ("subject", "body", "cc", "bcc"):
        s.pop(k, None)
    if s.get("default_template") not in s["templates"]:
        s["default_template"] = next(iter(s["templates"]))
    return s


def save_settings(s: dict) -> None:
    _write(SETTINGS_FILE, s)


# --- Spectrum login (password protected with DPAPI) ---
def _protect(text: str) -> str:
    import win32crypt
    blob = win32crypt.CryptProtectData(text.encode("utf-8"), "Remind", None, None, None, 0)
    return base64.b64encode(blob).decode("ascii")


def _unprotect(token: str) -> str:
    import win32crypt
    return win32crypt.CryptUnprotectData(base64.b64decode(token), None, None, None, 0)[1].decode("utf-8")


def load_credentials() -> dict | None:
    c = _read(CRED_FILE, None)
    if not (c and c.get("User") and c.get("Password")):
        return None
    try:
        return {"User": c["User"], "Password": _unprotect(c["Password"])}
    except Exception:
        return None  # e.g. the file was copied from another user or PC


def saved_user() -> str:
    return (_read(CRED_FILE, {}) or {}).get("User", "")


def save_credentials(user: str, password: str) -> None:
    _write(CRED_FILE, {"User": user, "Password": _protect(password)})


# --- send history ---
def add_history(entry: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    entry = {"when": datetime.now().isoformat(timespec="seconds"), **entry}
    with HISTORY_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def load_history() -> list[dict]:
    out = []
    try:
        for line in HISTORY_FILE.read_text(encoding="utf-8").splitlines():
            try:
                out.append(json.loads(line))
            except ValueError:
                pass  # skip a damaged line
    except OSError:
        pass
    return list(reversed(out))  # newest first

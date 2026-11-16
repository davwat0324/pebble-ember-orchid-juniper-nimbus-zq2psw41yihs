"""Self-update from the latest GitHub Release.

check() asks the GitHub API for the latest release. If its tag is newer than this version and it has
Remind-win64.zip attached, download() fetches it (checking the SHA-256 that GitHub publishes for the
asset) and unpacks it, and apply() starts a hidden PowerShell step that waits for Remind to close,
copies the new files over the install folder and starts Remind again. Settings, the saved login and
history live in %APPDATA%\\Remind and are not touched.

Only the installed .exe can update itself (sys.frozen); running from source just reports.
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path

from version import ASSET, REPO, __version__

API = f"https://api.github.com/repos/{REPO}/releases/latest"
UA = {"User-Agent": "Remind-updater", "Accept": "application/vnd.github+json"}


@dataclass
class Release:
    version: str
    url: str
    size: int
    sha256: str        # '' when GitHub didn't publish a digest
    notes: str


def _parse(v: str) -> tuple[int, ...]:
    parts = []
    for p in v.strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits or 0))
    return tuple(parts + [0] * (3 - len(parts)))


def is_newer(candidate: str, current: str = __version__) -> bool:
    return _parse(candidate) > _parse(current)


def check(timeout: float = 10) -> Release | None:
    """The latest release if it is newer than this copy, else None. Raises on network errors."""
    req = urllib.request.Request(API, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.load(resp)
    tag = data.get("tag_name", "")
    if data.get("draft") or data.get("prerelease") or not is_newer(tag):
        return None
    for a in data.get("assets", []):
        if a.get("name") == ASSET:
            digest = a.get("digest") or ""
            return Release(tag.lstrip("vV"), a["browser_download_url"], a.get("size", 0),
                           digest.split(":", 1)[1] if digest.startswith("sha256:") else "", data.get("body") or "")
    return None


def install_dir() -> Path | None:
    """The folder Remind.exe runs from, or None when running from source."""
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else None


def download(rel: Release, progress=None) -> Path:
    """Download and unpack the release; returns the folder that holds the new Remind.exe."""
    work = Path(tempfile.mkdtemp(prefix=f"Remind-{rel.version}-"))
    zpath = work / ASSET
    h = hashlib.sha256()
    req = urllib.request.Request(rel.url, headers={"User-Agent": UA["User-Agent"]})
    with urllib.request.urlopen(req, timeout=60) as resp, open(zpath, "wb") as out:
        total = int(resp.headers.get("Content-Length") or rel.size or 0)
        done = 0
        while chunk := resp.read(1 << 16):
            out.write(chunk)
            h.update(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)
    if rel.sha256 and h.hexdigest().lower() != rel.sha256.lower():
        raise RuntimeError("The downloaded update is damaged (checksum mismatch). Try again later.")
    dest = (work / "unpacked").resolve()
    with zipfile.ZipFile(zpath) as z:
        for name in z.namelist():                      # refuse entries that would land outside the folder
            if not (dest / name).resolve().is_relative_to(dest):
                raise RuntimeError("The update archive is not valid.")
        z.extractall(dest)
    exe = next(dest.rglob("Remind.exe"), None)
    if not exe:
        raise RuntimeError("The update archive does not contain Remind.exe.")
    return exe.parent


def _ps_quote(s: str) -> str:
    return "'" + s.replace("'", "''") + "'"


def apply(new_app: Path) -> None:
    """Start the hidden step that replaces the install folder once Remind exits and relaunches it.

    The caller must quit the app right after this returns.
    """
    target = install_dir()
    if target is None:
        raise RuntimeError("Updating only works in the installed app (this is running from source).")
    exe = target / "Remind.exe"
    script = f"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
try {{
    Wait-Process -Id {os.getpid()} -Timeout 60 -ErrorAction SilentlyContinue
    robocopy {_ps_quote(str(new_app))} {_ps_quote(str(target))} /MIR /R:5 /W:2 /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) {{ throw "Copying the new version failed (robocopy code $LASTEXITCODE)." }}
    Get-ChildItem -LiteralPath {_ps_quote(str(target))} -Recurse -File | Unblock-File
    Start-Process -FilePath {_ps_quote(str(exe))}
}} catch {{
    [void][System.Windows.MessageBox]::Show("Remind could not be updated.`n`n$($_.Exception.Message)", 'Remind', 'OK', 'Error')
}}
"""
    ps1 = new_app.parent / "apply-update.ps1"
    ps1.write_text(script, encoding="utf-8-sig")
    # Run as text, not as a script file, so an execution policy set by Group Policy can't block it.
    cmd = f"& ([scriptblock]::Create((Get-Content -Raw -LiteralPath {_ps_quote(str(ps1))})))"
    # CREATE_NO_WINDOW alone: combined with DETACHED_PROCESS the helper silently never ran (no copy, no restart).
    # CREATE_NEW_PROCESS_GROUP keeps it alive after Remind exits.
    subprocess.Popen(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", cmd],
                     creationflags=subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP, close_fds=True)

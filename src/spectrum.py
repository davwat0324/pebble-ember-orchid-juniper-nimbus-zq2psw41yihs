"""Employee names, status and email from Spectrum, read with the user's own Spectrum SQL login.

Reads the custom email view set in the sign-in box (one row per employee). If no view is set, or
the login can't see it yet (not granted), it falls back to the same query on the employee master
views.

The server, database and view are deliberately NOT in the source or in the released .exe (the
GitHub repo and its releases are public). Each user types them once in the Spectrum sign-in box
and they are saved in %APPDATA%\\Remind\\settings.json. REMIND_SPECTRUM_SERVER / _DATABASE /
_EMAIL_VIEW environment variables override them (handy for development).
"""
from __future__ import annotations

import os
from dataclasses import dataclass

import loader
import store

DEFAULT_DATABASE = "Spectrum"

FALLBACK_SQL = """
SELECT LTRIM(RTRIM(m.Employee_Code)) AS Employee,
       RTRIM(m.First_Name) AS First_Name,
       RTRIM(m.Last_Name) AS Last_Name,
       m.Employment_Status AS Status,
       COALESCE(NULLIF(RTRIM(e.Employee_Email), ''), NULLIF(RTRIM(e.Employee_Personal_Email), '')) AS Email
FROM PR_EMPLOYEE_MASTER_1 m
LEFT JOIN PR_EMPLOYEE_MASTER_3 e ON e.Company_Code = m.Company_Code AND e.Employee_Code = m.Employee_Code
"""

STATUS_NAMES = {"A": "Active", "T": "Terminated", "I": "Inactive", "D": "Deceased"}


def site() -> dict:
    """{'server', 'database', 'view'} from the environment, else the user's saved settings."""
    s = store.load_settings()
    return {
        "server": os.environ.get("REMIND_SPECTRUM_SERVER") or s.get("spectrum_server", ""),
        "database": os.environ.get("REMIND_SPECTRUM_DATABASE") or s.get("spectrum_database") or DEFAULT_DATABASE,
        "view": os.environ.get("REMIND_SPECTRUM_EMAIL_VIEW") or s.get("spectrum_view", ""),
    }


def save_site(server: str, database: str, view: str) -> None:
    s = store.load_settings()
    s.update(spectrum_server=server.strip(), spectrum_database=database.strip() or DEFAULT_DATABASE,
             spectrum_view=view.strip())
    store.save_settings(s)


class LoginRequired(Exception):
    pass


@dataclass
class Employee:
    number: str
    first: str
    last: str
    status: str
    email: str

    @property
    def status_name(self) -> str:
        return STATUS_NAMES.get(self.status, f"Status {self.status}" if self.status else "")


def _driver() -> str:
    import pyodbc
    have = pyodbc.drivers()
    for d in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
        if d in have:
            return d
    raise RuntimeError("Microsoft ODBC Driver 18 for SQL Server is not installed on this PC "
                       "(free from Microsoft). Install it, then try again.")


def connect(cred: dict | None = None):
    import pyodbc
    cfg = site()
    cred = cred or store.load_credentials()
    if not cfg["server"]:
        raise LoginRequired("Enter the Spectrum server (ask whoever set up Remind) and your login.")
    if not cred:
        raise LoginRequired("Sign in to Spectrum to look up emails.")
    pw = cred["Password"].replace("}", "}}")  # braces must be doubled inside {...} in an ODBC string
    cs = (f"DRIVER={{{_driver()}}};SERVER={cfg['server']};DATABASE={cfg['database']};UID="
          f"{cred['User']};PWD={{{pw}}};Encrypt=yes;TrustServerCertificate=yes;Connection Timeout=20")
    try:
        cn = pyodbc.connect(cs, timeout=20)
    except pyodbc.Error as e:
        if "Login failed" in str(e):
            raise LoginRequired("Spectrum didn't accept that user name and password.") from e
        raise
    cn.timeout = 120
    return cn


def _view_missing(e: Exception, view: str) -> bool:
    msg = str(e)
    return bool(view) and view in msg or "Invalid object name" in msg or "SELECT permission" in msg


def fetch_all(cn=None) -> dict[str, Employee]:
    """Every employee, keyed by employee number without padding ('1234')."""
    import pyodbc
    view = site()["view"]
    own = cn is None
    cn = cn or connect()
    try:
        cur = cn.cursor()
        try:
            if not view:
                raise LookupError("no email view set")
            cur.execute(f"SELECT Employee, First_Name, Last_Name, Status, Email FROM {view}")
        except (pyodbc.Error, LookupError) as e:
            if isinstance(e, pyodbc.Error) and not _view_missing(e, view):
                raise
            cur = cn.cursor()
            cur.execute(FALLBACK_SQL)
        out = {}
        for r in cur.fetchall():
            num = loader.normalize(r.Employee or "")
            out[num] = Employee(num, (r.First_Name or "").strip(), (r.Last_Name or "").strip(),
                                (r.Status or "").strip(), (r.Email or "").strip())
        return out
    finally:
        if own:
            cn.close()

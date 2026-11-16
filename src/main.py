"""Remind - email a reminder to every employee on a list.

Load the employee list (CSV or Excel) -> duplicates are removed and each number's email is looked
up in Spectrum -> adjust who gets it -> write the message (starts from the saved default) -> send
through Outlook, one email per person.
"""
from __future__ import annotations

import dataclasses
import re
import sys
from pathlib import Path

from PySide6.QtCore import (QAbstractTableModel, QEvent, QModelIndex, QSortFilterProxyModel, Qt, QThread, QTimer,
                            Signal)
from PySide6.QtGui import QAction, QColor, QFont, QIcon, QKeySequence, QPalette, QTextCharFormat, QTextListFormat
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QFrame, QHBoxLayout, QHeaderView, QInputDialog, QLabel, QLineEdit, QMainWindow, QMenu, QMessageBox,
    QProgressDialog, QPushButton, QStackedWidget, QTableView, QTableWidget, QTableWidgetItem, QTextBrowser,
    QTextEdit, QToolButton, QVBoxLayout, QWidget,
)

import loader
import mailer
import spectrum
import store
import updater
from mailer import Message, Recipient
from version import __version__ as VERSION

APP_NAME = "Remind"
ACCENT = "#005DB4"
EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")


def resource(name: str) -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))  # PyInstaller unpacks data files to _MEIPASS
    return str(base / name)


STYLE = f"""
* {{ font-family: 'Segoe UI'; font-size: 10pt; color: #1F2328; }}
QMainWindow, QDialog, #page {{ background: #FFFFFF; }}
#header {{ background: #FFFFFF; border-bottom: 1px solid #ECEEF1; }}
#brand {{ font-size: 13pt; font-weight: 600; color: {ACCENT}; }}
#step {{ color: #9AA1AA; }}
#stepOn {{ color: #1F2328; font-weight: 600; }}
#title {{ font-size: 16pt; font-weight: 600; }}
#muted {{ color: #6B7280; }}
#warn {{ color: #B54708; }}
#drop {{ border: 2px dashed #CBD2DA; border-radius: 14px; background: #F8FAFC; }}
#drop[hover="true"] {{ border-color: {ACCENT}; background: #EEF5FC; }}
#dropTitle {{ font-size: 14pt; font-weight: 600; }}
QPushButton {{ background: #FFFFFF; border: 1px solid #D5DAE0; border-radius: 8px; padding: 7px 16px; }}
QPushButton:hover {{ background: #F3F5F7; }}
QPushButton:disabled {{ color: #A0A7B0; }}
QPushButton#primary {{ background: {ACCENT}; border-color: {ACCENT}; color: white; font-weight: 600; }}
QPushButton#primary:hover {{ background: #004E98; }}
QPushButton#primary:disabled {{ background: #9DBEE0; border-color: #9DBEE0; }}
QPushButton#link {{ border: none; background: transparent; color: {ACCENT}; padding: 4px 6px; }}
QPushButton#link:hover {{ text-decoration: underline; }}
QToolButton {{ border: none; border-radius: 6px; padding: 4px 8px; background: transparent; }}
QToolButton:hover {{ background: #F0F2F5; }}
QToolButton:checked {{ background: #E3EDF8; }}
QToolButton::menu-indicator {{ image: none; width: 0; }}
QLineEdit, QComboBox, QTextEdit {{ border: 1px solid #D5DAE0; border-radius: 8px; padding: 6px 8px;
    background: white; selection-background-color: #CFE2F6; selection-color: #1F2328; }}
QLineEdit:focus, QTextEdit:focus, QComboBox:focus {{ border-color: {ACCENT}; }}
QComboBox::drop-down {{ border: none; width: 24px; }}
QTableView {{ border: 1px solid #ECEEF1; border-radius: 10px; gridline-color: transparent;
    selection-background-color: #EEF5FC; selection-color: #1F2328; alternate-background-color: #FAFBFC; }}
QTableView::item {{ padding: 4px 6px; border: none; }}
QHeaderView::section {{ background: white; border: none; border-bottom: 1px solid #ECEEF1;
    padding: 8px 6px; color: #6B7280; font-weight: 600; }}
QToolButton#chip {{ background: #F0F2F5; border-radius: 10px; padding: 3px 10px; }}
#toolbar {{ border: 1px solid #D5DAE0; border-bottom: none; border-top-left-radius: 8px;
    border-top-right-radius: 8px; background: #FAFBFC; }}
QTextEdit#body {{ border-top-left-radius: 0; border-top-right-radius: 0; padding: 10px 12px; }}
QStatusBar {{ background: white; color: #6B7280; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: #D5DAE0; border-radius: 4px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: #B8C0CA; }}
QScrollBar::add-line, QScrollBar::sub-line, QScrollBar::add-page, QScrollBar::sub-page {{ background: none; height: 0; }}
"""


def light_palette(app: QApplication) -> None:
    """Always light, even when Windows is in dark mode (Fusion would follow it)."""
    try:
        app.styleHints().setColorScheme(Qt.ColorScheme.Light)
    except AttributeError:
        pass  # Qt < 6.8 has no setColorScheme
    pal = app.style().standardPalette()
    for role, color in ((QPalette.Window, "#FFFFFF"), (QPalette.Base, "#FFFFFF"),
                        (QPalette.AlternateBase, "#FAFBFC"), (QPalette.Text, "#1F2328"),
                        (QPalette.WindowText, "#1F2328"), (QPalette.Button, "#FFFFFF"),
                        (QPalette.ButtonText, "#1F2328"), (QPalette.Highlight, "#CFE2F6"),
                        (QPalette.HighlightedText, "#1F2328"), (QPalette.ToolTipBase, "#FFFFFF"),
                        (QPalette.ToolTipText, "#1F2328"), (QPalette.Link, ACCENT)):
        pal.setColor(role, QColor(color))
    app.setPalette(pal)


def label(text="", name=None, wrap=False) -> QLabel:
    lab = QLabel(text)
    if name:
        lab.setObjectName(name)
    lab.setWordWrap(wrap)
    return lab


def button(text, name=None, slot=None) -> QPushButton:
    b = QPushButton(text)
    if name:
        b.setObjectName(name)
    if slot:
        b.clicked.connect(slot)
    b.setCursor(Qt.PointingHandCursor)
    return b


class Worker(QThread):
    done = Signal(object)
    failed = Signal(object)

    def __init__(self, fn, parent=None):
        super().__init__(parent)
        self.fn = fn

    def run(self):
        try:
            self.done.emit(self.fn())
        except Exception as e:  # handed to the GUI thread through the signal
            self.failed.emit(e)


# --- dialogs -------------------------------------------------------------------------------------------------------
class LoginDialog(QDialog):
    def __init__(self, parent, note=""):
        super().__init__(parent)
        self.setWindowTitle("Spectrum login")
        self.setMinimumWidth(400)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 22, 24, 18)
        lay.setSpacing(12)
        lay.addWidget(label("Sign in to Spectrum", "title"))
        lay.addWidget(label((note + "\n\n" if note else "") + "Your Spectrum SQL login is used only to look up emails. "
                            "It is saved, encrypted, for your Windows user on this PC.", "muted", True))
        site = spectrum.site()
        form = QFormLayout()
        form.setSpacing(10)
        self.server = QLineEdit(site["server"])
        self.server.setPlaceholderText("SQL Server host name")
        self.database = QLineEdit(site["database"])
        self.view = QLineEdit(site["view"])
        self.view.setPlaceholderText("optional - custom view with employee emails")
        self.user, self.pw = QLineEdit(store.saved_user()), QLineEdit()
        self.pw.setEchoMode(QLineEdit.Password)
        form.addRow("Server", self.server)
        form.addRow("Database", self.database)
        form.addRow("Email view", self.view)
        form.addRow("User name", self.user)
        form.addRow("Password", self.pw)
        lay.addLayout(form)
        self.error = label("", "warn", True)
        lay.addWidget(self.error)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Save).setObjectName("primary")
        bb.accepted.connect(self._save)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        first_empty = next((e for e in (self.server, self.user, self.pw) if not e.text()), self.pw)
        first_empty.setFocus()

    def _save(self):
        if not (self.server.text().strip() and self.user.text().strip() and self.pw.text()):
            self.error.setText("Server, user name and password are needed.")
            return
        if not re.fullmatch(r"[\w.\[\] ]*", self.view.text().strip()):
            self.error.setText("The email view name can only contain letters, digits, _ . and [ ].")
            return
        spectrum.save_site(self.server.text(), self.database.text(), self.view.text())
        for w in QApplication.topLevelWidgets():          # keep the main window's in-memory copy in step
            if isinstance(w, MainWindow):
                w.settings.update({k: v for k, v in store.load_settings().items() if k.startswith("spectrum_")})
        store.save_credentials(self.user.text().strip(), self.pw.text())
        self.accept()


class AddDialog(QDialog):
    """Add someone by employee number (looked up) or by name + email."""

    def __init__(self, parent, employees: dict):
        super().__init__(parent)
        self.employees = employees
        self.filled = False
        self.setWindowTitle("Add person")
        self.setMinimumWidth(380)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(22, 20, 22, 16)
        lay.setSpacing(10)
        lay.addWidget(label("Add person", "title"))
        form = QFormLayout()
        form.setSpacing(10)
        self.num, self.first, self.last, self.email = QLineEdit(), QLineEdit(), QLineEdit(), QLineEdit()
        self.num.setPlaceholderText("optional - fills in the rest")
        self.num.textChanged.connect(self._lookup)
        form.addRow("Employee #", self.num)
        form.addRow("First name", self.first)
        form.addRow("Last name", self.last)
        form.addRow("Email", self.email)
        lay.addLayout(form)
        self.msg = label("", "warn")
        lay.addWidget(self.msg)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText("Add")
        bb.button(QDialogButtonBox.Ok).setObjectName("primary")
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _lookup(self, text):
        e = self.employees.get(loader.normalize(text))
        if e or self.filled:
            self.first.setText(e.first if e else "")
            self.last.setText(e.last if e else "")
            self.email.setText(e.email if e else "")
        self.filled = bool(e)
        self.msg.setText("No email in Spectrum - type one in." if e and not e.email else "")

    def _ok(self):
        if not EMAIL_RE.match(self.email.text().strip()):
            self.msg.setText("Enter a valid email address.")
            return
        self.accept()

    def recipient(self) -> tuple[Recipient, str]:
        num = loader.normalize(self.num.text())
        e = self.employees.get(num)
        r = Recipient(num if e else "", self.first.text().strip(), self.last.text().strip(),
                      self.email.text().strip(), "Added")
        return r, (e.status if e else "")


class PreviewDialog(QDialog):
    def __init__(self, parent, msg: Message, people: list[Recipient]):
        super().__init__(parent)
        self.msg, self.people = msg, people
        self.setWindowTitle("Preview")
        self.resize(720, 620)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(22, 18, 22, 16)
        lay.setSpacing(10)
        top = QHBoxLayout()
        top.addWidget(label("Preview as", "muted"))
        self.who = QComboBox()
        for r in people:
            self.who.addItem(f"{r.full or r.email}  <{r.email}>")
        self.who.currentIndexChanged.connect(self._show)
        top.addWidget(self.who, 1)
        lay.addLayout(top)
        self.view = QTextBrowser()
        self.view.setOpenExternalLinks(True)
        lay.addWidget(self.view, 1)
        note = "Your Outlook signature is added below the message." if msg.signature else ""
        if msg.attachments:
            note += ("  " if note else "") + "Attached: " + ", ".join(Path(a).name for a in msg.attachments)
        lay.addWidget(label(note, "muted", True))
        row = QHBoxLayout()
        row.addWidget(button("Send a test to myself", slot=self._test))
        row.addStretch()
        row.addWidget(button("Close", "primary", self.accept))
        lay.addLayout(row)
        self._show()

    def _current(self) -> Recipient:
        return self.people[max(self.who.currentIndex(), 0)]

    def _show(self):
        r = self._current()
        body = mailer.merge(mailer.body_inner(self.msg.body_html), r, is_html=True)
        subject = mailer.merge(self.msg.subject, r, is_html=False)
        cc = f"<br><b>CC:</b> {self.msg.cc}" if self.msg.cc.strip() else ""
        self.view.setHtml(f"<div style='color:#6B7280'><b>To:</b> {r.email}{cc}<br><b>Subject:</b> "
                          f"{subject}</div><hr>{body}")

    def _test(self):
        try:
            app = mailer.outlook()
            me = mailer.my_address(app)
            if not me:
                me, ok = QInputDialog.getText(self, "Send a test", "Send the test to:")
                if not ok or not me.strip():
                    return
            test = dataclasses.replace(self.msg, subject="[TEST] " + self.msg.subject)
            note = mailer.deliver(app, test, self._current(), to=me)
            QMessageBox.information(self, "Test sent",
                                    f"Test email sent to {me}." + (f"\n\nIt was {note}." if note else ""))
        except Exception as e:
            QMessageBox.warning(self, "Could not send the test", mailer.describe(e))


class HistoryDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("History")
        self.resize(860, 480)
        self.items = store.load_history()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(22, 18, 22, 16)
        lay.setSpacing(10)
        lay.addWidget(label("History", "title"))
        self.table = QTableWidget(len(self.items), 5)
        self.table.setHorizontalHeaderLabels(["When", "Subject", "", "People", "File"])
        self.table.verticalHeader().hide()
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setShowGrid(False)
        for i, h in enumerate(self.items):
            ok = sum(1 for r in h.get("recipients", []) if r.get("ok"))
            fail = len(h.get("recipients", [])) - ok
            vals = [h.get("when", "").replace("T", "  "), h.get("subject", ""),
                    "Drafts" if h.get("mode") == "drafts" else "Sent",
                    f"{ok}" + (f"  ({fail} failed)" if fail else ""), h.get("file", "")]
            for c, v in enumerate(vals):
                self.table.setItem(i, c, QTableWidgetItem(v))
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.cellDoubleClicked.connect(self._detail)
        lay.addWidget(self.table, 1)
        row = QHBoxLayout()
        row.addWidget(label("Double-click a row to see who it went to." if self.items else "Nothing sent yet.",
                            "muted"))
        row.addStretch()
        exp = button("Export to Excel", slot=self._export)
        exp.setEnabled(bool(self.items))
        row.addWidget(exp)
        row.addWidget(button("Close", "primary", self.accept))
        lay.addLayout(row)

    def _detail(self, row, _col):
        h = self.items[row]
        lines = [f"{r.get('name') or ''}  <{r.get('email')}>" + ("" if r.get("ok") else f"  - FAILED: {r.get('error')}")
                 for r in h.get("recipients", [])]
        box = QMessageBox(self)
        box.setWindowTitle(APP_NAME)
        box.setText(f"<b>{h.get('subject', '')}</b><br>{len(lines)} people, {h.get('when', '').replace('T', ' ')}")
        box.setDetailedText("\n".join(lines))
        box.exec()

    def _export(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export history", str(Path.home() / "Remind history.xlsx"),
                                              "Excel (*.xlsx)")
        if not path:
            return
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "History"
        ws.append(["When", "Subject", "Mode", "File", "Employee #", "Name", "Email", "Result"])
        for h in self.items:
            for r in h.get("recipients", []):
                ws.append([h.get("when", "").replace("T", " "), h.get("subject"), h.get("mode"), h.get("file"),
                           r.get("number"), r.get("name"), r.get("email"),
                           "OK" if r.get("ok") else f"Failed: {r.get('error')}"])
        for col, w in zip("ABCDEFGH", (20, 40, 8, 30, 11, 26, 34, 30)):
            ws.column_dimensions[col].width = w
        try:
            wb.save(path)
        except OSError as e:
            QMessageBox.warning(self, "Export", f"Could not save the file:\n{e}")


class SettingsDialog(QDialog):
    def __init__(self, parent, settings: dict):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setMinimumWidth(440)
        self.reset_template = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 16)
        lay.setSpacing(12)
        lay.addWidget(label("Settings", "title"))
        row = QHBoxLayout()
        self.login = label()
        self._show_login()
        row.addWidget(self.login)
        row.addStretch()
        row.addWidget(button("Change", "link", self._change_login))
        lay.addLayout(row)
        self.sig = QCheckBox("Add my Outlook signature to each email")
        self.sig.setChecked(settings.get("signature", True))
        lay.addWidget(self.sig)
        self.reset = QCheckBox("Reset the default template to the original placeholder message")
        lay.addWidget(self.reset)
        ver = QHBoxLayout()
        ver.addWidget(label(f"Remind {VERSION}  ·  settings in {store.DATA_DIR}", "muted", True), 1)
        ver.addWidget(button("Check for updates", "link", lambda: parent.check_updates(manual=True)))
        lay.addLayout(ver)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setObjectName("primary")
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _show_login(self):
        user = store.saved_user()
        self.login.setText(f"Spectrum login: <b>{user}</b>" if user else "Spectrum login: not set")

    def _change_login(self):
        if LoginDialog(self).exec():
            self._show_login()


# --- the recipient table ------------------------------------------------------------------------------------------
class Row:
    __slots__ = ("r", "checked", "code")

    def __init__(self, r: Recipient, checked: bool, code: str = ""):
        self.r, self.checked, self.code = r, checked, code  # code: Spectrum status code (T and D are never ticked)


def make_row(number: str, employees: dict) -> Row:
    """A table row for one employee #: ticked when there is an email and they are not terminated."""
    e = employees.get(number)
    if not e:
        return Row(Recipient(number, "", "", "", "Not in Spectrum"), False)
    ok = e.status not in ("T", "D") and bool(EMAIL_RE.match(e.email))
    return Row(Recipient(number, e.first, e.last, e.email, e.status_name if e.email else "No email"), ok, e.status)


class RecipientModel(QAbstractTableModel):
    HEAD = ["", "First name", "Last name", "Employee #", "Email", "Status"]
    EDITABLE = {1: "first", 2: "last", 3: "number", 4: "email"}  # column -> Recipient attribute
    changed = Signal()

    def __init__(self):
        super().__init__()
        self.rows = []
        self.employees = {}

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()
        self.changed.emit()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(self.HEAD)

    def headerData(self, s, o, role=Qt.DisplayRole):
        if o == Qt.Horizontal and role == Qt.DisplayRole:
            return self.HEAD[s]
        return None

    def flags(self, ix):
        f = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if ix.column() == 0:
            f |= Qt.ItemIsUserCheckable
        elif ix.column() in self.EDITABLE:
            f |= Qt.ItemIsEditable
        return f

    def data(self, ix, role=Qt.DisplayRole):
        row = self.rows[ix.row()]
        r, c = row.r, ix.column()
        if role == Qt.CheckStateRole and c == 0:
            return Qt.Checked if row.checked else Qt.Unchecked
        if role == Qt.EditRole and c in self.EDITABLE:
            return getattr(r, self.EDITABLE[c])
        if role == Qt.DisplayRole:
            return ["", r.first, r.last, r.number, r.email or "No email - double-click to add", r.status][c]
        if role == Qt.ForegroundRole:
            if (c == 4 and not r.email) or (c == 5 and r.status in ("No email", "Not in Spectrum")):
                return QColor("#B54708")
            if c == 5:
                return QColor("#9AA1AA" if row.code in ("T", "D") else "#6B7280")
        if role == Qt.ToolTipRole and c in self.EDITABLE:
            return "Double-click to change it for this send (Spectrum is not changed)"
        if role == Qt.UserRole:  # sort key
            if c == 0:
                return int(row.checked)
            if c == 3:
                return int(r.number) if r.number.isdigit() else 10**9
            return self.data(ix, Qt.DisplayRole).lower()
        return None

    def setData(self, ix, value, role=Qt.EditRole):
        row, c = self.rows[ix.row()], ix.column()
        if role == Qt.CheckStateRole and c == 0:
            row.checked = Qt.CheckState(value) == Qt.Checked and bool(EMAIL_RE.match(row.r.email))
        elif role == Qt.EditRole and c in self.EDITABLE:
            text = str(value).strip()
            old = getattr(row.r, self.EDITABLE[c])
            if text == old:
                return False
            if c == 3:  # a different employee #: look the person up again
                self.rows[ix.row()] = make_row(loader.normalize(text), self.employees)
            else:
                setattr(row.r, self.EDITABLE[c], text)
            if c == 4:
                valid = bool(EMAIL_RE.match(text))
                row.checked = valid and (row.checked or not EMAIL_RE.match(old))
                if text and row.r.status in ("No email", "Not in Spectrum"):
                    row.r.status = "Email added"
        else:
            return False
        self.dataChanged.emit(self.index(ix.row(), 0), self.index(ix.row(), self.columnCount() - 1))
        self.changed.emit()
        return True

    def set_all(self, checked: bool, rows: list[int] | None = None):
        for i in (rows if rows is not None else range(len(self.rows))):
            self.rows[i].checked = checked and bool(EMAIL_RE.match(self.rows[i].r.email))
        if self.rows:
            self.dataChanged.emit(self.index(0, 0), self.index(len(self.rows) - 1, 0))
        self.changed.emit()

    def add(self, row: Row):
        self.beginInsertRows(QModelIndex(), len(self.rows), len(self.rows))
        self.rows.append(row)
        self.endInsertRows()
        self.changed.emit()

    def remove(self, idx: list[int]):
        drop = set(idx)
        self.set_rows([x for i, x in enumerate(self.rows) if i not in drop])

    def selected(self) -> list[Recipient]:
        return [x.r for x in self.rows if x.checked and EMAIL_RE.match(x.r.email)]


class DropZone(QFrame):
    dropped = Signal(str)

    def __init__(self, browse):
        super().__init__()
        self.setObjectName("drop")
        self.setAcceptDrops(True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(40, 56, 40, 56)
        lay.setSpacing(10)
        lay.setAlignment(Qt.AlignCenter)
        icon = QLabel()
        icon.setPixmap(QIcon(resource("remind.ico")).pixmap(56, 56))
        icon.setAlignment(Qt.AlignCenter)
        lay.addWidget(icon)
        for text, name in (("Drop the employee list here", "dropTitle"),
                           ("CSV or Excel. Duplicates are removed and emails come from Spectrum.", "muted")):
            lab = label(text, name)
            lab.setAlignment(Qt.AlignCenter)
            lay.addWidget(lab)
        b = button("Choose file…", "primary", browse)
        b.setFixedWidth(160)
        lay.addSpacing(6)
        lay.addWidget(b, 0, Qt.AlignCenter)

    def _hover(self, on):
        self.setProperty("hover", "true" if on else "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self._hover(True)

    def dragLeaveEvent(self, e):
        self._hover(False)

    def dropEvent(self, e):
        self._hover(False)
        urls = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        if urls:
            self.dropped.emit(urls[0])


def page() -> tuple[QWidget, QVBoxLayout]:
    w = QWidget()
    w.setObjectName("page")
    lay = QVBoxLayout(w)
    lay.setContentsMargins(36, 26, 36, 22)
    lay.setSpacing(12)
    return w, lay


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(QIcon(resource("remind.ico")))
        self.settings = store.load_settings()
        self.employees = {}
        self.sheet = None
        self.attachments = []
        self.worker = None

        root = QWidget()
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.addWidget(self._header())
        self.stack = QStackedWidget()
        self.stack.addWidget(self._file_page())
        self.stack.addWidget(self._people_page())
        self.stack.addWidget(self._message_page())
        outer.addWidget(self.stack, 1)
        self.setCentralWidget(root)
        self.statusBar().setSizeGripEnabled(False)
        self.go(0)
        self.resize(*(self.settings.get("size") or (1000, 720)))
        self.update_worker = None
        if updater.install_dir():                  # installed copy: look for a newer release shortly after start
            QTimer.singleShot(2500, self.check_updates)

    # ---- updates ---------------------------------------------------------------------------
    def check_updates(self, manual=False):
        if self.update_worker and self.update_worker.isRunning():
            return
        self.update_worker = Worker(updater.check, self)
        self.update_worker.done.connect(lambda rel: self._update_checked(rel, manual))
        if manual:
            self.update_worker.failed.connect(lambda e: QMessageBox.warning(
                self, APP_NAME, f"Couldn't check for updates.\n\n{e}"))
        self.update_worker.start()

    def _update_checked(self, rel, manual):
        if rel is None:
            if manual:
                QMessageBox.information(self, APP_NAME, f"You have the latest version ({VERSION}).")
            return
        notes = rel.notes.strip()
        notes = (notes[:700] + "…") if len(notes) > 700 else notes
        text = f"Remind {rel.version} is available. You have {VERSION}."
        box = QMessageBox(QMessageBox.Information, "Update available", text, parent=self)
        box.setInformativeText("Update now? Remind will close and reopen in a few seconds.")
        if notes:
            box.setDetailedText(notes)
        now = box.addButton("Update now", QMessageBox.AcceptRole)
        box.addButton("Later", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() is not now:
            return
        if not updater.install_dir():
            QMessageBox.information(self, APP_NAME, "This copy runs from source; pull the new code instead.")
            return
        prog = QProgressDialog("Downloading the update…", "", 0, 0, self)
        prog.setCancelButton(None)
        prog.setWindowTitle(APP_NAME)
        prog.setWindowModality(Qt.WindowModal)
        prog.setMinimumDuration(0)
        prog.show()
        self.update_worker = Worker(lambda: updater.download(rel), self)

        def ready(path):
            prog.close()
            try:
                updater.apply(path)
            except Exception as e:
                QMessageBox.warning(self, APP_NAME, f"Couldn't install the update.\n\n{e}")
                return
            self.close()                           # saves settings; the update step restarts Remind

        def failed(e):
            prog.close()
            QMessageBox.warning(self, APP_NAME, f"Couldn't download the update.\n\n{e}")

        self.update_worker.done.connect(ready)
        self.update_worker.failed.connect(failed)
        self.update_worker.start()

    # ---- header with the three steps -------------------------------------------------------
    def _header(self):
        h = QFrame()
        h.setObjectName("header")
        lay = QHBoxLayout(h)
        lay.setContentsMargins(24, 12, 16, 12)
        lay.addWidget(label("Remind", "brand"))
        lay.addSpacing(28)
        self.steps = []
        for i, t in enumerate(["File", "Recipients", "Message"]):
            if i:
                lay.addWidget(label("›", "step"))
            s = label(f"{i + 1}  {t}", "step")
            self.steps.append(s)
            lay.addWidget(s)
        lay.addStretch()
        for text, slot in (("History", self._history), ("Settings", self._settings)):
            b = QToolButton()
            b.setText(text)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(slot)
            lay.addWidget(b)
        return h

    def go(self, i: int):
        self.stack.setCurrentIndex(i)
        for n, s in enumerate(self.steps):
            s.setObjectName("stepOn" if n == i else "step")
            s.style().unpolish(s)
            s.style().polish(s)
        if i == 2:
            self._update_to()

    # ---- step 1: file ----------------------------------------------------------------------
    def _file_page(self):
        w, lay = page()
        lay.addStretch(1)
        self.drop = DropZone(self._browse)
        self.drop.dropped.connect(self.load_file)
        self.drop.setFixedWidth(620)
        lay.addWidget(self.drop, 0, Qt.AlignHCenter)
        self.file_status = label("", "muted")
        self.file_status.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.file_status)
        lay.addStretch(2)
        return w

    def _browse(self):
        start = self.settings.get("last_dir") or str(Path.home())
        path, _ = QFileDialog.getOpenFileName(self, "Employee list", start,
                                              "Employee list (*.csv *.xlsx *.xlsm *.txt);;All files (*)")
        if path:
            self.load_file(path)

    def load_file(self, path: str):
        try:
            sheet = loader.read(path)
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, f"Could not read that file:\n{e}")
            return
        if not sheet.rows:
            QMessageBox.warning(self, APP_NAME, "That file has no rows.")
            return
        self.sheet = sheet
        self.attachments.clear()
        self._show_attachments()
        self.settings["last_dir"] = str(Path(path).parent)
        many = len(sheet.columns) > 1
        self.col.blockSignals(True)
        self.col.clear()
        self.col.addItems(sheet.columns)
        self.col.setCurrentIndex(loader.guess_column(sheet))
        self.col.blockSignals(False)
        self.col.setVisible(many)
        self.col_label.setVisible(many)
        if self.employees:
            self._build_rows()
            self.go(1)
        else:
            self._lookup()

    def _lookup(self, note=""):
        signed_in = store.load_credentials() and spectrum.site()["server"]
        if not signed_in and not LoginDialog(self, note).exec():
            return
        self.file_status.setText("Looking up emails in Spectrum…")
        self.drop.setEnabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.worker = Worker(spectrum.fetch_all, self)
        self.worker.done.connect(self._looked_up)
        self.worker.failed.connect(self._lookup_failed)
        self.worker.start()

    def _lookup_finished(self):
        QApplication.restoreOverrideCursor()
        self.drop.setEnabled(True)
        self.file_status.setText("")

    def _looked_up(self, employees):
        self._lookup_finished()
        self.employees = employees
        self._build_rows()
        self.go(1)

    def _lookup_failed(self, e):
        self._lookup_finished()
        if isinstance(e, spectrum.LoginRequired):
            if LoginDialog(self, str(e)).exec():
                self._lookup()
            return
        msg = str(e)
        if "permission" in msg.lower():
            where = spectrum.site()["view"] or "the employee master views"
            msg = ("Your Spectrum login does not have access to the employee email list ("
                   f"{where}). Ask the Spectrum administrator to grant SELECT on it.\n\n{msg}")
        QMessageBox.warning(self, APP_NAME, f"Could not read Spectrum.\n\n{msg}")

    # ---- step 2: recipients ----------------------------------------------------------------
    def _people_page(self):
        w, lay = page()
        top = QHBoxLayout()
        box = QVBoxLayout()
        box.setSpacing(2)
        box.addWidget(label("Recipients", "title"))
        self.source = label("", "muted")
        box.addWidget(self.source)
        top.addLayout(box)
        top.addStretch()
        self.col_label = label("Employee # is in", "muted")
        self.col = QComboBox()
        self.col.currentIndexChanged.connect(lambda _: self.sheet and self._build_rows())
        top.addWidget(self.col_label)
        top.addWidget(self.col)
        lay.addLayout(top)

        bar = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search name, number or email")
        self.search.setClearButtonEnabled(True)
        self.search.setFixedWidth(300)
        bar.addWidget(self.search)
        bar.addSpacing(10)
        self.counts = label("", "muted")
        bar.addWidget(self.counts, 1)
        bar.addWidget(button("Select all", "link", lambda: self._select(True)))
        bar.addWidget(button("Select none", "link", lambda: self._select(False)))
        lay.addLayout(bar)

        self.model = RecipientModel()
        self.model.changed.connect(self._update_counts)
        self.proxy = QSortFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterKeyColumn(-1)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy.setSortRole(Qt.UserRole)
        self.search.textChanged.connect(self.proxy.setFilterFixedString)
        t = self.table = QTableView()
        t.setModel(self.proxy)
        t.setSortingEnabled(True)
        t.sortByColumn(2, Qt.AscendingOrder)
        t.setAlternatingRowColors(True)
        t.setShowGrid(False)
        t.setWordWrap(False)
        t.setSelectionBehavior(QAbstractItemView.SelectRows)
        t.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        t.verticalHeader().hide()
        t.verticalHeader().setDefaultSectionSize(34)
        hh = t.horizontalHeader()
        hh.setHighlightSections(False)
        for c, mode in enumerate((QHeaderView.Fixed, QHeaderView.Stretch, QHeaderView.Stretch,
                                  QHeaderView.ResizeToContents, QHeaderView.Stretch, QHeaderView.ResizeToContents)):
            hh.setSectionResizeMode(c, mode)
        hh.resizeSection(0, 40)
        delete = QAction(t)
        delete.setShortcut(QKeySequence.Delete)
        delete.triggered.connect(self._remove)
        t.addAction(delete)
        lay.addWidget(t, 1)

        foot = QHBoxLayout()
        foot.addWidget(button("Back", slot=lambda: self.go(0)))
        foot.addSpacing(12)
        foot.addWidget(button("+ Add person", "link", self._add))
        foot.addWidget(button("Remove selected rows", "link", self._remove))
        foot.addStretch()
        self.next1 = button("Next: Message", "primary", lambda: self.go(2))
        foot.addWidget(self.next1)
        lay.addLayout(foot)
        return w

    def _build_rows(self):
        numbers, total = loader.unique_numbers(self.sheet, self.col.currentIndex())
        self.search.clear()
        self.model.employees = self.employees
        self.model.set_rows([make_row(n, self.employees) for n in numbers])
        self.source.setText(f"{self.sheet.path.name}  ·  {total:,} rows  →  {len(numbers):,} unique employees")

    def _update_counts(self):
        rows = self.model.rows
        sel = len(self.model.selected())
        missing = sum(1 for x in rows if x.r.status == "Not in Spectrum")
        no_email = sum(1 for x in rows if x.r.status == "No email")
        term = sum(1 for x in rows if x.code in ("T", "D") and x.r.email and not x.checked)
        parts = [f"<b style='color:#1F2328'>{sel}</b> selected"]
        if no_email:
            parts.append(f"<span style='color:#B54708'>{no_email} no email</span>")
        if missing:
            parts.append(f"<span style='color:#B54708'>{missing} not in Spectrum</span>")
        if term:
            parts.append(f"{term} terminated")
        self.counts.setText("  ·  ".join(parts))
        self.next1.setEnabled(sel > 0)

    def _visible_rows(self) -> list[int]:
        return [self.proxy.mapToSource(self.proxy.index(i, 0)).row() for i in range(self.proxy.rowCount())]

    def _select(self, on: bool):
        self.model.set_all(on, self._visible_rows() if self.search.text() else None)

    def _add(self):
        d = AddDialog(self, self.employees)
        if d.exec():
            r, code = d.recipient()
            self.model.add(Row(r, True, code))

    def _remove(self):
        idx = [self.proxy.mapToSource(i).row() for i in self.table.selectionModel().selectedRows()]
        if idx:
            self.model.remove(idx)

    # ---- step 3: message ------------------------------------------------------------------
    def _message_page(self):
        w, lay = page()
        head = QHBoxLayout()
        head.addWidget(label("Message", "title"))
        head.addStretch()
        head.addWidget(label("Template", "muted"))
        self.tpl = QComboBox()
        self.tpl.setMinimumWidth(220)
        self.tpl.setToolTip("Start from a saved message. The default one is loaded when Remind opens.")
        self.tpl.activated.connect(lambda _i: self._load_template(self.tpl.currentData(), quiet=False))
        head.addWidget(self.tpl)
        opts = QToolButton()
        opts.setText("Template  ▾")
        opts.setPopupMode(QToolButton.InstantPopup)
        opts.setCursor(Qt.PointingHandCursor)
        tm = QMenu(opts)
        tm.addAction("Save changes to this template", self._save_template)
        tm.addAction("Save as new template…", self._save_template_as)
        tm.addSeparator()
        tm.addAction("Set as default", self._set_default_template)
        tm.addAction("Rename…", self._rename_template)
        tm.addAction("Delete", self._delete_template)
        opts.setMenu(tm)
        head.addWidget(opts)
        lay.addLayout(head)
        self.to_label = label("", "muted")
        lay.addWidget(self.to_label)
        form = QFormLayout()
        form.setSpacing(8)
        form.setHorizontalSpacing(14)
        self.cc, self.bcc, self.subject = QLineEdit(), QLineEdit(), QLineEdit()
        self.subject.installEventFilter(self)
        self.cc.setPlaceholderText("optional - copied on every email")
        self.bcc.setPlaceholderText("optional")
        form.addRow("Subject", self.subject)
        form.addRow("CC", self.cc)
        form.addRow("BCC", self.bcc)
        lay.addLayout(form)

        bar = QFrame()
        bar.setObjectName("toolbar")
        tl = QHBoxLayout(bar)
        tl.setContentsMargins(6, 4, 6, 4)
        tl.setSpacing(2)
        self.body = QTextEdit()
        self.body.setObjectName("body")
        self.body.document().setDefaultFont(QFont("Calibri", 11))
        self.body.setTabChangesFocus(True)
        self.body.installEventFilter(self)

        def tool(text, tip, slot, checkable=False, **font):
            b = QToolButton()
            b.setText(text)
            b.setToolTip(tip)
            b.setCheckable(checkable)
            b.setCursor(Qt.PointingHandCursor)
            f = b.font()
            f.setBold(font.get("bold", False))
            f.setItalic(font.get("italic", False))
            f.setUnderline(font.get("underline", False))
            b.setFont(f)
            b.clicked.connect(slot)
            tl.addWidget(b)
            return b

        self.b_bold = tool("B", "Bold (Ctrl+B)", lambda on: self._fmt(bold=on), True, bold=True)
        self.b_ital = tool("I", "Italic (Ctrl+I)", lambda on: self._fmt(italic=on), True, italic=True)
        self.b_und = tool("U", "Underline (Ctrl+U)", lambda on: self._fmt(underline=on), True, underline=True)
        tool("•  List", "Bulleted list", self._bullets)
        tool("Link", "Make the selected text a link", self._link)
        for key, b in ((QKeySequence.Bold, self.b_bold), (QKeySequence.Italic, self.b_ital),
                       (QKeySequence.Underline, self.b_und)):
            a = QAction(self.body)
            a.setShortcut(key)
            a.setShortcutContext(Qt.WidgetShortcut)
            a.triggered.connect(b.click)
            self.body.addAction(a)
        self.body.currentCharFormatChanged.connect(self._sync_fmt)
        tl.addSpacing(10)
        fields = QToolButton()
        fields.setText("Insert field  ▾")
        fields.setToolTip("Filled in for each person. Or type { in the subject or message.")
        fields.setPopupMode(QToolButton.InstantPopup)
        fields.setCursor(Qt.PointingHandCursor)
        menu = QMenu(fields)
        for token, (_attr, text) in mailer.FIELDS.items():
            menu.addAction(text, lambda t=token: (self.body.insertPlainText(t), self.body.setFocus()))
        fields.setMenu(menu)
        tl.addWidget(fields)
        tl.addStretch()
        tool("Attach file", "Attach a file to every email", self._attach)
        box = QVBoxLayout()
        box.setSpacing(0)
        box.addWidget(bar)
        box.addWidget(self.body, 1)
        lay.addLayout(box, 1)

        self.att_row = QHBoxLayout()
        self.att_row.setSpacing(6)
        self.att_row.addStretch()
        lay.addLayout(self.att_row)

        foot = QHBoxLayout()
        foot.addWidget(button("Back", slot=lambda: self.go(1)))
        foot.addStretch()
        foot.addWidget(button("Preview", slot=self._preview))
        foot.addWidget(button("Save as drafts", slot=lambda: self._send(drafts=True)))
        self.send_btn = button("Send", "primary", lambda: self._send(drafts=False))
        self.send_btn.setMinimumWidth(110)
        foot.addWidget(self.send_btn)
        lay.addLayout(foot)
        self._load_default(True)
        return w

    def _fmt(self, bold=None, italic=None, underline=None):
        f = QTextCharFormat()
        if bold is not None:
            f.setFontWeight(QFont.Bold if bold else QFont.Normal)
        if italic is not None:
            f.setFontItalic(italic)
        if underline is not None:
            f.setFontUnderline(underline)
        self.body.mergeCurrentCharFormat(f)
        self.body.setFocus()

    def _sync_fmt(self, f: QTextCharFormat):
        self.b_bold.setChecked(f.fontWeight() >= QFont.Bold)
        self.b_ital.setChecked(f.fontItalic())
        self.b_und.setChecked(f.fontUnderline())

    def _bullets(self):
        cur = self.body.textCursor()
        lst = cur.currentList()
        if lst:
            lst.remove(cur.block())
            bf = cur.blockFormat()
            bf.setIndent(0)
            cur.setBlockFormat(bf)
        else:
            cur.createList(QTextListFormat.ListDisc)
        self.body.setFocus()

    def _link(self):
        cur = self.body.textCursor()
        if not cur.hasSelection():
            QMessageBox.information(self, APP_NAME, "Select the text to turn into a link first.")
            return
        url, ok = QInputDialog.getText(self, "Link", "Web address:", text="https://")
        if ok and url.strip():
            f = QTextCharFormat()
            f.setAnchor(True)
            f.setAnchorHref(url.strip())
            f.setForeground(QColor("#0563C1"))
            f.setFontUnderline(True)
            cur.mergeCharFormat(f)

    def eventFilter(self, obj, e):
        """Typing { in the subject or message pops up the fields to pick from, like a mail merge."""
        if e.type() == QEvent.KeyPress and e.text() == "{" and obj in (self.subject, self.body):
            QTimer.singleShot(0, lambda: self._suggest(obj))
        return super().eventFilter(obj, e)

    def _suggest(self, w):
        menu = QMenu(self)
        for token, (_attr, text) in mailer.FIELDS.items():
            menu.addAction(text).setData(token[1:])
        menu.setActiveAction(menu.actions()[0])
        if w is self.body:
            pos = w.viewport().mapToGlobal(w.cursorRect().bottomLeft())
        else:
            pos = w.mapToGlobal(w.cursorRect().bottomLeft())
        act = menu.exec(pos)
        if act:
            w.insertPlainText(act.data()) if w is self.body else w.insert(act.data())
        w.setFocus()

    def _attach(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Attach files", self.settings.get("last_dir", ""))
        for p in paths:
            if p not in self.attachments:
                self.attachments.append(p)
        self._show_attachments()

    def _show_attachments(self):
        while self.att_row.count() > 1:
            it = self.att_row.takeAt(0)
            if it.widget():
                it.widget().deleteLater()
        for i, p in enumerate(self.attachments):
            chip = QToolButton()
            chip.setObjectName("chip")
            chip.setText(f"{Path(p).name}   ✕")
            chip.setToolTip(f"{p}\nClick to remove")
            chip.setCursor(Qt.PointingHandCursor)
            chip.clicked.connect(lambda _=False, p=p: (self.attachments.remove(p), self._show_attachments()))
            self.att_row.insertWidget(i, chip)

    # ---- templates -------------------------------------------------------------------------
    def _templates(self) -> dict:
        return self.settings["templates"]

    def _refresh_templates(self, select: str):
        default = self.settings["default_template"]
        self.tpl.blockSignals(True)
        self.tpl.clear()
        for name in sorted(self._templates(), key=str.lower):
            self.tpl.addItem(f"{name}  (default)" if name == default else name, name)
        self.tpl.setCurrentIndex(max(self.tpl.findData(select), 0))
        self.tpl.blockSignals(False)
        self.tpl_name = self.tpl.currentData()

    def _load_default(self, quiet=True):
        self._load_template(self.settings["default_template"], quiet)

    def _load_template(self, name: str, quiet=True):
        t = self._templates().get(name) or store.blank_template()
        self.subject.setText(t["subject"])
        self.cc.setText(t["cc"])
        self.bcc.setText(t["bcc"])
        self.body.setHtml(t["body"])
        self._refresh_templates(name)
        if not quiet:
            self.statusBar().showMessage(f"Template “{name}” loaded.", 3000)

    def _current_template(self) -> dict:
        return {"subject": self.subject.text(), "body": self.body.toHtml(), "cc": self.cc.text(),
                "bcc": self.bcc.text()}

    def _save_templates(self, select: str, note: str):
        store.save_settings(self.settings)
        self._refresh_templates(select)
        self.statusBar().showMessage(note, 4000)

    def _ask_name(self, title: str, text: str = "") -> str | None:
        name, ok = QInputDialog.getText(self, title, "Template name:", text=text)
        name = name.strip()
        if not ok or not name:
            return None
        if name in self._templates() and name != text:
            if QMessageBox.question(self, APP_NAME, f"Replace the template “{name}”?") != QMessageBox.Yes:
                return None
        return name

    def _save_template(self):
        self._templates()[self.tpl_name] = self._current_template()
        self._save_templates(self.tpl_name, f"Template “{self.tpl_name}” saved.")

    def _save_template_as(self):
        name = self._ask_name("Save as new template")
        if name:
            self._templates()[name] = self._current_template()
            self._save_templates(name, f"Template “{name}” saved.")

    def _set_default_template(self):
        self.settings["default_template"] = self.tpl_name
        self._save_templates(self.tpl_name, f"“{self.tpl_name}” is now the default. Remind opens with it.")

    def _rename_template(self):
        old = self.tpl_name
        name = self._ask_name("Rename template", old)
        if not name or name == old:
            return
        self._templates()[name] = self._templates().pop(old)
        if self.settings["default_template"] == old:
            self.settings["default_template"] = name
        self._save_templates(name, f"Renamed to “{name}”.")

    def _delete_template(self):
        name = self.tpl_name
        if len(self._templates()) == 1:
            QMessageBox.information(self, APP_NAME, "This is the only template, so it can't be deleted.")
            return
        if QMessageBox.question(self, APP_NAME, f"Delete the template “{name}”?") != QMessageBox.Yes:
            return
        del self._templates()[name]
        if self.settings["default_template"] == name:
            self.settings["default_template"] = sorted(self._templates(), key=str.lower)[0]
        self._save_templates(self.settings["default_template"], f"Template “{name}” deleted.")
        self._load_default(True)

    def _update_to(self):
        n = len(self.model.selected())
        self.to_label.setText(f"To: <b style='color:#1F2328'>{n} {'person' if n == 1 else 'people'}</b>, "
                              "each gets their own email.")
        self.send_btn.setText(f"Send {n}")

    def _message(self) -> Message:
        return Message(self.subject.text(), self.body.toHtml(), self.cc.text(), self.bcc.text(),
                       list(self.attachments), self.settings.get("signature", True))

    def _preview(self):
        people = self.model.selected()
        if people:
            PreviewDialog(self, self._message(), people).exec()

    def _checks(self, msg: Message) -> bool:
        problems = []
        if not msg.subject.strip():
            problems.append("The subject is empty.")
        text = self.body.toPlainText() + " " + msg.subject
        tokens = set(re.findall(r"\{[A-Za-z]+\}", text))
        broken = {t for t in tokens & set(mailer.FIELDS) if t not in msg.body_html + msg.subject}
        unknown = tokens - set(mailer.FIELDS) | broken
        if unknown:
            problems.append("These will not be filled in: " + ", ".join(sorted(unknown))
                            + (" (remove the formatting inside the braces)" if broken else ""))
        if re.search(r"\[[^\]]+\]", text):
            problems.append("The message still has placeholder text in [brackets].")
        for name, val in (("CC", msg.cc), ("BCC", msg.bcc)):
            bad = [a.strip() for a in re.split("[;,]", val) if a.strip() and not EMAIL_RE.match(a.strip())]
            if bad:
                problems.append(f"{name} has an address that does not look right: {bad[0]}")
        if not problems:
            return True
        return QMessageBox.warning(self, APP_NAME, "\n\n".join(problems) + "\n\nContinue anyway?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes

    def _send(self, drafts: bool):
        people = self.model.selected()
        msg = self._message()
        if not people or not self._checks(msg):
            return
        n = len(people)
        where = "saved to your Outlook Drafts folder" if drafts else "sent from your Outlook"
        if QMessageBox.question(self, APP_NAME, f"{n} email{'s' if n != 1 else ''} will be {where}.\n"
                                "Each person gets their own email.\n\nContinue?",
                                QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel) != QMessageBox.Yes:
            return
        try:
            app = mailer.outlook()
        except Exception as e:
            QMessageBox.warning(self, APP_NAME, str(e))
            return
        verb = "Saving" if drafts else "Sending"
        prog = QProgressDialog(f"{verb}…", "Stop", 0, n, self)
        prog.setWindowTitle(APP_NAME)
        prog.setWindowModality(Qt.WindowModal)
        prog.setMinimumDuration(0)
        results = []
        for i, r in enumerate(people):
            if prog.wasCanceled():
                break
            prog.setLabelText(f"{verb} {i + 1} of {n}: {r.full or r.email}")
            entry = {"number": r.number, "name": r.full, "email": r.email, "ok": True}
            try:
                note = mailer.deliver(app, msg, r, drafts=drafts)
                if note:
                    entry["note"] = note
            except Exception as e:
                entry.update(ok=False, error=mailer.describe(e))
            results.append(entry)
            prog.setValue(i + 1)
            QApplication.processEvents()
        prog.close()
        store.add_history({"mode": "drafts" if drafts else "sent", "subject": msg.subject,
                           "file": self.sheet.path.name if self.sheet else "", "recipients": results})
        ok = sum(1 for x in results if x["ok"])
        failed = [x for x in results if not x["ok"]]
        text = f"{ok} email{'s' if ok != 1 else ''} {'saved to Drafts' if drafts else 'sent'}."
        if len(results) < n:
            text += f"\nStopped: {n - len(results)} not {'saved' if drafts else 'sent'}."
        noted = [x for x in results if x.get("note")]
        if noted:
            text += f"\n{len(noted)} {noted[0]['note']}."
        box = QMessageBox(QMessageBox.Warning if failed else QMessageBox.Information, APP_NAME, text,
                          parent=self)
        if failed:
            box.setInformativeText(f"{len(failed)} failed. See details.")
            box.setDetailedText("\n".join(f"{x['name']} <{x['email']}>: {x.get('error')}" for x in failed))
        box.exec()

    # ---- header buttons ------------------------------------------------------------------
    def _history(self):
        HistoryDialog(self).exec()

    def _settings(self):
        d = SettingsDialog(self, self.settings)
        if not d.exec():
            return
        self.settings["signature"] = d.sig.isChecked()
        if d.reset.isChecked():
            self._templates()[self.settings["default_template"]] = store.blank_template()
            self._load_default(False)
        store.save_settings(self.settings)

    def closeEvent(self, e):
        self.settings["size"] = [self.width(), self.height()]
        store.save_settings(self.settings)
        super().closeEvent(e)


def main():
    if sys.platform == "win32":
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Remind.EmailReminders")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    light_palette(app)
    app.setStyleSheet(STYLE)
    w = MainWindow()
    w.show()
    if len(sys.argv) > 1 and Path(sys.argv[1]).is_file():
        w.load_file(sys.argv[1])
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

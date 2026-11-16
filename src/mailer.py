"""Builds and sends the emails through the user's own classic Outlook (COM), one email per person.

Nothing about the mail server lives in the app: Outlook sends from its default account and the
message lands in the user's Sent Items. Outlook's default signature is added the same way Outlook
does it for a new message (opening the item's inspector fills the body with the signature).
"""
from __future__ import annotations

import html
import os
import re
from dataclasses import dataclass, field

FIELDS = {
    "{FirstName}": ("first", "First name"),
    "{LastName}": ("last", "Last name"),
    "{EmployeeNumber}": ("number", "Employee #"),
}


@dataclass
class Recipient:
    number: str
    first: str
    last: str
    email: str
    status: str = ""

    @property
    def full(self) -> str:
        return f"{self.first} {self.last}".strip()


@dataclass
class Message:
    subject: str
    body_html: str
    cc: str = ""
    bcc: str = ""
    attachments: list[str] = field(default_factory=list)
    signature: bool = True


def merge(text: str, r: Recipient, is_html: bool) -> str:
    for token, (attr, _label) in FIELDS.items():
        val = getattr(r, attr) or ""
        if not val and attr == "first":
            val = "there"
        text = text.replace(token, html.escape(val) if is_html else val)
    return text


_BODY_TAG = re.compile(r"<body[^>]*>", re.I)


def body_inner(doc: str) -> str:
    """The part of an HTML document between <body> and </body> (Qt writes full documents)."""
    m = _BODY_TAG.search(doc)
    if not m:
        return doc
    end = doc.lower().rfind("</body>")
    return doc[m.end():end if end > m.end() else len(doc)]


def _style(doc: str) -> str:
    """Qt's own CSS block, kept so lists/bold/links look the same in Outlook."""
    m = re.search(r"<style[^>]*>.*?</style>", doc, re.I | re.S)
    return m.group(0) if m else ""


def outlook():
    import win32com.client
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise RuntimeError("Couldn't start Outlook. Remind needs classic Outlook installed and signed in.") from e


def my_address(app=None) -> str:
    """The SMTP address Outlook sends from (its first account), or '' if Outlook won't say."""
    app = app or outlook()
    try:
        if app.Session.Accounts.Count:
            addr = app.Session.Accounts.Item(1).SmtpAddress
            if addr:
                return addr
    except Exception:
        pass
    try:
        ex = app.Session.CurrentUser.AddressEntry.GetExchangeUser()
        return ex.PrimarySmtpAddress if ex else ""
    except Exception:
        return ""


class SendError(Exception):
    """One email could not be built or sent. The message says which step failed and why."""


E_INVALIDARG = -2147024809          # 0x80070057 "The parameter is incorrect."
OL_SAVE = 0                         # OlInspectorClose.olSave


def _hresult(e: Exception) -> int | None:
    """The HRESULT of a pywin32 COM error (the inner one when Outlook wraps it), else None."""
    args = getattr(e, "args", ())
    info = getattr(e, "excepinfo", None)
    if info and len(info) > 5 and info[5]:
        return info[5]
    return args[0] if args and isinstance(args[0], int) else None


def describe(e: BaseException) -> str:
    """Readable text for an exception, preferring Outlook's own message for COM errors."""
    if isinstance(e, SendError):
        return str(e)
    info = getattr(e, "excepinfo", None)
    text = info[2] if info and len(info) > 2 and info[2] else ""
    args = getattr(e, "args", ())
    if not text and len(args) > 1 and isinstance(args[1], str):
        text = args[1]
    code = _hresult(e)
    text = text or str(e) or type(e).__name__
    return f"{text.strip()} (0x{code & 0xFFFFFFFF:08X})" if isinstance(code, int) else text.strip()


def _step(what: str, fn, *args):
    """Run one Outlook call; on failure raise SendError naming the step."""
    try:
        return fn(*args)
    except SendError:
        raise
    except Exception as e:
        raise SendError(f"{what}: {describe(e)}") from e


def _with_signature(item, ours: str) -> bool:
    """Insert our text above Outlook's default signature. False if Outlook wouldn't give the signature."""
    try:
        item.GetInspector                    # opening the (hidden) inspector makes Outlook add the signature
        current = item.HTMLBody or ""
        m = _BODY_TAG.search(current)
        item.HTMLBody = current[:m.end()] + ours + current[m.end():] if m else ours + current
        return True
    except Exception:
        return False


def _plain_html(ours: str) -> str:
    return f"<html><head><meta charset='utf-8'></head><body>{ours}</body></html>"


def build(app, msg: Message, r: Recipient, to: str | None = None, signature: bool | None = None):
    """A ready-to-send Outlook MailItem for one recipient (not saved or sent).

    Each Outlook call is wrapped so a failure says which step broke. If the signature can't be
    added, the email is built without it rather than failing.
    """
    signature = msg.signature if signature is None else signature
    item = _step("Creating the email in Outlook", app.CreateItem, 0)
    _step("Setting the To address", setattr, item, "To", to or r.email)
    if msg.cc.strip() and not to:
        _step("Setting CC", setattr, item, "CC", msg.cc)
    if msg.bcc.strip() and not to:
        _step("Setting BCC", setattr, item, "BCC", msg.bcc)
    _step("Setting the subject", setattr, item, "Subject", merge(msg.subject, r, is_html=False))
    ours = (_style(msg.body_html) + '<div style="font-family:Calibri,sans-serif;font-size:11pt">'
            + merge(body_inner(msg.body_html), r, is_html=True) + "</div>")
    if not (signature and _with_signature(item, ours)):
        _step("Setting the message text", setattr, item, "HTMLBody", _plain_html(ours))
    for path in msg.attachments:
        path = os.path.normpath(os.path.abspath(path))   # Qt gives forward slashes; Outlook wants a full Windows path
        _step(f"Attaching {os.path.basename(path)}", lambda p: item.Attachments.Add(p), path)
    return item


def _close_inspector(app, item):
    """Save the item, close the hidden inspector opened for the signature, and return the item reloaded."""
    item.Save()
    try:
        item.GetInspector.Close(OL_SAVE)
    except Exception:
        pass
    return app.Session.GetItemFromID(item.EntryID)


def deliver(app, msg: Message, r: Recipient, *, drafts: bool = False, to: str | None = None) -> str:
    """Build one email and send it (or save it to Drafts). Returns a note ('' normally); raises SendError.

    Some Outlook builds refuse Send() with "The parameter is incorrect" (0x80070057) for an email whose
    hidden inspector is still open (it is opened to add the signature). When that happens the email is
    saved, the inspector closed and Send() retried; if Outlook still refuses, the email is rebuilt without
    the signature (no inspector at all) and sent that way.
    """
    item = build(app, msg, r, to)
    if drafts:
        try:
            _close_inspector(app, item)      # saves; closing the inspector keeps hundreds of them from piling up
        except Exception as e:
            raise SendError(f"Saving to Drafts: {describe(e)}") from e
        return ""
    try:
        item.Send()
        return ""
    except Exception as e:
        if _hresult(e) != E_INVALIDARG:
            raise SendError(f"Sending: {describe(e)}") from e
        first = e
    try:
        _close_inspector(app, item).Send()
        return ""
    except Exception:
        pass
    try:
        item.Delete()                         # the saved draft from the retry above
    except Exception:
        pass
    if msg.signature:
        try:
            build(app, msg, r, to, signature=False).Send()
            return "sent without your Outlook signature (Outlook refused it with the signature)"
        except Exception:
            pass
    raise SendError(f"Sending: {describe(first)}") from first

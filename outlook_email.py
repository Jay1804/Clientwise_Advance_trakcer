"""Outlook (desktop) COM automation for emailing split report files.

Drives a locally installed Outlook client via pywin32 - no SMTP credentials
involved. This is deliberately swappable: send_via_outlook is the only
function that talks to Outlook, so switching to SMTP later just means
swapping this module's implementation, not the callers in app.py.
"""

import os
import tempfile


def outlook_available() -> bool:
    """True if pywin32 looks importable. Checked lazily (not at module import
    time) so this module still imports fine in environments without pywin32
    or Outlook - the UI just disables the email section instead of crashing."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def send_via_outlook(
    to_address: str,
    cc_address: str,
    subject: str,
    body: str,
    attachment_bytes: bytes,
    attachment_filename: str,
) -> None:
    """Compose and send one email through the local Outlook desktop app.

    Outlook's COM API needs a real file path for attachments, not in-memory
    bytes, so this writes the attachment to a private temp directory first
    and removes it again once the send attempt finishes (success or not).

    COM must be initialized on whichever thread calls this - fine on a plain
    script's main thread, but Streamlit runs the script in a worker
    ("ScriptRunner") thread that never gets an implicit CoInitialize, which
    surfaces as "CoInitialize has not been called." pythoncom.CoInitialize()/
    CoUninitialize() bracket the call so it works from either caller.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        tmp_dir = tempfile.mkdtemp(prefix="advancetracker_email_")
        tmp_path = os.path.join(tmp_dir, attachment_filename)
        try:
            with open(tmp_path, "wb") as f:
                f.write(attachment_bytes)

            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # olMailItem
            mail.To = to_address
            if cc_address:
                mail.CC = cc_address
            mail.Subject = subject
            mail.Body = body
            mail.Attachments.Add(tmp_path)
            mail.Send()
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            try:
                os.rmdir(tmp_dir)
            except OSError:
                pass
    finally:
        pythoncom.CoUninitialize()

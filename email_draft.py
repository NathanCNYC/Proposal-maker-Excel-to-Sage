import os, time, psutil, win32com.client as win32, pythoncom
import pythoncom
from pathlib import Path


def _get_outlook(max_wait: int = 90):
    """
    Return a live Outlook.Application COM object.
    • Starts Outlook.exe if it's not running.
    • First tries GetActiveObject().
    • If that fails, tries Dispatch() until it succeeds (up to max_wait s).
    """
    # 1) start Outlook if needed
    if not any(p.name().lower() == "outlook.exe" for p in psutil.process_iter()):
        os.startfile("outlook")          # shell resolves full path automatically

    # 2) poll for COM object
    for _ in range(max_wait):
        # a) grab an already-registered instance
        try:
            return win32.GetActiveObject("Outlook.Application")
        except pythoncom.com_error:
            pass

        # b) attempt to create a new COM server (works once Outlook has finished loading)
        try:
            return win32.Dispatch("Outlook.Application")
        except pythoncom.com_error:
            time.sleep(1)

    raise RuntimeError(
        f"Outlook COM server not available after {max_wait}s.\n"
        "Make sure Outlook opens correctly; then run the script again."
    )


def create_outlook_draft(data: dict, pdf_path: Path, quote_id: str) -> None:
    outlook = _get_outlook()          # ← COM object is guaranteed now
    mail    = outlook.CreateItem(0)   # olMailItem

    mail.To      = data["EMAIL"]
    mail.CC      = "" # Enter any emails you want CCd on every email
    mail.Subject = f"Quote #{quote_id} {data['JOB']}"

    # --- plain-text body -------------------------------------------------
    body = """Dear Customer,

Our proposal is attached as a PDF.

........................................

whatever message you want in here
"""

    mail.Body = body
    mail.Attachments.Add(str(pdf_path))
    mail.Display(True)

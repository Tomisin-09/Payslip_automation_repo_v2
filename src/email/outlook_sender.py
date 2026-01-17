import os
from typing import Optional

import win32com.client as win32
import pythoncom


class OutlookEmailSender:
    """
    Windows-only Outlook email sender using COM automation.
    """

    def __init__(self, send_mode: str = "send"):
        if send_mode not in ("send", "display"):
            raise ValueError("send_mode must be 'send' or 'display'")
        self.send_mode = send_mode
        self.outlook = None

    def __enter__(self):
        # Ensure COM is initialized for this thread/process
        pythoncom.CoInitialize()
        try:
            # Dispatch attaches to the running Outlook session if present
            self.outlook = win32.Dispatch("Outlook.Application")
        except Exception as e:
            raise RuntimeError(
                "Failed to start Outlook COM automation. "
                "Ensure Outlook desktop is installed, opened at least once, "
                "and you are running under the same Windows user profile."
            ) from e
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.outlook = None
        # Clean up COM for this thread/process
        pythoncom.CoUninitialize()

    def send_email(
        self,
        to_address: str,
        subject: str,
        body: str,
        attachment_path: Optional[str] = None,
        sender_mailbox: Optional[str] = None,
    ):
        if not to_address or not isinstance(to_address, str):
            raise ValueError("Recipient email address is missing or invalid")
        if not subject:
            raise ValueError("Email subject cannot be empty")
        if not body:
            raise ValueError("Email body cannot be empty")

        mail = self.outlook.CreateItem(0)  # 0 = MailItem
        mail.To = to_address
        mail.Subject = subject
        mail.Body = body

        if sender_mailbox:
            mail.SentOnBehalfOfName = sender_mailbox

        if attachment_path:
            attachment_path = os.path.abspath(attachment_path)
            if not os.path.exists(attachment_path):
                raise FileNotFoundError(f"Attachment not found: {attachment_path}")
            mail.Attachments.Add(attachment_path)

        if self.send_mode == "display":
            mail.Display()
        else:
            mail.Send()

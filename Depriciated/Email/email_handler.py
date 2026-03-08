print(">>> Loaded email_handler from:", __file__)

import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
from typing import Iterable, Optional
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

SENDER_EMAIL = os.getenv("PROVA_EMAIL")
APP_PASSWORD = os.getenv("PROVA_APP_PASSWORD")
SENDER_NAME = os.getenv("PROVA_SENDER_NAME", "ProVA Assistant")


# -----------------------------
# Helper: ensure input is a list
# -----------------------------
def _ensure_list(x):
    """
    Converts input to a list.
    - None -> []
    - string -> [string]
    - iterable -> list(iterable)
    """
    if x is None:
        return []
    if isinstance(x, (str, bytes)):
        return [x]
    try:
        return list(x)
    except Exception:
        return [x]


# -----------------------------
# Helper: add file attachment
# -----------------------------
def add_attachment(msg: MIMEMultipart, filepath: str):
    filename = os.path.basename(filepath)

    with open(filepath, "rb") as f:
        attachment = MIMEBase("application", "octet-stream")
        attachment.set_payload(f.read())

    encoders.encode_base64(attachment)
    attachment.add_header(
        "Content-Disposition",
        f'attachment; filename="{filename}"'
    )

    msg.attach(attachment)


# -----------------------------
# Main function: send email
# -----------------------------
def send_email(
    to_email: Iterable[str],
    subject: str,
    message_text: str,
    attachments: Optional[Iterable[str]] = None,
    is_html: bool = False,
    cc: Optional[Iterable[str]] = None,
    bcc: Optional[Iterable[str]] = None,
):
    """
    Send an email using Gmail SMTP.

    to_email : str or list[str]
    subject : str
    message_text : str (plain text or HTML)
    attachments : list[str] (file paths)
    is_html : bool
    cc, bcc : optional recipients
    """

    to_list = _ensure_list(to_email)
    cc_list = _ensure_list(cc)
    bcc_list = _ensure_list(bcc)

    all_recipients = to_list + cc_list + bcc_list

    if not all_recipients:
        raise ValueError("No recipients provided")

    msg = MIMEMultipart()
    msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    # Body
    body_type = "html" if is_html else "plain"
    msg.attach(MIMEText(message_text, body_type))

    # Attachments
    for file in _ensure_list(attachments):
        add_attachment(msg, file)

    # Send email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())

    return "Email sent successfully!"

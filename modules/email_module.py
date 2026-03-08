"""
ProVA — modules/email_module.py  (v2.0 — gated flow)
======================================================
Voice-driven email composition and sending via Gmail SMTP.

Voice flow:
  1. "Send email to John"
  2. ProVA asks subject   → voice confirm → retry → typed fallback
  3. ProVA asks message   → multi-sentence "Add more?" loop → confirm
  4. ProVA reads summary  → "Shall I send it?"
     • No → "Change subject / message / cancel?" → loops back
  5. Send

Key architectural change vs v1
─────────────────────────────────
handle() now accepts gate_event (the _module_active Event from voice_module).
It holds the gate for the ENTIRE email session so the main voice loop cannot
race in and interpret subject/body dictation as a new command.

Router must pass gate_event:
  _exec(_get_em_handle(), cmd, speak_fn, listen_fn, gate_event, on_done=on_done)

Contacts:
  - contacts.json maps spoken names → email addresses
  - Unknown name → ask → normalize → confirm → offer to save (3 attempts)

Dependencies:
  pip install python-dotenv
"""

from __future__ import annotations

import json
import logging
import os
import re
import smtplib
import threading
import time
from dataclasses import dataclass
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path
from typing import Callable, Iterable, Optional

log = logging.getLogger("ProVA.Email")

# ─────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────
_MODULE_DIR    = Path(__file__).parent
_ENV_PATH      = _MODULE_DIR.parent / ".env"
_CONTACTS_FILE = _MODULE_DIR.parent / "data" / "contacts.json"
_MAX_ATTACH_MB = 20
_SMTP_TIMEOUT  = 15
_SMALL_PAUSE   = 0.25   # seconds between speak() and listen() to prevent echo


# ─────────────────────────────────────────────────────────────────
# RESULT
# ─────────────────────────────────────────────────────────────────
@dataclass
class Result:
    success: bool
    message: str


# ─────────────────────────────────────────────────────────────────
# CREDENTIALS
# ─────────────────────────────────────────────────────────────────
def _load_credentials() -> tuple[Optional[str], Optional[str], str]:
    try:
        from dotenv import load_dotenv
        load_dotenv(dotenv_path=_ENV_PATH, override=False)
    except ImportError:
        log.warning("python-dotenv not installed — reading env vars directly.")

    email       = os.getenv("PROVA_EMAIL")
    password    = os.getenv("PROVA_APP_PASSWORD")
    sender_name = os.getenv("PROVA_SENDER_NAME", "ProVA Assistant")

    if not email or not password:
        log.warning(
            "Email credentials not set. Add PROVA_EMAIL and "
            "PROVA_APP_PASSWORD to %s", _ENV_PATH
        )
    return email, password, sender_name


SENDER_EMAIL, APP_PASSWORD, SENDER_NAME = _load_credentials()


# ─────────────────────────────────────────────────────────────────
# CONTACTS BOOK
# ─────────────────────────────────────────────────────────────────
def _load_contacts() -> dict[str, str]:
    if _CONTACTS_FILE.exists():
        try:
            return json.loads(_CONTACTS_FILE.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning("Could not load contacts.json: %s", e)
    return {}


def _save_contact(name: str, email: str) -> None:
    contacts = _load_contacts()
    contacts[name.lower().strip()] = email.strip()
    try:
        _CONTACTS_FILE.write_text(
            json.dumps(contacts, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        log.info("Saved contact: %s → %s", name, email)
    except Exception as e:
        log.warning("Could not save contact: %s", e)


def _resolve_recipient(name: str) -> Optional[str]:
    """
    Resolve spoken name → email. Returns address or None.

    If the input looks like an email (contains @), normalize it first and
    reject if the local part still has spaces — that means STT garbled it
    (e.g. "nida qureshi n 04@gmail.com") and the user should type instead.
    """
    name = name.strip()
    # Looks like an email address — normalize first, then validate cleanly
    if "@" in name or re.search(r"\bat\b", name, re.IGNORECASE):
        normalized = _normalize_spoken_email(name)
        # If local part has spaces it's a voice-garbled address — don't accept it
        local = normalized.split("@")[0] if "@" in normalized else normalized
        if " " in local:
            log.info("Rejecting garbled voice email (spaces in local): %r", normalized)
            return None
        if _is_valid_email(normalized):
            return normalized
        return None
    contacts   = _load_contacts()
    name_lower = name.lower()
    if name_lower in contacts:
        return contacts[name_lower]
    for key, addr in contacts.items():
        if key in name_lower or name_lower in key:
            return addr
    return None


def _is_valid_email(addr: str) -> bool:
    return bool(re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+", addr.strip()))


# ─────────────────────────────────────────────────────────────────
# ATTACHMENT HELPERS
# ─────────────────────────────────────────────────────────────────
def _validate_attachment(filepath: str) -> Optional[str]:
    p = Path(filepath)
    if not p.exists():  return f"Attachment not found: {p.name}"
    if not p.is_file(): return f"'{p.name}' is not a file."
    size_mb = p.stat().st_size / (1024 * 1024)
    if size_mb > _MAX_ATTACH_MB:
        return f"'{p.name}' is {size_mb:.1f} MB — max is {_MAX_ATTACH_MB} MB."
    return None


def _add_attachment(msg: MIMEMultipart, filepath: str) -> None:
    p    = Path(filepath)
    part = MIMEBase("application", "octet-stream")
    part.set_payload(p.read_bytes())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{p.name}"')
    msg.attach(part)


def _ensure_list(x) -> list:
    if x is None:                   return []
    if isinstance(x, (str, bytes)): return [x]
    try:                            return list(x)
    except Exception:               return [x]


# ─────────────────────────────────────────────────────────────────
# CORE SEND — pure function, no voice logic
# ─────────────────────────────────────────────────────────────────
def send_email(
    to_email:     Iterable[str],
    subject:      str,
    message_text: str,
    attachments:  Optional[Iterable[str]] = None,
    is_html:      bool = False,
    cc:           Optional[Iterable[str]] = None,
    bcc:          Optional[Iterable[str]] = None,
) -> Result:
    """Send an email via Gmail SMTP. Never raises — returns Result."""
    if not SENDER_EMAIL or not APP_PASSWORD:
        return Result(
            False,
            f"Email credentials are not configured. "
            f"Please add PROVA_EMAIL and PROVA_APP_PASSWORD to {_ENV_PATH}."
        )

    to_list  = _ensure_list(to_email)
    cc_list  = _ensure_list(cc)
    bcc_list = _ensure_list(bcc)
    all_rcpt = to_list + cc_list + bcc_list

    if not all_rcpt:
        return Result(False, "No recipients provided.")

    invalid = [a for a in all_rcpt if not _is_valid_email(a)]
    if invalid:
        return Result(False, f"Invalid email address: {', '.join(invalid)}.")

    attach_list = _ensure_list(attachments)
    for filepath in attach_list:
        err = _validate_attachment(filepath)
        if err:
            return Result(False, err)

    msg            = MIMEMultipart()
    msg["From"]    = formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"]      = ", ".join(to_list)
    msg["Subject"] = subject
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.attach(MIMEText(message_text, "html" if is_html else "plain", "utf-8"))
    for filepath in attach_list:
        _add_attachment(msg, filepath)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=_SMTP_TIMEOUT) as smtp:
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.sendmail(SENDER_EMAIL, all_rcpt, msg.as_string())
        log.info("Email sent to %s | subject: %s", to_list, subject)
        return Result(True, "Email sent successfully.")
    except smtplib.SMTPAuthenticationError:
        return Result(False, "Authentication failed. Check your Gmail app password in the .env file.")
    except smtplib.SMTPRecipientsRefused as e:
        return Result(False, f"Recipient refused: {e}")
    except TimeoutError:
        return Result(False, "Email timed out. Check your internet connection.")
    except Exception as e:
        log.exception("Unexpected SMTP error")
        return Result(False, f"Failed to send email: {e}")


# ─────────────────────────────────────────────────────────────────
# GATE HELPERS
# ─────────────────────────────────────────────────────────────────

def _hold(gate_event: Optional[threading.Event]) -> None:
    """Set gate_event (blocks main voice loop). Idempotent, safe if None."""
    if gate_event is not None:
        gate_event.set()


def _release(gate_event: Optional[threading.Event]) -> None:
    """Clear gate_event (allows main voice loop). Safe if None."""
    if gate_event is not None:
        gate_event.clear()


# ─────────────────────────────────────────────────────────────────
# YES / NO WORD SETS
# ─────────────────────────────────────────────────────────────────
_YES_WORDS = frozenset({
    "yes", "yeah", "yep", "yup", "sure", "ok", "okay", "correct",
    "send", "do it", "confirm", "go ahead", "right", "done", "proceed",
    "agree", "sounds good", "that's right", "thats right",
    # Indian-English STT variants
    "han", "haan", "ya", "yaa", "yeh", "yas", "ha",
})
_NO_WORDS = frozenset({
    "no", "nope", "cancel", "stop", "abort", "nah",
    "never mind", "nevermind", "reject", "negative", "wrong",
    # Indian-English STT variants
    "nahi", "na",
})


# ─────────────────────────────────────────────────────────────────
# LOW-LEVEL VOICE PRIMITIVES
# ─────────────────────────────────────────────────────────────────

def _gated_listen(
    prompt:     str,
    speak_fn:   Callable,
    listen_fn:  Callable,
    gate_event: Optional[threading.Event],
    retries:    int = 3,
    retry_msg:  Optional[str] = None,
) -> Optional[str]:
    """
    Hold gate → speak prompt → listen.
    Retries on silence up to `retries` times.
    Re-holds gate immediately after listen_fn (which clears it internally).
    Returns stripped transcription or None.
    """
    for attempt in range(retries):
        _hold(gate_event)
        speak_fn(prompt if attempt == 0 else (retry_msg or f"Sorry, didn't catch that. {prompt}"))
        time.sleep(_SMALL_PAUSE)

        response = listen_fn()
        _hold(gate_event)   # re-hold immediately after listen_fn clears it

        if response and response.strip():
            return response.strip()

    return None


def _gated_confirm(
    prompt:     str,
    speak_fn:   Callable,
    listen_fn:  Callable,
    gate_event: Optional[threading.Event],
) -> bool:
    """
    Hold gate → speak prompt → listen for yes/no.
    Retries once on silence or ambiguous answer.
    Includes fuzzy matching for STT mishearings.
    Returns True/False (safe default on failure: False).
    """
    for attempt in range(2):
        _hold(gate_event)
        speak_fn(prompt if attempt == 0 else "Please say yes or no.")
        time.sleep(_SMALL_PAUSE)

        response = listen_fn()
        _hold(gate_event)

        if not response:
            if attempt == 0:
                speak_fn("I didn't hear anything.")
            continue

        r     = response.lower().strip()
        words = set(r.split())

        if words & _YES_WORDS: return True
        if words & _NO_WORDS:  return False

        # Fuzzy fallback for phonetic STT mishearings ("yeas", "noo", etc.)
        _fuzz = None
        try:
            from thefuzz import fuzz as _fuzz
        except ImportError:
            try:
                from fuzzywuzzy import fuzz as _fuzz
            except ImportError:
                pass

        if _fuzz is not None:
            best_yes = max(_fuzz.ratio(r, w) for w in {"yes", "yeah", "okay", "correct"})
            best_no  = max(_fuzz.ratio(r, w) for w in {"no", "nope", "cancel"})
            log.debug("Confirm fuzzy: yes=%d no=%d for %r", best_yes, best_no, r)
            if best_yes >= 70 and best_yes > best_no: return True
            if best_no  >= 70 and best_no  > best_yes: return False

        if attempt == 0:
            speak_fn("I didn't catch that.")

    return False  # safe default — never send without explicit yes


# ─────────────────────────────────────────────────────────────────
# FIELD COLLECTORS
# ─────────────────────────────────────────────────────────────────

def _collect_field(
    field_name: str,
    prompt:     str,
    speak_fn:   Callable,
    listen_fn:  Callable,
    gate_event: Optional[threading.Event],
) -> Optional[str]:
    """
    Collect a single short field (subject, recipient name, etc.) via voice.

    Flow:
      1. Ask by voice (2 attempts max)
      2. "I heard X — is that correct?"
      3. If rejected → one more voice attempt
      4. Still rejected → typed terminal fallback (confirmed by voice)

    Returns confirmed value or None if user abandons.
    """
    for voice_attempt in range(2):
        heard = _gated_listen(
            prompt if voice_attempt == 0 else f"Let's try again. {prompt}",
            speak_fn, listen_fn, gate_event,
            retries=2,
        )
        if heard:
            confirmed = _gated_confirm(
                f"I heard: {heard}. Is that correct?",
                speak_fn, listen_fn, gate_event,
            )
            if confirmed:
                return heard

    # Typed fallback — ask user to type in the chat box
    speak_fn(f"No problem — please type the {field_name} in the chat box and press Enter.")
    typed = _gated_listen("", speak_fn, listen_fn, gate_event, retries=3)
    if not typed:
        speak_fn(f"Nothing received.")
        return None

    confirmed = _gated_confirm(
        f"I got: {typed}. Is that correct?",
        speak_fn, listen_fn, gate_event,
    )
    return typed if confirmed else None


def _collect_body(
    speak_fn:   Callable,
    listen_fn:  Callable,
    gate_event: Optional[threading.Event],
) -> Optional[str]:
    """
    Collect email body with multi-sentence dictation.

    Flow:
      1. "What's the message?" → listen → append
      2. "Anything to add? Say done when finished." → loop
      3. Confirm full assembled body
      4. If rejected → offer typed fallback or re-dictate
    """
    _DONE_PHRASES = frozenset({
        "done", "that's it", "thats it", "nothing else",
        "that's all", "thats all", "send it", "finish", "finished",
        "nothing more", "end",
    })

    parts: list[str] = []

    while True:
        prompt = "What's the message?" if not parts else "Anything to add? Say done when finished."
        heard  = _gated_listen(prompt, speak_fn, listen_fn, gate_event, retries=3)

        if not heard:
            if parts:
                break   # silence after content = done
            break       # silence on first ask → fall to typed below

        # "done" signal only counts after we have something
        heard_lower = heard.lower().strip()
        if parts and any(d in heard_lower for d in _DONE_PHRASES):
            break

        parts.append(heard)

        more = _gated_confirm(
            "Got it. Do you want to add anything else to the message?",
            speak_fn, listen_fn, gate_event,
        )
        if not more:
            break

    # Nothing collected → ask user to type in chat box
    if not parts:
        speak_fn("Please type your message in the chat box and press Enter.")
        typed = _gated_listen("", speak_fn, listen_fn, gate_event, retries=3)
        return typed or None

    full_body = " ".join(parts)

    # Confirm assembled body
    preview   = full_body if len(full_body) <= 150 else full_body[:150] + "..."
    confirmed = _gated_confirm(
        f"Message: {preview}. Is that correct?",
        speak_fn, listen_fn, gate_event,
    )
    if confirmed:
        return full_body

    # Rejected — offer typed or re-dictate
    retype = _gated_confirm(
        "Would you like to type the message instead?",
        speak_fn, listen_fn, gate_event,
    )
    if retype:
        speak_fn("Please type your message in the chat box and press Enter.")
        typed = _gated_listen("", speak_fn, listen_fn, gate_event, retries=3)
        if typed:
            confirmed_typed = _gated_confirm(
                f"I got: {typed[:100]}{'...' if len(typed) > 100 else ''}. Is that correct?",
                speak_fn, listen_fn, gate_event,
            )
            return typed if confirmed_typed else None
        return None

    # Re-dictate from scratch (recursive — bounded by user patience)
    speak_fn("Let's start the message again.")
    return _collect_body(speak_fn, listen_fn, gate_event)


def _collect_recipient(
    initial_name: str,
    speak_fn:     Callable,
    listen_fn:    Callable,
    gate_event:   Optional[threading.Event],
) -> Optional[tuple[str, str]]:
    """
    Resolve recipient to (display_name, email_address).

    - Known contact or bare email → immediate return (voice works fine for names)
    - Unknown contact → go STRAIGHT to typed input for the address.

    Email addresses are NOT collected by voice. STT is far too unreliable for
    addresses (dots, @, mixed letters/numbers). Typed input is used every time
    so the address is always exactly what the user intended.

    After a new address is typed and validated, ProVA offers to save it as a
    contact so it never needs to be typed again.
    """
    name = initial_name.strip()

    # Recipient name not given in original command — ask by voice
    if not name:
        name = _gated_listen(
            "Who should I send the email to?",
            speak_fn, listen_fn, gate_event, retries=3,
        ) or ""
        if not name:
            speak_fn("No recipient given. Cancelling.")
            return None

    # Known contact or already a valid email address — done
    resolved = _resolve_recipient(name)
    if resolved:
        log.info("Resolved '%s' → %s", name, resolved)
        return (name, resolved)

    # Unknown contact — ask for address via TYPED INPUT only
    # (voice dictation of email addresses is unreliable with all STT engines)
    _hold(gate_event)
    speak_fn(
        f"I don't have an address for {name}. "
        "Please type their email address in the chat box."
    )

    for attempt in range(3):
        typed = _gated_listen("", speak_fn, listen_fn, gate_event, retries=2)

        if not typed:
            if attempt < 2:
                speak_fn("Nothing received. Please type the address in the chat box.")
                continue
            speak_fn("No address provided. Cancelling.")
            return None

        # Normalize the typed address (handles "john dot smith at gmail dot com"
        # if user still types spoken-style, or just cleans whitespace)
        normalized = _normalize_spoken_email(typed)

        if not _is_valid_email(normalized):
            if attempt < 2:
                speak_fn(f"That doesn't look like a valid email: {normalized}. Please try again.")
                continue
            speak_fn("Couldn't get a valid address. Cancelling.")
            return None

        log.info("Typed email for '%s': %s", name, normalized)

        # Offer to save — voice confirm is fine here (it's a simple yes/no)
        save = _gated_confirm(
            f"Got it — {normalized}. "
            f"Should I save {name} as a contact so you don't have to type next time?",
            speak_fn, listen_fn, gate_event,
        )
        if save:
            _save_contact(name, normalized)
            speak_fn(f"Saved. I'll remember {name} from now on.")

        return (name, normalized)

    speak_fn("Couldn't get a valid address. Cancelling.")
    return None


# ─────────────────────────────────────────────────────────────────
# MAIN FLOW
# ─────────────────────────────────────────────────────────────────

def _run_email_flow(
    cmd,
    speak_fn:   Callable,
    listen_fn:  Callable,
    gate_event: Optional[threading.Event],
) -> None:
    """Full email composition flow. Gate held throughout by handle()."""

    if (cmd.action or "") == "cancel":
        speak_fn("Email cancelled.")
        return

    # ── Step 1: Recipient ─────────────────────────────────────────
    recipient = _collect_recipient(
        initial_name=(cmd.target or "").strip(),
        speak_fn=speak_fn, listen_fn=listen_fn, gate_event=gate_event,
    )
    if recipient is None:
        return
    recipient_name, recipient_email = recipient

    speak_fn(f"Composing email to {recipient_name}.")

    # ── Step 2: Subject ───────────────────────────────────────────
    subject = _collect_field(
        field_name="subject",
        prompt="What's the subject?",
        speak_fn=speak_fn, listen_fn=listen_fn, gate_event=gate_event,
    )
    if subject is None:
        speak_fn("Email cancelled.")
        return

    # ── Step 3: Body ──────────────────────────────────────────────
    body = _collect_body(speak_fn, listen_fn, gate_event)
    if not body:
        speak_fn("Email cancelled.")
        return

    # ── Step 4: Attachments (future scope — from cmd.extra) ───────
    attachments = cmd.extra.get("attachments", [])

    # ── Step 5: Summary + confirm loop (with edit support) ────────
    while True:
        attach_note = ""
        if attachments:
            n = len(attachments)
            attach_note = f" with {n} attachment{'s' if n != 1 else ''}"

        preview = body if len(body) <= 120 else body[:120] + "..."
        summary = (
            f"Ready to send to {recipient_name}. "
            f"Subject: {subject}. "
            f"Message: {preview}{attach_note}. "
            f"Shall I send it?"
        )
        send_confirmed = _gated_confirm(summary, speak_fn, listen_fn, gate_event)

        if send_confirmed:
            break

        # Ask what to change
        what = _gated_listen(
            "What would you like to change? Say subject, message, or cancel.",
            speak_fn, listen_fn, gate_event, retries=2,
        )
        if not what:
            speak_fn("Email cancelled. Nothing was sent.")
            return

        what_lower = what.lower()

        if "subject" in what_lower:
            new_sub = _collect_field("subject", "What's the new subject?",
                                     speak_fn, listen_fn, gate_event)
            if new_sub:
                subject = new_sub

        elif any(w in what_lower for w in ("message", "body", "content", "text", "mail")):
            new_body = _collect_body(speak_fn, listen_fn, gate_event)
            if new_body:
                body = new_body

        else:
            # Treat as cancel
            speak_fn("Email cancelled. Nothing was sent.")
            return

    # ── Step 6: Send ──────────────────────────────────────────────
    speak_fn("Sending...")
    result = send_email(
        to_email     = recipient_email,
        subject      = subject,
        message_text = body,
        attachments  = attachments if attachments else None,
    )
    speak_fn(result.message)
    if result.success:
        log.info("Email flow complete: to=%s subject=%s", recipient_email, subject)
    else:
        log.warning("Email send failed: %s", result.message)


# ─────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────

def handle(
    cmd,
    speak_fn:   Callable[[str], None],
    listen_fn:  Callable[[], Optional[str]],
    gate_event: Optional[threading.Event] = None,
) -> None:
    """
    Entry point called by router via run_async.

    Holds gate_event for the ENTIRE email session — the main voice loop
    cannot race in and treat subject/body dictation as new commands.

    Router call (must include gate_event):
        _exec(_get_em_handle(), cmd, speak_fn, listen_fn, gate_event, on_done=on_done)
    """
    _hold(gate_event)
    try:
        _run_email_flow(cmd, speak_fn, listen_fn, gate_event)
    except Exception as e:
        log.exception("Email flow error: %s", e)
        speak_fn("Something went wrong with the email. Please try again.")
    finally:
        # Brief cooldown drains residual mic audio before the main loop resumes
        time.sleep(0.8)
        _release(gate_event)


# ─────────────────────────────────────────────────────────────────
# STT EMAIL ADDRESS NORMALISATION
# "john at gmail dot com" → "john@gmail.com"
# ─────────────────────────────────────────────────────────────────

def _collapse_spaced_letters(s: str) -> str:
    """
    Collapse letter-by-letter STT dictation into a compact string.
    'n i d a 04' → 'nida04'
    """
    tokens = s.split()
    result = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        if len(tok) == 1:
            run = tok
            while i + 1 < len(tokens) and len(tokens[i + 1]) == 1:
                i += 1
                run += tokens[i]
            result.append(run)
        else:
            result.append(tok)
        i += 1
    return "".join(result)


_EMAIL_NOISE = re.compile(
    # "a" excluded — could be a legitimate single-letter dictation
    r'\b(?:and|the|is|are|was)\b', re.IGNORECASE
)


def _normalize_spoken_email(text: str) -> str:
    """
    Convert spoken / STT-mangled email to standard format.

    Handles:
      "at the rate" / "at the"  → @
      " at "                    → @
      " dot "                   → .
      " dash " / " hyphen "     → -
      " underscore "            → _
      "n i d a 04"              → "nida04"
      noise words (and/the)     → stripped from local part only
    """
    t = text.lower().strip()

    # Spoken @ variants (longest first)
    t = re.sub(r'\bat\s+the\s+rate\b', '@', t)
    t = re.sub(r'\bat\s+the\b',        '@', t)
    t = re.sub(r'\s+at\s+',            '@', t)

    # Other symbol words
    t = re.sub(r'\s+dot\s+',        '.', t)
    t = re.sub(r'\s+dash\s+',       '-', t)
    t = re.sub(r'\s+hyphen\s+',     '-', t)
    t = re.sub(r'\s+underscore\s+', '_', t)

    if '@' in t:
        local, domain = t.split('@', 1)
        local  = _EMAIL_NOISE.sub(' ', local).strip()
        local  = _collapse_spaced_letters(local).replace(' ', '')
        domain = domain.replace(' ', '')
        return local + '@' + domain

    # No @ found — strip spaces wholesale
    return _EMAIL_NOISE.sub('', t).replace(' ', '')
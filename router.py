"""
router.py — ProVA module dispatcher.

Fixes in this version:
  - Import path corrected: modules.excel_module.handle → modules.excel_module.handle
  - reminder handle() signature updated: handle(cmd, speak_fn, runner)
  - xl_handle import path fixed: excel_module not excel_module (was excel_module)
  - All imports wrapped in a lazy-load pattern so missing optional deps
    (win32com, winsdk) don't crash the whole app at startup

Confirmation race fix (this version):
  - _voice_confirm now accepts a module_active_event (threading.Event) and holds
    it for the entire confirm window (speak prompt + listen), not just during listen_fn.
  - Without this, the main loop races in between speak_fn("please say yes...")
    returning and listen_fn() being called, grabs the mic first, and captures
    the user's "yes" as a stray unknown command — producing "I didn't catch that"
    followed by deletion a few seconds later when listen_fn finally gets the mic.
  - voice_module passes its _module_active event via the new gate_event param to route().
  - Fallback: if gate_event is None (old callers), behaviour is unchanged.
"""

from __future__ import annotations

import logging
import threading
import time
from typing import Callable, Optional

from parser import Command

log = logging.getLogger("ProVA.Router")

_reminder_runner = None


def init(reminder_runner) -> None:
    """Called once from voice_module.run() after startup."""
    global _reminder_runner
    _reminder_runner = reminder_runner


# ─────────────────────────────────────────────────────────────────
# LAZY MODULE IMPORTS
# Deferred so a missing optional dep (pywin32, winsdk) doesn't
# crash ProVA startup — only the affected module fails.
# ─────────────────────────────────────────────────────────────────
def _get_cc_handle():
    from modules.computer_control import handle as cc_handle
    return cc_handle

def _get_fm_handle():
    from modules.file_manager import handle as fm_handle
    return fm_handle

def _get_rm_handle():
    from modules.reminder_module import handle as rm_handle
    return rm_handle

def _get_em_handle():
    from modules.email_module import handle as em_handle
    return em_handle

def _get_xl_handle():
    from modules.excel_module.handle import handle as xl_handle
    return xl_handle

def _get_help_handle():
    from modules.help_module import handle as help_handle
    return help_handle

def _get_start_reminder_system():
    from modules.reminder_module import start_reminder_system
    return start_reminder_system


# ─────────────────────────────────────────────────────────────────
# ASYNC RUNNER
# ─────────────────────────────────────────────────────────────────
def run_async(fn, *args, on_done: Callable = None, **kwargs) -> threading.Thread:
    """
    Fire-and-forget background thread.
    on_done(): called when the module function returns (success or error).
    """
    def _safe():
        try:
            fn(*args, **kwargs)
        except Exception as e:
            log.exception("Crash in '%s': %s", getattr(fn, "__name__", "?"), e)
            for arg in args:
                if callable(arg):
                    try: arg("Something went wrong. Please try again.")
                    except Exception: pass
                    break
        finally:
            if on_done:
                try: on_done()
                except Exception: pass
    t = threading.Thread(target=_safe, daemon=True)
    t.start()
    return t


# ─────────────────────────────────────────────────────────────────
# VOICE CONFIRMATION
# ─────────────────────────────────────────────────────────────────
# Cooldown after confirmation resolves — prevents the main loop from
# immediately capturing residual mic audio (room echo, TTS tail, etc.)
_POST_CONFIRM_COOLDOWN = 1.2   # seconds

def _voice_confirm(
    speak_fn:   Callable,
    listen_fn:  Callable,
    gate_event: Optional[threading.Event] = None,
) -> bool:
    """
    Ask the user yes/no by voice for a destructive operation.

    Race-safe design
    ─────────────────
    The confirmation window has three phases:
      1. speak_fn("Please say yes…")  — TTS plays the prompt
      2. listen_fn()                  — mic captures the answer
      3. post-confirm cooldown        — drains any residual audio

    The bug: _module_active is only set INSIDE listen_fn(), so there is a
    gap between speak_fn() returning and listen_fn() being called. The main
    loop can race into that gap, grab the mic, and capture "yes" as an
    unknown command → "I didn't catch that", while the deletion still
    proceeds moments later when listen_fn finally gets the mic.

    Fix: if gate_event (_module_active from voice_module) is provided, we
    SET it before speaking and only CLEAR it after the post-confirm cooldown.
    listen_fn()'s own set/clear of the same event is harmless (idempotent).
    The main loop checks `if _module_active.is_set(): continue` so it will
    not touch the mic for the entire confirm window.

    Retries once on silence or unclear answer.
    """
    # Indian-English STT often returns phonetic approximations of yes/no.
    # Exact word-set match handles the common ones; fuzzy fallback catches the rest.
    yes_words = {
        # Standard
        "yes", "yeah", "yep", "yup", "confirm", "sure",
        "go ahead", "ok", "okay", "correct", "do it", "done",
        # Indian-English STT variants (Google en-IN mishearings)
        "yas", "ya", "yaa", "yeh", "yes please", "yes yes",
        "han", "haan", "ha",           # Hindi affirmatives
        "proceed", "continue", "agree",
    }
    no_words = {
        # Standard
        "no", "nope", "cancel", "stop", "abort",
        "never mind", "nevermind", "nah", "don't",
        # Indian-English STT variants
        "nahi", "nope nope", "na", "no no",
        "reject", "decline",
    }

    # ── Hold the gate for the entire confirm window ───────────────
    if gate_event is not None:
        gate_event.set()

    try:
        for attempt in range(2):
            prompt = (
                "Please say yes to confirm, or no to cancel."
                if attempt == 0
                else "Say yes to confirm, or no to cancel."
            )
            speak_fn(prompt)
            time.sleep(0.3)

            text = listen_fn()   # listen_fn also sets/clears gate_event — harmless

            # Re-hold the gate immediately after listen_fn clears it, so the
            # cooldown + any retry speak is still protected.
            if gate_event is not None:
                gate_event.set()

            if not text:
                if attempt == 0:
                    # One retry on silence before giving up
                    speak_fn("I didn't hear anything.")
                    continue
                speak_fn("No answer heard. Cancelled.")
                return False

            log.info("Confirmation response (attempt %d): %r", attempt + 1, text)
            words = set(text.lower().split())
            text_lower = text.lower().strip()

            if words & yes_words:
                return True

            if words & no_words:
                speak_fn("Cancelled.")
                return False

            # Fuzzy fallback — catches STT mishearings like "yas ya", "yeas", "knoe"
            # that aren't exact matches but are phonetically close
            _yes_fuzzy = {"yes", "yeah", "confirm", "okay"}
            _no_fuzzy  = {"no", "nope", "cancel"}
            try:
                from fuzzywuzzy import fuzz as _fuzz
            except ImportError:
                try:
                    from thefuzz import fuzz as _fuzz  # type: ignore
                except ImportError:
                    _fuzz = None

            if _fuzz is not None:
                best_yes = max(_fuzz.ratio(text_lower, w) for w in _yes_fuzzy)
                best_no  = max(_fuzz.ratio(text_lower, w) for w in _no_fuzzy)
                log.info("Confirm fuzzy: yes=%d no=%d for %r", best_yes, best_no, text_lower)
                if best_yes >= 70 and best_yes > best_no:
                    log.info("Confirm: fuzzy yes match (%d) for %r", best_yes, text_lower)
                    return True
                if best_no >= 70 and best_no > best_yes:
                    log.info("Confirm: fuzzy no match (%d) for %r", best_no, text_lower)
                    speak_fn("Cancelled.")
                    return False

            # Unclear response — retry once
            if attempt == 0:
                speak_fn("I didn't catch that.")
                continue

        speak_fn("Not sure about that. Cancelled to be safe.")
        return False

    finally:
        # ── Post-confirm cooldown before releasing the gate ───────
        # Drains any residual audio (room echo, late mic packets) so the
        # main loop doesn't immediately capture them as a new command.
        time.sleep(_POST_CONFIRM_COOLDOWN)
        if gate_event is not None:
            gate_event.clear()




# ─────────────────────────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────────────────────────
def route(
    cmd:        Command,
    speak_fn:   Callable[[str], None],
    listen_fn:  Callable[[], Optional[str]],
    on_done:    Callable = None,
    gate_event: Optional[threading.Event] = None,
    status_fn:  Optional[Callable[[str], None]] = None,
) -> None:
    """
    Dispatch cmd to the appropriate module.

    gate_event: the _module_active threading.Event from voice_module.
        When provided, _voice_confirm holds it for the entire confirmation
        window (speak + listen + cooldown) to prevent the main loop racing
        in and consuming the user's "yes/no" as an unknown command.
        Pass voice_module._module_active here from the call site.

    status_fn: optional callback to push UI status strings. Used to fire
        "EXECUTING" once the intent is resolved and a module is dispatched,
        distinct from "THINKING" (parsing) so users see the task is running.
    """
    log.info("Routing → %s", cmd.summary())

    def _exec(fn, *args, **kwargs):
        """Fire EXECUTING status then dispatch — keeps call sites clean."""
        if status_fn:
            status_fn("EXECUTING")
        run_async(fn, *args, **kwargs)

    if cmd.confidence < 75 and cmd.intent not in ("unknown", "system"):
        speak_fn("I'm not quite sure what you meant. Could you rephrase?")
        if on_done: on_done()
        return

    if cmd.intent == "unknown":
        speak_fn("I didn't catch that. Could you rephrase?")
        if on_done: on_done()
        return

    if cmd.intent == "help":
        _exec(_get_help_handle(), cmd, speak_fn, on_done=on_done)

    elif cmd.intent == "computer_control":
        _exec(_get_cc_handle(), cmd, speak_fn, on_done=on_done)

    elif cmd.intent == "file_manager":
        # Gate-armed confirm: set _module_active BEFORE speaking the prompt so the
        # main listen loop cannot race in and steal "yes"/"no" between the time
        # file_manager speaks "Delete folder X?" and _voice_confirm calls listen_fn.
        def confirm_fn(prompt: str = "") -> bool:
            if gate_event is not None:
                gate_event.set()
            if prompt:
                speak_fn(prompt)
            return _voice_confirm(speak_fn, listen_fn, gate_event)
        _exec(_get_fm_handle(), cmd, speak_fn, confirm_fn, on_done=on_done)

    elif cmd.intent == "email":
        _exec(_get_em_handle(), cmd, speak_fn, listen_fn, gate_event, on_done=on_done)

    elif cmd.intent == "excel":
        _exec(_get_xl_handle(), cmd, speak_fn, listen_fn, on_done=on_done)

    elif cmd.intent == "reminder":
        if _reminder_runner is None:
            speak_fn("The reminder system isn't running. Please restart ProVA.")
            if on_done: on_done()
            return
        _exec(_get_rm_handle(), cmd, speak_fn, _reminder_runner, on_done=on_done)

    elif cmd.intent == "system":
        if cmd.action == "exit":
            speak_fn("Goodbye! ProVA is shutting down.")
            if on_done: on_done()
            raise SystemExit
        elif cmd.action == "sleep":
            speak_fn("Going to sleep. Say 'hey ProVA' to wake me.")
        elif cmd.action == "pause":
            speak_fn("Pausing. Click Resume in the window when you're ready.")
        if on_done: on_done()

    else:
        speak_fn("I didn't understand that command.")
        if on_done: on_done()
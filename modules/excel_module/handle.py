"""
ProVA — excel_module/handle.py
Conversation flow for dashboard creation and file open.
Asks for filename, listens back, finds file or shows picker,
then builds and opens dashboard.
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable, Optional

from .commands    import Session
from .file_finder import find_data_file, pick_file_dialog

log = logging.getLogger("ProVA.Excel.Handle")

_session: Optional[Session] = None

# Words that are part of the voice command itself, not the filename.
# "create excel dashboard for restaurant sales data"
#   → strip: create, excel, dashboard, for  → "restaurant sales data"  ✓
# "create an excel dashboard"
#   → strip: an, excel, dashboard            → ""  → ask/picker  ✓
_EXCEL_CMD_WORDS = frozenset({
    "excel", "dashboard", "create", "make", "build", "open",
    "an", "a", "the", "for", "from", "using", "with", "my",
    "me", "please", "generate", "produce",
})


def _clean_filename(raw: str) -> str:
    """
    Strip excel command words from the parser-extracted target.
    Returns the cleaned filename, or an empty string if nothing meaningful remains.
    """
    if not raw:
        return ""
    words = [w for w in raw.lower().split() if w not in _EXCEL_CMD_WORDS]
    return " ".join(words).strip()


def _get_session() -> Session:
    global _session
    if _session is None:
        _session = Session()
    return _session


def handle(
    cmd,
    speak_fn:  Callable[[str], None],
    listen_fn: Callable[[], Optional[str]],
) -> None:
    """
    Entry point called by router via run_async.
    Signature: (cmd, speak_fn, listen_fn)
    """
    action   = (cmd.action or "dashboard").lower()
    # Clean the extracted target before doing anything with it so that
    # phrases like "create an excel dashboard" don't leak command words
    # into the file-search query and produce false-positive matches.
    filename = _clean_filename((cmd.target or "").strip())

    # ── Open file (just launch in Excel, no dashboard) ─────────
    if action == "open":
        _handle_open(filename, speak_fn, listen_fn)
        return

    # ── Dashboard creation ──────────────────────────────────────
    _handle_dashboard(filename, speak_fn, listen_fn)


# ─────────────────────────────────────────────────────────────────
# OPEN
# ─────────────────────────────────────────────────────────────────

def _handle_open(filename: str, speak_fn, listen_fn) -> None:
    import os, subprocess

    # Empty filename after stripping command words ("open excel") means
    # the user wants to launch Excel the application, not open a data file.
    if not filename:
        _excel_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
            r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
        ]
        for path in _excel_paths:
            if os.path.exists(path):
                speak_fn("Opening Excel.")
                try:
                    subprocess.Popen([path])
                except Exception as e:
                    speak_fn(f"Couldn't launch Excel. {e}")
                return
        # Excel not found at known paths — fall through to file picker
        speak_fn("I couldn't find Excel on this machine. Which file would you like to open?")

    data_path = _resolve_file(filename, speak_fn, listen_fn)
    if data_path is None:
        return
    speak_fn(f"Opening {data_path.name}.")
    try:
        os.startfile(str(data_path))
    except Exception as e:
        speak_fn(f"Could not open the file. {e}")


# ─────────────────────────────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────────────────────────────

def _handle_dashboard(filename: str, speak_fn, listen_fn) -> None:
    session = _get_session()

    # ── Ask for filename if not provided (or only command words were given) ─
    if not filename:
        speak_fn("Which file should I use for the dashboard?")
        response = listen_fn()
        if not response:
            speak_fn("I didn't catch that. Dashboard cancelled.")
            return
        # Clean the spoken response too, in case they repeat command words
        filename = _clean_filename(response.strip()) or response.strip()

    # ── Resolve file ────────────────────────────────────────────
    data_path = _resolve_file(filename, speak_fn, listen_fn)
    if data_path is None:
        return

    # ── Load ────────────────────────────────────────────────────
    speak_fn(f"Found {data_path.name}. Loading data now.")
    try:
        session.load(str(data_path))
    except FileNotFoundError:
        speak_fn(f"Could not find {data_path.name}.")
        return
    except ValueError as e:
        speak_fn(str(e))
        return
    except Exception as e:
        log.exception("Load failed")
        speak_fn(f"Failed to load {data_path.name}.")
        return

    n_rows = len(session.df)
    n_cols = len(session.df.columns)
    speak_fn(
        f"Loaded {n_rows:,} rows and {n_cols} columns. "
        "Building your dashboard now — please wait."
    )

    # ── Build ────────────────────────────────────────────────────
    try:
        out_path = session.create_dashboard()
        speak_fn(
            f"Dashboard is ready. Opening {out_path.name} in Excel now."
        )
    except ValueError as e:
        speak_fn(str(e))
    except Exception as e:
        log.exception("Dashboard build failed")
        speak_fn(f"Something went wrong building the dashboard. {e}")


# ─────────────────────────────────────────────────────────────────
# FILE RESOLUTION
# ─────────────────────────────────────────────────────────────────

def _resolve_file(
    filename: str,
    speak_fn: Callable,
    listen_fn: Callable,
) -> Optional[Path]:
    """
    Try to find the file by spoken name.
    If a cleaned, non-empty filename is given → search and return if found.
    If not found, or filename is empty → open file picker.
    Returns None if user cancels or nothing is found.
    """
    if filename:
        data_path = find_data_file(filename)
        if data_path:
            speak_fn(f"Found {data_path.name}.")
            return data_path
        else:
            speak_fn(
                f"I couldn't find a file called {filename}. "
                "Opening a file picker — please select your file."
            )
    else:
        speak_fn("Opening a file picker so you can select your file.")

    data_path = pick_file_dialog()
    if data_path is None:
        speak_fn("No file selected. Cancelled.")
        return None

    return data_path
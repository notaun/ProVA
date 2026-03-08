"""
modules/help_module.py — ProVA spoken help guide.

Triggered by: "help", "help with email", "what can you do", etc.

Speaks a concise overview of ProVA's capabilities.
Specific section: "help with <module>" speaks only that module's guide.

MODULE_HELP is also imported by the UI to render a help panel.
"""

from __future__ import annotations

import logging
from typing import Callable

log = logging.getLogger("ProVA.Help")


# ─────────────────────────────────────────────────────────────────
# HELP CONTENT
# ─────────────────────────────────────────────────────────────────
MODULE_HELP: dict[str, str] = {

    "general": (
        "I'm ProVA, your productivity voice assistant. "
        "Say 'hey ProVA' to wake me up, then give a command. "
        "I can open apps and search the web, "
        "manage your files and folders, "
        "send and compose emails, "
        "build Excel dashboards from your data, "
        "and set reminders and alarms. "
        "Say 'help with' followed by a module name to learn more — "
        "for example, help with email, or help with files."
    ),

    "computer": (
        "Computer control. "
        "Say 'open' followed by any app — for example, open Chrome, open Notepad. "
        "Say 'search for' or 'google' followed by your query to search the web. "
        "For example: google latest news, or search for Python tutorials."
    ),

    "files": (
        "File manager. "
        "I can create, delete, rename, copy, move, and find files and folders. "
        "Examples: create file notes dot txt, "
        "delete folder old projects, "
        "rename report to final report, "
        "copy file budget to documents, "
        "move notes to desktop, "
        "or find files in downloads. "
        "I always ask for confirmation before deleting anything."
    ),

    "email": (
        "Email. "
        "Say 'send email to' followed by a contact name. "
        "For example: send email to John. "
        "I'll then ask for the subject and body. "
        "You can include a hint directly: "
        "send email to Sarah about the project deadline. "
        "Make sure your email credentials are set up in the dot env file first."
    ),

    "excel": (
        "Excel and dashboards. "
        "Say 'create dashboard from' followed by a filename. "
        "For example: create dashboard from sales data. "
        "I'll search your Desktop, Downloads, Documents, and ProVA folders. "
        "If I can't find the file, a picker window will open. "
        "You can specify columns: create dashboard from sales data, metric revenue, category region. "
        "Other commands: analyze data, make chart, or open dashboard."
    ),

    "reminders": (
        "Reminders and alarms. "
        "Set a reminder: remind me to call the client at 3 pm. "
        "Countdown reminder: remind me in 20 minutes to check the oven. "
        "Daily alarm: set daily alarm at 8 am. "
        "View reminders: list reminders or show reminders. "
        "Cancel a reminder: cancel reminder — I'll ask which one."
    ),
}

_ALIASES: dict[str, str] = {
    "computer": "computer", "app": "computer",    "apps": "computer",
    "web": "computer",      "search": "computer",
    "file": "files",        "files": "files",     "folder": "files",
    "folders": "files",     "file manager": "files",
    "email": "email",       "mail": "email",      "emails": "email",
    "excel": "excel",       "dashboard": "excel", "chart": "excel",
    "data": "excel",        "spreadsheet": "excel",
    "reminder": "reminders","reminders": "reminders",
    "alarm": "reminders",   "alarms": "reminders",
}


def _detect_section(text: str) -> str:
    t = text.lower()
    for alias, key in _ALIASES.items():
        if alias in t:
            return key
    return "general"


def handle(cmd, speak_fn: Callable[[str], None]) -> None:
    section = _detect_section(cmd.raw)
    log.info("Help section: %s", section)
    speak_fn(MODULE_HELP[section])
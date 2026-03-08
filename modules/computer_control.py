"""
ProVA — modules/computer_control.py

Fixes in this version:
  - f-string log calls → %-style (PEP 8 / W1203 warning fix)
  - `dict[str, list]` type hints → `Dict[str, List]` for Python 3.8 compat
  - `tuple[str, str]` return type → `Optional[Tuple[str, str]]`
  - `_find_browser` return type annotation fixed
  - Removed bare `except: pass` in _resolve_path (now logs warning)
  - `shutil.which(path.replace(".exe",""))` → only called if path has no sep
"""

import os
import re
import glob
import json
import shutil
import logging
import subprocess
import webbrowser
from pathlib import Path
from urllib.parse import quote_plus
from dataclasses import dataclass
from typing import Optional, Dict, List, Tuple

try:
    from fuzzywuzzy import fuzz, process as fuzz_process
except ImportError:
    try:
        from thefuzz import fuzz, process as fuzz_process  # type: ignore
    except ImportError:
        class fuzz:  # type: ignore
            @staticmethod
            def token_sort_ratio(a: str, b: str) -> int:
                return 100 if a.lower() in b.lower() else 0
        class fuzz_process:  # type: ignore
            @staticmethod
            def extractOne(query, choices, scorer=None):
                return (choices[0] if choices else "", 0)

log = logging.getLogger("ProVA.ComputerControl")

# ─────────────────────────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────────────────────────
_PF    = os.environ.get("ProgramFiles",       r"C:\Program Files")
_PF86  = os.environ.get("ProgramFiles(x86)",  r"C:\Program Files (x86)")
_LOCAL = os.environ.get("LOCALAPPDATA",        r"C:\Users\Default\AppData\Local")
_USER_APPS_FILE = str(Path(__file__).resolve().parent.parent / "data" / "user_apps.json")


@dataclass
class Result:
    success: bool
    message: str


ELEVATED_APPS = {
    "taskmgr.exe", "regedit.exe", "gpedit.msc",
    "compmgmt.msc", "diskmgmt.msc", "eventvwr.msc",
}

APP_LOOKUP: Dict[str, List[str]] = {
    "excel": [
        rf"{_PF}\Microsoft Office\root\Office16\EXCEL.EXE",
        rf"{_PF86}\Microsoft Office\root\Office16\EXCEL.EXE",
        rf"{_PF}\Microsoft Office\Office16\EXCEL.EXE",
    ],
    "word": [
        rf"{_PF}\Microsoft Office\root\Office16\WINWORD.EXE",
        rf"{_PF86}\Microsoft Office\root\Office16\WINWORD.EXE",
    ],
    "powerpoint": [
        rf"{_PF}\Microsoft Office\root\Office16\POWERPNT.EXE",
        rf"{_PF86}\Microsoft Office\root\Office16\POWERPNT.EXE",
    ],
    "outlook": [
        rf"{_PF}\Microsoft Office\root\Office16\OUTLOOK.EXE",
        rf"{_PF86}\Microsoft Office\root\Office16\OUTLOOK.EXE",
    ],
    "onenote": [
        rf"{_PF}\Microsoft Office\root\Office16\ONENOTE.EXE",
    ],
    "notepad":        ["notepad.exe"],
    "calculator":     ["calc.exe"],
    "paint":          ["mspaint.exe"],
    "task manager":   ["taskmgr.exe"],
    "file explorer":  ["explorer.exe"],
    "explorer":       ["explorer.exe"],
    "cmd":            ["cmd.exe"],
    "command prompt": ["cmd.exe"],
    "powershell":     ["powershell.exe"],
    "control panel":  ["control.exe"],
    "settings":       ["ms-settings:"],
    "snipping tool":  ["SnippingTool.exe"],
    "wordpad":        ["wordpad.exe"],
    "sticky notes":   ["stikynot.exe"],
    "registry":       ["regedit.exe"],
    "chrome": [
        rf"{_PF}\Google\Chrome\Application\chrome.exe",
        rf"{_PF86}\Google\Chrome\Application\chrome.exe",
    ],
    "edge": [
        rf"{_PF}\Microsoft\Edge\Application\msedge.exe",
        rf"{_PF86}\Microsoft\Edge\Application\msedge.exe",
    ],
    "firefox": [
        rf"{_PF}\Mozilla Firefox\firefox.exe",
        rf"{_PF86}\Mozilla Firefox\firefox.exe",
    ],
    "vscode": [
        rf"{_LOCAL}\Programs\Microsoft VS Code\Code.exe",
        rf"{_PF}\Microsoft VS Code\Code.exe",
    ],
    "visual studio code": [
        rf"{_LOCAL}\Programs\Microsoft VS Code\Code.exe",
    ],
    "visual studio": [
        rf"{_PF}\Microsoft Visual Studio\*\Community\Common7\IDE\devenv.exe",
        rf"{_PF}\Microsoft Visual Studio\*\Professional\Common7\IDE\devenv.exe",
        rf"{_PF}\Microsoft Visual Studio\*\Enterprise\Common7\IDE\devenv.exe",
    ],
    "pycharm": [
        rf"{_PF}\JetBrains\PyCharm Community Edition*\bin\pycharm64.exe",
        rf"{_PF}\JetBrains\PyCharm Professional Edition*\bin\pycharm64.exe",
        rf"{_LOCAL}\Programs\PyCharm Community Edition*\bin\pycharm64.exe",
    ],
    "git bash": [
        rf"{_PF}\Git\bin\bash.exe",
        rf"{_PF86}\Git\bin\bash.exe",
    ],
    "teams": [
        rf"{_LOCAL}\Microsoft\Teams\current\Teams.exe",
        rf"{_PF}\Microsoft\Teams\current\Teams.exe",
    ],
    "zoom":     [rf"{_LOCAL}\Zoom\bin\Zoom.exe"],
    "discord":  [rf"{_LOCAL}\Discord\Update.exe"],
    "slack":    [rf"{_LOCAL}\slack\slack.exe"],
    "whatsapp": [rf"{_LOCAL}\WhatsApp\WhatsApp.exe"],
    "spotify": [
        rf"{_LOCAL}\Microsoft\WindowsApps\Spotify.exe",
        rf"{_LOCAL}\Spotify\Spotify.exe",
    ],
    "vlc": [
        rf"{_PF}\VideoLAN\VLC\vlc.exe",
        rf"{_PF86}\VideoLAN\VLC\vlc.exe",
    ],
    "7zip": [
        rf"{_PF}\7-Zip\7zFM.exe",
        rf"{_PF86}\7-Zip\7zFM.exe",
    ],
}

APP_SYNONYMS: Dict[str, str] = {
    "microsoft excel":       "excel",
    "ms excel":              "excel",
    "excel spreadsheet":     "excel",
    "spreadsheet":           "excel",
    "microsoft word":        "word",
    "ms word":               "word",
    "word document":         "word",
    "microsoft powerpoint":  "powerpoint",
    "ms powerpoint":         "powerpoint",
    "slides":                "powerpoint",
    "presentation":          "powerpoint",
    "microsoft outlook":     "outlook",
    "mail":                  "outlook",
    "email client":          "outlook",
    "terminal":              "cmd",
    "shell":                 "powershell",
    "browser":               "chrome",
    "internet":              "chrome",
    "google chrome":         "chrome",
    "microsoft edge":        "edge",
    "mozilla firefox":       "firefox",
    "text editor":           "notepad",
    "files":                 "file explorer",
    "my computer":           "file explorer",
    "this pc":               "file explorer",
    "vs code":               "vscode",
    "code":                  "vscode",
    "vsc":                   "vscode",
    "vs":                    "visual studio",
    "devenv":                "visual studio",
    "task manager":          "task manager",
    "taskmgr":               "task manager",
    "calculator app":        "calculator",
    "calc":                  "calculator",
    "microsoft teams":       "teams",
}

_STT_FILLER = [
    "please", "the", "an", "a", "app", "application",
    "program", "software", "tool", "window", "open",
    "launch", "start", "run", "browser", "now",
]

WEBSITE_MAP: Dict[str, str] = {
    "youtube":       "https://www.youtube.com",
    "youtube.com":   "https://www.youtube.com",
    "gmail":         "https://mail.google.com",
    "google":        "https://www.google.com",
    "google drive":  "https://drive.google.com",
    "google docs":   "https://docs.google.com",
    "google sheets": "https://sheets.google.com",
    "google maps":   "https://maps.google.com",
    "facebook":      "https://www.facebook.com",
    "instagram":     "https://www.instagram.com",
    "twitter":       "https://www.twitter.com",
    "x":             "https://www.x.com",
    "reddit":        "https://www.reddit.com",
    "github":        "https://www.github.com",
    "linkedin":      "https://www.linkedin.com",
    "netflix":       "https://www.netflix.com",
    "amazon":        "https://www.amazon.com",
    "wikipedia":     "https://www.wikipedia.org",
    "chatgpt":       "https://chat.openai.com",
    "openai":        "https://www.openai.com",
    "notion":        "https://www.notion.so",
    "trello":        "https://www.trello.com",
    "figma":         "https://www.figma.com",
}

BROWSER_PATHS: Dict[str, List[str]] = {
    "chrome":  APP_LOOKUP["chrome"],
    "edge":    APP_LOOKUP["edge"],
    "firefox": APP_LOOKUP["firefox"],
}
BROWSER_FALLBACK_ORDER = ["chrome", "edge", "firefox"]


# ─────────────────────────────────────────────────────────────────
# USER APP MEMORY
# ─────────────────────────────────────────────────────────────────
def _load_user_apps() -> dict:
    if os.path.exists(_USER_APPS_FILE):
        try:
            with open(_USER_APPS_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_user_app(name: str, path: str) -> None:
    apps = _load_user_apps()
    apps[name.lower()] = path
    try:
        with open(_USER_APPS_FILE, "w") as f:
            json.dump(apps, f, indent=2)
        log.info("Saved user app: '%s' → %s", name, path)
    except Exception as e:
        log.warning("Could not save user app: %s", e)


# ─────────────────────────────────────────────────────────────────
# NAME CLEANING
# ─────────────────────────────────────────────────────────────────
def _clean_spoken_name(raw: str) -> str:
    text = raw.lower().strip()
    for word in _STT_FILLER:
        text = re.sub(rf"\b{re.escape(word)}\b", "", text)
    return re.sub(r"\s+", " ", text).strip()


# ─────────────────────────────────────────────────────────────────
# APP RESOLUTION
# ─────────────────────────────────────────────────────────────────
def _resolve_glob(pattern: str) -> Optional[str]:
    matches = sorted(glob.glob(pattern), reverse=True)
    return matches[0] if matches else None


def _resolve_path(path: str) -> Optional[str]:
    """Resolve one candidate path entry to a real path or None."""
    if path.startswith("ms-"):
        return path
    if "*" in path:
        return _resolve_glob(path)
    # Only use shutil.which for bare executables (no path separator)
    if os.sep not in path and "/" not in path:
        result = shutil.which(path)
        if result:
            return result
        # Also try without .exe extension
        no_ext = path.replace(".exe", "").replace(".EXE", "")
        return shutil.which(no_ext)
    return path if os.path.exists(path) else None


def _fuzzy_lookup_key(name: str) -> Optional[str]:
    all_keys = list(APP_LOOKUP.keys()) + list(APP_SYNONYMS.keys())
    result = fuzz_process.extractOne(name, all_keys, scorer=fuzz.token_sort_ratio)
    if result is None:
        return None
    match, score = result[0], result[1]
    log.debug("Fuzzy key match: '%s' → '%s' (score=%d)", name, match, score)
    if score >= 72:
        return APP_SYNONYMS.get(match, match)
    return None


def _scan_start_menu(app_name: str) -> Optional[str]:
    roots = [
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
        os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs"),
    ]
    name_lower = app_name.lower()
    for root in roots:
        if not os.path.isdir(root):
            continue
        for lnk in Path(root).rglob("*.lnk"):
            if name_lower in lnk.stem.lower():
                return str(lnk)
    return None


def resolve_app_name(spoken_name: str) -> Optional[str]:
    cleaned = _clean_spoken_name(spoken_name)
    log.info("Resolving app: '%s' → cleaned: '%s'", spoken_name, cleaned)

    user_apps = _load_user_apps()
    if cleaned in user_apps:
        path = user_apps[cleaned]
        if os.path.exists(path):
            log.info("Resolved via user_apps.json → %s", path)
            return path

    canonical = APP_SYNONYMS.get(cleaned, cleaned)

    for candidate in APP_LOOKUP.get(canonical, []):
        resolved = _resolve_path(candidate)
        if resolved:
            log.info("Resolved via lookup table → %s", resolved)
            return resolved

    fuzzy_key = _fuzzy_lookup_key(cleaned)
    if fuzzy_key and fuzzy_key != canonical:
        for candidate in APP_LOOKUP.get(fuzzy_key, []):
            resolved = _resolve_path(candidate)
            if resolved:
                log.info("Resolved via fuzzy match (key='%s') → %s", fuzzy_key, resolved)
                return resolved

    via_which = shutil.which(cleaned) or shutil.which(canonical)
    if via_which:
        log.info("Resolved via shutil.which → %s", via_which)
        return via_which

    result = _scan_start_menu(cleaned) or _scan_start_menu(canonical)
    if result:
        log.info("Resolved via Start Menu scan → %s", result)
        return result

    log.warning("Could not resolve app: '%s'", spoken_name)
    return None


# ─────────────────────────────────────────────────────────────────
# APP LAUNCHER
# ─────────────────────────────────────────────────────────────────
def _needs_elevation(path: str) -> bool:
    return os.path.basename(path).lower() in ELEVATED_APPS


def launch_app(path: str) -> Result:
    try:
        if ":" in path and os.sep not in path and "/" not in path:
            os.startfile(path)
            return Result(True, "Done.")

        path = os.path.normpath(path)

        if _needs_elevation(path):
            import ctypes
            log.info("Launching elevated: %s", path)
            ctypes.windll.shell32.ShellExecuteW(None, "runas", path, None, None, 1)
            return Result(True, "Done.")

        if path.lower().endswith(".lnk"):
            os.startfile(path)
            return Result(True, "Done.")

        subprocess.Popen([path], shell=False)
        return Result(True, "Done.")

    except FileNotFoundError:
        return Result(False, "I couldn't find the application file.")
    except PermissionError:
        return Result(False, "Permission denied. You may need to run ProVA as administrator.")
    except Exception as e:
        log.exception("launch_app failed for '%s': %s", path, e)
        return Result(False, "Something went wrong launching the app.")


# ─────────────────────────────────────────────────────────────────
# BROWSER & WEB SEARCH
# ─────────────────────────────────────────────────────────────────
def _find_browser(preferred: Optional[str] = None) -> Optional[Tuple[str, str]]:
    order = (
        [preferred] + [b for b in BROWSER_FALLBACK_ORDER if b != preferred]
        if preferred else BROWSER_FALLBACK_ORDER
    )
    for browser in order:
        for path in BROWSER_PATHS.get(browser, []):
            if os.path.exists(path):
                return browser, path
    for browser in order:
        found = shutil.which(browser)
        if found:
            return browser, found
    return None


def web_search(query: str, preferred_browser: Optional[str] = None) -> Result:
    if not query or not query.strip():
        return Result(False, "What would you like me to search for?")

    url = f"https://www.google.com/search?q={quote_plus(query.strip())}"
    browser = _find_browser(preferred_browser)

    if browser:
        name, exe = browser
        try:
            subprocess.Popen([exe, url])
            return Result(True, f"Searching for {query}.")
        except Exception as e:
            log.warning("Browser launch failed for %s: %s", name, e)

    try:
        webbrowser.open(url)
        return Result(True, f"Searching for {query}.")
    except Exception:
        return Result(False, "Could not open a browser to search.")


def open_url(url: str, preferred_browser: Optional[str] = None) -> Result:
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    browser = _find_browser(preferred_browser)
    if browser:
        name, exe = browser
        try:
            subprocess.Popen([exe, url])
            return Result(True, f"Opening {url}.")
        except Exception as e:
            log.warning("open_url failed for %s: %s", name, e)
    try:
        webbrowser.open(url)
        return Result(True, f"Opening {url}.")
    except Exception:
        return Result(False, "Could not open the page.")


# ─────────────────────────────────────────────────────────────────
# MAIN HANDLER
# ─────────────────────────────────────────────────────────────────
def handle(cmd, speak_fn) -> None:
    action = cmd.action
    target = (cmd.target or "").strip()

    if action == "open":
        if not target:
            speak_fn("What application would you like me to open?")
            return

        path = resolve_app_name(target)

        if path is None:
            cleaned_target = _clean_spoken_name(target)
            website_url = WEBSITE_MAP.get(cleaned_target) or WEBSITE_MAP.get(target.lower())
            if website_url:
                speak_fn(f"Opening {target} in your browser.")
                result = open_url(website_url)
                if not result.success:
                    speak_fn(result.message)
                return
            speak_fn(
                f"I couldn't find {target} on your computer. "
                "You can add it by telling me its full path, "
                "or try searching for it instead."
            )
            log.warning("Unresolved app '%s' — not in app table or website map", target)
            return

        speak_fn(f"Opening {target}.")
        result = launch_app(path)
        if not result.success:
            speak_fn(result.message)

    elif action == "search":
        if not target:
            speak_fn("What would you like me to search for?")
            return
        result = web_search(target)
        speak_fn(result.message)

    elif action == "url":
        if not target:
            speak_fn("Please give me a web address to open.")
            return
        result = open_url(target)
        speak_fn(result.message)

    else:
        speak_fn("I'm not sure what to do with that command.")
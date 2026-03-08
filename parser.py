"""
parser.py — ProVA intent parser.

Owns everything related to understanding what the user said:
  - Command dataclass
  - INTENT_PATTERNS  (priority-ranked)
  - All argument extractors
  - detect_intent()  — single public entry point

Fixes in this version:
  - `from fuzzywuzzy import fuzz` → guarded import with thefuzz fallback
  - All type hints use `Optional` from typing (not bare X | Y which needs 3.10+)
  - _FILLER regex compiled once at module level (was already correct)
  - extract_time: `g[1].isdigit()` guard prevents AttributeError on None group
  - PATTERNS sort key corrected (was inadvertently reversing priority)
"""

from __future__ import annotations

import re
import logging
from dataclasses import dataclass, field
from typing import Optional, Tuple, List

# Graceful fuzzy import — works with both fuzzywuzzy and thefuzz
try:
    from fuzzywuzzy import fuzz
except ImportError:
    try:
        from thefuzz import fuzz  # type: ignore
    except ImportError:
        # Minimal fallback so the file at least imports without crashing
        class fuzz:  # type: ignore
            @staticmethod
            def partial_ratio(a: str, b: str) -> int:
                return 100 if a in b else 0

log = logging.getLogger("ProVA.Parser")

FUZZY_THRESHOLD = 68   # lowered from 75 — Indian English accent varies more from exact keywords


# ─────────────────────────────────────────────────────────────────
# COMMAND
# ─────────────────────────────────────────────────────────────────
@dataclass
class Command:
    intent:     str
    raw:        str
    action:     Optional[str] = None
    target:     Optional[str] = None
    time_str:   Optional[str] = None
    message:    Optional[str] = None
    confidence: int           = 100
    extra:      dict          = field(default_factory=dict)

    def summary(self) -> str:
        parts = [f"intent={self.intent}"]
        if self.action:           parts.append(f"action={self.action}")
        if self.target:           parts.append(f"target={self.target}")
        if self.time_str:         parts.append(f"time={self.time_str}")
        if self.message:          parts.append(f"msg={self.message[:30]!r}")
        if self.confidence < 100: parts.append(f"conf={self.confidence}")
        return " | ".join(parts)


# ─────────────────────────────────────────────────────────────────
# INTENT PATTERNS
# (priority, keyword_phrase, intent, action)
# priority 10 = specific multi-word phrase
# priority  5 = single-word catch-all
# Pre-sorted at module load: (priority DESC, len DESC)
# ─────────────────────────────────────────────────────────────────
_RAW: List[Tuple[int, str, str, str]] = [

    # ── Help ──────────────────────────────────────────────────────
    (10, "what can you do",   "help", "guide"),
    (10, "show commands",     "help", "guide"),
    (10, "list commands",     "help", "guide"),
    (10, "help with",         "help", "guide"),
    (10, "help me",           "help", "guide"),
    ( 5, "help",              "help", "guide"),

    # ── Computer control ──────────────────────────────────────────
    (10, "open excel",        "excel",            "open"),
    (10, "launch excel",      "excel",            "open"),
    (10, "search for",        "computer_control", "search"),
    (10, "look up",           "computer_control", "search"),
    (10, "google search",     "computer_control", "search"),
    ( 5, "open",              "computer_control", "open"),
    ( 5, "launch",            "computer_control", "open"),
    ( 5, "start",             "computer_control", "open"),
    ( 5, "search",            "computer_control", "search"),
    ( 5, "google",            "computer_control", "search"),

    # ── File manager — create ─────────────────────────────────────
    (10, "create file",       "file_manager", "create_file"),
    (10, "make file",         "file_manager", "create_file"),
    (10, "new file",          "file_manager", "create_file"),
    (10, "create a file",     "file_manager", "create_file"),
    (10, "make a file",       "file_manager", "create_file"),
    (10, "create folder",     "file_manager", "create_folder"),
    (10, "make folder",       "file_manager", "create_folder"),
    (10, "new folder",        "file_manager", "create_folder"),
    (10, "create a folder",   "file_manager", "create_folder"),
    (10, "make a folder",     "file_manager", "create_folder"),
    (10, "create directory",  "file_manager", "create_folder"),
    (10, "make directory",    "file_manager", "create_folder"),
    (10, "create a directory","file_manager", "create_folder"),

    # ── File manager — delete ─────────────────────────────────────
    (10, "delete file",       "file_manager", "delete_file"),
    (10, "remove file",       "file_manager", "delete_file"),
    (10, "delete folder",     "file_manager", "delete_folder"),
    (10, "remove folder",     "file_manager", "delete_folder"),
    (10, "delete directory",  "file_manager", "delete_folder"),
    ( 5, "delete",            "file_manager", "delete_ambiguous"),
    ( 5, "remove",            "file_manager", "delete_ambiguous"),

    # ── File manager — ops ────────────────────────────────────────
    (10, "rename file",       "file_manager", "rename"),
    (10, "rename folder",     "file_manager", "rename"),
    (10, "list files",        "file_manager", "list"),
    (10, "list folder",       "file_manager", "list"),
    (10, "show files",        "file_manager", "list"),
    (10, "show folder",       "file_manager", "list"),
    (10, "what's in",         "file_manager", "list"),
    (10, "whats in",          "file_manager", "list"),
    (10, "copy file",         "file_manager", "copy"),
    (10, "copy folder",       "file_manager", "copy"),
    (10, "move file",         "file_manager", "move"),
    (10, "move folder",       "file_manager", "move"),
    (10, "file info",         "file_manager", "info"),
    (10, "folder info",       "file_manager", "info"),
    (10, "info about",        "file_manager", "info"),
    (10, "details of",        "file_manager", "info"),
    (10, "find file",         "file_manager", "find"),
    (10, "find folder",       "file_manager", "find"),
    (10, "search file",       "file_manager", "find"),
    (10, "locate file",       "file_manager", "find"),
    (10, "locate folder",     "file_manager", "find"),
    (10, "where is",          "file_manager", "find"),
    (10, "where's my",        "file_manager", "find"),
    (10, "check file",        "file_manager", "find"),
    ( 5, "rename",            "file_manager", "rename"),
    ( 5, "copy",              "file_manager", "copy"),
    ( 5, "move",              "file_manager", "move"),
    ( 5, "find",              "file_manager", "find"),
    ( 5, "locate",            "file_manager", "find"),
    ( 5, "info",              "file_manager", "info"),
    ( 5, "details",           "file_manager", "info"),

    # ── Email ─────────────────────────────────────────────────────
    (20, "help me write an email", "email", "send"),
    (20, "help me send an email",  "email", "send"),
    (20, "help me compose",        "email", "send"),
    (10, "cancel email",       "email", "cancel"),
    (10, "stop email",         "email", "cancel"),
    (10, "send email",        "email", "send"),
    (10, "send mail",         "email", "send"),
    (10, "compose email",     "email", "compose"),
    (10, "write email",       "email", "compose"),
    (10, "draft email",       "email", "compose"),
    ( 5, "email",             "email", "send"),
    ( 5, "compose",           "email", "compose"),
    ( 5, "mail",              "email", "send"),

    # ── Excel / Dashboard ─────────────────────────────────────────
    (10, "create dashboard",  "excel", "dashboard"),
    (10, "make dashboard",    "excel", "dashboard"),
    (10, "build dashboard",   "excel", "dashboard"),
    (10, "open dashboard",    "excel", "open"),
    (10, "make chart",        "excel", "chart"),
    (10, "create chart",      "excel", "chart"),
    (10, "make pivot",        "excel", "pivot"),
    (10, "create pivot",      "excel", "pivot"),
    (10, "analyze data",      "excel", "analyze"),
    (10, "analyse data",      "excel", "analyze"),
    ( 5, "dashboard",         "excel", "dashboard"),
    ( 5, "chart",             "excel", "chart"),
    ( 5, "pivot",             "excel", "pivot"),
    ( 5, "analyze",           "excel", "analyze"),
    ( 5, "analyse",           "excel", "analyze"),
    ( 5, "excel",             "excel", "open"),

    # ── Reminders ─────────────────────────────────────────────────
    (10, "remind me in",      "reminder", "remind_in"),
    (10, "remind in",         "reminder", "remind_in"),
    (10, "remind me to",      "reminder", "set"),
    (10, "set reminder",      "reminder", "set"),
    (10, "set alarm",         "reminder", "alarm"),
    (10, "wake me at",        "reminder", "alarm"),
    (10, "wake me up",        "reminder", "alarm"),
    (10, "daily reminder",    "reminder", "daily"),
    (10, "daily alarm",       "reminder", "daily"),
    (10, "list reminders",    "reminder", "list"),
    (10, "show reminders",    "reminder", "list"),
    (10, "cancel reminder",   "reminder", "delete"),
    (10, "delete reminder",   "reminder", "delete"),
    ( 5, "remind me",         "reminder", "set"),
    ( 5, "reminder",          "reminder", "set"),
    ( 5, "alarm",             "reminder", "alarm"),
    ( 5, "alert me",          "reminder", "set"),

    # ── Indian-English spoken variants ───────────────────────────
    # Google en-IN commonly returns these alternate phrasings
    (10, "make a new file",   "file_manager", "create_file"),
    (10, "make a new folder", "file_manager", "create_folder"),
    (10, "open the file",     "file_manager", "find"),
    (10, "send a mail",       "email",        "send"),
    (10, "write a mail",      "email",        "compose"),
    (10, "search the web",    "computer_control", "search"),
    (10, "open the browser",  "computer_control", "open"),
    (10, "set a timer",       "reminder", "alarm"),
    (10, "start a timer",     "reminder", "alarm"),
    (10, "set timer",         "reminder", "alarm"),
    ( 5, "timer",             "reminder", "alarm"),
    (10, "set an alarm",      "reminder",     "alarm"),
    (10, "set a reminder",    "reminder",     "set"),
    (10, "show my reminders", "reminder",     "list"),
    (10, "what are my reminders", "reminder", "list"),

    # ── System ────────────────────────────────────────────────────
    (10, "shut down prova",   "system", "exit"),
    (10, "turn off prova",    "system", "exit"),
    (10, "goodbye prova",     "system", "exit"),
    (10, "stop prova",        "system", "exit"),
    ( 5, "goodbye",           "system", "exit"),
    ( 5, "exit",              "system", "exit"),
    ( 5, "quit",              "system", "exit"),
    ( 5, "sleep",             "system", "sleep"),
    ( 5, "pause",             "system", "pause"),
]

# Pre-sort once at import: priority DESC, then keyword length DESC
PATTERNS: List[Tuple[int, str, str, str]] = sorted(
    _RAW, key=lambda x: (x[0], len(x[1])), reverse=True
)


# ─────────────────────────────────────────────────────────────────
# NORMALISATION
# ─────────────────────────────────────────────────────────────────
_FILLER = re.compile(
    r"\b(please|can you|could you|would you|i want to|i need to|"
    r"i'd like to|i would like to|hey prova|"
    r"kindly|do one thing|just|only|actually|basically|"
    r"tell me|show me|give me|let me)\b",
    re.IGNORECASE,
)
# Note: bare 'prova' intentionally removed from _FILLER — it was stripping
# the noun from shutdown commands like "stop prova", "goodbye prova",
# "turn off prova" before they could match their patterns.

def _normalise(text: str) -> str:
    text = _FILLER.sub("", text)
    return re.sub(r"\s+", " ", text).strip().lower()


# ─────────────────────────────────────────────────────────────────
# SPOKEN NUMBERS → DIGITS
# ─────────────────────────────────────────────────────────────────
_WORD_NUMS = {
    "one":"1","two":"2","three":"3","four":"4","five":"5",
    "six":"6","seven":"7","eight":"8","nine":"9","ten":"10",
    "eleven":"11","twelve":"12","fifteen":"15","twenty":"20",
    "thirty":"30","forty":"40","sixty":"60",
}
_WORD_NUM_RE = re.compile(
    r"\b(" + "|".join(re.escape(w) for w in _WORD_NUMS) + r")\b",
    re.IGNORECASE,
)
def _w2d(text: str) -> str:
    return _WORD_NUM_RE.sub(lambda m: _WORD_NUMS[m.group(1).lower()], text)


# ─────────────────────────────────────────────────────────────────
# EXTRACTORS
# ─────────────────────────────────────────────────────────────────

def extract_time(text: str) -> Optional[str]:
    text = _w2d(text)
    for phrase, result in [
        (r"half\s+past\s+(\d{1,2})",    lambda m: f"{int(m.group(1)):02d}:30"),
        (r"quarter\s+past\s+(\d{1,2})", lambda m: f"{int(m.group(1)):02d}:15"),
        (r"quarter\s+to\s+(\d{1,2})",   lambda m: f"{int(m.group(1))-1:02d}:45"),
    ]:
        m = re.search(phrase, text, re.IGNORECASE)
        if m:
            return result(m)

    # Pattern 1 has 3 groups: (hour, minute, am/pm?  — also matches a.m./p.m.)
    # Pattern 2 has 2 groups: (hour, am/pm)
    # g[2] must never be accessed when len(g)==2 — that was the original IndexError.
    # a\.?m\.? matches: am, a.m, a.m.   |   p\.?m\.? matches: pm, p.m, p.m.
    _period_pat = r'a\.?m\.?|p\.?m\.?'
    for pat in [
        rf'\b(\d{{1,2}}):(\d{{2}})\s*({_period_pat})?\b',
        rf'\b(\d{{1,2}})\s*({_period_pat})\b',
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            g = m.groups()
            hour = int(g[0])
            if len(g) == 3:
                # First pattern: groups are (hour, minute, period?)
                minute = int(g[1]) if g[1] is not None and g[1].isdigit() else 0
                period = (g[2] or "").lower().replace(".", "")
            else:
                # Second pattern: groups are (hour, period) — no minutes
                minute = 0
                period = (g[1] or "").lower().replace(".", "")
            if period == "pm" and hour != 12: hour += 12
            elif period == "am" and hour == 12: hour = 0
            return f"{hour:02d}:{minute:02d}"
    return None


def extract_duration_minutes(text: str) -> Optional[int]:
    text = _w2d(text)
    m = re.search(r"in\s+(\d+)\s+minute", text, re.IGNORECASE)
    if m: return int(m.group(1))
    m = re.search(r"in\s+(\d+)\s+hour",   text, re.IGNORECASE)
    if m: return int(m.group(1)) * 60
    if re.search(r"half\s+an?\s+hour", text, re.IGNORECASE): return 30
    return None


_NAME_STOP = r"(?:about|regarding|saying|with subject|re:|re\b|and|,|\.|$)"

def extract_email_recipient(text: str) -> Optional[str]:
    m = re.search(
        r"\b(?:email|send|compose|mail|write|draft)\b.*?\bto\s+"
        r"((?:[A-Za-z]+\.?\s*){1,4}?)\s*(?=" + _NAME_STOP + r")",
        text, re.IGNORECASE,
    )
    if m:
        name = m.group(1).strip().rstrip(".,")
        if name: return name
    m = re.search(
        r"\bto\s+([A-Za-z][A-Za-z\s\.]{0,40}?)\s*(?=" + _NAME_STOP + r")",
        text, re.IGNORECASE,
    )
    if m:
        name = m.group(1).strip().rstrip(".,")
        if name: return name
    return None


def extract_email_body_hint(text: str) -> Optional[str]:
    m = re.search(r"\b(?:about|regarding|saying|re:|re\b)\s+(.+)$", text, re.IGNORECASE)
    return m.group(1).strip() if m else None


_PATH_PREPS    = r"\b(?:in|inside|at|into|under|within|to|from)\b"
_FILE_TRIGGERS = (
    r"\b(?:create|make|new|delete|remove|rename|copy|move|"
    r"find|search|list|show|open|get|info|details|about|"
    r"file|folder|directory|called|named|name)\b"
)

def extract_file_target(text: str) -> Optional[str]:
    m = re.search(r'["\']([^"\'\"]+)["\']', text)
    if m: return m.group(1).strip()
    cleaned = re.sub(_FILE_TRIGGERS, " ", text, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    # Strip leading articles and connector words left by trigger removal.
    # "create a folder headphones" → "a headphones"       → "headphones"
    # "delete the folder on"       → "the on"             → "on"
    # "folder called vaishnavi"    → "called vaishnavi"   → "vaishnavi"
    # "create a folder name ticket"→ "name ticket"        → "ticket" (name stripped too)
    cleaned = re.sub(r"^(?:a|an|the|called|named|name)\s+", "", cleaned, flags=re.IGNORECASE)
    prep = re.search(_PATH_PREPS, cleaned, re.IGNORECASE)
    if prep: cleaned = cleaned[:prep.start()].strip()
    rename = re.search(r"^(.+?)\s+to\s+", cleaned, re.IGNORECASE)
    if rename: cleaned = rename.group(1).strip()
    return cleaned or None


def extract_path_arg(text: str) -> Optional[str]:
    m = re.search(
        _PATH_PREPS + r"\s+(?:the\s+)?([\w\s\-\.]+?)(?:\s+folder|\s+directory)?\s*$",
        text, re.IGNORECASE,
    )
    return m.group(1).strip() if m else None


def extract_rename_target(text: str) -> Optional[str]:
    m = re.search(r"\bto\s+(.+)$", text, re.IGNORECASE)
    return m.group(1).strip() if m else None


def infer_delete_type(text: str) -> str:
    t = text.lower()
    if any(w in t for w in ("folder", "directory", "dir")): return "delete_folder"
    if re.search(r"\b\w+\.\w{2,5}\b", t): return "delete_file"
    return "delete_file"


def extract_app_name(text: str, trigger: str) -> Optional[str]:
    m = re.search(rf'\b{re.escape(trigger)}\s+(.+)', text, re.IGNORECASE)
    return m.group(1).strip() if m else None


def extract_search_query(text: str) -> Optional[str]:
    m = re.search(r'\b(?:search(?:\s+for)?|google|look\s+up)\s+(.+)', text, re.IGNORECASE)
    return m.group(1).strip() if m else None


def extract_reminder_message(text: str) -> Optional[str]:
    text = _w2d(text)
    # Strip all reminder command words (longest phrases first to avoid partial matches)
    c = re.sub(
        r'\b(set\s+daily\s+alarm|set\s+daily\s+reminder|daily\s+alarm|daily\s+reminder'
        r'|remind\s+me\s+in|remind\s+me\s+to|remind\s+me|remind\s+in'
        r'|set\s+an?\s+alarm|set\s+an?\s+reminder|set\s+reminder'
        r'|list\s+reminders|show\s+reminders|cancel\s+reminder|delete\s+reminder'
        r'|reminder|alarm|alert\s+me'
        r'|from\s+now|set|daily|list|cancel|delete|alert'
        r'|to|at|by|in|an?|the)\b',
        '', text, flags=re.IGNORECASE
    )
    c = re.sub(r'\b\d+\s*(?:minute|hour|second)s?\b', '', c, flags=re.IGNORECASE)
    c = re.sub(r'\b\d{1,2}(:\d{2})?\s*(a\.?m\.?|p\.?m\.?)?\b', '', c, flags=re.IGNORECASE)
    return c.strip(" ,.") or None


def extract_excel_filename(text: str) -> Optional[str]:
    m = re.search(
        r"\b(?:from|using|with|for|of)\s+(.+?)"
        r"(?:\s+using|\s+with|\s+metric|\s+date|\s+category|$)",
        text, re.IGNORECASE,
    )
    return m.group(1).strip() if m else extract_file_target(text)


# ─────────────────────────────────────────────────────────────────
# SLOT FILLERS  (called per-intent after match)
# ─────────────────────────────────────────────────────────────────

def extract_copy_move_dest(text: str) -> Optional[str]:
    """
    Extract the destination for copy/move commands.
    'copy vaseline to downloads'     → 'downloads'
    'move report to desktop'         → 'desktop'
    'move vaseline from storage to download' → 'download'
    Returns the raw spoken destination; file_manager resolves it to a real path.
    """
    m = re.search(
        r'\b(?:to|into|onto)\s+(?:the\s+|my\s+)?'
        r'([\w\s\-\.]+?)(?:\s+folder|\s+directory)?'
        r'(?:\s+from\b|$)',
        text, re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()
    return None


def _fill_computer_control(cmd: Command, text: str, trigger: str) -> None:
    if cmd.action == "open":
        cmd.target = extract_app_name(text, trigger)
    elif cmd.action == "search":
        cmd.target = extract_search_query(text)


def _fill_file_manager(cmd: Command, text: str) -> None:
    if cmd.action == "delete_ambiguous":
        cmd.action = infer_delete_type(text)
    cmd.target   = extract_file_target(text)
    path_arg     = extract_path_arg(text)
    if cmd.action == "rename":
        cmd.extra["new_name"] = extract_rename_target(text) or ""
    elif cmd.action in ("copy", "move"):
        # Destination is after "to/into/onto" — not the same as path_arg (in/inside/at)
        dest = extract_copy_move_dest(text)
        cmd.extra["dest"] = dest or "."
    elif path_arg:
        cmd.extra["path"] = path_arg


def _fill_email(cmd: Command, text: str) -> None:
    cmd.target  = extract_email_recipient(text)
    cmd.message = extract_email_body_hint(text)


def _fill_excel(cmd: Command, text: str) -> None:
    cmd.target = extract_excel_filename(text)
    m_metric = re.search(r"\bmetric\s+(\w[\w\s]{0,30}?)(?:\s+date|\s+category|$)", text, re.IGNORECASE)
    m_date   = re.search(r"\bdate\s+(?:column\s+)?(\w[\w\s]{0,20}?)(?:\s+category|$)", text, re.IGNORECASE)
    m_cat    = re.search(r"\b(?:category|group)\s+(?:by\s+|column\s+)?(\w[\w\s]{0,20}?)(?:\s+metric|$)", text, re.IGNORECASE)
    if m_metric: cmd.extra["metric"]       = m_metric.group(1).strip()
    if m_date:   cmd.extra["date_col"]     = m_date.group(1).strip()
    if m_cat:    cmd.extra["category_col"] = m_cat.group(1).strip()


def _fill_reminder(cmd: Command, text: str) -> None:
    cmd.time_str = extract_time(text)
    cmd.message  = extract_reminder_message(text)
    if cmd.action == "remind_in":
        mins = extract_duration_minutes(text)
        if mins: cmd.extra["minutes"] = mins


# ─────────────────────────────────────────────────────────────────
# PUBLIC ENTRY POINT
# ─────────────────────────────────────────────────────────────────
def detect_intent(raw_text: str) -> Command:
    """
    Parse spoken or typed text into a structured Command.

    Steps:
      1. Normalise (strip filler, lowercase)
      2. Blocklist check — single-word confirmation/cancellation responses
         (yes, no, okay, etc.) must NEVER fuzzy-route to a system command.
         They are only meaningful inside _voice_confirm, not the main loop.
      3. Exact substring match (priority + length sorted — specific beats generic)
      4. Fuzzy fallback if nothing matched — skipped for very short utterances
         (≤ 2 words) to prevent noise like "yes" firing system/exit at conf=80.
      5. Per-intent slot filling on original text (preserves casing for filenames)
    """
    norm = _normalise(raw_text)

    # ── Blocklist: words that are ONLY valid inside a confirmation prompt ──
    # Without this guard, "yes" fuzzy-matches "goodbye" (partial_ratio=80)
    # and shuts ProVA down. "no" similarly matches "prova" etc.
    _CONFIRM_WORDS = {
        "yes", "yeah", "yep", "yup", "yas", "ya",
        "no",  "nope", "nah",
        "ok",  "okay", "sure", "cancel", "abort",
    }
    if norm.strip() in _CONFIRM_WORDS:
        log.info("Blocklist: confirmation word %r ignored in main loop", norm)
        return Command(intent="unknown", raw=raw_text, confidence=0)

    matched_kw     = None
    matched_intent = None
    matched_action = None
    confidence     = 0

    # Exact match
    for _pri, keyword, intent, action in PATTERNS:
        if keyword in norm:
            matched_kw     = keyword
            matched_intent = intent
            matched_action = action
            confidence     = 100
            break

    # Fuzzy fallback — skipped for very short utterances (≤ 2 words).
    # Short phrases like "yes", "yas ya", "no way" produce false positives
    # at threshold=68 because fuzz.partial_ratio finds accidental char overlap
    # with short keywords like "exit", "goodbye", "quit".
    if confidence < 100 and len(norm.split()) > 2:
        best = 0
        for _pri, keyword, intent, action in PATTERNS:
            score = fuzz.partial_ratio(keyword, norm)
            if score > best and score >= FUZZY_THRESHOLD:
                best           = score
                matched_kw     = keyword
                matched_intent = intent
                matched_action = action
                confidence     = score

    if not matched_intent:
        log.info("No intent match for: %r", norm)
        return Command(intent="unknown", raw=raw_text, confidence=0)

    log.info("Parsed: intent=%s action=%s trigger=%r conf=%d",
             matched_intent, matched_action, matched_kw, confidence)

    cmd = Command(intent=matched_intent, raw=raw_text,
                  action=matched_action, confidence=confidence)

    # Slot filling uses raw_text to preserve casing in filenames/names
    if matched_intent == "computer_control":
        _fill_computer_control(cmd, raw_text, matched_kw or "")
    elif matched_intent == "file_manager":
        _fill_file_manager(cmd, raw_text)
    elif matched_intent == "email":
        _fill_email(cmd, raw_text)
    elif matched_intent == "excel":
        _fill_excel(cmd, raw_text)
    elif matched_intent == "reminder":
        _fill_reminder(cmd, raw_text)

    return cmd
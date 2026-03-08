"""
ProVA — modules/file_manager.py
=================================
Voice-friendly file manager.

Changes in this version:
  FIX 1 — BASE_DIR is now project_root/storage (created on first use).
           Falls back to ~/ProVA/storage if __file__ is unavailable.

  FIX 2 — extract_file_target() rewritten in parser_helpers section.
           Old regex left "a xyz" after stripping trigger words because
           articles (a, an, the, my) and filler ("called", "named") were
           not stripped. Now strips them and handles quoted names too.

  FIX 3 — ALLOWED_ROOTS expanded with all common Windows user folders
           (Desktop, Downloads, Documents, Pictures, Music, Videos) and
           a KNOWN_LOCATIONS dict that maps spoken words → real Paths.
           "copy notes to desktop" now resolves to the real Desktop path.

  FIX 4 — copy_item / move_item: _is_allowed() now checks DEST separately
           and permits known locations even when src is in BASE_DIR.

  FIX 5 — find_files() now speaks the full path of each match, not just
           the name. "Where is my file budget?" gives a complete path.

  FIX 6 — Smarter rename parsing: "rename old name to new name" correctly
           splits on " to " without consuming "to" in middle-of-word positions.

  FIX 7 — handle() validates destination for copy/move against KNOWN_LOCATIONS
           so "copy xyz to desktop" works end-to-end.
"""

import os
import re
import shutil
import logging
import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional, List

log = logging.getLogger("ProVA.FileManager")


# ─────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────

def _resolve_base_dir() -> Path:
    """
    Resolve BASE_DIR to <project_root>/storage.
    Uses __file__ so it works regardless of cwd.
    Falls back to ~/ProVA/storage if running outside the project.
    """
    try:
        project_root = Path(__file__).resolve().parent.parent
        d = project_root / "storage"
    except NameError:
        d = Path.home() / "ProVA" / "storage"
    d.mkdir(parents=True, exist_ok=True)
    return d


BASE_DIR = _resolve_base_dir()

# Human-readable location names → real Paths
# Used to resolve spoken destinations like "copy to desktop"
KNOWN_LOCATIONS: dict = {
    # Plural forms
    "desktop":          Path.home() / "Desktop",
    "downloads":        Path.home() / "Downloads",
    "documents":        Path.home() / "Documents",
    "pictures":         Path.home() / "Pictures",
    "photos":           Path.home() / "Pictures",
    "music":            Path.home() / "Music",
    "videos":           Path.home() / "Videos",
    # Singular aliases (STT often drops the plural S)
    "download":         Path.home() / "Downloads",
    "document":         Path.home() / "Documents",
    "picture":          Path.home() / "Pictures",
    "photo":            Path.home() / "Pictures",
    "video":            Path.home() / "Videos",
    # ProVA storage
    "storage":          BASE_DIR,
    "prova":            BASE_DIR,
    "prova storage":    BASE_DIR,
    "home":             Path.home(),
}

ALLOWED_ROOTS: List[Path] = [
    BASE_DIR,
    Path.home() / "Desktop",
    Path.home() / "Downloads",
    Path.home() / "Documents",
    Path.home() / "Pictures",
    Path.home() / "Music",
    Path.home() / "Videos",
    Path.home(),
]

LIST_CAP = 30

WINDOWS_RESERVED = {
    "con", "prn", "aux", "nul",
    "com1","com2","com3","com4","com5","com6","com7","com8","com9",
    "lpt1","lpt2","lpt3","lpt4","lpt5","lpt6","lpt7","lpt8","lpt9",
}

INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

_lock = threading.Lock()


# ─────────────────────────────────────────────────────────────────
# RESULT
# ─────────────────────────────────────────────────────────────────
@dataclass
class Result:
    success:  bool
    message:  str
    data:     dict = field(default_factory=dict)
    needs_confirmation: bool = False
    confirm_prompt: str = ""


# ─────────────────────────────────────────────────────────────────
# PARSER HELPERS — voice text → clean names / paths
# These live here (not in parser.py) so the file manager fully owns
# its own text-understanding without depending on parser internals.
# ─────────────────────────────────────────────────────────────────

# Words that introduce the item name but are not part of it
_NAME_ARTICLES  = re.compile(
    r"^\s*(?:a|an|the|my|our|this|that|some)\s+", re.IGNORECASE
)  # Note: "old" and "new" intentionally excluded — they appear in real file names
# Trigger verbs/nouns that should be stripped before extracting the name
_FM_TRIGGERS = re.compile(
    r"\b(?:create|make|delete|remove|rename|copy|move|"
    r"find|search|locate|where|check|is|"
    r"list|show|open|get|info|details|about|"
    r"file|files|folder|folders|directory|directories|"
    r"called|named|with\s+name)\b",
    re.IGNORECASE,
)  # Note: "new" excluded — it appears in rename targets ("rename X to new name")
# Prepositions that introduce destinations / paths
_PATH_PREPS = re.compile(
    r"\b(?:in|inside|at|into|under|within|onto|on)\b",
    re.IGNORECASE,
)  # Note: "to", "from", "for" excluded — they appear in "rename X to Y", "search for X"
# Filler words that sometimes precede a name
_FILLER = re.compile(
    r"\b(?:please|just|kindly|actually|basically|the|a|an|my)\b",
    re.IGNORECASE,
)


def _clean_name(raw: str) -> str:
    """
    Strip articles, trigger words and filler from a raw name fragment.
    "a xyz"       → "xyz"
    "my projects" → "projects"
    "the report"  → "report"
    "please notes" → "notes"
    """
    raw = re.sub(r"^(?:for|from)\s+", "", raw.strip(), flags=re.IGNORECASE)
    s = _NAME_ARTICLES.sub("", raw)
    s = _FILLER.sub(" ", s)
    # Strip leading "new" — it's filler from "make a new X" phrasing, not part of the name
    s = re.sub(r"^new\s+", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_file_target(text: str) -> Optional[str]:
    """
    Extract the filename/folder name from a spoken command.
    Handles:
      - Quoted names:  create folder "my project"   → my project
      - called/named:  make a folder called reports  → reports
      - bare:          create folder xyz             → xyz
      - multi-word:    create folder test project    → test project
    """
    # 0. Special case: "list files in X" -- the target is what follows "in/inside"
    # Expand common contractions that confuse trigger stripping
    text = re.sub(r"\bwhere's\b", "where is", text, flags=re.IGNORECASE)
    text = re.sub(r"\bwhat's\b",  "what is",  text, flags=re.IGNORECASE)
    text = re.sub(r"\bwhats\b",   "what is",  text, flags=re.IGNORECASE)
    text = re.sub(r"\b's\s+(?:my|the)\b", " ", text, flags=re.IGNORECASE)

    if re.match(r"^\s*(?:list|show)\b", text, re.IGNORECASE):
        m_in = re.search(
            r"\b(?:in|inside)\s+(?:the\s+)?(.+?)(?:\s+folder|\s+directory)?\s*$",
            text, re.IGNORECASE
        )
        if m_in:
            return _clean_name(m_in.group(1))

    # 1. Quoted name takes absolute priority
    m = re.search(r'["\u2018\u2019\u201c\u201d]([^"\']+)["\u2018\u2019\u201c\u201d]', text)
    if m:
        return m.group(1).strip()

    # 2. "called X" or "named X" pattern
    m = re.search(r'\b(?:called|named|with\s+name)\s+(.+?)(?:\s+in\s|\s+inside\s|\s+at\s|$)',
                  text, re.IGNORECASE)
    if m:
        return _clean_name(m.group(1))

    # 3. Strip trigger words → what remains is the name
    s = _FM_TRIGGERS.sub(" ", text)
    s = re.sub(r"\s+", " ", s).strip()

    # Strip destination clause (anything after a path preposition)
    prep = _PATH_PREPS.search(s)
    if prep:
        s = s[:prep.start()].strip()

    # Also strip trailing "from <location>" clause — STT often produces
    # "move vaseline from storage to download" and trigger-stripping leaves
    # "vaseline from storage" before "to download".
    s = re.sub(r'\s+from\s+\S+.*$', '', s, flags=re.IGNORECASE).strip()

    # Strip rename target ("old to new" → take old part)
    m_rename = re.match(r'^(.+?)\s+to\s+\S', s, re.IGNORECASE)
    if m_rename:
        s = m_rename.group(1).strip()

    return _clean_name(s) or None


def parse_rename_parts(text: str):
    """
    Extract (old_name, new_name) from a rename command.
    Handles:
      rename folder old name to new name
      rename old to new
      rename 'old file' to 'new file'
    Returns (old, new) or (raw_target, "") if only one name found.
    """
    # Quoted form
    m = re.search(
        r'["\']([^"\']+)["\']\s+to\s+["\']([^"\']+)["\']', text, re.IGNORECASE
    )
    if m:
        return m.group(1).strip(), m.group(2).strip()

    # Strip rename-specific boilerplate (not _FM_TRIGGERS — it would eat "new name")
    s = re.sub(
        r"\b(?:rename|move|my|the|a|an|this|that|folder|file|"
        r"directory|directories|item)\b",
        " ", text, flags=re.IGNORECASE
    )
    s = re.sub(r"\s+", " ", s).strip()

    # Use word-boundary " to " split (avoids splitting "Toronto" etc.)
    parts = re.split(r'\s+to\s+', s, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) == 2:
        old = _clean_name(parts[0])
        # new_name: strip articles but keep 'new' (it may be part of the name)
        new = re.sub(r'^\s*(?:a|an|the|my|our|this|that|some)\s+', '', parts[1].strip(), flags=re.IGNORECASE)
        new = re.sub(r'\s+(?:in|inside|at|into)\s+.*$', '', new, flags=re.IGNORECASE).strip()
        return old, new

    return _clean_name(s), ""


def parse_destination(text: str) -> Optional[str]:
    """
    Extract destination location from spoken command.
    "copy notes to desktop"   → "desktop"
    "move report to downloads" → "downloads"
    "copy project to D:/work"  → "D:/work"
    Returns the raw destination string; resolve_location() maps it to a Path.
    """
    m = re.search(
        r'\b(?:to|into|onto)\s+(?:the\s+|my\s+)?(.+?)(?:\s+folder|\s+directory)?\s*$',
        text, re.IGNORECASE
    )
    if m:
        return m.group(1).strip()
    return None


def resolve_location(dest_str: str) -> Optional[Path]:
    """
    Map a spoken destination string to a real filesystem Path.
    "desktop"   → C:/Users/you/Desktop
    "downloads" → C:/Users/you/Downloads
    Absolute paths pass through as-is.
    Relative paths are resolved against BASE_DIR.
    """
    if not dest_str:
        return None

    # Check known location names first (case-insensitive)
    key = dest_str.strip().lower()
    if key in KNOWN_LOCATIONS:
        return KNOWN_LOCATIONS[key]

    # Absolute path?
    p = Path(dest_str)
    if p.is_absolute():
        return p.resolve()

    # Relative → resolve against BASE_DIR
    return (BASE_DIR / dest_str).resolve()


# ─────────────────────────────────────────────────────────────────
# PATH SAFETY
# ─────────────────────────────────────────────────────────────────

def _ensure_base_dir() -> None:
    BASE_DIR.mkdir(parents=True, exist_ok=True)


def _is_allowed(path: Path) -> bool:
    resolved = path.resolve()
    return any(
        str(resolved).startswith(str(root.resolve()))
        for root in ALLOWED_ROOTS
    )


def _resolve(raw: str, base: Path = BASE_DIR) -> Path:
    p = Path(os.path.normpath(raw))
    if not p.is_absolute():
        p = base / p
    return p.resolve()


def _validate_name(name: str) -> Optional[str]:
    if not name or not name.strip():
        return "Please provide a file or folder name."
    name = name.strip()
    stem = Path(name).stem.lower()
    if stem in WINDOWS_RESERVED:
        return f"'{name}' is a reserved Windows name and cannot be used."
    if INVALID_CHARS_RE.search(name):
        return f"'{name}' contains characters not allowed in file names."
    if len(name) > 255:
        return "That name is too long. Please use a shorter name."
    if name in ("", ".", ".."):
        return "That's not a valid file or folder name."
    if name.endswith((" ", ".")):
        return "File names cannot end with a space or dot."
    return None


def _classify_error(e: Exception, path: Path) -> str:
    err = str(e).lower()
    if isinstance(e, PermissionError):
        if path.exists():
            return f"'{path.name}' is currently in use or you don't have permission. Try closing it first."
        return f"Permission denied for '{path.name}'."
    if isinstance(e, FileNotFoundError):
        return f"'{path.name}' was not found."
    if isinstance(e, FileExistsError):
        return f"'{path.name}' already exists."
    if "winerror 32" in err or "being used" in err:
        return f"'{path.name}' is open in another program. Please close it and try again."
    if "winerror 5" in err:
        return f"Access denied for '{path.name}'. You may need administrator rights."
    return f"Operation failed: {e}"


# ─────────────────────────────────────────────────────────────────
# OPERATIONS
# ─────────────────────────────────────────────────────────────────

def list_items(path: str = ".") -> Result:
    """List files and folders in a directory."""
    target = _resolve(path)

    if not _is_allowed(target):
        return Result(False, "I'm not allowed to list files there.")

    if not target.exists():
        return Result(False, f"The folder '{target.name}' doesn't exist.")

    if not target.is_dir():
        return Result(False, f"'{target.name}' is a file, not a folder.")

    try:
        with _lock:
            entries = list(target.iterdir())
    except PermissionError:
        return Result(False, f"Permission denied for '{target}'.")

    folders = sorted([e for e in entries if e.is_dir()],  key=lambda x: x.name.lower())
    files   = sorted([e for e in entries if e.is_file()], key=lambda x: x.name.lower())
    total   = len(folders) + len(files)

    if total == 0:
        return Result(True, f"'{target.name}' is empty.", data={"folders": [], "files": []})

    parts = []
    if folders: parts.append(f"{len(folders)} folder{'s' if len(folders) != 1 else ''}")
    if files:   parts.append(f"{len(files)} file{'s' if len(files) != 1 else ''}")
    summary = f"Found {' and '.join(parts)} in {target.name}."

    shown_folders = [f.name for f in folders[:LIST_CAP]]
    shown_files   = [f.name for f in files[:LIST_CAP]]
    truncated     = total > LIST_CAP

    if truncated:
        summary += f" Showing the first {LIST_CAP}. Ask me to find something specific."

    log.info("Listed %s: %d folders, %d files", target, len(folders), len(files))

    return Result(True, summary, data={
        "path":      str(target),
        "folders":   shown_folders,
        "files":     shown_files,
        "total":     total,
        "truncated": truncated,
    })


def create_file(name: str, path: str = ".") -> Result:
    """Create an empty file inside BASE_DIR."""
    err = _validate_name(name)
    if err:
        return Result(False, err)

    _ensure_base_dir()
    target = _resolve(os.path.join(path, name))

    if not _is_allowed(target):
        return Result(False, "I can only create files inside your ProVA storage folder.")

    if target.exists():
        return Result(False, f"'{name}' already exists. Choose a different name.")

    try:
        with _lock:
            target.touch()
        log.info("Created file: %s", target)
        return Result(True, f"File '{name}' created in {target.parent.name}.")
    except Exception as e:
        return Result(False, _classify_error(e, target))


def create_folder(name: str, path: str = ".") -> Result:
    """Create a new folder inside BASE_DIR."""
    err = _validate_name(Path(name).parts[0])
    if err:
        return Result(False, err)

    _ensure_base_dir()
    target = _resolve(os.path.join(path, name))

    if not _is_allowed(target):
        return Result(False, "I can only create folders inside your ProVA storage folder.")

    if target.exists():
        return Result(False, f"A folder named '{name}' already exists.")

    try:
        with _lock:
            target.mkdir(parents=True, exist_ok=False)
        log.info("Created folder: %s", target)
        return Result(True, f"Folder '{name}' created in {target.parent.name}.")
    except Exception as e:
        return Result(False, _classify_error(e, target))


def delete_file(name: str, path: str = ".", confirmed: bool = False) -> Result:
    """Delete a file. Requires confirmation."""
    target = _resolve(os.path.join(path, name))

    if not _is_allowed(target):
        return Result(False, "I'm not allowed to delete files there.")

    if not target.exists():
        return Result(False, f"'{name}' was not found.")

    if not target.is_file():
        return Result(False, f"'{name}' is a folder. Use 'delete folder' instead.")

    if not confirmed:
        return Result(False, "", needs_confirmation=True,
                      confirm_prompt=f"Are you sure you want to delete '{name}'? This cannot be undone.")

    try:
        with _lock:
            target.unlink()
        log.info("Deleted file: %s", target)
        return Result(True, f"'{name}' deleted.")
    except Exception as e:
        return Result(False, _classify_error(e, target))


def delete_folder(name: str, path: str = ".", confirmed: bool = False) -> Result:
    """Delete a folder and all its contents. Always requires confirmation."""
    target = _resolve(os.path.join(path, name))

    if not _is_allowed(target):
        return Result(False, "I'm not allowed to delete folders there.")

    if not target.exists():
        return Result(False, f"Folder '{name}' was not found.")

    if not target.is_dir():
        return Result(False, f"'{name}' is a file. Use 'delete file' instead.")

    try:
        contents     = list(target.rglob("*"))
        file_count   = sum(1 for x in contents if x.is_file())
        folder_count = sum(1 for x in contents if x.is_dir())
    except Exception:
        file_count = folder_count = 0

    if not confirmed:
        detail = ""
        if file_count or folder_count:
            detail = (f" It contains {file_count} file{'s' if file_count != 1 else ''} "
                      f"and {folder_count} subfolder{'s' if folder_count != 1 else ''}.")
        return Result(False, "", needs_confirmation=True,
                      confirm_prompt=f"Delete folder '{name}'?{detail} Everything inside will be permanently removed.")

    try:
        with _lock:
            shutil.rmtree(target)
        log.info("Deleted folder: %s (%d files, %d folders)", target, file_count, folder_count)
        return Result(True, f"Folder '{name}' and all its contents deleted.")
    except Exception as e:
        return Result(False, _classify_error(e, target))


def rename_item(old_name: str, new_name: str, path: str = ".") -> Result:
    """Rename a file or folder."""
    if not old_name:
        return Result(False, "Please tell me which file or folder to rename.")
    if not new_name:
        return Result(False, "Please tell me what to rename it to.")

    err = _validate_name(new_name)
    if err:
        return Result(False, err)

    src  = _resolve(os.path.join(path, old_name))
    dest = _resolve(os.path.join(path, new_name))

    if not _is_allowed(src) or not _is_allowed(dest):
        return Result(False, "I can only rename items inside your ProVA workspace.")

    if not src.exists():
        # Try a fuzzy name search within the same directory
        parent = src.parent
        if parent.is_dir():
            candidates = [p for p in parent.iterdir()
                          if old_name.lower() in p.name.lower()]
            if len(candidates) == 1:
                src  = candidates[0]
                dest = src.parent / new_name
            elif len(candidates) > 1:
                names = ", ".join(p.name for p in candidates[:5])
                return Result(False, f"Found multiple items matching '{old_name}': {names}. Please be more specific.")
            else:
                return Result(False, f"'{old_name}' was not found in {parent.name}.")
        else:
            return Result(False, f"'{old_name}' was not found.")

    if dest.exists():
        return Result(False, f"'{new_name}' already exists. Please choose a different name.")

    try:
        with _lock:
            src.rename(dest)
        log.info("Renamed: %s → %s", src, dest)
        return Result(True, f"Renamed '{src.name}' to '{new_name}'.")
    except Exception as e:
        return Result(False, _classify_error(e, src))


def copy_item(name: str, dest_path: str, src_path: str = ".") -> Result:
    """
    Copy a file or folder to a destination.
    dest_path can be a known location name ("desktop") or an absolute path.
    """
    # Detect missing destination — dest_path "." means the user didn't specify one
    if not dest_path or dest_path in (".", ""):
        return Result(False,
            f"Please tell me where to copy '{name}' to. "
            f"For example: copy {name} to desktop, or copy {name} to downloads.")

    src  = _resolve(os.path.join(src_path, name))
    dest = resolve_location(dest_path) or _resolve(dest_path)

    if not _is_allowed(src):
        return Result(False, f"'{name}' is not in an allowed location.")

    if not _is_allowed(dest):
        return Result(False, f"I'm not allowed to copy files to '{dest_path}'.")

    if not src.exists():
        return Result(False, f"'{name}' was not found.")

    if dest.is_dir():
        dest = dest / src.name

    if dest.exists():
        return Result(False, f"'{dest.name}' already exists at the destination.")

    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        with _lock:
            if src.is_dir():
                shutil.copytree(src, dest)
            else:
                shutil.copy2(src, dest)
        log.info("Copied: %s → %s", src, dest)
        return Result(True, f"Copied '{name}' to '{dest.parent.name}'.")
    except Exception as e:
        return Result(False, _classify_error(e, src))


def move_item(name: str, dest_path: str, src_path: str = ".", confirmed: bool = False) -> Result:
    """
    Move a file or folder to a destination.
    dest_path can be a known location name ("downloads") or a path.
    """
    src  = _resolve(os.path.join(src_path, name))
    dest = resolve_location(dest_path) or _resolve(dest_path)

    if not _is_allowed(src):
        return Result(False, f"'{name}' is not in an allowed location.")

    if not _is_allowed(dest):
        return Result(False, f"I'm not allowed to move files to '{dest_path}'.")

    if not src.exists():
        return Result(False, f"'{name}' was not found.")

    final_dest = dest / src.name if dest.is_dir() else dest

    if final_dest.exists() and not confirmed:
        return Result(False, "", needs_confirmation=True,
                      confirm_prompt=f"'{final_dest.name}' already exists at the destination. Overwrite it?")

    try:
        final_dest.parent.mkdir(parents=True, exist_ok=True)
        with _lock:
            shutil.move(str(src), str(final_dest))
        log.info("Moved: %s → %s", src, final_dest)
        return Result(True, f"Moved '{name}' to '{dest.name}'.")
    except Exception as e:
        return Result(False, _classify_error(e, src))


def get_info(name: str, path: str = ".") -> Result:
    """Get human-readable info including full path."""
    target = _resolve(os.path.join(path, name))

    if not _is_allowed(target):
        return Result(False, "I can only read info about items in your ProVA workspace.")

    if not target.exists():
        # Try fuzzy match
        parent = target.parent
        if parent.is_dir():
            candidates = [p for p in parent.iterdir()
                          if name.lower() in p.name.lower()]
            if len(candidates) == 1:
                target = candidates[0]
            elif len(candidates) > 1:
                names = ", ".join(p.name for p in candidates[:5])
                return Result(False, f"Found multiple items matching '{name}': {names}.")
            else:
                return Result(False, f"'{name}' was not found.")
        else:
            return Result(False, f"'{name}' was not found.")

    try:
        stat     = target.stat()
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%B %d, %Y at %I:%M %p")
        full_path = str(target)

        if target.is_file():
            size_bytes = stat.st_size
            if size_bytes < 1024:
                size_str = f"{size_bytes} bytes"
            elif size_bytes < 1024 ** 2:
                size_str = f"{size_bytes / 1024:.1f} kilobytes"
            else:
                size_str = f"{size_bytes / 1024**2:.1f} megabytes"

            suffix = target.suffix.lstrip(".").upper() or "unknown type"
            msg = (f"'{target.name}' is a {suffix} file, {size_str}, "
                   f"last modified {modified}. "
                   f"Full path: {full_path}.")
        else:
            contents  = list(target.iterdir())
            n_files   = sum(1 for x in contents if x.is_file())
            n_folders = sum(1 for x in contents if x.is_dir())
            msg = (f"'{target.name}' is a folder containing "
                   f"{n_files} file{'s' if n_files != 1 else ''} and "
                   f"{n_folders} subfolder{'s' if n_folders != 1 else ''}, "
                   f"last modified {modified}. "
                   f"Full path: {full_path}.")

        log.info("Info: %s", target)
        return Result(True, msg, data={"path": full_path, "modified": modified})

    except Exception as e:
        return Result(False, _classify_error(e, target))


def find_files(pattern: str, search_path: str = ".") -> Result:
    """
    Search for files/folders matching a name pattern (case-insensitive).
    Reports full paths so the user knows exactly where items are.
    Searches both BASE_DIR and all ALLOWED_ROOTS.
    """
    # If search_path is "." use BASE_DIR; if it's a known location, resolve it
    if search_path in (".", ""):
        base = BASE_DIR
    else:
        base = resolve_location(search_path) or _resolve(search_path)

    if not base.is_dir():
        return Result(False, f"'{base.name}' is not a valid folder to search in.")

    pattern_lower = pattern.lower().strip()

    # Support wildcards: "*.txt" → match by suffix; else substring match
    try:
        matches = []
        for p in base.rglob("*"):
            name_lower = p.name.lower()
            if "*" in pattern_lower:
                # Glob-style: "*.txt" → p.suffix == ".txt"
                stub = pattern_lower.replace("*", "")
                if stub in name_lower:
                    matches.append(p)
            else:
                if pattern_lower in name_lower:
                    matches.append(p)

        matches = matches[:LIST_CAP]

        if not matches:
            return Result(False,
                          f"No files matching '{pattern}' found in {base.name}. "
                          f"Try searching a specific location, e.g. 'find {pattern} in downloads'.")

        # Spoken: name + location for first few results
        spoken_items = []
        for m in matches[:5]:
            rel = m.relative_to(base) if m.is_relative_to(base) else m
            spoken_items.append(f"{m.name} in {m.parent.name}")

        count = len(matches)
        msg = (f"Found {count} match{'es' if count != 1 else ''} for '{pattern}'. "
               + "; ".join(spoken_items)
               + ("." if not spoken_items[-1].endswith(".") else ""))

        if count > 5:
            msg += f" And {count - 5} more."

        log.info("Find '%s' in %s: %d results", pattern, base, count)
        return Result(True, msg, data={
            "matches": [m.name for m in matches],
            "paths":   [str(m) for m in matches],
        })

    except Exception as e:
        return Result(False, f"Search failed: {e}")


# ─────────────────────────────────────────────────────────────────
# VOICE CONFIRMATION FLOW
# ─────────────────────────────────────────────────────────────────

def _handle_confirmation(result: Result, cmd, speak_fn, confirm_fn) -> None:
    """
    Pass confirm_prompt to confirm_fn, which arms the gate_event BEFORE speaking it
    and then listens for yes/no.  Retries once on silence or unclear answer.

    Race-fix: the old code called speak_fn(prompt) here and then confirm_fn().
    That left a window between TTS finishing and the confirm listener arming the
    mic lock where the main voice loop could race in and capture "yes" as an
    unknown command.  Passing the prompt INTO confirm_fn (which sets gate_event
    first) closes that window completely.
    """
    confirmed = confirm_fn(result.confirm_prompt)

    if not confirmed:
        speak_fn("Cancelled.")
        return

    action   = cmd.action
    name     = cmd.target or ""
    path     = cmd.extra.get("path", ".")
    dest     = cmd.extra.get("dest", ".")

    if action == "delete_file":
        final = delete_file(name, path, confirmed=True)
    elif action == "delete_folder":
        final = delete_folder(name, path, confirmed=True)
    elif action == "move":
        final = move_item(name, dest, path, confirmed=True)
    else:
        final = Result(False, "Unknown confirmed action.")

    speak_fn(final.message)


# ─────────────────────────────────────────────────────────────────
# MAIN HANDLER — called by voice_module router
# ─────────────────────────────────────────────────────────────────

def handle(cmd, speak_fn, confirm_fn=None) -> None:
    """
    Entry point called by voice_module.route() via run_async.

    cmd.action values:
        list          → cmd.target = folder path (optional)
        create_file   → cmd.target = filename
        create_folder → cmd.target = folder name
        delete_file   → cmd.target = filename
        delete_folder → cmd.target = folder name
        rename        → cmd.target = old name, cmd.extra["new_name"] = new name
        copy          → cmd.target = item name, cmd.extra["dest"] = destination
        move          → cmd.target = item name, cmd.extra["dest"] = destination
        info          → cmd.target = item name
        find          → cmd.target = search pattern
    """
    action   = cmd.action or ""
    raw_text = cmd.raw or ""

    # ── Re-parse name and destination from raw text ──────────────
    # The generic parser often leaves articles ("a", "the", "my") in
    # cmd.target. Re-parse here for precision.
    if action in ("create_file", "create_folder", "delete_file", "delete_folder",
                  "find", "info", "list"):
        name = parse_file_target(raw_text) or (cmd.target or "").strip()
    elif action == "rename":
        old, new = parse_rename_parts(raw_text)
        name     = old
        if new:
            cmd.extra["new_name"] = new
        else:
            name = (cmd.target or "").strip()
    elif action in ("copy", "move"):
        name = parse_file_target(raw_text) or (cmd.target or "").strip()
        # Re-parse destination from raw text → resolve to real path
        spoken_dest = parse_destination(raw_text) or cmd.extra.get("dest", "")
        if spoken_dest:
            resolved = resolve_location(spoken_dest)
            cmd.extra["dest"] = str(resolved) if resolved else spoken_dest
    else:
        name = (cmd.target or "").strip()

    path     = cmd.extra.get("path", ".")
    dest     = cmd.extra.get("dest", ".")
    new_name = cmd.extra.get("new_name", "")

    _ensure_base_dir()

    dispatch = {
        "list":          lambda: list_items(path if not name else name),
        "create_file":   lambda: create_file(name, path),
        "create_folder": lambda: create_folder(name, path),
        "delete_file":   lambda: delete_file(name, path, confirmed=False),
        "delete_folder": lambda: delete_folder(name, path, confirmed=False),
        "rename":        lambda: rename_item(name, new_name, path),
        "copy":          lambda: copy_item(name, dest, path),
        "move":          lambda: move_item(name, dest, path, confirmed=False),
        "info":          lambda: get_info(name, path),
        "find":          lambda: find_files(name, path),
    }

    fn = dispatch.get(action)
    if fn is None:
        speak_fn(f"I don't know how to '{action}' files.")
        return

    # For "list files" / "list folder" / "show files": the word "files"/"folder"
    # is part of the trigger phrase, not a real directory name.  Clear it so
    # list_items() uses the default storage dir instead of looking for a
    # directory literally named "files".
    _LIST_NOISE = {"file", "files", "folder", "folders", "directory", "directories",
                   "content", "contents", "all", "everything"}
    if action == "list" and name.lower() in _LIST_NOISE:
        name = ""

    if not name and action not in ("list",):
        speak_fn("Please tell me which file or folder.")
        return

    result = fn()

    if result.needs_confirmation:
        if confirm_fn is None:
            speak_fn("Confirmation required for that operation.")
            return
        _handle_confirmation(result, cmd, speak_fn, confirm_fn)
        return

    speak_fn(result.message)

    # Read out item names for small listings
    if result.success and action == "list" and result.data:
        folders = result.data.get("folders", [])
        files   = result.data.get("files",   [])
        total   = result.data.get("total",   0)
        if 0 < total <= 10:
            if folders: speak_fn("Folders: " + ", ".join(folders) + ".")
            if files:   speak_fn("Files: "   + ", ".join(files)   + ".")
        elif total > 10:
            speak_fn("That's a lot to read out. Ask me to find something specific.")

    if result.success and action == "find" and result.data:
        paths = result.data.get("paths", [])
        if len(paths) == 1:
            speak_fn(f"Full path: {paths[0]}.")
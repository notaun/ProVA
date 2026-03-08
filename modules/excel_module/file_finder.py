"""
ProVA — modules/excel_module/file_finder.py
============================================
Searches for data files by spoken filename across known directories.
Falls back to a Windows file picker dialog if not found.

Search order (fastest first, then broader):
  1. ~/ProVA/           (ProVA workspace)
  2. ~/Desktop/
  3. ~/Documents/
  4. ~/Downloads/
  5. ~/OneDrive/        (if exists)
  6. Recursive sub-search of the above (up to 2 levels deep)

Supported file types: .xlsx, .xls, .csv, .txt
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

log = logging.getLogger("ProVA.Excel.FileFinder")

# Directories to search, in priority order
# Project root resolved from this file's location (works regardless of where ProVA is installed)
_PROJECT_ROOT = Path(__file__).parent.parent.parent

_SEARCH_ROOTS = [
    _PROJECT_ROOT,                    # ProVA_final_v3/ itself
    _PROJECT_ROOT / "modules" / "excel_module" / "data",  # bundled datasets
    Path.home() / "Desktop",
    Path.home() / "Documents",
    Path.home() / "Downloads",
    Path.home() / "OneDrive",
]

_SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv", ".txt"}


def find_data_file(spoken_name: str) -> Optional[Path]:
    """
    Find a data file by spoken name.

    Matching strategy:
      1. Exact filename match (case-insensitive, with or without extension)
      2. Partial name match (spoken name is substring of filename)
      3. Fuzzy: each word of spoken name appears in filename

    Returns the best matching Path, or None if nothing found.
    """
    name_lower = spoken_name.lower().strip()

    # Strip common spoken filler
    for filler in ("file", "spreadsheet", "excel", "the", "my", "dot", "xlsx", "csv"):
        name_lower = name_lower.replace(filler, "").strip()

    name_lower = name_lower.strip()
    if not name_lower:
        return None

    candidates: list[tuple[int, Path]] = []   # (score, path)

    for root in _SEARCH_ROOTS:
        if not root.exists():
            continue

        # Flat search first (fast), then recurse 2 levels
        for depth, pattern in [(0, "*"), (1, "*/*"), (2, "*/*/*")]:
            for p in root.glob(pattern):
                if not p.is_file():
                    continue
                if p.suffix.lower() not in _SUPPORTED_SUFFIXES:
                    continue

                stem_lower = p.stem.lower()
                full_lower = p.name.lower()

                # Score the match quality
                if stem_lower == name_lower or full_lower == name_lower:
                    score = 100                          # exact
                elif name_lower in stem_lower:
                    score = 80 - depth * 5              # substring, penalise depth
                elif all(w in stem_lower for w in name_lower.split()):
                    score = 60 - depth * 5              # all words present
                else:
                    continue

                candidates.append((score, p))

    if not candidates:
        log.info("File finder: no match for '%s'", spoken_name)
        return None

    # Return highest-scoring match (ties broken by earliest in list = shallower)
    candidates.sort(key=lambda x: -x[0])
    best = candidates[0][1]
    log.info("File finder: '%s' → %s (score=%d)", spoken_name, best, candidates[0][0])
    return best


def pick_file_dialog() -> Optional[Path]:
    """
    Open a Windows file picker dialog so user can manually select a file.
    Returns the selected Path, or None if cancelled.

    Used as fallback when find_data_file() returns None.
    Falls back gracefully if tkinter is not available.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()          # hide the empty tk window
        root.attributes("-topmost", True)

        path_str = filedialog.askopenfilename(
            title="ProVA — Select your data file",
            filetypes=[
                ("Excel files",       "*.xlsx *.xls"),
                ("CSV files",         "*.csv"),
                ("Text files",        "*.txt"),
                ("All supported",     "*.xlsx *.xls *.csv *.txt"),
            ],
            initialdir=str(_PROJECT_ROOT),
        )
        root.destroy()

        if path_str:
            log.info("File picker: user selected %s", path_str)
            return Path(path_str)

    except Exception as e:
        log.warning("File picker dialog failed: %s", e)

    return None
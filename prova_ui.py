"""
prova_ui.py — ProVA desktop UI  (v3 — full redesign)

New in this version:
  ─ THEMES: dark (default) + light mode, switchable live
  ─ SETTINGS PANEL: voice selector, TTS rate, TTS volume,
      wake-word toggle, noise-suppression toggle, theme switcher
  ─ VOICE SELECTOR: enumerates all installed SAPI voices at runtime,
      lets user pick and test any voice
  ─ STATE BAR: rich per-state animated banner replacing the thin listen bar
      IDLE / CALIBRATING / LISTENING / THINKING / EXECUTING / SPEAKING / PAUSED / ERROR
  ─ CHAT READABILITY: larger font, bubble-style rows with alternating
      alignment, icon prefix per speaker, timestamp always visible,
      full-width highlight row on hover
  ─ NOISE SUPPRESSION: optional noisereduce pass on captured audio
      (toggled in Settings; graceful no-op if library absent)
  ─ SEND button always visible; Enter sends; Shift+Enter inserts newline
  ─ "THINKING…" and "EXECUTING…" animated dots so users know work is happening
"""

from __future__ import annotations

import json
import os
import queue
import threading
import tkinter as tk
import tkinter.ttk as ttk
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────────
# THEME DEFINITIONS
# ─────────────────────────────────────────────────────────────────
THEMES: dict[str, dict] = {
    "dark": {
        "BG_DEEP":     "#0d0e10",
        "BG_PANEL":    "#13151a",
        "BG_INPUT":    "#1c1f26",
        "BG_BUBBLE_U": "#1a2540",   # user bubble background
        "BG_BUBBLE_P": "#1e1608",   # prova bubble background
        "BG_SETTINGS": "#0f1115",
        "BORDER":      "#272b35",
        "ACCENT":      "#f0a500",
        "ACCENT_DIM":  "#6b4800",
        "ACCENT2":     "#3b82f6",   # secondary — listen blue
        "TEXT_PRI":    "#e8e9eb",
        "TEXT_SEC":    "#5a6175",
        "TEXT_USER":   "#93c5fd",
        "TEXT_PROVA":  "#f0a500",
        "TEXT_SYS":    "#4ade80",
        "TEXT_ERR":    "#f87171",
        "TEXT_BUBBLE_U": "#c5d9ff",
        "TEXT_BUBBLE_P": "#ffd98a",
        "STATE_IDLE":      "#4ade80",
        "STATE_LISTEN":    "#3b82f6",
        "STATE_THINK":     "#f59e0b",
        "STATE_EXEC":      "#a78bfa",
        "STATE_SPEAK":     "#ec4899",
        "STATE_CALIB":     "#06b6d4",
        "STATE_PAUSE":     "#6b7280",
        "STATE_ERR":       "#f87171",
        "BAR_IDLE":    "#0d1a0e",
        "BAR_LISTEN":  "#0d1630",
        "BAR_THINK":   "#1a1400",
        "BAR_EXEC":    "#130d1f",
        "BAR_SPEAK":   "#1a0812",
        "BAR_CALIB":   "#041215",
        "BAR_PAUSE":   "#111318",
        "BAR_ERR":     "#1a0808",
    },
    "light": {
        "BG_DEEP":     "#f4f5f7",
        "BG_PANEL":    "#ffffff",
        "BG_INPUT":    "#eef0f4",
        "BG_BUBBLE_U": "#dbeafe",
        "BG_BUBBLE_P": "#fef3c7",
        "BG_SETTINGS": "#e9ebef",
        "BORDER":      "#d1d5db",
        "ACCENT":      "#d97706",
        "ACCENT_DIM":  "#fde68a",
        "ACCENT2":     "#2563eb",
        "TEXT_PRI":    "#111827",
        "TEXT_SEC":    "#6b7280",
        "TEXT_USER":   "#1d4ed8",
        "TEXT_PROVA":  "#b45309",
        "TEXT_SYS":    "#059669",
        "TEXT_ERR":    "#dc2626",
        "TEXT_BUBBLE_U": "#1e3a8a",
        "TEXT_BUBBLE_P": "#78350f",
        "STATE_IDLE":      "#059669",
        "STATE_LISTEN":    "#2563eb",
        "STATE_THINK":     "#d97706",
        "STATE_EXEC":      "#7c3aed",
        "STATE_SPEAK":     "#db2777",
        "STATE_CALIB":     "#0891b2",
        "STATE_PAUSE":     "#6b7280",
        "STATE_ERR":       "#dc2626",
        "BAR_IDLE":    "#dcfce7",
        "BAR_LISTEN":  "#dbeafe",
        "BAR_THINK":   "#fef3c7",
        "BAR_EXEC":    "#ede9fe",
        "BAR_SPEAK":   "#fce7f3",
        "BAR_CALIB":   "#cffafe",
        "BAR_PAUSE":   "#f3f4f6",
        "BAR_ERR":     "#fee2e2",
    },
}

# State → (icon, label, bar_key, dot_key, animate?)
STATE_META: dict[str, tuple] = {
    "IDLE":        ("●",  "IDLE",         "BAR_IDLE",  "STATE_IDLE",  False),
    "CALIBRATING": ("⟳",  "CALIBRATING",  "BAR_CALIB", "STATE_CALIB", True),
    "LISTENING":   ("🎤", "LISTENING",    "BAR_LISTEN","STATE_LISTEN",True),
    "THINKING":    ("⟳",  "THINKING",     "BAR_THINK", "STATE_THINK", True),
    "EXECUTING":   ("⚙",  "EXECUTING",    "BAR_EXEC",  "STATE_EXEC",  True),
    "BUSY":        ("⚙",  "EXECUTING",    "BAR_EXEC",  "STATE_EXEC",  True),
    "SPEAKING":    ("🔊", "SPEAKING",     "BAR_SPEAK", "STATE_SPEAK", False),
    "PAUSED":      ("⏸",  "PAUSED",       "BAR_PAUSE", "STATE_PAUSE", False),
    "STOPPED":     ("■",  "STOPPED",      "BAR_IDLE",  "STATE_PAUSE", False),
    "ERROR":       ("⚠",  "ERROR",        "BAR_ERR",   "STATE_ERR",   True),
}

STATE_HINTS: dict[str, str] = {
    "IDLE":        "Say  'hey ProVA'  to give a command",
    "CALIBRATING": "Measuring background noise — please stay quiet…",
    "LISTENING":   "Listening now — speak your command",
    "THINKING":    "Working on it…",
    "EXECUTING":   "Running your command…",
    "BUSY":        "Running your command…",
    "SPEAKING":    "ProVA is responding…",
    "PAUSED":      "Paused — click RESUME to continue",
    "STOPPED":     "ProVA has stopped",
    "ERROR":       "Something went wrong — check the log",
}

FONT_UI   = ("Segoe UI", 10)
FONT_CHAT = ("Segoe UI", 13)
FONT_TIME = ("Segoe UI", 9)
FONT_HEAD = ("Segoe UI", 13, "bold")
FONT_BTN  = ("Segoe UI", 10, "bold")
FONT_BAR  = ("Segoe UI", 11, "bold")
FONT_HINT = ("Segoe UI", 9)
FONT_SET  = ("Segoe UI", 10)

PREFS_PATH = Path(__file__).resolve().parent / "data" / "ui_prefs.json"


def _load_prefs() -> dict:
    try:
        if PREFS_PATH.exists():
            return json.loads(PREFS_PATH.read_text())
    except Exception:
        pass
    return {}


def _save_prefs(d: dict):
    try:
        PREFS_PATH.parent.mkdir(parents=True, exist_ok=True)
        PREFS_PATH.write_text(json.dumps(d, indent=2))
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────
# MESSAGE MODEL
# ─────────────────────────────────────────────────────────────────
class _Msg:
    def __init__(self, kind: str, text: str):
        self.kind = kind   # "user" | "prova" | "system"
        self.text = text
        self.ts   = datetime.now().strftime("%H:%M")


# ─────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────
class ProVAApp:
    _PLACEHOLDER = "Type a command, or speak…"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("ProVA")
        self.root.minsize(680, 600)
        self.root.geometry("820x740")

        # ── state ────────────────────────────────────────────────
        self._prefs              = _load_prefs()
        self._theme_name         = self._prefs.get("theme", "dark")
        self._T                  = THEMES[self._theme_name]
        self._voice_thread       = None
        self._stop_event         = threading.Event()
        self._pause_event        = threading.Event()
        self._text_queue         = queue.Queue()
        self._ui_queue           = queue.Queue()
        self._muted              = [False]
        self._paused             = False
        self._running            = False
        self._current_status     = "IDLE"
        self._dot_anim_id        = None
        self._dot_frame          = 0
        self._settings_open      = False
        self._input_has_placeholder = True

        # ── voice list (populated lazily) ────────────────────────

        # ── noise suppression toggle ─────────────────────────────
        self._noise_suppress = tk.BooleanVar(
            value=self._prefs.get("noise_suppress", False))

        self._build_ui()
        self._apply_theme(self._theme_name, initial=True)
        self._poll_ui_queue()
        # Enumerate voices in background so startup isn't delayed

    # ─────────────────────────────────────────────────────────────
    # UI CONSTRUCTION
    # ─────────────────────────────────────────────────────────────
    def _build_ui(self):
        r = self.root

        # ── separator helpers ─────────────────────────────────────
        def sep(parent, side=tk.BOTTOM):
            self._seps = getattr(self, "_seps", [])
            f = tk.Frame(parent, height=1)
            f.pack(side=side, fill=tk.X)
            self._seps.append(f)
            return f

        # ── BOTTOM: controls ──────────────────────────────────────
        sep(r, tk.BOTTOM)
        self._ctrl_frame = tk.Frame(r, height=64)
        self._ctrl_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self._ctrl_frame.pack_propagate(False)

        self._start_btn = tk.Button(
            self._ctrl_frame, text="▶  START",
            font=FONT_BTN, relief=tk.FLAT,
            cursor="hand2", padx=20, pady=5,
            command=self._start,
        )
        self._start_btn.pack(side=tk.LEFT, padx=(16, 6), pady=14)

        self._pause_btn = tk.Button(
            self._ctrl_frame, text="⏸  PAUSE",
            font=FONT_BTN, relief=tk.FLAT,
            cursor="hand2", padx=20, pady=5, state=tk.DISABLED,
            command=self._toggle_pause,
        )
        self._pause_btn.pack(side=tk.LEFT, padx=6, pady=14)

        self._stop_btn = tk.Button(
            self._ctrl_frame, text="■  STOP",
            font=FONT_BTN, relief=tk.FLAT,
            cursor="hand2", padx=20, pady=5, state=tk.DISABLED,
            command=self._stop,
        )
        self._stop_btn.pack(side=tk.LEFT, padx=6, pady=14)

        self._wake_var = tk.BooleanVar(value=self._prefs.get("wake_word", True))
        self._wake_chk = tk.Checkbutton(
            self._ctrl_frame, text="Wake word",
            variable=self._wake_var,
            font=FONT_UI, cursor="hand2",
            command=self._toggle_wake_word,
        )
        self._wake_chk.pack(side=tk.RIGHT, padx=(0, 16), pady=14)

        # ── BOTTOM: input row ────────────────────────────────────
        sep(r, tk.BOTTOM)
        self._inp_frame = tk.Frame(r, height=50)
        self._inp_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self._inp_frame.pack_propagate(False)

        self._send_btn = tk.Button(
            self._inp_frame, text="SEND",
            font=FONT_BTN, relief=tk.FLAT,
            cursor="hand2", padx=14,
            command=self._send_typed,
        )
        self._send_btn.pack(side=tk.RIGHT, padx=(0, 12), pady=10)

        self._input_var = tk.StringVar()
        self._input_var.trace_add("write", self._on_input_change)
        self._input = tk.Entry(
            self._inp_frame, textvariable=self._input_var,
            font=FONT_CHAT, relief=tk.FLAT, bd=0,
        )
        self._input.pack(side=tk.LEFT, fill=tk.BOTH, expand=True,
                         padx=(14, 8), pady=14)
        self._input.bind("<Return>",    lambda e: self._send_typed())
        self._input.bind("<FocusIn>",   self._clear_ph)
        self._input.bind("<FocusOut>",  self._restore_ph)
        # Clicking the chat area returns focus to root so OS voice dictation
        # (Windows Speech Recognition) can't type into the input box while
        # ProVA is speaking — that caused "Type a command, or sOpen Spotifypeak..."
        self._set_placeholder()

        # ── BOTTOM: state bar ─────────────────────────────────────
        self._state_bar = tk.Frame(r, height=0)
        self._state_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self._state_bar.pack_propagate(False)

        self._state_icon_lbl = tk.Label(
            self._state_bar, font=("Segoe UI Emoji", 13),
        )
        self._state_icon_lbl.pack(side=tk.LEFT, padx=(14, 4))

        self._state_name_lbl = tk.Label(
            self._state_bar, font=FONT_BAR,
        )
        self._state_name_lbl.pack(side=tk.LEFT)

        self._state_dot_lbl = tk.Label(
            self._state_bar, text="", font=FONT_BAR, width=3,
        )
        self._state_dot_lbl.pack(side=tk.LEFT)

        self._state_hint_lbl = tk.Label(
            self._state_bar, font=FONT_HINT,
        )
        self._state_hint_lbl.pack(side=tk.RIGHT, padx=14)

        # ── TOP: header ───────────────────────────────────────────
        self._hdr = tk.Frame(r, height=48)
        self._hdr.pack(side=tk.TOP, fill=tk.X)
        self._hdr.pack_propagate(False)

        tk.Label(
            self._hdr, text="◆ ProVA",
            font=FONT_HEAD, padx=16,
        ).pack(side=tk.LEFT, pady=10)

        # Settings gear
        self._settings_btn = tk.Button(
            self._hdr, text="⚙  SETTINGS",
            font=FONT_BTN, relief=tk.FLAT,
            cursor="hand2", padx=10,
            command=self._toggle_settings,
        )
        self._settings_btn.pack(side=tk.RIGHT, padx=(0, 10), pady=10)

        tk.Button(
            self._hdr, text="HELP",
            font=FONT_BTN, relief=tk.FLAT,
            cursor="hand2", padx=10,
            command=self._send_help,
        ).pack(side=tk.RIGHT, padx=(0, 4), pady=10)

        sep(r, tk.TOP)

        # ── MIDDLE: settings panel (hidden by default) ───────────
        self._settings_panel = tk.Frame(r)
        # NOT packed yet — toggled on demand

        self._build_settings_panel()

        # ── MIDDLE: chat ──────────────────────────────────────────
        self._chat_outer = tk.Frame(r)
        self._chat_outer.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self._chat = tk.Text(
            self._chat_outer,
            font=FONT_CHAT, relief=tk.FLAT, bd=0,
            wrap=tk.WORD, state=tk.DISABLED, cursor="arrow",
            padx=14, pady=10, spacing1=3, spacing3=8,
        )
        self._chat.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        # Clicking chat returns focus to root so Windows voice dictation
        # can't type into the input box while ProVA is listening/speaking
        self._chat.bind("<Button-1>", lambda e: self.root.focus_set())

        sb = tk.Scrollbar(
            self._chat_outer, command=self._chat.yview,
            relief=tk.FLAT, bd=0, width=10,
        )
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._chat.configure(yscrollcommand=sb.set)

        # Chat tags — configured again in _apply_theme
        self._chat.tag_configure("ts")
        self._chat.tag_configure("you_label")
        self._chat.tag_configure("prova_label")
        self._chat.tag_configure("user_text")
        self._chat.tag_configure("prova_text")
        self._chat.tag_configure("sys_text")
        self._chat.tag_configure("bubble_u")
        self._chat.tag_configure("bubble_p")
        self._chat.tag_configure("newline")

        self._append_msg(_Msg("system", "ProVA ready — click ▶ START"))

    def _build_settings_panel(self):
        """Build the collapsible settings panel contents."""
        T = self._T  # will be re-skinned in _apply_theme
        p = self._settings_panel

        # ── row 1: theme ──────────────────────────────────────────
        r1 = tk.Frame(p)
        r1.pack(fill=tk.X, padx=16, pady=(10, 4))
        tk.Label(r1, text="Theme:", font=FONT_SET, width=12, anchor="w").pack(side=tk.LEFT)
        self._theme_var = tk.StringVar(value=self._theme_name)
        for name, label in [("dark", "🌑 Dark"), ("light", "☀ Light")]:
            tk.Radiobutton(
                r1, text=label, variable=self._theme_var,
                value=name, font=FONT_SET, cursor="hand2",
                command=lambda n=name: self._apply_theme(n),
            ).pack(side=tk.LEFT, padx=8)

        # ── row 3: TTS rate ───────────────────────────────────────
        r3 = tk.Frame(p)
        r3.pack(fill=tk.X, padx=16, pady=4)
        tk.Label(r3, text="Speed:", font=FONT_SET, width=12, anchor="w").pack(side=tk.LEFT)
        self._rate_var = tk.IntVar(value=self._prefs.get("tts_rate", 172))
        self._rate_lbl = tk.Label(r3, text=f"{self._rate_var.get()} wpm",
                                  font=FONT_SET, width=8)
        tk.Scale(
            r3, from_=100, to=250, orient=tk.HORIZONTAL,
            variable=self._rate_var, length=200,
            showvalue=False, relief=tk.FLAT,
            command=lambda v: (
                self._rate_lbl.configure(text=f"{int(float(v))} wpm"),
                self._apply_voice_settings(),
            ),
        ).pack(side=tk.LEFT)
        self._rate_lbl.pack(side=tk.LEFT, padx=8)

        # ── row 4: TTS volume ─────────────────────────────────────
        r4 = tk.Frame(p)
        r4.pack(fill=tk.X, padx=16, pady=4)
        tk.Label(r4, text="Volume:", font=FONT_SET, width=12, anchor="w").pack(side=tk.LEFT)
        self._vol_var = tk.DoubleVar(value=self._prefs.get("tts_vol", 1.0))
        self._vol_lbl = tk.Label(r4, text=f"{int(self._vol_var.get()*100)}%",
                                 font=FONT_SET, width=8)
        tk.Scale(
            r4, from_=0.1, to=1.0, resolution=0.05, orient=tk.HORIZONTAL,
            variable=self._vol_var, length=200,
            showvalue=False, relief=tk.FLAT,
            command=lambda v: (
                self._vol_lbl.configure(text=f"{int(float(v)*100)}%"),
                self._apply_voice_settings(),
            ),
        ).pack(side=tk.LEFT)
        self._vol_lbl.pack(side=tk.LEFT, padx=8)

        # ── row 5: noise suppression ──────────────────────────────
        r5 = tk.Frame(p)
        r5.pack(fill=tk.X, padx=16, pady=(4, 10))
        tk.Label(r5, text="Noise filter:", font=FONT_SET, width=12, anchor="w").pack(side=tk.LEFT)
        self._ns_chk = tk.Checkbutton(
            r5, text="Enable noisereduce (adds ~150ms latency, improves BT clarity)",
            variable=self._noise_suppress, font=FONT_SET, cursor="hand2",
            command=self._apply_noise_setting,
        )
        self._ns_chk.pack(side=tk.LEFT)

        # separator at bottom
        self._settings_sep = tk.Frame(p, height=1)
        self._settings_sep.pack(fill=tk.X)

    # ─────────────────────────────────────────────────────────────
    # THEME ENGINE
    # ─────────────────────────────────────────────────────────────
    def _apply_theme(self, name: str, initial: bool = False):
        self._theme_name = name
        T = THEMES[name]
        self._T = T
        r = self.root
        r.configure(bg=T["BG_DEEP"])

        def w(widget, **kw):
            try: widget.configure(**kw)
            except tk.TclError: pass

        # root + structural frames
        for frame in [self._ctrl_frame, self._inp_frame, self._hdr,
                      self._chat_outer, self._settings_panel]:
            w(frame, bg=T["BG_DEEP"])

        # separators
        for s in getattr(self, "_seps", []):
            w(s, bg=T["BORDER"])

        # state bar
        w(self._state_bar, bg=T["BG_DEEP"])
        w(self._state_icon_lbl, bg=T["BG_DEEP"], fg=T["STATE_IDLE"])
        w(self._state_name_lbl, bg=T["BG_DEEP"], fg=T["STATE_IDLE"])
        w(self._state_dot_lbl,  bg=T["BG_DEEP"], fg=T["STATE_IDLE"])
        w(self._state_hint_lbl, bg=T["BG_DEEP"], fg=T["TEXT_SEC"])

        # header label "◆ ProVA"
        for child in self._hdr.winfo_children():
            if isinstance(child, tk.Label):
                w(child, bg=T["BG_DEEP"], fg=T["ACCENT"])
            elif isinstance(child, tk.Button):
                w(child, bg=T["BG_DEEP"], fg=T["TEXT_SEC"],
                  activebackground=T["BORDER"], activeforeground=T["TEXT_PRI"])

        # control buttons
        w(self._start_btn, bg=T["ACCENT"], fg=T["BG_DEEP"],
          activebackground=T["ACCENT_DIM"], activeforeground=T["TEXT_PRI"])
        w(self._pause_btn, bg=T["BG_INPUT"], fg=T["TEXT_SEC"],
          activebackground=T["BORDER"], activeforeground=T["TEXT_PRI"],
          disabledforeground=T["TEXT_SEC"])
        w(self._stop_btn,  bg=T["BG_INPUT"], fg=T["TEXT_ERR"],
          activebackground=T["BORDER"], activeforeground=T["TEXT_ERR"],
          disabledforeground=T["TEXT_SEC"])
        w(self._wake_chk,  bg=T["BG_DEEP"], fg=T["TEXT_SEC"],
          selectcolor=T["BG_INPUT"],
          activebackground=T["BG_DEEP"], activeforeground=T["TEXT_PRI"])

        # input row
        w(self._inp_frame, bg=T["BG_INPUT"])
        w(self._input, bg=T["BG_INPUT"], fg=T["TEXT_PRI"],
          insertbackground=T["ACCENT"],
          disabledbackground=T["BG_INPUT"])
        if self._input_has_placeholder:
            w(self._input, fg=T["TEXT_SEC"])
        w(self._send_btn, bg=T["ACCENT"], fg=T["BG_DEEP"],
          activebackground=T["ACCENT_DIM"], activeforeground=T["TEXT_PRI"])

        # settings panel
        w(self._settings_panel, bg=T["BG_SETTINGS"])
        w(self._settings_sep, bg=T["BORDER"])
        for child in self._settings_panel.winfo_children():
            self._theme_widget_recursive(child, T)

        # chat
        w(self._chat, bg=T["BG_PANEL"], fg=T["TEXT_PRI"])
        self._reconfigure_chat_tags(T)

        # re-skin state bar to current state
        if not initial:
            self._set_status(self._current_status)

        # persist
        self._prefs["theme"] = name
        _save_prefs(self._prefs)

    def _theme_widget_recursive(self, widget, T):
        """Recursively apply theme to settings-panel children."""
        try:
            if isinstance(widget, tk.Label):
                widget.configure(bg=T["BG_SETTINGS"], fg=T["TEXT_PRI"])
            elif isinstance(widget, tk.Frame):
                widget.configure(bg=T["BG_SETTINGS"])
                for c in widget.winfo_children():
                    self._theme_widget_recursive(c, T)
            elif isinstance(widget, tk.Checkbutton):
                widget.configure(bg=T["BG_SETTINGS"], fg=T["TEXT_PRI"],
                                  selectcolor=T["BG_INPUT"],
                                  activebackground=T["BG_SETTINGS"],
                                  activeforeground=T["ACCENT"])
            elif isinstance(widget, tk.Radiobutton):
                widget.configure(bg=T["BG_SETTINGS"], fg=T["TEXT_PRI"],
                                  selectcolor=T["BG_INPUT"],
                                  activebackground=T["BG_SETTINGS"],
                                  activeforeground=T["ACCENT"])
            elif isinstance(widget, tk.Scale):
                widget.configure(bg=T["BG_SETTINGS"], fg=T["TEXT_PRI"],
                                  troughcolor=T["BG_INPUT"],
                                  activebackground=T["ACCENT"],
                                  highlightthickness=0)
            elif isinstance(widget, tk.Button):
                widget.configure(bg=T["BG_INPUT"], fg=T["TEXT_PRI"],
                                  activebackground=T["BORDER"],
                                  activeforeground=T["TEXT_PRI"],
                                  relief=tk.FLAT)
        except tk.TclError:
            pass

    def _reconfigure_chat_tags(self, T):
        self._chat.tag_configure("ts",         foreground=T["TEXT_SEC"],      font=FONT_TIME)
        self._chat.tag_configure("you_label",  foreground=T["TEXT_USER"],     font=("Segoe UI", 12, "bold"))
        self._chat.tag_configure("prova_label",foreground=T["TEXT_PROVA"],    font=("Segoe UI", 12, "bold"))
        self._chat.tag_configure("user_text",  foreground=T["TEXT_BUBBLE_U"], font=FONT_CHAT,
                                  background=T["BG_BUBBLE_U"],
                                  lmargin1=8, lmargin2=8, rmargin=8)
        self._chat.tag_configure("prova_text", foreground=T["TEXT_BUBBLE_P"], font=FONT_CHAT,
                                  background=T["BG_BUBBLE_P"],
                                  lmargin1=8, lmargin2=8, rmargin=8)
        self._chat.tag_configure("sys_text",   foreground=T["TEXT_SYS"],      font=("Segoe UI", 11, "italic"))
        self._chat.tag_configure("newline",    font=("Segoe UI", 4))

    # ─────────────────────────────────────────────────────────────
    # SETTINGS PANEL TOGGLE
    # ─────────────────────────────────────────────────────────────
    def _toggle_settings(self):
        if self._settings_open:
            self._settings_panel.pack_forget()
            self._settings_open = False
            self._settings_btn.configure(text="⚙  SETTINGS")
        else:
            # Pack settings between header sep and chat
            self._settings_panel.pack(
                in_=self.root, side=tk.TOP, fill=tk.X,
                before=self._chat_outer,
            )
            self._settings_open = True
            self._settings_btn.configure(text="✕  CLOSE")
            self._apply_theme(self._theme_name)  # re-skin freshly

    # ─────────────────────────────────────────────────────────────
    # VOICE / SETTINGS HELPERS
    # ─────────────────────────────────────────────────────────────
    def _apply_voice_settings(self):
        """Push speed/volume to Config so live changes take effect."""
        try:
            from voice_module import Config
            Config.TTS_RATE   = self._rate_var.get()
            Config.TTS_VOLUME = round(self._vol_var.get(), 2)
            self._prefs["tts_rate"] = Config.TTS_RATE
            self._prefs["tts_vol"]  = Config.TTS_VOLUME
            _save_prefs(self._prefs)
        except ImportError:
            pass

    def _apply_noise_setting(self):
        try:
            from voice_module import Config
            Config.NOISE_SUPPRESS = self._noise_suppress.get()
            self._prefs["noise_suppress"] = self._noise_suppress.get()
            _save_prefs(self._prefs)
        except ImportError:
            pass

    # ─────────────────────────────────────────────────────────────
    # CHAT
    # ─────────────────────────────────────────────────────────────
    def _append_msg(self, msg: _Msg):
        T = self._T
        self._chat.configure(state=tk.NORMAL)

        if msg.kind == "user":
            # Newline spacer
            self._chat.insert(tk.END, "\n", "newline")
            # Header line: icon + label + timestamp
            self._chat.insert(tk.END,
                f"  🧑 You   ", "you_label")
            self._chat.insert(tk.END,
                f"{msg.ts}\n", "ts")
            # Bubble text
            self._chat.insert(tk.END,
                f"  {msg.text}\n", "user_text")

        elif msg.kind == "prova":
            self._chat.insert(tk.END, "\n", "newline")
            self._chat.insert(tk.END,
                f"  ◆ ProVA  ", "prova_label")
            self._chat.insert(tk.END,
                f"{msg.ts}\n", "ts")
            self._chat.insert(tk.END,
                f"  {msg.text}\n", "prova_text")

        else:  # system
            self._chat.insert(tk.END,
                f"\n  ─  {msg.text}  ─  {msg.ts}\n", "sys_text")

        self._chat.configure(state=tk.DISABLED)
        self._chat.see(tk.END)

    # ─────────────────────────────────────────────────────────────
    # STATE BAR
    # ─────────────────────────────────────────────────────────────
    def _set_status(self, status: str):
        self._current_status = status
        T = self._T
        meta = STATE_META.get(status, STATE_META["IDLE"])
        icon, label, bar_key, dot_key, animate = meta
        hint = STATE_HINTS.get(status, "")

        bar_bg  = T.get(bar_key, T["BAR_IDLE"])
        dot_col = T.get(dot_key, T["STATE_IDLE"])

        # Expand bar
        self._state_bar.configure(bg=bar_bg, height=38)
        self._state_icon_lbl.configure(text=icon, bg=bar_bg, fg=dot_col)
        self._state_name_lbl.configure(text=f"  {label}", bg=bar_bg, fg=dot_col)
        self._state_dot_lbl.configure(bg=bar_bg, fg=dot_col, text="")
        self._state_hint_lbl.configure(text=hint, bg=bar_bg, fg=T["TEXT_SEC"])

        # Also update header bg subtly
        hdr_bg = bar_bg if status not in ("IDLE", "STOPPED", "PAUSED") else T["BG_DEEP"]
        for child in self._hdr.winfo_children():
            try: child.configure(bg=hdr_bg)
            except tk.TclError: pass
        self._hdr.configure(bg=hdr_bg)

        # Animated dots for active states
        if self._dot_anim_id:
            self.root.after_cancel(self._dot_anim_id)
            self._dot_anim_id = None

        if animate:
            self._dot_frame = 0
            self._tick_dots(dot_col, bar_bg)

    _DOT_FRAMES = ["", ".", "..", "..."]

    def _tick_dots(self, color: str, bg: str):
        self._dot_frame = (self._dot_frame + 1) % len(self._DOT_FRAMES)
        try:
            self._state_dot_lbl.configure(
                text=self._DOT_FRAMES[self._dot_frame],
                fg=color, bg=bg,
            )
        except tk.TclError:
            return
        self._dot_anim_id = self.root.after(400, lambda: self._tick_dots(color, bg))

    # ─────────────────────────────────────────────────────────────
    # INPUT PLACEHOLDER
    # ─────────────────────────────────────────────────────────────
    def _set_placeholder(self):
        self._input_var.set(self._PLACEHOLDER)
        self._input.configure(fg=self._T["TEXT_SEC"])
        self._input_has_placeholder = True

    def _on_input_change(self, *_):
        if self._input_has_placeholder:
            val = self._input_var.get()
            if val != self._PLACEHOLDER:
                if val.startswith(self._PLACEHOLDER):
                    real = val[len(self._PLACEHOLDER):]
                    self._input_has_placeholder = False
                    self._input.configure(fg=self._T["TEXT_PRI"])
                    self._input_var.set(real)
                    self._input.icursor(tk.END)
                elif val:
                    self._input_has_placeholder = False
                    self._input.configure(fg=self._T["TEXT_PRI"])
        # Notify listen_fn that the user is actively typing.
        # This resets the capture timeout so it doesn't expire mid-sentence.
        # Only sent when a module is waiting (text_queue is being monitored).
        if self._running and not self._input_has_placeholder:
            val = self._input_var.get()
            if val and val != self._PLACEHOLDER:
                self._text_queue.put("__typing__")

    def _clear_ph(self, _=None):
        if self._input_has_placeholder:
            self._input_var.set("")
            self._input.configure(fg=self._T["TEXT_PRI"])
            self._input_has_placeholder = False

    def _restore_ph(self, _=None):
        if not self._input_var.get().strip():
            self._set_placeholder()

    # ─────────────────────────────────────────────────────────────
    # UI QUEUE POLL
    # ─────────────────────────────────────────────────────────────
    def _poll_ui_queue(self):
        try:
            while True:
                kind, payload = self._ui_queue.get_nowait()
                if kind == "msg":
                    self._append_msg(payload)
                elif kind == "status":
                    self._set_status(payload)
        except queue.Empty:
            pass
        self.root.after(80, self._poll_ui_queue)

    # ─────────────────────────────────────────────────────────────
    # CALLBACKS (called from voice_module background threads)
    # ─────────────────────────────────────────────────────────────
    def _cb_user(self, text: str):
        self._ui_queue.put(("msg", _Msg("user", text)))

    def _cb_prova(self, text: str):
        self._ui_queue.put(("msg",    _Msg("prova", text)))
        self._ui_queue.put(("status", "SPEAKING"))

    def _cb_status(self, status: str):
        self._ui_queue.put(("status", status))
        if status == "STOPPED":
            self._ui_queue.put(("msg", _Msg("system", "ProVA stopped")))
            self.root.after(60, self._on_stopped)

    # ─────────────────────────────────────────────────────────────
    # CONTROLS
    # ─────────────────────────────────────────────────────────────
    def _on_stopped(self):
        T = self._T
        self._running = False
        self._paused  = False
        self._start_btn.configure(state=tk.NORMAL,   bg=T["ACCENT"],    fg=T["BG_DEEP"])
        self._pause_btn.configure(state=tk.DISABLED, text="⏸  PAUSE",   fg=T["TEXT_SEC"])
        self._stop_btn.configure (state=tk.DISABLED)

    def _start(self):
        if self._running:
            return
        self._stop_event.clear()
        self._pause_event.clear()
        self._muted[0] = False
        self._paused   = False

        from voice_module import run, ProVACallbacks, Config
        Config.WAKE_WORD_ENABLED = self._wake_var.get()
        Config.NOISE_SUPPRESS    = self._noise_suppress.get()
        # Push saved voice settings
        Config.TTS_RATE   = self._rate_var.get()
        Config.TTS_VOLUME = round(self._vol_var.get(), 2)

        cbs = ProVACallbacks(
            on_user_speech  = self._cb_user,
            on_prova_speech = self._cb_prova,
            on_status       = self._cb_status,
        )
        self._voice_thread = threading.Thread(
            target=run,
            kwargs=dict(
                callbacks   = cbs,
                stop_event  = self._stop_event,
                pause_event = self._pause_event,
                text_queue  = self._text_queue,
                muted_flag  = self._muted,
            ),
            daemon=True,
        )
        self._voice_thread.start()
        self._running = True

        T = self._T
        self._start_btn.configure(state=tk.DISABLED, bg=T["BG_INPUT"], fg=T["TEXT_SEC"])
        self._pause_btn.configure(state=tk.NORMAL)
        self._stop_btn.configure (state=tk.NORMAL)
        self._append_msg(_Msg("system", "ProVA starting…"))
        self.root.focus_set()   # ensure input box is not focused during voice operation

    def _toggle_pause(self):
        if not self._running:
            return
        if self._paused:
            self._pause_event.clear()
            self._paused = False
            self._pause_btn.configure(text="⏸  PAUSE")
            self._append_msg(_Msg("system", "Resumed"))
        else:
            self._pause_event.set()
            self._paused = True
            self._pause_btn.configure(text="▶  RESUME", fg=self._T["ACCENT"])
            self._append_msg(_Msg("system", "Paused"))

    def _stop(self):
        if self._running:
            self._stop_event.set()


    def _toggle_wake_word(self):
        try:
            from voice_module import Config
            Config.WAKE_WORD_ENABLED = self._wake_var.get()
        except ImportError:
            pass
        self._prefs["wake_word"] = self._wake_var.get()
        _save_prefs(self._prefs)
        hint = "Say 'hey ProVA' to activate" if self._wake_var.get() else "Always listening"
        state = "ON" if self._wake_var.get() else "OFF"
        self._append_msg(_Msg("system", f"Wake word {state} — {hint}"))

    def _send_typed(self):
        text = self._input_var.get().strip()
        if not text or text == self._PLACEHOLDER:
            return
        self._set_placeholder()
        if not self._running:
            self._append_msg(_Msg("system",
                "ProVA isn't running — click ▶ START first"))
            return
        # Do NOT append user message here — voice_module calls callbacks.on_user_speech(typed)
        # which fires _cb_user → _ui_queue → _append_msg. Adding it here too = double messages.
        self._text_queue.put(text)
        # Return focus to root so Windows voice dictation doesn't type
        # the next spoken command into the input box
        self.root.focus_set()

    def _send_help(self):
        if self._running:
            self._text_queue.put("help")
        else:
            try:
                from modules.help_module import MODULE_HELP
                self._append_msg(_Msg("system", "── Help ──"))
                for section, text in MODULE_HELP.items():
                    self._append_msg(_Msg("system", f"[{section.upper()}]  {text}"))
            except ImportError:
                self._append_msg(_Msg("system",
                    "Start ProVA first, then say 'help'"))

    def _on_close(self):
        self._stop_event.set()
        self.root.destroy()


# ─────────────────────────────────────────────────────────────────
def main():
    root = tk.Tk()
    app  = ProVAApp(root)
    root.protocol("WM_DELETE_WINDOW", app._on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
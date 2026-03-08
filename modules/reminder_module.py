"""
ProVA — modules/reminder_module.py
=====================================
Async reminder manager integrated with ProVA's voice pipeline.

Architecture:
  - AsyncReminderManager runs on its own event loop in a background thread
    (via start_manager_sync). ProVA's voice loop stays unblocked.
  - All public calls from voice_module go through thread-safe
    run_coroutine_threadsafe() wrappers exposed as plain sync functions.
  - handle(cmd, speak_fn) is the router entry point.

Fixes vs original:
  - asyncio.Lock lazy-initialized inside the loop (not in __init__)
  - asyncio.get_running_loop() replaces deprecated get_event_loop()
  - Notifications fire OUTSIDE the lock (prevents deadlock on slow toast)
  - XML special chars escaped in toast content
  - remind_in() checks `is not None` not truthiness (minutes=0 bug)
  - DEFAULT_STORE saved to <project_root>/data/reminders.json
    Falls back to ~/ProVA/ if running outside the project directory
  - Time parser handles "5 pm" / "5:30 pm" / "17:00" / "5" (hour only)
  - speak_fn wired in: ProVA speaks reminder aloud when it fires
  - list_reminders() returns spoken summary (count + next upcoming)
  - Result dataclass for all handler returns

Dependencies:
    pip install winsdk        (optional — falls back to PowerShell toast)

Usage:
    from modules.reminder_module import start_reminder_system, handle
    runner, manager = start_reminder_system(speak_fn)
    # then route voice commands to handle(cmd, speak_fn, manager)
"""

import asyncio
import base64
import json
import logging
import os
import uuid
from dataclasses import dataclass, asdict, field
from datetime import datetime, timedelta
from html import escape as xml_escape
from pathlib import Path
from typing import Optional, Callable, List
from zoneinfo import ZoneInfo

log = logging.getLogger("ProVA.Reminder")

# ─────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────
# Store reminders next to the project root — works regardless of where
# Python is launched from. Falls back to ~/ProVA/ if __file__ is unavailable.
def _resolve_store_path() -> Path:
    try:
        # __file__ is modules/reminder_module.py → go up one level to project root
        project_root = Path(__file__).resolve().parent.parent
        store_dir = project_root / "data"
    except NameError:
        store_dir = Path.home() / "ProVA"
    store_dir.mkdir(parents=True, exist_ok=True)
    return store_dir / "reminders.json"

DEFAULT_STORE = _resolve_store_path()
CHECK_INTERVAL     = 1.0          # seconds between due-check ticks
DEFAULT_TZ         = ZoneInfo("Asia/Kolkata")   # change to your timezone

# ─────────────────────────────────────────────────────────────────
# WINDOWS TOAST — three-layer system
#
# Layer 1 (best): winsdk native WinRT — requires `pip install winsdk`
# Layer 2 (good): PowerShell WinRT bridge — built into every Windows 10/11
# Layer 3 (safe): plyer cross-platform fallback — requires `pip install plyer`
#
# Root cause of silent failures:
#   Windows ToastNotificationManager requires a registered App User Model ID
#   (AUMID). An unregistered string like "ProVA" is silently dropped.
#   Fix: use the AUMID of a real registered Windows app as the notifier ID.
#   We use the Windows PowerShell AUMID which is always present on Win10/11:
#   {1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe
# ─────────────────────────────────────────────────────────────────

# Known-registered AUMID — Windows will always accept notifications from this
_TOAST_APP_ID = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe"

_HAS_WINSDK = False
try:
    from winsdk.windows.ui.notifications import (
        ToastNotificationManager, ToastNotification
    )
    from winsdk.windows.data.xml.dom import XmlDocument
    _HAS_WINSDK = True
    log.info("winsdk available — using native Toast API")
except Exception:
    log.info("winsdk not available — will use PowerShell toast")

_HAS_PLYER = False
try:
    from plyer import notification as _plyer_notification
    _HAS_PLYER = True
    log.info("plyer available — using as final fallback")
except Exception:
    pass


# ─────────────────────────────────────────────────────────────────
# RESULT — returned to voice layer for every operation
# ─────────────────────────────────────────────────────────────────
@dataclass
class Result:
    success: bool
    message: str          # What ProVA speaks


# ─────────────────────────────────────────────────────────────────
# REMINDER DATACLASS
# ─────────────────────────────────────────────────────────────────
@dataclass
class Reminder:
    id:             str
    title:          str
    message:        str
    when_iso:       str
    repeat_seconds: Optional[int] = None

    @property
    def when(self) -> datetime:
        return datetime.fromisoformat(self.when_iso)

    @when.setter
    def when(self, dt: datetime) -> None:
        self.when_iso = dt.isoformat()

    def to_dict(self) -> dict:
        return asdict(self)

    def spoken_time(self) -> str:
        """Human-friendly time string for TTS: '5:30 PM today' / 'tomorrow at 9:00 AM'"""
        now  = datetime.now(DEFAULT_TZ)
        when = self.when.astimezone(DEFAULT_TZ)
        time_str = when.strftime("%-I:%M %p").lstrip("0") if os.name != "nt" \
                   else when.strftime("%I:%M %p").lstrip("0")

        if when.date() == now.date():
            return f"today at {time_str}"
        elif when.date() == (now + timedelta(days=1)).date():
            return f"tomorrow at {time_str}"
        else:
            return when.strftime(f"%A, %B %-d at {time_str}")

    @classmethod
    def create(cls, title: str, message: str, when: datetime,
               repeat_seconds: Optional[int] = None) -> "Reminder":
        if when.tzinfo is None:
            raise ValueError("'when' must be timezone-aware.")
        return cls(
            id=str(uuid.uuid4()),
            title=title,
            message=message,
            when_iso=when.isoformat(),
            repeat_seconds=repeat_seconds,
        )

    @classmethod
    def from_dict(cls, d: dict) -> "Reminder":
        return cls(
            id=d["id"],
            title=d["title"],
            message=d["message"],
            when_iso=d["when_iso"],
            repeat_seconds=d.get("repeat_seconds"),
        )


# ─────────────────────────────────────────────────────────────────
# TIME PARSER
# Handles all formats voice_module might produce:
#   "17:00"  "5:30 pm"  "5 pm"  "9"  "9 AM"
# ─────────────────────────────────────────────────────────────────
def parse_time_string(time_str: str, tz: ZoneInfo = DEFAULT_TZ) -> datetime:
    """
    Parse a time string into a timezone-aware datetime (today or tomorrow).

    Supported formats:
      "17:00"     24-hour with minutes
      "5:30 pm"   12-hour with minutes + period
      "5 pm"      12-hour hour-only + period    ← was broken in original
      "9"         bare hour (assumed AM if < 12, else 24h)
      "9 AM"      hour + period, no minutes

    If the resolved time is in the past, schedules for tomorrow.
    """
    s = time_str.strip().lower()

    # Normalize period markers
    s = s.replace(".", "").replace("a m", "am").replace("p m", "pm")

    now = datetime.now(tz)
    hour = minute = None
    is_12h = False
    is_pm  = "pm" in s
    is_am  = "am" in s

    # Strip am/pm for parsing
    s_clean = s.replace("pm", "").replace("am", "").strip()

    if ":" in s_clean:
        # "17:00" or "5:30"
        parts  = s_clean.split(":")
        hour   = int(parts[0])
        minute = int(parts[1])
        is_12h = is_am or is_pm
    else:
        # Bare number like "5" or "17"
        try:
            hour   = int(s_clean)
            minute = 0
            is_12h = is_am or is_pm
        except ValueError:
            raise ValueError(
                f"Can't understand time '{time_str}'. "
                "Try something like '5 pm', '17:00', or '9:30 am'."
            )

    # Convert 12h → 24h
    if is_12h:
        if is_pm and hour != 12:
            hour += 12
        elif is_am and hour == 12:
            hour = 0

    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        raise ValueError(f"Invalid time: {time_str}")

    target = datetime(now.year, now.month, now.day, hour, minute, 0, tzinfo=tz)

    # If already passed today, schedule tomorrow
    if target <= now:
        target += timedelta(days=1)

    return target


# ─────────────────────────────────────────────────────────────────
# TOAST NOTIFICATIONS
# ─────────────────────────────────────────────────────────────────
def _build_toast_xml(title: str, message: str, is_alarm: bool = False) -> str:
    """
    Build Windows toast XML.

    scenario="alarm"   → sticky until dismissed, plays alarm sound,
                         bypasses Focus Assist / Do Not Disturb
    scenario="reminder"→ sticky until dismissed, softer chime, for repeating
    duration="long"    → stays on screen ~25 s (vs 7 s default)
    audio loop="false" → plays once (set True for a looping alarm)
    Dismiss button     → required on newer Win11 builds for alarm scenario
    """
    t = xml_escape(title)
    m = xml_escape(message)
    scenario  = "alarm"    if is_alarm else "reminder"
    audio_src = "ms-winsoundevent:Notification.Looping.Alarm2" if is_alarm \
                else "ms-winsoundevent:Notification.Reminder"
    return (
        f'<toast scenario="{scenario}" duration="long">'
        '<visual><binding template="ToastGeneric">'
        f'<text hint-maxLines="1">{t}</text>'
        f'<text>{m}</text>'
        '</binding></visual>'
        f'<audio src="{audio_src}" loop="false"/>'
        '<actions>'
        '<action activationType="system" arguments="dismiss" content="Dismiss"/>'
        '</actions>'
        '</toast>'
    )


async def _notify_native(reminder: Reminder) -> None:
    """
    Layer 1 — winsdk native WinRT.
    Uses the registered PowerShell AUMID so Windows actually shows the toast.
    Unregistered app IDs (e.g. bare "ProVA") are silently dropped by Windows.
    """
    is_alarm = reminder.repeat_seconds is None
    xml_str  = _build_toast_xml(reminder.title, reminder.message, is_alarm=is_alarm)
    xml = XmlDocument()
    xml.load_xml(xml_str)
    toast       = ToastNotification(xml)
    toast.tag   = reminder.id[:16]
    toast.group = "ProVA"
    # Key fix: use a registered AUMID — bare "ProVA" is not registered in Windows
    notifier = ToastNotificationManager.create_toast_notifier(_TOAST_APP_ID)
    notifier.show(toast)
    log.info("Native toast shown for '%s'", reminder.title)


async def _notify_powershell(reminder: Reminder) -> None:
    """
    Layer 2 — PowerShell WinRT bridge.
    Works on every Windows 10/11 machine without any pip installs.
    Also uses the registered PowerShell AUMID to guarantee visibility.
    """
    is_alarm   = reminder.repeat_seconds is None
    xml_str    = _build_toast_xml(reminder.title, reminder.message, is_alarm=is_alarm)
    xml_escaped = xml_str.replace("'", "''")
    app_id      = _TOAST_APP_ID.replace("\\", "\\\\")   # survive PS string escaping

    script = f"""
[void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime]
[void][Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType=WindowsRuntime]
[void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType=WindowsRuntime]

$appId = '{_TOAST_APP_ID}'
$xml   = [Windows.Data.Xml.Dom.XmlDocument]::new()
$xml.LoadXml('{xml_escaped}')
$toast         = [Windows.UI.Notifications.ToastNotification]::new($xml)
$toast.Tag     = '{reminder.id[:16]}'
$toast.Group   = 'ProVA'
$notifier      = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId)
$notifier.Show($toast)
"""
    encoded = base64.b64encode(script.encode("utf-16-le")).decode()
    proc = await asyncio.create_subprocess_exec(
        "powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded,
        stdout=asyncio.subprocess.DEVNULL,
        stderr=asyncio.subprocess.PIPE,
    )
    try:
        _, stderr = await asyncio.wait_for(proc.communicate(), timeout=15)
        if stderr:
            err = stderr.decode(errors="replace").strip()
            log.warning("PowerShell toast stderr for '%s': %s", reminder.title, err)
        else:
            log.info("PowerShell toast shown for '%s'", reminder.title)
    except asyncio.TimeoutError:
        proc.kill()
        log.warning("PowerShell toast timed out for '%s'", reminder.title)


def _notify_plyer(reminder: Reminder) -> None:
    """
    Layer 3 — plyer fallback (cross-platform, no WinRT required).
    pip install plyer
    Shows a simple system notification without alarm sound.
    """
    _plyer_notification.notify(
        title=reminder.title,
        message=reminder.message or reminder.title,
        app_name="ProVA",
        timeout=30,
    )
    log.info("plyer notification shown for '%s'", reminder.title)


async def _notify(reminder: Reminder) -> None:
    """
    Dispatch notification through layers until one succeeds.
    Layer 1: winsdk native (best — full alarm toast with sound)
    Layer 2: PowerShell WinRT (good — same toast XML, no extra installs)
    Layer 3: plyer (safe — simple popup, no sound, cross-platform)
    Always logs success/failure at each layer.
    """
    # Layer 1
    if _HAS_WINSDK:
        try:
            await _notify_native(reminder)
            return
        except Exception as e:
            log.warning("winsdk toast failed for '%s': %s — trying PowerShell", reminder.title, e)

    # Layer 2
    try:
        await _notify_powershell(reminder)
        return
    except Exception as e:
        log.warning("PowerShell toast failed for '%s': %s — trying plyer", reminder.title, e)

    # Layer 3
    if _HAS_PLYER:
        try:
            _notify_plyer(reminder)
            return
        except Exception as e:
            log.warning("plyer notification failed for '%s': %s", reminder.title, e)

    log.warning(
        "All toast layers failed for '%s'. "
        "Reminder was still spoken aloud via TTS. "
        "To fix: pip install plyer  or  pip install winsdk",
        reminder.title
    )


# ─────────────────────────────────────────────────────────────────
# REMINDER MANAGER
# ─────────────────────────────────────────────────────────────────
class ReminderManager:
    """
    Async reminder manager that runs on its own event loop thread.

    Key fixes vs original AsyncReminderManagerWindows:
      - Lock lazy-initialized inside the loop (asyncio.Lock must be
        created inside a running event loop — not in __init__)
      - start() uses get_running_loop() not deprecated get_event_loop()
      - Notifications fire OUTSIDE the lock so a slow/hung toast
        cannot freeze the entire check loop
      - speak_fn registered at construction so fired reminders
        are spoken aloud through ProVA's TTS
    """

    def __init__(
        self,
        speak_fn:       Optional[Callable[[str], None]] = None,
        store_path:     Path  = DEFAULT_STORE,
        tz:             ZoneInfo = DEFAULT_TZ,
        check_interval: float = CHECK_INTERVAL,
    ):
        self.speak_fn       = speak_fn or (lambda s: None)
        self.store_path     = Path(store_path)
        self.tz             = tz
        self.check_interval = check_interval

        self._reminders:  List[Reminder]              = []
        self._callbacks:  List[Callable[[Reminder], None]] = []
        self._task:       Optional[asyncio.Task]       = None
        self._running:    bool                         = False
        self._lock:       Optional[asyncio.Lock]       = None   # lazy init

        self._load()

    # ── Lazy lock ────────────────────────────────────────────────
    def _get_lock(self) -> asyncio.Lock:
        """
        Create the Lock on first use inside the running event loop.
        Fixes: asyncio.Lock() in __init__ raises DeprecationWarning
        in Python 3.10 and breaks entirely in 3.12+.
        """
        if self._lock is None:
            self._lock = asyncio.Lock()
        return self._lock

    # ── Persistence ───────────────────────────────────────────────
    def _load(self) -> None:
        self.store_path.parent.mkdir(parents=True, exist_ok=True)
        if self.store_path.exists():
            try:
                data = json.loads(self.store_path.read_text(encoding="utf-8"))
                self._reminders = [Reminder.from_dict(d) for d in data]
                log.info("Loaded %d reminder(s) from disk", len(self._reminders))
            except Exception:
                log.exception("Failed loading reminders — starting fresh")
                self._reminders = []
        self._sort()

    async def _save(self) -> None:
        tmp = self.store_path.with_suffix(".tmp")
        tmp.write_text(
            json.dumps([r.to_dict() for r in self._reminders],
                       indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        os.replace(tmp, self.store_path)

    def _sort(self) -> None:
        self._reminders.sort(key=lambda r: r.when)

    # ── Public API (async) ────────────────────────────────────────
    async def add_reminder(
        self,
        title:          str,
        message:        str,
        when:           datetime,
        repeat_seconds: Optional[int] = None,
    ) -> str:
        r = Reminder.create(title, message, when, repeat_seconds)
        async with self._get_lock():
            self._reminders.append(r)
            self._sort()
            await self._save()
        log.info("Added reminder '%s' for %s", r.title, r.spoken_time())
        return r.id

    async def delete_reminder(self, reminder_id: str) -> bool:
        async with self._get_lock():
            before = len(self._reminders)
            self._reminders = [r for r in self._reminders if r.id != reminder_id]
            removed = len(self._reminders) < before
            if removed:
                await self._save()
        return removed

    async def list_reminders(self) -> List[Reminder]:
        async with self._get_lock():
            return list(self._reminders)

    async def set_alarm(
        self, time_str: str, title: str = "Alarm", message: str = "Alarm ringing"
    ) -> str:
        when = parse_time_string(time_str, self.tz)
        return await self.add_reminder(title, message, when)

    async def set_daily_alarm(
        self, time_str: str, title: str = "Daily Alarm", message: str = "Daily reminder"
    ) -> str:
        when = parse_time_string(time_str, self.tz)
        return await self.add_reminder(title, message, when, repeat_seconds=86400)

    async def remind_in(
        self,
        minutes: Optional[int] = None,
        seconds: Optional[int] = None,
        title:   str = "Reminder",
        message: str = "Reminder triggered",
    ) -> str:
        """
        Fix: original checked `if minutes:` which treated minutes=0 as falsy.
        Now uses `is not None` so remind_in(minutes=0) works correctly.
        """
        now = datetime.now(self.tz)
        if minutes is not None:
            when = now + timedelta(minutes=minutes)
        elif seconds is not None:
            when = now + timedelta(seconds=seconds)
        else:
            raise ValueError("Provide minutes or seconds.")
        return await self.add_reminder(title, message, when)

    def register_callback(self, cb: Callable[[Reminder], None]) -> None:
        self._callbacks.append(cb)

    # ── Background loop ───────────────────────────────────────────
    async def _loop(self) -> None:
        """
        Check loop — runs every CHECK_INTERVAL seconds.

        Key fix: notifications fire OUTSIDE the lock.
        Original held the lock while awaiting _notify_powershell(),
        which could hang for 10+ seconds and freeze all reminder ops.

        Pattern now:
          1. Acquire lock → collect due reminders → update list → release lock
          2. Fire notifications and speak outside the lock
        """
        self._running = True
        while self._running:
            now = datetime.now(self.tz)
            due: List[Reminder] = []

            # ── Critical section: collect due, update list ────────
            async with self._get_lock():
                for r in self._reminders:
                    if r.when <= now:
                        due.append(r)

                for r in due:
                    if r.repeat_seconds:
                        r.when = r.when + timedelta(seconds=r.repeat_seconds)
                    else:
                        self._reminders.remove(r)

                if due:
                    self._sort()
                    await self._save()

            # ── Fire callbacks + notify OUTSIDE the lock ──────────
            for r in due:
                for cb in self._callbacks:
                    try:
                        cb(r)
                    except Exception as e:
                        log.warning("Callback error for reminder '%s': %s", r.title, e)

                # Speak through ProVA TTS
                try:
                    self.speak_fn(f"Reminder: {r.title}. {r.message}")
                except Exception as e:
                    log.warning("speak_fn error: %s", e)

                # Windows toast
                await _notify(r)

            await asyncio.sleep(self.check_interval)

    # ── Lifecycle ─────────────────────────────────────────────────
    def start(self) -> None:
        """
        Fix: original used asyncio.get_event_loop() which is deprecated
        in Python 3.10+ and raises RuntimeError in 3.12 if no loop is running.
        Must be called from within a running event loop.
        """
        if self._task and not self._task.done():
            return
        loop = asyncio.get_running_loop()
        self._task = loop.create_task(self._loop())
        log.info("Reminder loop started.")

    async def stop(self) -> None:
        self._running = False
        if self._task:
            self._task.cancel()
            try:
                await self._task
            except asyncio.CancelledError:
                pass

    async def shutdown(self) -> None:
        await self.stop()
        async with self._get_lock():
            await self._save()
        log.info("Reminder manager shut down.")


# ─────────────────────────────────────────────────────────────────
# SYNC BRIDGE
# Runs the async manager on its own thread so ProVA's voice loop
# (which uses threading, not asyncio) can call it safely.
# ─────────────────────────────────────────────────────────────────
class _SyncRunner:
    """
    Wraps ReminderManager with sync-callable methods using
    run_coroutine_threadsafe(). Safe to call from any thread.
    """

    def __init__(self, manager: ReminderManager, loop: asyncio.AbstractEventLoop):
        self._manager = manager
        self._loop    = loop

    def _run(self, coro):
        """Submit a coroutine to the manager's event loop and wait for result."""
        future = asyncio.run_coroutine_threadsafe(coro, self._loop)
        return future.result(timeout=10)

    def add_reminder(self, title, message, when, repeat_seconds=None) -> str:
        return self._run(
            self._manager.add_reminder(title, message, when, repeat_seconds)
        )

    def delete_reminder(self, reminder_id: str) -> bool:
        return self._run(self._manager.delete_reminder(reminder_id))

    def list_reminders(self) -> List[Reminder]:
        return self._run(self._manager.list_reminders())

    def set_alarm(self, time_str, title="Alarm", message="Alarm ringing") -> str:
        return self._run(self._manager.set_alarm(time_str, title, message))

    def set_daily_alarm(self, time_str, title="Daily Alarm", message="Daily reminder") -> str:
        return self._run(self._manager.set_daily_alarm(time_str, title, message))

    def remind_in(self, minutes=None, seconds=None, title="Reminder", message="Reminder") -> str:
        return self._run(
            self._manager.remind_in(minutes=minutes, seconds=seconds,
                                     title=title, message=message)
        )

    def stop(self) -> None:
        future = asyncio.run_coroutine_threadsafe(
            self._manager.shutdown(), self._loop
        )
        try:
            future.result(timeout=5)
        except Exception:
            pass
        self._loop.call_soon_threadsafe(self._loop.stop)


def start_reminder_system(
    speak_fn: Callable[[str], None],
    store_path: Path = DEFAULT_STORE,
    tz: ZoneInfo = DEFAULT_TZ,
) -> tuple["_SyncRunner", "ReminderManager"]:
    """
    Start the reminder system in a background thread.
    Returns (_SyncRunner, ReminderManager) — use _SyncRunner from voice_module.

    Call once at ProVA startup:
        runner, manager = start_reminder_system(speak_fn)

    Then pass `runner` to handle():
        handle(cmd, speak_fn, runner)
    """
    import threading

    manager = ReminderManager(speak_fn=speak_fn, store_path=store_path, tz=tz)
    loop    = asyncio.new_event_loop()
    runner  = _SyncRunner(manager, loop)

    def _thread_main():
        asyncio.set_event_loop(loop)

        async def _start_and_run():
            manager.start()
            # Keep loop alive indefinitely
            while True:
                await asyncio.sleep(3600)

        try:
            loop.run_until_complete(_start_and_run())
        except asyncio.CancelledError:
            pass
        finally:
            pending = asyncio.all_tasks(loop)
            for t in pending:
                t.cancel()
            loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
            loop.run_until_complete(loop.shutdown_asyncgens())
            loop.close()

    t = threading.Thread(target=_thread_main, daemon=True, name="ProVA-ReminderLoop")
    t.start()
    log.info("Reminder system started on background thread.")
    return runner, manager


# ─────────────────────────────────────────────────────────────────
# SPOKEN SUMMARY HELPERS
# ─────────────────────────────────────────────────────────────────
def _spoken_list_summary(reminders: List[Reminder]) -> str:
    """
    'Just count + next upcoming' format as chosen.
    Examples:
      "You have 3 reminders. Next up: Meeting, today at 5:00 PM."
      "You have no upcoming reminders."
    """
    if not reminders:
        return "You have no upcoming reminders."

    count = len(reminders)
    next_r = reminders[0]   # already sorted by when

    count_str = f"You have {count} reminder{'s' if count != 1 else ''}."
    next_str  = f"Next up: {next_r.title}, {next_r.spoken_time()}."

    return f"{count_str} {next_str}"


def _spoken_confirmation(reminder: Reminder, repeat: bool = False) -> str:
    """What ProVA says right after setting a reminder."""
    base = f"Reminder set. I'll remind you about {reminder.title} {reminder.spoken_time()}."
    if repeat:
        base += " This will repeat daily."
    return base


# ─────────────────────────────────────────────────────────────────
# MAIN HANDLER — called by voice_module router
# ─────────────────────────────────────────────────────────────────
def handle(cmd, speak_fn: Callable[[str], None], runner: "_SyncRunner") -> None:
    """
    Entry point called by voice_module.route() via run_async.

    Args:
        cmd:      Command dataclass from voice_module
        speak_fn: lambda text: speak(engine, text)
        runner:   _SyncRunner instance from start_reminder_system()

    cmd.action values:
        "set"       → set a one-time reminder
        "alarm"     → set a one-time alarm (same, different phrasing)
        "daily"     → set a daily repeating alarm
        "remind_in" → remind in X minutes
        "list"      → list upcoming reminders
        "delete"    → cancel a reminder (future: by title match)
    """
    action   = cmd.action   or "set"
    time_str = cmd.time_str or ""
    # Use the extracted subject ("call John") as title when no explicit target given.
    # Fall back to "ProVA Reminder" only when message is empty or single-word noise.
    _msg_as_title = (
        cmd.message
        if cmd.message and len(cmd.message.split()) > 1
        else None
    )
    title   = cmd.target or _msg_as_title or "ProVA Reminder"
    message = cmd.message or cmd.target or title

    try:
        # ── set / alarm ───────────────────────────────────────────
        if action in ("set", "alarm"):
            if not time_str:
                speak_fn("What time should I set the reminder for?")
                return

            rid  = runner.set_alarm(time_str, title=title, message=message)
            rems = runner.list_reminders()
            just_set = next((r for r in rems if r.id == rid), None)

            if just_set:
                speak_fn(_spoken_confirmation(just_set, repeat=False))
            else:
                speak_fn(f"Reminder set for {time_str}.")

        # ── daily ─────────────────────────────────────────────────
        elif action == "daily":
            if not time_str:
                speak_fn("What time should the daily reminder be?")
                return

            rid      = runner.set_daily_alarm(time_str, title=title, message=message)
            rems     = runner.list_reminders()
            just_set = next((r for r in rems if r.id == rid), None)

            if just_set:
                speak_fn(_spoken_confirmation(just_set, repeat=True))
            else:
                speak_fn(f"Daily reminder set for {time_str}.")

        # ── remind_in ─────────────────────────────────────────────
        elif action == "remind_in":
            # cmd.extra may carry {"minutes": 10} or {"seconds": 30}
            minutes = cmd.extra.get("minutes")
            seconds = cmd.extra.get("seconds")

            if minutes is None and seconds is None:
                # Try to parse from message: "remind me in 10 minutes"
                import re
                m = re.search(r"(\d+)\s*(minute|min|second|sec)", message, re.IGNORECASE)
                if m:
                    val  = int(m.group(1))
                    unit = m.group(2).lower()
                    if "min" in unit:
                        minutes = val
                    else:
                        seconds = val

            if minutes is None and seconds is None:
                speak_fn("How many minutes should I wait before reminding you?")
                return

            runner.remind_in(minutes=minutes, seconds=seconds,
                             title=title, message=message)

            if minutes is not None:
                speak_fn(f"Got it. I'll remind you in {minutes} minute{'s' if minutes != 1 else ''}.")
            else:
                speak_fn(f"Got it. I'll remind you in {seconds} second{'s' if seconds != 1 else ''}.")

        # ── list ──────────────────────────────────────────────────
        elif action == "list":
            reminders = runner.list_reminders()
            speak_fn(_spoken_list_summary(reminders))

        # ── delete / cancel ───────────────────────────────────────
        elif action in ("delete", "cancel"):
            reminders = runner.list_reminders()
            if not reminders:
                speak_fn("You have no reminders to cancel.")
                return

            # Match by title (case-insensitive substring)
            target_title = (cmd.target or "").lower()
            match = next(
                (r for r in reminders if target_title in r.title.lower()),
                None
            )

            if match is None:
                # Fallback: cancel the next upcoming one
                match = reminders[0]
                speak_fn(f"Cancelling your next reminder: {match.title}, {match.spoken_time()}.")
            else:
                speak_fn(f"Cancelling reminder: {match.title}.")

            runner.delete_reminder(match.id)

        else:
            speak_fn(f"I'm not sure how to handle that reminder command.")

    except ValueError as e:
        speak_fn(str(e))
    except Exception as e:
        log.exception("Reminder handler error: %s", e)
        speak_fn("Something went wrong with the reminder. Please try again.")
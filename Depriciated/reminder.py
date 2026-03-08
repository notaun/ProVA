import asyncio
import json
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import os
import uuid
import logging
import base64
from typing import Optional, Callable, List

# ---- Config ----
DEFAULT_STORE = "reminders_windows.json"
CHECK_INTERVAL_SECONDS = 1.0

try:
    from zoneinfo import ZoneInfo
    DEFAULT_TZ = ZoneInfo("UTC")
except Exception:
    # Fallback if tzdata not available
    from datetime import timezone
    DEFAULT_TZ = timezone.utc


# ---- Logging ----
logger = logging.getLogger("reminder_windows")
logger.setLevel(logging.INFO)
if not logger.handlers:
    h = logging.StreamHandler()
    h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(h)

# ---- Prefer the modern winsdk package (Windows Runtime) ----
_HAS_WINSKD = False
try:
    # these imports may raise if winsdk package is not installed
    from winsdk.windows.ui.notifications import ToastNotificationManager, ToastNotification
    from winsdk.windows.data.xml.dom import XmlDocument
    _HAS_WINSKD = True
    logger.info("winsdk available: using native Windows Toast API")
except Exception:
    _HAS_WINSKD = False
    logger.info("winsdk not available: will use PowerShell fallback")


@dataclass
class Reminder:
    id: str
    title: str
    message: str
    when_iso: str  # timezone-aware ISO datetime
    repeat_seconds: Optional[int] = None

    @property
    def when(self) -> datetime:
        return datetime.fromisoformat(self.when_iso)

    @when.setter
    def when(self, dt: datetime):
        self.when_iso = dt.isoformat()

    def to_json(self):
        return asdict(self)

    @classmethod
    def create(cls, title: str, message: str, when: datetime, repeat_seconds: Optional[int] = None):
        if when.tzinfo is None:
            raise ValueError("when must be timezone-aware")
        return cls(id=str(uuid.uuid4()), title=title, message=message, when_iso=when.isoformat(),
                   repeat_seconds=repeat_seconds)

    @classmethod
    def from_dict(cls, d):
        return cls(id=d["id"], title=d["title"], message=d["message"], when_iso=d["when_iso"],
                   repeat_seconds=d.get("repeat_seconds"))


class AsyncReminderManagerWindows:
    """
    Async reminder manager for Windows.
    Use start() within an asyncio event loop to run background checking.
    Register callbacks via register_on_fire(cb) to receive fired reminders in your VA.
    """

    def __init__(self, store_path: str = DEFAULT_STORE, tz: ZoneInfo = DEFAULT_TZ,
                 check_interval: float = CHECK_INTERVAL_SECONDS):
        self.store_path = store_path
        self.tz = tz
        self.check_interval = check_interval
        self._reminders: List[Reminder] = []
        self._task: Optional[asyncio.Task] = None
        self._running = False
        self._lock = asyncio.Lock()
        self._on_fire_callbacks: List[Callable[[Reminder], None]] = []
        self._load()

    # Persistence
    def _load(self):
        if os.path.exists(self.store_path):
            try:
                with open(self.store_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._reminders = [Reminder.from_dict(d) for d in data]
                self._sort()
            except Exception:
                logger.exception("Failed to load reminders; starting empty.")
                self._reminders = []
        else:
            self._reminders = []

    async def _save(self):
        try:
            tmp = self.store_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump([r.to_json() for r in self._reminders], f, indent=2, ensure_ascii=False)
            os.replace(tmp, self.store_path)
        except Exception:
            logger.exception("Failed to save reminders")

    def _sort(self):
        self._reminders.sort(key=lambda r: r.when)

    # Public API
    async def add_reminder(self, title: str, message: str, when: datetime, repeat_seconds: Optional[int] = None) -> str:
        if when.tzinfo is None:
            raise ValueError("when must be timezone-aware")
        r = Reminder.create(title=title, message=message, when=when, repeat_seconds=repeat_seconds)
        async with self._lock:
            self._reminders.append(r)
            self._sort()
            await self._save()
        logger.info("Added reminder %s for %s", r.id, r.when_iso)
        return r.id

    async def delete_reminder(self, reminder_id: str) -> bool:
        async with self._lock:
            before = len(self._reminders)
            self._reminders = [r for r in self._reminders if r.id != reminder_id]
            changed = len(self._reminders) != before
            if changed:
                await self._save()

        if changed:
            logger.info("Deleted reminder %s", reminder_id)

        return changed

    async def list_reminders(self) -> List[Reminder]:
        async with self._lock:
            return list(self._reminders)

    def register_on_fire(self, cb: Callable[[Reminder], None]):
        """Register a sync callback. VA can wrap coroutine if desired."""
        self._on_fire_callbacks.append(cb)

    # Notification backends (Windows)
    async def _notify_native_winsdk(self, reminder: Reminder) -> bool:
        """Use winsdk native Toast API (requires winsdk package)."""
        try:
            # Build simple Toast XML
            toast_xml = f"""
            <toast>
              <visual>
                <binding template='ToastGeneric'>
                  <text><![CDATA[{reminder.title}]]></text>
                  <text><![CDATA[{reminder.message}]]></text>
                </binding>
              </visual>
            </toast>
            """
            xml_doc = XmlDocument()
            xml_doc.load_xml(toast_xml)
            toast = ToastNotification(xml_doc)
            notifier = ToastNotificationManager.create_toast_notifier("Python-Reminder")
            notifier.show(toast)
            logger.debug("winsdk toast shown for %s", reminder.id)
            return True
        except Exception:
            logger.exception("winsdk notify failed")
            return False

    async def _notify_powershell(self, reminder: Reminder) -> bool:
        """
        Safe PowerShell fallback using -EncodedCommand (UTF16LE base64) to avoid injection.
        The PowerShell script uses the Windows Runtime Toast API.
        """
        # Build PS script (no user interpolation into command string; embed as here-doc)
        ps_script = f'''
Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
$xml = @"
<toast>
  <visual>
    <binding template='ToastGeneric'>
      <text>{_ps_escape_xml(reminder.title)}</text>
      <text>{_ps_escape_xml(reminder.message)}</text>
    </binding>
  </visual>
</toast>
"@
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
$xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
$xmlDoc.LoadXml($xml)
$toast = [Windows.UI.Notifications.ToastNotification]::new($xmlDoc)
$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Python-Reminder")
$notifier.Show($toast)
'''
        # Encode UTF-16LE and base64 for -EncodedCommand
        ps_bytes = ps_script.encode('utf-16-le')
        b64 = base64.b64encode(ps_bytes).decode('ascii')
        # Launch PowerShell non-blocking (async)
        proc = await asyncio.create_subprocess_exec("powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", b64,
                                                    stdout=asyncio.subprocess.DEVNULL,
                                                    stderr=asyncio.subprocess.DEVNULL)
        await proc.wait()
        logger.debug("PowerShell toast launched for %s", reminder.id)
        return True

    async def _notify_system(self, reminder: Reminder) -> bool:
        """Top-level notify: try winsdk then PowerShell fallback."""
        if _HAS_WINSKD:
            ok = await self._notify_native_winsdk(reminder)
            if ok:
                return True
            else:
                logger.info("winsdk failed; trying PowerShell fallback")
        # fallback
        return await self._notify_powershell(reminder)

    # Background loop
    async def _loop(self):
        logger.info("Reminder loop starting")
        self._running = True
        try:
            while self._running:
                now = datetime.now(tz=self.tz)
                due = []
                async with self._lock:
                    for r in list(self._reminders):
                        if r.when <= now:
                            due.append(r)
                    for r in due:
                        # Trigger callbacks immediately (sync) so VA can respond quickly
                        for cb in self._on_fire_callbacks:
                            try:
                                cb(r)
                            except Exception:
                                logger.exception("on_fire callback failed")
                        # Fire system notification (async)
                        try:
                            await self._notify_system(r)
                        except Exception:
                            logger.exception("System notify failed for %s", r.id)

                        # Reschedule or remove
                        if r.repeat_seconds:
                            # advance by repeat_seconds until in future (catch-up)
                            next_when = r.when
                            while next_when <= now:
                                next_when = next_when + timedelta(seconds=r.repeat_seconds)
                            r.when = next_when
                        else:
                            self._reminders = [x for x in self._reminders if x.id != r.id]

                    if due:
                        self._sort()
                        await self._save()
                # interruptible sleep
                await asyncio.sleep(self.check_interval)
        except asyncio.CancelledError:
            logger.info("Reminder loop cancelled")
        finally:
            self._running = False
            logger.info("Reminder loop stopped")

    # Control methods
    def start(self):
        """Start the background task. Must be called from the running event loop."""
        if self._task and not self._task.done():
            return
        loop = asyncio.get_event_loop()
        self._task = loop.create_task(self._loop())

    async def stop(self):
        """Stop the loop and wait for it to finish."""
        self._running = False
        if self._task:
            self._task.cancel()
            try:
                await self._task
            except asyncio.CancelledError:
                pass
            self._task = None

    async def shutdown(self):
        await self.stop()
        async with self._lock:
            await self._save()

    # Direct notify helper (immediately notify, not scheduling)
    async def notify_now(self, title: str, message: str):
        r = Reminder.create(title=title, message=message, when=datetime.now(tz=self.tz))
        await self._notify_system(r)


# ---- Helper: escape XML for PowerShell embedded xml ---
def _ps_escape_xml(s: str) -> str:
    """Escape text for safe insertion into XML used in the PowerShell script."""
    # minimal escaping for &,<,>,',"
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace("'", "&apos;")
             .replace('"', "&quot;"))


# ---- Sync helper for non-async entrypoints ----
def start_manager_sync(manager: AsyncReminderManagerWindows):
    """
    Convenience to start the manager from synchronous code:
    - creates a background asyncio event loop in a dedicated thread.
    - returns a 'runner' object with stop() method to shutdown.
    """
    import threading

    loop = asyncio.new_event_loop()

    def run_loop():
        asyncio.set_event_loop(loop)
        manager.start()
        try:
            loop.run_forever()
        finally:
            # ensure proper cleanup
            pending = asyncio.all_tasks(loop=loop)
            for t in pending:
                t.cancel()
            loop.run_until_complete(loop.shutdown_asyncgens())
            loop.close()

    t = threading.Thread(target=run_loop, daemon=True)
    t.start()

    class Runner:
        def stop(self):
            # schedule shutdown coroutine and stop loop
            async def shutdown_and_stop():
                await manager.shutdown()
                loop.stop()
            asyncio.run_coroutine_threadsafe(shutdown_and_stop(), loop).result(timeout=5)
            t.join(timeout=2)

    return Runner()


# ---- Initiation example (usage) ----
if __name__ == "__main__":
    # Simple demo: start manager, add a reminder 10 seconds from now, run for 20s.
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    mgr = AsyncReminderManagerWindows()

    async def demo():
        mgr.register_on_fire(lambda r: logger.info("Callback: fired %s - %s", r.id, r.title))
        mgr.start()
        tz = DEFAULT_TZ
        when = datetime.now(tz=tz) + timedelta(seconds=10)
        await mgr.add_reminder("Demo Reminder", "This fired from reminder_windows demo", when)
        # let it run for 25 seconds, then shutdown
        await asyncio.sleep(25)
        await mgr.shutdown()

    try:
        loop.run_until_complete(demo())
    finally:
        loop.close()

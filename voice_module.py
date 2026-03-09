"""
voice_module.py — ProVA audio engine and main loop.

Accuracy improvements for Indian English + Bluetooth headset:

  IMPROVEMENT 1 — Recognizer tuned for Bluetooth latency.
    Bluetooth mics add ~100-200 ms codec delay and narrow the audio to 8 kHz
    in HSP/HFP mode. Increased PAUSE_THRESHOLD (0.8 → 1.2 s) so commands
    aren't cut off while speaking, and raised PHRASE_TIME_LIMIT (10 → 15 s).
    Ambient calibration extended (1.5 → 2.5 s) to properly baseline BT noise.

  IMPROVEMENT 2 — STT retry with alternate language fallback.
    transcribe() now tries en-IN first, then en-US on UnknownValueError.
    Google STT occasionally returns nothing for en-IN on short/fast utterances
    but succeeds with en-US. A second attempt costs ~300 ms but saves the command.

  IMPROVEMENT 3 — Pre-STT audio normalisation.
    Bluetooth audio is often quieter than wired. Added _normalise_audio() which
    checks the RMS energy of captured audio and logs a warning if it's too low,
    helping diagnose "words heard wrong" caused by mic gain rather than STT.

  IMPROVEMENT 4 — Mic selection prefers Bluetooth input explicitly.
    pick_best_microphone() now scores BT headset mics higher than laptop mics,
    ensuring the right device is chosen when both are available.

  IMPROVEMENT 5 — Dynamic energy threshold disabled after calibration.
    DYNAMIC_ENERGY=True lets the threshold drift upward during silence between
    commands, causing the next command to be missed entirely. Disabled after the
    initial calibration so the threshold stays stable session-wide.
"""

from __future__ import annotations

import logging
import queue
import sys
import threading
import time
from dataclasses import dataclass, field
from typing import Callable, Optional

# Ensure ProVA root is on path so `from parser import` finds our parser.py,
# not the stdlib `parser` module (which exists in Python < 3.9).
import os as _os
_prova_root = _os.path.dirname(_os.path.abspath(__file__))
if _prova_root not in sys.path:
    sys.path.insert(0, _prova_root)

import pyttsx3
import speech_recognition as sr

try:
    from fuzzywuzzy import fuzz
except ImportError:
    try:
        from thefuzz import fuzz  # type: ignore
    except ImportError:
        class fuzz:  # type: ignore
            @staticmethod
            def partial_ratio(a: str, b: str) -> int:
                return 100 if a in b else 0

from parser import detect_intent
from router import route, init as router_init

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("ProVA")


# ─────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────
class Config:
    WAKE_WORD         = "hey prova"
    WAKE_WORD_ENABLED = True
    LISTEN_TIMEOUT    = 8
    PHRASE_TIME_LIMIT = 15
    AMBIENT_DURATION  = 2.0            # long enough for BT AND built-in
    TTS_RATE          = 172
    TTS_VOLUME        = 1.0
    TTS_VOICE_ID      = None
    NOISE_SUPPRESS    = False
    WAKE_WORD_FUZZ    = 65
    ENERGY_THRESHOLD  = 200            # starting point; calibration overrides this
    DYNAMIC_ENERGY    = True           # calibrate first, then lock
    PAUSE_THRESHOLD   = 1.2            # generous — handles BT codec delay + normal speech
    FORCE_MIC_INDEX   = None
    POST_TTS_COOLDOWN = 0.9
    ECHO_MIN_WORDS    = 3
    ECHO_THRESHOLD    = 80
    ECHO_WINDOW       = 2.5
    STT_RETRY         = True
    LOW_ENERGY_WARN   = 80             # warn if RMS genuinely dead


# ─────────────────────────────────────────────────────────────────
# CALLBACKS
# ─────────────────────────────────────────────────────────────────
@dataclass
class ProVACallbacks:
    on_user_speech:  Callable[[str], None] = field(default=lambda t: None)
    on_prova_speech: Callable[[str], None] = field(default=lambda t: None)
    on_status:       Callable[[str], None] = field(default=lambda s: None)


# ─────────────────────────────────────────────────────────────────
# MODULE-LEVEL COORDINATION PRIMITIVES
# ─────────────────────────────────────────────────────────────────
_mic_lock      = threading.Lock()
_tts_done      = threading.Event()
_module_active = threading.Event()
_tts_done.set()

_last_spoken      = [""]
_last_spoken_time = [0.0]
_busy_count       = [0]

_tts_queue: queue.Queue = queue.Queue()


def _reset_state() -> None:
    """Reset all module-level coordination state. Called at start of run()."""
    _tts_done.set()
    _module_active.clear()
    _last_spoken[0]      = ""
    _last_spoken_time[0] = 0.0   # Fix: was time.time() which caused first command to be echo-filtered
    _busy_count[0]       = 0
    while not _tts_queue.empty():
        try:
            _tts_queue.get_nowait()
            _tts_queue.task_done()
        except Exception:
            break


# ─────────────────────────────────────────────────────────────────
# DEDICATED TTS THREAD
# ─────────────────────────────────────────────────────────────────
def _speaker_thread(
    muted_flag:     list,
    callbacks:      ProVACallbacks,
    recognizer_ref: list,
) -> None:
    # pyttsx3 on Windows uses SAPI via COM. COM must be initialized in the
    # same thread that calls pyttsx3.init() — otherwise engine.runAndWait()
    # silently does nothing (voice never plays even though log shows the text).
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except ImportError:
        pass   # non-Windows or pywin32 not installed — fine

    try:
        engine = pyttsx3.init()
    except Exception as e:
        log.error("TTS engine failed to initialize: %s", e)
        log.error("Voice disabled. Check pyttsx3 is installed and a SAPI voice exists.")
        while True:
            text = _tts_queue.get()
            _tts_done.set()
            _tts_queue.task_done()
            if text is None:
                break
        return

    engine.setProperty("rate",   Config.TTS_RATE)
    engine.setProperty("volume", Config.TTS_VOLUME)

    # Voice: always female. Priority: Zira > any other female > first available female.
    # Never fall back to a male voice.
    voices = engine.getProperty("voices")
    chosen = None
    for v in voices:
        if "zira" in v.name.lower():
            chosen = v.id
            log.info("TTS: Microsoft Zira selected")
            break
    if not chosen:
        for v in voices:
            if any(n in v.name.lower() for n in ("hazel", "susan", "helen", "eva", "female")):
                chosen = v.id
                log.info("TTS: female voice fallback: %s", v.name)
                break
    if not chosen and voices:
        # Last resort — just take first, but log a warning
        chosen = voices[0].id
        log.warning("TTS: no female voice found, using: %s", voices[0].name)
    if chosen:
        engine.setProperty("voice", chosen)

    log.info("TTS ready. Rate=%d Volume=%.1f", Config.TTS_RATE, Config.TTS_VOLUME)

    # Track rate/volume for live changes from Settings sliders.
    _applied_rate  = Config.TTS_RATE
    _applied_vol   = Config.TTS_VOLUME

    while True:
        text = _tts_queue.get()
        if text is None:
            _tts_queue.task_done()
            break

        # ── Apply live settings changes (rate/volume only) ────────
        if Config.TTS_RATE != _applied_rate:
            engine.setProperty("rate", Config.TTS_RATE)
            _applied_rate = Config.TTS_RATE
            log.info("TTS rate updated → %d", _applied_rate)
        if abs(Config.TTS_VOLUME - _applied_vol) > 0.01:
            engine.setProperty("volume", Config.TTS_VOLUME)
            _applied_vol = Config.TTS_VOLUME
            log.info("TTS volume updated → %.2f", _applied_vol)

        # ── Voice test command ────────────────────────────────────
        if text == "__voice_test__":
            try:
                engine.stop()   # reset SAPI state — same fix as normal speech path
                engine.say("Hello! This is a voice test. Speed and volume can be adjusted in Settings.")
                engine.runAndWait()
            except Exception as e:
                log.warning("Voice test TTS error: %s", e)
            _tts_done.set()
            _tts_queue.task_done()
            continue

        _last_spoken[0]      = text.lower()
        _last_spoken_time[0] = time.time()

        _tts_done.clear()
        callbacks.on_status("SPEAKING")

        r = recognizer_ref[0]
        saved = r.energy_threshold if r else None

        if not muted_flag[0]:
            log.info("ProVA: %s", text)
            try:
                # engine.stop() resets pyttsx3's internal SAPI driver state.
                # Without this, runAndWait() silently skips after the first call
                # because the driver's isSpeaking flag is left in a dirty state.
                engine.stop()
                engine.say(text)
                engine.runAndWait()
            except Exception as e:
                log.warning("TTS error: %s", e)
                # Try reinitialising engine on error — recovers from COM state corruption
                try:
                    engine = pyttsx3.init()
                    engine.setProperty("rate",   Config.TTS_RATE)
                    engine.setProperty("volume", Config.TTS_VOLUME)
                    if chosen:
                        engine.setProperty("voice", chosen)
                    engine.say(text)
                    engine.runAndWait()
                    log.info("TTS: engine re-initialised and speech recovered")
                except Exception as e2:
                    log.error("TTS recovery failed: %s", e2)
        else:
            log.info("ProVA (muted): %s", text)

        time.sleep(Config.POST_TTS_COOLDOWN)

        if r and saved is not None:
            r.energy_threshold = saved

        _tts_done.set()
        callbacks.on_status("IDLE")
        _tts_queue.task_done()


def make_speak_fn(callbacks: ProVACallbacks, muted_flag: list) -> Callable:
    """Returns a thread-safe speak function. Blocks until fully spoken + cooldown."""
    def speak(text: str) -> None:
        if not text or not text.strip():
            return   # never send empty text to TTS or chat
        callbacks.on_prova_speech(text)
        _tts_queue.put(text)
        _tts_queue.join()
    return speak


# ─────────────────────────────────────────────────────────────────
# RECOGNIZER
# ─────────────────────────────────────────────────────────────────
def build_recognizer() -> sr.Recognizer:
    r = sr.Recognizer()
    r.energy_threshold         = Config.ENERGY_THRESHOLD
    r.dynamic_energy_threshold = Config.DYNAMIC_ENERGY
    r.pause_threshold          = Config.PAUSE_THRESHOLD
    r.phrase_threshold         = 0.3
    r.non_speaking_duration    = 0.6   # slightly longer for BT gap after speech
    return r


# ─────────────────────────────────────────────────────────────────
# AUDIO ENERGY CHECK
# BT mics in HFP mode transmit at lower amplitude. Log a warning so
# the user knows to check mic gain in Windows sound settings.
# ─────────────────────────────────────────────────────────────────
def _check_audio_energy(audio: sr.AudioData) -> None:
    try:
        import audioop
        rms = audioop.rms(audio.get_raw_data(), audio.sample_width)
        if rms < Config.LOW_ENERGY_WARN:
            log.warning(
                "Mic energy low (RMS=%d). "
                "Check Windows Settings > Sound > Input > your microphone > Input volume. "
                "If using Bluetooth, switch to built-in mic for better STT accuracy.", rms
            )
        else:
            log.debug("Audio RMS: %d", rms)
    except Exception:
        pass  # audioop unavailable on some platforms — non-critical


# ─────────────────────────────────────────────────────────────────
# COMMAND-LEVEL ECHO FILTER
# ─────────────────────────────────────────────────────────────────
def _is_command_echo(text: str) -> bool:
    words = text.split()
    if len(words) < Config.ECHO_MIN_WORDS:
        return False
    elapsed = time.time() - _last_spoken_time[0]
    if elapsed > Config.ECHO_WINDOW:
        return False
    last = _last_spoken[0]
    if not last:
        return False
    score = fuzz.partial_ratio(text.lower(), last)
    if score >= Config.ECHO_THRESHOLD:
        log.info("Echo filter: discarding %r (score=%d)", text, score)
        return True
    return False


# ─────────────────────────────────────────────────────────────────
# MICROPHONE SELECTION
# ─────────────────────────────────────────────────────────────────
def pick_best_microphone() -> sr.Microphone:
    """
    Uses whatever mic Windows has set as the default input device.
    User controls mic choice via Windows Sound Settings > Input.
    Works with built-in, BT, or wired without any code changes.
    """
    try:
        names = sr.Microphone.list_microphone_names()
        log.info("Available audio devices:")
        for i, n in enumerate(names):
            log.info("  [%d] %s", i, n)
    except Exception:
        pass
    # device_index=None = Windows default input (whatever is set in Sound Settings)
    log.info("Using Windows default input device")
    return sr.Microphone(device_index=None)


# ─────────────────────────────────────────────────────────────────
# MIC TEST
# ─────────────────────────────────────────────────────────────────
def _do_mic_test(speak_fn, recognizer, mic):
    """
    Report the active mic device, measure its RMS level over 2s, and speak
    a summary. Triggered by __mic_test__ special command from the UI.
    """
    try:
        import pyaudio, struct, math
        pa = pyaudio.PyAudio()
        dev_idx  = mic._device_index if hasattr(mic, "_device_index") else None
        dev_name = "unknown"
        if dev_idx is not None:
            try:
                dev_name = pa.get_device_info_by_index(dev_idx).get("name", "unknown")
            except Exception:
                pass
        pa.terminate()
    except Exception:
        dev_name = "unknown"

    # Measure RMS over 2 seconds — save threshold first so calibration isn't lost
    log.info("Mic test: measuring RMS on %r", dev_name)
    _saved_threshold = recognizer.energy_threshold
    try:
        with mic as source:
            recognizer.adjust_for_ambient_noise(source, duration=1.5)
        rms = recognizer.energy_threshold
        level = "very low" if rms < 60 else "low" if rms < 120 else "good" if rms < 400 else "high"
        msg = (
            f"Mic test complete. Using: {dev_name}. "
            f"Signal level is {level} at {int(rms)}. "
            f"Threshold is set to {int(_saved_threshold)}."
        )
    except Exception as e:
        msg = f"Mic test failed: {e}"
    finally:
        # Always restore the calibrated threshold — mic test must not change it
        recognizer.energy_threshold = _saved_threshold

    log.info("Mic test result: %s", msg)
    speak_fn(msg)


# ─────────────────────────────────────────────────────────────────
# AUDIO CAPTURE
# ─────────────────────────────────────────────────────────────────
def _apply_noise_reduction(audio: sr.AudioData) -> sr.AudioData:
    """
    Apply spectral-gating noise reduction via the `noisereduce` library.
    Adds ~100–200 ms latency. Gracefully returns original audio if the
    library isn't installed or processing fails.
    """
    try:
        import numpy as np
        import noisereduce as nr

        rate    = audio.sample_rate
        width   = audio.sample_width        # bytes per sample
        raw     = audio.get_raw_data()
        dtype   = np.int16 if width == 2 else np.int8
        samples = np.frombuffer(raw, dtype=dtype).astype(np.float32)

        # Use first 0.5 s as noise profile (or whole clip if shorter)
        noise_len  = min(int(rate * 0.5), len(samples))
        noise_clip = samples[:noise_len]
        reduced    = nr.reduce_noise(y=samples, y_noise=noise_clip, sr=rate,
                                     prop_decrease=0.75)
        out = reduced.astype(dtype).tobytes()
        return sr.AudioData(out, rate, width)
    except ImportError:
        return audio   # noisereduce not installed — silent no-op
    except Exception as e:
        log.debug("Noise reduction skipped: %s", e)
        return audio


def _capture(
    recognizer: sr.Recognizer,
    mic: sr.Microphone,
    interrupt_event: Optional[threading.Event] = None,
    typed_queue: Optional[queue.Queue] = None,
) -> Optional[sr.AudioData | str]:
    """
    Capture one audio phrase from the microphone.

    Returns either:
      - sr.AudioData   — mic captured speech
      - str            — text typed in the chat box during the capture window
      - None           — timeout / interrupted

    interrupt_event: aborts capture within ~1s when set (releases _mic_lock
    so listen_fn on the module thread can grab it).

    typed_queue: when provided, _capture checks it between every 1s poll
    window. This is the key fix for "too slow to type" — if the user types
    during the mic-listening window the text is returned immediately rather
    than waiting for the full 8s timeout to expire.
    """
    _POLL_WINDOW = 1.0   # seconds per short listen window

    total_waited = 0.0
    with mic as source:
        while True:
            # 1. Abort if another thread claimed the gate
            if interrupt_event is not None and interrupt_event.is_set():
                log.debug("_capture: interrupt_event set — releasing mic early")
                return None

            # 2. Check for chat-box input typed during this capture window
            if typed_queue is not None:
                try:
                    item = typed_queue.get_nowait()
                    if item == "__typing__":
                        # User is mid-sentence in the chat box — reset the
                        # timeout so we don't expire while they're still typing.
                        log.debug("_capture: typing signal received — resetting timeout")
                        total_waited = 0.0
                        continue
                    if item and item.strip():
                        log.info("_capture: chat input during mic window: %r", item)
                        return item.strip()
                except queue.Empty:
                    pass

            remaining_timeout = max(0.0, Config.LISTEN_TIMEOUT - total_waited)
            if remaining_timeout <= 0:
                return None   # overall timeout exhausted — no speech heard

            window = min(_POLL_WINDOW, remaining_timeout)
            try:
                audio = recognizer.listen(
                    source,
                    timeout=window,
                    phrase_time_limit=Config.PHRASE_TIME_LIMIT,
                )
                # Got audio — break out of polling loop
                break
            except sr.WaitTimeoutError:
                total_waited += window
                # Loop back: re-check interrupt + typed_queue before next window

    if Config.NOISE_SUPPRESS:
        audio = _apply_noise_reduction(audio)

    return audio


# Indian-English STT artifacts → canonical forms
# Google en-IN commonly mishears these high-frequency command words.
_STT_CORRECTIONS = {
    # ── ProVA name mishearings ────────────────────────────────────
    "pro va":    "prova",
    "pro vha":   "prova",
    "pro baa":   "prova",
    "hey pro":   "hey prova",
    "a pro va":  "hey prova",
    # Wake word variants heard in logs
    "hey goa":   "hey prova",    # log: score=71 'hey goa'
    "hey prabhu":"hey prova",    # log: score=78 'hey prabhu'
    "hey prova": "hey prova",
    "a prova":   "hey prova",

    # ── App / action mishearings ──────────────────────────────────
    "open excel":  "open excel",
    "open axle":   "open excel",
    "open access": "open excel",
    "remind me two": "remind me to",
    "centre mail":   "send email",
    "sand email":    "send email",
    "sand mail":     "send mail",
    "create folder": "create folder",
    "create fuller": "create folder",
    "delete fuller": "delete folder",
    "search four":   "search for",
    "such four":     "search for",
}

def _apply_stt_corrections(text: str) -> str:
    """Fix common Indian-English STT mishearings before intent parsing."""
    t = text.lower().strip()
    for wrong, right in _STT_CORRECTIONS.items():
        if wrong in t:
            corrected = t.replace(wrong, right)
            log.info("STT correction: %r → %r", t, corrected)
            return corrected
    return t


def transcribe(recognizer: sr.Recognizer, audio: sr.AudioData) -> Optional[str]:
    """
    Transcribe audio with en-IN primary and en-US fallback retry.
    Also applies STT correction table for common Indian-English mishearings.
    """
    _check_audio_energy(audio)

    # Primary: Indian English
    try:
        result = recognizer.recognize_google(
            audio, language="en-IN", show_all=False
        )
        return _apply_stt_corrections(result.lower().strip())
    except sr.UnknownValueError:
        pass  # fall through to retry
    except sr.RequestError as e:
        log.error("STT network error: %s", e)
        now = time.time()
        if now - getattr(transcribe, "_last_warn", 0) > 30:
            transcribe._last_warn = now  # type: ignore[attr-defined]
            log.warning("Google STT needs internet connection.")
        return None

    # Retry: US English — often picks up fast/clipped BT audio better
    if Config.STT_RETRY:
        try:
            log.debug("en-IN returned nothing — retrying with en-US")
            result = recognizer.recognize_google(
                audio, language="en-US", show_all=False
            )
            corrected = _apply_stt_corrections(result.lower().strip())
            log.info("STT en-US fallback succeeded: %r", corrected)
            return corrected
        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            pass

    return None


# ─────────────────────────────────────────────────────────────────
# WAKE WORD
# ─────────────────────────────────────────────────────────────────
def wait_for_wake_word(
    recognizer: sr.Recognizer,
    mic: sr.Microphone,
    callbacks: ProVACallbacks,
) -> bool:
    if _module_active.is_set():
        return False
    if not _tts_done.wait(timeout=15):
        return False
    if _module_active.is_set():
        return False
    with _mic_lock:
        audio = _capture(recognizer, mic, interrupt_event=_module_active)
    if not audio:
        return False
    text = transcribe(recognizer, audio)
    if not text:
        return False
    score = fuzz.partial_ratio(Config.WAKE_WORD, text)
    if score >= Config.WAKE_WORD_FUZZ:
        log.info("Wake word detected (score=%d): %r", score, text)
        return True
    log.debug("Not wake word (score=%d): %r", score, text)
    return False


# ─────────────────────────────────────────────────────────────────
# MAIN LOOP
# ─────────────────────────────────────────────────────────────────
def run(
    callbacks:   ProVACallbacks  = None,
    stop_event:  threading.Event = None,
    pause_event: threading.Event = None,
    text_queue:  queue.Queue     = None,
    muted_flag:  list            = None,
) -> None:
    # Fix: create fresh objects each call instead of sharing mutable defaults
    if callbacks   is None: callbacks   = ProVACallbacks()
    if stop_event  is None: stop_event  = threading.Event()
    if pause_event is None: pause_event = threading.Event()
    if text_queue  is None: text_queue  = queue.Queue()
    if muted_flag  is None: muted_flag  = [False]

    _reset_state()

    recognizer     = build_recognizer()
    recognizer_ref = [recognizer]

    threading.Thread(
        target=_speaker_thread,
        args=(muted_flag, callbacks, recognizer_ref),
        daemon=True,
    ).start()

    speak = make_speak_fn(callbacks, muted_flag)

    if Config.FORCE_MIC_INDEX is not None:
        mic = sr.Microphone(device_index=Config.FORCE_MIC_INDEX)
    else:
        try:
            mic = pick_best_microphone()
        except RuntimeError as e:
            log.critical(str(e))
            callbacks.on_status("ERROR")
            speak("No microphone found. Please check your audio settings.")
            _tts_queue.put(None)
            return

    callbacks.on_status("CALIBRATING")
    log.info("Calibrating microphone...")

    # Simple calibration: measure ambient noise, floor/cap, lock.
    # No JSON profiles — they were causing drift issues with BT mics.
    with mic as source:
        recognizer.adjust_for_ambient_noise(source, duration=Config.AMBIENT_DURATION)

    # Adaptive floor/cap — works for both BT (low RMS ~60-150) and built-in (200-500).
    # Calibration tells us what the room + mic actually produces; we just sanity-check it.
    # Floor: if calibration goes below 50, something is wrong — use 50 minimum.
    # Cap: if above 600, room was very noisy during calibration — use 600 to avoid deafness.
    if recognizer.energy_threshold < 50:
        recognizer.energy_threshold = 50
    if recognizer.energy_threshold > 600:
        recognizer.energy_threshold = 600

    log.info("Calibration done. Energy threshold: %.0f", recognizer.energy_threshold)
    # Lock — prevents threshold drifting upward during silence between commands.
    recognizer.dynamic_energy_threshold = False

    def listen_fn() -> Optional[str]:
        """
        Capture one phrase from the user.

        Chat-first design: checks text_queue BEFORE opening the mic.
        This means the user can type in the chat box at any point where
        ProVA is waiting for input (subject, body, confirmation, email
        address, yes/no) — voice and chat are interchangeable.

        Flow:
          1. Check text_queue immediately (non-blocking) — typed input wins
          2. Wait for any TTS to finish
          3. Re-check text_queue (user may have typed while TTS was playing)
          4. Fall back to mic capture
        """
        _module_active.set()
        callbacks.on_status("LISTENING")
        try:
            # 1. Immediate check — something already typed before we were called
            try:
                typed = text_queue.get_nowait()
                if typed == "__typing__":
                    pass  # ignore mid-typing signals at this stage
                elif typed and typed.strip():
                    log.info("listen_fn: chat input (immediate): %r", typed)
                    callbacks.on_user_speech(typed)
                    return typed.strip()
            except queue.Empty:
                pass

            # 2. Wait for TTS to finish so we don't pick up ProVA's own voice
            if not _tts_done.wait(timeout=15):
                log.warning("TTS gate timeout in listen_fn")

            # 3. Re-check after TTS — user may have typed while ProVA was speaking
            try:
                typed = text_queue.get_nowait()
                if typed == "__typing__":
                    pass  # ignore mid-typing signals at this stage
                elif typed and typed.strip():
                    log.info("listen_fn: chat input (post-TTS): %r", typed)
                    callbacks.on_user_speech(typed)
                    return typed.strip()
            except queue.Empty:
                pass

            # 4. Fall back to mic capture — also polls text_queue every 1s
            #    so typing in the chat box during the listen window works.
            with _mic_lock:
                result = _capture(recognizer, mic, typed_queue=text_queue)
            if result is None:
                return None
            # _capture returns str when chat-box input arrived mid-window.
            # __typing__ signals are consumed internally by _capture — never returned.
            if isinstance(result, str):
                log.info("listen_fn: chat input (mid-capture): %r", result)
                callbacks.on_user_speech(result)
                return result
            # Otherwise it's AudioData — transcribe normally
            text = transcribe(recognizer, result)
            if text and _is_command_echo(text):
                return None
            return text
        finally:
            _module_active.clear()
            callbacks.on_status("IDLE")

    # Start reminder system
    reminder_runner = None
    try:
        from modules.reminder_module import start_reminder_system
        reminder_runner, _ = start_reminder_system(speak)
        log.info("Reminder system started.")
    except Exception as e:
        log.error("Could not start reminder system: %s", e)

    router_init(reminder_runner)

    callbacks.on_status("IDLE")
    if Config.WAKE_WORD_ENABLED:
        speak("ProVA is ready. Say 'hey ProVA' to give a command.")
    else:
        speak("ProVA is ready. How can I help?")

    # ── Main loop ─────────────────────────────────────────────────
    while not stop_event.is_set():
        try:
            if pause_event.is_set():
                callbacks.on_status("PAUSED")
                time.sleep(0.2)
                continue

            # Typed input
            # IMPORTANT: check _module_active BEFORE reading text_queue.
            # If a module (email, file_manager confirm, etc.) is active,
            # it is waiting on listen_fn() which will drain text_queue itself.
            # Reading here would steal the typed text before listen_fn sees it,
            # causing the module to time out and the typed input to be lost.
            if _module_active.is_set():
                time.sleep(0.1)
                continue

            try:
                typed = text_queue.get_nowait()
                log.info("Typed: %r", typed)

                # ── Special internal commands ─────────────────────
                # These must be caught BEFORE on_user_speech and detect_intent
                # so they don't show as user messages or trigger "I didn't catch that".
                if typed == "__voice_test__":
                    _tts_queue.put("Hello! This is a voice test. Speed and volume can be adjusted in Settings.")
                    continue
                if typed == "__mic_test__":
                    _do_mic_test(speak, recognizer, mic)
                    continue
                if typed.lower() in ("hey prova",):
                    speak("I'm already listening. Just give me a command.")
                    continue

                callbacks.on_user_speech(typed)
                callbacks.on_status("THINKING")
                _busy_count[0] += 1

                # Fix: named function, not lambda, avoids late-binding closure bug
                def _typed_done():
                    _busy_count[0] = max(0, _busy_count[0] - 1)
                    if _busy_count[0] == 0:
                        callbacks.on_status("IDLE")

                route(detect_intent(typed), speak, listen_fn, _typed_done,
                      gate_event=_module_active,
                      status_fn=callbacks.on_status)
                continue
            except queue.Empty:
                pass

            if Config.WAKE_WORD_ENABLED:
                callbacks.on_status("IDLE")
                if not wait_for_wake_word(recognizer, mic, callbacks):
                    continue
                speak("Yes?")

            if _module_active.is_set():
                time.sleep(0.1)
                continue

            callbacks.on_status("LISTENING")
            if not _tts_done.wait(timeout=15):
                log.warning("TTS gate timeout in main loop")

            if _module_active.is_set():
                time.sleep(0.1)
                continue

            with _mic_lock:
                # Pass _module_active so _capture() aborts within ~1s if a module
                # (e.g. file_manager confirmation) claims the gate while we're
                # blocked inside recognizer.listen(). This releases _mic_lock quickly
                # so listen_fn() can grab it and capture the user's "yes/no".
                audio = _capture(recognizer, mic, interrupt_event=_module_active)

            if audio is None:
                # In wake-word mode, catching nothing after "Yes?" is surprising
                # so we tell the user. In always-listening mode, silence is normal
                # (user may just not be speaking) — silently loop back.
                if Config.WAKE_WORD_ENABLED:
                    speak("I didn't catch that.")
                callbacks.on_status("IDLE")
                continue

            text = transcribe(recognizer, audio)
            if text is None:
                # Same logic: in always-listening mode, failed transcription is
                # usually BT echo bleed or ambient noise — don't spam "Sorry I couldn't
                # understand" every 2 seconds. Just loop silently.
                if Config.WAKE_WORD_ENABLED:
                    speak("Sorry, I couldn't understand. Please try again.")
                callbacks.on_status("IDLE")
                continue

            if _is_command_echo(text):
                log.info("Command echo discarded: %r", text)
                callbacks.on_status("IDLE")
                continue

            log.info("Command: %r", text)
            callbacks.on_user_speech(text)
            callbacks.on_status("THINKING")
            _busy_count[0] += 1

            # Fix: named function for correct closure
            def _cmd_done():
                _busy_count[0] = max(0, _busy_count[0] - 1)
                if _busy_count[0] == 0:
                    callbacks.on_status("IDLE")

            route(detect_intent(text), speak, listen_fn, _cmd_done,
                  gate_event=_module_active,
                  status_fn=callbacks.on_status)

        except SystemExit:
            log.info("Shutdown requested.")
            break
        except KeyboardInterrupt:
            speak("Interrupted. Shutting down.")
            break
        except Exception as e:
            log.exception("Unexpected error: %s", e)
            speak("Something went wrong. Still listening.")

    _tts_queue.put(None)
    callbacks.on_status("STOPPED")
    log.info("ProVA stopped.")


if __name__ == "__main__":
    run()
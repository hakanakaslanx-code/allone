from __future__ import annotations

from dataclasses import dataclass
import importlib
import importlib.util
import queue
import threading
from typing import Optional

pythoncom = None
win32com = None
winsound = None

if importlib.util.find_spec("pythoncom") and importlib.util.find_spec("win32com.client"):
    pythoncom = importlib.import_module("pythoncom")
    win32com = importlib.import_module("win32com.client")

if importlib.util.find_spec("winsound"):
    winsound = importlib.import_module("winsound")


@dataclass(frozen=True)
class SpeechRequest:
    text: str
    speak_digits: bool
    feedback_mode: str
    success: bool


class SpeechQueue:
    def __init__(self) -> None:
        self._queue: "queue.Queue[Optional[SpeechRequest]]" = queue.Queue()
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._run, name="speech-queue", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        self._queue.put(None)

    def speak_barcode(self, text: str, *, speak_digits: bool, feedback_mode: str, success: bool) -> None:
        clean = (text or "").strip()
        if not clean:
            return
        self._queue.put(
            SpeechRequest(
                text=clean,
                speak_digits=speak_digits,
                feedback_mode=feedback_mode,
                success=success,
            )
        )

    def _run(self) -> None:
        voice = None
        com_initialized = False
        if pythoncom is not None and win32com is not None:
            try:
                pythoncom.CoInitialize()
                com_initialized = True
                voice = win32com.client.Dispatch("SAPI.SpVoice")
            except Exception:
                voice = None

        try:
            while not self._stop_event.is_set():
                request = self._queue.get()
                if request is None:
                    break
                self._handle_request(request, voice)
        finally:
            if com_initialized and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _handle_request(self, request: SpeechRequest, voice) -> None:
        if voice is not None:
            text = self._format_text(request.text, request.speak_digits)
            try:
                voice.Speak(text)
            except Exception:
                pass

        if request.feedback_mode == "tts" and voice is not None:
            try:
                voice.Speak("OK" if request.success else "Error")
            except Exception:
                pass
        elif request.feedback_mode == "beep":
            self._play_beep(request.success)

    @staticmethod
    def _format_text(text: str, speak_digits: bool) -> str:
        if not speak_digits:
            return text
        return " ".join(list(text))

    @staticmethod
    def _play_beep(success: bool) -> None:
        if winsound is None:
            return
        freq = 980 if success else 420
        duration = 120 if success else 200
        try:
            winsound.Beep(freq, duration)
        except Exception:
            pass

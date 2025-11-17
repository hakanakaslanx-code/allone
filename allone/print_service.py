"""LAN üzerinde paylaşılan etiket yazıcısı için gömülü Flask sunucusu."""

from __future__ import annotations

import logging
import os
import platform
import socket
import tempfile
import threading
from functools import wraps
from typing import Callable, Optional

from flask import Flask, jsonify, request
from werkzeug.utils import secure_filename

logger = logging.getLogger("allone.shared_printer")

app = Flask(__name__)
app.config.setdefault("MAX_CONTENT_LENGTH", 20 * 1024 * 1024)

_state_lock = threading.RLock()
_server_thread: Optional[threading.Thread] = None
_server_token: Optional[str] = None
_server_port: Optional[int] = None
_server_host: str = "127.0.0.1"
_sharing_enabled: bool = False
_active_server: Optional["SharedLabelPrinterServer"] = None


def resolve_local_ip() -> str:
    """Yerel ağdaki erişilebilir IP adresini döndürür."""
    try:
        # İnternete çıkış yapmaya gerek olmadan yerel IP'yi öğrenmek için UDP kullanıyoruz.
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"


def _is_thread_running() -> bool:
    with _state_lock:
        thread = _server_thread
    return bool(thread and thread.is_alive())


def _translate(key: str) -> str:
    with _state_lock:
        server = _active_server
    return server._tr(key) if server else key


def _run_flask_app() -> None:
    global _server_thread, _sharing_enabled
    try:
        with _state_lock:
            port = _server_port
        if port is None:
            logger.error("Shared printer server port is not configured.")
            return
        app.run(host="0.0.0.0", port=port, threaded=True, use_reloader=False)
    except Exception as exc:  # pragma: no cover - dayanıklılık
        logger.exception("Shared printer server error: %s", exc)
        with _state_lock:
            _sharing_enabled = False
    finally:
        with _state_lock:
            _server_thread = None


def _start_thread_locked() -> None:
    global _server_thread
    if _server_thread and _server_thread.is_alive():
        return
    if _server_port is None:
        raise RuntimeError(_translate("Server port is not configured."))
    thread = threading.Thread(
        target=_run_flask_app,
        name="SharedLabelPrinterServer",
        daemon=True,
    )
    _server_thread = thread
    thread.start()


def _ensure_port_available(port: int) -> None:
    test_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        test_socket.bind(("0.0.0.0", port))
    except OSError as exc:
        raise RuntimeError(
            _translate("Port {port} is not available: {error}").format(port=port, error=exc)
        ) from exc
    finally:
        test_socket.close()


def require_token(func: Callable) -> Callable:
    """Bearer token doğrulaması ekleyen dekoratör."""

    @wraps(func)
    def wrapper(*args, **kwargs):
        header = request.headers.get("Authorization", "")
        token = None
        if header.startswith("Bearer "):
            token = header.split(" ", 1)[1].strip()

        with _state_lock:
            expected = _server_token

        if not expected:
            return jsonify({"error": _translate("Server token is not configured.")}), 503

        if token != expected:
            return jsonify({"error": _translate("Invalid or missing authorization token.")}), 401

        return func(*args, **kwargs)

    return wrapper


@app.route("/status", methods=["GET"])
@require_token
def status():
    """Sunucu durumunu JSON olarak döndürür."""
    with _state_lock:
        running = _sharing_enabled and _is_thread_running()
        host = _server_host
        port = _server_port
        server = _active_server
    printer_name = server._printer_name if server else "DYMO LabelWriter 450"
    return jsonify(
        {
            "running": running,
            "printer_name": printer_name,
            "port": port,
            "host": host,
        }
    )


@app.route("/print", methods=["POST"])
@require_token
def print_job():
    """Gönderilen dosyayı yerel yazıcıya yollar."""
    with _state_lock:
        server = _active_server
        enabled = _sharing_enabled

    if not enabled:
        return jsonify({"error": _translate("SHARED_PRINTER_DISABLED")}), 409

    if server is None:
        return jsonify({"error": _translate("SHARED_PRINTER_NOT_READY")}), 503

    if "file" not in request.files:
        return jsonify({"error": server._tr("No file found in request.")}), 400

    upload = request.files["file"]
    if not upload or not upload.filename:
        return jsonify({"error": server._tr("No valid filename provided.")}), 400

    filename = secure_filename(upload.filename) or "job.bin"
    suffix = os.path.splitext(filename)[1] or ".bin"

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            upload.save(tmp.name)
            temp_path = tmp.name

        server._send_to_printer(temp_path)
    except Exception as exc:  # pragma: no cover - ortam bağımlı
        logger.exception(server._tr("Print error: %s"), exc)
        return jsonify({"error": str(exc)}), 500
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass

    return jsonify({"ok": True})


class SharedLabelPrinterServer:
    """Flask sunucusunu arka planda çalıştırıp yazdırma isteklerini yöneten sınıf."""

    def __init__(
        self,
        log_callback: Optional[Callable[[str], None]] = None,
        printer_name: str = "DYMO LabelWriter 450",
        translator: Optional[Callable[[str], str]] = None,
    ) -> None:
        # Kullanıcıya log mesajı gönderebilmek için isteğe bağlı geri çağırım saklanır.
        self._log_callback = log_callback
        self._lock = threading.RLock()
        self._port: Optional[int] = None
        self._token: Optional[str] = None
        self._host: str = "127.0.0.1"
        self._printer_name = printer_name
        self._win32print = None
        self._translator: Callable[[str], str] = translator or (lambda key: key)
        with _state_lock:
            global _active_server
            _active_server = self

    # ------------------------------------------------------------------
    # Yardımcı metodlar
    # ------------------------------------------------------------------
    def _log(self, message: str) -> None:
        """GUI günlük alanına ve uygulama loglarına mesaj gönderir."""
        if self._log_callback:
            self._log_callback(message)
        logger.info(message)

    def _load_win32print(self):
        """Windows yazdırma API'sini dinamik olarak yükler."""
        if platform.system() != "Windows":
            raise RuntimeError(self._tr("win32print is only available on Windows."))
        if self._win32print is None:
            try:
                import win32print  # type: ignore
            except ImportError as exc:  # pragma: no cover - ortam bağımlı
                raise RuntimeError(
                    self._tr(
                        "win32print module could not be loaded. Please check the pywin32 installation."
                    )
                ) from exc
            self._win32print = win32print
        return self._win32print

    def _send_to_printer(self, file_path: str) -> None:
        """Belirtilen dosyayı Windows yazdırma altyapısına iletir."""
        win32print = self._load_win32print()
        printer_name = self._printer_name or win32print.GetDefaultPrinter()
        if not printer_name:
            raise RuntimeError(self._tr("Printer name could not be determined."))

        handle = win32print.OpenPrinter(printer_name)
        try:
            job_info = ("Shared Label Print Job", None, "RAW")
            win32print.StartDocPrinter(handle, 1, job_info)
            win32print.StartPagePrinter(handle)
            with open(file_path, "rb") as job_file:
                data = job_file.read()
                if not data:
                    raise RuntimeError(self._tr("File content is empty; cannot print."))
                win32print.WritePrinter(handle, data)
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
        finally:
            win32print.ClosePrinter(handle)

    # ------------------------------------------------------------------
    # Genel API
    # ------------------------------------------------------------------
    def print_file(self, file_path: str) -> None:
        """Public wrapper to send a file to the configured printer."""

        if not file_path or not os.path.exists(file_path):
            raise RuntimeError(self._tr("Please select a file to print."))

        self._send_to_printer(file_path)

    def start(self, port: int, token: str) -> None:
        """Sunucuyu belirtilen port ve jetonla başlatır."""
        cleaned_token = token.strip()
        if not cleaned_token:
            raise RuntimeError(self._tr("Authorization token cannot be empty."))

        port_int = int(port)
        if not (1 <= port_int <= 65535):
            raise RuntimeError(self._tr("Please enter a valid port number."))

        host_ip = resolve_local_ip()

        with _state_lock:
            global _server_token, _server_port, _server_host, _sharing_enabled, _active_server
            thread_running = _server_thread and _server_thread.is_alive()
            if thread_running and _server_port not in (None, port_int):
                raise RuntimeError(self._tr("Server is already running."))

            if not thread_running:
                _ensure_port_available(port_int)
                _server_port = port_int
                _start_thread_locked()
            else:
                _server_port = _server_port or port_int

            _server_token = cleaned_token
            _server_host = host_ip
            _sharing_enabled = True
            _active_server = self

        with self._lock:
            self._port = port_int
            self._token = cleaned_token
            self._host = host_ip

        self._log(
            self._tr("SHARED_PRINTER_STARTED").format(host=self._host, port=self._port)
        )

    def stop(self, timeout: float = 5.0) -> None:  # noqa: ARG002 - timeout uyumluluk için
        """Sunucuyu nazikçe durdurur."""
        with _state_lock:
            global _sharing_enabled
            _sharing_enabled = False

        self._log(self._tr("SHARED_PRINTER_STOPPED"))

    def _tr(self, key: str) -> str:
        return self._translator(key) if self._translator else key

    def set_translator(self, translator: Optional[Callable[[str], str]]) -> None:
        self._translator = translator or (lambda key: key)

    def is_running(self) -> bool:
        """Sunucunun hala aktif olup olmadığını döndürür."""
        with _state_lock:
            return _sharing_enabled and _is_thread_running()

    def current_port(self) -> Optional[int]:
        """Aktif port bilgisini verir."""
        with _state_lock:
            return _server_port

    def current_token(self) -> Optional[str]:
        """Mevcut bearer jetonunu döndürür."""
        with _state_lock:
            return _server_token

    def current_host(self) -> str:
        """Ağa ilan edilen IP bilgisini döndürür."""
        with _state_lock:
            return _server_host


__all__ = ["SharedLabelPrinterServer", "resolve_local_ip"]

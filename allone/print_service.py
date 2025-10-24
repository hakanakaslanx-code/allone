"""LAN üzerinde paylaşılan etiket yazıcısı için gömülü Flask sunucusu."""

from __future__ import annotations

import logging
import os
import platform
import socket
import tempfile
import threading
from typing import Callable, Optional

from flask import Flask, jsonify, request
from werkzeug.serving import make_server
from werkzeug.utils import secure_filename

logger = logging.getLogger("allone.shared_printer")


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


class SharedLabelPrinterServer:
    """Flask sunucusunu arka planda çalıştırıp yazdırma isteklerini yöneten sınıf."""

    def __init__(
        self,
        log_callback: Optional[Callable[[str], None]] = None,
        printer_name: str = "DYMO LabelWriter 450",
    ) -> None:
        # Kullanıcıya log mesajı gönderebilmek için isteğe bağlı geri çağırım saklanır.
        self._log_callback = log_callback
        # Sunucu ve iş parçacığı durumunu korumak için kilit kullanıyoruz.
        self._lock = threading.RLock()
        self._server = None
        self._thread: Optional[threading.Thread] = None
        self._app: Optional[Flask] = None
        self._port: Optional[int] = None
        self._token: Optional[str] = None
        self._host: str = "127.0.0.1"
        self._printer_name = printer_name
        self._win32print = None

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
            raise RuntimeError("win32print sadece Windows üzerinde kullanılabilir.")
        if self._win32print is None:
            try:
                import win32print  # type: ignore
            except ImportError as exc:  # pragma: no cover - ortam bağımlı
                raise RuntimeError(
                    "win32print modülü yüklenemedi. Lütfen pywin32 kurulumunu kontrol edin."
                ) from exc
            self._win32print = win32print
        return self._win32print

    def _require_token(self, func):
        """Bearer token doğrulamasını sağlayan dekoratör."""

        def wrapper(*args, **kwargs):
            header = request.headers.get("Authorization", "")
            token = None
            if header.startswith("Bearer "):
                token = header.split(" ", 1)[1].strip()

            with self._lock:
                expected = self._token

            if not expected:
                return jsonify({"error": "Sunucu jetonu yapılandırılmamış."}), 503

            if token != expected:
                return jsonify({"error": "Geçersiz veya eksik yetkilendirme jetonu."}), 401

            return func(*args, **kwargs)

        return wrapper

    def _create_app(self) -> Flask:
        """Flask uygulamasını ve uç noktalarını tanımlar."""
        app = Flask("shared_label_printer")
        app.config.setdefault("MAX_CONTENT_LENGTH", 20 * 1024 * 1024)

        @app.route("/status", methods=["GET"])
        @self._require_token
        def status():
            """Sunucu durumunu JSON olarak döndürür."""
            return jsonify(
                {
                    "running": self.is_running(),
                    "printer_name": self._printer_name,
                    "port": self.current_port(),
                    "host": self._host,
                }
            )

        @app.route("/print", methods=["POST"])
        @self._require_token
        def print_job():
            """Gönderilen dosyayı yerel yazıcıya yollar."""
            if "file" not in request.files:
                return jsonify({"error": "Yüklenecek dosya bulunamadı."}), 400

            upload = request.files["file"]
            if not upload or not upload.filename:
                return jsonify({"error": "Geçerli bir dosya adı gönderilmedi."}), 400

            filename = secure_filename(upload.filename) or "job.bin"
            suffix = os.path.splitext(filename)[1] or ".bin"

            temp_path = None
            try:
                # Dosyayı geçici dizine kaydediyoruz.
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    upload.save(tmp.name)
                    temp_path = tmp.name

                # Kaydedilen dosyayı yazıcıya gönderiyoruz.
                self._send_to_printer(temp_path)
            except Exception as exc:
                logger.exception("Yazdırma sırasında hata oluştu: %s", exc)
                return jsonify({"error": str(exc)}), 500
            finally:
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except OSError:
                        pass

            return jsonify({"ok": True})

        return app

    def _send_to_printer(self, file_path: str) -> None:
        """Belirtilen dosyayı Windows yazdırma altyapısına iletir."""
        win32print = self._load_win32print()
        printer_name = self._printer_name or win32print.GetDefaultPrinter()
        if not printer_name:
            raise RuntimeError("Yazıcı adı bulunamadı.")

        handle = win32print.OpenPrinter(printer_name)
        try:
            job_info = ("Shared Label Print Job", None, "RAW")
            win32print.StartDocPrinter(handle, 1, job_info)
            win32print.StartPagePrinter(handle)
            with open(file_path, "rb") as job_file:
                data = job_file.read()
                if not data:
                    raise RuntimeError("Dosya içeriği boş olduğu için yazdırma yapılamadı.")
                win32print.WritePrinter(handle, data)
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
        finally:
            win32print.ClosePrinter(handle)

    def _serve(self) -> None:
        """Werkzeug sunucusunu arka planda çalıştırır."""
        try:
            if self._server is not None:
                self._server.serve_forever()
        except Exception as exc:  # pragma: no cover - dayanıklılık
            self._log(f"Paylaşılan yazıcı sunucusu hata verdi: {exc}")
        finally:
            self._cleanup()

    def _cleanup(self) -> None:
        """Sunucu kapandığında tüm durum değişkenlerini sıfırlar."""
        with self._lock:
            self._server = None
            self._thread = None
            self._app = None
            self._port = None
            self._token = None
            self._host = "127.0.0.1"

    # ------------------------------------------------------------------
    # Genel API
    # ------------------------------------------------------------------
    def start(self, port: int, token: str) -> None:
        """Sunucuyu belirtilen port ve jetonla başlatır."""
        with self._lock:
            if self.is_running():
                raise RuntimeError("Sunucu zaten çalışıyor.")

            self._port = int(port)
            self._token = token.strip()
            if not self._token:
                raise RuntimeError("Yetkilendirme jetonu boş olamaz.")

            self._host = resolve_local_ip()
            self._app = self._create_app()

            try:
                self._server = make_server("0.0.0.0", self._port, self._app)
            except OSError as exc:
                self._cleanup()
                raise RuntimeError(f"Port {self._port} kullanılamıyor: {exc}") from exc

            self._thread = threading.Thread(
                target=self._serve,
                name="SharedLabelPrinterServer",
                daemon=True,
            )
            self._thread.start()

        self._log(
            f"Etiket yazıcısı paylaşımı {self._host}:{self._port} adresinde başlatıldı."
        )

    def stop(self, timeout: float = 5.0) -> None:
        """Sunucuyu nazikçe durdurur."""
        with self._lock:
            server = self._server
            thread = self._thread

        if not server or not thread:
            return

        try:
            server.shutdown()
        finally:
            thread.join(timeout=timeout)
            self._cleanup()

        self._log("Etiket yazıcısı paylaşımı durduruldu.")

    def is_running(self) -> bool:
        """Sunucunun hala aktif olup olmadığını döndürür."""
        with self._lock:
            return bool(self._thread and self._thread.is_alive())

    def current_port(self) -> Optional[int]:
        """Aktif port bilgisini verir."""
        with self._lock:
            return self._port

    def current_token(self) -> Optional[str]:
        """Mevcut bearer jetonunu döndürür."""
        with self._lock:
            return self._token

    def current_host(self) -> str:
        """Ağa ilan edilen IP bilgisini döndürür."""
        with self._lock:
            return self._host


__all__ = ["SharedLabelPrinterServer", "resolve_local_ip"]

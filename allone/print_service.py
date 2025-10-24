"""Embedded LAN printing infrastructure for AllOne Tools GUI."""

from __future__ import annotations

import base64
import binascii
import ipaddress
import logging
import os
import platform
import socket
import tempfile
import threading
from dataclasses import dataclass
from functools import wraps
from typing import Callable, Dict, List, Optional, Tuple

from flask import Flask, jsonify, request
from werkzeug.serving import make_server
from werkzeug.utils import secure_filename
from zeroconf import ServiceBrowser, ServiceInfo, ServiceStateChange, Zeroconf

logger = logging.getLogger("allone.print_service")


class PrinterBackendError(RuntimeError):
    """Raised when the local printer backend cannot be initialised."""


class PrinterBackend:
    """Abstract interface for accessing platform printers."""

    def list_printers(self) -> List[Dict[str, str]]:
        raise NotImplementedError

    def print_job(self, printer_name: str, file_path: str) -> None:
        raise NotImplementedError


class Win32PrinterBackend(PrinterBackend):
    """Windows printer backend using the win32print module."""

    def __init__(self, win32print_module):
        self._win32print = win32print_module

    def list_printers(self) -> List[Dict[str, str]]:
        flags = (
            self._win32print.PRINTER_ENUM_LOCAL
            | self._win32print.PRINTER_ENUM_CONNECTIONS
        )
        printers = []
        for _, _, name, comment in self._win32print.EnumPrinters(flags):
            printers.append({"name": name, "comment": comment or ""})
        return printers

    def print_job(self, printer_name: str, file_path: str) -> None:
        with open(file_path, "rb") as handle:
            data = handle.read()

        if not data:
            raise RuntimeError("Print job is empty.")

        printer = self._win32print.OpenPrinter(printer_name)
        try:
            job_info = ("AllOneTools Print Job", None, "RAW")
            self._win32print.StartDocPrinter(printer, 1, job_info)
            self._win32print.StartPagePrinter(printer)
            self._win32print.WritePrinter(printer, data)
            self._win32print.EndPagePrinter(printer)
            self._win32print.EndDocPrinter(printer)
        finally:
            self._win32print.ClosePrinter(printer)


class CUPSPrinterBackend(PrinterBackend):
    """CUPS printer backend for Linux and macOS."""

    def __init__(self, cups_module):
        self._cups = cups_module
        self._connection = self._cups.Connection()

    def list_printers(self) -> List[Dict[str, str]]:
        printers = []
        for name, attrs in self._connection.getPrinters().items():
            printers.append(
                {
                    "name": name,
                    "info": attrs.get("printer-info", ""),
                    "location": attrs.get("printer-location", ""),
                }
            )
        return printers

    def print_job(self, printer_name: str, file_path: str) -> None:
        self._connection.printFile(printer_name, file_path, "AllOneTools Print Job", {})


def build_backend() -> PrinterBackend:
    system = platform.system()
    if system == "Windows":
        try:
            import win32print  # type: ignore
        except ImportError as exc:  # pragma: no cover - import-time failure
            raise PrinterBackendError(
                "pywin32 is required to access printers on Windows."
            ) from exc
        return Win32PrinterBackend(win32print)

    try:
        import cups  # type: ignore
    except ImportError as exc:  # pragma: no cover - import-time failure
        raise PrinterBackendError(
            "pycups is required to access printers on this platform."
        ) from exc
    return CUPSPrinterBackend(cups)


HOSTNAME = socket.gethostname()


def make_network_printer_name(printer_name: str, hostname: Optional[str] = None) -> str:
    host = (hostname or HOSTNAME).strip()
    base = (printer_name or "").strip()
    if not host:
        return base
    suffix = f" - {host}"
    if base.endswith(suffix):
        return base
    return f"{base}{suffix}" if base else suffix[3:]


def split_network_printer_name(printer_name: str) -> Tuple[str, Optional[str]]:
    parts = (printer_name or "").rsplit(" - ", 1)
    if len(parts) == 2:
        base, host = parts
        return base.strip(), host.strip() or None
    return printer_name.strip(), None


def decorate_printer_entry(entry: Dict[str, str], hostname: Optional[str] = None) -> Dict[str, str]:
    decorated = dict(entry)
    raw_value = decorated.get("name") or decorated.get("info") or ""
    if not isinstance(raw_value, str):
        raw_value = str(raw_value)
    raw_name = raw_value.strip() or "Unknown Printer"
    host = hostname or HOSTNAME
    decorated["raw_name"] = raw_name
    decorated["hostname"] = host
    decorated["name"] = make_network_printer_name(raw_name, host)
    return decorated


def resolve_local_ip() -> Optional[str]:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            return sock.getsockname()[0]
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return None


def is_local_network_ip(ip: str) -> bool:
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return False
    return addr.is_private or addr.is_loopback or addr.is_link_local


@dataclass
class DiscoveredPrinter:
    name: str
    raw_name: str
    hostname: str
    address: str
    port: int
    source: str
    properties: Dict[str, str]


class EmbeddedPrintServer:
    """Flask server embedded into the GUI for sharing printers over LAN."""

    def __init__(self, log_callback: Optional[Callable[[str], None]] = None):
        self._lock = threading.RLock()
        self._server = None
        self._thread: Optional[threading.Thread] = None
        self._app: Optional[Flask] = None
        self._token: Optional[str] = None
        self._port: Optional[int] = None
        self._shared = False
        self._service_infos: Dict[str, ServiceInfo] = {}
        self._zeroconf: Optional[Zeroconf] = None
        self._backend: Optional[PrinterBackend] = None
        self._backend_error: Optional[str] = None
        self._log_callback = log_callback

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _log(self, message: str) -> None:
        if self._log_callback:
            self._log_callback(message)
        logger.info(message)

    def _ensure_backend(self) -> Optional[PrinterBackend]:
        with self._lock:
            if self._backend or self._backend_error:
                return self._backend

        try:
            backend = build_backend()
        except PrinterBackendError as exc:
            with self._lock:
                self._backend = None
                self._backend_error = str(exc)
            self._log(f"Printer backend unavailable: {exc}")
            return None
        except Exception as exc:  # pragma: no cover - unexpected failure
            with self._lock:
                self._backend = None
                self._backend_error = str(exc)
            self._log(f"Unexpected printer backend error: {exc}")
            return None

        with self._lock:
            self._backend = backend
            self._backend_error = None
        return backend

    def backend_status(self) -> Tuple[bool, Optional[str]]:
        with self._lock:
            return self._backend is not None, self._backend_error

    def _require_token(self, func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            header = request.headers.get("Authorization", "")
            token = None
            if header.startswith("Bearer "):
                token = header.split(" ", 1)[1].strip()

            with self._lock:
                expected = self._token

            if not expected:
                return jsonify({"error": "Print server token is not configured."}), 503

            if token != expected:
                return jsonify({"error": "Invalid bearer token."}), 401

            return func(*args, **kwargs)

        return wrapper

    def _create_app(self) -> Flask:
        app = Flask("embedded_print_server")
        app.config.setdefault("MAX_CONTENT_LENGTH", 20 * 1024 * 1024)

        @app.before_request
        def restrict_network():
            remote_ip = request.headers.get("X-Forwarded-For")
            if remote_ip:
                remote_ip = remote_ip.split(",")[0].strip()
            else:
                remote_ip = request.remote_addr

            if remote_ip is None:
                return jsonify({"error": "Client IP could not be verified."}), 403

            if not is_local_network_ip(remote_ip):
                return (
                    jsonify({"error": "Only local network clients may access the server."}),
                    403,
                )

            return None

        @app.route("/status", methods=["GET"])
        @self._require_token
        def status():
            ready, error = self.backend_status()
            return jsonify(
                {
                    "sharing": self.is_shared(),
                    "backend_ready": ready,
                    "backend_error": error,
                    "zeroconf": bool(self._service_infos),
                }
            )

        @app.route("/printers", methods=["GET"])
        @self._require_token
        def list_printers():
            backend = self._ensure_backend()
            if backend is None:
                _, error = self.backend_status()
                return (
                    jsonify({"error": error or "Printer backend unavailable."}),
                    503,
                )

            try:
                printers = [
                    decorate_printer_entry(printer)
                    for printer in backend.list_printers()
                ]
            except Exception as exc:
                logger.exception("Failed to list printers: %s", exc)
                return jsonify({"error": str(exc)}), 500

            return jsonify({"printers": printers})

        @app.route("/enable", methods=["POST"])
        @self._require_token
        def enable():
            try:
                enabled = self.enable_local_sharing()
            except RuntimeError as exc:
                return jsonify({"error": str(exc)}), 503
            return jsonify({"sharing": bool(enabled)})

        @app.route("/disable", methods=["POST"])
        @self._require_token
        def disable():
            self.disable_local_sharing()
            return jsonify({"sharing": False})

        @app.route("/print", methods=["POST"])
        @self._require_token
        def print_endpoint():
            if not self.is_shared():
                return jsonify({"error": "Printer sharing is disabled."}), 409

            backend = self._ensure_backend()
            if backend is None:
                _, error = self.backend_status()
                return (
                    jsonify({"error": error or "Printer backend unavailable."}),
                    503,
                )

            printer_name = request.form.get("printer_name")
            if request.is_json and not printer_name:
                printer_name = (request.json or {}).get("printer_name")

            if not printer_name:
                return jsonify({"error": "printer_name is required."}), 400

            raw_printer_name, _ = split_network_printer_name(printer_name)

            try:
                temp_path = self._prepare_job_file()
                backend.print_job(raw_printer_name, temp_path)
            except ValueError as exc:
                return jsonify({"error": str(exc)}), 400
            except Exception as exc:
                logger.exception("Print job failed: %s", exc)
                return jsonify({"error": str(exc)}), 500
            finally:
                if "temp_path" in locals() and temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except OSError:
                        pass

            return jsonify({"status": "printed", "printer_name": printer_name})

        return app

    def _prepare_job_file(self) -> str:
        if request.files:
            upload = request.files.get("file")
            if upload is None or not upload.filename:
                raise ValueError("No file uploaded.")

            suffix = os.path.splitext(secure_filename(upload.filename))[1] or ".bin"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                upload.save(tmp.name)
                return tmp.name

        payload = request.get_json(silent=True) or {}
        data_bytes: Optional[bytes] = None

        if "data_base64" in payload:
            try:
                data_bytes = base64.b64decode(payload["data_base64"], validate=True)
            except (ValueError, binascii.Error) as exc:
                raise ValueError(f"Invalid base64 payload: {exc}") from exc
        elif "text" in payload:
            data_bytes = payload["text"].encode("utf-8")

        if not data_bytes:
            raise ValueError("No printable data provided.")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as tmp:
            tmp.write(data_bytes)
            tmp.flush()
            return tmp.name

    def _serve(self) -> None:
        try:
            if self._server is not None:
                self._server.serve_forever()
        except Exception as exc:  # pragma: no cover - best effort logging
            self._log(f"Print server encountered an error: {exc}")
        finally:
            self._finalize_stop()

    def _finalize_stop(self) -> None:
        self.disable_local_sharing()
        with self._lock:
            self._server = None
            self._thread = None
            self._app = None
            self._token = None
            self._port = None
            self._shared = False

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def start(self, port: int, token: str, host: str = "0.0.0.0") -> None:
        backend = self._ensure_backend()
        if backend is None:
            ready, error = self.backend_status()
            raise RuntimeError(error or "Printer backend unavailable.")

        with self._lock:
            if self.is_running():
                raise RuntimeError("Server is already running.")

            self._token = token.strip()
            self._port = port
            self._app = self._create_app()

            try:
                self._server = make_server(host, port, self._app)
            except OSError as exc:
                self._finalize_stop()
                raise exc

            self._thread = threading.Thread(
                target=self._serve,
                name="EmbeddedPrintServer",
                daemon=True,
            )
            self._thread.start()

    def stop(self, timeout: float = 5.0) -> None:
        with self._lock:
            server = self._server
            thread = self._thread

        if server is None or thread is None:
            return

        try:
            server.shutdown()
        finally:
            thread.join(timeout=timeout)
            self._finalize_stop()

    def is_running(self) -> bool:
        with self._lock:
            return bool(self._thread and self._thread.is_alive())

    def is_shared(self) -> bool:
        with self._lock:
            return self._shared

    def current_port(self) -> Optional[int]:
        with self._lock:
            return self._port

    def current_token(self) -> Optional[str]:
        with self._lock:
            return self._token

    def list_local_printers(self) -> List[Dict[str, str]]:
        backend = self._ensure_backend()
        if backend is None:
            return []

        try:
            printers = backend.list_printers()
        except Exception as exc:
            self._log(f"Failed to enumerate printers: {exc}")
            return []

        return [decorate_printer_entry(printer) for printer in printers]

    # ------------------------------------------------------------------
    # Sharing helpers
    # ------------------------------------------------------------------
    def enable_local_sharing(self) -> bool:
        backend = self._ensure_backend()
        if backend is None:
            raise RuntimeError(self._backend_error or "Printer backend unavailable.")

        with self._lock:
            if self._shared:
                return True
            port = self._port
            if port is None:
                raise RuntimeError("Server is not running.")
            self._shared = True

        self._register_services()
        return True

    def disable_local_sharing(self) -> bool:
        with self._lock:
            was_shared = self._shared
            self._shared = False

        if was_shared:
            self._unregister_services()
        return was_shared

    def _register_services(self) -> None:
        printers = self.list_local_printers()
        if not printers:
            printers = [
                {
                    "name": make_network_printer_name("AllOneTools Print Server"),
                    "raw_name": "AllOneTools Print Server",
                    "hostname": HOSTNAME,
                }
            ]

        ip_address = resolve_local_ip()
        if ip_address is None:
            self._log("Could not determine local IP for Zeroconf broadcast.")
            return

        try:
            address_bytes = socket.inet_aton(ip_address)
        except OSError:
            self._log(f"Invalid IP address for Zeroconf broadcast: {ip_address}")
            return

        zeroconf = Zeroconf()
        failures = 0

        for entry in printers:
            service_name = f"{entry['name']}._printer._tcp.local."
            try:
                info = ServiceInfo(
                    type_="_printer._tcp.local.",
                    name=service_name,
                    addresses=[address_bytes],
                    port=self.current_port() or 0,
                    properties={
                        b"sharing": b"true",
                        b"hostname": entry.get("hostname", HOSTNAME).encode("utf-8", "ignore"),
                        b"raw_name": entry.get("raw_name", "").encode("utf-8", "ignore"),
                    },
                    server=f"{HOSTNAME}.local.",
                )
            except Exception as exc:
                failures += 1
                logger.warning("Failed to prepare Zeroconf service info: %s", exc)
                continue

            try:
                zeroconf.register_service(info, cooperating_responders=True)
                self._service_infos[entry["name"]] = info
            except Exception as exc:
                failures += 1
                logger.warning("Failed to register Zeroconf service for %s: %s", entry, exc)

        if failures and not self._service_infos:
            try:
                zeroconf.close()
            except Exception:
                pass
            return

        self._zeroconf = zeroconf

    def _unregister_services(self) -> None:
        zeroconf = self._zeroconf
        services = list(self._service_infos.items())
        self._service_infos = {}
        self._zeroconf = None

        if zeroconf is None:
            return

        for name, info in services:
            try:
                zeroconf.unregister_service(info)
            except Exception:
                logger.warning("Failed to unregister Zeroconf service for %s", name)

        try:
            zeroconf.close()
        except Exception:
            pass


class NetworkPrinterBrowser:
    """Discovers printers advertised over Zeroconf on the local network."""

    def __init__(self, callback: Optional[Callable[[List[DiscoveredPrinter]], None]] = None, log_callback: Optional[Callable[[str], None]] = None):
        self._callback = callback
        self._log = log_callback
        self._lock = threading.RLock()
        self._zeroconf: Optional[Zeroconf] = None
        self._browser: Optional[ServiceBrowser] = None
        self._printers: Dict[str, DiscoveredPrinter] = {}

        try:
            self._zeroconf = Zeroconf()
            self._browser = ServiceBrowser(
                self._zeroconf,
                "_printer._tcp.local.",
                handlers=[self._handle_service_state],
            )
        except Exception as exc:
            if self._log:
                self._log(f"Zeroconf discovery unavailable: {exc}")
            logger.warning("Failed to start Zeroconf browser: %s", exc)
            self._zeroconf = None
            self._browser = None

    def is_available(self) -> bool:
        return self._zeroconf is not None and self._browser is not None

    def _notify(self) -> None:
        if not self._callback:
            return
        with self._lock:
            printers = list(self._printers.values())
        try:
            self._callback(printers)
        except Exception:  # pragma: no cover - protect GUI callback
            logger.exception("Printer discovery callback failed")

    def _handle_service_state(self, zeroconf: Zeroconf, service_type: str, name: str, state_change: ServiceStateChange) -> None:
        if state_change is ServiceStateChange.Removed:
            with self._lock:
                self._printers.pop(name, None)
            self._notify()
            return

        info = zeroconf.get_service_info(service_type, name)
        if not info:
            return

        addresses: List[str] = []
        for raw in info.addresses:
            if len(raw) == 4:  # IPv4 only
                addresses.append(socket.inet_ntoa(raw))

        if not addresses:
            return

        props = {}
        for key, value in info.properties.items():
            try:
                props[key.decode("utf-8")] = value.decode("utf-8") if isinstance(value, (bytes, bytearray)) else str(value)
            except Exception:
                continue

        display_name = name.split("._printer._tcp.local.")[0]
        raw_name = props.get("raw_name", display_name)
        hostname = props.get("hostname", info.server.rstrip("."))

        printer = DiscoveredPrinter(
            name=display_name,
            raw_name=raw_name,
            hostname=hostname,
            address=addresses[0],
            port=info.port,
            source="remote",
            properties=props,
        )

        with self._lock:
            self._printers[name] = printer
        self._notify()

    def printers(self) -> List[DiscoveredPrinter]:
        with self._lock:
            return list(self._printers.values())

    def close(self) -> None:
        if self._browser:
            try:
                self._browser.cancel()
            except Exception:
                pass
        if self._zeroconf:
            try:
                self._zeroconf.close()
            except Exception:
                pass
        self._browser = None
        self._zeroconf = None


__all__ = [
    "EmbeddedPrintServer",
    "NetworkPrinterBrowser",
    "DiscoveredPrinter",
    "make_network_printer_name",
    "split_network_printer_name",
]

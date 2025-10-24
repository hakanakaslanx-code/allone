#!/usr/bin/env python3
"""AllOneTools Print Server.

Expose locally attached printers (such as the Dymo LabelWriter 450) over the
local network via a simple authenticated Flask API. The service discovers
printers using the platform's native APIs (pywin32 on Windows, pycups on
Linux/macOS) and advertises itself via Zeroconf so that other machines on the
same network can find it automatically.
"""

from __future__ import annotations

import argparse
import base64
import binascii
import ipaddress
import logging
import os
import platform
import socket
import tempfile
from abc import ABC, abstractmethod
from functools import wraps
from typing import Dict, Iterable, List, Optional, Tuple

from flask import Flask, jsonify, request
from werkzeug.utils import secure_filename
from zeroconf import ServiceInfo, Zeroconf

import atexit
import secrets
import threading


class PrinterBackend(ABC):
    """Abstract interface for platform-specific printer access."""

    @abstractmethod
    def list_printers(self) -> List[Dict[str, str]]:
        """Return a list of available printers."""

    @abstractmethod
    def print_job(self, printer_name: str, file_path: str) -> None:
        """Send a file to the specified printer."""


class Win32PrinterBackend(PrinterBackend):
    """Windows printer backend that uses pywin32 to talk to the spooler."""

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
    """cups-based backend for Linux and macOS hosts."""

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
            raise RuntimeError(
                "pywin32 is required to access printers on Windows."
            ) from exc
        return Win32PrinterBackend(win32print)

    try:
        import cups  # type: ignore
    except ImportError as exc:  # pragma: no cover - import-time failure
        raise RuntimeError("pycups is required to access printers on Linux/macOS.") from exc
    return CUPSPrinterBackend(cups)


app = Flask(__name__)
app.config.setdefault("MAX_CONTENT_LENGTH", 20 * 1024 * 1024)  # 20 MB default

logger = logging.getLogger("allonetools.print_server")

SHARING_LOCK = threading.Lock()
SHARING_ENABLED = False
SERVICE_INFOS: Dict[str, ServiceInfo] = {}
ZEROCONF: Optional[Zeroconf] = None
BACKEND: Optional[PrinterBackend] = None
BACKEND_ERROR: Optional[str] = None
AUTH_TOKEN: Optional[str] = None
SERVER_PORT: int = 5151
HOSTNAME: str = socket.gethostname()


def initialise_backend() -> None:
    global BACKEND, BACKEND_ERROR
    try:
        BACKEND = build_backend()
        BACKEND_ERROR = None
    except Exception as exc:  # pragma: no cover - platform specific
        BACKEND = None
        BACKEND_ERROR = str(exc)
        logger.error("Unable to initialise printer backend: %s", exc)


def make_network_printer_name(printer_name: str, hostname: Optional[str] = None) -> str:
    host = (hostname or HOSTNAME).strip()
    base = printer_name.strip()
    if not host:
        return base
    suffix = f" - {host}"
    if base.endswith(suffix):
        return base
    return f"{base}{suffix}"


def split_network_printer_name(printer_name: str) -> Tuple[str, Optional[str]]:
    parts = printer_name.rsplit(" - ", 1)
    if len(parts) == 2:
        base, host = parts
        return base.strip(), host.strip() or None
    return printer_name.strip(), None


def decorate_printer_entry(entry: Dict[str, str]) -> Dict[str, str]:
    # Copy entry to avoid mutating backend data
    decorated = dict(entry)
    raw_name_value = decorated.get("name") or decorated.get("info") or ""
    if not isinstance(raw_name_value, str):
        raw_name_value = str(raw_name_value)
    raw_name = raw_name_value.strip()

    decorated["raw_name"] = raw_name
    decorated["hostname"] = HOSTNAME
    decorated["name"] = make_network_printer_name(raw_name or "Unknown Printer")
    return decorated


def require_token(view_func):
    @wraps(view_func)
    def wrapper(*args, **kwargs):
        if not AUTH_TOKEN:
            return error_response(503, "Print server token is not configured.")

        header = request.headers.get("Authorization", "")
        if not header.startswith("Bearer "):
            return error_response(401, "Authorization header missing or invalid.")

        token = header.split(" ", 1)[1].strip()
        if token != AUTH_TOKEN:
            return error_response(401, "Invalid bearer token.")

        return view_func(*args, **kwargs)

    return wrapper


@app.before_request
def restrict_network() -> Optional[object]:
    remote_ip = extract_request_ip()
    if remote_ip is None:
        logger.warning("Could not determine client IP address.")
        return error_response(403, "Client IP could not be verified.")

    if not is_local_network_ip(remote_ip):
        logger.warning("Rejected request from non-local address: %s", remote_ip)
        return error_response(403, "Only local network clients may access the server.")

    return None


@app.route("/status", methods=["GET"])
@require_token
def status() -> object:
    return jsonify(
        {
            "sharing": SHARING_ENABLED,
            "backend_ready": BACKEND is not None,
            "backend_error": BACKEND_ERROR,
            "zeroconf": bool(SERVICE_INFOS),
        }
    )


@app.route("/printers", methods=["GET"])
@require_token
def list_printers() -> object:
    if BACKEND is None:
        return error_response(503, BACKEND_ERROR or "Printer backend unavailable.")

    try:
        printers = BACKEND.list_printers()
    except Exception as exc:
        logger.exception("Failed to list printers: %s", exc)
        return error_response(500, f"Failed to list printers: {exc}")

    enriched = [decorate_printer_entry(printer) for printer in printers]

    return jsonify({"printers": enriched})


@app.route("/enable", methods=["POST"])
@require_token
def enable() -> object:
    global SHARING_ENABLED

    if BACKEND is None:
        return error_response(503, BACKEND_ERROR or "Printer backend unavailable.")

    with SHARING_LOCK:
        if SHARING_ENABLED:
            return jsonify({"status": "already-enabled"})

        SHARING_ENABLED = True
        register_service()

    logger.info("Printer sharing enabled.")
    return jsonify({"status": "enabled"})


@app.route("/disable", methods=["POST"])
@require_token
def disable() -> object:
    global SHARING_ENABLED

    with SHARING_LOCK:
        if not SHARING_ENABLED:
            return jsonify({"status": "already-disabled"})

        SHARING_ENABLED = False
        unregister_service()

    logger.info("Printer sharing disabled.")
    return jsonify({"status": "disabled"})


@app.route("/print", methods=["POST"])
@require_token
def print_endpoint() -> object:
    if not SHARING_ENABLED:
        return error_response(409, "Printer sharing is disabled.")

    if BACKEND is None:
        return error_response(503, BACKEND_ERROR or "Printer backend unavailable.")

    printer_name = (
        request.form.get("printer_name")
        or (request.json or {}).get("printer_name")
        if request.is_json
        else None
    )

    if not printer_name:
        return error_response(400, "printer_name is required.")

    raw_printer_name, requested_host = split_network_printer_name(printer_name)

    try:
        temp_path = prepare_job_file()
        BACKEND.print_job(raw_printer_name, temp_path)
    except ValueError as exc:
        logger.warning("Invalid print payload: %s", exc)
        return error_response(400, str(exc))
    except Exception as exc:
        logger.exception("Print job failed: %s", exc)
        return error_response(500, f"Print job failed: {exc}")
    finally:
        if 'temp_path' in locals() and temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                logger.warning("Could not remove temporary file: %s", temp_path)

    logger.info(
        "Print job submitted to %s (requested via %s)",
        raw_printer_name,
        make_network_printer_name(raw_printer_name, requested_host or HOSTNAME),
    )
    return jsonify(
        {
            "status": "printed",
            "printer_name": make_network_printer_name(raw_printer_name, requested_host or HOSTNAME),
            "raw_printer_name": raw_printer_name,
        }
    )


def prepare_job_file() -> str:
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


def extract_request_ip() -> Optional[str]:
    header = request.headers.get("X-Forwarded-For")
    if header:
        # take left-most address
        ip = header.split(",")[0].strip()
        if ip:
            return ip
    return request.remote_addr


def is_local_network_ip(ip: str) -> bool:
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return False
    return addr.is_private or addr.is_loopback or addr.is_link_local


def register_service() -> None:
    global SERVICE_INFOS, ZEROCONF

    if SERVICE_INFOS:
        return

    if BACKEND is None:
        logger.warning("Cannot register Zeroconf service without an initialised backend.")
        return

    ip_address = resolve_local_ip()
    if ip_address is None:
        logger.warning("Could not determine local IP for Zeroconf broadcast.")
        return

    try:
        address_bytes = socket.inet_aton(ip_address)
    except OSError as exc:
        logger.warning("Invalid IP address for Zeroconf broadcast (%s): %s", ip_address, exc)
        return

    try:
        printers = BACKEND.list_printers()
    except Exception as exc:  # pragma: no cover - backend specific
        logger.warning("Failed to enumerate printers for Zeroconf: %s", exc)
        printers = []

    entries = [decorate_printer_entry(printer) for printer in printers]

    if not entries:
        # Advertise the server itself so discovery clients still learn about it.
        entries = [
            {
                "name": make_network_printer_name("AllOneTools Print Server"),
                "raw_name": "AllOneTools Print Server",
                "hostname": HOSTNAME,
            }
        ]

    ZEROCONF = Zeroconf()

    for entry in entries:
        service_name = f"{entry['name']}._printer._tcp.local."
        try:
            info = ServiceInfo(
                type_="_printer._tcp.local.",
                name=service_name,
                addresses=[address_bytes],
                port=SERVER_PORT,
                properties={
                    b"sharing": b"true",
                    b"hostname": entry.get("hostname", HOSTNAME).encode("utf-8", "ignore"),
                    b"raw_name": entry.get("raw_name", "").encode("utf-8", "ignore"),
                },
                server=f"{HOSTNAME}.local.",
                weight=0,
                priority=0,
            )
        except Exception as exc:  # pragma: no cover - data validation specifics
            logger.warning("Failed to prepare Zeroconf service info for %s: %s", entry, exc)
            continue

        try:
            ZEROCONF.register_service(info, cooperating_responders=True)
            SERVICE_INFOS[entry["name"]] = info
            logger.info(
                "Zeroconf service registered for printer '%s' as '%s' on %s:%s",
                entry.get("raw_name"),
                entry["name"],
                ip_address,
                SERVER_PORT,
            )
        except Exception as exc:  # pragma: no cover - zeroconf specifics
            logger.warning(
                "Failed to register Zeroconf service for %s (%s): %s",
                entry.get("raw_name"),
                entry["name"],
                exc,
            )

    if not SERVICE_INFOS:
        # If registration failed for all entries, ensure the Zeroconf socket is closed.
        if ZEROCONF:
            try:
                ZEROCONF.close()
            except Exception:
                pass
            ZEROCONF = None


def unregister_service() -> None:
    global SERVICE_INFOS, ZEROCONF

    if ZEROCONF and SERVICE_INFOS:
        for name, info in list(SERVICE_INFOS.items()):
            try:
                ZEROCONF.unregister_service(info)
                logger.info("Zeroconf service unregistered for %s", name)
            except Exception as exc:  # pragma: no cover
                logger.warning("Failed to unregister Zeroconf service for %s: %s", name, exc)

    if ZEROCONF:
        try:
            ZEROCONF.close()
        except Exception:
            pass
        ZEROCONF = None

    SERVICE_INFOS = {}


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


def error_response(status: int, message: str) -> object:
    response = jsonify({"error": message, "status": status})
    response.status_code = status
    return response


def shutdown_service() -> None:
    unregister_service()


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="AllOneTools Print Server")
    parser.add_argument("--host", default="0.0.0.0", help="Host/IP to bind the Flask server")
    parser.add_argument("--port", type=int, default=5151, help="TCP port to bind the Flask server")
    parser.add_argument(
        "--token",
        default=os.environ.get("PRINT_SERVER_TOKEN", ""),
        help="Bearer token required for API access (defaults to PRINT_SERVER_TOKEN env)",
    )
    parser.add_argument(
        "--log-level",
        default=os.environ.get("PRINT_SERVER_LOG_LEVEL", "INFO"),
        help="Logging level (DEBUG, INFO, WARNING, ERROR)",
    )
    return parser.parse_args(argv)


def configure_logging(level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


def main(argv: Optional[Iterable[str]] = None) -> None:
    global AUTH_TOKEN, SERVER_PORT

    args = parse_args(argv)
    configure_logging(args.log_level)

    initialise_backend()

    SERVER_PORT = args.port
    AUTH_TOKEN = args.token.strip()

    if not AUTH_TOKEN:
        AUTH_TOKEN = secrets.token_urlsafe(24)
        logger.warning(
            "No token supplied. Generated a temporary token: %s", AUTH_TOKEN
        )

    logger.info("Starting AllOneTools Print Server on %s:%s", args.host, args.port)
    logger.info("Authentication token: %s", AUTH_TOKEN)

    atexit.register(shutdown_service)

    try:
        app.run(host=args.host, port=args.port, use_reloader=False)
    finally:
        shutdown_service()


if __name__ == "__main__":  # pragma: no cover - CLI entry
    main()

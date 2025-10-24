"""Example client for the AllOneTools Print Server REST API."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List

import requests


def make_headers(token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {token}"}


def check_status(base_url: str, token: str) -> Dict[str, object]:
    response = requests.get(f"{base_url}/status", headers=make_headers(token), timeout=10)
    response.raise_for_status()
    return response.json()


def enable_sharing(base_url: str, token: str) -> Dict[str, object]:
    response = requests.post(f"{base_url}/enable", headers=make_headers(token), timeout=10)
    response.raise_for_status()
    return response.json()


def list_printers(base_url: str, token: str) -> List[Dict[str, str]]:
    response = requests.get(f"{base_url}/printers", headers=make_headers(token), timeout=10)
    response.raise_for_status()
    payload = response.json()
    return payload.get("printers", [])


def send_print_job(base_url: str, token: str, printer_name: str, file_path: Path) -> Dict[str, object]:
    with file_path.open("rb") as handle:
        files = {"file": (file_path.name, handle)}
        data = {"printer_name": printer_name}
        response = requests.post(
            f"{base_url}/print",
            headers=make_headers(token),
            data=data,
            files=files,
            timeout=30,
        )
    response.raise_for_status()
    return response.json()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Interact with the AllOneTools Print Server")
    parser.add_argument("--url", default="http://127.0.0.1:5151", help="Base URL of the print server")
    parser.add_argument("--token", required=True, help="Bearer token configured on the server")
    parser.add_argument(
        "--printer",
        help="Printer name to target when sending a job (use the value shown in /printers)",
    )
    parser.add_argument("--file", type=Path, help="File to send to the printer")
    parser.add_argument(
        "--enable", action="store_true", help="Enable sharing before listing printers or printing"
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    base_url = args.url.rstrip("/")

    if args.enable:
        enable_response = enable_sharing(base_url, args.token)
        print("Enable response:", enable_response)

    status = check_status(base_url, args.token)
    print("Server status:", status)

    printers = list_printers(base_url, args.token)
    if not printers:
        print("No printers reported by the server.")
    else:
        print("Available printers:")
        for printer in printers:
            name = printer.get("name")
            comment = printer.get("comment") or printer.get("info") or ""
            location = printer.get("location", "")
            raw_name = printer.get("raw_name") or ""
            hostname = printer.get("hostname") or ""

            details = []
            if raw_name and raw_name != name:
                details.append(f"raw={raw_name}")
            if hostname:
                details.append(f"host={hostname}")
            if comment:
                details.append(comment)
            if location:
                details.append(location)

            extra = "; ".join(filter(None, details))
            print(f"  - {name}{' (' + extra + ')' if extra else ''}")

    if args.printer and args.file:
        if not args.file.exists():
            raise SystemExit(f"File not found: {args.file}")
        response = send_print_job(base_url, args.token, args.printer, args.file)
        print("Print response:", response)
    elif args.printer or args.file:
        print("Both --printer and --file must be provided to send a job.")


if __name__ == "__main__":
    main()

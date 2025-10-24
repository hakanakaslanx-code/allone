# AllOneTools Print Server

A lightweight Flask service that shares locally attached printers (including the
USB Dymo LabelWriter 450) with other computers on the same Wi-Fi/LAN. The
service exposes a secured REST API, enforces Bearer token authentication, and
announces each shared printer over Zeroconf using the pattern
`{printer_name} - {hostname}` so that clients on the network can discover
uniquely named devices automatically.

## Features

- Unified REST API with endpoints to list printers, toggle sharing, check
  status, and submit print jobs.
- Works on Windows via `pywin32` (`win32print`) and on Linux/macOS via
  `pycups`.
- Only accepts requests from private/loopback network addresses.
- Mandatory `Authorization: Bearer <token>` header on every request.
- Zeroconf/mDNS advertisement as `_printer._tcp.local.` with names formatted as
  `{printer_name} - {hostname}` when sharing is enabled.
- Detailed logging and descriptive error responses for easier troubleshooting.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\\Scripts\\activate
pip install -r requirements.txt
```

## Running the server

```bash
python server.py --host 0.0.0.0 --port 5151 --token YOUR_SECURE_TOKEN
```

If you omit `--token` the server generates a random token at startup (logged to
the console). Use the same token inside the desktop app (AllOneTools → Print
Server tab) and in API clients.

### Environment variables

- `PRINT_SERVER_TOKEN`: default token if `--token` is not provided.
- `PRINT_SERVER_LOG_LEVEL`: override the logging level (e.g. `DEBUG`).
- `PRINT_SERVER_MAX_BYTES`: optional maximum request size in bytes (defaults to
  20 MB) when set before launching the app.

### Zeroconf advertisement

The server advertises every locally attached printer using its real system
name combined with the host computer name. For example, a Dymo printer on the
machine `social` is announced as `Dymo LabelWriter 450 - social` under the
`_printer._tcp.local.` service type. mDNS browsers and other discovery clients
can enumerate these names to differentiate identical printer models connected to
different machines on the same network.

## API overview

All endpoints require the `Authorization: Bearer <token>` header and must be
invoked from the same local network.

| Method | Endpoint   | Description                                          |
| ------ | ---------- | ---------------------------------------------------- |
| GET    | `/status`  | Returns sharing state and backend availability.      |
| GET    | `/printers` | Lists printers detected by the host system with names including the host computer. |
| POST   | `/enable`  | Enables sharing and registers the Zeroconf service.  |
| POST   | `/disable` | Disables sharing and unregisters Zeroconf.           |
| POST   | `/print`   | Sends a print job to `printer_name`.                 |

### `/print` payloads

Submit either a `multipart/form-data` request with `file` and `printer_name`
fields, or a JSON body:

```json
{
  "printer_name": "DYMO LabelWriter 450 - social",
  "data_base64": "..."  // optional base64-encoded raw data
}
```

Alternatively, use the `text` property to print UTF-8 strings without base64.

## Example client

The repository contains [`client_example.py`](client_example.py) showing how to
enable sharing, list printers, and send a file:

```bash
python client_example.py --url http://192.168.1.10:5151 --token YOUR_TOKEN --enable
python client_example.py --url http://192.168.1.10:5151 --token YOUR_TOKEN
python client_example.py --url http://192.168.1.10:5151 --token YOUR_TOKEN \
    --printer "DYMO LabelWriter 450 - social" --file sample_label.pdf
```

## Troubleshooting

- Ensure the service is running on the machine with the USB printer connected.
- The firewall must allow inbound traffic on the chosen port.
- On Linux/macOS install the CUPS development headers before `pycups`:
  `sudo apt install libcups2-dev` (Debian/Ubuntu) or the equivalent package.
- On Windows install the latest Dymo drivers and ensure the printer appears in
  the system printer list.

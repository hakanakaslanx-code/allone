"""Google Sheets synchronisation utilities for RugBase."""

from __future__ import annotations

import json
import os
import shutil
import sqlite3
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from ..paths import (
    get_backups_dir,
    get_lock_path,
    get_service_account_path,
    get_token_path,
)

EXPECTED_CLIENT_EMAIL = "rugbase-sync@rugbase-sync.iam.gserviceaccount.com"
DEFAULT_SHEET_ID = "1n6_7L-8fPtQBN_QodxBXj3ZMzOPpMzdx8tpdRZZe5F8"
SCOPES = (
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
)
ITEM_COLUMNS = [
    "id",
    "rug_no",
    "sku",
    "title",
    "collection",
    "size",
    "price",
    "qty",
    "updated_at",
]


class GoogleSheetsSyncError(RuntimeError):
    """Raised when synchronisation fails."""


@dataclass
class SyncState:
    local_latest: Optional[str] = None
    remote_latest: Optional[str] = None
    last_sync: Optional[str] = None

    @classmethod
    def load(cls) -> "SyncState":
        path = get_token_path()
        try:
            with path.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
        except FileNotFoundError:
            return cls()
        except json.JSONDecodeError:
            return cls()
        return cls(
            local_latest=data.get("local_latest"),
            remote_latest=data.get("remote_latest"),
            last_sync=data.get("last_sync"),
        )

    def save(self) -> None:
        path = get_token_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as handle:
            json.dump(
                {
                    "local_latest": self.local_latest,
                    "remote_latest": self.remote_latest,
                    "last_sync": self.last_sync,
                },
                handle,
                indent=2,
            )


class GoogleSheetsSync:
    """Coordinator for synchronising the local database with Google Sheets."""

    def __init__(
        self,
        sheet_id: str = DEFAULT_SHEET_ID,
        *,
        db_path: Optional[Path] = None,
        status_callback: Optional[Callable[[str], None]] = None,
        logger: Optional[Callable[[str], None]] = None,
    ) -> None:
        self.sheet_id = sheet_id
        self.db_path = Path(db_path or "rugbase.db").resolve()
        self._status_callback = status_callback
        self._logger = logger
        self._service_account_info: Optional[Dict[str, str]] = None
        self._credentials: Optional[service_account.Credentials] = None
        self._sheets_service = None
        self._sheet_title: Optional[str] = None
        self._state = SyncState.load()
        self._state_lock = threading.RLock()
        self._batch_size = 5000

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------
    def test_connection(self) -> str:
        """Validate the configured Google Sheets credentials."""
        self._set_status("Google Sheets bağlantısı test ediliyor…")
        with self._synchronisation_lock():
            credentials = self._get_credentials()
            if credentials.service_account_email != EXPECTED_CLIENT_EMAIL:
                raise GoogleSheetsSyncError(
                    "Sheets bağlantısı başarısız: client_email eşleşmiyor."
                )
            service = self._get_sheets_service()
            metadata = self._execute_request(
                service.spreadsheets().get(
                    spreadsheetId=self.sheet_id, includeGridData=False
                )
            )
            sheets = metadata.get("sheets", [])
            if not sheets:
                raise GoogleSheetsSyncError("Sheets bağlantısı başarısız: sayfa bulunamadı.")
            self._sheet_title = sheets[0]["properties"].get("title", "Sheet1")
            self._set_status(f"Sheet '{self._sheet_title}' bağlantısı başarılı.")
            return self._sheet_title

    def upload_database(self) -> None:
        """Upload the contents of the items table to Google Sheets."""
        with self._synchronisation_lock():
            self._ensure_ready()
            self._set_status("Sheets'e veri yükleniyor…")
            service = self._get_sheets_service()
            sheet_title = self._get_sheet_title()
            self._ensure_headers(service, sheet_title)
            self._clear_sheet_body(service, sheet_title)
            last_col = self._column_letter(len(ITEM_COLUMNS))
            next_row = 2
            for batch in self._iter_database_rows():
                if not batch:
                    continue
                end_row = next_row + len(batch) - 1
                cell_range = f"{sheet_title}!A{next_row}:{last_col}{end_row}"
                self._update_values(service, cell_range, batch)
                next_row = end_row + 1
            self._set_status("Sheets'e veri yükleme tamamlandı.")
            self._refresh_state(after_upload=True)

    def download_to_database(self) -> Tuple[int, int]:
        """Pull rows from Google Sheets and upsert them into the database."""
        with self._synchronisation_lock():
            self._ensure_ready()
            self._set_status("Sheets'ten veriler indiriliyor…")
            service = self._get_sheets_service()
            sheet_title = self._get_sheet_title()
            values = self._execute_request(
                service.spreadsheets().values().get(
                    spreadsheetId=self.sheet_id, range=sheet_title
                )
            ).get("values", [])
            if not values:
                self._set_status("Sheets'te veri bulunamadı.")
                return (0, 0)
            header, rows = values[0], values[1:]
            header_map = {key: idx for idx, key in enumerate(header)}
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            inserted = 0
            updated = 0
            conflicts: List[Dict[str, Dict[str, str]]] = []
            try:
                for raw in rows:
                    record = self._row_to_record(header_map, raw)
                    if record is None:
                        continue
                    local = self._fetch_local_row(conn, record["id"])
                    if local is None:
                        self._insert_row(conn, record)
                        inserted += 1
                        continue
                    outcome = self._compare_rows(local, record)
                    if outcome == "remote":
                        self._update_row(conn, record)
                        updated += 1
                        conflicts.append({"id": record["id"], "winner": "remote", "local": dict(local), "remote": record})
                    elif outcome == "local":
                        conflicts.append({"id": record["id"], "winner": "local", "local": dict(local), "remote": record})
                conn.commit()
            finally:
                conn.close()
            if conflicts:
                self._write_conflicts(conflicts)
            self._set_status("Sheets verileri yerel veritabanına işlendi.")
            self._refresh_state(after_download=True)
            return (inserted, updated)

    def bidirectional_sync(self) -> Tuple[int, int]:
        """Perform a full two-way sync."""
        inserted, updated = self.download_to_database()
        self.upload_database()
        return inserted, updated

    def background_sync_cycle(self) -> bool:
        """Check for changes and run a silent sync when necessary."""
        try:
            local_latest = self._latest_local_timestamp()
            remote_latest = self._latest_remote_timestamp()
        except GoogleSheetsSyncError:
            raise
        except Exception as exc:  # pragma: no cover - defensive
            self._set_status(f"Sheets kontrolü başarısız: {exc}")
            raise
        with self._state_lock:
            state = self._state
            if state.local_latest == local_latest and state.remote_latest == remote_latest:
                return True
        try:
            self.bidirectional_sync()
            return True
        except GoogleSheetsSyncError as exc:
            self._set_status(str(exc))
            raise

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _ensure_ready(self) -> None:
        if not self.db_path.exists():
            raise GoogleSheetsSyncError(f"Veritabanı bulunamadı: {self.db_path}")
        self._get_credentials()
        self._get_sheets_service()

    def _get_credentials(self) -> service_account.Credentials:
        if self._credentials is not None:
            return self._credentials
        info = self._load_service_account_info()
        credentials = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES
        )
        self._credentials = credentials
        return credentials

    def _load_service_account_info(self) -> Dict[str, str]:
        if self._service_account_info is not None:
            return self._service_account_info
        env_value = os.environ.get("RUGBASE_SA_JSON")
        info: Optional[Dict[str, str]] = None
        if env_value:
            try:
                info = json.loads(env_value)
            except json.JSONDecodeError as exc:
                raise GoogleSheetsSyncError("Sheets bağlantısı başarısız: RUGBASE_SA_JSON geçersiz.") from exc
        else:
            sa_path = get_service_account_path()
            if sa_path.exists():
                with sa_path.open("r", encoding="utf-8") as handle:
                    info = json.load(handle)
            else:
                self._ensure_sample_credentials_exists(sa_path)
                raise GoogleSheetsSyncError(
                    f"Sheets bağlantısı başarısız: kimlik dosyası bulunamadı ({sa_path})."
                )
        if not info:
            raise GoogleSheetsSyncError("Sheets bağlantısı başarısız: kimlik bilgisi okunamadı.")
        self._service_account_info = info
        return info

    def _ensure_sample_credentials_exists(self, target: Path) -> None:
        credentials_dir = target.parent
        credentials_dir.mkdir(parents=True, exist_ok=True)
        sample_target = credentials_dir / "service_account.sample.json"
        if sample_target.exists():
            return
        package_root = Path(__file__).resolve().parents[2]
        sample_source = package_root / "resources" / "service_account.sample.json"
        if sample_source.exists():
            shutil.copy2(sample_source, sample_target)

    def _get_sheets_service(self):
        if self._sheets_service is None:
            credentials = self._get_credentials()
            self._sheets_service = build(  # pragma: no cover - network call
                "sheets", "v4", credentials=credentials, cache_discovery=False
            )
        return self._sheets_service

    def _get_sheet_title(self) -> str:
        if self._sheet_title:
            return self._sheet_title
        service = self._get_sheets_service()
        metadata = self._execute_request(
            service.spreadsheets().get(
                spreadsheetId=self.sheet_id, includeGridData=False
            )
        )
        sheets = metadata.get("sheets", [])
        if not sheets:
            raise GoogleSheetsSyncError("Sheet yapılandırması bulunamadı.")
        self._sheet_title = sheets[0]["properties"].get("title", "Sheet1")
        return self._sheet_title

    def _ensure_headers(self, service, sheet_title: str) -> None:
        last_col = self._column_letter(len(ITEM_COLUMNS))
        self._update_values(
            service,
            f"{sheet_title}!A1:{last_col}1",
            [ITEM_COLUMNS],
        )

    def _clear_sheet_body(self, service, sheet_title: str) -> None:
        last_col = self._column_letter(len(ITEM_COLUMNS))
        self._execute_request(
            service.spreadsheets().values().clear(
                spreadsheetId=self.sheet_id,
                range=f"{sheet_title}!A2:{last_col}",
            )
        )

    def _update_values(self, service, cell_range: str, values: Sequence[Sequence[object]]) -> None:
        self._execute_request(
            service.spreadsheets().values().update(
                spreadsheetId=self.sheet_id,
                range=cell_range,
                valueInputOption="RAW",
                body={"values": values},
            )
        )

    def _iter_database_rows(self) -> Iterable[List[str]]:
        conn = sqlite3.connect(self.db_path)
        try:
            cursor = conn.execute(
                "SELECT id, rug_no, sku, title, collection, size, price, qty, updated_at "
                "FROM items ORDER BY id"
            )
            while True:
                rows = cursor.fetchmany(self._batch_size)
                if not rows:
                    break
                formatted = []
                for row in rows:
                    formatted.append([
                        self._coerce_value(row[idx]) for idx in range(len(ITEM_COLUMNS))
                    ])
                yield formatted
        finally:
            conn.close()

    def _row_to_record(self, header_map: Dict[str, int], raw: Sequence[str]) -> Optional[Dict[str, str]]:
        if "id" not in header_map:
            return None
        idx = header_map["id"]
        if idx >= len(raw):
            return None
        raw_id = raw[idx]
        if not raw_id:
            return None
        try:
            record_id = int(str(raw_id).strip())
        except ValueError:
            return None
        record: Dict[str, str] = {"id": record_id}
        for column in ITEM_COLUMNS:
            if column == "id":
                continue
            value_idx = header_map.get(column)
            if value_idx is None or value_idx >= len(raw):
                record[column] = ""
            else:
                record[column] = str(raw[value_idx])
        return record

    def _fetch_local_row(self, conn: sqlite3.Connection, record_id: int) -> Optional[sqlite3.Row]:
        cursor = conn.execute(
            "SELECT id, rug_no, sku, title, collection, size, price, qty, updated_at "
            "FROM items WHERE id = ?",
            (record_id,),
        )
        return cursor.fetchone()

    def _insert_row(self, conn: sqlite3.Connection, record: Dict[str, str]) -> None:
        columns = ", ".join(ITEM_COLUMNS)
        placeholders = ", ".join(["?"] * len(ITEM_COLUMNS))
        values = [record.get(column, "") if column != "id" else record["id"] for column in ITEM_COLUMNS]
        conn.execute(
            f"INSERT INTO items ({columns}) VALUES ({placeholders})",
            values,
        )

    def _update_row(self, conn: sqlite3.Connection, record: Dict[str, str]) -> None:
        assignments = ", ".join([f"{column} = ?" for column in ITEM_COLUMNS if column != "id"])
        values = [record.get(column, "") for column in ITEM_COLUMNS if column != "id"]
        values.append(record["id"])
        conn.execute(
            f"UPDATE items SET {assignments} WHERE id = ?",
            values,
        )

    def _compare_rows(self, local: sqlite3.Row, remote: Dict[str, str]) -> str:
        local_ts = self._parse_timestamp(local["updated_at"])
        remote_ts = self._parse_timestamp(remote.get("updated_at"))
        if remote_ts and (not local_ts or remote_ts > local_ts):
            return "remote"
        if local_ts and (not remote_ts or local_ts > remote_ts):
            return "local"
        return "equal"

    def _write_conflicts(self, conflicts: List[Dict[str, Dict[str, str]]]) -> None:
        timestamp = datetime.utcnow().strftime("%Y%m%d")
        path = get_backups_dir() / f"conflicts-{timestamp}.json"
        existing: List[Dict[str, Dict[str, str]]] = []
        if path.exists():
            try:
                with path.open("r", encoding="utf-8") as handle:
                    existing = json.load(handle)
            except json.JSONDecodeError:
                existing = []
        existing.extend(conflicts)
        with path.open("w", encoding="utf-8") as handle:
            json.dump(existing, handle, indent=2, ensure_ascii=False)

    def _latest_local_timestamp(self) -> Optional[str]:
        if not self.db_path.exists():
            return None
        conn = sqlite3.connect(self.db_path)
        try:
            cursor = conn.execute("SELECT MAX(updated_at) FROM items")
            value = cursor.fetchone()[0]
            if value:
                return str(value)
            return None
        finally:
            conn.close()

    def _latest_remote_timestamp(self) -> Optional[str]:
        self._ensure_ready()
        service = self._get_sheets_service()
        sheet_title = self._get_sheet_title()
        headers = self._execute_request(
            service.spreadsheets().values().get(
                spreadsheetId=self.sheet_id, range=f"{sheet_title}!1:1"
            )
        ).get("values", [[]])
        header_row = headers[0] if headers else []
        if "updated_at" not in header_row:
            return None
        index = header_row.index("updated_at")
        column = self._column_letter(index + 1)
        result = self._execute_request(
            service.spreadsheets().values().get(
                spreadsheetId=self.sheet_id,
                range=f"{sheet_title}!{column}2:{column}",
            )
        ).get("values", [])
        latest: Optional[datetime] = None
        latest_raw: Optional[str] = None
        for row in result:
            if not row:
                continue
            candidate_raw = row[0]
            candidate_dt = self._parse_timestamp(candidate_raw)
            if candidate_dt and (latest is None or candidate_dt > latest):
                latest = candidate_dt
                latest_raw = candidate_raw
        return latest_raw

    def _refresh_state(self, *, after_upload: bool = False, after_download: bool = False) -> None:
        with self._state_lock:
            if after_download:
                self._state.local_latest = self._latest_local_timestamp()
            if after_upload:
                self._state.remote_latest = self._latest_remote_timestamp()
            now = datetime.utcnow().isoformat()
            self._state.last_sync = now
            self._state.save()

    def _parse_timestamp(self, value: Optional[str]) -> Optional[datetime]:
        if not value:
            return None
        try:
            return datetime.fromisoformat(str(value))
        except ValueError:
            return None

    def _coerce_value(self, value: object) -> str:
        if value is None:
            return ""
        if isinstance(value, datetime):
            return value.isoformat()
        return str(value)

    def _column_letter(self, index: int) -> str:
        result = ""
        while index > 0:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result or "A"

    def _execute_request(self, request):
        try:
            return request.execute()
        except HttpError as exc:  # pragma: no cover - network call
            status = getattr(exc.resp, "status", None)
            if status == 403:
                message = "Sheets bağlantısı başarısız: Sheet erişimi reddedildi."
            elif status == 404:
                message = "Sheets bağlantısı başarısız: sayfa bulunamadı."
            else:
                message = f"Sheets API hatası: {exc}"
            raise GoogleSheetsSyncError(message) from exc

    @contextmanager
    def _synchronisation_lock(self, timeout: float = 10.0):
        lock_path = get_lock_path()
        start = time.monotonic()
        acquired = False
        while time.monotonic() - start < timeout:
            try:
                fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, str(os.getpid()).encode("utf-8"))
                os.close(fd)
                acquired = True
                break
            except FileExistsError:
                time.sleep(0.25)
        if not acquired:
            raise GoogleSheetsSyncError("Başka bir senkronizasyon devam ediyor.")
        try:
            yield
        finally:
            try:
                os.remove(lock_path)
            except FileNotFoundError:
                pass

    def _set_status(self, message: str) -> None:
        if self._logger:
            self._logger(message)
        if self._status_callback:
            self._status_callback(message)

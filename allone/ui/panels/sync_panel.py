"""UI panel for configuring Google Sheets synchronisation."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import messagebox, ttk

from core.backup import create_database_backup
from core.sync.google_sheets_sync import (
    DEFAULT_SHEET_ID,
    EXPECTED_CLIENT_EMAIL,
    GoogleSheetsSync,
    GoogleSheetsSyncError,
)
from settings_manager import save_settings


class SyncPanel(ttk.Frame):
    """Synchronisation settings embedded into the main settings view."""

    def __init__(self, parent: ttk.Widget, app, *, db_path: Optional[Path] = None) -> None:
        super().__init__(parent)
        self.app = app
        self.settings = app.settings.setdefault("sync", {})
        sheet_id = self.settings.get("sheet_id", DEFAULT_SHEET_ID)
        self.sheet_id_var = tk.StringVar(value=sheet_id)
        self.status_var = tk.StringVar(value=app.tr("SyncStatusIdle"))
        self._background_job: Optional[str] = None
        self._background_failures = 0
        self._background_running = False
        self._db_path = Path(db_path or self.settings.get("db_path", "rugbase.db")).resolve()

        self.sync = GoogleSheetsSync(
            sheet_id=sheet_id,
            db_path=self._db_path,
            status_callback=self._update_status_async,
            logger=self.app.log,
        )

        self._build_layout()
        self._bind_events()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_layout(self) -> None:
        padding = {"padx": 6, "pady": 4}

        sheet_label = ttk.Label(self, text=self.app.tr("Google Sheet ID"))
        sheet_label.grid(row=0, column=0, sticky="w", **padding)
        self.sheet_entry = ttk.Entry(self, textvariable=self.sheet_id_var, width=60)
        self.sheet_entry.grid(row=0, column=1, sticky="we", **padding)

        email_label = ttk.Label(self, text=self.app.tr("Service Account Email"))
        email_label.grid(row=1, column=0, sticky="w", **padding)
        email_entry = ttk.Entry(self, width=60)
        email_entry.grid(row=1, column=1, sticky="we", **padding)
        email_entry.insert(0, EXPECTED_CLIENT_EMAIL)
        email_entry.configure(state="readonly")

        status_label = ttk.Label(self, text=self.app.tr("SyncStatus"))
        status_label.grid(row=2, column=0, sticky="w", **padding)
        status_value = ttk.Label(self, textvariable=self.status_var)
        status_value.grid(row=2, column=1, sticky="w", **padding)

        button_frame = ttk.Frame(self)
        button_frame.grid(row=3, column=0, columnspan=2, sticky="w", padx=6, pady=(10, 6))

        test_button = ttk.Button(
            button_frame,
            text=self.app.tr("Test Connection"),
            command=self._handle_test_connection,
        )
        test_button.grid(row=0, column=0, padx=(0, 8))

        sync_button = ttk.Menubutton(
            button_frame,
            text=self.app.tr("Manual Sync (Upload/Download)"),
            direction="below",
        )
        sync_menu = tk.Menu(sync_button, tearoff=False)
        sync_menu.add_command(
            label=self.app.tr("Upload Local → Sheets"),
            command=self._handle_manual_upload,
        )
        sync_menu.add_command(
            label=self.app.tr("Download Sheets → Local"),
            command=self._handle_manual_download,
        )
        sync_menu.add_separator()
        sync_menu.add_command(
            label=self.app.tr("Two-way Sync"),
            command=self._handle_bidirectional_sync,
        )
        sync_button.configure(menu=sync_menu)
        sync_button.grid(row=0, column=1, padx=(0, 8))

        backup_button = ttk.Button(
            button_frame,
            text=self.app.tr("Backup Now"),
            command=self._handle_backup,
        )
        backup_button.grid(row=0, column=2)

        self.columnconfigure(1, weight=1)

    def _bind_events(self) -> None:
        def on_sheet_id_change(_event=None) -> None:
            sheet_id = self.sheet_id_var.get().strip() or DEFAULT_SHEET_ID
            if sheet_id != self.sync.sheet_id:
                self.sync.sheet_id = sheet_id
            self.settings["sheet_id"] = sheet_id
            self.app.settings["sync"] = self.settings
            self.app.log(self.app.tr("SyncSheetIdSaved").format(sheet_id=sheet_id))
            save_settings(self.app.settings)

        self.bind("<Destroy>", self._on_destroy)
        self.sheet_entry.bind("<FocusOut>", on_sheet_id_change)
        self.sheet_entry.bind("<Return>", on_sheet_id_change)

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------
    def _handle_test_connection(self) -> None:
        self._set_status(self.app.tr("SyncStatusTesting"))

        def task() -> None:
            try:
                title = self.sync.test_connection()
                self._notify_success(
                    self.app.tr("SyncTestSuccess").format(title=title)
                )
            except GoogleSheetsSyncError as exc:
                self._notify_error(str(exc))
            except Exception as exc:  # pragma: no cover - defensive
                self._notify_error(str(exc))

        self.app.run_in_thread(task)

    def _handle_manual_upload(self) -> None:
        self._set_status(self.app.tr("SyncStatusUploading"))

        def task() -> None:
            try:
                self.sync.upload_database()
                self._notify_success(self.app.tr("SyncUploadSuccess"))
            except GoogleSheetsSyncError as exc:
                self._notify_error(str(exc))

        self.app.run_in_thread(task)

    def _handle_manual_download(self) -> None:
        self._set_status(self.app.tr("SyncStatusDownloading"))

        def task() -> None:
            try:
                inserted, updated = self.sync.download_to_database()
                self._notify_success(
                    self.app.tr("SyncDownloadSuccess").format(inserted=inserted, updated=updated)
                )
            except GoogleSheetsSyncError as exc:
                self._notify_error(str(exc))

        self.app.run_in_thread(task)

    def _handle_bidirectional_sync(self) -> None:
        self._set_status(self.app.tr("SyncStatusBidirectional"))

        def task() -> None:
            try:
                inserted, updated = self.sync.bidirectional_sync()
                self._notify_success(
                    self.app.tr("SyncBidirectionalSuccess").format(inserted=inserted, updated=updated)
                )
            except GoogleSheetsSyncError as exc:
                self._notify_error(str(exc))

        self.app.run_in_thread(task)

    def _handle_backup(self) -> None:
        def task() -> None:
            try:
                backup_path, metadata = create_database_backup(self._db_path)
            except FileNotFoundError as exc:
                self._notify_error(str(exc))
                return
            message = self.app.tr("SyncBackupSuccess").format(path=backup_path)
            self._notify_success(message)
            self.app.log(
                self.app.tr("SyncBackupMetadata").format(
                    timestamp=metadata["timestamp"],
                    size=metadata["size_bytes"],
                    digest=metadata["sha256"],
                )
            )

        self.app.run_in_thread(task)

    # ------------------------------------------------------------------
    # Background sync management
    # ------------------------------------------------------------------
    def start_background_sync(self) -> None:
        self._schedule_background_check(60000)

    def _schedule_background_check(self, delay_ms: int) -> None:
        if self._background_job is not None:
            try:
                self.after_cancel(self._background_job)
            except ValueError:
                pass
        self._background_job = self.after(delay_ms, self._background_check)

    def _background_check(self) -> None:
        if self._background_running:
            return
        self._background_running = True

        def worker() -> None:
            delay = 60000
            try:
                self.sync.background_sync_cycle()
                self._background_failures = 0
                self._set_status(self.app.tr("SyncStatusIdle"))
            except GoogleSheetsSyncError as exc:
                self._background_failures += 1
                delay = min(60000 * (2 ** self._background_failures), 600000)
                self._notify_error(str(exc), toast=False)
            except Exception as exc:  # pragma: no cover - defensive
                self._background_failures += 1
                delay = min(60000 * (2 ** self._background_failures), 600000)
                self._notify_error(str(exc), toast=False)
            finally:
                def done() -> None:
                    self._background_running = False
                    self._schedule_background_check(delay)

                self.app.after(0, done)

        self.app.run_in_thread(worker)

    def _on_destroy(self, _event=None) -> None:
        if self._background_job is not None:
            try:
                self.after_cancel(self._background_job)
            except ValueError:
                pass
        self._background_job = None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _update_status_async(self, message: str) -> None:
        def apply() -> None:
            self.status_var.set(message)

        self.app.after(0, apply)

    def _set_status(self, message: str) -> None:
        self.app.after(0, lambda: self.status_var.set(message))

    def _notify_success(self, message: str) -> None:
        self._set_status(message)
        self.app.log(message)
        self.app.after(0, lambda: messagebox.showinfo(self.app.tr("Success"), message))

    def _notify_error(self, message: str, *, toast: bool = True) -> None:
        self._set_status(message)
        self.app.log(message)
        if toast:
            self.app.after(0, lambda: messagebox.showerror(self.app.tr("Error"), message))


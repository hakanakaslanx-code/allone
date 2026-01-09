from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from modules.maps_scraper.models import BusinessList
from modules.maps_scraper.scraper import build_output_filename, build_query, scrape_google_maps


class GoogleMapsScraperTab:
    def __init__(self, parent: ttk.Frame, app) -> None:
        self.app = app
        self.parent = parent
        self.thread: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.business_list: BusinessList | None = None

        self.search_term_var = tk.StringVar()
        self.location_var = tk.StringVar()
        self.max_listings_var = tk.IntVar(value=20)
        self.headless_var = tk.BooleanVar(value=False)
        self.include_socials_var = tk.BooleanVar(value=True)
        self.progress_label_var = tk.StringVar(value=self.app.tr("Ready"))

        self._build_ui()

    def _build_ui(self) -> None:
        card = self.app.create_section_card(self.parent, "Google Maps Scraper")
        body = card.body
        body.columnconfigure(1, weight=1)
        body.rowconfigure(6, weight=1)

        search_label = ttk.Label(body, text=self.app.tr("Search Term"), style="Secondary.TLabel")
        search_label.grid(
            row=0, column=0, sticky="w", padx=6, pady=6
        )
        self.app.register_widget(search_label, "Search Term")
        ttk.Entry(body, textvariable=self.search_term_var).grid(
            row=0, column=1, sticky="ew", padx=6, pady=6
        )

        location_label = ttk.Label(body, text=self.app.tr("Location"), style="Secondary.TLabel")
        location_label.grid(
            row=1, column=0, sticky="w", padx=6, pady=6
        )
        self.app.register_widget(location_label, "Location")
        ttk.Entry(body, textvariable=self.location_var).grid(
            row=1, column=1, sticky="ew", padx=6, pady=6
        )

        listings_label = ttk.Label(body, text=self.app.tr("Max Listings"), style="Secondary.TLabel")
        listings_label.grid(
            row=2, column=0, sticky="w", padx=6, pady=6
        )
        self.app.register_widget(listings_label, "Max Listings")
        ttk.Entry(body, textvariable=self.max_listings_var, width=10).grid(
            row=2, column=1, sticky="w", padx=6, pady=6
        )

        options_frame = ttk.Frame(body, style="PanelBody.TFrame")
        options_frame.grid(row=3, column=0, columnspan=2, sticky="w", padx=6, pady=6)
        headless_check = ttk.Checkbutton(
            options_frame,
            text=self.app.tr("Headless"),
            variable=self.headless_var,
        )
        headless_check.grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.app.register_widget(headless_check, "Headless")
        socials_check = ttk.Checkbutton(
            options_frame,
            text=self.app.tr("Include Socials/Email"),
            variable=self.include_socials_var,
        )
        socials_check.grid(row=0, column=1, sticky="w")
        self.app.register_widget(socials_check, "Include Socials/Email")

        button_frame = ttk.Frame(body, style="PanelBody.TFrame")
        button_frame.grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=6)

        self.start_button = ttk.Button(button_frame, text=self.app.tr("Start"), command=self.start_scrape)
        self.start_button.grid(row=0, column=0, padx=(0, 8))
        self.app.register_widget(self.start_button, "Start")

        self.stop_button = ttk.Button(button_frame, text=self.app.tr("Stop/Cancel"), command=self.stop_scrape)
        self.stop_button.grid(row=0, column=1, padx=(0, 8))
        self.app.register_widget(self.stop_button, "Stop/Cancel")

        self.export_excel_button = ttk.Button(
            button_frame,
            text=self.app.tr("Export Excel"),
            command=self.export_excel,
        )
        self.export_excel_button.grid(row=0, column=2, padx=(0, 8))
        self.app.register_widget(self.export_excel_button, "Export Excel")

        self.export_csv_button = ttk.Button(
            button_frame,
            text=self.app.tr("Export CSV"),
            command=self.export_csv,
        )
        self.export_csv_button.grid(row=0, column=3)
        self.app.register_widget(self.export_csv_button, "Export CSV")

        progress_frame = ttk.Frame(body, style="PanelBody.TFrame")
        progress_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 0))
        progress_frame.columnconfigure(0, weight=1)
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        ttk.Label(progress_frame, textvariable=self.progress_label_var, style="Secondary.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )

        self.log_text = ScrolledText(body, height=8)
        self.log_text.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=6, pady=6)
        self.log_text.config(state=tk.DISABLED)

    def _set_progress(self, value: int, total: int) -> None:
        def apply() -> None:
            self.progress_bar.config(maximum=max(total, 1))
            self.progress_bar["value"] = value
            self.progress_label_var.set(self.app.tr("Progress").format(value=value, total=total))

        if threading.current_thread() is threading.main_thread():
            apply()
        else:
            self.progress_bar.after(0, apply)

    def log(self, message: str) -> None:
        if message is None:
            return

        def append() -> None:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, str(message) + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

        if threading.current_thread() is threading.main_thread():
            append()
        else:
            self.log_text.after(0, append)

    def start_scrape(self) -> None:
        if self.thread and self.thread.is_alive():
            self.log(self.app.tr("Scraper already running."))
            return

        search_term = self.search_term_var.get().strip()
        if not search_term:
            self.log(self.app.tr("Search term is required."))
            return

        self.stop_event.clear()
        self.business_list = BusinessList()
        self.progress_bar["value"] = 0
        self.progress_label_var.set(self.app.tr("Starting scrape..."))

        def worker() -> None:
            self.log(self.app.tr("Starting Google Maps scrape..."))
            try:
                self.business_list = scrape_google_maps(
                    search_term=search_term,
                    location=self.location_var.get().strip(),
                    max_listings=max(1, int(self.max_listings_var.get())),
                    headless=self.headless_var.get(),
                    include_socials=self.include_socials_var.get(),
                    log=self.log,
                    progress=self._set_progress,
                    stop_event=self.stop_event,
                )
            finally:
                self._set_progress(
                    len(self.business_list.business_list) if self.business_list else 0,
                    max(1, int(self.max_listings_var.get())),
                )

        self.thread = threading.Thread(target=worker, daemon=True)
        self.thread.start()

    def stop_scrape(self) -> None:
        if self.thread and self.thread.is_alive():
            self.stop_event.set()
            self.log(self.app.tr("Stop requested."))
        else:
            self.log(self.app.tr("No active scrape to stop."))

    def _export(self, extension: str) -> None:
        if not self.business_list or not self.business_list.business_list:
            self.log(self.app.tr("No data to export."))
            return

        query = build_query(self.search_term_var.get(), self.location_var.get())
        filename = build_output_filename(query, extension)
        if extension == "xlsx":
            output_path = self.business_list.save_to_excel(filename)
        else:
            output_path = self.business_list.save_to_csv(filename)
        self.log(self.app.tr("Saved: {path}").format(path=output_path))

    def export_excel(self) -> None:
        self._export("xlsx")

    def export_csv(self) -> None:
        self._export("csv")

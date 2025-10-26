"""Wayfair export formatter widget for the AllOne desktop tool."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


class WayfairFormatter(ttk.Frame):
    """Interactive formatter that maps Excel columns to Wayfair export fields."""

    REQUIRED_FIELDS: List[str] = [
        "SKU",
        "Title",
        "Description",
        "UPC",
        "Brand",
        "Price",
        "Quantity",
        "Size",
        "Color",
        "Material",
        "Country of Origin",
    ]

    SIZE_COMBINE_LABEL = "Width × Length"

    def __init__(self, parent: tk.Widget, *args, **kwargs) -> None:
        super().__init__(parent, *args, **kwargs)

        self.file_path = tk.StringVar()
        self.status_var = tk.StringVar()
        self.missing_var = tk.StringVar()

        self._dataframe: Optional[pd.DataFrame] = None
        self._column_names: List[str] = []

        self.mapping_widgets: Dict[str, ttk.Combobox] = {}

        self.size_width_var = tk.StringVar()
        self.size_length_var = tk.StringVar()

        self._build_ui()

    # ---------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """Create all UI widgets for the formatter."""

        file_frame = ttk.LabelFrame(self, text="1. Excel Dosyası Seç")
        file_frame.pack(fill=tk.X, padx=10, pady=10)

        file_entry = ttk.Entry(file_frame, textvariable=self.file_path)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 5), pady=10)

        ttk.Button(file_frame, text="Gözat", command=self._choose_file).pack(
            side=tk.LEFT, padx=(0, 10), pady=10
        )

        columns_frame = ttk.LabelFrame(self, text="2. Sütunlar")
        columns_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.columns_list = tk.Text(columns_frame, height=6, state=tk.DISABLED)
        self.columns_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        mapping_frame = ttk.LabelFrame(self, text="3. Wayfair Alan Eşlemesi")
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        for row, field in enumerate(self.REQUIRED_FIELDS):
            ttk.Label(mapping_frame, text=field).grid(
                row=row, column=0, sticky=tk.W, padx=10, pady=5
            )

            combo = ttk.Combobox(mapping_frame, state="readonly")
            combo.grid(row=row, column=1, sticky=tk.EW, padx=10, pady=5)
            self.mapping_widgets[field] = combo

        mapping_frame.columnconfigure(1, weight=1)

        size_frame = ttk.Frame(mapping_frame)
        size_frame.grid(row=self.REQUIRED_FIELDS.index("Size"), column=2, padx=10, pady=5)

        ttk.Label(size_frame, text="Genişlik:").grid(row=0, column=0, padx=(0, 5))
        self.size_width_combo = ttk.Combobox(size_frame, state="readonly", textvariable=self.size_width_var)
        self.size_width_combo.grid(row=0, column=1, padx=(0, 10))

        ttk.Label(size_frame, text="Uzunluk:").grid(row=0, column=2, padx=(0, 5))
        self.size_length_combo = ttk.Combobox(size_frame, state="readonly", textvariable=self.size_length_var)
        self.size_length_combo.grid(row=0, column=3)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(button_frame, text="Mapping Yükle", command=self._load_mapping).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(button_frame, text="Mapping Kaydet", command=self._save_mapping).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(button_frame, text="Dosya Oluştur", command=self._generate_file).pack(
            side=tk.RIGHT, padx=5
        )

        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Label(status_frame, textvariable=self.status_var, foreground="#006400").pack(
            fill=tk.X
        )
        ttk.Label(status_frame, textvariable=self.missing_var, foreground="#8B0000").pack(
            fill=tk.X
        )

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _choose_file(self) -> None:
        """Prompt the user to choose an Excel file and load the columns."""

        path = filedialog.askopenfilename(
            title="Excel Dosyası Seç",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")],
        )
        if not path:
            return

        self.file_path.set(path)
        self._load_excel(path)

    def _load_excel(self, path: str) -> None:
        """Load the Excel file into memory and populate UI elements."""

        try:
            df = pd.read_excel(path)
        except Exception as exc:  # pragma: no cover - guarded by GUI
            messagebox.showerror("Hata", f"Excel dosyası okunamadı: {exc}")
            return

        self._dataframe = df
        self._column_names = list(df.columns.astype(str))

        self._refresh_columns_display()
        self._populate_mapping_options()

        self.status_var.set(f"{len(self._column_names)} sütun yüklendi.")
        self.missing_var.set("")

    def _refresh_columns_display(self) -> None:
        """Show the available column names in the text widget."""

        self.columns_list.configure(state=tk.NORMAL)
        self.columns_list.delete("1.0", tk.END)
        for name in self._column_names:
            self.columns_list.insert(tk.END, f"• {name}\n")
        self.columns_list.configure(state=tk.DISABLED)

    def _populate_mapping_options(self) -> None:
        """Update the combobox values with the loaded Excel columns."""

        options = [self.SIZE_COMBINE_LABEL] + self._column_names
        for field, widget in self.mapping_widgets.items():
            if field == "Size":
                widget.configure(values=options)
            else:
                widget.configure(values=self._column_names)

        self.size_width_combo.configure(values=self._column_names)
        self.size_length_combo.configure(values=self._column_names)

    def _generate_file(self) -> None:
        """Create the formatted Wayfair Excel file based on the mapping."""

        if self._dataframe is None:
            messagebox.showwarning("Uyarı", "Lütfen önce bir Excel dosyası seçin.")
            return

        mapping = {field: widget.get() for field, widget in self.mapping_widgets.items()}

        missing_fields = [field for field, column in mapping.items() if not column]

        if mapping.get("Size") == self.SIZE_COMBINE_LABEL:
            if not self.size_width_var.get() or not self.size_length_var.get():
                missing_fields.append("Size (Width/Length seçilmeli)")

        if missing_fields:
            self._update_missing_fields(missing_fields)
            messagebox.showwarning(
                "Eksik Alanlar",
                "Lütfen aşağıdaki alanlar için eşleme yapın:\n" + "\n".join(missing_fields),
            )
            return

        output_df = pd.DataFrame()

        for field, selection in mapping.items():
            if selection == self.SIZE_COMBINE_LABEL:
                width_col = self.size_width_var.get()
                length_col = self.size_length_var.get()
                output_df[field] = self._build_size_column(width_col, length_col)
            elif selection in self._dataframe.columns:
                output_df[field] = self._dataframe[selection]
            else:
                output_df[field] = ""

        missing_cells = self._detect_missing_cells(output_df)

        try:
            source_dir = Path(self.file_path.get()).parent
            output_path = source_dir / "wayfair_ready.xlsx"
            output_df.to_excel(output_path, index=False)
        except Exception as exc:  # pragma: no cover - guarded by GUI
            messagebox.showerror("Hata", f"Dosya kaydedilemedi: {exc}")
            return

        self.status_var.set(f"{output_path} oluşturuldu.")
        self._update_missing_fields(missing_cells)

        if missing_cells:
            messagebox.showwarning(
                "Eksik Zorunlu Hücreler",
                "Bazı zorunlu hücreler boş. Ayrıntılar için aşağıya bakın.",
            )
        else:
            messagebox.showinfo("Başarılı", "Wayfair formatında dosya hazır.")

    def _build_size_column(self, width_col: str, length_col: str) -> pd.Series:
        """Create the combined Size column from width and length columns."""

        width_series = self._dataframe.get(width_col, pd.Series(dtype=object))
        length_series = self._dataframe.get(length_col, pd.Series(dtype=object))

        return width_series.astype(str).str.strip() + " x " + length_series.astype(str).str.strip()

    def _detect_missing_cells(self, dataframe: pd.DataFrame) -> List[str]:
        """Detect required fields with empty values."""

        missing = []
        for field in self.REQUIRED_FIELDS:
            if dataframe[field].isna().any() or (dataframe[field] == "").any():
                missing.append(field)
        return missing

    def _update_missing_fields(self, missing_fields: List[str]) -> None:
        """Update the UI with missing required field information."""

        if missing_fields:
            details = "Eksik alanlar: " + ", ".join(sorted(set(missing_fields)))
        else:
            details = "Tüm zorunlu alanlar dolu görünüyor."
        self.missing_var.set(details)

    # ------------------------------------------------------------------
    # Mapping persistence
    # ------------------------------------------------------------------
    def _save_mapping(self) -> None:
        """Persist the current mapping configuration to a JSON file."""

        if not self.mapping_widgets:
            return

        mapping = {field: widget.get() for field, widget in self.mapping_widgets.items()}
        mapping_payload = {
            "file": self.file_path.get(),
            "mapping": mapping,
            "size_width": self.size_width_var.get(),
            "size_length": self.size_length_var.get(),
        }

        save_path = filedialog.asksaveasfilename(
            title="Mapping Kaydet",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )

        if not save_path:
            return

        try:
            with open(save_path, "w", encoding="utf-8") as handle:
                json.dump(mapping_payload, handle, ensure_ascii=False, indent=2)
        except OSError as exc:  # pragma: no cover - GUI feedback
            messagebox.showerror("Hata", f"Mapping kaydedilemedi: {exc}")
            return

        self.status_var.set(f"Mapping {os.path.basename(save_path)} olarak kaydedildi.")

    def _load_mapping(self) -> None:
        """Load a previously saved mapping from JSON."""

        load_path = filedialog.askopenfilename(
            title="Mapping Yükle",
            filetypes=[("JSON", "*.json")],
        )
        if not load_path:
            return

        try:
            with open(load_path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except (OSError, json.JSONDecodeError) as exc:  # pragma: no cover - GUI feedback
            messagebox.showerror("Hata", f"Mapping yüklenemedi: {exc}")
            return

        mapping: Dict[str, str] = payload.get("mapping", {})
        size_width = payload.get("size_width", "")
        size_length = payload.get("size_length", "")

        for field, widget in self.mapping_widgets.items():
            value = mapping.get(field, "")
            if value:
                widget.set(value)

        if size_width:
            self.size_width_var.set(size_width)
        if size_length:
            self.size_length_var.set(size_length)

        linked_file = payload.get("file")
        if linked_file and not self.file_path.get():
            self.file_path.set(linked_file)

        self.status_var.set(f"Mapping {os.path.basename(load_path)} yüklendi.")


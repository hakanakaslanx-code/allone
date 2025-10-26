"""Wayfair Export Formatter Tkinter module.

This module provides a small Tkinter based utility that allows a user to
select an Excel workbook, map the workbook's columns to the required Wayfair
fields, and export a Wayfair ready workbook.  The interface also supports
storing and loading mappings as JSON files for convenience.

The script is intentionally self contained so it can be executed directly for
testing or integrated into the broader AllOne tool suite.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


# Fields that are required by the Wayfair export template.  The interface uses
# this list directly to create the mapping widgets, so updating it will
# automatically update the UI.
REQUIRED_FIELDS: List[str] = [
    "Vendor SKU",
    "Product Name",
    "Brand",
    "Color",
    "Material",
    "Width (in)",
    "Length (in)",
    "Price",
    "Inventory Qty",
    "Country of Origin",
    "Description",
    "Image URL 1",
]


@dataclass
class MappingState:
    """Container for the mapping selections."""

    selections: Dict[str, str]

    def to_json(self) -> str:
        return json.dumps(self.selections, indent=2)

    @classmethod
    def from_json(cls, data: str) -> "MappingState":
        return cls(json.loads(data))


class WayfairExportFormatter(tk.Tk):
    """Tkinter application implementing the Wayfair export workflow."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Wayfair Export Formatter")
        self.geometry("720x600")

        self.source_path: Optional[Path] = None
        self.dataframe: Optional[pd.DataFrame] = None

        self._build_ui()

    # ------------------------------------------------------------------ UI ---
    def _build_ui(self) -> None:
        """Create and layout the Tkinter widgets used by the application."""

        top_frame = ttk.Frame(self, padding=10)
        top_frame.pack(fill=tk.X, anchor=tk.N)

        ttk.Label(top_frame, text="1. Excel dosyasını seçin:").pack(anchor=tk.W)
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(fill=tk.X, pady=(5, 10))

        self.file_label = ttk.Label(button_frame, text="Henüz dosya seçilmedi")
        self.file_label.pack(side=tk.LEFT, expand=True, fill=tk.X)

        ttk.Button(button_frame, text="Excel Seç", command=self.select_file).pack(
            side=tk.RIGHT
        )

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

        mapping_frame = ttk.Frame(self, padding=10)
        mapping_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            mapping_frame, text="2. Wayfair alanlarını Excel sütunlarıyla eşleştirin"
        ).pack(anchor=tk.W)

        self.mapping_container = ttk.Frame(mapping_frame)
        self.mapping_container.pack(fill=tk.BOTH, expand=True, pady=10)

        self.mapping_widgets: Dict[str, ttk.Combobox] = {}

        ttk.Button(mapping_frame, text="JSON Kaydet", command=self.save_mapping).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        ttk.Button(mapping_frame, text="JSON Yükle", command=self.load_mapping).pack(
            side=tk.LEFT
        )

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

        action_frame = ttk.Frame(self, padding=10)
        action_frame.pack(fill=tk.X)

        ttk.Button(
            action_frame, text="Dosya Oluştur", command=self.create_wayfair_file
        ).pack(side=tk.RIGHT)

        self.warning_var = tk.StringVar(value="")
        ttk.Label(action_frame, textvariable=self.warning_var, foreground="red").pack(
            side=tk.LEFT, expand=True, fill=tk.X
        )

    # --------------------------------------------------------------- Actions ---
    def select_file(self) -> None:
        """Prompt the user to select an Excel file and load its columns."""

        file_path = filedialog.askopenfilename(
            title="Excel dosyası seçin",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
        except Exception as exc:  # pragma: no cover - Tkinter UI feedback only
            messagebox.showerror("Dosya okunamadı", str(exc))
            return

        self.source_path = Path(file_path)
        self.dataframe = df
        self.file_label.configure(text=self.source_path.name)
        self._populate_mapping_widgets()

    def _populate_mapping_widgets(self) -> None:
        """Create a combobox for each required field using the dataframe columns."""

        for child in self.mapping_container.winfo_children():
            child.destroy()
        self.mapping_widgets.clear()

        columns = list(self.dataframe.columns) if self.dataframe is not None else []
        combobox_values = ["(Seçiniz)"] + columns

        for index, field in enumerate(REQUIRED_FIELDS):
            row = ttk.Frame(self.mapping_container)
            row.grid(row=index, column=0, sticky="ew", pady=3)
            row.columnconfigure(1, weight=1)

            ttk.Label(row, text=field).grid(row=0, column=0, sticky="w", padx=(0, 10))

            combo = ttk.Combobox(
                row,
                values=combobox_values,
                state="readonly",
            )
            combo.current(0)
            combo.grid(row=0, column=1, sticky="ew")
            self.mapping_widgets[field] = combo

    def _collect_mapping(self) -> MappingState:
        mapping: Dict[str, str] = {}
        for field, widget in self.mapping_widgets.items():
            selection = widget.get()
            if selection and selection != "(Seçiniz)":
                mapping[field] = selection
        return MappingState(mapping)

    def _validate_mapping(self, mapping: MappingState) -> List[str]:
        missing = [field for field in REQUIRED_FIELDS if field not in mapping.selections]
        return missing

    def _generate_output_dataframe(self, mapping: MappingState) -> pd.DataFrame:
        assert self.dataframe is not None

        output = pd.DataFrame()
        for field in REQUIRED_FIELDS:
            column_name = mapping.selections.get(field)
            if column_name and column_name in self.dataframe.columns:
                output[field] = self.dataframe[column_name]
            else:
                output[field] = ""

        # Automatically derive Size column from Width and Length when available.
        width_col = mapping.selections.get("Width (in)")
        length_col = mapping.selections.get("Length (in)")
        if width_col and length_col:
            width_series = self.dataframe[width_col]
            length_series = self.dataframe[length_col]
            size_series = width_series.astype(str).str.strip() + "W x " + length_series.astype(
                str
            ).str.strip() + "L"
            output["Size"] = size_series

        return output

    def create_wayfair_file(self) -> None:
        if self.dataframe is None or self.source_path is None:
            messagebox.showwarning("Dosya seçin", "Lütfen önce bir Excel dosyası seçin.")
            return

        mapping = self._collect_mapping()
        missing_fields = self._validate_mapping(mapping)
        if missing_fields:
            self.warning_var.set(
                "Eksik alanlar: " + ", ".join(sorted(missing_fields))
            )
            messagebox.showwarning(
                "Eksik alanlar",
                "Tüm zorunlu alanlar için bir eşleme yapmalısınız.",
            )
            return

        self.warning_var.set("")
        output_df = self._generate_output_dataframe(mapping)

        output_path = self.source_path.parent / "wayfair_ready.xlsx"
        try:
            output_df.to_excel(output_path, index=False)
        except Exception as exc:  # pragma: no cover - Tkinter UI feedback only
            messagebox.showerror("Kayıt hatası", str(exc))
            return

        messagebox.showinfo(
            "Başarılı",
            f"Wayfair dosyası oluşturuldu:\n{output_path}",
        )

    # ------------------------------------------------------------- JSON I/O ---
    def save_mapping(self) -> None:
        if not self.mapping_widgets:
            messagebox.showwarning(
                "Eşleme yok", "Önce bir Excel dosyası seçerek eşleme oluşturun."
            )
            return

        mapping = self._collect_mapping()
        file_path = filedialog.asksaveasfilename(
            title="Eşlemeyi JSON olarak kaydet",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as fh:
                fh.write(mapping.to_json())
        except OSError as exc:  # pragma: no cover - Tkinter UI feedback only
            messagebox.showerror("Kaydetme hatası", str(exc))
            return

        messagebox.showinfo("Kaydedildi", f"Eşleme kaydedildi:\n{file_path}")

    def load_mapping(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Eşleme JSON dosyasını seç",
            filetypes=[("JSON", "*.json"), ("All Files", "*.*")],
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as fh:
                loaded = MappingState.from_json(fh.read())
        except (OSError, json.JSONDecodeError) as exc:  # pragma: no cover
            messagebox.showerror("Yükleme hatası", str(exc))
            return

        if not self.mapping_widgets:
            messagebox.showwarning(
                "Eşleme uygulanamadı",
                "Önce bir Excel dosyası seçin, ardından eşlemeyi yükleyin.",
            )
            return

        for field, widget in self.mapping_widgets.items():
            selection = loaded.selections.get(field, "")
            if selection and selection in widget["values"]:
                widget.set(selection)
            else:
                widget.set("(Seçiniz)")

        messagebox.showinfo(
            "Yüklendi",
            "Eşleme seçimleri yüklendi. Lütfen değerleri kontrol edin.",
        )


def main() -> None:
    app = WayfairExportFormatter()
    app.mainloop()


if __name__ == "__main__":
    main()

"""Wayfair Export Formatter module for the AllOne tool.

This module provides a ``WayfairFormatter`` widget that can be embedded inside any
Tkinter/ttk container.  It allows the user to load an Excel workbook, preview its
content, map the columns to Wayfair's required fields, optionally derive a
``Size`` column from width and length selections, and export the Wayfair ready
file.  Column mappings can be stored and restored as JSON files for convenience.

The module also exposes a runnable ``main`` function which builds a minimal
Tkinter application placing the ``WayfairFormatter`` inside an ``ttk.Notebook``
(tabbed interface).
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk


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
    """Container that stores the selected mapping for the required fields."""

    selections: Dict[str, str]

    def to_json(self) -> str:
        return json.dumps(self.selections, indent=2, ensure_ascii=False)

    @classmethod
    def from_json(cls, data: str) -> "MappingState":
        return cls(json.loads(data))


class WayfairFormatter(ttk.Frame):
    """ttk frame implementing the Wayfair export workflow."""

    preview_rows: int = 25

    def __init__(self, parent: tk.Misc, *, required_fields: Iterable[str] = REQUIRED_FIELDS):
        super().__init__(parent, padding=10)

        self.required_fields: List[str] = list(required_fields)
        self.source_path: Optional[Path] = None
        self.dataframe: Optional[pd.DataFrame] = None
        self.mapping_widgets: Dict[str, ttk.Combobox] = {}

        self.warning_var = tk.StringVar(value="")
        self.derive_size_var = tk.BooleanVar(value=True)

        self._build_ui()

    # ------------------------------------------------------------------ UI ---
    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)

        file_frame = ttk.LabelFrame(self, text="1. Excel dosyasını seçin")
        file_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)

        select_btn = ttk.Button(file_frame, text="Excel Seç", command=self.select_file)
        select_btn.grid(row=0, column=0, sticky="w", pady=5)

        self.file_label = ttk.Label(file_frame, text="Henüz dosya seçilmedi")
        self.file_label.grid(row=0, column=1, sticky="w", padx=(10, 0))

        preview_frame = ttk.Frame(file_frame)
        preview_frame.grid(row=1, column=0, columnspan=2, sticky="nsew")
        file_frame.rowconfigure(1, weight=1)

        self.preview_tree = ttk.Treeview(
            preview_frame,
            show="headings",
            height=self.preview_rows,
        )
        preview_scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)
        preview_scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(xscrollcommand=preview_scroll_x.set, yscrollcommand=preview_scroll_y.set)

        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        preview_scroll_y.grid(row=0, column=1, sticky="ns")
        preview_scroll_x.grid(row=1, column=0, sticky="ew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        mapping_frame = ttk.LabelFrame(self, text="2. Wayfair alanlarını Excel sütunlarıyla eşleştirin")
        mapping_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        mapping_frame.columnconfigure(1, weight=1)

        for index, field in enumerate(self.required_fields):
            ttk.Label(mapping_frame, text=field).grid(row=index, column=0, sticky="w", padx=(0, 10), pady=3)
            combo = ttk.Combobox(mapping_frame, state="readonly", values=["(Seçiniz)"])
            combo.current(0)
            combo.grid(row=index, column=1, sticky="ew", pady=3)
            self.mapping_widgets[field] = combo

        options_frame = ttk.Frame(mapping_frame)
        options_frame.grid(row=len(self.required_fields), column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Checkbutton(
            options_frame,
            text="Width/Length alanlarından Size sütunu üret",
            variable=self.derive_size_var,
        ).grid(row=0, column=0, sticky="w")

        mapping_buttons = ttk.Frame(mapping_frame)
        mapping_buttons.grid(row=len(self.required_fields) + 1, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(mapping_buttons, text="JSON Kaydet", command=self.save_mapping).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(mapping_buttons, text="JSON Yükle", command=self.load_mapping).grid(row=0, column=1, padx=(0, 5))

        action_frame = ttk.LabelFrame(self, text="3. Wayfair dosyasını oluşturun")
        action_frame.grid(row=2, column=0, sticky="ew")
        action_frame.columnconfigure(0, weight=1)

        ttk.Button(action_frame, text="Dosya Oluştur", command=self.create_wayfair_file).grid(
            row=0, column=0, sticky="e", pady=5
        )
        ttk.Label(action_frame, textvariable=self.warning_var, foreground="red").grid(
            row=1, column=0, sticky="w"
        )

    # --------------------------------------------------------------- Helpers ---
    def _update_preview(self) -> None:
        """Refresh the preview treeview with the dataframe content."""

        for column in self.preview_tree["columns"]:
            self.preview_tree.heading(column, text="")
        self.preview_tree.delete(*self.preview_tree.get_children())

        if self.dataframe is None:
            self.preview_tree.configure(columns=())
            return

        columns = list(self.dataframe.columns)
        self.preview_tree.configure(columns=columns)

        for column in columns:
            self.preview_tree.heading(column, text=column)
            self.preview_tree.column(column, width=120, anchor="w")

        for _, row in self.dataframe.head(self.preview_rows).iterrows():
            values = ["" if pd.isna(row[col]) else str(row[col]) for col in columns]
            self.preview_tree.insert("", "end", values=values)

    def _populate_mapping_values(self) -> None:
        columns = ["(Seçiniz)"]
        if self.dataframe is not None:
            columns += list(self.dataframe.columns)
        for combo in self.mapping_widgets.values():
            current_selection = combo.get()
            combo.configure(values=columns)
            if current_selection in columns:
                combo.set(current_selection)
            else:
                combo.current(0)

    def _collect_mapping(self) -> MappingState:
        mapping: Dict[str, str] = {}
        for field, widget in self.mapping_widgets.items():
            selection = widget.get()
            if selection and selection != "(Seçiniz)":
                mapping[field] = selection
        return MappingState(mapping)

    def _generate_output_dataframe(self, mapping: MappingState) -> pd.DataFrame:
        assert self.dataframe is not None

        output = pd.DataFrame()
        for field in self.required_fields:
            column_name = mapping.selections.get(field)
            if column_name and column_name in self.dataframe.columns:
                output[field] = self.dataframe[column_name]
            else:
                output[field] = ""

        if self.derive_size_var.get():
            width_col = mapping.selections.get("Width (in)")
            length_col = mapping.selections.get("Length (in)")
            if width_col and length_col and width_col in self.dataframe.columns and length_col in self.dataframe.columns:
                width_series = self.dataframe[width_col].fillna("").astype(str).str.strip()
                length_series = self.dataframe[length_col].fillna("").astype(str).str.strip()
                output["Size"] = width_series + "W x " + length_series + "L"
        return output

    def _missing_fields_text(self, mapping: MappingState) -> str:
        missing = [field for field in self.required_fields if field not in mapping.selections]
        if not missing:
            return ""
        return ", ".join(missing)

    # --------------------------------------------------------------- Actions ---
    def select_file(self) -> None:
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
        self._update_preview()
        self._populate_mapping_values()

    def create_wayfair_file(self) -> None:
        if self.dataframe is None or self.source_path is None:
            messagebox.showwarning("Dosya seçilmedi", "Lütfen önce bir Excel dosyası seçin.")
            return

        mapping = self._collect_mapping()
        output_df = self._generate_output_dataframe(mapping)
        missing_text = self._missing_fields_text(mapping)

        output_path = self.source_path.parent / "wayfair_ready.xlsx"
        try:
            output_df.to_excel(output_path, index=False)
        except Exception as exc:  # pragma: no cover - Tkinter UI feedback only
            messagebox.showerror("Kayıt hatası", str(exc))
            return

        if missing_text:
            self.warning_var.set(f"Eksik alanlar: {missing_text}")
            messagebox.showwarning(
                "Eksik alanlar",
                "Aşağıdaki alanlar boş bırakıldı ve dosya yine de oluşturuldu:\n"
                + missing_text,
            )
        else:
            self.warning_var.set("")
            messagebox.showinfo(
                "Başarılı",
                f"Wayfair dosyası oluşturuldu:\n{output_path}",
            )

    def save_mapping(self) -> None:
        if not self.mapping_widgets:
            messagebox.showwarning(
                "Eşleme yok",
                "Önce bir Excel dosyası seçerek eşleme oluşturun.",
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
            with open(file_path, "w", encoding="utf-8") as handle:
                handle.write(mapping.to_json())
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
            with open(file_path, "r", encoding="utf-8") as handle:
                mapping = MappingState.from_json(handle.read())
        except (OSError, json.JSONDecodeError) as exc:  # pragma: no cover - Tkinter UI feedback only
            messagebox.showerror("Yükleme hatası", str(exc))
            return

        if not self.mapping_widgets:
            messagebox.showwarning(
                "Eşleme uygulanamadı",
                "Önce bir Excel dosyası seçin, ardından eşlemeyi yükleyin.",
            )
            return

        for field, widget in self.mapping_widgets.items():
            selection = mapping.selections.get(field, "")
            values = widget["values"] if isinstance(widget["values"], tuple) else tuple(widget["values"])
            if selection and selection in values:
                widget.set(selection)
            else:
                widget.current(0)

        messagebox.showinfo(
            "Yüklendi",
            "Eşleme seçimleri yüklendi. Lütfen değerleri kontrol edin.",
        )


def main() -> None:
    root = tk.Tk()
    root.title("AllOne - Excel Tools")
    root.geometry("900x700")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tab_frame = ttk.Frame(notebook)
    tab_frame.pack(fill="both", expand=True)
    notebook.add(tab_frame, text="Excel Tools")

    for index in range(2):
        tab_frame.rowconfigure(index, weight=1)
    tab_frame.columnconfigure(0, weight=1)

    formatter = WayfairFormatter(tab_frame)
    formatter.grid(row=0, column=0, sticky="nsew")

    root.mainloop()


if __name__ == "__main__":
    main()

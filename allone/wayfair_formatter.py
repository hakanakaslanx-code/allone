"""Wayfair export formatter widget for the AllOne desktop tool."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

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

    SIZE_COMBINE_LABEL_KEY = "Width × Length"
    SELECTION_PLACEHOLDER_KEY = "(Select)"

    def __init__(
        self,
        parent: tk.Widget,
        *args,
        translator: Optional[Callable[[str], str]] = None,
        **kwargs,
    ) -> None:
        super().__init__(parent, *args, **kwargs)

        self.file_path = tk.StringVar()
        self.status_var = tk.StringVar()
        self.missing_var = tk.StringVar()

        self._dataframe: Optional[pd.DataFrame] = None
        self._column_names: List[str] = []

        self.mapping_widgets: Dict[str, ttk.Combobox] = {}

        self.size_width_var = tk.StringVar()
        self.size_length_var = tk.StringVar()

        self._translator: Callable[[str], str] = translator or (lambda key: key)
        self._translatable_widgets: List[Tuple[tk.Widget, str, str]] = []
        self._status_key: Optional[str] = None
        self._status_kwargs: Dict[str, object] = {}
        self._missing_key: Optional[str] = None
        self._missing_kwargs: Dict[str, object] = {}

        self._placeholder_history: set[str] = set()
        self._size_label_history: set[str] = {self.SIZE_COMBINE_LABEL_KEY}

        self._build_ui()
        self._refresh_translations()

    # ---------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        """Create all UI widgets for the formatter."""

        file_frame = ttk.LabelFrame(self, text=self._tr("1. Choose Excel File"))
        self._register(file_frame, "text", "1. Choose Excel File")
        file_frame.pack(fill=tk.X, padx=10, pady=10)

        file_entry = ttk.Entry(file_frame, textvariable=self.file_path)
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 5), pady=10)

        browse_button = ttk.Button(file_frame, text=self._tr("Browse"), command=self._choose_file)
        browse_button.pack(
            side=tk.LEFT, padx=(0, 10), pady=10
        )
        self._register(browse_button, "text", "Browse")

        columns_frame = ttk.LabelFrame(self, text=self._tr("2. Columns"))
        self._register(columns_frame, "text", "2. Columns")
        columns_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.columns_list = tk.Text(columns_frame, height=6, state=tk.DISABLED)
        self.columns_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        mapping_frame = ttk.LabelFrame(self, text=self._tr("3. Map Wayfair Fields"))
        self._register(mapping_frame, "text", "3. Map Wayfair Fields")
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        for row, field in enumerate(self.REQUIRED_FIELDS):
            ttk.Label(mapping_frame, text=field).grid(
                row=row, column=0, sticky=tk.W, padx=10, pady=5
            )

            combo = ttk.Combobox(mapping_frame, state="readonly")
            combo.grid(row=row, column=1, sticky=tk.EW, padx=10, pady=5)
            placeholder = self._tr(self.SELECTION_PLACEHOLDER_KEY)
            combo.configure(values=[placeholder])
            combo.set(placeholder)
            self.mapping_widgets[field] = combo

        mapping_frame.columnconfigure(1, weight=1)

        size_frame = ttk.Frame(mapping_frame)
        size_frame.grid(row=self.REQUIRED_FIELDS.index("Size"), column=2, padx=10, pady=5)

        width_label = ttk.Label(size_frame, text=self._tr("Width:"))
        width_label.grid(row=0, column=0, padx=(0, 5))
        self._register(width_label, "text", "Width:")
        self.size_width_combo = ttk.Combobox(size_frame, state="readonly", textvariable=self.size_width_var)
        self.size_width_combo.grid(row=0, column=1, padx=(0, 10))

        length_label = ttk.Label(size_frame, text=self._tr("Length:"))
        length_label.grid(row=0, column=2, padx=(0, 5))
        self._register(length_label, "text", "Length:")
        self.size_length_combo = ttk.Combobox(size_frame, state="readonly", textvariable=self.size_length_var)
        self.size_length_combo.grid(row=0, column=3)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        load_button = ttk.Button(button_frame, text=self._tr("Load Mapping"), command=self._load_mapping)
        load_button.pack(
            side=tk.LEFT, padx=5
        )
        self._register(load_button, "text", "Load Mapping")
        save_button = ttk.Button(button_frame, text=self._tr("Save Mapping"), command=self._save_mapping)
        save_button.pack(
            side=tk.LEFT, padx=5
        )
        self._register(save_button, "text", "Save Mapping")
        create_button = ttk.Button(button_frame, text=self._tr("Create File"), command=self._generate_file)
        create_button.pack(
            side=tk.RIGHT, padx=5
        )
        self._register(create_button, "text", "Create File")

        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Label(status_frame, textvariable=self.status_var, foreground="#006400").pack(
            fill=tk.X
        )
        ttk.Label(status_frame, textvariable=self.missing_var, foreground="#8B0000").pack(
            fill=tk.X
        )

    # ------------------------------------------------------------------
    # Translation helpers
    # ------------------------------------------------------------------
    def _register(self, widget: tk.Widget, attr: str, key: str) -> None:
        self._translatable_widgets.append((widget, attr, key))

    def _tr(self, key: str) -> str:
        return self._translator(key) if self._translator else key

    def set_translator(self, translator: Optional[Callable[[str], str]]) -> None:
        self._translator = translator or (lambda key: key)
        self._refresh_translations()

    def _refresh_translations(self) -> None:
        for widget, attr, key in self._translatable_widgets:
            try:
                widget.configure(**{attr: self._tr(key)})
            except tk.TclError:
                continue
        placeholder = self._tr(self.SELECTION_PLACEHOLDER_KEY)
        self._placeholder_history.add(placeholder)
        size_label = self._tr(self.SIZE_COMBINE_LABEL_KEY)
        self._size_label_history.add(size_label)
        self._populate_mapping_options()
        self._refresh_status()
        self._refresh_missing_info()

    def _set_status(self, key: Optional[str], **kwargs) -> None:
        self._status_key = key
        self._status_kwargs = kwargs
        self._refresh_status()

    def _refresh_status(self) -> None:
        if not self._status_key:
            self.status_var.set("")
            return
        self.status_var.set(self._tr(self._status_key).format(**self._status_kwargs))

    def _set_missing_info(self, key: Optional[str], **kwargs) -> None:
        self._missing_key = key
        self._missing_kwargs = kwargs
        self._refresh_missing_info()

    def _refresh_missing_info(self) -> None:
        if not self._missing_key:
            self.missing_var.set("")
            return
        self.missing_var.set(self._tr(self._missing_key).format(**self._missing_kwargs))

    def _is_placeholder(self, value: str) -> bool:
        return (not value) or (value in self._placeholder_history)

    def _is_size_label(self, value: str) -> bool:
        return value in self._size_label_history

    def _format_missing_list(self, missing: List[str]) -> List[str]:
        unique_items = []
        seen = set()
        for item in missing:
            if item not in seen:
                seen.add(item)
                unique_items.append(item)
        return [self._tr(item) for item in unique_items]

    def _normalize_selection(self, value: str) -> str:
        return "" if self._is_placeholder(value) else value

    def _serialize_mapping_value(self, field: str, value: str) -> str:
        if self._is_placeholder(value):
            return ""
        if field == "Size" and self._is_size_label(value):
            return self.SIZE_COMBINE_LABEL_KEY
        return value

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _choose_file(self) -> None:
        """Prompt the user to choose an Excel file and load the columns."""

        path = filedialog.askopenfilename(
            title=self._tr("Select Excel File"),
            filetypes=[(self._tr("Excel Files"), "*.xlsx *.xlsm *.xls")],
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
            messagebox.showerror(
                self._tr("Error"),
                self._tr("Excel file could not be read: {error}").format(error=exc),
            )
            return

        self._dataframe = df
        self._column_names = list(df.columns.astype(str))

        self._refresh_columns_display()
        self._populate_mapping_options()

        self._set_status("{count} columns loaded.", count=len(self._column_names))
        self._set_missing_info(None)

    def _refresh_columns_display(self) -> None:
        """Show the available column names in the text widget."""

        self.columns_list.configure(state=tk.NORMAL)
        self.columns_list.delete("1.0", tk.END)
        for name in self._column_names:
            self.columns_list.insert(tk.END, f"• {name}\n")
        self.columns_list.configure(state=tk.DISABLED)

    def _populate_mapping_options(self) -> None:
        """Update the combobox values with the loaded Excel columns."""

        placeholder = self._tr(self.SELECTION_PLACEHOLDER_KEY)
        size_label = self._tr(self.SIZE_COMBINE_LABEL_KEY)
        base_columns = [col for col in self._column_names if isinstance(col, str)]

        for field, widget in self.mapping_widgets.items():
            current = widget.get()
            if field == "Size":
                values = [placeholder, size_label] + base_columns
            else:
                values = [placeholder] + base_columns
            widget.configure(values=values)
            if self._is_placeholder(current) or current not in values:
                widget.set(placeholder)

        combo_values = ["" if placeholder == "" else placeholder] + base_columns

        width_current = self.size_width_var.get()
        self.size_width_combo.configure(values=combo_values)
        if self._is_placeholder(width_current) or width_current not in combo_values:
            self.size_width_var.set("" if placeholder == "" else placeholder)

        length_current = self.size_length_var.get()
        self.size_length_combo.configure(values=combo_values)
        if self._is_placeholder(length_current) or length_current not in combo_values:
            self.size_length_var.set("" if placeholder == "" else placeholder)

    def _generate_file(self) -> None:
        """Create the formatted Wayfair Excel file based on the mapping."""

        if self._dataframe is None:
            messagebox.showwarning(
                self._tr("Warning"),
                self._tr("Please choose an Excel file first."),
            )
            return

        mapping: Dict[str, str] = {}
        for field, widget in self.mapping_widgets.items():
            value = widget.get()
            if self._is_placeholder(value):
                mapping[field] = ""
            elif field == "Size" and self._is_size_label(value):
                mapping[field] = self.SIZE_COMBINE_LABEL_KEY
            else:
                mapping[field] = value

        missing_fields = [field for field, column in mapping.items() if not column]

        combine_selected = mapping.get("Size") == self.SIZE_COMBINE_LABEL_KEY
        width_value = self.size_width_var.get()
        length_value = self.size_length_var.get()
        if combine_selected:
            if self._is_placeholder(width_value) or self._is_placeholder(length_value):
                missing_fields.append("Size (Width/Length must be selected)")

        if missing_fields:
            self._update_missing_fields(missing_fields)
            messagebox.showwarning(
                self._tr("Missing Fields"),
                self._tr("Please map the following fields:\n{fields}").format(
                    fields="\n".join(self._format_missing_list(missing_fields))
                ),
            )
            return

        output_df = pd.DataFrame()

        for field, selection in mapping.items():
            if selection == self.SIZE_COMBINE_LABEL_KEY:
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
            messagebox.showerror(
                self._tr("Error"),
                self._tr("File could not be saved: {error}").format(error=exc),
            )
            return

        self._set_status("Wayfair formatted file saved to {path}.", path=output_path)
        self._update_missing_fields(missing_cells)

        if missing_cells:
            messagebox.showwarning(
                self._tr("Missing Required Cells"),
                self._tr("Some required cells are empty. See below for details."),
            )
        else:
            messagebox.showinfo(
                self._tr("Success"),
                self._tr("Wayfair formatted file is ready."),
            )

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
            details = ", ".join(self._format_missing_list(missing_fields))
            self._set_missing_info("Missing fields: {fields}", fields=details)
        else:
            self._set_missing_info("All required fields look filled.")

    # ------------------------------------------------------------------
    # Mapping persistence
    # ------------------------------------------------------------------
    def _save_mapping(self) -> None:
        """Persist the current mapping configuration to a JSON file."""

        if not self.mapping_widgets:
            return

        mapping = {
            field: self._serialize_mapping_value(field, widget.get())
            for field, widget in self.mapping_widgets.items()
        }
        mapping_payload = {
            "file": self.file_path.get(),
            "mapping": mapping,
            "size_width": self._normalize_selection(self.size_width_var.get()),
            "size_length": self._normalize_selection(self.size_length_var.get()),
        }

        save_path = filedialog.asksaveasfilename(
            title=self._tr("Save Mapping"),
            defaultextension=".json",
            filetypes=[(self._tr("JSON"), "*.json")],
        )

        if not save_path:
            return

        try:
            with open(save_path, "w", encoding="utf-8") as handle:
                json.dump(mapping_payload, handle, ensure_ascii=False, indent=2)
        except OSError as exc:  # pragma: no cover - GUI feedback
            messagebox.showerror(
                self._tr("Error"),
                self._tr("Mapping could not be saved: {error}").format(error=exc),
            )
            return

        self._set_status(
            "Mapping saved as {filename}.", filename=os.path.basename(save_path)
        )

    def _load_mapping(self) -> None:
        """Load a previously saved mapping from JSON."""

        load_path = filedialog.askopenfilename(
            title=self._tr("Load Mapping"),
            filetypes=[(self._tr("JSON"), "*.json")],
        )
        if not load_path:
            return

        try:
            with open(load_path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except (OSError, json.JSONDecodeError) as exc:  # pragma: no cover - GUI feedback
            messagebox.showerror(
                self._tr("Error"),
                self._tr("Mapping could not be loaded: {error}").format(error=exc),
            )
            return

        mapping: Dict[str, str] = payload.get("mapping", {})
        size_width = payload.get("size_width", "")
        size_length = payload.get("size_length", "")

        for field, widget in self.mapping_widgets.items():
            value = mapping.get(field, "")
            if not value:
                widget.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
                continue

            if field == "Size":
                self._size_label_history.add(value)
                if self._is_size_label(value):
                    widget.set(self._tr(self.SIZE_COMBINE_LABEL_KEY))
                    continue

            current_values = widget.cget("values")
            options = list(current_values) if isinstance(current_values, tuple) else list(current_values)
            if value in options:
                widget.set(value)
            else:
                widget.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))

        if size_width and not self._is_placeholder(size_width):
            self.size_width_var.set(size_width)
        else:
            self.size_width_var.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
        if size_length and not self._is_placeholder(size_length):
            self.size_length_var.set(size_length)
        else:
            self.size_length_var.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))

        linked_file = payload.get("file")
        if linked_file and not self.file_path.get():
            self.file_path.set(linked_file)

        self._set_status("Mapping {filename} loaded.", filename=os.path.basename(load_path))


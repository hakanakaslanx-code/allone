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
        self.field_enable_vars: Dict[str, tk.BooleanVar] = {}

        self.size_width_var = tk.StringVar()
        self.size_length_var = tk.StringVar()
        self.compliance_message_var = tk.StringVar()

        self._suppress_compliance_update = False

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

        placeholder = self._tr(self.SELECTION_PLACEHOLDER_KEY)

        for row, field in enumerate(self.REQUIRED_FIELDS):
            label = ttk.Label(mapping_frame, text=field)
            label.grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)

            combo = ttk.Combobox(mapping_frame, state="disabled")
            combo.grid(row=row, column=1, sticky=tk.EW, padx=10, pady=5)
            combo.configure(values=[placeholder])
            combo.set(placeholder)
            combo.bind("<<ComboboxSelected>>", lambda _event, f=field: self._on_mapping_change(f))
            self.mapping_widgets[field] = combo

            toggle_var = tk.BooleanVar(value=False)
            self.field_enable_vars[field] = toggle_var
            toggle = ttk.Checkbutton(
                mapping_frame,
                text=self._tr("Map this field"),
                variable=toggle_var,
                command=lambda f=field: self._on_field_toggle(f),
            )
            toggle.grid(row=row, column=2, sticky=tk.W, padx=10, pady=5)
            self._register(toggle, "text", "Map this field")

        mapping_frame.columnconfigure(1, weight=1)

        size_frame = ttk.Frame(mapping_frame)
        size_frame.grid(row=self.REQUIRED_FIELDS.index("Size"), column=3, padx=10, pady=5, sticky=tk.W)

        width_label = ttk.Label(size_frame, text=self._tr("Width:"))
        width_label.grid(row=0, column=0, padx=(0, 5))
        self._register(width_label, "text", "Width:")
        self.size_width_combo = ttk.Combobox(size_frame, state="disabled", textvariable=self.size_width_var)
        self.size_width_combo.grid(row=0, column=1, padx=(0, 10))
        self.size_width_combo.bind("<<ComboboxSelected>>", lambda _event: self._update_compliance_panel())

        length_label = ttk.Label(size_frame, text=self._tr("Length:"))
        length_label.grid(row=0, column=2, padx=(0, 5))
        self._register(length_label, "text", "Length:")
        self.size_length_combo = ttk.Combobox(size_frame, state="disabled", textvariable=self.size_length_var)
        self.size_length_combo.grid(row=0, column=3)
        self.size_length_combo.bind("<<ComboboxSelected>>", lambda _event: self._update_compliance_panel())

        self.size_width_var.set(placeholder)
        self.size_length_var.set(placeholder)

        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        load_button = ttk.Button(button_frame, text=self._tr("Load Mapping"), command=self._load_mapping)
        load_button.pack(side=tk.LEFT, padx=5)
        self._register(load_button, "text", "Load Mapping")

        save_button = ttk.Button(button_frame, text=self._tr("Save Mapping"), command=self._save_mapping)
        save_button.pack(side=tk.LEFT, padx=5)
        self._register(save_button, "text", "Save Mapping")

        compliance_frame = ttk.LabelFrame(self, text=self._tr("Compliance Panel"))
        self._register(compliance_frame, "text", "Compliance Panel")
        compliance_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))

        self.compliance_message = ttk.Label(
            compliance_frame,
            textvariable=self.compliance_message_var,
            anchor=tk.W,
            wraplength=500,
        )
        self.compliance_message.pack(fill=tk.X, padx=10, pady=(8, 4))

        self.compliance_details = tk.Text(
            compliance_frame,
            height=5,
            state=tk.DISABLED,
            wrap=tk.WORD,
        )
        self.compliance_details.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        export_button_frame = ttk.Frame(compliance_frame)
        export_button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        partial_button = ttk.Button(
            export_button_frame,
            text=self._tr("Export Only Mapped Fields"),
            command=lambda: self._export_file(compliant=False),
        )
        partial_button.pack(side=tk.LEFT, padx=5)
        self._register(partial_button, "text", "Export Only Mapped Fields")

        compliant_button = ttk.Button(
            export_button_frame,
            text=self._tr("Export Compliant File"),
            command=lambda: self._export_file(compliant=True),
        )
        compliant_button.pack(side=tk.LEFT, padx=5)
        self._register(compliant_button, "text", "Export Compliant File")

        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Label(status_frame, textvariable=self.status_var, foreground="#006400").pack(
            fill=tk.X
        )
        ttk.Label(status_frame, textvariable=self.missing_var, foreground="#8B0000").pack(
            fill=tk.X
        )

        self._update_compliance_panel()

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

    # ------------------------------------------------------------------
    # Mapping helpers
    # ------------------------------------------------------------------
    def _on_field_toggle(self, field: str) -> None:
        if field not in self.mapping_widgets:
            return
        enabled = self.field_enable_vars[field].get()
        if not enabled:
            self.mapping_widgets[field].set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
        self._set_field_state(field)
        self._update_compliance_panel()

    def _on_mapping_change(self, _field: str) -> None:
        if not self._suppress_compliance_update:
            self._update_compliance_panel()

    def _set_field_state(self, field: str) -> None:
        widget = self.mapping_widgets[field]
        if self.field_enable_vars[field].get():
            widget.configure(state="readonly")
        else:
            widget.configure(state="disabled")
            widget.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
        if field == "Size":
            self._refresh_size_inputs()

    def _refresh_field_states(self) -> None:
        for field in self.mapping_widgets:
            self._set_field_state(field)

    def _refresh_size_inputs(self) -> None:
        placeholder = self._tr(self.SELECTION_PLACEHOLDER_KEY)
        size_enabled = self.field_enable_vars.get("Size")
        enabled = bool(size_enabled.get()) if size_enabled else False
        state = "readonly" if enabled else "disabled"
        self.size_width_combo.configure(state=state)
        self.size_length_combo.configure(state=state)
        if not enabled:
            self.size_width_var.set(placeholder)
            self.size_length_var.set(placeholder)

    def _collect_mapping_state(self) -> Dict[str, Dict[str, object]]:
        states: Dict[str, Dict[str, object]] = {}
        for field, widget in self.mapping_widgets.items():
            enabled = self.field_enable_vars[field].get()
            raw_value = widget.get()
            normalized = self._normalize_selection(raw_value)
            state: Dict[str, object] = {
                "enabled": enabled,
                "raw_value": raw_value,
                "selection": normalized,
                "valid": False,
                "reason": "disabled" if not enabled else "",
            }
            if not enabled:
                states[field] = state
                continue

            if not normalized:
                state["reason"] = "mapping_missing"
                states[field] = state
                continue

            if field == "Size" and normalized == self.SIZE_COMBINE_LABEL_KEY:
                width_val = self.size_width_var.get()
                length_val = self.size_length_var.get()
                if self._is_placeholder(width_val) or self._is_placeholder(length_val):
                    state["reason"] = "size_components_missing"
                else:
                    state["valid"] = True
                    state["width"] = width_val
                    state["length"] = length_val
            else:
                if normalized in self._column_names:
                    state["valid"] = True
                else:
                    state["reason"] = "mapping_missing"

            states[field] = state

        return states

    def _analyze_states(
        self, states: Dict[str, Dict[str, object]], fields: List[str]
    ) -> Tuple[List[Tuple[str, pd.Series]], List[Tuple[str, str]], List[Tuple[str, str]], List[str]]:
        included: List[Tuple[str, pd.Series]] = []
        skipped: List[Tuple[str, str]] = []
        invalid_required: List[Tuple[str, str]] = []
        empty_required: List[str] = []

        for field in fields:
            state = states.get(field, {})
            enabled = bool(state.get("enabled"))
            valid = bool(state.get("valid"))
            reason = str(state.get("reason", ""))
            if not enabled or not valid:
                reason_code = reason or ("disabled" if not enabled else "mapping_missing")
                skipped.append((field, reason_code))
                if field in self.REQUIRED_FIELDS:
                    invalid_required.append((field, reason_code))
                continue

            series = self._get_series_for_state(field, state)
            included.append((field, series))
            if (
                field in self.REQUIRED_FIELDS
                and self._dataframe is not None
                and self._series_has_missing(series)
            ):
                empty_required.append(field)

        return included, skipped, invalid_required, empty_required

    def _get_series_for_state(self, field: str, state: Dict[str, object]) -> pd.Series:
        if self._dataframe is None:
            return pd.Series(dtype=object)

        selection = state.get("selection", "")
        if field == "Size" and selection == self.SIZE_COMBINE_LABEL_KEY:
            width_col = str(state.get("width", ""))
            length_col = str(state.get("length", ""))
            return self._build_size_column(width_col, length_col)

        if selection and selection in self._dataframe.columns:
            return self._get_clean_series(str(selection))

        return pd.Series(dtype=object)

    def _get_clean_series(self, column: str) -> pd.Series:
        if self._dataframe is None:
            return pd.Series(dtype=object)
        series = self._dataframe.get(column, pd.Series(dtype=object))
        return self._clean_series(series)

    def _clean_series(self, series: pd.Series) -> pd.Series:
        if series.empty:
            return series.copy()

        def normalize(value):
            if pd.isna(value):
                return ""
            if isinstance(value, str):
                return " ".join(value.replace("\u00a0", " ").split()).strip()
            return value

        return series.apply(normalize)

    def _series_has_missing(self, series: pd.Series) -> bool:
        if series.isna().any():
            return True
        if series.dtype == object:
            return series.eq("").any()
        return False

    def _describe_reason(self, reason: str) -> str:
        mapping = {
            "disabled": self._tr("Field not selected."),
            "mapping_missing": self._tr("Mapping not selected."),
            "size_components_missing": self._tr("Width and Length must be selected."),
        }
        return mapping.get(reason, self._tr("Mapping not selected."))

    def _update_compliance_panel(self) -> None:
        if self._suppress_compliance_update:
            return

        states = self._collect_mapping_state()
        fields = list(self.mapping_widgets.keys())
        included, skipped, invalid_required, empty_required = self._analyze_states(states, fields)

        self.compliance_details.configure(state=tk.NORMAL)
        self.compliance_details.delete("1.0", tk.END)

        messages: List[str] = []

        for field, reason in invalid_required:
            messages.append(f"☐ {field} — {self._describe_reason(reason)}")

        for field in empty_required:
            messages.append(f"☐ {field} — {self._tr('Contains empty cells.')}")

        if not messages:
            if self._dataframe is None and not included:
                self.compliance_message_var.set(self._tr("Load an Excel file to evaluate compliance."))
                self.compliance_details.insert(
                    tk.END, f"• {self._tr('Enable mappings to begin validation.')}\n"
                )
            else:
                self.compliance_message_var.set(self._tr("All required fields mapped and filled."))
                for field, _series in included:
                    if field in self.REQUIRED_FIELDS:
                        self.compliance_details.insert(tk.END, f"• ☑ {field}\n")
        else:
            self.compliance_message_var.set(self._tr("Required fields needing attention:"))
            for line in messages:
                self.compliance_details.insert(tk.END, f"• {line}\n")

        self.compliance_details.configure(state=tk.DISABLED)

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

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
        self._update_compliance_panel()

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

        self._suppress_compliance_update = True
        try:
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
        finally:
            self._suppress_compliance_update = False

        self._refresh_field_states()
        self._update_compliance_panel()

    def _export_file(self, compliant: bool) -> None:
        """Export a Wayfair-ready file either partially or with full compliance."""

        if self._dataframe is None:
            messagebox.showwarning(
                self._tr("Warning"),
                self._tr("Please choose an Excel file first."),
            )
            return

        states = self._collect_mapping_state()
        fields = list(self.mapping_widgets.keys())
        included, skipped, invalid_required, empty_required = self._analyze_states(states, fields)

        if compliant:
            if invalid_required:
                self._set_status("Export cancelled: missing required mappings.")
                self._update_compliance_panel()
                return
            if empty_required:
                self._set_status("Export cancelled: required fields contain empty values.")
                self._update_compliance_panel()
                return
            ordered_fields = [
                field for field in self.REQUIRED_FIELDS if any(f == field for f, _series in included)
            ]
        else:
            ordered_fields = [field for field, _series in included]

        if not ordered_fields:
            self._set_status("No fields selected to export.")
            return

        data = {field: series for field, series in included if field in ordered_fields}
        output_df = pd.DataFrame(data, columns=ordered_fields)

        try:
            source_path = self.file_path.get()
            source_dir = Path(source_path).parent if source_path else Path.cwd()
            output_name = "wayfair_ready.xlsx" if compliant else "wayfair_partial.xlsx"
            output_path = source_dir / output_name
            output_df.to_excel(output_path, index=False)
        except Exception as exc:  # pragma: no cover - guarded by GUI
            messagebox.showerror(
                self._tr("Error"),
                self._tr("File could not be saved: {error}").format(error=exc),
            )
            return

        rows = len(output_df.index)
        mapped_fields = len(ordered_fields)
        skipped_names = [field for field, _reason in skipped if field not in ordered_fields]
        skipped_display = ", ".join(dict.fromkeys(skipped_names)) if skipped_names else self._tr("None")

        self._set_status(
            "Export summary: {rows} rows | {fields} fields exported | Skipped: {skipped}",
            rows=rows,
            fields=mapped_fields,
            skipped=skipped_display,
        )
        self._set_missing_info(None)
        messagebox.showinfo(
            self._tr("Success"),
            self._tr("Wayfair formatted file saved to {path}.").format(path=output_path),
        )
        self._update_compliance_panel()

    def _build_size_column(self, width_col: str, length_col: str) -> pd.Series:
        """Create the combined Size column from width and length columns."""

        width_series = self._get_clean_series(width_col)
        length_series = self._get_clean_series(length_col)

        if self._dataframe is not None:
            width_series = width_series.reindex(self._dataframe.index, fill_value="")
            length_series = length_series.reindex(self._dataframe.index, fill_value="")

        combined: List[str] = []
        for w, l in zip(width_series.tolist(), length_series.tolist()):
            w_text = str(w).strip()
            l_text = str(l).strip()
            if not w_text or not l_text:
                combined.append("")
            else:
                combined.append(f"{w_text} x {l_text}")

        if self._dataframe is not None:
            return pd.Series(combined, index=self._dataframe.index)
        return pd.Series(combined)

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
            "enabled_fields": {field: var.get() for field, var in self.field_enable_vars.items()},
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
        enabled_fields = payload.get("enabled_fields", {})

        self._suppress_compliance_update = True
        try:
            for field, widget in self.mapping_widgets.items():
                self.field_enable_vars[field].set(bool(enabled_fields.get(field, False)))
                value = mapping.get(field, "")
                if not value:
                    widget.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
                    self._set_field_state(field)
                    continue

                if field == "Size":
                    self._size_label_history.add(value)
                    if self._is_size_label(value):
                        widget.set(self._tr(self.SIZE_COMBINE_LABEL_KEY))
                        self._set_field_state(field)
                        continue

                current_values = widget.cget("values")
                options = (
                    list(current_values)
                    if isinstance(current_values, tuple)
                    else list(current_values)
                )
                if value in options:
                    widget.set(value)
                else:
                    widget.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
                self._set_field_state(field)

            if size_width and not self._is_placeholder(size_width):
                self.size_width_var.set(size_width)
            else:
                self.size_width_var.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
            if size_length and not self._is_placeholder(size_length):
                self.size_length_var.set(size_length)
            else:
                self.size_length_var.set(self._tr(self.SELECTION_PLACEHOLDER_KEY))
        finally:
            self._suppress_compliance_update = False

        self._refresh_field_states()
        self._update_compliance_panel()

        linked_file = payload.get("file")
        if linked_file and not self.file_path.get():
            self.file_path.set(linked_file)

        self._set_status("Mapping {filename} loaded.", filename=os.path.basename(load_path))


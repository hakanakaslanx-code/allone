# app_ui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os

from settings_manager import load_settings, save_settings
from updater import check_for_updates
import backend_logic as backend
from i18n import (
    LANGUAGE_DISPLAY,
    display_for_language,
    sanitize_language,
    translate,
    validate_translations,
)

__version__ = "3.4.4"

DYMO_LABELS = {
    'Address (30252)': {'w_in': 3.5, 'h_in': 1.125},
    'Shipping (30256)': {'w_in': 4.0, 'h_in': 2.3125},
    'Small Multipurpose (30336)': {'w_in': 2.125, 'h_in': 1.0},
    'File Folder (30258)': {'w_in': 3.5, 'h_in': 0.5625},
}


class ToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        initial_lang = sanitize_language(self.settings.get("language", "en"))
        self.language_var = tk.StringVar(value=initial_lang)
        self.language_var.trace_add("write", self.on_language_change)
        self.language_display = tk.StringVar(value=display_for_language(initial_lang))
        self.translatables = []

        self.title(self.t("app_title", version=__version__))
        self.geometry("900x750")

        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=10, pady=(10, 0))
        language_label = ttk.Label(header_frame, text=self.t("language_label"))
        language_label.pack(side="left")
        self.add_translation_target(lambda text, widget=language_label: widget.config(text=text), "language_label")

        self.language_combo = ttk.Combobox(
            header_frame,
            textvariable=self.language_display,
            values=list(LANGUAGE_DISPLAY.values()),
            state="readonly",
            width=12,
        )
        self.language_combo.pack(side="left", padx=(5, 0))
        self.language_combo.bind("<<ComboboxSelected>>", self.on_language_combo_selected)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)

        self.create_file_image_tab()
        self.create_data_calc_tab()
        self.create_code_gen_tab()
        self.create_about_tab()

        self.log_area = ScrolledText(self, height=12)
        self.log_area.pack(pady=10, padx=10, fill="both", expand=True)
        self.log_area.config(state=tk.DISABLED)

        self.apply_language()
        self.log_translation_issues()
        self.log(self.t("welcome_log"))

        self.run_in_thread(check_for_updates, self, self.log, __version__, silent=True)

    def t(self, key, **kwargs):
        return translate(self.language_var.get(), key, **kwargs)

    def log_translation_issues(self):
        translation_issues = validate_translations()
        messages = []
        for lang, data in translation_issues.items():
            parts = []
            if data["missing"]:
                parts.append(f"missing keys: {', '.join(data['missing'])}")
            if data["extra"]:
                parts.append(f"extra keys: {', '.join(data['extra'])}")
            if parts:
                message = f"⚠️ Translation issue for '{lang}': {'; '.join(parts)}"
                self.log(message)
                messages.append(message)

        if messages:
            messagebox.showwarning(
                title=self.t("warning_title"),
                message="\n".join(messages),
            )

    def add_translation_target(self, setter, key, fmt=None):
        self.translatables.append((setter, key, fmt))
        if fmt:
            setter(self.t(key, **fmt))
        else:
            setter(self.t(key))

    def apply_language(self):
        self.title(self.t("app_title", version=__version__))
        for setter, key, fmt in self.translatables:
            if fmt:
                setter(self.t(key, **fmt))
            else:
                setter(self.t(key))
        if hasattr(self, "help_text_area"):
            self.help_text_area.config(state=tk.NORMAL)
            self.help_text_area.delete("1.0", tk.END)
            self.help_text_area.insert(tk.END, self.t("help_content", version=__version__))
            self.help_text_area.config(state=tk.DISABLED)
        self.language_display.set(display_for_language(self.language_var.get()))

    def on_language_combo_selected(self, _event):
        selected = self.language_display.get()
        for code, name in LANGUAGE_DISPLAY.items():
            if name == selected and self.language_var.get() != code:
                self.language_var.set(code)
                break

    def on_language_change(self, *_):
        lang = sanitize_language(self.language_var.get())
        if lang != self.language_var.get():
            self.language_var.set(lang)
            return
        self.settings["language"] = lang
        save_settings(self.settings)
        self.apply_language()
        self.log(self.t("language_changed_log", language=display_for_language(lang)))

    def log(self, message):
        if not isinstance(message, str):
            message = str(message)

        self.log_area.config(state=tk.NORMAL)

        if '\r' in message:
            self.log_area.delete("end-1l", "end")
            self.log_area.insert(tk.END, message.replace('\r', ''))
        else:
            self.log_area.insert(tk.END, message + "\n")

        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)

    def run_in_thread(self, target_func, *args, **kwargs):
        thread = threading.Thread(target=target_func, args=args, kwargs=kwargs)
        thread.daemon = True
        thread.start()

    def task_completion_popup(self, title, message):
        title_map = {
            "Success": self.t("success_title"),
            "Error": self.t("error_title"),
            "Complete": self.t("complete_title"),
        }
        translated_title = title_map.get(title, title)
        self.after(0, messagebox.showinfo, translated_title, message)

    def create_file_image_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=self.t("tab_file_image"))
        self.add_translation_target(lambda text, nb=self.notebook, frame=tab: nb.tab(frame, text=text), "tab_file_image")

        file_ops_frame = ttk.LabelFrame(tab, text=self.t("section_copy_move"))
        file_ops_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=file_ops_frame: widget.config(text=text), "section_copy_move")

        self.source_folder = tk.StringVar(value=self.settings.get("source_folder", ""))
        self.target_folder = tk.StringVar(value=self.settings.get("target_folder", ""))
        self.numbers_file = tk.StringVar()

        source_label = ttk.Label(file_ops_frame, text=self.t("label_source_folder"))
        source_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=source_label: widget.config(text=text), "label_source_folder")

        ttk.Entry(file_ops_frame, textvariable=self.source_folder, width=60).grid(row=0, column=1, padx=5, pady=5)
        browse_src_btn = ttk.Button(
            file_ops_frame,
            text=self.t("btn_browse"),
            command=lambda: self.source_folder.set(filedialog.askdirectory()),
        )
        browse_src_btn.grid(row=0, column=2, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=browse_src_btn: widget.config(text=text), "btn_browse")

        target_label = ttk.Label(file_ops_frame, text=self.t("label_target_folder"))
        target_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=target_label: widget.config(text=text), "label_target_folder")

        ttk.Entry(file_ops_frame, textvariable=self.target_folder, width=60).grid(row=1, column=1, padx=5, pady=5)
        browse_tgt_btn = ttk.Button(
            file_ops_frame,
            text=self.t("btn_browse"),
            command=lambda: self.target_folder.set(filedialog.askdirectory()),
        )
        browse_tgt_btn.grid(row=1, column=2, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=browse_tgt_btn: widget.config(text=text), "btn_browse")

        numbers_label = ttk.Label(file_ops_frame, text=self.t("label_numbers_file"))
        numbers_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=numbers_label: widget.config(text=text), "label_numbers_file")

        ttk.Entry(file_ops_frame, textvariable=self.numbers_file, width=60).grid(row=2, column=1, padx=5, pady=5)
        browse_numbers_btn = ttk.Button(
            file_ops_frame,
            text=self.t("btn_browse"),
            command=lambda: self.numbers_file.set(filedialog.askopenfilename()),
        )
        browse_numbers_btn.grid(row=2, column=2, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=browse_numbers_btn: widget.config(text=text), "btn_browse")

        btn_frame = ttk.Frame(file_ops_frame)
        btn_frame.grid(row=3, column=1, pady=10)

        copy_btn = ttk.Button(btn_frame, text=self.t("btn_copy_files"), command=lambda: self.start_process_files("copy"))
        copy_btn.pack(side="left", padx=5)
        self.add_translation_target(lambda text, widget=copy_btn: widget.config(text=text), "btn_copy_files")

        move_btn = ttk.Button(btn_frame, text=self.t("btn_move_files"), command=lambda: self.start_process_files("move"))
        move_btn.pack(side="left", padx=5)
        self.add_translation_target(lambda text, widget=move_btn: widget.config(text=text), "btn_move_files")

        save_btn = ttk.Button(btn_frame, text=self.t("btn_save_settings"), command=self.save_folder_settings)
        save_btn.pack(side="left", padx=5)
        self.add_translation_target(lambda text, widget=save_btn: widget.config(text=text), "btn_save_settings")

        heic_frame = ttk.LabelFrame(tab, text=self.t("section_heic"))
        heic_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=heic_frame: widget.config(text=text), "section_heic")

        heic_label = ttk.Label(heic_frame, text=self.t("label_heic_folder"))
        heic_label.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=heic_label: widget.config(text=text), "label_heic_folder")

        self.heic_folder = tk.StringVar()
        ttk.Entry(heic_frame, textvariable=self.heic_folder, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        browse_heic_btn = ttk.Button(
            heic_frame,
            text=self.t("btn_browse"),
            command=lambda: self.heic_folder.set(filedialog.askdirectory()),
        )
        browse_heic_btn.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=browse_heic_btn: widget.config(text=text), "btn_browse")

        convert_btn = ttk.Button(heic_frame, text=self.t("btn_convert"), command=self.start_heic_conversion)
        convert_btn.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=convert_btn: widget.config(text=text), "btn_convert")

        resize_frame = ttk.LabelFrame(tab, text=self.t("section_resize"))
        resize_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=resize_frame: widget.config(text=text), "section_resize")

        self.resize_folder = tk.StringVar()
        self.quality = tk.StringVar(value="75")
        self.resize_mode = tk.StringVar(value="width")
        self.max_width = tk.StringVar(value="1920")
        self.resize_percentage = tk.StringVar(value="50")

        image_folder_label = ttk.Label(resize_frame, text=self.t("label_image_folder"))
        image_folder_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=image_folder_label: widget.config(text=text), "label_image_folder")

        ttk.Entry(resize_frame, textvariable=self.resize_folder, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="ew")
        browse_resize_btn = ttk.Button(
            resize_frame,
            text=self.t("btn_browse"),
            command=lambda: self.resize_folder.set(filedialog.askdirectory()),
        )
        browse_resize_btn.grid(row=0, column=4, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=browse_resize_btn: widget.config(text=text), "btn_browse")

        resize_mode_label = ttk.Label(resize_frame, text=self.t("label_resize_mode"))
        resize_mode_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=resize_mode_label: widget.config(text=text), "label_resize_mode")

        radio_frame = ttk.Frame(resize_frame)
        radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")

        width_radio = ttk.Radiobutton(radio_frame, text=self.t("radio_by_width"), variable=self.resize_mode, value="width", command=self.toggle_resize_mode)
        width_radio.pack(side="left")
        self.add_translation_target(lambda text, widget=width_radio: widget.config(text=text), "radio_by_width")

        percentage_radio = ttk.Radiobutton(radio_frame, text=self.t("radio_by_percentage"), variable=self.resize_mode, value="percentage", command=self.toggle_resize_mode)
        percentage_radio.pack(side="left", padx=10)
        self.add_translation_target(lambda text, widget=percentage_radio: widget.config(text=text), "radio_by_percentage")

        max_width_label = ttk.Label(resize_frame, text=self.t("label_max_width"))
        max_width_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=max_width_label: widget.config(text=text), "label_max_width")

        self.width_entry = ttk.Entry(resize_frame, textvariable=self.max_width, width=10)
        self.width_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        percentage_label = ttk.Label(resize_frame, text=self.t("label_percentage"))
        percentage_label.grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.add_translation_target(lambda text, widget=percentage_label: widget.config(text=text), "label_percentage")

        self.percentage_entry = ttk.Entry(resize_frame, textvariable=self.resize_percentage, width=10)
        self.percentage_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        quality_label = ttk.Label(resize_frame, text=self.t("label_quality"))
        quality_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=quality_label: widget.config(text=text), "label_quality")

        ttk.Entry(resize_frame, textvariable=self.quality, width=10).grid(row=3, column=1, padx=5, pady=5, sticky="w")

        resize_btn = ttk.Button(resize_frame, text=self.t("btn_resize"), command=self.start_resize_task)
        resize_btn.grid(row=4, column=1, columnspan=2, pady=10)
        self.add_translation_target(lambda text, widget=resize_btn: widget.config(text=text), "btn_resize")

        self.toggle_resize_mode()

    def toggle_resize_mode(self):
        if self.resize_mode.get() == "width":
            self.width_entry.config(state="normal")
            self.percentage_entry.config(state="disabled")
        else:
            self.width_entry.config(state="disabled")
            self.percentage_entry.config(state="normal")

    def create_data_calc_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=self.t("tab_data_calc"))
        self.add_translation_target(lambda text, nb=self.notebook, frame=tab: nb.tab(frame, text=text), "tab_data_calc")

        format_frame = ttk.LabelFrame(tab, text=self.t("section_format_numbers"))
        format_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=format_frame: widget.config(text=text), "section_format_numbers")

        self.format_file = tk.StringVar()

        format_label = ttk.Label(format_frame, text=self.t("label_format_file"))
        format_label.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=format_label: widget.config(text=text), "label_format_file")

        ttk.Entry(format_frame, textvariable=self.format_file, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        format_browse_btn = ttk.Button(
            format_frame,
            text=self.t("btn_browse"),
            command=lambda: self.format_file.set(filedialog.askopenfilename()),
        )
        format_browse_btn.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=format_browse_btn: widget.config(text=text), "btn_browse")

        format_btn = ttk.Button(format_frame, text=self.t("btn_format"), command=self.start_format_numbers)
        format_btn.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=format_btn: widget.config(text=text), "btn_format")

        single_rug_frame = ttk.LabelFrame(tab, text=self.t("section_rug_single"))
        single_rug_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=single_rug_frame: widget.config(text=text), "section_rug_single")

        self.rug_dim_input = tk.StringVar()
        self.rug_result_label = tk.StringVar()

        rug_label = ttk.Label(single_rug_frame, text=self.t("label_rug_dimension"))
        rug_label.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=rug_label: widget.config(text=text), "label_rug_dimension")

        ttk.Entry(single_rug_frame, textvariable=self.rug_dim_input, width=20).pack(side="left", padx=5, pady=5)
        rug_btn = ttk.Button(single_rug_frame, text=self.t("btn_calculate"), command=self.calculate_single_rug)
        rug_btn.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=rug_btn: widget.config(text=text), "btn_calculate")

        rug_result_display = ttk.Label(single_rug_frame, textvariable=self.rug_result_label, font=("Helvetica", 10, "bold"))
        rug_result_display.pack(side="left", padx=15, pady=5)

        bulk_rug_frame = ttk.LabelFrame(tab, text=self.t("section_rug_bulk"))
        bulk_rug_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=bulk_rug_frame: widget.config(text=text), "section_rug_bulk")

        self.bulk_rug_file = tk.StringVar()
        self.bulk_rug_col = tk.StringVar(value="Size")

        bulk_file_label = ttk.Label(bulk_rug_frame, text=self.t("label_bulk_file"))
        bulk_file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=bulk_file_label: widget.config(text=text), "label_bulk_file")

        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        bulk_browse_btn = ttk.Button(
            bulk_rug_frame,
            text=self.t("btn_browse"),
            command=lambda: self.bulk_rug_file.set(filedialog.askopenfilename()),
        )
        bulk_browse_btn.grid(row=0, column=2, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bulk_browse_btn: widget.config(text=text), "btn_browse")

        bulk_column_label = ttk.Label(bulk_rug_frame, text=self.t("label_bulk_column"))
        bulk_column_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=bulk_column_label: widget.config(text=text), "label_bulk_column")

        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_col, width=20).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        bulk_process_btn = ttk.Button(bulk_rug_frame, text=self.t("btn_process_file"), command=self.start_bulk_rug_sizer)
        bulk_process_btn.grid(row=1, column=2, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bulk_process_btn: widget.config(text=text), "btn_process_file")

        unit_frame = ttk.LabelFrame(tab, text=self.t("section_unit_converter"))
        unit_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=unit_frame: widget.config(text=text), "section_unit_converter")

        self.unit_input = tk.StringVar(value="182 cm to ft")
        self.unit_result_label = tk.StringVar()

        unit_label = ttk.Label(unit_frame, text=self.t("label_conversion"))
        unit_label.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=unit_label: widget.config(text=text), "label_conversion")

        ttk.Entry(unit_frame, textvariable=self.unit_input, width=20).pack(side="left", padx=5, pady=5)
        unit_btn = ttk.Button(unit_frame, text=self.t("btn_convert_units"), command=self.convert_units)
        unit_btn.pack(side="left", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=unit_btn: widget.config(text=text), "btn_convert_units")

        unit_result_display = ttk.Label(unit_frame, textvariable=self.unit_result_label, font=("Helvetica", 10, "bold"))
        unit_result_display.pack(side="left", padx=15, pady=5)

        image_link_frame = ttk.LabelFrame(tab, text=self.t("section_match_links"))
        image_link_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=image_link_frame: widget.config(text=text), "section_match_links")

        self.input_excel_file = tk.StringVar()
        self.image_links_file = tk.StringVar(value="image link shopify.csv")
        self.key_column = tk.StringVar(value="A")

        source_excel_label = ttk.Label(image_link_frame, text=self.t("label_source_excel"))
        source_excel_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=source_excel_label: widget.config(text=text), "label_source_excel")

        ttk.Entry(image_link_frame, textvariable=self.input_excel_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        source_excel_btn = ttk.Button(
            image_link_frame,
            text=self.t("btn_browse"),
            command=lambda: self.input_excel_file.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])),
        )
        source_excel_btn.grid(row=0, column=2, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=source_excel_btn: widget.config(text=text), "btn_browse")

        image_links_label = ttk.Label(image_link_frame, text=self.t("label_image_links"))
        image_links_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=image_links_label: widget.config(text=text), "label_image_links")

        ttk.Entry(image_link_frame, textvariable=self.image_links_file, width=50).grid(row=1, column=1, padx=5, pady=5)
        image_links_btn = ttk.Button(
            image_link_frame,
            text=self.t("btn_browse"),
            command=lambda: self.image_links_file.set(filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])),
        )
        image_links_btn.grid(row=1, column=2, padx=5, pady=5)
        self.add_translation_target(lambda text, widget=image_links_btn: widget.config(text=text), "btn_browse")

        key_column_label = ttk.Label(image_link_frame, text=self.t("label_key_column"))
        key_column_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.add_translation_target(lambda text, widget=key_column_label: widget.config(text=text), "label_key_column")

        ttk.Entry(image_link_frame, textvariable=self.key_column, width=10).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        match_links_btn = ttk.Button(image_link_frame, text=self.t("btn_match_links"), command=self.start_add_image_links)
        match_links_btn.grid(row=3, column=1, pady=10)
        self.add_translation_target(lambda text, widget=match_links_btn: widget.config(text=text), "btn_match_links")

    def create_code_gen_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=self.t("tab_code_gen"))
        self.add_translation_target(lambda text, nb=self.notebook, frame=tab: nb.tab(frame, text=text), "tab_code_gen")

        def toggle_dymo_options(output_var, combobox, entry):
            if output_var.get() == "Dymo":
                combobox.config(state="readonly")
                entry.config(state="normal")
            else:
                combobox.config(state="disabled")
                entry.config(state="disabled")

        qr_frame = ttk.LabelFrame(tab, text=self.t("section_qr"))
        qr_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=qr_frame: widget.config(text=text), "section_qr")

        self.qr_data = tk.StringVar()
        self.qr_filename = tk.StringVar(value="qrcode.png")
        self.qr_output_type = tk.StringVar(value="PNG")
        self.qr_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0])
        self.qr_bottom_text = tk.StringVar()

        qr_data_label = ttk.Label(qr_frame, text=self.t("label_data_url"))
        qr_data_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=qr_data_label: widget.config(text=text), "label_data_url")

        ttk.Entry(qr_frame, textvariable=self.qr_data, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5)

        output_type_label = ttk.Label(qr_frame, text=self.t("label_output_type"))
        output_type_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=output_type_label: widget.config(text=text), "label_output_type")

        qr_radio_frame = ttk.Frame(qr_frame)
        qr_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")

        qr_dymo_combo = ttk.Combobox(qr_frame, textvariable=self.qr_dymo_size, values=list(DYMO_LABELS.keys()), state="disabled", width=30)
        qr_bottom_entry = ttk.Entry(qr_frame, textvariable=self.qr_bottom_text, state="disabled", width=32)

        qr_png_radio = ttk.Radiobutton(
            qr_radio_frame,
            text=self.t("radio_png"),
            variable=self.qr_output_type,
            value="PNG",
            command=lambda: toggle_dymo_options(self.qr_output_type, qr_dymo_combo, qr_bottom_entry),
        )
        qr_png_radio.pack(side="left", padx=5)
        self.add_translation_target(lambda text, widget=qr_png_radio: widget.config(text=text), "radio_png")

        qr_dymo_radio = ttk.Radiobutton(
            qr_radio_frame,
            text=self.t("radio_dymo"),
            variable=self.qr_output_type,
            value="Dymo",
            command=lambda: toggle_dymo_options(self.qr_output_type, qr_dymo_combo, qr_bottom_entry),
        )
        qr_dymo_radio.pack(side="left", padx=5)
        self.add_translation_target(lambda text, widget=qr_dymo_radio: widget.config(text=text), "radio_dymo")

        dymo_size_label = ttk.Label(qr_frame, text=self.t("label_dymo_size"))
        dymo_size_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=dymo_size_label: widget.config(text=text), "label_dymo_size")

        qr_dymo_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        bottom_text_label = ttk.Label(qr_frame, text=self.t("label_bottom_text"))
        bottom_text_label.grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bottom_text_label: widget.config(text=text), "label_bottom_text")

        qr_bottom_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        filename_label = ttk.Label(qr_frame, text=self.t("label_filename"))
        filename_label.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=filename_label: widget.config(text=text), "label_filename")

        ttk.Entry(qr_frame, textvariable=self.qr_filename, width=60).grid(row=3, column=1, columnspan=3, padx=5, pady=5)

        qr_generate_btn = ttk.Button(qr_frame, text=self.t("btn_generate_qr"), command=self.start_generate_qr)
        qr_generate_btn.grid(row=4, column=1, columnspan=2, pady=10)
        self.add_translation_target(lambda text, widget=qr_generate_btn: widget.config(text=text), "btn_generate_qr")

        bc_frame = ttk.LabelFrame(tab, text=self.t("section_barcode"))
        bc_frame.pack(fill="x", padx=10, pady=10)
        self.add_translation_target(lambda text, widget=bc_frame: widget.config(text=text), "section_barcode")

        self.bc_data = tk.StringVar()
        self.bc_filename = tk.StringVar(value="barcode.png")
        self.bc_type = tk.StringVar(value='code128')
        self.bc_output_type = tk.StringVar(value="PNG")
        self.bc_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0])
        self.bc_bottom_text = tk.StringVar()

        bc_data_label = ttk.Label(bc_frame, text=self.t("label_barcode_data"))
        bc_data_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bc_data_label: widget.config(text=text), "label_barcode_data")

        ttk.Entry(bc_frame, textvariable=self.bc_data, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="w")

        bc_format_label = ttk.Label(bc_frame, text=self.t("label_format"))
        bc_format_label.grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bc_format_label: widget.config(text=text), "label_format")

        ttk.Combobox(bc_frame, textvariable=self.bc_type, values=['code39', 'code128', 'ean13', 'upca'], state="readonly", width=15).grid(row=0, column=3, padx=5, pady=5, sticky="w")

        bc_output_label = ttk.Label(bc_frame, text=self.t("label_output_type"))
        bc_output_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bc_output_label: widget.config(text=text), "label_output_type")

        bc_radio_frame = ttk.Frame(bc_frame)
        bc_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")

        bc_dymo_combo = ttk.Combobox(bc_frame, textvariable=self.bc_dymo_size, values=list(DYMO_LABELS.keys()), state="disabled", width=30)
        bc_bottom_entry = ttk.Entry(bc_frame, textvariable=self.bc_bottom_text, state="disabled", width=32)

        bc_png_radio = ttk.Radiobutton(
            bc_radio_frame,
            text=self.t("radio_png"),
            variable=self.bc_output_type,
            value="PNG",
            command=lambda: toggle_dymo_options(self.bc_output_type, bc_dymo_combo, bc_bottom_entry),
        )
        bc_png_radio.pack(side="left", padx=5)
        self.add_translation_target(lambda text, widget=bc_png_radio: widget.config(text=text), "radio_png")

        bc_dymo_radio = ttk.Radiobutton(
            bc_radio_frame,
            text=self.t("radio_dymo"),
            variable=self.bc_output_type,
            value="Dymo",
            command=lambda: toggle_dymo_options(self.bc_output_type, bc_dymo_combo, bc_bottom_entry),
        )
        bc_dymo_radio.pack(side="left", padx=5)
        self.add_translation_target(lambda text, widget=bc_dymo_radio: widget.config(text=text), "radio_dymo")

        bc_dymo_label = ttk.Label(bc_frame, text=self.t("label_dymo_size"))
        bc_dymo_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bc_dymo_label: widget.config(text=text), "label_dymo_size")

        bc_dymo_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        bc_bottom_label = ttk.Label(bc_frame, text=self.t("label_bottom_text"))
        bc_bottom_label.grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bc_bottom_label: widget.config(text=text), "label_bottom_text")

        bc_bottom_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")

        bc_filename_label = ttk.Label(bc_frame, text=self.t("label_filename"))
        bc_filename_label.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.add_translation_target(lambda text, widget=bc_filename_label: widget.config(text=text), "label_filename")

        ttk.Entry(bc_frame, textvariable=self.bc_filename, width=60).grid(row=3, column=1, columnspan=3, padx=5, pady=5)

        bc_generate_btn = ttk.Button(bc_frame, text=self.t("btn_generate_barcode"), command=self.start_generate_barcode)
        bc_generate_btn.grid(row=4, column=1, columnspan=2, pady=10)
        self.add_translation_target(lambda text, widget=bc_generate_btn: widget.config(text=text), "btn_generate_barcode")

    def create_about_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=self.t("tab_about"))
        self.add_translation_target(lambda text, nb=self.notebook, frame=tab: nb.tab(frame, text=text), "tab_about")

        top_frame = ttk.Frame(tab)
        top_frame.pack(fill="x", padx=10, pady=5)

        update_btn = ttk.Button(top_frame, text=self.t("btn_check_updates"), command=lambda: self.run_in_thread(check_for_updates, self, self.log, __version__, silent=False))
        update_btn.pack(side="left")
        self.add_translation_target(lambda text, widget=update_btn: widget.config(text=text), "btn_check_updates")

        self.help_text_area = ScrolledText(tab, wrap=tk.WORD, padx=10, pady=10, font=("Helvetica", 10))
        self.help_text_area.pack(fill="both", expand=True)
        self.help_text_area.insert(tk.END, self.t("help_content", version=__version__))
        self.help_text_area.config(state=tk.DISABLED)

    def save_folder_settings(self):
        src = self.source_folder.get()
        tgt = self.target_folder.get()
        if not src or not tgt:
            messagebox.showwarning(self.t("warning_title"), self.t("settings_empty"))
            return
        self.settings['source_folder'] = src
        self.settings['target_folder'] = tgt
        save_settings(self.settings)
        self.log(self.t("log_settings_saved"))
        messagebox.showinfo(self.t("success_title"), self.t("settings_saved_popup"))

    def start_process_files(self, action):
        src = self.source_folder.get()
        tgt = self.target_folder.get()
        nums_f = self.numbers_file.get()
        if not all([src, tgt, nums_f]):
            messagebox.showerror(self.t("error_title"), self.t("missing_required_fields"))
            return
        self.run_in_thread(backend.process_files_task, src, tgt, nums_f, action, self.log, self.task_completion_popup)

    def start_heic_conversion(self):
        folder = self.heic_folder.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror(self.t("error_title"), self.t("invalid_folder"))
            return
        self.run_in_thread(backend.convert_heic_task, folder, self.log, self.task_completion_popup)

    def start_resize_task(self):
        src_folder = self.resize_folder.get()
        if not src_folder or not os.path.isdir(src_folder):
            messagebox.showerror(self.t("error_title"), self.t("invalid_folder"))
            return
        mode = self.resize_mode.get()
        try:
            value = int(self.max_width.get()) if mode == 'width' else int(self.resize_percentage.get())
            quality = int(self.quality.get())
            if not (value > 0 and 1 <= quality <= 95):
                raise ValueError
        except ValueError:
            messagebox.showerror(self.t("error_title"), self.t("resize_value_error"))
            return
        self.run_in_thread(backend.resize_images_task, src_folder, mode, value, quality, self.log, self.task_completion_popup)

    def start_format_numbers(self):
        file_path = self.format_file.get()
        if not file_path:
            messagebox.showerror(self.t("error_title"), self.t("select_file"))
            return
        err, success_msg = backend.format_numbers_task(file_path)
        if err:
            self.log(f"Error: {err}")
            messagebox.showerror(self.t("error_title"), err)
        else:
            self.log(success_msg)
            messagebox.showinfo(self.t("success_title"), success_msg)

    def calculate_single_rug(self):
        dim_str = self.rug_dim_input.get()
        if not dim_str:
            self.rug_result_label.set(self.t("rug_enter_dimension"))
            return
        w, h = backend.size_to_inches_wh(dim_str)
        s = backend.calculate_sqft(dim_str)
        if w is not None and h is not None and s is not None:
            self.rug_result_label.set(self.t("rug_result", width=w, height=h, area=s))
        else:
            self.rug_result_label.set(self.t("rug_invalid"))

    def start_bulk_rug_sizer(self):
        path = self.bulk_rug_file.get()
        col = self.bulk_rug_col.get()
        if not path or not col:
            messagebox.showerror(self.t("error_title"), self.t("missing_file_and_column"))
            return
        self.run_in_thread(backend.bulk_rug_sizer_task, path, col, self.log, self.task_completion_popup)

    def convert_units(self):
        input_str = self.unit_input.get()
        result_str = backend.convert_units_logic(input_str)
        self.unit_result_label.set(result_str)

    def start_add_image_links(self):
        input_path = self.input_excel_file.get()
        links_path = self.image_links_file.get()
        key_col = self.key_column.get()
        if not all([input_path, links_path, key_col]):
            messagebox.showerror(self.t("error_title"), self.t("missing_inputs"))
            return
        self.run_in_thread(
            backend.add_image_links_task,
            input_path,
            links_path,
            key_col,
            self.log,
            self.task_completion_popup,
        )

    def start_generate_qr(self):
        data = self.qr_data.get()
        fname = self.qr_filename.get()
        if not data or not fname:
            messagebox.showerror(self.t("error_title"), self.t("data_filename_required"))
            return
        dymo_info = DYMO_LABELS[self.qr_dymo_size.get()] if self.qr_output_type.get() == "Dymo" else None
        log_msg, success_msg = backend.generate_qr_task(data, fname, self.qr_output_type.get(), dymo_info, self.qr_bottom_text.get())
        self.log(log_msg)
        if success_msg:
            self.task_completion_popup("Success", success_msg)
        else:
            messagebox.showerror(self.t("error_title"), log_msg)

    def start_generate_barcode(self):
        data = self.bc_data.get()
        fname = self.bc_filename.get()
        if not data or not fname:
            messagebox.showerror(self.t("error_title"), self.t("data_filename_required"))
            return
        dymo_info = DYMO_LABELS[self.bc_dymo_size.get()] if self.bc_output_type.get() == "Dymo" else None
        log_msg, success_msg = backend.generate_barcode_task(
            data,
            fname,
            self.bc_type.get(),
            self.bc_output_type.get(),
            dymo_info,
            self.bc_bottom_text.get() or data,
        )
        self.log(log_msg)
        if success_msg:
            self.task_completion_popup("Success", success_msg)
        else:
            messagebox.showerror(self.t("error_title"), log_msg)

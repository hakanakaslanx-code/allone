# app_ui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import os

# Diğer modüllerimizden gerekli fonksiyonları ve değişkenleri import ediyoruz
from settings_manager import load_settings, save_settings
from updater import check_for_updates
import backend_logic as backend

__version__ = "3.1.6"

DYMO_LABELS = {
    'Address (30252)': {'w_in': 3.5, 'h_in': 1.125},
    'Shipping (30256)': {'w_in': 4.0, 'h_in': 2.3125},
    'Small Multipurpose (30336)': {'w_in': 2.125, 'h_in': 1.0},
    'File Folder (30258)': {'w_in': 3.5, 'h_in': 0.5625},
}

class ToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Combined Utility Tool v{__version__}")
        self.geometry("900x750")

        self.settings = load_settings()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.gemini_api_key = tk.StringVar(value=self.settings.get("gemini_api_key", ""))
        self.gemini_model = None

        self.create_ai_assistant_tab()
        self.create_file_image_tab()
        self.create_data_calc_tab()
        self.create_code_gen_tab()
        self.create_about_tab()

        self.log_area = ScrolledText(self, height=12)
        self.log_area.pack(pady=10, padx=10, fill="both", expand=True)
        self.log_area.config(state=tk.DISABLED)
        self.log("Welcome to the Combined Utility Tool!")
        
        self.run_in_thread(check_for_updates, self, self.log, __version__, silent=True)
        
        if self.gemini_api_key.get():
            self.configure_gemini()

    def log(self, message):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, str(message) + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
    
    def run_in_thread(self, target_func, *args, **kwargs):
        thread = threading.Thread(target=target_func, args=args, kwargs=kwargs)
        thread.daemon = True
        thread.start()

    def task_completion_popup(self, title, message):
        """Shows a messagebox popup from the main thread."""
        self.after(0, messagebox.showinfo, title, message)

    def create_ai_assistant_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="✨ AI Assistant")
        api_frame = ttk.LabelFrame(tab, text="Gemini API Configuration")
        api_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(api_frame, text="Google API Key:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        api_entry = ttk.Entry(api_frame, textvariable=self.gemini_api_key, width=50, show="*")
        api_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        api_frame.grid_columnconfigure(1, weight=1)
        ttk.Button(api_frame, text="Set & Save Key", command=self.configure_gemini).grid(row=0, column=2, padx=5, pady=5)
        self.ai_status_label = ttk.Label(api_frame, text="Status: Not Configured", foreground="red")
        self.ai_status_label.grid(row=0, column=3, padx=10, pady=5)
        chat_frame = ttk.LabelFrame(tab, text="Chat")
        chat_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.chat_display = ScrolledText(chat_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Helvetica", 10))
        self.chat_display.pack(fill="both", expand=True, padx=5, pady=5)
        input_frame = ttk.Frame(chat_frame)
        input_frame.pack(fill="x", padx=5, pady=5)
        self.user_input_entry = ttk.Entry(input_frame, font=("Helvetica", 10))
        self.user_input_entry.pack(fill="x", expand=True, side="left", padx=(0, 5))
        self.user_input_entry.bind("<Return>", self.on_send_message)
        send_button = ttk.Button(input_frame, text="Send", command=self.on_send_message)
        send_button.pack(side="right")
        
    def configure_gemini(self):
        """Gets key from UI and asks backend to initialize the AI model."""
        api_key = self.gemini_api_key.get()
        if not api_key:
            messagebox.showerror("Error", "Please enter a Google API Key.")
            return
        
        self.log("Configuring Gemini API...")
        
        model, error = backend.initialize_gemini_model(api_key)
        
        if error:
            self.gemini_model = None
            self.ai_status_label.config(text="Status: Configuration Failed", foreground="red")
            error_message = f"Could not configure the Gemini API.\nCheck project settings (API enabled, billing, etc.).\n\nError: {error}"
            self.log(f"ERROR: {error_message}")
            messagebox.showerror("Configuration Failed", error_message)
        else:
            self.gemini_model = model
            self.ai_status_label.config(text="Status: Ready", foreground="green")
            self.log("✅ Gemini API configured successfully.")
            self.settings['gemini_api_key'] = api_key
            save_settings(self.settings)
            self.log("API Key has been securely saved to settings.json.")

    def on_send_message(self, event=None):
        if not self.gemini_model: messagebox.showwarning("Warning", "Please set API key first."); return
        user_prompt = self.user_input_entry.get().strip()
        if not user_prompt: return
        self._update_chat_window(f"You: {user_prompt}")
        self.user_input_entry.delete(0, tk.END)
        self.ai_status_label.config(text="Status: AI is thinking...")
        self.run_in_thread(self.get_and_display_ai_response, user_prompt)
        
    def get_and_display_ai_response(self, prompt):
        ai_response = backend.ask_ai(self.gemini_model, prompt)
        self.after(0, self._update_chat_window, f"AI: {ai_response}")
        self.after(0, self.ai_status_label.config, {"text": "Status: Ready"})

    def _update_chat_window(self, message):
        self.chat_display.config(state=tk.NORMAL)
        self.chat_display.insert(tk.END, message + "\n\n")
        self.chat_display.see(tk.END)
        self.chat_display.config(state=tk.DISABLED)

    def create_file_image_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="File & Image Tools")
        file_ops_frame = ttk.LabelFrame(tab, text="1. Copy/Move Files by List")
        file_ops_frame.pack(fill="x", padx=10, pady=10)
        self.source_folder = tk.StringVar(value=self.settings.get("source_folder", ""))
        self.target_folder = tk.StringVar(value=self.settings.get("target_folder", ""))
        self.numbers_file = tk.StringVar()
        ttk.Label(file_ops_frame, text="Source Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(file_ops_frame, textvariable=self.source_folder, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_ops_frame, text="Browse...", command=lambda: self.source_folder.set(filedialog.askdirectory())).grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(file_ops_frame, text="Target Folder:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(file_ops_frame, textvariable=self.target_folder, width=60).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_ops_frame, text="Browse...", command=lambda: self.target_folder.set(filedialog.askdirectory())).grid(row=1, column=2, padx=5, pady=5)
        ttk.Label(file_ops_frame, text="Numbers File (List):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(file_ops_frame, textvariable=self.numbers_file, width=60).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(file_ops_frame, text="Browse...", command=lambda: self.numbers_file.set(filedialog.askopenfilename())).grid(row=2, column=2, padx=5, pady=5)
        btn_frame = ttk.Frame(file_ops_frame)
        btn_frame.grid(row=3, column=1, pady=10)
        ttk.Button(btn_frame, text="Copy Files", command=lambda: self.start_process_files("copy")).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Move Files", command=lambda: self.start_process_files("move")).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Save Settings", command=self.save_folder_settings).pack(side="left", padx=5)
        heic_frame = ttk.LabelFrame(tab, text="2. Convert HEIC to JPG")
        heic_frame.pack(fill="x", padx=10, pady=10)
        self.heic_folder = tk.StringVar()
        ttk.Label(heic_frame, text="Folder with HEIC files:").pack(side="left", padx=5, pady=5)
        ttk.Entry(heic_frame, textvariable=self.heic_folder, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        ttk.Button(heic_frame, text="Browse...", command=lambda: self.heic_folder.set(filedialog.askdirectory())).pack(side="left", padx=5, pady=5)
        ttk.Button(heic_frame, text="Convert", command=self.start_heic_conversion).pack(side="left", padx=5, pady=5)
        resize_frame = ttk.LabelFrame(tab, text="3. Batch Image Resizer")
        resize_frame.pack(fill="x", padx=10, pady=10)
        self.resize_folder = tk.StringVar()
        self.quality = tk.StringVar(value="75")
        self.resize_mode = tk.StringVar(value="width")
        self.max_width = tk.StringVar(value="1920")
        self.resize_percentage = tk.StringVar(value="50")
        ttk.Label(resize_frame, text="Image Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(resize_frame, textvariable=self.resize_folder, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky="ew")
        ttk.Button(resize_frame, text="Browse...", command=lambda: self.resize_folder.set(filedialog.askdirectory())).grid(row=0, column=4, padx=5, pady=5)
        ttk.Label(resize_frame, text="Resize Mode:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        radio_frame = ttk.Frame(resize_frame)
        radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")
        ttk.Radiobutton(radio_frame, text="By Width", variable=self.resize_mode, value="width", command=self.toggle_resize_mode).pack(side="left")
        ttk.Radiobutton(radio_frame, text="By Percentage", variable=self.resize_mode, value="percentage", command=self.toggle_resize_mode).pack(side="left", padx=10)
        ttk.Label(resize_frame, text="Max Width:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.width_entry = ttk.Entry(resize_frame, textvariable=self.max_width, width=10)
        self.width_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(resize_frame, text="Percentage (%):").grid(row=2, column=2, padx=5, pady=5, sticky="e")
        self.percentage_entry = ttk.Entry(resize_frame, textvariable=self.resize_percentage, width=10)
        self.percentage_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(resize_frame, text="JPEG Quality (1-95):").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(resize_frame, textvariable=self.quality, width=10).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(resize_frame, text="Resize & Compress", command=self.start_resize_task).grid(row=4, column=1, columnspan=2, pady=10)
        self.toggle_resize_mode()

    def toggle_resize_mode(self):
        if self.resize_mode.get() == "width":
            self.width_entry.config(state="normal"); self.percentage_entry.config(state="disabled")
        else:
            self.width_entry.config(state="disabled"); self.percentage_entry.config(state="normal")

    def create_data_calc_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Data & Calculation")
        format_frame = ttk.LabelFrame(tab, text="4. Format Numbers from File")
        format_frame.pack(fill="x", padx=10, pady=10)
        self.format_file = tk.StringVar()
        ttk.Label(format_frame, text="Excel/CSV/TXT File:").pack(side="left", padx=5, pady=5)
        ttk.Entry(format_frame, textvariable=self.format_file, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        ttk.Button(format_frame, text="Browse...", command=lambda: self.format_file.set(filedialog.askopenfilename())).pack(side="left", padx=5, pady=5)
        ttk.Button(format_frame, text="Format", command=self.start_format_numbers).pack(side="left", padx=5, pady=5)
        single_rug_frame = ttk.LabelFrame(tab, text="5. Rug Size Calculator (Single)")
        single_rug_frame.pack(fill="x", padx=10, pady=10)
        self.rug_dim_input = tk.StringVar()
        self.rug_result_label = tk.StringVar()
        ttk.Label(single_rug_frame, text="Dimension (e.g., 5'2\" x 8'):").pack(side="left", padx=5, pady=5)
        ttk.Entry(single_rug_frame, textvariable=self.rug_dim_input, width=20).pack(side="left", padx=5, pady=5)
        ttk.Button(single_rug_frame, text="Calculate", command=self.calculate_single_rug).pack(side="left", padx=5, pady=5)
        ttk.Label(single_rug_frame, textvariable=self.rug_result_label, font=("Helvetica", 10, "bold")).pack(side="left", padx=15, pady=5)
        bulk_rug_frame = ttk.LabelFrame(tab, text="6. BULK Process Rug Sizes from File")
        bulk_rug_frame.pack(fill="x", padx=10, pady=10)
        self.bulk_rug_file = tk.StringVar()
        self.bulk_rug_col = tk.StringVar(value="Size")
        ttk.Label(bulk_rug_frame, text="Excel/CSV File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(bulk_rug_frame, text="Browse...", command=lambda: self.bulk_rug_file.set(filedialog.askopenfilename())).grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(bulk_rug_frame, text="Column Name/Letter:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(bulk_rug_frame, textvariable=self.bulk_rug_col, width=20).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(bulk_rug_frame, text="Process File", command=self.start_bulk_rug_sizer).grid(row=1, column=2, padx=5, pady=5)
        unit_frame = ttk.LabelFrame(tab, text="7. Unit Converter")
        unit_frame.pack(fill="x", padx=10, pady=10)
        self.unit_input = tk.StringVar(value="182 cm to ft")
        self.unit_result_label = tk.StringVar()
        ttk.Label(unit_frame, text="Conversion:").pack(side="left", padx=5, pady=5)
        ttk.Entry(unit_frame, textvariable=self.unit_input, width=20).pack(side="left", padx=5, pady=5)
        ttk.Button(unit_frame, text="Convert", command=self.convert_units).pack(side="left", padx=5, pady=5)
        ttk.Label(unit_frame, textvariable=self.unit_result_label, font=("Helvetica", 10, "bold")).pack(side="left", padx=15, pady=5)

    def create_code_gen_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Code Generators")
        def toggle_dymo_options(output_var, combobox, entry):
            if output_var.get() == "Dymo": combobox.config(state="readonly"); entry.config(state="normal")
            else: combobox.config(state="disabled"); entry.config(state="disabled")
        qr_frame = ttk.LabelFrame(tab, text="8. QR Code Generator")
        qr_frame.pack(fill="x", padx=10, pady=10)
        self.qr_data = tk.StringVar(); self.qr_filename = tk.StringVar(value="qrcode.png")
        self.qr_output_type = tk.StringVar(value="PNG"); self.qr_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0])
        self.qr_bottom_text = tk.StringVar()
        ttk.Label(qr_frame, text="Data/URL:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(qr_frame, textvariable=self.qr_data, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5)
        ttk.Label(qr_frame, text="Output Type:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        qr_radio_frame = ttk.Frame(qr_frame); qr_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")
        qr_dymo_combo = ttk.Combobox(qr_frame, textvariable=self.qr_dymo_size, values=list(DYMO_LABELS.keys()), state="disabled", width=30)
        qr_bottom_entry = ttk.Entry(qr_frame, textvariable=self.qr_bottom_text, state="disabled", width=32)
        ttk.Radiobutton(qr_radio_frame, text="Standard PNG", variable=self.qr_output_type, value="PNG", command=lambda: toggle_dymo_options(self.qr_output_type, qr_dymo_combo, qr_bottom_entry)).pack(side="left", padx=5)
        ttk.Radiobutton(qr_radio_frame, text="Dymo Label", variable=self.qr_output_type, value="Dymo", command=lambda: toggle_dymo_options(self.qr_output_type, qr_dymo_combo, qr_bottom_entry)).pack(side="left", padx=5)
        ttk.Label(qr_frame, text="Dymo Size:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        qr_dymo_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(qr_frame, text="Bottom Text:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        qr_bottom_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(qr_frame, text="Filename:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(qr_frame, textvariable=self.qr_filename, width=60).grid(row=3, column=1, columnspan=3, padx=5, pady=5)
        ttk.Button(qr_frame, text="Generate QR Code", command=self.start_generate_qr).grid(row=4, column=1, columnspan=2, pady=10)
        bc_frame = ttk.LabelFrame(tab, text="9. Barcode Generator")
        bc_frame.pack(fill="x", padx=10, pady=10)
        self.bc_data = tk.StringVar(); self.bc_filename = tk.StringVar(value="barcode.png")
        self.bc_type = tk.StringVar(value='code128'); self.bc_output_type = tk.StringVar(value="PNG")
        self.bc_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0]); self.bc_bottom_text = tk.StringVar()
        ttk.Label(bc_frame, text="Data:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(bc_frame, textvariable=self.bc_data, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(bc_frame, text="Format:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        ttk.Combobox(bc_frame, textvariable=self.bc_type, values=['code39', 'code128', 'ean13', 'upca'], state="readonly", width=15).grid(row=0, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(bc_frame, text="Output Type:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        bc_radio_frame = ttk.Frame(bc_frame); bc_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")
        bc_dymo_combo = ttk.Combobox(bc_frame, textvariable=self.bc_dymo_size, values=list(DYMO_LABELS.keys()), state="disabled", width=30)
        bc_bottom_entry = ttk.Entry(bc_frame, textvariable=self.bc_bottom_text, state="disabled", width=32)
        ttk.Radiobutton(bc_radio_frame, text="Standard PNG", variable=self.bc_output_type, value="PNG", command=lambda: toggle_dymo_options(self.bc_output_type, bc_dymo_combo, bc_bottom_entry)).pack(side="left", padx=5)
        ttk.Radiobutton(bc_radio_frame, text="Dymo Label", variable=self.bc_output_type, value="Dymo", command=lambda: toggle_dymo_options(self.bc_output_type, bc_dymo_combo, bc_bottom_entry)).pack(side="left", padx=5)
        ttk.Label(bc_frame, text="Dymo Size:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        bc_dymo_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(bc_frame, text="Bottom Text:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        bc_bottom_entry.grid(row=2, column=3, padx=5, pady=5, sticky="w")
        ttk.Label(bc_frame, text="Filename:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(bc_frame, textvariable=self.bc_filename, width=60).grid(row=3, column=1, columnspan=3, padx=5, pady=5)
        ttk.Button(bc_frame, text="Generate Barcode", command=self.start_generate_barcode).grid(row=4, column=1, columnspan=2, pady=10)

    def create_about_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Help & About")
        top_frame = ttk.Frame(tab); top_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(top_frame, text="Check for Updates", command=lambda: self.run_in_thread(check_for_updates, self, self.log, __version__, silent=False)).pack(side="left")
        help_text_area = ScrolledText(tab, wrap=tk.WORD, padx=10, pady=10, font=("Helvetica", 10))
        help_text_area.pack(fill="both", expand=True)
        help_content = f"""
Combined Utility Tool - v{__version__}
This application combines common file, image, and data processing tasks into a single interface.
--- FEATURES ---
✨ AI Assistant:
   A chat interface powered by Google Gemini. Configure it with your own API key to ask questions or get help.
1. Copy/Move Files by List:
   Finds and copies or moves image files based on a list in an Excel or text file.
2. Convert HEIC to JPG:
   Converts Apple's HEIC format images to the universal JPG format.
3. Batch Image Resizer:
   Resizes images by a fixed width or by a percentage of the original dimensions.
4. Format Numbers from File:
   Formats items from a file's first column into a single, comma-separated line.
5. Rug Size Calculator (Single):
   Calculates dimensions in inches and square feet from a text entry (e.g., "5'2\\" x 8'").
6. BULK Process Rug Sizes from File:
   Processes a column of dimensions in an Excel/CSV file, adding calculated width, height, and area.
7. Unit Converter:
   Quickly converts between units like cm, m, ft, and inches.
8. QR Code Generator:
   Creates a QR code from text or a URL, savable as a PNG or Dymo label.
9. Barcode Generator:
   Creates common barcodes, savable as a PNG or Dymo label.
---------------------------------
Created by Hakan Akaslan
"""
        help_text_area.insert(tk.END, help_content)
        help_text_area.config(state=tk.DISABLED)

    # --- UI Starter Methods ---
    def save_folder_settings(self):
        src = self.source_folder.get(); tgt = self.target_folder.get()
        if not src or not tgt: messagebox.showwarning("Warning", "Source and Target folders cannot be empty."); return
        self.settings['source_folder'] = src; self.settings['target_folder'] = tgt
        save_settings(self.settings)
        self.log("✅ Settings saved to settings.json")
        messagebox.showinfo("Success", "Folder settings have been saved.")

    def start_process_files(self, action):
        src = self.source_folder.get(); tgt = self.target_folder.get(); nums_f = self.numbers_file.get()
        if not all([src, tgt, nums_f]): messagebox.showerror("Error", "Please specify Source, Target, and Numbers File."); return
        self.run_in_thread(backend.process_files_task, src, tgt, nums_f, action, self.log, self.task_completion_popup)

    def start_heic_conversion(self):
        folder = self.heic_folder.get()
        if not folder or not os.path.isdir(folder): messagebox.showerror("Error", "Please select a valid folder."); return
        self.run_in_thread(backend.convert_heic_task, folder, self.log, self.task_completion_popup)

    def start_resize_task(self):
        src_folder = self.resize_folder.get()
        if not src_folder or not os.path.isdir(src_folder): messagebox.showerror("Error", "Please select a valid image folder."); return
        mode = self.resize_mode.get()
        try:
            value = int(self.max_width.get()) if mode == 'width' else int(self.resize_percentage.get())
            quality = int(self.quality.get())
            if not (value > 0 and 1 <= quality <= 95): raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Resize values and quality must be valid numbers."); return
        self.run_in_thread(backend.resize_images_task, src_folder, mode, value, quality, self.log, self.task_completion_popup)

    def start_format_numbers(self):
        file_path = self.format_file.get()
        if not file_path: messagebox.showerror("Error", "Please select a file."); return
        err, success_msg = backend.format_numbers_task(file_path)
        if err:
            self.log(f"Error: {err}"); messagebox.showerror("Error", err)
        else:
            self.log(success_msg); messagebox.showinfo("Success", success_msg)

    def calculate_single_rug(self):
        dim_str = self.rug_dim_input.get()
        if not dim_str: self.rug_result_label.set("Please enter a dimension."); return
        w, h = backend.size_to_inches_wh(dim_str); s = backend.calculate_sqft(dim_str)
        if w is not None: self.rug_result_label.set(f"W: {w} in | H: {h} in | Area: {s} sqft")
        else: self.rug_result_label.set("Invalid Format")

    def start_bulk_rug_sizer(self):
        path = self.bulk_rug_file.get(); col = self.bulk_rug_col.get()
        if not path or not col: messagebox.showerror("Error", "Please select a file and specify a column."); return
        self.run_in_thread(backend.bulk_rug_sizer_task, path, col, self.log, self.task_completion_popup)

    def convert_units(self):
        input_str = self.unit_input.get()
        result_str = backend.convert_units_logic(input_str)
        self.unit_result_label.set(result_str)
    
    def start_generate_qr(self):
        data = self.qr_data.get(); fname = self.qr_filename.get()
        if not data or not fname: messagebox.showerror("Error", "Data and filename are required."); return
        dymo_info = DYMO_LABELS[self.qr_dymo_size.get()] if self.qr_output_type.get() == "Dymo" else None
        log_msg, success_msg = backend.generate_qr_task(data, fname, self.qr_output_type.get(), dymo_info, self.qr_bottom_text.get())
        self.log(log_msg)
        if success_msg: self.task_completion_popup("Success", success_msg)
        else: messagebox.showerror("Error", log_msg)

    def start_generate_barcode(self):
        data = self.bc_data.get(); fname = self.bc_filename.get()
        if not data or not fname: messagebox.showerror("Error", "Data and filename are required."); return
        dymo_info = DYMO_LABELS[self.bc_dymo_size.get()] if self.bc_output_type.get() == "Dymo" else None
        log_msg, success_msg = backend.generate_barcode_task(data, fname, self.bc_type.get(), self.bc_output_type.get(), dymo_info, self.bc_bottom_text.get() or data)
        self.log(log_msg)
        if success_msg: self.task_completion_popup("Success", success_msg)
        else: messagebox.showerror("Error", log_msg)








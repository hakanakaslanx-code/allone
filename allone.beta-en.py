# -*- coding: utf-8 -*-
__version__ = "2.3-GUI-EN"

"""
This script combines multiple utility programs into a single GUI application:
- Features Auto-Update functionality.
- English UI.
- Dymo Label support.
"""

import sys
import subprocess
import os
import re
import requests
import json
import logging
import shutil
import threading

# --- GUI Imports ---
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# --- Global Dymo Label Definitions ---
DYMO_LABELS = {
    'Address (30252)': {'w_in': 3.5, 'h_in': 1.125},
    'Shipping (30256)': {'w_in': 4.0, 'h_in': 2.3125},
    'Small Multipurpose (30336)': {'w_in': 2.125, 'h_in': 1.0},
    'File Folder (30258)': {'w_in': 3.5, 'h_in': 0.5625},
}

# --- Automatic Setup ---
def install_and_check():
    required_packages = [
        'tqdm', 'openpyxl', 'Pillow', 'pillow-heif',
        'pandas', 'requests', 'xlrd', 'qrcode', 'python-barcode'
    ]
    print("Checking for required libraries...")
    for package in required_packages:
        try:
            import_name = package.replace('-', '_')
            __import__(import_name)
        except ImportError:
            try:
                print(f"'{package}' not found. Installing...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            except subprocess.CalledProcessError:
                print(f"ERROR: Failed to install '{package}'. Please install manually."); sys.exit(1)
    print("\n✅ Setup checks complete. Starting GUI...")

# --- Backend Logic ---
logging.basicConfig(filename="tool_operations.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
SETTINGS_FILE = "settings.json"

def clean_file_path(file_path: str) -> str: return file_path.strip().strip('"').strip("'")

def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding='utf-8') as f: return json.load(f)
        except json.JSONDecodeError: return {}
    return {}

def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding='utf-8') as f: json.dump(settings, f, indent=4)

def load_numbers_from_file(file_path: str) -> list:
    import pandas as pd
    try:
        p = clean_file_path(file_path)
        if not os.path.exists(p): return []
        if p.lower().endswith((".xlsx",".xls")): df = pd.read_excel(p, header=None, usecols=[0], dtype=str)
        else: df = pd.read_csv(p, header=None, usecols=[0], dtype=str, on_bad_lines='skip', sep=r'\s+|\t|,', engine='python')
        return df[0].dropna().str.strip().tolist()
    except Exception as e:
        print(f"Error reading file '{p}': {e}"); return []

def parse_feet_inches(value_str: str):
    if not isinstance(value_str, str) or not value_str.strip(): return None
    try:
        s = value_str.strip().lower().replace("”",'"').replace("″",'"').replace("′","'").replace("’","'").replace("inches",'"').replace("inch",'"').replace("in",'"')
        s = re.sub(r"\s+", "", s)
        m = re.fullmatch(r"(\d+(?:\.\d+)?)\'(\d+(?:\.\d+)?)?\"?", s)
        if m: return float(m.group(1)) + (float(m.group(2)) if m.group(2) else 0.0) / 12.0
        m = re.fullmatch(r'(\d+(?:\.\d+)?)"', s)
        if m: return float(m.group(1)) / 12.0
        if "'" not in s and "." in s: p=s.split(".",1); return float(p[0] or 0) + float(p[1] or 0) / 12.0
        if re.fullmatch(r'\d+(?:\.\d+)?', s): return float(s)
    except: return None

def size_to_inches_wh(s: str):
    m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(s))
    if not m: return (None, None)
    w = parse_feet_inches(m.group(1)); h = parse_feet_inches(m.group(2))
    return (round(w*12, 2), round(h*12, 2)) if w is not None and h is not None else (None, None)

def calculate_sqft(s: str):
    try:
        m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(s))
        if not m: return None
        w, h = parse_feet_inches(m.group(1)), parse_feet_inches(m.group(2))
        return round(w * h, 2) if w is not None and h is not None else None
    except: return None

def create_label_image(code_image, label_info, bottom_text=""):
    from PIL import Image, ImageDraw, ImageFont
    DPI = 300
    label_width_px = int(label_info['w_in'] * DPI)
    label_height_px = int(label_info['h_in'] * DPI)
    label_bg = Image.new('RGB', (label_width_px, label_height_px), 'white')
    padding = int(0.1 * DPI)
    text_area_height = int(0.25 * DPI) if bottom_text else 0
    max_code_w = label_width_px - (2 * padding)
    max_code_h = label_height_px - (2 * padding) - text_area_height
    code_image.thumbnail((max_code_w, max_code_h), Image.Resampling.LANCZOS)
    paste_x = (label_width_px - code_image.width) // 2
    paste_y = (label_height_px - text_area_height - code_image.height) // 2
    label_bg.paste(code_image, (paste_x, paste_y))
    if bottom_text:
        draw = ImageDraw.Draw(label_bg)
        font, font_size = None, int(0.15 * DPI)
        for font_name in ["arial.ttf", "calibri.ttf", "Helvetica.ttf", "Verdana.ttf"]:
            try: font = ImageFont.truetype(font_name, size=font_size); break
            except IOError: continue
        if not font: font = ImageFont.load_default()
        text_bbox = draw.textbbox((0, 0), bottom_text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_x = (label_width_px - text_width) // 2
        text_y = paste_y + code_image.height + int(padding * 0.2)
        draw.text((text_x, text_y), bottom_text, font=font, fill='black')
    return label_bg


class ToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Combined Utility Tool v{__version__}")
        self.geometry("850x700")

        self.settings = load_settings()
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)

        self.create_file_image_tab()
        self.create_data_calc_tab()
        self.create_code_gen_tab()
        self.create_about_tab()

        self.log_area = ScrolledText(self, height=10)
        self.log_area.pack(pady=10, padx=10, fill="both", expand=True)
        self.log_area.config(state=tk.DISABLED)
        self.log("Welcome to the Combined Utility Tool!")
        
        # --- AUTO-UPDATE ---
        # Run the update check in a separate thread on startup
        self.run_in_thread(self.check_for_updates, silent=True)

    def log(self, message):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
    
    def run_in_thread(self, target_func, *args, **kwargs):
        thread = threading.Thread(target=target_func, args=args, kwargs=kwargs)
        thread.daemon = True
        thread.start()
        
    def check_for_updates(self, silent=False):
        """Checks for a new version of the script on GitHub and self-updates if one is found."""
        self.log("Checking for updates...")
        
        # IMPORTANT: This URL should point to the raw file of THIS script on your GitHub.
        script_url = "https://raw.githubusercontent.com/hakanakaslanx-code/allone/refs/heads/main/allone.beta-en.py"
        
        current_script_name = os.path.basename(sys.argv[0])
        
        try:
            response = requests.get(script_url, timeout=10)
            response.raise_for_status()
            remote_script_content = response.text
            
            match = re.search(r"__version__\s*=\s*[\"'](.+?)[\"']", remote_script_content)
            if not match:
                self.log("Warning: Could not determine remote version. Skipping update.")
                if not silent: messagebox.showwarning("Update Check", "Could not determine remote version number. Update skipped.")
                return
            
            remote_version = match.group(1)
            
            if remote_version > __version__:
                self.log(f"--> New version ({remote_version}) available.")
                
                # Ask user for confirmation
                update_confirmation = messagebox.askyesno("Update Available", f"A new version ({remote_version}) is available.\nYour current version is {__version__}.\n\nDo you want to update now?")
                
                if not update_confirmation:
                    self.log("Update declined by user.")
                    return
                
                messagebox.showinfo("Updating...", "The application will now close, update itself, and restart. Please wait.")
                self.log("Updating...")
                
                new_script_path = current_script_name + ".new"
                with open(new_script_path, "w", encoding='utf-8') as f:
                    f.write(remote_script_content)
                
                if sys.platform == 'win32':
                    updater_script_path = "updater.bat"
                    updater_content = f"""
@echo off
echo Updating script... please wait.
timeout /t 2 /nobreak > NUL
del "{current_script_name}"
rename "{new_script_path}" "{current_script_name}"
echo ✅ Update complete. Relaunching...
start "" "{sys.executable}" "{current_script_name}"
del "{updater_script_path}"
                    """
                else:  # for macOS and Linux
                    updater_script_path = "updater.sh"
                    updater_content = f"""
#!/bin/bash
echo "Updating script... please wait."
sleep 2
rm "{current_script_name}"
mv "{new_script_path}" "{current_script_name}"
echo "✅ Update complete. Relaunching..."
chmod +x "{current_script_name}"
"{sys.executable}" "{current_script_name}" &
rm -- "$0"
                    """
                
                with open(updater_script_path, "w", encoding='utf-8') as f:
                    f.write(updater_content)
                
                if sys.platform != 'win32':
                    os.chmod(updater_script_path, 0o755)
                
                subprocess.Popen([updater_script_path])
                self.destroy() # Close the GUI gracefully
                sys.exit(0)
                
            else:
                self.log(f"✅ Your application is up-to-date. (Version: {__version__})")
                if not silent: messagebox.showinfo("Update Check", "You are running the latest version.")
                
        except Exception as e:
            self.log(f"Warning: Update check failed. Reason: {e}")
            if not silent: messagebox.showerror("Update Check Failed", f"Could not check for updates.\n\nReason: {e}")


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
        ttk.Button(btn_frame, text="Copy Files", command=lambda: self.run_in_thread(self.process_files, "copy")).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Move Files", command=lambda: self.run_in_thread(self.process_files, "move")).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Save Settings", command=self.save_folder_settings).pack(side="left", padx=5)

        heic_frame = ttk.LabelFrame(tab, text="2. Convert HEIC to JPG")
        heic_frame.pack(fill="x", padx=10, pady=10)
        self.heic_folder = tk.StringVar()
        ttk.Label(heic_frame, text="Folder with HEIC files:").pack(side="left", padx=5, pady=5)
        ttk.Entry(heic_frame, textvariable=self.heic_folder, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        ttk.Button(heic_frame, text="Browse...", command=lambda: self.heic_folder.set(filedialog.askdirectory())).pack(side="left", padx=5, pady=5)
        ttk.Button(heic_frame, text="Convert", command=lambda: self.run_in_thread(self.convert_heic)).pack(side="left", padx=5, pady=5)
        
        resize_frame = ttk.LabelFrame(tab, text="3. Batch Image Resizer")
        resize_frame.pack(fill="x", padx=10, pady=10)
        self.resize_folder = tk.StringVar()
        self.max_width = tk.StringVar(value="1920")
        self.quality = tk.StringVar(value="75")
        
        ttk.Label(resize_frame, text="Image Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(resize_frame, textvariable=self.resize_folder, width=60).grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        ttk.Button(resize_frame, text="Browse...", command=lambda: self.resize_folder.set(filedialog.askdirectory())).grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(resize_frame, text="Max Width:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(resize_frame, textvariable=self.max_width, width=10).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(resize_frame, text="JPEG Quality (1-95):").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        ttk.Entry(resize_frame, textvariable=self.quality, width=10).grid(row=1, column=3, padx=5, pady=5, sticky="w")
        
        ttk.Button(resize_frame, text="Resize & Compress", command=lambda: self.run_in_thread(self.resize_images)).grid(row=2, column=1, columnspan=2, pady=10)

    def create_data_calc_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Data & Calculation")
        
        format_frame = ttk.LabelFrame(tab, text="4. Format Numbers from File")
        format_frame.pack(fill="x", padx=10, pady=10)
        self.format_file = tk.StringVar()
        ttk.Label(format_frame, text="Excel/CSV/TXT File:").pack(side="left", padx=5, pady=5)
        ttk.Entry(format_frame, textvariable=self.format_file, width=60).pack(side="left", padx=5, pady=5, expand=True, fill="x")
        ttk.Button(format_frame, text="Browse...", command=lambda: self.format_file.set(filedialog.askopenfilename())).pack(side="left", padx=5, pady=5)
        ttk.Button(format_frame, text="Format", command=self.format_numbers).pack(side="left", padx=5, pady=5)

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
        ttk.Button(bulk_rug_frame, text="Process File", command=lambda: self.run_in_thread(self.bulk_process_rugs)).grid(row=1, column=2, padx=5, pady=5)

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
            if output_var.get() == "Dymo":
                combobox.config(state="readonly")
                entry.config(state="normal")
            else:
                combobox.config(state="disabled")
                entry.config(state="disabled")

        qr_frame = ttk.LabelFrame(tab, text="8. QR Code Generator")
        qr_frame.pack(fill="x", padx=10, pady=10)
        self.qr_data = tk.StringVar()
        self.qr_filename = tk.StringVar(value="qrcode.png")
        self.qr_output_type = tk.StringVar(value="PNG")
        self.qr_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0])
        self.qr_bottom_text = tk.StringVar()

        ttk.Label(qr_frame, text="Data/URL:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(qr_frame, textvariable=self.qr_data, width=60).grid(row=0, column=1, columnspan=3, padx=5, pady=5)
        
        ttk.Label(qr_frame, text="Output Type:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        qr_radio_frame = ttk.Frame(qr_frame)
        qr_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")
        
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
        
        ttk.Button(qr_frame, text="Generate QR Code", command=self.generate_qr).grid(row=4, column=1, columnspan=2, pady=10)

        bc_frame = ttk.LabelFrame(tab, text="9. Barcode Generator")
        bc_frame.pack(fill="x", padx=10, pady=10)
        self.bc_data = tk.StringVar()
        self.bc_filename = tk.StringVar(value="barcode.png")
        self.bc_type = tk.StringVar(value='code128')
        self.bc_output_type = tk.StringVar(value="PNG")
        self.bc_dymo_size = tk.StringVar(value=list(DYMO_LABELS.keys())[0])
        self.bc_bottom_text = tk.StringVar()
        
        ttk.Label(bc_frame, text="Data:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(bc_frame, textvariable=self.bc_data, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(bc_frame, text="Format:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        ttk.Combobox(bc_frame, textvariable=self.bc_type, values=['code39', 'code128', 'ean13', 'upca'], state="readonly", width=15).grid(row=0, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(bc_frame, text="Output Type:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        bc_radio_frame = ttk.Frame(bc_frame)
        bc_radio_frame.grid(row=1, column=1, columnspan=3, sticky="w")

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
        
        ttk.Button(bc_frame, text="Generate Barcode", command=self.generate_barcode).grid(row=4, column=1, columnspan=2, pady=10)

    def create_about_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Help & About")
        
        top_frame = ttk.Frame(tab)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(top_frame, text="Check for Updates", command=lambda: self.run_in_thread(self.check_for_updates, silent=False)).pack(side="left")

        help_text_area = ScrolledText(tab, wrap=tk.WORD, padx=10, pady=10)
        help_text_area.pack(fill="both", expand=True)

        help_content = f"""
Combined Utility Tool - v{__version__}

This application combines common file, image, and data processing tasks into a single interface.

--- FEATURES ---

1. Copy/Move Files by List:
   Finds and copies or moves image files from a source folder to a target folder based on a list in an Excel or text file.

2. Convert HEIC to JPG:
   Converts all of Apple's HEIC format images in a selected folder to the universal JPG format.

3. Batch Image Resizer:
   Resizes and compresses all images in a folder to a specified maximum width while preserving the aspect ratio.

4. Format Numbers from File:
   Reads the first column of a file and formats all items into a single, comma-separated line of text.

5. Rug Size Calculator (Single):
   Calculates the dimensions in inches and square feet from a single text entry (e.g., "5'2\\" x 8'").

6. BULK Process Rug Sizes from File:
   Processes a column of dimensions in an Excel/CSV file, calculating width (in), height (in), and area (sqft) for each row.

7. Unit Converter:
   Quickly converts between units like cm, m, ft, and inches.

8. QR Code Generator:
   Creates a QR code from your text or URL. Can be saved as a standard PNG or formatted for a Dymo label.

9. Barcode Generator:
   Creates common barcodes like Code 128, EAN13, etc. Can be saved as a standard PNG or formatted for a Dymo label.

---------------------------------
Created by Hakan Akaslan
"""
        
        help_text_area.insert(tk.END, help_content)
        help_text_area.config(state=tk.DISABLED)


    # --- Action Methods ---
    def save_folder_settings(self):
        src = self.source_folder.get()
        tgt = self.target_folder.get()
        if not src or not tgt:
            messagebox.showwarning("Warning", "Source and Target folders cannot be empty.")
            return
        
        self.settings['source_folder'] = src
        self.settings['target_folder'] = tgt
        save_settings(self.settings)
        self.log("✅ Settings saved to settings.json")
        messagebox.showinfo("Success", "Folder settings have been saved.")

    def process_files(self, action):
        from tqdm import tqdm
        src = self.source_folder.get()
        tgt = self.target_folder.get()
        nums_f = self.numbers_file.get()

        if not all([src, tgt, nums_f]):
            messagebox.showerror("Error", "Please specify Source, Target, and Numbers File.")
            return

        self.log(f"Starting file {action} process...")
        nums = load_numbers_from_file(nums_f)
        if not nums:
            self.log("No numbers found in the file."); return
        
        src, tgt = clean_file_path(src), clean_file_path(tgt)
        proc, missing = [], set(nums)
        exts = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tiff'}
        files = [f for f in os.listdir(src) if os.path.isfile(os.path.join(src, f))]
        map_ = {n: [f for f in files if n in f and os.path.splitext(f)[1].lower() in exts] for n in nums}
        
        for n in tqdm(nums, desc=f"{action.title()}ing files"):
            if map_.get(n):
                for f in map_[n]:
                    try:
                        if action == "copy": shutil.copy2(os.path.join(src, f), os.path.join(tgt, f))
                        else: shutil.move(os.path.join(src, f), os.path.join(tgt, f))
                        proc.append(f); missing.discard(n)
                    except Exception as e: self.log(f"Error processing '{f}': {e}")
            else: logging.warning(f"No match for: {n}")
        
        summary = f"\n--- Summary ---\nProcessed: {len(proc)}\nNot Found: {len(missing)}"
        self.log(summary)
        if missing: self.log(f"Unfound: {', '.join(list(missing))}")
        messagebox.showinfo("Complete", f"File {action} process finished. See log for details.")

    def convert_heic(self):
        from PIL import Image
        import pillow_heif
        from tqdm import tqdm
        
        folder = self.heic_folder.get()
        if not folder or not os.path.isdir(folder): messagebox.showerror("Error", "Please select a valid folder."); return
        
        self.log("Starting HEIC to JPG conversion...")
        try:
            files = [f for f in os.listdir(folder) if f.lower().endswith(".heic")]
            if not files: self.log("No HEIC files found."); return
            
            for f in tqdm(files, desc="HEIC -> JPG"):
                src = os.path.join(folder, f); dst = f"{os.path.splitext(src)[0]}.jpg"
                try:
                    heif = pillow_heif.read_heif(src)
                    img = Image.frombytes(heif.mode, heif.size, heif.data, "raw")
                    img.save(dst, "JPEG")
                    self.log(f"Converted: {f} -> {os.path.basename(dst)}")
                except Exception as e: self.log(f"Error converting '{f}': {e}")
            
            self.log("\n✅ Conversion complete.")
            messagebox.showinfo("Success", "HEIC conversion is complete.")
        except Exception as e:
            self.log(f"An error occurred: {e}"); messagebox.showerror("Error", f"An error occurred: {e}")

    def resize_images(self):
        from PIL import Image
        from tqdm import tqdm
        
        src_folder = self.resize_folder.get()
        if not src_folder or not os.path.isdir(src_folder): messagebox.showerror("Error", "Please select a valid image folder."); return

        try:
            w = int(self.max_width.get()); q = int(self.quality.get())
            if not 1 <= q <= 95: q = 75
        except ValueError: messagebox.showerror("Error", "Max width and quality must be valid numbers."); return
        
        tgt_folder = os.path.join(src_folder, "resized")
        os.makedirs(tgt_folder, exist_ok=True)
        self.log(f"Resized images will be saved in: {tgt_folder}")
        
        files = [f for f in os.listdir(src_folder) if f.lower().endswith(('.jpg','.jpeg','.png','.webp'))]
        if not files: self.log("No compatible images found."); return
        
        for f in tqdm(files, desc="Resizing images"):
            try:
                with Image.open(os.path.join(src_folder, f)) as img:
                    if img.width > w:
                        r = w / float(img.width); h = int(float(img.height) * r)
                        resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                        img = img.resize((w, h), resample)
                    if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                    
                    if f.lower().endswith(('.jpg','.jpeg')): img.save(os.path.join(tgt_folder, f), "JPEG", quality=q, optimize=True)
                    else: img.save(os.path.join(tgt_folder, f))
                    self.log(f"Resized: {f}")
            except Exception as e: self.log(f"Error with {f}: {e}")
        
        self.log("\n✅ Image resizing complete.")
        messagebox.showinfo("Success", "Image processing is complete.")

    def format_numbers(self):
        file_path = self.format_file.get()
        if not file_path: messagebox.showerror("Error", "Please select a file."); return
        
        nums = load_numbers_from_file(file_path)
        if not nums: self.log("No numbers found in file."); return
        
        out_str = ",".join(nums)
        try:
            with open("formatted_numbers.txt", "w", encoding='utf-8') as f: f.write(out_str)
            self.log("Formatted numbers saved to 'formatted_numbers.txt'.")
            messagebox.showinfo("Success", f"Formatted text saved to {os.path.abspath('formatted_numbers.txt')}")
        except Exception as e:
            self.log(f"Error saving file: {e}"); messagebox.showerror("Error", f"Could not save file: {e}")

    def calculate_single_rug(self):
        dim_str = self.rug_dim_input.get()
        if not dim_str: self.rug_result_label.set("Please enter a dimension."); return
        
        w, h = size_to_inches_wh(dim_str)
        s = calculate_sqft(dim_str)
        
        if w is not None: self.rug_result_label.set(f"W: {w} in | H: {h} in | Area: {s} sqft")
        else: self.rug_result_label.set("Invalid Format")

    def bulk_process_rugs(self):
        import pandas as pd
        from tqdm import tqdm
        tqdm.pandas(desc="Calculating Dimensions")
        
        path = self.bulk_rug_file.get()
        col = self.bulk_rug_col.get()
        if not path or not col: messagebox.showerror("Error", "Please select a file and specify a column."); return

        self.log(f"Processing rug sizes from: {path}")
        try: df = pd.read_excel(path) if path.lower().endswith((".xlsx",".xls")) else pd.read_csv(path)
        except Exception as e: self.log(f"Error reading file: {e}"); messagebox.showerror("Error", f"Could not read file: {e}"); return
        
        sel_col = None
        if len(col) == 1 and col.isalpha():
            idx = ord(col.upper()) - ord('A')
            if idx < len(df.columns): sel_col = df.columns[idx]
        elif col in df.columns: sel_col = col
        
        if not sel_col: messagebox.showerror("Error", f"Column '{col}' not found."); return

        res = df[sel_col].progress_apply(lambda s: {'w': size_to_inches_wh(s)[0], 'h': size_to_inches_wh(s)[1], 'a': calculate_sqft(s)})
        df["Width_in"] = [r['w'] for r in res]; df["Height_in"] = [r['h'] for r in res]; df["Area_sqft"] = [r['a'] for r in res]
        
        out_path = f"{os.path.splitext(path)[0]}_with_sizes.xlsx"
        try:
            df.to_excel(out_path, index=False)
            self.log(f"✅ Saved to: {out_path}")
            messagebox.showinfo("Success", f"Processed file saved to:\n{out_path}")
        except Exception as e:
            csv_path = f"{os.path.splitext(path)[0]}_with_sizes.csv"
            df.to_csv(csv_path, index=False)
            self.log(f"Could not save as Excel ({e}). ✅ Saved to CSV instead: {csv_path}")
            messagebox.showwarning("Saved as CSV", f"Could not save as Excel. Saved as CSV instead:\n{csv_path}")

    def convert_units(self):
        i = self.unit_input.get().lower()
        if not i: return
        
        m = re.match(r"^\s*(.+?)\s*(cm|m|ft|in)\s+to\s+(cm|m|ft|in)\s*$", i, re.I)
        if not m: self.unit_result_label.set("Invalid Format"); return

        v_str, fu, tu = m.groups(); cm = None
        try:
            if fu == 'ft': cm = parse_feet_inches(v_str) * 30.48 if parse_feet_inches(v_str) else None
            else: val = float(v_str); cm = val if fu == 'cm' else val * 100 if fu == 'm' else val * 2.54 if fu == 'in' else None
        except: pass
        
        if cm is None: self.unit_result_label.set(f"Could not parse '{v_str}'."); return
        
        res = ""
        if tu == 'cm': res = f"{cm:.2f} cm"
        elif tu == 'm': res = f"{cm / 100:.2f} m"
        elif tu == 'in': res = f"{cm / 2.54:.2f} in"
        elif tu == 'ft': total_in = cm / 2.54; res = f"{int(total_in // 12)}' {total_in % 12:.2f}\""
        
        self.unit_result_label.set(f"--> {res}")
    
    def generate_qr(self):
        import qrcode
        data = self.qr_data.get()
        fname = self.qr_filename.get()
        output_type = self.qr_output_type.get()
        
        if not data or not fname: messagebox.showerror("Error", "Data and filename are required."); return
        
        try:
            if output_type == "PNG":
                qrcode.make(data).save(fname)
            else: # Dymo
                label_name = self.qr_dymo_size.get()
                label_info = DYMO_LABELS[label_name]
                bottom_text = self.qr_bottom_text.get()
                
                qr_img = qrcode.make(data)
                label_image = create_label_image(qr_img, label_info, bottom_text)
                label_image.save(fname)
            
            self.log(f"✅ QR Code saved as '{fname}'")
            messagebox.showinfo("Success", f"QR Code saved to:\n{os.path.abspath(fname)}")

        except Exception as e:
            self.log(f"Error generating QR Code: {e}"); messagebox.showerror("Error", f"Error: {e}")

    def generate_barcode(self):
        import barcode
        from barcode.writer import ImageWriter
        
        data = self.bc_data.get()
        fname = self.bc_filename.get()
        bc_format = self.bc_type.get()
        output_type = self.bc_output_type.get()
        
        if not data or not fname: messagebox.showerror("Error", "Data and filename are required."); return

        try:
            if output_type == "PNG":
                BarcodeClass = barcode.get_barcode_class(bc_format)
                saved_fname = BarcodeClass(data, writer=ImageWriter()).save(fname.replace('.png',''))
                self.log(f"✅ Barcode saved as '{saved_fname}'")
                messagebox.showinfo("Success", f"Barcode saved to:\n{os.path.abspath(saved_fname)}")
            else: # Dymo
                label_name = self.bc_dymo_size.get()
                label_info = DYMO_LABELS[label_name]
                bottom_text = self.bc_bottom_text.get() or data
                
                BarcodeClass = barcode.get_barcode_class(bc_format)
                barcode_pil_img = BarcodeClass(data, writer=ImageWriter()).render()
                label_image = create_label_image(barcode_pil_img, label_info, bottom_text)
                label_image.save(fname)
                self.log(f"✅ Dymo Barcode Label saved as '{fname}'")
                messagebox.showinfo("Success", f"Dymo Label saved to:\n{os.path.abspath(fname)}")
                
        except Exception as e:
            self.log(f"Error generating barcode: {e}"); messagebox.showerror("Error", f"Error: {e}")

if __name__ == "__main__":
    install_and_check()
    app = ToolApp()
    app.mainloop()


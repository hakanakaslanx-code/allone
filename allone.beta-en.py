# -*- coding: utf-8 -*-

# ATTENTION: When you update the code, you must increment this version number (e.g., "1.2").
__version__ = "1.8"

"""
This script combines multiple utility programs into a single interface:
1. File Copy/Move: Manages files based on specific numbers or identifiers.
2. Excel/Text Formatter: Formats data from a column into a single text file line.
3. HEIC to JPG Converter: Converts HEIC format images to JPG.
4. Rug Size Calculator: Calculates dimensions (inches and sqft) from feet/inch format.
5. Image Resizer/Compressor: Batch resizes and compresses images in a folder.
6. BULK Excel/CSV Rug Sizer: Reads a column of dimensions and adds Width_in / Height_in / Area_sqft.
7. Unit Converter: Converts between cm, m, feet, and inches.
8. QR Code Generator: Creates a QR code image from text or a URL.
9. Barcode Generator: Creates a barcode image in various formats.
"""

import sys
import subprocess
import os
import re
import requests # Required for the updater function
import json
import logging
import shutil

# --- Automatic Setup and Self-Update Mechanism ---

def install_and_check():
    """Checks for required libraries and installs them if they are missing."""
    # GÜNCELLENDİ: 'python-barcode' kütüphanesi eklendi
    required_packages = [
        'tqdm', 'openpyxl', 'Pillow', 'pillow-heif',
        'pandas', 'requests', 'xlrd', 'qrcode', 'python-barcode'
    ]
    
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("ERROR: pip is not available. Please ensure you have pip installed and in your system's PATH.")
        sys.exit(1)

    print("Checking for required libraries...")
    for package in required_packages:
        try:
            # python-barcode kütüphanesi 'barcode' olarak import edilir.
            import_name = package.replace('-', '_')
            __import__(import_name)
        except ImportError:
            try:
                print(f"'{package}' not found. Installing...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            except subprocess.CalledProcessError:
                print(f"ERROR: Failed to install '{package}'. Please install it manually: pip install {package}")
                sys.exit(1)

    print("\n✅ Setup checks complete.")

def check_for_updates():
    """
    Checks for a new version of the script on GitHub and self-updates if one is found.
    """
    print("Checking for updates...")
    script_url = "https://raw.githubusercontent.com/hakanakaslanx-code/allone/refs/heads/main/allone.beta-en.py" # Örnek URL, kendi reponuzla değiştirin
    current_script_name = os.path.basename(sys.argv[0])
    try:
        response = requests.get(script_url)
        response.raise_for_status()
        remote_script_content = response.text
        match = re.search(r"__version__\s*=\s*[\"'](.+?)[\"']", remote_script_content)
        if not match:
            print("Warning: Could not determine remote version. Skipping update.")
            return
        remote_version = match.group(1)
        if remote_version > __version__:
            print(f"--> New version ({remote_version}) available. Updating...")
            new_script_path = current_script_name + ".new"
            with open(new_script_path, "w", encoding='utf-8') as f: f.write(remote_script_content)
            if sys.platform == 'win32':
                updater_script_path = "updater.bat"
                updater_content = f"""
@echo off
timeout /t 2 /nobreak > NUL
del "{current_script_name}"
rename "{new_script_path}" "{current_script_name}"
echo ✅ Update complete. Relaunching...
start "" "{sys.executable}" "{current_script_name}"
del "{updater_script_path}"
                """
            else:
                updater_script_path = "updater.sh"
                updater_content = f"""
#!/bin/bash
sleep 2
rm "{current_script_name}"
mv "{new_script_path}" "{current_script_name}"
echo "✅ Update complete. Relaunching..."
"{sys.executable}" "{current_script_name}" &
rm -- "$0"
                """
            with open(updater_script_path, "w", encoding='utf-8') as f: f.write(updater_content)
            if sys.platform != 'win32': os.chmod(updater_script_path, 0o755)
            subprocess.Popen([updater_script_path])
            sys.exit(0)
        else:
            print(f"✅ Your code is up-to-date. (Version: {__version__})")
    except Exception as e:
        print(f"Warning: Update check failed. Reason: {e}")

# --- Global Settings ---
logging.basicConfig(filename="tool_operations.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
SETTINGS_FILE = "settings.json"

# --- Helper Functions ---
def clean_file_path(file_path: str) -> str: return file_path.strip().strip('"').strip("'")
def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding='utf-8') as f: return json.load(f)
        except json.JSONDecodeError:
            print(f"Warning: '{SETTINGS_FILE}' is corrupt. New settings will be created.")
            return {}
    return {}
def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding='utf-8') as f: json.dump(settings, f, indent=4)

def ask_for_folder_settings(existing_settings: dict) -> dict:
    print("\n--- Folder Settings for File Operations ---")
    print(f"Current Source Folder: {existing_settings.get('source_folder', 'Not set')}")
    print(f"Current Target Folder: {existing_settings.get('target_folder', 'Not set')}")
    if input("Update settings? (y/n): ").strip().lower() == "y":
        source_folder = input("Enter new source folder path: ").strip()
        target_folder = input("Enter new target folder path: ").strip()
        if not os.path.isdir(clean_file_path(source_folder)):
            print(f"Error: Source folder '{source_folder}' not found.")
            return existing_settings
        if not os.path.isdir(clean_file_path(target_folder)):
            if input(f"Target folder '{target_folder}' not found. Create it? (y/n): ").strip().lower() == 'y':
                try: os.makedirs(clean_file_path(target_folder))
                except OSError as e: print(f"Error creating folder: {e}"); return existing_settings
            else: return existing_settings
        new_settings = {"source_folder": source_folder, "target_folder": target_folder}
        save_settings(new_settings)
        print("✅ Settings updated.")
    return load_settings()

def load_numbers_from_file(file_path: str) -> list:
    import pandas as pd
    try:
        p = clean_file_path(file_path)
        if not os.path.exists(p): return []
        if p.lower().endswith(".csv"): df = pd.read_csv(p, header=None, usecols=[0], dtype=str, on_bad_lines='skip')
        elif p.lower().endswith((".xlsx", ".xls")): df = pd.read_excel(p, header=None, usecols=[0], dtype=str, engine=None)
        else: df = pd.read_csv(p, header=None, usecols=[0], dtype=str, sep='\t', on_bad_lines='skip')
        return df[0].dropna().str.strip().tolist()
    except Exception as e: print(f"Error reading file '{p}': {e}"); return []

def parse_feet_inches(value_str: str):
    if not isinstance(value_str, str) or not value_str.strip(): return None
    try:
        s = value_str.strip().lower().replace("”", '"').replace("″", '"').replace("′", "'").replace("’", "'").replace("inches", '"').replace("inch", '"').replace("in", '"')
        s = re.sub(r"\s+", "", s)
        m = re.fullmatch(r"(\d+(?:\.\d+)?)\'(\d+(?:\.\d+)?)?\"?", s)
        if m: return float(m.group(1)) + (float(m.group(2)) if m.group(2) else 0.0) / 12.0
        m = re.fullmatch(r'(\d+(?:\.\d+)?)"', s)
        if m: return float(m.group(1)) / 12.0
        if "'" not in s and "." in s:
            p = s.split(".", 1); return float(p[0] or 0) + float(p[1] or 0) / 12.0
        if re.fullmatch(r'\d+(?:\.\d+)?', s): return float(s)
    except (ValueError, TypeError): return None

def size_to_inches_wh(s: str):
    m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(s))
    if not m: return (None, None)
    w = parse_feet_inches(m.group(1)); h = parse_feet_inches(m.group(2))
    if w is None or h is None: return (None, None)
    return (round(w * 12, 2), round(h * 12, 2))

def calculate_sqft(s: str):
    try:
        m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(s))
        if not m: return None
        w = parse_feet_inches(m.group(1)); h = parse_feet_inches(m.group(2))
        return round(w * h, 2) if w is not None and h is not None else None
    except Exception: return None

# --- Main Modules ---
def rug_size_calculator():
    print("\n=== Rug Size Calculator (inches and sqft) ===")
    while True:
        i = input("Enter dimension ('width x height') (or 'q' to quit): ").strip()
        if i.lower() == 'q': break
        w, h = size_to_inches_wh(i); s = calculate_sqft(i)
        if w is not None: print(f"Result: Width: {w} in | Height: {h} in | Area: {s} sqft")
        else: print("Error: Invalid format.")

def bulk_sizes_from_sheet():
    import pandas as pd
    from tqdm import tqdm
    tqdm.pandas(desc="Calculating Dimensions")
    path = clean_file_path(input("Enter Excel/CSV file path: ").strip())
    if not os.path.exists(path): print("Error: File not found."); return
    try: df = pd.read_excel(path) if path.lower().endswith((".xlsx",".xls")) else pd.read_csv(path)
    except Exception as e: print(f"Error reading file: {e}"); return
    print("\nColumns:", ", ".join(map(str, df.columns)))
    col = input("Enter column name/letter for dimensions (e.g., Size or A): ").strip()
    sel_col = None
    if len(col) == 1 and col.isalpha():
        idx = ord(col.upper()) - ord('A')
        if idx < len(df.columns): sel_col = df.columns[idx]
        else: print("Error: Column letter out of range."); return
    elif col in df.columns: sel_col = col
    else: print(f"Error: Column '{col}' not found."); return
    res = df[sel_col].progress_apply(lambda s: {'w': size_to_inches_wh(s)[0], 'h': size_to_inches_wh(s)[1], 'a': calculate_sqft(s)})
    df["Width_in"] = [r['w'] for r in res]; df["Height_in"] = [r['h'] for r in res]; df["Area_sqft"] = [r['a'] for r in res]
    out_path = f"{os.path.splitext(path)[0]}_with_sizes.xlsx"
    try:
        df.to_excel(out_path, index=False)
        print(f"\n✅ Saved to: {out_path}")
    except Exception as e:
        csv_path = f"{os.path.splitext(path)[0]}_with_sizes.csv"
        df.to_csv(csv_path, index=False)
        print(f"\nCould not save as Excel ({e}). ✅ Saved to CSV instead: {csv_path}")

def format_numbers_from_file():
    f = input("Enter Excel/CSV/TXT file path: ").strip()
    nums = load_numbers_from_file(f)
    if not nums: print("No numbers found."); return
    out = ",".join(nums); print("\n--- Formatted ---\n", out)
    with open("formatted_numbers.txt", "w", encoding='utf-8') as f: f.write(out)
    print("\n✅ Saved to 'formatted_numbers.txt'.")

def convert_heic_to_jpg_in_directory():
    from PIL import Image
    import pillow_heif
    from tqdm import tqdm
    d = clean_file_path(input("Enter folder path with HEIC files: ").strip())
    if not os.path.isdir(d): print("Error: Not a valid directory."); return
    try:
        files = [f for f in os.listdir(d) if f.lower().endswith(".heic")]
        if not files: print("No HEIC files found."); return
        for f in tqdm(files, desc="HEIC -> JPG"):
            src = os.path.join(d, f); dst = f"{os.path.splitext(src)[0]}.jpg"
            try:
                heif = pillow_heif.read_heif(src)
                img = Image.frombytes(heif.mode, heif.size, heif.data, "raw")
                img.save(dst, "JPEG")
            except Exception as e: print(f"\nError converting '{f}': {e}")
        print("\n✅ Conversion complete.")
    except Exception as e: print(f"\nAn error occurred: {e}")

def process_files_main(settings):
    from tqdm import tqdm
    src = settings.get("source_folder"); tgt = settings.get("target_folder")
    if not src or not tgt: print("Set folders first (Option 's')."); return
    nums_file = input("Enter file path with numbers: ").strip()
    nums = load_numbers_from_file(nums_file)
    if not nums: print("No numbers to process."); return
    act = input("Copy (c) or Move (m)?: ").strip().lower()
    if act not in ['c', 'm']: print("Invalid choice."); return
    action = "copy" if act == 'c' else "move"
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
                except Exception as e: print(f"\nError processing '{f}': {e}")
        else: logging.warning(f"No match for: {n}")
    print(f"\n--- Summary ---\nProcessed: {len(proc)}\nNot Found: {len(missing)}")
    if missing: print("Unfound:", ", ".join(list(missing)))

def resize_and_compress_images():
    from PIL import Image
    from tqdm import tqdm
    print("\n=== Bulk Image Resizer & Compressor ===")
    src = clean_file_path(input("Enter folder path with images: ").strip())
    if not os.path.isdir(src): print("Error: Directory not found."); return
    tgt = os.path.join(src, "resized"); os.makedirs(tgt, exist_ok=True)
    print(f"Resized images will be in: {tgt}")
    try:
        w = int(input("Enter max width (e.g., 1920): ").strip())
        q = int(input("Enter JPEG quality (1-95, default 75): ").strip() or 75)
        if not 1 <= q <= 95: q = 75
    except ValueError: print("Invalid number."); return
    files = [f for f in os.listdir(src) if f.lower().endswith(('.jpg','.jpeg','.png','.webp'))]
    if not files: print("No compatible images found."); return
    for f in tqdm(files, desc="Resizing images"):
        try:
            with Image.open(os.path.join(src, f)) as img:
                if img.width > w:
                    r = w / float(img.width); h = int(float(img.height) * r)
                    resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                    img = img.resize((w, h), resample)
                if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                if f.lower().endswith(('.jpg','.jpeg')): img.save(os.path.join(tgt, f), "JPEG", quality=q, optimize=True)
                else: img.save(os.path.join(tgt, f))
        except Exception as e: print(f"\nError with {f}: {e}")
    print("\n✅ Processing complete.")

def unit_converter():
    print("\n=== Unit Converter ===")
    print("Examples: '182 cm to ft', '5\\'11\" to cm'")
    while True:
        i = input("Enter conversion (or 'q' to quit): ").strip().lower()
        if i == 'q': break
        m = re.match(r"^\s*(.+?)\s*(cm|m|ft|in)\s+to\s+(cm|m|ft|in)\s*$", i, re.I)
        if not m: print("Invalid format."); continue
        v, fu, tu = m.groups(); cm = None
        try:
            if fu == 'ft': cm = parse_feet_inches(v) * 30.48 if parse_feet_inches(v) else None
            else: val = float(v); cm = val if fu == 'cm' else val * 100 if fu == 'm' else val * 2.54 if fu == 'in' else None
        except: pass
        if cm is None: print(f"Could not parse '{v}'."); continue
        res = ""
        if tu == 'cm': res = f"{cm:.2f} cm"
        elif tu == 'm': res = f"{cm / 100:.2f} m"
        elif tu == 'in': res = f"{cm / 2.54:.2f} in"
        elif tu == 'ft': total_in = cm / 2.54; res = f"{int(total_in // 12)}' {total_in % 12:.2f}\""
        print(f"--> {v} {fu}  =  {res}")

def generate_qr_code():
    import qrcode
    print("\n=== QR Code Generator ===")
    data = input("Enter text/URL for QR code: ").strip()
    if not data: print("No data provided."); return
    fname = input("Enter filename (default: qrcode.png): ").strip() or "qrcode.png"
    if not fname.lower().endswith('.png'): fname += '.png'
    try:
        qrcode.make(data).save(fname)
        print(f"✅ QR Code saved as '{fname}'.")
    except Exception as e: print(f"An error occurred: {e}")

def generate_barcode():
    import barcode
    from barcode.writer import ImageWriter
    print("\n=== Barcode Generator ===")
    supported_formats = {
        '1': ('ean13', "EAN-13 (Retail - 12 digits)"),
        '2': ('upca', "UPC-A (Retail - 11 digits)"),
        '3': ('code128', "Code 128 (Alphanumeric)"),
        '4': ('code39', "Code 39 (Alphanumeric, simpler)"),
    }
    print("Choose a barcode format:")
    for key, (name, desc) in supported_formats.items(): print(f"  {key}. {desc}")
    choice = input("Your choice: ").strip()
    if choice not in supported_formats: print("Invalid choice."); return
    barcode_format, _ = supported_formats[choice]
    data = input(f"Enter data for {barcode_format.upper()}: ").strip()
    if not data: print("No data provided."); return
    fname = input("Enter filename (default: barcode.png): ").strip() or "barcode"
    try:
        BarcodeClass = barcode.get_barcode_class(barcode_format)
        my_barcode = BarcodeClass(data, writer=ImageWriter())
        my_barcode.save(fname)
        print(f"✅ Barcode saved as '{fname}.png'.")
    except barcode.errors.BarcodeError as e: print(f"\nERROR: Could not generate barcode: {e}")
    except Exception as e: print(f"An unexpected error occurred: {e}")

# --- Central Menu Configuration ---
MENU_OPTIONS = {
    '1': {'d': 'Copy/Move Files by List', 'f': process_files_main, 'c': 'File & Image Tools', 's': True},
    '2': {'d': 'Convert HEIC to JPG', 'f': convert_heic_to_jpg_in_directory, 'c': 'File & Image Tools', 's': False},
    '3': {'d': 'Batch Image Resizer', 'f': resize_and_compress_images, 'c': 'File & Image Tools', 's': False},
    '4': {'d': 'Format Numbers from File', 'f': format_numbers_from_file, 'c': 'Data & Calculation Tools', 's': False},
    '5': {'d': 'Rug Size Calculator (Single)', 'f': rug_size_calculator, 'c': 'Data & Calculation Tools', 's': False},
    '6': {'d': 'BULK Process Rug Sizes from File', 'f': bulk_sizes_from_sheet, 'c': 'Data & Calculation Tools', 's': False},
    '7': {'d': 'Unit Converter (cm, m, ft, in)', 'f': unit_converter, 'c': 'Data & Calculation Tools', 's': False},
    '8': {'d': 'QR Code Generator', 'f': generate_qr_code, 'c': 'Data & Calculation Tools', 's': False},
    '9': {'d': 'Barcode Generator', 'f': generate_barcode, 'c': 'Data & Calculation Tools', 's': False},
    's': {'d': 'Set Folders for File Operations', 'f': ask_for_folder_settings, 'c': 'Settings & Other', 's': True},
    'h': {'d': 'Help / Guide', 'f': None, 'c': 'Settings & Other', 's': False},
    'q': {'d': 'Quit', 'f': None, 'c': 'Settings & Other', 's': False}
}

def show_usage():
    green, reset = "\033[92m", "\033[0m"
    print(f"\n{green}=== Combined Utility Tool - Guide ==={reset}")
    current_category = ""
    for key, opts in MENU_OPTIONS.items():
        if opts['c'] != current_category:
            current_category = opts['c']
            print(f"\n--- {green}{current_category}{reset} ---")
        print(f"  {key}. {green}{opts['d']}{reset}")
    print("\n")

def main():
    settings = load_settings()
    while True:
        print("\n" + "="*15 + " MAIN MENU " + "="*15 + f"\n (Version: {__version__})")
        categories = {}
        for key, opt in MENU_OPTIONS.items():
            cat = opt['c']
            if cat not in categories: categories[cat] = []
            categories[cat].append(f"  {key}. {opt['d']}")
        cat_order = ['File & Image Tools', 'Data & Calculation Tools', 'Settings & Other']
        for category in cat_order:
            if category in categories:
                print(f"\n--- {category} ---")
                for item in categories[category]: print(item)
        print("-" * 37)
        choice = input("Your choice: ").strip().lower()
        if choice == 'q': print("Thank you for using the tool. Goodbye!"); break
        elif choice == 'h': show_usage()
        elif choice in MENU_OPTIONS:
            opt = MENU_OPTIONS[choice]
            func = opt['f']
            if func:
                if opt['s']: settings = func(settings)
                else: func()
        else: print("Invalid choice. Please enter a key from the menu.")
        input("\nPress Enter to return to the menu...")

if __name__ == "__main__":
    install_and_check()
    # check_for_updates() # Kendi reponuza yüklediğinizde bu satırı aktif edebilirsiniz
    main()

# -*- coding: utf-8 -*-

# ATTENTION: When you update the code, you must increment this version number (e.g., "1.2").
__version__ = "1.7.1"

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
    required_packages = [
        'tqdm', 'openpyxl', 'Pillow', 'pillow-heif',
        'pandas', 'requests', 'xlrd', 'qrcode'
    ]
    
    # Check if pip is available
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("ERROR: pip is not available. Please ensure you have pip installed and in your system's PATH.")
        sys.exit(1)

    print("Checking for required libraries...")
    for package in required_packages:
        try:
            # Check if package can be imported
            __import__(package.replace('-', '_'))
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
    # Bu fonksiyonun içeriği değişmediği için aynı kalıyor.
    print("Checking for updates...")
    script_url = "https://raw.githubusercontent.com/hakanakaslanx-code/allone/refs/heads/main/allone.beta-en.py"
    current_script_name = os.path.basename(sys.argv[0])
    try:
        response = requests.get(script_url)
        response.raise_for_status()
        remote_script_content = response.text
        match = re.search(r"__version__\s*=\s*[\"'](.+?)[\"']", remote_script_content)
        if not match:
            print("Warning: Could not determine the remote script's version. Skipping update.")
            return
        remote_version = match.group(1)
        if remote_version > __version__:
            print(f"--> A new version ({remote_version}) is available. Current version: {__version__}. Updating...")
            new_script_path = current_script_name + ".new"
            with open(new_script_path, "w", encoding='utf-8') as f: f.write(remote_script_content)
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
            else:
                updater_script_path = "updater.sh"
                updater_content = f"""
#!/bin/bash
echo "Updating script... please wait."
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
    except requests.exceptions.RequestException as e:
        print(f"Warning: Update check failed. Could not connect to the server. Reason: {e}")
    except Exception as e:
        print(f"Warning: An unexpected error occurred during the update check: {e}")


logging.basicConfig(
    filename="tool_operations.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
SETTINGS_FILE = "settings.json"

# --- Helper Functions ---
def clean_file_path(file_path: str) -> str: return file_path.strip().strip('"').strip("'")
def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding='utf-8') as file: return json.load(file)
        except json.JSONDecodeError:
            print(f"Warning: '{SETTINGS_FILE}' is corrupt or empty. New settings will be created.")
            return {}
    return {}
def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding='utf-8') as file: json.dump(settings, file, indent=4)

def ask_for_folder_settings(existing_settings: dict) -> dict:
    print("\n--- Folder Settings for File Operations ---")
    print(f"Current Source Folder: {existing_settings.get('source_folder', 'Not set')}")
    print(f"Current Target Folder: {existing_settings.get('target_folder', 'Not set')}")
    update_settings = input("Do you want to update the settings? (y/n): ").strip().lower()
    if update_settings == "y":
        source_folder = input("Enter the new source folder path: ").strip()
        target_folder = input("Enter the new target folder path: ").strip()
        if not os.path.isdir(clean_file_path(source_folder)):
            print(f"Error: Source folder '{source_folder}' not found or is not a valid directory.")
            return existing_settings
        if not os.path.isdir(clean_file_path(target_folder)):
            create_folder = input(f"Warning: Target folder '{target_folder}' not found. Create it? (y/n): ").strip().lower()
            if create_folder == 'y':
                try:
                    os.makedirs(clean_file_path(target_folder))
                    print(f"Folder '{target_folder}' created successfully.")
                except OSError as e:
                    print(f"Error: Could not create folder. Reason: {e}")
                    return existing_settings
            else:
                print("Operation cancelled. Settings were not updated.")
                return existing_settings
        new_settings = {"source_folder": source_folder, "target_folder": target_folder}
        save_settings(new_settings)
        print("✅ Settings updated successfully.")
        return new_settings
    return existing_settings

def load_numbers_from_file(file_path: str) -> list:
    import pandas as pd
    try:
        p = clean_file_path(file_path)
        if not os.path.exists(p):
            print(f"Error: File '{file_path}' not found.")
            return []
        if p.lower().endswith(".csv"):
            df = pd.read_csv(p, header=None, usecols=[0], dtype=str, on_bad_lines='skip')
        elif p.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(p, header=None, usecols=[0], dtype=str, engine=None)
        else:
            df = pd.read_csv(p, header=None, usecols=[0], dtype=str, sep='\t', on_bad_lines='skip')
        return df[0].dropna().str.strip().tolist()
    except Exception as e:
        logging.error(f"Error reading numbers from file: {e}")
        print(f"Error: Could not read the file. Reason: {e}")
        return []

def parse_feet_inches(value_str: str):
    if not isinstance(value_str, str) or not value_str.strip(): return None
    try:
        s = value_str.strip().lower()
        s = s.replace("”", '"').replace("″", '"').replace("′", "'").replace("’", "'")
        s = s.replace("inches", '"').replace("inch", '"').replace("in", '"')
        s = re.sub(r"\s+", "", s)
        m_ft_in = re.fullmatch(r"(\d+(?:\.\d+)?)\'(\d+(?:\.\d+)?)?\"?", s)
        if m_ft_in:
            feet = float(m_ft_in.group(1))
            inches = float(m_ft_in.group(2)) if m_ft_in.group(2) else 0.0
            return feet + inches / 12.0
        m_in_only = re.fullmatch(r'(\d+(?:\.\d+)?)"', s)
        if m_in_only: return float(m_in_only.group(1)) / 12.0
        if "'" not in s and "." in s:
            parts = s.split(".", 1)
            feet = int(parts[0]) if parts[0] else 0
            inches = int(parts[1]) if parts[1] else 0
            return float(feet) + float(inches) / 12.0
        if re.fullmatch(r'\d+(?:\.\d+)?', s): return float(s)
        return None
    except (ValueError, TypeError): return None

def size_to_inches_wh(size_str: str):
    m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(size_str))
    if not m: return (None, None)
    w_ft = parse_feet_inches(m.group(1))
    h_ft = parse_feet_inches(m.group(2))
    if w_ft is None or h_ft is None: return (None, None)
    return (round(w_ft * 12, 2), round(h_ft * 12, 2))

def calculate_sqft(size_str: str):
    try:
        m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(size_str))
        if not m: return None
        width_ft = parse_feet_inches(m.group(1))
        height_ft = parse_feet_inches(m.group(2))
        if width_ft is None or height_ft is None: return None
        return round(width_ft * height_ft, 2)
    except Exception as e:
        logging.error(f"Error calculating sqft for '{size_str}': {e}")
        return None

# --- Main Modules ---
def rug_size_calculator():
    print("\n=== Rug Size Calculator (inches and sqft) ===")
    while True:
        user_input = input("Enter dimension ('width x height') (or 'q' to quit): ").strip()
        if user_input.lower() == 'q': break
        w_in, h_in = size_to_inches_wh(user_input)
        sqft = calculate_sqft(user_input)
        if w_in is not None: print(f"Result: Width: {w_in} in | Height: {h_in} in | Area: {sqft} sqft")
        else: print("Error: Invalid format. Please use 'width x height'.")

def bulk_sizes_from_sheet():
    import pandas as pd
    from tqdm import tqdm
    tqdm.pandas(desc="Calculating Dimensions")
    path = clean_file_path(input("Enter the path to the Excel/CSV file: ").strip())
    if not os.path.exists(path):
        print("Error: File not found."); return
    try:
        df = pd.read_excel(path, engine=None) if path.lower().endswith((".xlsx", ".xls")) else pd.read_csv(path)
    except Exception as e:
        print(f"Error reading file: {e}"); return
    print("\nColumns:", ", ".join(map(str, df.columns)))
    col_input = input("Enter the column name or letter for dimensions (e.g., Size or A): ").strip()
    selected_col_name = None
    if len(col_input) == 1 and col_input.isalpha():
        col_index = ord(col_input.upper()) - ord('A')
        if col_index < len(df.columns): selected_col_name = df.columns[col_index]
        else: print(f"Error: Column letter '{col_input}' is out of range."); return
    elif col_input in df.columns: selected_col_name = col_input
    else: print(f"Error: Column '{col_input}' not found."); return
    results = df[selected_col_name].progress_apply(lambda s: {'w_in': size_to_inches_wh(s)[0], 'h_in': size_to_inches_wh(s)[1], 'area': calculate_sqft(s)})
    df["Width_in"] = [r['w_in'] if r and r['w_in'] is not None else '' for r in results]
    df["Height_in"] = [r['h_in'] if r and r['h_in'] is not None else '' for r in results]
    df["Area_sqft"] = [r['area'] if r and r['area'] is not None else '' for r in results]
    base, _ = os.path.splitext(path)
    out_xlsx = f"{base}_with_sizes.xlsx"
    try:
        df.to_excel(out_xlsx, index=False)
        print(f"\n✅ Successfully saved to: {out_xlsx}")
    except Exception as e:
        print(f"\nCould not write to Excel ({e}). Saving as CSV...")
        out_csv = f"{base}_with_sizes.csv"
        df.to_csv(out_csv, index=False)
        print(f"✅ Successfully saved to: {out_csv}")

def format_numbers_from_file():
    file_path = input("Enter path to the Excel/CSV/TXT file: ").strip()
    numbers = load_numbers_from_file(file_path)
    if not numbers: print("No numbers to process or file could not be read."); return
    formatted_numbers = ",".join(numbers)
    print("\n--- Formatted Numbers ---\n", formatted_numbers)
    output_file = "formatted_numbers.txt"
    with open(output_file, "w", encoding='utf-8') as file: file.write(formatted_numbers)
    print(f"\n✅ Saved to '{output_file}'.")

def convert_heic_to_jpg_in_directory():
    from PIL import Image
    import pillow_heif
    from tqdm import tqdm
    directory = clean_file_path(input("Enter path to folder with HEIC files: ").strip())
    if not os.path.isdir(directory): print(f"Error: '{directory}' is not a valid directory."); return
    try:
        heic_files = [f for f in os.listdir(directory) if f.lower().endswith(".heic")]
        if not heic_files: print("No '.heic' files found."); return
        print(f"Found {len(heic_files)} HEIC files. Converting...")
        for heic_file in tqdm(heic_files, desc="HEIC -> JPG", unit="file"):
            source_path = os.path.join(directory, heic_file)
            target_path = os.path.splitext(source_path)[0] + ".jpg"
            try:
                heif_file = pillow_heif.read_heif(source_path)
                image = Image.frombytes(heif_file.mode, heif_file.size, heif_file.data, "raw")
                image.save(target_path, "JPEG")
                logging.info(f"Converted: {source_path} -> {target_path}")
            except Exception as e:
                logging.error(f"Conversion error for '{source_path}': {e}")
                print(f"\nError converting '{heic_file}': {e}")
        print("\n✅ All possible HEIC files converted.")
    except Exception as e:
        logging.error(f"General conversion error: {e}")
        print(f"\nError during conversion process: {e}")

def process_files_main(settings):
    from tqdm import tqdm
    source_folder = settings.get("source_folder")
    target_folder = settings.get("target_folder")
    if not source_folder or not target_folder:
        print("Source and target folders must be set first (Option 's')."); return
    file_path = input("Enter path to file with numbers (Excel/CSV/TXT): ").strip()
    target_numbers = load_numbers_from_file(file_path)
    if not target_numbers: print("No numbers to process."); return
    action_choice = input("Select action: Copy (c) / Move (m): ").strip().lower()
    if action_choice not in ['c', 'm']: print("Invalid choice."); return
    action = "copy" if action_choice == 'c' else "move"
    source_folder, target_folder = clean_file_path(source_folder), clean_file_path(target_folder)
    processed_files, missing_numbers = [], set(target_numbers)
    valid_extensions = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tiff'}
    all_files_in_source = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f))]
    num_to_file_map = {num: [file for file in all_files_in_source if num in file and os.path.splitext(file)[1].lower() in valid_extensions] for num in target_numbers}
    action_str = "Copying" if action == "copy" else "Moving"
    for num in tqdm(target_numbers, desc=f"{action_str} files"):
        if num_to_file_map.get(num):
            for file_to_process in num_to_file_map[num]:
                src_path = os.path.join(source_folder, file_to_process)
                dst_path = os.path.join(target_folder, file_to_process)
                try:
                    if action == "copy": shutil.copy2(src_path, dst_path)
                    else: shutil.move(src_path, dst_path)
                    processed_files.append(file_to_process)
                    missing_numbers.discard(num)
                except Exception as e:
                    logging.error(f"Failed to process '{file_to_process}': {e}")
                    print(f"\nError processing '{file_to_process}': {e}")
        else: logging.warning(f"No match found for number: {num}")
    print("\n--- Operation Summary ---")
    print(f"Files processed: {len(processed_files)}")
    if processed_files: print("Some processed files:", ", ".join(list(set(processed_files))[:5]) + '...')
    print(f"Identifiers not found: {len(missing_numbers)}")
    if missing_numbers: print("Unfound identifiers:", ", ".join(list(missing_numbers)))

def resize_and_compress_images():
    from PIL import Image
    from tqdm import tqdm
    print("\n=== Bulk Image Resizer and Compressor ===")
    source_dir = clean_file_path(input("Enter path to the folder with images: ").strip())
    if not os.path.isdir(source_dir): print(f"Error: Source directory '{source_dir}' not found."); return
    target_dir = os.path.join(source_dir, "resized")
    os.makedirs(target_dir, exist_ok=True)
    print(f"Resized images will be saved to: {target_dir}")
    try:
        max_width = int(input("Enter maximum width (e.g., 1920): ").strip())
        quality = int(input("Enter JPEG quality (1-95, default 75): ").strip() or 75)
        if not 1 <= quality <= 95: quality = 75; print("Invalid quality. Defaulting to 75.")
    except ValueError: print("Invalid number. Operation cancelled."); return
    valid_extensions = ('.jpg', '.jpeg', '.png', '.webp')
    image_files = [f for f in os.listdir(source_dir) if f.lower().endswith(valid_extensions)]
    if not image_files: print("No compatible images found."); return
    print(f"Found {len(image_files)} images to process.")
    for filename in tqdm(image_files, desc="Resizing images"):
        try:
            source_path = os.path.join(source_dir, filename)
            target_path = os.path.join(target_dir, filename)
            with Image.open(source_path) as img:
                if img.width > max_width:
                    ratio = max_width / float(img.width)
                    new_height = int(float(img.height) * ratio)
                    resample_filter = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                    resized_img = img.resize((max_width, new_height), resample_filter)
                else: resized_img = img.copy()
                if resized_img.mode in ("RGBA", "P"): resized_img = resized_img.convert("RGB")
                if os.path.splitext(filename)[1].lower() in ['.jpg', '.jpeg']:
                    resized_img.save(target_path, "JPEG", quality=quality, optimize=True)
                else: resized_img.save(target_path)
        except Exception as e:
            print(f"\nError processing {filename}: {e}"); logging.error(f"Failed to resize {filename}: {e}")
    print(f"\n✅ Processing complete. Resized images are in '{target_dir}'.")

def unit_converter():
    print("\n=== Unit Converter ===")
    print("Convert between 'cm', 'm', 'ft', 'in'.\nExamples: '182 cm to ft', '5\\'11\" to cm'")
    while True:
        user_input = input("Enter conversion (or 'q' to quit): ").strip().lower()
        if user_input == 'q': break
        match = re.match(r"^\s*(.+?)\s*(cm|m|ft|in)\s+to\s+(cm|m|ft|in)\s*$", user_input, re.IGNORECASE)
        if not match:
            print("Invalid format. Please use 'value unit to target_unit' (e.g., '100 cm to in')."); continue
        value_str, from_unit, to_unit = match.groups()
        INCH_TO_CM, FOOT_TO_CM, METER_TO_CM = 2.54, 30.48, 100
        cm_value = None
        try:
            if from_unit == 'ft':
                decimal_feet = parse_feet_inches(value_str)
                if decimal_feet is not None: cm_value = decimal_feet * FOOT_TO_CM
            else:
                value = float(value_str)
                if from_unit == 'cm': cm_value = value
                elif from_unit == 'm': cm_value = value * METER_TO_CM
                elif from_unit == 'in': cm_value = value * INCH_TO_CM
        except (ValueError, TypeError): pass
        if cm_value is None:
            print(f"Could not parse value '{value_str}'."); continue
        if to_unit == 'cm': result = f"{cm_value:.2f} cm"
        elif to_unit == 'm': result = f"{cm_value / METER_TO_CM:.2f} m"
        elif to_unit == 'in': result = f"{cm_value / INCH_TO_CM:.2f} in"
        elif to_unit == 'ft':
            total_inches = cm_value / INCH_TO_CM
            feet = int(total_inches // 12)
            inches = total_inches % 12
            result = f"{feet}' {inches:.2f}\""
        print(f"--> {value_str} {from_unit}  =  {result}")

def generate_qr_code():
    """Generates a QR code image from user-provided text."""
    import qrcode
    print("\n=== QR Code Generator ===")
    data_to_encode = input("Enter the text or URL for the QR code: ").strip()
    if not data_to_encode:
        print("No data provided. Operation cancelled."); return
    default_filename = "qrcode.png"
    output_filename = input(f"Enter filename (default: {default_filename}): ").strip() or default_filename
    if not output_filename.lower().endswith('.png'): output_filename += '.png'
    try:
        print("Generating QR code...")
        qr_img = qrcode.make(data_to_encode)
        qr_img.save(output_filename)
        print(f"✅ QR Code successfully saved as '{output_filename}'.")
    except Exception as e:
        print(f"An error occurred: {e}"); logging.error(f"QR Code generation failed: {e}")

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
        if choice == 'q':
            print("Thank you for using the tool. Goodbye!"); break
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
    # check_for_updates() # You can uncomment this if you host the new version
    main()

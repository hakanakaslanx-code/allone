# -*- coding: utf-8 -*-

# ATTENTION: When you update the code, you must increment this version number (e.g., "1.2").
__version__ = "1.4"

"""
This script combines multiple utility programs into a single interface:
1. File Copy/Move: Manages files based on specific numbers or identifiers.
2. Excel/Text Formatter: Formats data from a column into a single text file line.
3. HEIC to JPG Converter: Converts HEIC format images to JPG.
4. Rug Size Calculator: Calculates dimensions (inches and sqft) from feet/inch format.
5. Image Resizer/Compressor: Batch resizes and compresses images in a folder.
6. BULK Excel/CSV Rug Sizer: Reads a column of dimensions and adds Width_in / Height_in / Area_sqft.
7. Unit Converter: Converts between cm, m, feet, and inches.
"""

import sys
import subprocess
import os
import re
import requests # Required for the updater function

# --- Automatic Setup and Self-Update Mechanism ---

def install_and_check():
    """Checks for required libraries and installs them if they are missing."""
    required_packages = [
        'tqdm', 'openpyxl', 'Pillow', 'pillow-heif',
        'pandas', 'requests', 'xlrd'
    ]
    
    try:
        import tqdm, openpyxl, PIL, pillow_heif, pandas, requests, xlrd
    except ImportError:
        print("Some required libraries are missing. Starting installation...")
        for package in required_packages:
            try:
                __import__(package.replace('-', '_'))
            except ImportError:
                try:
                    print(f"Installing '{package}'...")
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
            else: # for macOS and Linux
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
            
            with open(updater_script_path, "w", encoding='utf-8') as f:
                f.write(updater_content)
            
            if sys.platform != 'win32':
                os.chmod(updater_script_path, 0o755)

            subprocess.Popen([updater_script_path])
            sys.exit(0)
        else:
            print(f"✅ Your code is up-to-date. (Version: {__version__})")

    except requests.exceptions.RequestException as e:
        print(f"Warning: Update check failed. Could not connect to the server. Reason: {e}")
    except Exception as e:
        print(f"Warning: An unexpected error occurred during the update check: {e}")

# --- Required Libraries and Settings (Post-Setup) ---
import shutil
import logging
import json
from tqdm import tqdm
from PIL import Image
import pillow_heif
import pandas as pd

logging.basicConfig(
    filename="tool_operations.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
SETTINGS_FILE = "settings.json"

# --- Helper Functions ---
def clean_file_path(file_path: str) -> str:
    return file_path.strip().strip('"').strip("'")

def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding='utf-8') as file:
                return json.load(file)
        except json.JSONDecodeError:
            print(f"Warning: '{SETTINGS_FILE}' is corrupt or empty. New settings will be created.")
            return {}
    return {}

def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding='utf-8') as file:
        json.dump(settings, file, indent=4)

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
            print(f"Warning: Target folder '{target_folder}' not found.")
            create_folder = input(f"Do you want to create the folder '{target_folder}'? (y/n): ").strip().lower()
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
    try:
        p = clean_file_path(file_path)
        if not os.path.exists(p):
            print(f"Error: File '{file_path}' not found.")
            return []
        
        if p.lower().endswith(".csv"):
            df = pd.read_csv(p, header=None, usecols=[0], dtype=str, on_bad_lines='skip')
        elif p.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(p, header=None, usecols=[0], dtype=str, engine=None)
        else: # Assume it's a .txt file
            df = pd.read_csv(p, header=None, usecols=[0], dtype=str, sep='\t', on_bad_lines='skip')
        
        return df[0].dropna().str.strip().tolist()

    except Exception as e:
        logging.error(f"Error reading numbers from file: {e}")
        print(f"Error: Could not read the file. Reason: {e}")
        return []

# --- Rug Dimension Parsing (IMPROVED) ---
def parse_feet_inches(value_str: str):
    if not isinstance(value_str, str) or not value_str.strip():
        return None
    try:
        s = value_str.strip().lower()
        s = s.replace("”", '"').replace("″", '"').replace("′", "'").replace("’", "'")
        s = s.replace("inches", '"').replace("inch", '"').replace("in", '"')
        s = re.sub(r"\s+", "", s)

        m_ft_in = re.fullmatch(r"(\d+(?:\.\d+)?)\'(\d+(?:\.\d+)?)?\"?", s)
        if m_ft_in:
            feet = float(m_ft_in.group(1))
            inches_str = m_ft_in.group(2)
            inches = float(inches_str) if inches_str else 0.0
            return feet + inches / 12.0

        m_in_only = re.fullmatch(r'(\d+(?:\.\d+)?)"', s)
        if m_in_only:
            return float(m_in_only.group(1)) / 12.0

        if "'" not in s and "." in s:
            parts = s.split(".", 1)
            feet = int(parts[0]) if parts[0] else 0
            inches = int(parts[1]) if parts[1] else 0
            return float(feet) + float(inches) / 12.0
            
        if re.fullmatch(r'\d+(?:\.\d+)?', s):
            return float(s)

        return None
    except (ValueError, TypeError):
        return None

def size_to_inches_wh(size_str: str):
    m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(size_str))
    if not m:
        return (None, None)
    w_ft = parse_feet_inches(m.group(1))
    h_ft = parse_feet_inches(m.group(2))
    if w_ft is None or h_ft is None:
        return (None, None)
    return (round(w_ft * 12, 2), round(h_ft * 12, 2))

def calculate_sqft(size_str: str):
    try:
        m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", str(size_str))
        if not m:
            return None
        width_ft = parse_feet_inches(m.group(1))
        height_ft = parse_feet_inches(m.group(2))
        if width_ft is None or height_ft is None:
            return None
        return round(width_ft * height_ft, 2)
    except Exception as e:
        logging.error(f"Error calculating sqft for '{size_str}': {e}")
        return None

# --- Main Modules ---
def rug_size_calculator():
    print("\n=== Rug Size Calculator (inches and sqft) ===")
    print('Examples: 11\'.5" x 5\'.6"  |  11.5x5.6  | 11\'x5\'5.5" ')
    while True:
        user_input = input("Enter dimension ('width x height') (or 'q' to quit): ").strip()
        if user_input.lower() == 'q':
            break
        w_in, h_in = size_to_inches_wh(user_input)
        sqft = calculate_sqft(user_input)
        if w_in is not None and h_in is not None and sqft is not None:
            print(f"Result: Width: {w_in} in | Height: {h_in} in | Area: {sqft} sqft")
        else:
            print("Error: Invalid format. Please enter in 'width x height' format.")

def bulk_sizes_from_sheet():
    path = clean_file_path(input("Enter the path to the Excel/CSV file to be processed: ").strip())
    if not os.path.exists(path):
        print("Error: File not found.")
        return
    try:
        if path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(path, engine=None)
        else:
            df = pd.read_csv(path)
    except Exception as e:
        print(f"Error: An error occurred while reading the file: {e}")
        return

    print("\nColumns in the file:", ", ".join(map(str, df.columns)))
    col_input = input("Enter the name OR the letter of the column with the dimensions (e.g., Size or A): ").strip()
    
    selected_col_name = None
    if len(col_input) == 1 and col_input.isalpha():
        col_index = ord(col_input.upper()) - ord('A')
        if col_index < len(df.columns):
            selected_col_name = df.columns[col_index]
            print(f"Column '{col_input}' selected, which corresponds to header '{selected_col_name}'.")
        else:
            print(f"Error: Column letter '{col_input}' is out of range for this file.")
            return
    else:
        if col_input in df.columns:
            selected_col_name = col_input
        else:
            print(f"Error: A column named '{col_input}' was not found.")
            return
    
    tqdm.pandas(desc="Calculating Dimensions")
    
    results = df[selected_col_name].progress_apply(lambda s: {
        'w_in': size_to_inches_wh(s)[0],
        'h_in': size_to_inches_wh(s)[1],
        'area': calculate_sqft(s)
    })

    df["Width_in"] = [r['w_in'] if r and r['w_in'] is not None else '' for r in results]
    df["Height_in"] = [r['h_in'] if r and r['h_in'] is not None else '' for r in results]
    df["Area_sqft"] = [r['area'] if r and r['area'] is not None else '' for r in results]

    base, _ = os.path.splitext(path)
    out_xlsx = f"{base}_with_sizes.xlsx"
    try:
        df.to_excel(out_xlsx, index=False)
        print(f"\n✅ Successfully saved to: {out_xlsx}")
    except Exception as e:
        print(f"\nCould not write to Excel ({e}). Saving as CSV instead...")
        out_csv = f"{base}_with_sizes.csv"
        df.to_csv(out_csv, index=False)
        print(f"✅ Successfully saved to: {out_csv}")


def format_numbers_from_file():
    file_path = input("Enter the path to the Excel/CSV/TXT file to be processed: ").strip()
    numbers = load_numbers_from_file(file_path)
    if not numbers:
        print("No numbers to process or file could not be read.")
        return
    formatted_numbers = ",".join(numbers)
    print("\n--- Formatted Numbers ---")
    print(formatted_numbers)
    output_file = "formatted_numbers.txt"
    with open(output_file, "w", encoding='utf-8') as file:
        file.write(formatted_numbers)
    print(f"\n✅ Saved to '{output_file}'.")

def convert_heic_to_jpg_in_directory():
    directory = input("Enter the path to the folder containing HEIC files: ").strip()
    clean_dir = clean_file_path(directory)
    if not os.path.isdir(clean_dir):
        print(f"Error: '{clean_dir}' is not a valid directory.")
        return
    try:
        heic_files = [f for f in os.listdir(clean_dir) if f.lower().endswith(".heic")]
        if not heic_files:
            print("No files with the '.heic' extension were found in this folder.")
            return
        
        print(f"Found {len(heic_files)} HEIC files. Starting conversion...")
        for heic_file in tqdm(heic_files, desc="HEIC -> JPG", unit="file"):
            source_path = os.path.join(clean_dir, heic_file)
            target_path = os.path.splitext(source_path)[0] + ".jpg"
            try:
                heif_file = pillow_heif.read_heif(source_path)
                image = Image.frombytes(heif_file.mode, heif_file.size, heif_file.data, "raw")
                image.save(target_path, "JPEG")
                logging.info(f"Converted: {source_path} -> {target_path}")
            except Exception as e:
                logging.error(f"Conversion error for '{source_path}': {e}")
                print(f"\nError: Could not convert '{heic_file}'. Reason: {e}")
        print("\n✅ All HEIC files converted successfully.")
    except Exception as e:
        logging.error(f"General conversion error: {e}")
        print(f"\nError: The conversion process could not be completed. Reason: {e}")

def process_files(source_folder, target_folder, target_numbers, action):
    processed_files, missing_numbers = [], set(target_numbers)
    valid_extensions = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tiff'}
    
    all_files_in_source = [f for f in os.listdir(source_folder) if os.path.isfile(os.path.join(source_folder, f))]
    
    num_to_file_map = {num: [] for num in target_numbers}
    for file in all_files_in_source:
        for num in target_numbers:
            if num in file and os.path.splitext(file)[1].lower() in valid_extensions:
                num_to_file_map[num].append(file)
    
    action_str = "copying" if action == "copy" else "moving"
    for num in tqdm(target_numbers, desc=f"Files are being {action_str}"):
        if num in num_to_file_map and num_to_file_map[num]:
            for file_to_process in num_to_file_map[num]:
                src_path = os.path.join(source_folder, file_to_process)
                dst_path = os.path.join(target_folder, file_to_process)
                try:
                    if action == "copy":
                        shutil.copy2(src_path, dst_path)
                    else:
                        shutil.move(src_path, dst_path)
                    processed_files.append(file_to_process)
                    missing_numbers.discard(num)
                except Exception as e:
                    logging.error(f"Failed to process '{file_to_process}': {e}")
                    print(f"\nError: Could not process '{file_to_process}'. Reason: {e}")
        else:
            logging.warning(f"No matching file found for number: {num}")

    return processed_files, list(missing_numbers)

def process_files_main(settings):
    source_folder = settings.get("source_folder")
    target_folder = settings.get("target_folder")
    if not source_folder or not target_folder:
        print("Please set the source and target folders first (Main Menu -> Option 's').")
        return

    file_path = input("Enter the path to the file with numbers (Excel/CSV/TXT): ").strip()
    target_numbers = load_numbers_from_file(file_path)
    if not target_numbers:
        print("No numbers to process or the file could not be read.")
        return

    action_choice = input("Select action: Copy (c) / Move (m): ").strip().lower()
    if action_choice not in ['c', 'm']:
        print("Invalid choice. Please enter 'c' or 'm'.")
        return
    
    action = "copy" if action_choice == 'c' else "move"
    
    processed, missing = process_files(source_folder, target_folder, target_numbers, action)
    
    print("\n--- Operation Summary ---")
    print(f"Number of files processed: {len(processed)}")
    if processed:
        print("Some processed files:", ", ".join(list(set(processed))[:10]) + ('...' if len(set(processed)) > 10 else ''))
        
    print(f"Number of identifiers not found: {len(missing)}")
    if missing:
        print("Unfound identifiers:", ", ".join(missing))

def resize_and_compress_images():
    print("\n=== Bulk Image Resizer and Compressor ===")
    source_dir = clean_file_path(input("Enter the path to the folder with images: ").strip())
    
    if not os.path.isdir(source_dir):
        print(f"Error: Source directory '{source_dir}' not found.")
        return

    target_dir = os.path.join(source_dir, "resized")
    os.makedirs(target_dir, exist_ok=True)
    print(f"Resized images will be saved to: {target_dir}")

    try:
        max_width = int(input("Enter the maximum width for the images (e.g., 1920): ").strip())
        quality = int(input("Enter the JPEG quality (1-95, default is 75): ").strip() or 75)
        if not 1 <= quality <= 95:
            quality = 75
            print("Invalid quality value. Defaulting to 75.")
    except ValueError:
        print("Invalid number entered. Operation cancelled.")
        return

    valid_extensions = ('.jpg', '.jpeg', '.png', '.webp')
    image_files = [f for f in os.listdir(source_dir) if f.lower().endswith(valid_extensions)]

    if not image_files:
        print("No compatible image files found in the directory.")
        return

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
                else:
                    resized_img = img.copy()

                if resized_img.mode in ("RGBA", "P"):
                    resized_img = resized_img.convert("RGB")
                
                file_ext = os.path.splitext(filename)[1].lower()
                if file_ext in ['.jpg', '.jpeg']:
                    resized_img.save(target_path, "JPEG", quality=quality, optimize=True)
                else:
                    resized_img.save(target_path)

        except Exception as e:
            print(f"\nError processing {filename}: {e}")
            logging.error(f"Failed to resize {filename}: {e}")
    
    print(f"\n✅ Processing complete. Resized images are in the '{target_dir}' folder.")

# --- Unit Converter Helpers ---
INCH_TO_CM = 2.54
FOOT_TO_CM = 30.48
METER_TO_CM = 100

def convert_to_cm(value_str, unit):
    """Converts a given value from its unit to centimeters."""
    try:
        if unit == 'ft':
            decimal_feet = parse_feet_inches(value_str)
            return decimal_feet * FOOT_TO_CM if decimal_feet is not None else None
        
        value = float(value_str)
        if unit == 'cm': return value
        if unit == 'm': return value * METER_TO_CM
        if unit == 'in': return value * INCH_TO_CM
        return None
    except (ValueError, TypeError):
        return None

def convert_from_cm(cm_value, unit):
    """Converts a value from centimeters to the target unit."""
    if unit == 'cm': return f"{cm_value:.2f} cm"
    if unit == 'm': return f"{cm_value / METER_TO_CM:.2f} m"
    if unit == 'in': return f"{cm_value / INCH_TO_CM:.2f} in"
    if unit == 'ft':
        total_inches = cm_value / INCH_TO_CM
        feet = int(total_inches // 12)
        inches = total_inches % 12
        return f"{feet}' {inches:.2f}\""
    return "Invalid target unit"

def unit_converter():
    print("\n=== Unit Converter ===")
    print("Convert between 'cm', 'm', 'ft', 'in'.")
    print("Examples: '182 cm to ft', '5\\'11\" to cm', '2.5 m to in'")
    
    while True:
        user_input = input("Enter conversion (or 'q' to quit): ").strip().lower()
        if user_input == 'q':
            break

        match = re.match(r"^\s*(.+?)\s*(cm|m|ft|in)\s+to\s+(cm|m|ft|in)\s*$", user_input, re.IGNORECASE)
        
        if not match:
            print("Invalid format. Please use 'value unit to target_unit' (e.g., '100 cm to in').")
            continue
            
        value_str, from_unit, to_unit = match.groups()

        cm_value = convert_to_cm(value_str, from_unit)

        if cm_value is None:
            print(f"Could not parse value '{value_str}'. Please check your input.")
            continue
            
        result = convert_from_cm(cm_value, to_unit)
        
        original_input_formatted = f"{value_str} {from_unit}"
        print(f"--> {original_input_formatted}  =  {result}")

# --- Central Menu Configuration ---
MENU_OPTIONS = {
    # File & Image Tools
    '1': {'description': 'Copy/Move Files by List', 'function': process_files_main, 'category': 'File & Image Tools', 'requires_settings': True},
    '2': {'description': 'Convert HEIC to JPG', 'function': convert_heic_to_jpg_in_directory, 'category': 'File & Image Tools', 'requires_settings': False},
    '3': {'description': 'Batch Image Resizer', 'function': resize_and_compress_images, 'category': 'File & Image Tools', 'requires_settings': False},
    
    # Data & Calculation Tools
    '4': {'description': 'Format Numbers from File', 'function': format_numbers_from_file, 'category': 'Data & Calculation Tools', 'requires_settings': False},
    '5': {'description': 'Rug Size Calculator (Single)', 'function': rug_size_calculator, 'category': 'Data & Calculation Tools', 'requires_settings': False},
    '6': {'description': 'BULK Process Rug Sizes from File', 'function': bulk_sizes_from_sheet, 'category': 'Data & Calculation Tools', 'requires_settings': False},
    '7': {'description': 'Unit Converter (cm, m, ft, in)', 'function': unit_converter, 'category': 'Data & Calculation Tools', 'requires_settings': False},
    
    # Settings & Other
    's': {'description': 'Set Folders for File Operations', 'function': ask_for_folder_settings, 'category': 'Settings & Other', 'requires_settings': True},
    'h': {'description': 'Help / Guide', 'function': None, 'category': 'Settings & Other', 'requires_settings': False},
    'q': {'description': 'Quit', 'function': None, 'category': 'Settings & Other', 'requires_settings': False}
}

def show_usage():
    green = "\033[92m"
    reset = "\033[0m"
    print(f"\n{green}=== Combined Utility Tool - Guide ==={reset}")

    current_category = ""
    for key, options in MENU_OPTIONS.items():
        if options['category'] != current_category:
            current_category = options['category']
            print(f"\n--- {green}{current_category}{reset} ---")
        
        print(f"  {key}. {green}{options['description']}{reset}")
    
    print("\n")

def main():
    """The main function that runs the menu loop and handles user choices."""
    settings = load_settings()

    while True:
        print("\n" + "="*15 + " MAIN MENU " + "="*15)
        print(f" (Version: {__version__})")
        
        categories = {}
        for key, option in MENU_OPTIONS.items():
            cat = option['category']
            if cat not in categories:
                categories[cat] = []
            
            # Help and Quit don't need a number, but others do.
            if key in ('h', 'q', 's'):
                 categories[cat].append(f"  {key}. {option['description']}")
            else:
                 categories[cat].append(f"{key}. {option['description']}")

        # Ensure consistent order
        cat_order = ['File & Image Tools', 'Data & Calculation Tools', 'Settings & Other']
        for category in cat_order:
            if category in categories:
                print(f"\n--- {category} ---")
                for item in categories[category]:
                    print(item)
        
        print("-" * 37)
        choice = input("Your choice: ").strip().lower()

        if choice == 'q':
            print("Thank you for using the tool. Goodbye!")
            break
        elif choice == 'h':
            show_usage()
        elif choice in MENU_OPTIONS:
            selected_option = MENU_OPTIONS[choice]
            function_to_call = selected_option['function']
            
            if selected_option['requires_settings']:
                settings = function_to_call(settings)
            else:
                function_to_call()
        else:
            print("Invalid choice. Please enter one of the keys from the menu.")
        
        input("\nPress Enter to return to the menu...")


if __name__ == "__main__":
    install_and_check()
    check_for_updates()
    main()

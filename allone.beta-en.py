# -*- coding: utf-8 -*-

# ATTENTION: When you update the code, you must increment this version number (e.g., "1.2").
__version__ = "1.1"

"""
This script combines multiple utility programs into a single interface:
1. File Copy/Move: Manages files based on specific numbers or identifiers.
2. Excel/Text Formatter: Formats data from a column into a single text file line.
3. HEIC to JPG Converter: Converts HEIC format images to JPG.
4. Carpet Size Calculator: Calculates dimensions (inches and sqft) from feet/inch format.
5. Google Maps Data Scraper: Scrapes business data from Google Maps using Playwright.
6. BULK Excel/CSV Carpet Sizer: Reads a column of dimensions and adds Width_in / Height_in / Area_sqft.
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
        'pandas', 'requests', 'beautifulsoup4', 'playwright',
        'xlrd' # Added for legacy .xls Excel format support
    ]
    
    try:
        import tqdm, openpyxl, PIL, pillow_heif, pandas, requests, bs4, playwright, xlrd
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

    print("\nInstalling/updating Playwright browsers (if necessary)...")
    try:
        subprocess.check_call([sys.executable, "-m", "playwright", "install"])
    except subprocess.CalledProcessError:
        print("ERROR: Failed to install Playwright browsers. Please run manually: playwright install")
        sys.exit(1)
    
    print("\n✅ Setup checks complete.")

def check_for_updates():
    """
    Checks for a new version of the script on Google Drive and self-updates if one is found.
    """
    print("Checking for updates...")
    
    script_url = "https://drive.google.com/uc?export=download&id=1T00kGXIW-tBHa9n00xRsFcKOOhC7ZuZA"
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
import time
from dataclasses import dataclass, asdict, field
from tqdm import tqdm
from PIL import Image
import pillow_heif
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import pandas as pd
from bs4 import BeautifulSoup

logging.basicConfig(
    filename="tool_operations.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
SETTINGS_FILE = "settings.json"

# --- Google Maps Data Classes ---
@dataclass
class Business:
    name: str = None
    address: str = None
    website: str = None
    phone_number: str = None
    reviews_count: int = None
    reviews_average: float = None
    latitude: float = None
    longitude: float = None
    facebook: str = None
    instagram: str = None
    email: str = None

@dataclass
class BusinessList:
    business_list: list[Business] = field(default_factory=list)
    save_at = 'output'

    def dataframe(self):
        return pd.json_normalize((asdict(b) for b in self.business_list), sep="_").fillna("")

    def save_to_excel(self, filename):
        os.makedirs(self.save_at, exist_ok=True)
        self.dataframe().to_excel(f"{self.save_at}/{filename}.xlsx", index=False)
        print(f"✅ Data successfully saved to '{self.save_at}/{filename}.xlsx'")

    def save_to_csv(self, filename):
        os.makedirs(self.save_at, exist_ok=True)
        self.dataframe().to_csv(f"{self.save_at}/{filename}.csv", index=False)
        print(f"✅ Data successfully saved to '{self.save_at}/{filename}.csv'")

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

# --- Carpet Dimension Parsing (IMPROVED) ---
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
    print("\n=== Carpet Size Calculator (inches and sqft) ===")
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

# --- Google Maps Functions ---
def extract_coordinates_from_url(url: str) -> tuple[float | None, float | None]:
    match = re.search(r"/@(-?\d+\.\d+),(-?\d+\.\d+)", url)
    if match:
        return float(match.group(1)), float(match.group(2))
    return None, None

def get_social_links_and_email(website_url):
    fb, ig, email = "", "", ""
    if not website_url: return fb, ig, email

    try:
        url = website_url if website_url.startswith("http") else "https://" + website_url
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        res = requests.get(url, timeout=10, headers=headers, allow_redirects=True)
        
        soup = BeautifulSoup(res.text, "html.parser")
        links = {a['href'] for a in soup.find_all("a", href=True)}
        
        fb = next((link for link in links if "facebook.com" in link), "")
        ig = next((link for link in links if "instagram.com" in link), "")
        
        emails = re.findall(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', res.text, re.IGNORECASE)
        if not emails:
             emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', res.text)
        
        email = emails[0] if emails else ""
        if email: print(f"  -> 📧 Email found: {email}")
            
    except Exception as e:
        print(f"  -> ❌ Error while scanning website ({website_url}): {e}")
        
    return fb, ig, email

def scrape_google_maps():
    print("\nWARNING: This function depends on the Google Maps interface and may break in the future.")
    search_term = input("Enter the search term for Google Maps (e.g., restaurants in New York): ").strip()
    total_str = input("Maximum number of results to scrape? (Leave empty to try for all): ").strip()
    total = int(total_str) if total_str.isdigit() and total_str else 1_000_000

    if not search_term:
        print("Search term not provided. Operation cancelled.")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        try:
            page.goto("https://www.google.com/maps", timeout=60000)
            page.locator('//input[@id="searchboxinput"]').fill(search_term)
            page.keyboard.press("Enter")
            
            print(f"🔍 Searching for '{search_term}'... Please wait for results to load.")
            results_panel_selector = 'div[role="feed"]'
            page.wait_for_selector(results_panel_selector, timeout=30000)

            prev_count = 0
            while True:
                listings_locator = page.locator(f'{results_panel_selector} > div > div[role="article"]')
                curr_count = listings_locator.count()
                
                print(f"🔄 Number of loaded listings: {curr_count}")
                if curr_count >= total or curr_count == prev_count:
                    print(f"✅ Scrolling complete. Found {curr_count} total results.")
                    break
                
                prev_count = curr_count
                page.evaluate(f"document.querySelector('{results_panel_selector}').scrollTop = document.querySelector('{results_panel_selector}').scrollHeight")
                time.sleep(3)
            
            listings = listings_locator.all()[:total]
            print(f"📦 Scraping information for {len(listings)} listings.")
            
            business_list = BusinessList()
            for i, listing in enumerate(listings):
                try:
                    listing.click()
                    print(f"\n➡️ Processing listing {i+1}/{len(listings)}...")
                    page.wait_for_selector('h1', timeout=10000)
                    
                    b = Business()
                    b.name = page.locator('h1').first.inner_text()
                    b.address = page.locator('button[data-item-id="address"] div.fontBodyMedium').first.inner_text(timeout=2000)
                    b.website = page.locator('a[data-item-id="authority"] div.fontBodyMedium').first.inner_text(timeout=2000)
                    b.phone_number = page.locator('button[data-item-id*="phone:tel:"] div.fontBodyMedium').first.inner_text(timeout=2000)
                    
                    review_text = page.locator('div.F7nice').first.inner_text(timeout=2000)
                    match = re.search(r'(\d[\.,]\d)\s+\((\d+)\)', review_text)
                    if match:
                        b.reviews_average = float(match.group(1).replace(',', '.'))
                        b.reviews_count = int(match.group(2).replace(',', ''))

                    b.latitude, b.longitude = extract_coordinates_from_url(page.url)
                    
                    if b.website:
                        print(f"  -> 🌐 Scanning website: {b.website}")
                        b.facebook, b.instagram, b.email = get_social_links_and_email(b.website)
                        
                    business_list.business_list.append(b)

                except PlaywrightTimeoutError:
                    print(f"  -> ⚠️ Some details (address, phone, etc.) for this listing could not be found or loaded.")
                    try: 
                        b.name = listing.get_attribute('aria-label')
                        if b.name: business_list.business_list.append(b)
                    except: pass
                except Exception as e:
                    logging.error(f"Error on listing {i+1}: {e}")

            if business_list.business_list:
                filename = f"google_maps_data_{search_term.strip().replace(' ', '_')}"
                business_list.save_to_excel(filename)
                business_list.save_to_csv(filename)
                
        except Exception as e:
            print(f"\nERROR: A problem occurred during the main process. {e}")
            logging.critical(f"Maps scraper failed: {e}")
        finally:
            print("Closing browser.")
            browser.close()

def show_usage():
    green = "\033[92m"
    reset = "\033[0m"
    usage_text = f"""
    {green}=== Combined Utility Tool - Guide ==={reset}

    1. {green}Copy/Move Files{reset}: Transfers image files from one folder to another based on a list of identifiers.
    2. {green}Format Numbers from File{reset}: Reads the first column from a file and formats the values into a single comma-separated line.
    3. {green}Convert HEIC to JPG{reset}: Converts all Apple HEIC format images in a folder to the universal JPG format.
    4. {green}Carpet Size Calculator (Single){reset}: Calculates the inches and sqft for a single dimension string.
    5. {green}Scrape Google Maps Data{reset}: Scrapes business info from Google Maps for a given search term.
    6. {green}BULK Process Carpet Sizes from File{reset}: Processes an entire column of dimensions in an Excel or CSV file.
    
    {green}--- Settings & Other ---{reset}
    s. {green}Set Folders for File Operations{reset}: Defines the source and target folders for the Copy/Move tool.
    h. {green}Help / Guide{reset}: Displays this help menu again.
    q. {green}Quit{reset}: Exits the program.
    """
    print(usage_text)

def main():
    """The main function that runs the menu loop and handles user choices."""
    while True:
        print("\n" + "="*15 + " MAIN MENU " + "="*15)
        print(f" (Version: {__version__})")
        print("1. Copy/Move Files")
        print("2. Format Numbers from File")
        print("3. Convert HEIC to JPG")
        print("4. Carpet Size Calculator (Single)")
        print("5. Scrape Google Maps Data")
        print("6. BULK Process Carpet Sizes from File")
        print("-----------------------------------")
        print("s. Set Folders   h. Help   q. Quit")
        choice = input("Your choice: ").strip().lower()

        settings = load_settings()

        if choice == "1": process_files_main(settings)
        elif choice == "2": format_numbers_from_file()
        elif choice == "3": convert_heic_to_jpg_in_directory()
        elif choice == "4": rug_size_calculator()
        elif choice == "5": scrape_google_maps()
        elif choice == "6": bulk_sizes_from_sheet()
        elif choice == "s": settings = ask_for_folder_settings(settings)
        elif choice == "h": show_usage()
        elif choice == "q":
            print("Thank you for using the tool. Goodbye!")
            break
        else:
            print("Invalid choice. Please enter one of the numbers or letters from the menu.")

if __name__ == "__main__":
    install_and_check()
    check_for_updates()
    main()
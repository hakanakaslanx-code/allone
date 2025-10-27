# backend_logic.py
import os
import re
import sys
import logging
import shutil
from pathlib import Path
from typing import Iterable, List, Optional, Set

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pillow_heif
import qrcode
import barcode
from barcode.writer import ImageWriter

# This file contains no Tkinter code.

# --- Helper Functions ---


def get_resource_path(*path_parts: str) -> Optional[str]:
    """Return an absolute path for bundled resources (PyInstaller friendly).

    The function searches several likely locations including the PyInstaller
    temporary directory (``sys._MEIPASS``), the package directory, and optional
    ``resources`` sub-directories. ``None`` is returned if the resource cannot
    be resolved.
    """

    if not path_parts:
        return None

    if len(path_parts) == 1 and isinstance(path_parts[0], (list, tuple)):
        path_parts = tuple(path_parts[0])  # type: ignore[assignment]

    relative_path = os.path.join(*map(str, path_parts))

    if os.path.isabs(relative_path):
        return relative_path if os.path.exists(relative_path) else None

    search_roots: List[Path] = []

    try:
        search_roots.append(Path(sys._MEIPASS))  # type: ignore[attr-defined]
    except AttributeError:
        pass

    module_dir = Path(__file__).resolve().parent
    search_roots.extend([module_dir, module_dir.parent])

    seen: Set[str] = set()
    for root in search_roots:
        for candidate in (
            root / relative_path,
            root / "resources" / relative_path,
        ):
            candidate_str = str(candidate)
            if candidate_str in seen:
                continue
            seen.add(candidate_str)
            if candidate.exists():
                return candidate_str

    return None

def clean_file_path(file_path: str) -> str:
    """Cleans quotes and spaces from the file path string."""
    return file_path.strip().strip('"').strip("'")


def _normalize_rug_number(value) -> str:
    """Normalize rug number values from spreadsheets for reliable comparison."""

    if value is None:
        return ""

    try:
        if pd.isna(value):  # type: ignore[arg-type]
            return ""
    except Exception:
        pass

    text = ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            text = str(int(value))
        else:
            text = str(value).strip()
    else:
        text = str(value).strip()

    if not text:
        return ""

    # Collapse values like "123.0" to "123" but keep leading zeros intact for
    # explicit string inputs such as "00123".
    if "." in text and text.replace(".", "", 1).isdigit():
        try:
            number_value = float(text)
            if number_value.is_integer():
                return str(int(number_value))
        except ValueError:
            pass

    return text


def normalize_rug_number(value) -> str:
    """Public helper to normalize rug number strings."""

    return _normalize_rug_number(value)


def _extract_rug_numbers(df: pd.DataFrame) -> List[str]:
    """Extract rug numbers from the most likely column in a dataframe."""

    if df.empty:
        return []

    candidate_column = None

    # Try to identify a column that contains "rug" in its name.
    for column in df.columns:
        if isinstance(column, tuple):
            column_name = " ".join(str(part) for part in column if part)
        else:
            column_name = str(column)
        if "rug" in column_name.lower():
            candidate_column = column
            break

    if candidate_column is None:
        candidate_column = df.columns[0]

    series = df[candidate_column]
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]

    values: Iterable = series.tolist()
    normalized = [_normalize_rug_number(value) for value in values]
    return [value for value in normalized if value]


def load_rug_numbers_from_file(file_path: str) -> List[str]:
    """Load rug numbers from an Excel/CSV file, returning normalized strings."""

    cleaned_path = clean_file_path(file_path)
    if not cleaned_path:
        raise ValueError("Master list file path is empty.")
    if not os.path.exists(cleaned_path):
        raise FileNotFoundError(f"File not found: {cleaned_path}")

    try:
        if cleaned_path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(cleaned_path, dtype=str)
        else:
            df = pd.read_csv(cleaned_path, dtype=str, on_bad_lines="skip")
    except Exception as exc:
        raise RuntimeError(f"Failed to read file '{cleaned_path}': {exc}") from exc

    return _extract_rug_numbers(df)


def compare_rug_numbers_task(
    sold_file: str,
    master_file: str,
    log_callback,
    completion_callback,
    result_callback,
):
    """Compare sold list rug numbers with a master list and report results."""

    log_callback("Starting rug number comparison...")

    try:
        sold_numbers = load_rug_numbers_from_file(sold_file)
        master_numbers = load_rug_numbers_from_file(master_file)
    except Exception as exc:
        error_message = f"{exc}"
        log_callback(f"Error: {error_message}")
        completion_callback("Error", error_message)
        return

    if not sold_numbers:
        log_callback("Sold list does not contain any rug numbers to compare.")

    master_set = set(master_numbers)
    found = sorted({number for number in sold_numbers if number in master_set})
    missing = sorted({number for number in sold_numbers if number not in master_set})

    result_callback(found, missing)

    summary = f"Comparison complete. Found: {len(found)} | Missing: {len(missing)}"
    log_callback(summary)
    completion_callback("Complete", "Rug number comparison completed.")

def parse_feet_inches(value_str: str):
    """Converts various string formats (e.g., 5'2", 5.2', 8") to a decimal feet value."""
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
    """Converts a dimension string (e.g., "5'2" x 8'") to a (width_in, height_in) tuple."""
    if not isinstance(s, str): return (None, None)
    m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", s)
    if not m: return (None, None)
    w = parse_feet_inches(m.group(1)); h = parse_feet_inches(m.group(2))
    return (round(w*12, 2), round(h*12, 2)) if w is not None and h is not None else (None, None)

def calculate_sqft(s: str):
    """Calculates the square footage from a dimension string."""
    if not isinstance(s, str): return None
    try:
        m = re.match(r"^\s*(.+?)\s*[xX×]\s*(.+?)\s*$", s)
        if not m: return None
        w, h = parse_feet_inches(m.group(1)), parse_feet_inches(m.group(2))
        return round(w * h, 2) if w is not None and h is not None else None
    except: return None
    
def create_label_image(code_image, label_info, bottom_text=""):
    """Creates a label image for Dymo printers with a centered code and optional text."""
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

def convert_units_logic(input_string):
    """Takes a conversion string and returns the result string."""
    i = input_string.lower()
    if not i:
        return ""
    
    m = re.match(r"^\s*(.+?)\s*(cm|m|ft|in)\s+to\s+(cm|m|ft|in)\s*$", i, re.I)
    if not m:
        return "Invalid Format"

    v_str, fu, tu = m.groups()
    cm = None
    try:
        if fu == 'ft':
            cm = parse_feet_inches(v_str) * 30.48 if parse_feet_inches(v_str) else None
        else:
            val = float(v_str)
            cm = val if fu == 'cm' else val * 100 if fu == 'm' else val * 2.54 if fu == 'in' else None
    except:
        pass
    
    if cm is None:
        return f"Could not parse '{v_str}'."
    
    res = ""
    if tu == 'cm': res = f"{cm:.2f} cm"
    elif tu == 'm': res = f"{cm / 100:.2f} m"
    elif tu == 'in': res = f"{cm / 2.54:.2f} in"
    elif tu == 'ft':
        total_in = cm / 2.54
        res = f"{int(total_in // 12)}' {total_in % 12:.2f}\""
    
    return f"--> {res}"

# --- Task Functions ---

def process_files_task(src, tgt, nums_f, action, log_callback, completion_callback):
    """Finds and copies or moves files based on a list."""
    log_callback(f"Starting file {action} process...")
    try:
        p = clean_file_path(nums_f)
        if not os.path.exists(p):
            log_callback(f"Error: Numbers file not found at '{p}'")
            completion_callback("Error", f"Numbers file not found at '{p}'")
            return
        if p.lower().endswith((".xlsx",".xls")):
            df = pd.read_excel(p, header=None, usecols=[0], dtype=str)
        else:
            df = pd.read_csv(p, header=None, usecols=[0], dtype=str, on_bad_lines='skip', sep=r'\s+|\t|,', engine='python')
        nums = df[0].dropna().str.strip().tolist()
    except Exception as e:
        log_callback(f"Error reading numbers file: {e}")
        completion_callback("Error", f"Could not read the numbers file: {e}")
        return

    if not nums:
        log_callback("No numbers found in the file to process."); return
    
    proc, missing = [], set(nums)
    exts = {'.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tiff'}
    files = [f for f in os.listdir(src) if os.path.isfile(os.path.join(src, f))]
    map_ = {n: [f for f in files if n in f and os.path.splitext(f)[1].lower() in exts] for n in nums}
    
    total_files = len(nums)
    log_callback(f"Processing {total_files} items from list...")
    for i, n in enumerate(nums):
        if (i + 1) % (total_files // 10 or 1) == 0:
            percentage = (i + 1) * 100 / total_files
            log_callback(f"  ...Progress: {percentage:.0f}% ({i + 1}/{total_files})")
            
        if map_.get(n):
            for f in map_[n]:
                try:
                    if action == "copy": shutil.copy2(os.path.join(src, f), os.path.join(tgt, f))
                    else: shutil.move(os.path.join(src, f), os.path.join(tgt, f))
                    proc.append(f)
                    if n in missing: missing.remove(n)
                except Exception as e: log_callback(f"Error processing '{f}': {e}")
        else:
            logging.warning(f"No match for: {n}")
    
    summary = f"--- Summary ---\nProcessed: {len(proc)}\nNot Found: {len(missing)}"
    log_callback(summary)
    if missing: log_callback(f"Unfound: {', '.join(list(missing))}")
    completion_callback("Complete", f"File {action} process finished. See log for details.")


def convert_heic_task(folder, log_callback, completion_callback):
    """Converts all HEIC files in a folder to JPG."""
    log_callback("Starting HEIC to JPG conversion...")
    try:
        files = [f for f in os.listdir(folder) if f.lower().endswith(".heic")]
        if not files:
            log_callback("No HEIC files found."); return
        
        total_files = len(files)
        for i, f in enumerate(files):
            log_callback(f"Converting ({i+1}/{total_files}): {f}")
            src, dst = os.path.join(folder, f), f"{os.path.splitext(os.path.join(folder, f))[0]}.jpg"
            try:
                heif = pillow_heif.read_heif(src)
                img = Image.frombytes(heif.mode, heif.size, heif.data, "raw")
                img.save(dst, "JPEG")
            except Exception as e: log_callback(f"Error converting '{f}': {e}")
        
        log_callback("\n✅ Conversion complete.")
        completion_callback("Success", "HEIC conversion is complete.")
    except Exception as e:
        log_callback(f"An error occurred: {e}")
        completion_callback("Error", f"An error occurred: {e}")

def resize_images_task(src_folder, mode, value, quality, log_callback, completion_callback):
    """Resizes all images in a folder by width or percentage."""
    tgt_folder = os.path.join(src_folder, "resized")
    os.makedirs(tgt_folder, exist_ok=True)
    log_callback(f"Resized images will be saved in: {tgt_folder}")
    files = [f for f in os.listdir(src_folder) if f.lower().endswith(('.jpg','.jpeg','.png','.webp'))]
    if not files:
        log_callback("No compatible images found."); return
    
    log_callback(f"Starting resize process for {len(files)} images...")
    for i, f in enumerate(files):
        if (i + 1) % 20 == 0: 
            log_callback(f"  ...resizing image {i+1} of {len(files)}")
        try:
            with Image.open(os.path.join(src_folder, f)) as img:
                ow, oh = img.size
                nw, nh = ow, oh
                if mode == "width":
                    if ow > value:
                        ratio = value / float(ow)
                        nw, nh = value, int(float(oh) * ratio)
                else:
                    nw, nh = int(ow * value / 100), int(oh * value / 100)
                
                if nw < 1: nw = 1
                if nh < 1: nh = 1

                if (nw, nh) != (ow, oh):
                    resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                    img = img.resize((nw, nh), resample)

                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")
                
                save_path = os.path.join(tgt_folder, f)
                if f.lower().endswith(('.jpg', '.jpeg')):
                    img.save(save_path, "JPEG", quality=quality, optimize=True)
                else:
                    img.save(save_path)
        except Exception as e:
            log_callback(f"Error with {f}: {e}")
            
    log_callback("\n✅ Image resizing complete.")
    completion_callback("Success", "Image processing is complete.")

def format_numbers_task(file_path):
    """Reads a column from a file and formats it into a single line."""
    try:
        p = clean_file_path(file_path)
        if p.lower().endswith((".xlsx",".xls")): df = pd.read_excel(p, header=None, usecols=[0], dtype=str)
        else: df = pd.read_csv(p, header=None, usecols=[0], dtype=str, on_bad_lines='skip', sep=r'\s+|\t|,', engine='python')
        nums = df[0].dropna().str.strip().tolist()
        if not nums: return ("No numbers found.", None)
        out_str = ",".join(nums)
        out_path = "formatted_numbers.txt"
        with open(out_path, "w", encoding='utf-8') as f: f.write(out_str)
        return (None, f"Formatted text saved to {os.path.abspath(out_path)}")
    except Exception as e:
        return (f"Could not process file: {e}", None)

def _process_rug_size_row(s):
    """Safely processes a rug size string."""
    try:
        w_in, h_in = size_to_inches_wh(s)
        area = calculate_sqft(s)
        return {'w': w_in, 'h': h_in, 'a': area}
    except Exception:
        return {'w': None, 'h': None, 'a': None}

def bulk_rug_sizer_task(path, col, log_callback, completion_callback):
    """Processes a sheet of rug sizes and adds calculated columns."""
    log_callback(f"Processing rug sizes from: {path}")
    try:
        df = pd.read_excel(path) if path.lower().endswith((".xlsx",".xls")) else pd.read_csv(path)
        log_callback(f"Successfully loaded {len(df)} rows from the file.")
    except Exception as e:
        log_callback(f"Error reading file: {e}"); completion_callback("Error", f"Could not read file: {e}"); return
    
    sel_col = None
    if len(col) == 1 and col.isalpha():
        idx = ord(col.upper()) - ord('A')
        if idx < len(df.columns): sel_col = df.columns[idx]
    elif col in df.columns: sel_col = col
    
    if not sel_col:
        completion_callback("Error", f"Column '{col}' not found."); return

    total_rows = len(df)
    results = []
    for i, row_value in enumerate(df[sel_col]):
        results.append(_process_rug_size_row(row_value))
        check_interval = min(100, total_rows // 10 or 1)
        if (i + 1) % check_interval == 0:
            percentage = (i + 1) * 100 / total_rows
            log_callback(f"  ...Progress: {percentage:.0f}% ({i + 1}/{total_rows})")

    res = pd.Series(results, index=df.index)
    
    df["Width_in"] = [r['w'] for r in res]; df["Height_in"] = [r['h'] for r in res]; df["Area_sqft"] = [r['a'] for r in res]
    
    out_path = f"{os.path.splitext(path)[0]}_with_sizes.xlsx"
    try:
        df.to_excel(out_path, index=False)
        log_callback(f"✅ Saved to: {out_path}")
        completion_callback("Success", f"Processed file saved to:\n{out_path}")
    except Exception as e:
        csv_path = f"{os.path.splitext(path)[0]}_with_sizes.csv"; df.to_csv(csv_path, index=False)
        log_callback(f"Could not save as Excel ({e}). ✅ Saved as CSV instead: {csv_path}")
        completion_callback("Saved as CSV", f"Could not save as Excel. Saved as CSV instead:\n{csv_path}")

def generate_qr_task(data, fname, output_type, dymo_size_info, bottom_text):
    """Generates a QR code as a PNG or Dymo label image."""
    try:
        img = qrcode.make(data)
        if output_type == "PNG":
            img.save(fname)
        else: # Dymo
            label_image = create_label_image(img, dymo_size_info, bottom_text)
            label_image.save(fname)
        return (f"✅ QR Code saved as '{fname}'", f"QR Code saved to:\n{os.path.abspath(fname)}")
    except Exception as e:
        return (f"Error generating QR Code: {e}", None)

def generate_barcode_task(data, fname, bc_format, output_type, dymo_size_info, bottom_text):
    """Generates a barcode as a PNG or Dymo label image."""
    try:
        BarcodeClass = barcode.get_barcode_class(bc_format)
        if output_type == "PNG":
            saved_fname = BarcodeClass(data, writer=ImageWriter()).save(fname.replace('.png',''))
            return (f"✅ Barcode saved as '{saved_fname}'", f"Barcode saved to:\n{os.path.abspath(saved_fname)}")
        else: # Dymo
            barcode_pil_img = BarcodeClass(data, writer=ImageWriter()).render()
            label_image = create_label_image(barcode_pil_img, dymo_size_info, bottom_text)
            label_image.save(fname)
            return (f"✅ Dymo Label saved as '{fname}'", f"Dymo Label saved to:\n{os.path.abspath(fname)}")
    except Exception as e:
        return (f"Error generating barcode: {e}", None)


def _load_font(preferred_names, size):
    for name in preferred_names:
        resource_candidates = [name]
        if not os.path.dirname(name):
            resource_candidates.append(os.path.join("fonts", name))
        for resource_name in resource_candidates:
            resource_path = get_resource_path(resource_name)
            if resource_path:
                try:
                    return ImageFont.truetype(resource_path, size=size)
                except (IOError, OSError):
                    logging.debug(
                        "Failed to load font from resource '%s'", resource_path
                    )

        if os.path.isabs(name) and os.path.exists(name):
            try:
                return ImageFont.truetype(name, size=size)
            except (IOError, OSError):
                logging.debug("Failed to load font from absolute path '%s'", name)
                continue

        try:
            return ImageFont.truetype(name, size=size)
        except (IOError, OSError):
            logging.debug("Failed to load font '%s' from system fonts", name)
            continue

    return ImageFont.load_default()


def _fit_font(draw, text, preferred_names, max_width, initial_size, min_size):
    """Return a font whose rendered width fits inside ``max_width``."""

    size = initial_size
    while size > min_size:
        font = _load_font(preferred_names, size=size)
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        if text_width <= max_width:
            return font, size
        size -= 1

    return _load_font(preferred_names, size=min_size), min_size


def generate_rinven_tag_label(details, fname, include_barcode, barcode_data):
    """Creates a Rinven tag label sized for a portrait Dymo 30256 (2.31" x 4") label."""

    DPI = 300
    width_in = 2.3125
    height_in = 4.0
    width_px = int(width_in * DPI)
    height_px = int(height_in * DPI)
    padding = int(0.14 * DPI)
    top_extra_padding = int(0.08 * DPI)

    try:
        template_candidates = [
            ("rinven_tag_template.png",),
            ("templates", "rinven_tag_template.png"),
            ("assets", "rinven_tag_template.png"),
        ]
        template_image = None
        for candidate in template_candidates:
            template_path = get_resource_path(*candidate)
            if not template_path:
                continue
            try:
                with Image.open(template_path) as img:
                    template_image = img.convert("RGB")
                break
            except (FileNotFoundError, OSError) as exc:
                logging.warning(
                    "Failed to load Rinven tag template from '%s': %s",
                    template_path,
                    exc,
                )
                template_image = None

        if template_image is not None:
            if template_image.size != (width_px, height_px):
                template_image = template_image.resize((width_px, height_px), Image.Resampling.LANCZOS)
            canvas = template_image.copy()
        else:
            canvas = Image.new("RGB", (width_px, height_px), "white")

        draw = ImageDraw.Draw(canvas)

        title_font_size = int(0.18 * DPI)
        title_pref_fonts = [
            "arialbd.ttf",
            "Arial Bold.ttf",
            "Helvetica-Bold.ttf",
            "Verdana Bold.ttf",
            "LiberationSans-Bold.ttf",
        ]
        text_font_size = int(0.14 * DPI)
        text_pref_fonts = [
            "arial.ttf",
            "Helvetica.ttf",
            "Verdana.ttf",
            "LiberationSans-Regular.ttf",
        ]

        current_y = padding + top_extra_padding

        barcode_img = None
        if include_barcode and barcode_data:
            last_barcode_error: Optional[Exception] = None
            for bc_format in ("code128", "code39"):
                try:
                    BarcodeClass = barcode.get_barcode_class(bc_format)
                    barcode_img = BarcodeClass(barcode_data, writer=ImageWriter()).render(
                        writer_options={"module_height": int(0.7 * DPI), "quiet_zone": 6}
                    )
                    break
                except Exception as exc:  # noqa: BLE001 - barcode library raises generic exceptions
                    last_barcode_error = exc
                    logging.debug("Failed to render %s barcode: %s", bc_format, exc)
            if barcode_img is None and last_barcode_error is not None:
                logging.warning(
                    "Skipping barcode rendering for Rinven tag due to error: %s",
                    last_barcode_error,
                )

        if barcode_img is not None:
            max_barcode_width = width_px - (padding * 2)
            max_barcode_height = int(height_px * 0.26)
            barcode_img.thumbnail((max_barcode_width, max_barcode_height), Image.Resampling.LANCZOS)
            barcode_x = (width_px - barcode_img.width) // 2
            canvas.paste(barcode_img, (barcode_x, current_y))
            current_y += barcode_img.height + int(0.05 * DPI)

        collection_name = details.get("collection", "").strip()
        if collection_name:
            title_max_width = width_px - (padding * 2)
            min_title_size = max(int(0.1 * DPI), 30)
            title_font, title_font_size = _fit_font(
                draw,
                collection_name,
                title_pref_fonts,
                title_max_width,
                title_font_size,
                min_title_size,
            )
            bbox = draw.textbbox((0, 0), collection_name, font=title_font)
            title_w = bbox[2] - bbox[0]
            title_h = bbox[3] - bbox[1]
            title_x = (width_px - title_w) // 2
            draw.text((title_x, current_y), collection_name, fill="black", font=title_font)
            current_y += title_h + int(0.07 * DPI)

        field_padding_top = max(current_y, padding)
        current_y = field_padding_top

        fields = [
            ("Design", details.get("design", "")),
            ("Color", details.get("color", "")),
            ("Size", details.get("size", "")),
            ("Origin", details.get("origin", "")),
            ("Style", details.get("style", "")),
            ("Content", details.get("content", "")),
            ("Type", details.get("type", "")),
            ("Rug #", details.get("rug_no", "")),
        ]

        base_colon_gap = int(0.035 * DPI)
        base_value_gap = int(0.06 * DPI)
        available_line_width = width_px - (padding * 2)
        available_height = height_px - padding - current_y
        min_text_size = max(int(0.06 * DPI), 18)

        def gap_values(font_size):
            colon_gap = max(base_colon_gap, int(font_size * 0.2))
            value_gap = max(base_value_gap, int(font_size * 0.4))
            return colon_gap, value_gap

        def measurements(font, font_size):
            colon_gap, value_gap = gap_values(font_size)
            colon_bbox = draw.textbbox((0, 0), ":", font=font)
            colon_width = colon_bbox[2] - colon_bbox[0]
            line_height = int(font_size * 1.35)
            widest_label = 0
            fits_width = True
            for label, value in fields:
                label_bbox = draw.textbbox((0, 0), label, font=font)
                label_width = label_bbox[2] - label_bbox[0]
                value_text = value.strip()
                value_bbox = draw.textbbox((0, 0), value_text, font=font)
                value_width = value_bbox[2] - value_bbox[0]
                line_width = label_width + colon_gap + colon_width + value_gap + value_width
                if line_width > available_line_width:
                    fits_width = False
                    break
                widest_label = max(widest_label, label_width)
            return fits_width, widest_label, colon_width, line_height, colon_gap, value_gap

        text_font = _load_font(text_pref_fonts, size=text_font_size)
        (
            fits_width,
            max_label_width,
            colon_width,
            line_height,
            colon_gap,
            value_gap,
        ) = measurements(text_font, text_font_size)
        total_height = line_height * len(fields)

        while (not fits_width or total_height > available_height) and text_font_size > min_text_size:
            text_font_size -= 1
            text_font = _load_font(text_pref_fonts, size=text_font_size)
            (
                fits_width,
                max_label_width,
                colon_width,
                line_height,
                colon_gap,
                value_gap,
            ) = measurements(text_font, text_font_size)
            total_height = line_height * len(fields)

        if not fits_width or total_height > available_height:
            text_font, text_font_size = _fit_font(
                draw,
                max((value.strip() for _, value in fields), key=len, default=""),
                text_pref_fonts,
                available_line_width,
                text_font_size,
                min_text_size,
            )
            (
                fits_width,
                max_label_width,
                colon_width,
                line_height,
                colon_gap,
                value_gap,
            ) = measurements(text_font, text_font_size)
            total_height = line_height * len(fields)

        for label, value in fields:
            label_text = label
            value_text = value.strip()
            draw.text((padding, current_y), label_text, fill="black", font=text_font)
            colon_x = padding + max_label_width + colon_gap
            draw.text((colon_x, current_y), ":", fill="black", font=text_font)
            value_x = colon_x + colon_width + value_gap
            draw.text((value_x, current_y), value_text, fill="black", font=text_font)
            current_y += line_height

        canvas.save(fname)
        return (
            f"✅ Rinven tag saved as '{fname}'",
            f"Rinven tag saved to:\n{os.path.abspath(fname)}",
        )
    except Exception as exc:
        logging.exception("Error generating Rinven tag")
        return (f"Error generating Rinven tag: {exc}", None)

def add_image_links_task(input_path, links_path, key_col, log_callback, completion_callback):
    log_callback("Starting process to add image links...")
    try:
        if input_path.lower().endswith((".xlsx", ".xls")):
            df_main = pd.read_excel(input_path, dtype={key_col: str})
        else:
            df_main = pd.read_csv(input_path, dtype={key_col: str})
            
        sel_col = None
        if len(key_col) == 1 and key_col.isalpha():
            idx = ord(key_col.upper()) - ord('A')
            if idx < len(df_main.columns):
                sel_col = df_main.columns[idx]
        elif key_col in df_main.columns:
            sel_col = key_col
            
        if not sel_col:
            completion_callback("Error", f"Key column '{key_col}' not found.")
            return

        log_callback("Loading image links file...")
        df_links = pd.read_csv(links_path, header=None, dtype=str)
        
        link_map = {}
        log_callback(f"Read a total of {len(df_links)} links.")
        for link in df_links[0].dropna().tolist():
            match = re.search(r"/files/([^/]+?)(-\d+)?\.(jpg|jpeg|png|webp|gif|bmp|tiff)", link, re.IGNORECASE)
            if match:
                key = match.group(1)
                clean_key = re.sub(r'(-\d+|\.|_|\s+).*$', '', key)
                
                if re.match(r'^\d+$', clean_key):
                    final_key = clean_key
                else:
                    final_key = key
                
                if final_key not in link_map:
                    link_map[final_key] = []
                link_map[final_key].append(link)
        
        log_callback(f"Found a total of {len(link_map)} unique keys.")

        for key in link_map:
            link_map[key].sort()

        log_callback("Matching and adding links to the data...")
        
        for index, row in df_main.iterrows():
            key_val = str(row[sel_col]).strip()
            
            log_callback(f"Searching for key: '{key_val}'")
            
            if key_val in link_map:
                links = link_map[key_val]
                for i, link in enumerate(links):
                    col_name = f"Image_Link_{i + 1}"
                    df_main.loc[index, col_name] = link
                log_callback(f"✅ Added {len(links)} links for key '{key_val}'.")
            else:
                log_callback(f"⚠️ Image link not found for key '{key_val}'.")
                
        out_path = f"{os.path.splitext(input_path)[0]}_with_images.xlsx"
        try:
            df_main.to_excel(out_path, index=False)
            log_callback(f"✅ File successfully saved: {out_path}")
            completion_callback("Success", f"Links have been added and the file is saved to:\n{out_path}")
        except Exception as e:
            csv_path = f"{os.path.splitext(input_path)[0]}_with_images.csv"
            df_main.to_csv(csv_path, index=False)
            log_callback(f"Could not save as Excel ({e}). ✅ Saved as CSV instead: {csv_path}")
            completion_callback("Saved as CSV", f"Could not save as Excel. Saved as CSV instead:\n{csv_path}")

    except Exception as e:
        log_callback(f"An error occurred: {e}")
        completion_callback("Error", f"An error occurred: {e}")